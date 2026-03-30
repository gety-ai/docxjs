import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlSmartTag,
	WmlAltChunk,
	WmlTableRow
} from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './options';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties, SectionType } from './document/section';
import { WmlRun, RunProperties } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlComment, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import { MountedWindowChange, VirtualizedRenderer } from './virtualized-renderer';
import { DocumentPager, PaginatedPage, PaginatedSection } from './document-pager';

const ns = {
	svg: "http://www.w3.org/2000/svg",
	mathML: "http://www.w3.org/1998/Math/MathML"
}

interface CellPos {
	col: number;
	row: number;
}

export interface RenderedDocumentHandle {
	destroy(): void;
	getMountedPages(): Array<{
		pageIndex: number;
		element: HTMLElement;
	}>;
	findMountedPage(pageIndex: number): HTMLElement | null;
	scrollToPage(
		pageIndex: number,
		options?: {
			block?: ScrollLogicalPosition;
			behavior?: ScrollBehavior;
		},
	): boolean;
}

declare const Highlight: any;

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

export class HtmlRenderer {

	className: string = "docx";
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;
	currentSectionProps: SectionProperties = null;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	defaultTabSize: string;
	currentTabs: any[] = [];

	commentHighlight: any;
	commentMap: Record<string, Range> = {};

	tasks: Promise<any>[] = [];
	postRenderTasks: any[] = [];
	pageVirtualizer: VirtualizedRenderer = null;
	lastMountedWindowSignature: string = null;

	constructor(public htmlDocument: Document) {
	}

	async render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options): Promise<RenderedDocumentHandle> {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.styleMap = null;
		this.tasks = [];
		this.postRenderTasks = [];
		this.currentTabs = [];
		this.currentEndnoteIds = [];
		this.commentMap = {};
		this.footnoteMap = {};
		this.endnoteMap = {};
		this.usedHederFooterParts = [];
		this.currentSectionProps = null;
		this.pageVirtualizer?.destroy();
		this.pageVirtualizer = null;

		if (this.options.renderComments && globalThis.Highlight) {
			this.commentHighlight = new Highlight();
		}

		styleContainer = styleContainer || bodyContainer;

		removeAllElements(styleContainer);
		removeAllElements(bodyContainer);

		styleContainer.appendChild(this.createComment("docxjs library predefined styles"));
		styleContainer.appendChild(this.renderDefaultStyle());

		if (document.themePart) {
			styleContainer.appendChild(this.createComment("docxjs document theme values"));
			this.renderTheme(document.themePart, styleContainer);
		}

		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			styleContainer.appendChild(this.createComment("docxjs document styles"));
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}

		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);

			styleContainer.appendChild(this.createComment("docxjs document numbering styles"));
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}

		if (document.footnotesPart) {
			this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
		}

		if (document.endnotesPart) {
			this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
		}

		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}

		if (!options.ignoreFonts && document.fontTablePart)
			this.renderFontTable(document.fontTablePart, styleContainer);

		const pages = document.pages ?? new DocumentPager(this.options, this.styleMap).buildPages(document.documentPart.body);
		const scrollElement = this.resolveVirtualScrollElement(bodyContainer, pages);
		let bodyHost = bodyContainer;

		if (scrollElement) {
			bodyHost = this.options.inWrapper ? this.renderWrapper([]) : this.createElement("div");
			bodyHost.dataset.docxPageCount = `${pages.length}`;
			bodyHost.dataset.docxVirtualized = "true";

			if (this.options.inWrapper || bodyHost !== bodyContainer) {
				bodyContainer.appendChild(bodyHost);
			}

			this.pageVirtualizer = new VirtualizedRenderer({
				document: this.htmlDocument,
				hostElement: bodyHost,
				scrollElement,
				items: pages.map(page => ({
					key: page.key,
					estimatedSize: page.estimatedHeight
				})),
				overscan: this.options.virtualizePagesOverscan,
				itemGap: this.options.inWrapper ? 30 : 0,
				centerItems: this.options.inWrapper,
				renderItem: index => this.renderPage(document.preparePageForRender(pages[index]), document.documentPart.body),
				onRendered: () => {
					this.flushPostRenderTasks();
					this.refreshTabStops();
				},
				onWindowChange: payload => {
					this.emitMountedPageWindowChange(pages, payload);
				}
			});
			this.pageVirtualizer.mount();
		} else {
			var sectionElements = pages.map(page => this.renderPage(document.preparePageForRender(page), document.documentPart.body));
			bodyHost.dataset.docxPageCount = `${pages.length}`;

			if (this.options.inWrapper) {
				bodyContainer.appendChild(this.renderWrapper(sectionElements));
			} else {
				appendChildren(bodyContainer, sectionElements);
			}

			this.emitMountedPageWindowChange(pages, {
				startIndex: pages[0]?.pageIndex ?? 0,
				endIndex: pages[pages.length - 1]?.pageIndex ?? 0,
				indices: pages.map((_, index) => index),
				addedIndices: pages.map((_, index) => index),
				removedIndices: [],
				items: sectionElements.map((element, index) => ({
					index,
					element
				})),
				isScrolling: false
			});
		}

		if (this.commentHighlight && options.renderComments) {
			(CSS as any).highlights.set(`${this.className}-comments`, this.commentHighlight);
		}

		this.flushPostRenderTasks();

		await Promise.allSettled(this.tasks);

		this.refreshTabStops();
		return this.createRenderHandle(document, bodyContainer, styleContainer, bodyHost, pages);
	}

	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (let [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		styleContainer.appendChild(this.createStyleElement(cssText));
	}

	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				this.tasks.push(this.document.loadFont(ref.id, ref.key).then(fontData => {
					const cssValues = {
						'font-family': encloseFontFamily(f.name),
						'src': `url(${fontData})`
					};

					if (ref.type == "bold" || ref.type == "boldItalic") {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == "italic" || ref.type == "boldItalic") {
						cssValues['font-style'] = 'italic';
					}

					const cssText = this.styleToString("@font-face", cssValues);
					styleContainer.appendChild(this.createComment(`docxjs ${f.name} font`));
					styleContainer.appendChild(this.createStyleElement(cssText));
				}));
			}
		}
	}

	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	processStyles(styles: IDomStyle[]) {
		return DocumentPager.createStyleMap(styles, this.options);
	}

	prodessNumberings(numberings: IDomNumbering[]) {
		for (let num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (var e of element.children) {
				e.parent = element;

				if (e.type == DomType.Table) {
					this.processTable(e);
				}
				else {
					this.processElement(e);
				}
			}

		}
	}

	processTable(table: WmlTable) {
		for (var r of table.children) {
			for (var c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left", "border-right", "border-top", "border-bottom",
					"padding-left", "padding-right", "padding-top", "padding-bottom"
				]);
				this.inheritTableCellPadding(table.cellStyle, c.cssStyle);

				this.processElement(c);
			}
		}
	}

	inheritTableCellPadding(input: Record<string, string>, output: Record<string, string>) {
		if (!input || !output)
			return;

		for (const key of ["padding-left", "padding-right", "padding-top", "padding-bottom"]) {
			if (input[key] != null && shouldInheritPadding(output[key])) {
				output[key] = input[key];
			}
		}
	}

	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input)
			return output;

		if (output == null) output = {};
		if (attrs == null) attrs = Object.getOwnPropertyNames(input);

		for (var key of attrs) {
			if (input.hasOwnProperty(key) && (!output.hasOwnProperty(key) || output[key] == null || output[key] === ""))
				output[key] = input[key];
		}

		return output;
	}

	createPageElement(className: string, props: SectionProperties): HTMLElement {
		var elem = this.createElement("section", { className });

		if (props) {
			if (props.pageMargins) {
				elem.style.paddingLeft = props.pageMargins.left;
				elem.style.paddingRight = props.pageMargins.right;
				elem.style.paddingTop = props.pageMargins.top;
				elem.style.paddingBottom = props.pageMargins.bottom;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth)
					elem.style.width = props.pageSize.width;
				if (!this.options.ignoreHeight)
					elem.style.minHeight = props.pageSize.height;
			}
		}

		return elem;
	}

	createSectionContent(props: SectionProperties): HTMLElement {
		var elem = this.createElement("article")

		if (props.columns && props.columns.numberOfColumns) {
			elem.style.columnCount = `${props.columns.numberOfColumns}`;
			elem.style.columnGap = props.columns.space;

			if (props.columns.separator) {
				elem.style.columnRule = "1px solid black";
			}
		}

		return elem;
	}	

	buildPages(document: DocumentElement): PaginatedPage[] {
		return new DocumentPager(this.options, this.styleMap).buildPages(document);
	}

	renderPage(page: PaginatedPage, document: DocumentElement): HTMLElement {
		this.currentFootnoteIds = [];
		this.currentEndnoteIds = page.initialEndnoteIds.slice();

		const pageElement = this.createPageElement(this.className, page.pageProps);
		pageElement.dataset.index = `${page.pageIndex}`;
		this.renderStyleValues(document.cssStyle, pageElement);

		this.options.renderHeaders && this.renderHeaderFooter(page.pageProps.headerRefs, page.pageProps,
			page.pageIndex, page.firstOfSection, pageElement);

		for (const sect of page.sections) {
			const contentElement = this.createSectionContent(sect.sectProps);
			sect.elements.forEach(element => {
				if (element.type == DomType.Table) {
					this.processTable(element as WmlTable);
				} else {
					this.processElement(element);
				}
			});
			if (this.options.mergeAdjacent) {
				sect.elements.forEach(element => this.ensureOptimizedTree(element));
			}
			this.currentSectionProps = sect.sectProps;
			this.renderSectionElements(sect, contentElement);
			pageElement.appendChild(contentElement);
		}

		if (this.options.renderFootnotes) {
			this.renderNotes(this.currentFootnoteIds, this.footnoteMap, pageElement);
		}

		if (this.options.renderEndnotes && page.isLastPage) {
			this.renderNotes(this.currentEndnoteIds, this.endnoteMap, pageElement);
		}

		this.options.renderFooters && this.renderHeaderFooter(page.footerProps.footerRefs, page.footerProps,
			page.pageIndex, page.firstOfSection, pageElement);

		return pageElement;
	}

	renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
		if (!refs) return;

		var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
			?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
			?? refs.find(x => x.type == "default");

		var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart;

		if (part) {
			const previousPart = this.currentPart;
			const previousSectionProps = this.currentSectionProps;
			this.currentPart = part;
			this.currentSectionProps = props;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			if (this.options.mergeAdjacent) {
				this.ensureOptimizedTree(part.rootElement);
			}
			const [el] = this.renderElements([part.rootElement], into) as HTMLElement[];

			if (props?.pageMargins) {
				if (part.rootElement.type === DomType.Header) {
					el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
					el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
				}
				else if (part.rootElement.type === DomType.Footer) {
					el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
					el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
				}
			}

			this.currentPart = previousPart;
			this.currentSectionProps = previousSectionProps;
		}
	}

	isPageBreakElement(elem: OpenXmlElement): boolean {
		if (elem.type != DomType.Break)
			return false;

		if ((elem as WmlBreak).break == "lastRenderedPageBreak")
			return !this.options.ignoreLastRenderedPageBreak;

		return (elem as WmlBreak).break == "page";
	}

	isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean {
		if (!prev) return false;
		if (!next) return false;

		return prev.pageSize?.orientation != next.pageSize?.orientation
			|| prev.pageSize?.width != next.pageSize?.width
			|| prev.pageSize?.height != next.pageSize?.height;
	}

	splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): PaginatedSection[] {
		var current: PaginatedSection = { sectProps: null, elements: [], pageBreak: false };
		var result = [current];

		for (let elem of elements) {
			if (elem.type == DomType.Paragraph) {
				const s = this.findStyle((elem as WmlParagraph).styleName);

				if (s?.paragraphProps?.pageBreakBefore && current.elements.length > 0) {
					current.pageBreak = true;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}
			}

			current.elements.push(elem);

			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;

				var sectProps = p.sectionProps;
				var pBreakIndex = -1;
				var rBreakIndex = -1;

				if (this.options.breakPages && p.children) {
					pBreakIndex = p.children.findIndex(r => {
						rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
						return rBreakIndex != -1;
					});
				}

				if (sectProps || pBreakIndex != -1) {
					current.sectProps = sectProps;
					current.pageBreak = pBreakIndex != -1;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}

				if (pBreakIndex != -1) {
					let breakRun = p.children[pBreakIndex];
					let splitRun = rBreakIndex < breakRun.children.length - 1;

					if (pBreakIndex < p.children.length - 1 || splitRun) {
						var children = elem.children;
						var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
						elem.children = children.slice(0, pBreakIndex);
						current.elements.push(newParagraph);

						if (splitRun) {
							let runChildren = breakRun.children;
							let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
							elem.children.push(newRun);
							breakRun.children = runChildren.slice(rBreakIndex);
						}
					}
				}
			}
		}

		let currentSectProps = null;

		for (let i = result.length - 1; i >= 0; i--) {
			if (result[i].sectProps == null) {
				result[i].sectProps = currentSectProps ?? defaultProps;
			} else {
				currentSectProps = result[i].sectProps
			}
		}

		return this.coalesceEmptySections(this.resolveSectionProps(result));
	}

	resolveSectionProps(sections: PaginatedSection[]) {
		let previous: SectionProperties = null;

		for (const section of sections) {
			if (previous) {
				section.sectProps = this.mergeSectionProps(previous, section.sectProps);
			}

			previous = section.sectProps;
		}

		return sections;
	}

	mergeSectionProps(base: SectionProperties, override: SectionProperties): SectionProperties {
		if (!base)
			return override;

		if (!override)
			return base;

		if (base === override)
			return base;

		return {
			...base,
			...override,
			type: override.type ?? base.type,
			pageSize: override.pageSize ?? base.pageSize,
			pageMargins: override.pageMargins ?? base.pageMargins,
			pageBorders: override.pageBorders ?? base.pageBorders,
			pageNumber: override.pageNumber ?? base.pageNumber,
			columns: override.columns,
			footerRefs: override.footerRefs ?? base.footerRefs,
			headerRefs: override.headerRefs ?? base.headerRefs,
			titlePage: override.titlePage ?? base.titlePage,
		};
	}

	coalesceEmptySections(sections: PaginatedSection[]) {
		const result: PaginatedSection[] = [];

		for (let i = 0; i < sections.length; i++) {
			const section = sections[i];
			const next = sections[i + 1];

			if (next && !section.pageBreak && !this.sectionHasVisibleContent(section) && !this.sectionForcesStandalonePage(section)) {
				next.elements = [...section.elements, ...next.elements];
				next.sectProps = this.mergeSectionProps(section.sectProps, next.sectProps);
				next.pageBreak = section.pageBreak || next.pageBreak;
				continue;
			}

			result.push(section);
		}

		return result;
	}

	sectionForcesStandalonePage(section: PaginatedSection) {
		switch (section.sectProps?.type) {
			case SectionType.EvenPage:
			case SectionType.OddPage:
				return true;
			default:
				return false;
		}
	}

	sectionHasVisibleContent(section: PaginatedSection) {
		return section.elements?.some(element => this.elementHasVisibleContent(element)) ?? false;
	}

	elementHasVisibleContent(element: OpenXmlElement) {
		if (!element)
			return false;

		switch (element.type) {
			case DomType.Text:
			case DomType.DeletedText:
				return !!(element as WmlText).text?.trim();

			case DomType.Image:
			case DomType.Drawing:
			case DomType.Table:
			case DomType.Symbol:
			case DomType.Tab:
			case DomType.NoBreakHyphen:
			case DomType.FootnoteReference:
			case DomType.EndnoteReference:
			case DomType.CommentReference:
			case DomType.AltChunk:
				return true;
		}

		return element.children?.some(child => this.elementHasVisibleContent(child)) ?? false;
	}

	groupByPageBreaks(sections: PaginatedSection[]): PaginatedSection[][] {
		let current = [];
		let prev: SectionProperties;
		const result: PaginatedSection[][] = [current];

		for (let s of sections) {
			current.push(s);

			if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
				result.push(current = []);

			prev = s.sectProps;
		}

		return result.filter(x => x.length > 0);
	}

	renderWrapper(children: HTMLElement[]) {
		return this.createElement("div", { className: `${this.className}-wrapper` }, children);
	}

	renderDefaultStyle() {
		var c = this.className;
		var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
		if (this.options.hideWrapperOnPrint) {
			wrapperStyle = `@media not print { ${wrapperStyle} }`;
		}
		var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
`;

		if (this.options.renderComments) {
			styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`
		};

		return this.createStyleElement(styleText);
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	//     let css = "";
	//     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	//     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	//     const topCounters = [];

	//     for(let num of numberingPart.numberings) {
	//         const absNum = numberingMap[num.abstractId];

	//         for(let lvl of absNum.levels) {
	//             const className = this.numberingClass(num.id, lvl.level);
	//             let listStyleType = "none";

	//             if(lvl.text && lvl.format == 'decimal') {
	//                 const counter = this.numberingCounter(num.id, lvl.level);

	//                 if (lvl.level > 0) {
	//                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	//                         "counter-reset": counter
	//                     });
	//                 } else {
	//                     topCounters.push(counter);
	//                 }

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": this.levelTextToContent(lvl.text, num.id),
	//                     "counter-increment": counter
	//                 });
	//             } else if(lvl.bulletPictureId) {
	//                 let pict = bulletMap[lvl.bulletPictureId];
	//                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": "' '",
	//                     "display": "inline-block",
	//                     "background": `var(${variable})`
	//                 }, pict.style);

	//                 this.document.loadNumberingImage(pict.referenceId).then(data => {
	//                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	//                     container.appendChild(createStyleElement(text));
	//                 });
	//             } else {
	//                 listStyleType = this.numFormatToCssValue(lvl.format);
	//             }

	//             css += this.styleToString(`p.${className}`, {
	//                 "display": "list-item",
	//                 "list-style-position": "inside",
	//                 "list-style-type": listStyleType,
	//                 //TODO
	//                 //...num.style
	//             });
	//         }
	//     }

	//     if (topCounters.length > 0) {
	//         css += this.styleToString(`.${this.className}-wrapper`, {
	//             "counter-reset": topCounters.join(" ")
	//         });
	//     }

	//     return createStyleElement(css);
	// }

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		var styleText = "";
		var resetCounters = [];

		for (var num of numberings) {
			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.tasks.push(this.document.loadNumberingImage(num.bullet.src).then(data => {
					var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(this.createStyleElement(text));
				}));
			}
			else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-set": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return this.createStyleElement(styleText);
	}

	renderStyles(styles: IDomStyle[]): HTMLElement {
		var styleText = "";
		const stylesMap = this.styleMap;
		const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

		for (const style of styles) {
			var subStyles = style.styles;

			if (style.linked) {
				var linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle)
					subStyles = subStyles.concat(linkedStyle.styles);
				else if (this.options.debug)
					console.warn(`Can't find linked style ${style.linked}`);
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				var selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''} 

				if (style.target != subStyle.target)
					selector += ` ${subStyle.target}`;

				if (defautStyles[style.target] == style)
					selector = `.${this.className} ${style.target}, ` + selector;

				styleText += this.styleToString(selector, subStyle.values);
			}
		}

		return this.createStyleElement(styleText);
	}

	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, into: HTMLElement) {
		var notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			notes.forEach(note => {
				this.processElement(note);
				if (this.options.mergeAdjacent) {
					this.ensureOptimizedTree(note);
				}
			});
			var result = this.createElement("ol", null, this.renderElements(notes));
			into.appendChild(result);
		}
	}

	renderElement(elem: OpenXmlElement): Node | Node[] {
		switch (elem.type) {
			case DomType.Paragraph:
				return this.renderParagraph(elem as WmlParagraph);

			case DomType.BookmarkStart:
				return this.renderBookmarkStart(elem as WmlBookmarkStart);

			case DomType.BookmarkEnd:
				return null; //ignore bookmark end

			case DomType.Run:
				return this.renderRun(elem as WmlRun);

			case DomType.Table:
				return this.renderTable(elem);

			case DomType.Row:
				return this.renderTableRow(elem);

			case DomType.Cell:
				return this.renderTableCell(elem);

			case DomType.Hyperlink:
				return this.renderHyperlink(elem);
			
			case DomType.SmartTag:
				return this.renderSmartTag(elem);

			case DomType.Drawing:
				return this.renderDrawing(elem);

			case DomType.Image:
				return this.renderImage(elem as IDomImage);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.DeletedText:
				return this.renderDeletedText(elem as WmlText);
	
			case DomType.Tab:
				return this.renderTab(elem);

			case DomType.Symbol:
				return this.renderSymbol(elem as WmlSymbol);

			case DomType.Break:
				return this.renderBreak(elem as WmlBreak);

			case DomType.Footer:
				return this.renderContainer(elem, "footer");

			case DomType.Header:
				return this.renderContainer(elem, "header");

			case DomType.Footnote:
			case DomType.Endnote:
				return this.renderContainer(elem, "li");

			case DomType.FootnoteReference:
				return this.renderFootnoteReference(elem as WmlNoteReference);

			case DomType.EndnoteReference:
				return this.renderEndnoteReference(elem as WmlNoteReference);

			case DomType.NoBreakHyphen:
				return this.createElement("wbr");

			case DomType.VmlPicture:
				return this.renderVmlPicture(elem);

			case DomType.VmlElement:
				return this.renderVmlElement(elem as VmlElement);
	
			case DomType.MmlMath:
				return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
	
			case DomType.MmlMathParagraph:
				return this.renderContainer(elem, "span");

			case DomType.MmlFraction:
				return this.renderContainerNS(elem, ns.mathML, "mfrac");

			case DomType.MmlBase:
				return this.renderContainerNS(elem, ns.mathML, 
					elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				return this.renderContainerNS(elem, ns.mathML, "mrow");

			case DomType.MmlGroupChar:
				return this.renderMmlGroupChar(elem);

			case DomType.MmlLimitLower:
				return this.renderContainerNS(elem, ns.mathML, "munder");

			case DomType.MmlMatrix:
				return this.renderContainerNS(elem, ns.mathML, "mtable");

			case DomType.MmlMatrixRow:
				return this.renderContainerNS(elem, ns.mathML, "mtr");
	
			case DomType.MmlRadical:
				return this.renderMmlRadical(elem);

			case DomType.MmlSuperscript:
				return this.renderContainerNS(elem, ns.mathML, "msup");

			case DomType.MmlSubscript:
				return this.renderContainerNS(elem, ns.mathML, "msub");

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				return this.renderContainerNS(elem, ns.mathML, "mn");

			case DomType.MmlFunctionName:
				return this.renderContainerNS(elem, ns.mathML, "ms");
	
			case DomType.MmlDelimiter:
				return this.renderMmlDelimiter(elem);

			case DomType.MmlRun:
				return this.renderMmlRun(elem);

			case DomType.MmlNary:
				return this.renderMmlNary(elem);

			case DomType.MmlPreSubSuper:
				return this.renderMmlPreSubSuper(elem);

			case DomType.MmlBar:
				return this.renderMmlBar(elem);
	
			case DomType.MmlEquationArray:
				return this.renderMllList(elem);

			case DomType.Inserted:
				return this.renderInserted(elem);

			case DomType.Deleted:
				return this.renderDeleted(elem);

			case DomType.CommentRangeStart:
				return this.renderCommentRangeStart(elem);

			case DomType.CommentRangeEnd:
				return this.renderCommentRangeEnd(elem);

			case DomType.CommentReference:
				return this.renderCommentReference(elem);

			case DomType.AltChunk:
				return this.renderAltChunk(elem);
		}

		return null;
	}
	renderElements(elems: OpenXmlElement[], into?: Node): Node[] {
		if (elems == null)
			return null;

		var result = elems.flatMap(e => e ? this.renderElement(e) : null).filter(e => e != null);

		if (into)
			appendChildren(into, result);

		return result;
	}

	renderContainer<T extends keyof HTMLElementTagNameMap>(elem: OpenXmlElement, tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>): HTMLElementTagNameMap[T] {
		return this.createElement<T>(tagName, props, this.renderElements(elem.children));
	}

	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		return this.createElementNS(ns, tagName, props, this.renderElements(elem.children));
	}

	renderParagraph(elem: WmlParagraph) {
		if (elem.sectionProps && !this.elementHasVisibleContent(elem)) {
			return null;
		}

		var result = this.renderContainer(elem, "p");

		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);
		this.renderCommonProperties(result.style, elem);

		const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
		}

		this.normalizeRenderedDrawingParagraph(result);

		return result;
	}

	renderRunProperties(style: any, props: RunProperties) {
		this.renderCommonProperties(style, props);
	}

	renderCommonProperties(style: any, props: CommonProperties) {
		if (props == null)
			return;

		if (props.color) {
			style["color"] = props.color;
		}

		if (props.fontSize) {
			style["font-size"] = props.fontSize;
		}
	}

	renderHyperlink(elem: WmlHyperlink) {
		var result = this.renderContainer(elem, "a");

		this.renderStyleValues(elem.cssStyle, result);

		let href = '';

		if (elem.id) {
			const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
			href = rel?.target ?? href;
		}

		if (elem.anchor) {
			href += `#${elem.anchor}`;
		}

		result.href = href;

		return result;
	}
	
	renderSmartTag(elem: WmlSmartTag) {
		return this.renderContainer(elem, "span");
	}
	
	renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = new Range();
		this.commentHighlight?.add(rng);

		const result = this.createComment(`start of comment #${commentStart.id}`);
		this.later(() => rng.setStart(result, 0));
		this.commentMap[commentStart.id] = rng;

		return result
	}

	renderCommentRangeEnd(commentEnd: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = this.commentMap[commentEnd.id];
		const result = this.createComment(`end of comment #${commentEnd.id}`);
		this.later(() => rng?.setEnd(result, 0));

		return result;
	}

	renderCommentReference(commentRef: WmlCommentReference) {
		if (!this.options.renderComments)
			return null;

		var comment = this.document.commentsPart?.commentMap[commentRef.id];

		if (!comment)
			return null;

		const frg = new DocumentFragment();
		const commentRefEl = this.createElement("span", { className: `${this.className}-comment-ref` }, ['💬']);
		const commentsContainerEl = this.createElement("div", { className: `${this.className}-comment-popover` });

		this.renderCommentContent(comment, commentsContainerEl);

		frg.appendChild(this.createComment(`comment #${comment.id} by ${comment.author} on ${comment.date}`));
		frg.appendChild(commentRefEl);
		frg.appendChild(commentsContainerEl);

		return frg;
	}

	renderAltChunk(elem: WmlAltChunk) {
		if (!this.options.renderAltChunks)
			return null;

		var result = this.createElement("iframe");
		
		this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
			result.srcdoc = x;
		}));

		return result;
	}

	renderCommentContent(comment: WmlComment, container: Node) {
		container.appendChild(this.createElement('div', { className: `${this.className}-comment-author` }, [comment.author]));
		container.appendChild(this.createElement('div', { className: `${this.className}-comment-date` }, [new Date(comment.date).toLocaleString()]));

		this.renderElements(comment.children, container);
	}

	renderDrawing(elem: OpenXmlElement) {
		var result = this.renderContainer(elem, "div");
		const anchor = elem.props?.drawingAnchor;
		const pageMargins = this.currentSectionProps?.pageMargins;
		const isPageRelativeAnchor = anchor?.wrapType == "wrapNone"
			&& pageMargins
			&& (anchor?.posX?.relative == "page" || anchor?.posY?.relative == "page");

		result.style.display = "inline-block";
		result.style.textIndent = "0px";

		this.renderStyleValues(elem.cssStyle, result);

		if (isPageRelativeAnchor) {
			result.style.display = "block";
			result.style.position = "absolute";

			if (anchor.posX?.relative == "page" && anchor.posX.offset) {
				result.style.left = `calc(${anchor.posX.offset} - ${pageMargins.left ?? "0px"})`;
			}

			if (anchor.posY?.relative == "page" && anchor.posY.offset) {
				result.style.top = `calc(${anchor.posY.offset} - ${pageMargins.top ?? "0px"})`;
			}
		}
		else if (!result.style.position) {
			result.style.position = "relative";
		}

		return result;
	}

	renderImage(elem: IDomImage) {
		let result = this.createElement("img");
		let transform = elem.cssStyle?.transform;

		this.renderStyleValues(elem.cssStyle, result);

		if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
			var [left, top, right, bottom] = elem.srcRect;
			transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
			result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
		}

		if (elem.rotation)
			transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;

		result.style.transform = transform?.trim();

		if (this.document) {
			this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
				result.src = x;
				if (this.options.onImageRendered && elem.docPrId) {
					result.onload = () => {
						this.options.onImageRendered({
							docPrId: elem.docPrId,
							imgEl: result,
							naturalSize: { width: result.naturalWidth, height: result.naturalHeight }
						});
					};
				}
			}));
		}

		return result;
	}

	renderText(elem: WmlText) {
		return this.htmlDocument.createTextNode(elem.text);
	}

	renderDeletedText(elem: WmlText) {
		return this.options.renderChanges ? this.renderText(elem) : null;
	}

	renderBreak(elem: WmlBreak) {
		if (elem.break == "column") {
			const result = this.createElement("span");
			result.style.display = "block";
			result.style.height = "0";
			result.style.breakAfter = "column";
			return result;
		}

		if (elem.break == "textWrapping") {
			return this.createElement("br");
		}

		return null;
	}

	renderSectionElements(section: PaginatedSection, contentElement: HTMLElement) {
		if (this.shouldRenderManualColumns(section)) {
			this.renderManualColumns(section, contentElement);
			return;
		}

		this.renderElements(section.elements, contentElement);
	}

	shouldRenderManualColumns(section: PaginatedSection) {
		const columns = section.sectProps?.columns;

		if (!columns || columns.numberOfColumns <= 1 || columns.equalWidth)
			return false;

		if (!columns.columns?.length || columns.columns.length < columns.numberOfColumns)
			return false;

		return this.hasColumnBreak(section.elements);
	}

	hasColumnBreak(elements: OpenXmlElement[]) {
		return elements?.some(element => this.elementHasBreak(element, "column")) ?? false;
	}

	elementHasBreak(element: OpenXmlElement, type: WmlBreak["break"]) {
		if (!element)
			return false;

		if (element.type == DomType.Break) {
			return (element as WmlBreak).break == type;
		}

		return element.children?.some(child => this.elementHasBreak(child, type)) ?? false;
	}

	renderManualColumns(section: PaginatedSection, contentElement: HTMLElement) {
		const columns = section.sectProps.columns;
		const groups = this.splitSectionByColumnBreaks(section.elements);
		const columnCount = Math.max(columns.numberOfColumns, groups.length);

		contentElement.style.columnCount = "";
		contentElement.style.columnGap = "";
		contentElement.style.columnRule = "";
		contentElement.style.display = "flex";
		contentElement.style.alignItems = "flex-start";
		contentElement.style.gap = columns.space ?? "0";

		for (let i = 0; i < columnCount; i++) {
			const column = this.createElement("div");
			const columnProps = columns.columns?.[i];

			column.style.boxSizing = "border-box";
			column.style.minWidth = "0";
			column.style.flex = "0 0 auto";
			column.style.width = columnProps?.width ?? this.getEqualColumnWidth(columns);

			if (!columnProps?.width) {
				column.style.flex = "1 1 0";
			}

			if (columns.separator && i > 0) {
				column.style.borderLeft = "1px solid black";
				column.style.paddingLeft = columns.space ?? "0";
			}

			this.renderElements(groups[i] ?? [], column);
			contentElement.appendChild(column);
		}
	}

	getEqualColumnWidth(columns: SectionProperties["columns"]) {
		const count = Math.max(columns?.numberOfColumns ?? 1, 1);
		return `calc((100% - (${columns?.space ?? "0"} * ${count - 1})) / ${count})`;
	}

	splitSectionByColumnBreaks(elements: OpenXmlElement[]) {
		const result: OpenXmlElement[][] = [[]];
		let current = result[0];

		for (const element of elements) {
			const fragments = this.splitElementByColumnBreaks(element);

			for (let i = 0; i < fragments.length; i++) {
				current.push(...fragments[i]);

				if (i < fragments.length - 1) {
					current = [];
					result.push(current);
				}
			}
		}

		return result;
	}

	splitElementByColumnBreaks(element: OpenXmlElement): OpenXmlElement[][] {
		if (element?.type != DomType.Paragraph || !element.children?.length) {
			return [[element]];
		}

		const fragments = this.splitParagraphChildrenByColumnBreaks(element.children);

		if (fragments.length == 1) {
			return [[element]];
		}

		return fragments.map(children => {
			if (!children.length)
				return [];

			return [{ ...element, children }];
		});
	}

	splitParagraphChildrenByColumnBreaks(children: OpenXmlElement[]) {
		const result: OpenXmlElement[][] = [[]];
		let current = result[0];

		for (const child of children) {
			const fragments = this.splitChildByColumnBreaks(child);

			for (let i = 0; i < fragments.length; i++) {
				const fragment = fragments[i];

				if (fragment) {
					current.push(fragment);
				}

				if (i < fragments.length - 1) {
					current = [];
					result.push(current);
				}
			}
		}

		return result;
	}

	splitChildByColumnBreaks(child: OpenXmlElement): (OpenXmlElement | null)[] {
		if (!child?.children?.length) {
			return [child];
		}

		const result: (OpenXmlElement | null)[] = [];
		let current: OpenXmlElement[] = [];
		let hasBreak = false;

		for (const nestedChild of child.children) {
			if (this.isColumnBreakElement(nestedChild)) {
				result.push(current.length ? { ...child, children: current } : null);
				current = [];
				hasBreak = true;
				continue;
			}

			current.push(nestedChild);
		}

		if (!hasBreak) {
			return [child];
		}

		result.push(current.length ? { ...child, children: current } : null);
		return result;
	}

	isColumnBreakElement(elem: OpenXmlElement) {
		return elem?.type == DomType.Break && (elem as WmlBreak).break == "column";
	}

	renderInserted(elem: OpenXmlElement): Node | Node[] {
		if (this.options.renderChanges)
			return this.renderContainer(elem, "ins");

		return this.renderElements(elem.children);
	}

	renderDeleted(elem: OpenXmlElement): Node {
		if (this.options.renderChanges)
			return this.renderContainer(elem, "del");

		return null;
	}

	renderSymbol(elem: WmlSymbol) {
		var span = this.createElement("span");
		span.style.fontFamily = elem.font;
		span.innerHTML = `&#x${elem.char};`
		return span;
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		var result = this.createElement("sup");
		this.currentFootnoteIds.push(elem.id);
		result.textContent = `${this.currentFootnoteIds.length}`;
		return result;
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		var result = this.createElement("sup");
		this.currentEndnoteIds.push(elem.id);
		result.textContent = `${this.currentEndnoteIds.length}`;
		return result;
	}

	renderTab(elem: OpenXmlElement) {
		var tabSpan = this.createElement("span");

		tabSpan.innerHTML = "&emsp;";//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			var stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
		return this.createElement("span", { id: elem.name });
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun)
			return null;

		const result = this.createElement("span");
		const style = elem.cssStyle ? { ...elem.cssStyle } : null;
		const drawingOnlyOffset = this.extractDrawingOnlyRunOffset(style, elem);

		if (elem.id)
			result.id = elem.id;

		this.renderClass(elem, result);
		this.renderStyleValues(style, result);

		if (drawingOnlyOffset) {
			result.style.display = "inline-block";
			result.style.lineHeight = "0";
			result.style.position = "relative";
			result.style.top = drawingOnlyOffset;
			result.style.verticalAlign = "top";
		}

		if (elem.verticalAlign) {
			const wrapper = this.createElement(elem.verticalAlign as any);
			this.renderElements(elem.children, wrapper);
			result.appendChild(wrapper);
		}
		else {
			this.renderElements(elem.children, result);
		}

		return result;
	}

	extractDrawingOnlyRunOffset(style: Record<string, string>, elem: WmlRun) {
		const verticalAlign = style?.["vertical-align"];

		if (!verticalAlign || !this.isDrawingOnlyRun(elem) || !/^[-\d.]+(?:pt|px)$/.test(verticalAlign)) {
			return null;
		}

		delete style["vertical-align"];
		return verticalAlign;
	}

	isDrawingOnlyRun(elem: WmlRun) {
		if (!elem.children?.length) {
			return false;
		}

		return elem.children.every(child => this.isDrawingOnlyInlineElement(child));
	}

	isDrawingOnlyInlineElement(elem: OpenXmlElement) {
		switch (elem?.type) {
			case DomType.Drawing:
			case DomType.Image:
				return true;
			case DomType.Run:
			case DomType.Hyperlink:
			case DomType.SmartTag:
				return elem.children?.every(child => this.isDrawingOnlyInlineElement(child)) ?? false;
			default:
				return false;
		}
	}

	normalizeRenderedDrawingParagraph(paragraph: HTMLParagraphElement) {
		if (!paragraph || paragraph.childElementCount != 1 || paragraph.textContent?.trim()) {
			return;
		}

		if (!paragraph.querySelector("img, svg")) {
			return;
		}

		paragraph.style.lineHeight = "0";
		paragraph.style.minHeight = "0";
		paragraph.style.fontSize = "0";
	}

	renderTable(elem: WmlTable) {
		let result = this.createElement("table");

		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		if (elem.columns)
			result.appendChild(this.renderTableColumns(elem.columns));

		this.renderClass(elem, result);
		this.renderElements(elem.children, result);
		this.renderStyleValues(elem.cssStyle, result);

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();

		return result;
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		let result = this.createElement("colgroup");

		for (let col of columns) {
			let colElem = this.createElement("col");

			if (col.width)
				colElem.style.width = col.width;

			result.appendChild(colElem);
		}

		return result;
	}

	renderTableRow(elem: WmlTableRow) {
		let result = this.createElement("tr");

		this.currentCellPosition.col = 0;

		if (elem.gridBefore)
			result.appendChild(this.renderTableCellPlaceholder(elem.gridBefore));

		this.renderClass(elem, result);
		this.renderElements(elem.children, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.gridAfter)
			result.appendChild(this.renderTableCellPlaceholder(elem.gridAfter));

		this.currentCellPosition.row++;

		return result;
	}

	renderTableCellPlaceholder(colSpan: number) {
		const result = this.createElement("td", { colSpan })
		result.style['border'] = 'none';
		return result;
	}

	renderTableCell(elem: WmlTableCell) {
		let result = this.renderContainer(elem, "td");

		const key = this.currentCellPosition.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == "restart") {
				this.currentVerticalMerge[key] = result;
				result.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				result.style.display = "none";
			}
		} else {
			this.currentVerticalMerge[key] = null;
		}

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.span)
			result.colSpan = elem.span;

		this.currentCellPosition.col += result.colSpan;

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		return this.renderContainer(elem, "div");
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		var container = this.createSvgElement("svg");

		container.setAttribute("style", elem.cssStyleText);

		const result = this.renderVmlChildElement(elem);

		if (elem.imageHref?.id) {
			this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x)));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			const bb = (container.firstElementChild as any).getBBox();

			container.setAttribute("width", `${Math.ceil(bb.x +  bb.width)}`);
			container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement): any {
		const result = this.createSvgElement(elem.tagName as any);
		Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				result.appendChild(...asArray(this.renderElement(child as any)));
			}
		}

		return result;
	}

	renderMmlRadical(elem: OpenXmlElement): HTMLElement {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return this.createElementNS(ns.mathML, "msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return this.createElementNS(ns.mathML, "mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement): HTMLElement {		
		const children = [];

		children.push(this.createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(this.createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));

		return this.createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement): HTMLElement {		
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;

		const charElem = this.createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(this.createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
		} else if(supElem) {
			children.push(this.createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
		} else if(subElem) {
			children.push(this.createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = this.createElementNS(ns.mathML, "mo", null);

		children.push(this.createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(this.createElementNS(ns.mathML, "mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const result = this.renderContainerNS(elem, ns.mathML, "mrow");

		switch(elem.props.position) {
			case "top": result.style.textDecoration = "overline"; break
			case "bottom": result.style.textDecoration = "underline"; break
		}

		return result;
	}

	renderMmlRun(elem: OpenXmlElement) {
		const result = this.createElementNS(ns.mathML, "ms", null, this.renderElements(elem.children));

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		return result;
	}

	renderMllList(elem: OpenXmlElement) {
		const result = this.createElementNS(ns.mathML, "mtable");

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		for (let child of this.renderElements(elem.children)) {
			result.appendChild(this.createElementNS(ns.mathML, "mtr", null, [
				this.createElementNS(ns.mathML, "mtd", null, [child])
			]));
		}

		return result;
	}


	renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
		if (!style)
			return;

		for (let k in style) {
			if (k.startsWith("$")) {
				ouput.setAttribute(k.slice(1), style[k]);
			} else {
				ouput.style[k] = style[k];
			}
		}
	}

	renderClass(input: OpenXmlElement, ouput: HTMLElement) {
		if (input.className)
			ouput.className = input.className;

		if (input.styleName)
			ouput.classList.add(this.processStyleName(input.styleName));
	}

	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$'))
				continue;
			
			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText)
			result += cssText;

		return result + "}\r\n";
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		var result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
	}

	numFormatToCssValue(format: string) {
		var mapping = {
			none: "none",
			bullet: "disc",
			decimal: "decimal",
			lowerLetter: "lower-alpha",
			upperLetter: "upper-alpha",
			lowerRoman: "lower-roman",
			upperRoman: "upper-roman",
			decimalZero: "decimal-leading-zero", // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: "katakana",
			aiueoFullWidth: "katakana",
			chineseCounting: "simp-chinese-informal",
			chineseCountingThousand: "simp-chinese-informal",
			chineseLegalSimplified: "simp-chinese-formal", // 中文大写
			chosung: "hangul-consonant",
			ideographDigital: "cjk-ideographic",
			ideographTraditional: "cjk-heavenly-stem", // 十天干
			ideographLegalTraditional: "trad-chinese-formal",
			ideographZodiac: "cjk-earthly-branch", // 十二地支
			iroha: "katakana-iroha",
			irohaFullWidth: "katakana-iroha",
			japaneseCounting: "japanese-informal",
			japaneseDigitalTenThousand: "cjk-decimal",
			japaneseLegal: "japanese-formal",
			thaiNumbers: "thai",
			koreanCounting: "korean-hangul-formal",
			koreanDigital: "korean-hangul-formal",
			koreanDigital2: "korean-hanja-informal",
			hebrew1: "hebrew",
			hebrew2: "hebrew",
			hindiNumbers: "devanagari",
			ganada: "hangul",
			taiwaneseCounting: "cjk-ideographic",
			taiwaneseCountingThousand: "cjk-ideographic",
			taiwaneseDigital:  "cjk-decimal",
		};

		return mapping[format] ?? format;
	}

	refreshTabStops() {
		if (!this.options.experimental)
			return;

		setTimeout(() => {
			const pixelToPoint = computePixelToPoint();

			for (let tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}

	createElementNS(ns: string, tagName: string, props?: Partial<Record<any, any>>, children?: ChildType[]): any {
		var result = ns ? this.htmlDocument.createElementNS(ns, tagName) : this.htmlDocument.createElement(tagName);
		Object.assign(result, props);
		children && appendChildren(result, children);
		return result;
	}

	createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>, children?: ChildType[]): HTMLElementTagNameMap[T] {
		return this.createElementNS(undefined, tagName, props, children);
	}

	createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>, children?: ChildType[]): SVGElementTagNameMap[T] {
		return this.createElementNS(ns.svg, tagName, props, children);
	}

	createStyleElement(cssText: string) {
		return this.createElement("style", { innerHTML: cssText });
	}
	
	createComment(text: string) {
		return this.htmlDocument.createComment(text);
	}

	later(func: Function) { 
		this.postRenderTasks.push(func);
	}

	flushPostRenderTasks(fromIndex: number = 0) {
		if (fromIndex >= this.postRenderTasks.length)
			return;

		const tasks = this.postRenderTasks.splice(fromIndex);
		tasks.forEach(task => task());
	}

	resolveVirtualScrollElement(bodyContainer: HTMLElement, pages: PaginatedPage[]): HTMLElement {
		if (!this.options.virtualizePages || pages.length < 2 || this.options.renderComments)
			return null;

		return findScrollableElement(bodyContainer, this.htmlDocument);
	}

	createRenderHandle(
		document: WordDocument,
		bodyContainer: HTMLElement,
		styleContainer: HTMLElement,
		bodyHost: HTMLElement,
		pages: PaginatedPage[]
	): RenderedDocumentHandle {
		const pageIndexMap = new Map(pages.map((page, index) => [page.pageIndex, index]));

		return {
			destroy: () => {
				if (this.commentHighlight && this.options.renderComments) {
					(CSS as any).highlights.delete(`${this.className}-comments`);
				}

				this.pageVirtualizer?.destroy();
				this.pageVirtualizer = null;

				removeAllElements(bodyContainer);

				if (styleContainer && styleContainer !== bodyContainer) {
					removeAllElements(styleContainer);
				}

				void document.dispose();
			},
			getMountedPages: () => {
				if (this.pageVirtualizer) {
					return this.pageVirtualizer.getMountedItems().map(item => ({
						pageIndex: pages[item.index].pageIndex,
						element: item.element
					}));
				}

				return this.getStaticMountedPages(bodyHost);
			},
			findMountedPage: (pageIndex: number) => {
				if (this.pageVirtualizer) {
					const virtualIndex = pageIndexMap.get(pageIndex);
					return virtualIndex == null ? null : this.pageVirtualizer.findMountedItem(virtualIndex);
				}

				return bodyHost.querySelector(`section.${this.className}[data-index="${pageIndex}"]`);
			},
			scrollToPage: (pageIndex: number, options = {}) => {
				const virtualIndex = pageIndexMap.get(pageIndex);

				if (virtualIndex == null) {
					return false;
				}

				if (this.pageVirtualizer) {
					this.pageVirtualizer.scrollToIndex(virtualIndex, options);
					return true;
				}

				const page = bodyHost.querySelector(`section.${this.className}[data-index="${pageIndex}"]`) as HTMLElement | null;

				if (!page) {
					return false;
				}

				page.scrollIntoView(options);
				return true;
			}
		};
	}

	getStaticMountedPages(bodyHost: HTMLElement) {
		return Array
			.from(bodyHost.querySelectorAll(`section.${this.className}[data-index]`))
			.map(element => ({
				pageIndex: Number((element as HTMLElement).dataset.index),
				element: element as HTMLElement
			}));
	}

	emitMountedPageWindowChange(pages: PaginatedPage[], payload: MountedWindowChange) {
		if (!this.options.onMountedPageWindowChange) {
			return;
		}

		const pageIndices = payload.indices.map(index => pages[index]?.pageIndex).filter(index => index != null);
		const signature = pageIndices.join(",");

		if (signature === this.lastMountedWindowSignature) {
			return;
		}

		this.lastMountedWindowSignature = signature;

		this.options.onMountedPageWindowChange({
			startPageIndex: pageIndices[0] ?? 0,
			endPageIndex: pageIndices[pageIndices.length - 1] ?? 0,
			pageIndices,
			addedPageIndices: payload.addedIndices.map(index => pages[index]?.pageIndex).filter(index => index != null),
			removedPageIndices: payload.removedIndices.map(index => pages[index]?.pageIndex).filter(index => index != null),
			pages: payload.items.map(item => ({
				pageIndex: pages[item.index]?.pageIndex,
				element: item.element
			})).filter(item => item.pageIndex != null),
			isScrolling: payload.isScrolling
		});
	}

	collectEndnoteIds(elements: OpenXmlElement[], output: string[]) {
		if (!elements)
			return;

		for (const element of elements) {
			if (element.type == DomType.EndnoteReference) {
				output.push((element as WmlNoteReference).id);
			}

			if (element.children?.length) {
				this.collectEndnoteIds(element.children, output);
			}
		}
	}

	estimatePageHeight(props: SectionProperties) {
		const defaultPageHeight = 1122;
		const pageHeight = parseSizeToPixels(props?.pageSize?.height) ?? defaultPageHeight;
		return pageHeight + (this.options.inWrapper ? 30 : 0);
	}

	optimizeChildren(children: OpenXmlElement[]) {
		const result: OpenXmlElement[] = [];

		for (const child of children) {
			const previous = result[result.length - 1];

			if (this.canMergeRuns(previous as WmlRun, child as WmlRun)) {
				for (const grandChild of child.children ?? []) {
					grandChild.parent = previous as WmlRun;
					(previous as WmlRun).children.push(grandChild);
				}
				continue;
			}

			if (this.canMergeText(previous as WmlText, child as WmlText)) {
				(previous as WmlText).text += (child as WmlText).text;
				continue;
			}

			result.push(child);
		}

		return result;
	}

	ensureOptimizedTree(element: OpenXmlElement) {
		if ((element as any).__optimizedRuns)
			return;

		if (element.children?.length) {
			element.children.forEach(child => this.ensureOptimizedTree(child));
			element.children = this.optimizeChildren(element.children);
			element.children.forEach(child => child.parent = element);
		}

		(element as any).__optimizedRuns = true;
	}

	canMergeRuns(left: WmlRun, right: WmlRun) {
		if (!left || !right || left.type != DomType.Run || right.type != DomType.Run)
			return false;

		if (left.fieldRun || right.fieldRun || left.id || right.id)
			return false;

		if (left.verticalAlign != right.verticalAlign || left.className != right.className || left.styleName != right.styleName)
			return false;

		if (!sameStyleMap(left.cssStyle, right.cssStyle))
			return false;

		return this.hasSimpleInlineChildren(left) && this.hasSimpleInlineChildren(right);
	}

	canMergeText(left: WmlText, right: WmlText) {
		return left?.type == DomType.Text && right?.type == DomType.Text;
	}

	hasSimpleInlineChildren(run: WmlRun) {
		return (run.children ?? []).every(child => simpleInlineChildTypes.has(child.type));
	}
}

type ChildType = Node | string;

function removeAllElements(elem: HTMLElement) {
	elem.innerHTML = '';
}

function appendChildren(elem: Node, children: (Node | string)[]) {
	const ownerDocument = elem.ownerDocument ?? document;
	children.forEach(c => elem.appendChild(isString(c) ? ownerDocument.createTextNode(c) : c));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	var parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}

function findScrollableElement(elem: HTMLElement, htmlDocument: Document): HTMLElement {
	const defaultView = htmlDocument.defaultView;
	let current: HTMLElement = elem;

	while (current) {
		const style = defaultView?.getComputedStyle(current);
		const overflowY = style?.overflowY ?? '';
		const overflow = style?.overflow ?? '';

		if (/(auto|scroll|overlay)/.test(`${overflowY} ${overflow}`)) {
			return current;
		}

		current = current.parentElement;
	}

	return null;
}

function parseSizeToPixels(value: string): number {
	if (!value)
		return null;

	const match = /^(-?\d*\.?\d+)(px|pt|pc|in|cm|mm|q)?$/i.exec(value.trim());

	if (!match)
		return null;

	const amount = parseFloat(match[1]);
	const unit = (match[2] ?? "px").toLowerCase();

	switch (unit) {
		case "pt": return amount * 96 / 72;
		case "pc": return amount * 16;
		case "in": return amount * 96;
		case "cm": return amount * 96 / 2.54;
		case "mm": return amount * 96 / 25.4;
		case "q": return amount * 96 / 101.6;
		default: return amount;
	}
}

function shouldInheritPadding(value: string) {
	if (value == null || value === "")
		return true;

	return /^0(?:\.0+)?(?:px|pt|pc|in|cm|mm|q)?$/i.test(value.trim());
}

const simpleInlineChildTypes = new Set([
	DomType.Text,
	DomType.DeletedText,
	DomType.Break,
	DomType.Tab,
	DomType.NoBreakHyphen,
	DomType.Symbol,
	DomType.FootnoteReference,
	DomType.EndnoteReference,
]);

function sameStyleMap(left: Record<string, string>, right: Record<string, string>) {
	const leftKeys = Object.keys(left ?? {});
	const rightKeys = Object.keys(right ?? {});

	if (leftKeys.length != rightKeys.length)
		return false;

	return leftKeys.every(key => left[key] == right[key]);
}
