import { WmlParagraph } from "./document/paragraph";
import { DocumentElement } from "./document/document";
import { FooterHeaderReference, SectionProperties, SectionType } from "./document/section";
import { DomType, OpenXmlElement, WmlBreak, WmlNoteReference, WmlText } from "./document/dom";
import { IDomStyle } from "./document/style";
import { escapeClassName, keyBy, mergeDeep } from "./utils";
import type { Options } from "./options";

export interface PaginatedSection {
    sectProps: SectionProperties;
    elements: OpenXmlElement[];
    pageBreak: boolean;
}

export interface PaginatedPage {
    pageIndex: number;
    key: number;
    sections: PaginatedSection[];
    pageProps: SectionProperties;
    footerProps: SectionProperties;
    firstOfSection: boolean;
    initialEndnoteIds: string[];
    estimatedHeight: number;
    isLastPage: boolean;
    headerRefs?: FooterHeaderReference[];
    footerRefs?: FooterHeaderReference[];
}

export interface DocumentPagingOptions extends Pick<Options,
    "breakPages" |
    "className" |
    "debug" |
    "ignoreLastRenderedPageBreak" |
    "inWrapper"
> {
}

export class DocumentPager {
    constructor(
        private options: DocumentPagingOptions,
        private styleMap: Record<string, IDomStyle> = {}
    ) {
    }

    static createStyleMap(styles: IDomStyle[], options: Pick<DocumentPagingOptions, "className" | "debug">) {
        const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);
        const className = options.className ?? "docx";

        for (const style of styles.filter(x => x.basedOn)) {
            const baseStyle = stylesMap[style.basedOn];

            if (baseStyle) {
                style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
                style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

                for (const baseValues of baseStyle.styles) {
                    const styleValues = style.styles.find(x => x.target == baseValues.target);

                    if (styleValues) {
                        copyStyleProperties(baseValues.values, styleValues.values);
                    } else {
                        style.styles.push({ ...baseValues, values: { ...baseValues.values } });
                    }
                }
            }
            else if (options.debug) {
                console.warn(`Can't find base style ${style.basedOn}`);
            }
        }

        for (const style of styles) {
            style.cssName = processStyleName(className, style.id);
        }

        return stylesMap;
    }

    buildPages(document: DocumentElement): PaginatedPage[] {
        const result: PaginatedPage[] = [];
        const allEndnoteIds: string[] = [];
        const sections = this.splitBySection(document.children, document.props);
        const pages = this.groupByPageBreaks(sections);
        let prevProps = null;

        for (let i = 0, l = pages.length; i < l; i++) {
            const pageSections = pages[i];
            const pageProps = pageSections[0].sectProps;
            let footerProps = pageProps;
            const initialEndnoteIds = allEndnoteIds.slice();

            for (const sect of pageSections) {
                this.collectEndnoteIds(sect.elements, allEndnoteIds);
                footerProps = sect.sectProps;
            }

            result.push({
                pageIndex: i,
                key: i,
                sections: pageSections,
                pageProps,
                footerProps,
                headerRefs: pageProps?.headerRefs,
                footerRefs: footerProps?.footerRefs,
                firstOfSection: prevProps != pageProps,
                initialEndnoteIds,
                estimatedHeight: this.estimatePageHeight(pageProps),
                isLastPage: i == l - 1
            });

            prevProps = footerProps;
        }

        return result;
    }

    findStyle(styleName: string) {
        return styleName && this.styleMap?.[styleName];
    }

    isPageBreakElement(elem: OpenXmlElement): boolean {
        if (elem.type != DomType.Break)
            return false;

        if ((elem as WmlBreak).break == "lastRenderedPageBreak")
            return !this.options.ignoreLastRenderedPageBreak;

        return (elem as WmlBreak).break == "page";
    }

    isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean {
        if (!prev || !next)
            return false;

        return prev.pageSize?.orientation != next.pageSize?.orientation
            || prev.pageSize?.width != next.pageSize?.width
            || prev.pageSize?.height != next.pageSize?.height;
    }

    splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): PaginatedSection[] {
        let current: PaginatedSection = { sectProps: null, elements: [], pageBreak: false };
        const result = [current];

        for (const elem of elements) {
            if (elem.type == DomType.Paragraph) {
                const style = this.findStyle((elem as WmlParagraph).styleName);

                if (style?.paragraphProps?.pageBreakBefore && current.elements.length > 0) {
                    current.pageBreak = true;
                    current = { sectProps: null, elements: [], pageBreak: false };
                    result.push(current);
                }
            }

            current.elements.push(elem);

            if (elem.type != DomType.Paragraph)
                continue;

            const paragraph = elem as WmlParagraph;
            const sectProps = paragraph.sectionProps;
            let paragraphBreakIndex = -1;
            let runBreakIndex = -1;

            if (this.options.breakPages && paragraph.children) {
                paragraphBreakIndex = paragraph.children.findIndex(run => {
                    runBreakIndex = run.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
                    return runBreakIndex != -1;
                });
            }

            if (sectProps || paragraphBreakIndex != -1) {
                current.sectProps = sectProps;
                current.pageBreak = paragraphBreakIndex != -1;
                current = { sectProps: null, elements: [], pageBreak: false };
                result.push(current);
            }

            if (paragraphBreakIndex == -1)
                continue;

            const breakRun = paragraph.children[paragraphBreakIndex];
            const splitRun = runBreakIndex < breakRun.children.length - 1;

            if (paragraphBreakIndex < paragraph.children.length - 1 || splitRun) {
                const children = elem.children;
                const newParagraph = { ...elem, children: children.slice(paragraphBreakIndex) };
                elem.children = children.slice(0, paragraphBreakIndex);
                current.elements.push(newParagraph);

                if (splitRun) {
                    const runChildren = breakRun.children;
                    const newRun = { ...breakRun, children: runChildren.slice(0, runBreakIndex) };
                    elem.children.push(newRun);
                    breakRun.children = runChildren.slice(runBreakIndex);
                }
            }
        }

        let currentSectProps = null;

        for (let i = result.length - 1; i >= 0; i--) {
            if (result[i].sectProps == null) {
                result[i].sectProps = currentSectProps ?? defaultProps;
            } else {
                currentSectProps = result[i].sectProps;
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

        for (const section of sections) {
            current.push(section);

            if (this.options.ignoreLastRenderedPageBreak || section.pageBreak || this.isPageBreakSection(prev, section.sectProps)) {
                result.push(current = []);
            }

            prev = section.sectProps;
        }

        return result.filter(x => x.length > 0);
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
}

function processStyleName(className: string, styleName: string): string {
    return styleName ? `${className}_${escapeClassName(styleName)}` : className;
}

function copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
    if (!input)
        return output;

    if (output == null) {
        output = {};
    }

    if (attrs == null) {
        attrs = Object.getOwnPropertyNames(input);
    }

    for (const key of attrs) {
        if (input.hasOwnProperty(key) && !output.hasOwnProperty(key)) {
            output[key] = input[key];
        }
    }

    return output;
}

function parseSizeToPixels(value: string) {
    if (!value)
        return null;

    if (value.endsWith("px"))
        return parseFloat(value);

    if (value.endsWith("pt"))
        return parseFloat(value) * (96 / 72);

    if (value.endsWith("cm"))
        return parseFloat(value) * (96 / 2.54);

    if (value.endsWith("mm"))
        return parseFloat(value) * (96 / 25.4);

    if (value.endsWith("in"))
        return parseFloat(value) * 96;

    return parseFloat(value);
}

