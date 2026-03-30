import { DocumentElement } from "./document/document";
import { FooterHeaderReference, SectionProperties } from "./document/section";
import { OpenXmlElement } from "./document/dom";
import { IDomStyle } from "./document/style";
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
export interface DocumentPagingOptions extends Pick<Options, "breakPages" | "className" | "debug" | "ignoreLastRenderedPageBreak" | "inWrapper"> {
}
export declare class DocumentPager {
    private options;
    private styleMap;
    constructor(options: DocumentPagingOptions, styleMap?: Record<string, IDomStyle>);
    static createStyleMap(styles: IDomStyle[], options: Pick<DocumentPagingOptions, "className" | "debug">): Record<any, IDomStyle>;
    buildPages(document: DocumentElement): PaginatedPage[];
    findStyle(styleName: string): IDomStyle;
    isPageBreakElement(elem: OpenXmlElement): boolean;
    isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean;
    splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): PaginatedSection[];
    resolveSectionProps(sections: PaginatedSection[]): PaginatedSection[];
    mergeSectionProps(base: SectionProperties, override: SectionProperties): SectionProperties;
    coalesceEmptySections(sections: PaginatedSection[]): PaginatedSection[];
    sectionForcesStandalonePage(section: PaginatedSection): boolean;
    sectionHasVisibleContent(section: PaginatedSection): boolean;
    elementHasVisibleContent(element: OpenXmlElement): any;
    groupByPageBreaks(sections: PaginatedSection[]): PaginatedSection[][];
    collectEndnoteIds(elements: OpenXmlElement[], output: string[]): void;
    estimatePageHeight(props: SectionProperties): number;
}
