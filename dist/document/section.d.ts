import { XmlParser } from "../parser/xml-parser";
import { Borders } from "./border";
import { Length } from "./common";
export interface Column {
    space: Length;
    width: Length;
}
export interface Columns {
    space: Length;
    numberOfColumns: number;
    separator: boolean;
    equalWidth: boolean;
    columns: Column[];
}
export interface PageSize {
    width: Length;
    height: Length;
    orientation: "landscape" | string;
}
export interface PageNumber {
    start: number;
    chapSep: "colon" | "emDash" | "endash" | "hyphen" | "period" | string;
    chapStyle: string;
    format: "none" | "cardinalText" | "decimal" | "decimalEnclosedCircle" | "decimalEnclosedFullstop" | "decimalEnclosedParen" | "decimalZero" | "lowerLetter" | "lowerRoman" | "ordinalText" | "upperLetter" | "upperRoman" | string;
}
export interface PageMargins {
    top: Length;
    right: Length;
    bottom: Length;
    left: Length;
    header: Length;
    footer: Length;
    gutter: Length;
}
export declare enum SectionType {
    Continuous = "continuous",
    NextPage = "nextPage",
    NextColumn = "nextColumn",
    EvenPage = "evenPage",
    OddPage = "oddPage"
}
export interface FooterHeaderReference {
    id: string;
    type: string | "first" | "even" | "default";
}
export interface SectionProperties {
    type: SectionType | string;
    pageSize: PageSize;
    pageMargins: PageMargins;
    pageBorders: Borders;
    pageNumber: PageNumber;
    columns: Columns;
    footerRefs: FooterHeaderReference[];
    headerRefs: FooterHeaderReference[];
    titlePage: boolean;
}
export declare function parseSectionProperties(elem: Element, xml?: XmlParser): SectionProperties;
