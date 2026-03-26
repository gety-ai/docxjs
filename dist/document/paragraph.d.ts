import { OpenXmlElement } from "./dom";
import { CommonProperties, Length } from "./common";
import { Borders } from "./border";
import { SectionProperties } from "./section";
import { LineSpacing } from "./line-spacing";
import { XmlParser } from "../parser/xml-parser";
import { RunProperties } from "./run";
export interface WmlParagraph extends OpenXmlElement, ParagraphProperties {
}
export interface ParagraphProperties extends CommonProperties {
    sectionProps: SectionProperties;
    tabs: ParagraphTab[];
    numbering: ParagraphNumbering;
    border: Borders;
    textAlignment: "auto" | "baseline" | "bottom" | "center" | "top" | string;
    lineSpacing: LineSpacing;
    keepLines: boolean;
    keepNext: boolean;
    pageBreakBefore: boolean;
    outlineLevel: number;
    styleName?: string;
    runProps: RunProperties;
}
export interface ParagraphTab {
    style: "bar" | "center" | "clear" | "decimal" | "end" | "num" | "start" | "left" | "right";
    leader: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore";
    position: Length;
}
export interface ParagraphNumbering {
    id: string;
    level: number;
}
export declare function parseParagraphProperties(elem: Element, xml: XmlParser): ParagraphProperties;
export declare function parseParagraphProperty(elem: Element, props: ParagraphProperties, xml: XmlParser): boolean;
export declare function parseTabs(elem: Element, xml: XmlParser): ParagraphTab[];
export declare function parseNumbering(elem: Element, xml: XmlParser): ParagraphNumbering;
