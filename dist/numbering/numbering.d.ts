import { ParagraphProperties } from "../document/paragraph";
import { RunProperties } from "../document/run";
import { XmlParser } from "../parser/xml-parser";
export interface NumberingPartProperties {
    numberings: Numbering[];
    abstractNumberings: AbstractNumbering[];
    bulletPictures: NumberingBulletPicture[];
}
export interface Numbering {
    id: string;
    abstractId: string;
    overrides: NumberingLevelOverride[];
}
export interface NumberingLevelOverride {
    level: number;
    start: number;
    numberingLevel: NumberingLevel;
}
export interface AbstractNumbering {
    id: string;
    name: string;
    multiLevelType: "singleLevel" | "multiLevel" | "hybridMultilevel" | string;
    levels: NumberingLevel[];
    numberingStyleLink: string;
    styleLink: string;
}
export interface NumberingLevel {
    level: number;
    start: string;
    restart: number;
    format: 'lowerRoman' | 'lowerLetter' | string;
    text: string;
    justification: string;
    bulletPictureId: string;
    paragraphStyle: string;
    paragraphProps: ParagraphProperties;
    runProps: RunProperties;
}
export interface NumberingBulletPicture {
    id: string;
    referenceId: string;
    style: string;
}
export declare function parseNumberingPart(elem: Element, xml: XmlParser): NumberingPartProperties;
export declare function parseNumbering(elem: Element, xml: XmlParser): Numbering;
export declare function parseAbstractNumbering(elem: Element, xml: XmlParser): AbstractNumbering;
export declare function parseNumberingLevel(elem: Element, xml: XmlParser): NumberingLevel;
export declare function parseNumberingLevelOverrride(elem: Element, xml: XmlParser): NumberingLevelOverride;
export declare function parseNumberingBulletPicture(elem: Element, xml: XmlParser): NumberingBulletPicture;
