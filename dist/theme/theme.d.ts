import { XmlParser } from "../parser/xml-parser";
export declare class DmlTheme {
    colorScheme: DmlColorScheme;
    fontScheme: DmlFontScheme;
}
export interface DmlColorScheme {
    name: string;
    colors: Record<string, string>;
}
export interface DmlFontScheme {
    name: string;
    majorFont: DmlFormInfo;
    minorFont: DmlFormInfo;
}
export interface DmlFormInfo {
    latinTypeface: string;
    eaTypeface: string;
    csTypeface: string;
}
export declare function parseTheme(elem: Element, xml: XmlParser): DmlTheme;
export declare function parseColorScheme(elem: Element, xml: XmlParser): DmlColorScheme;
export declare function parseFontScheme(elem: Element, xml: XmlParser): DmlFontScheme;
export declare function parseFontInfo(elem: Element, xml: XmlParser): DmlFormInfo;
