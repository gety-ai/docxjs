import { XmlParser } from "../parser/xml-parser";
export declare const ns: {
    wordml: string;
    drawingml: string;
    picture: string;
    compatibility: string;
    math: string;
};
export type LengthType = "px" | "pt" | "%" | "";
export type Length = string;
export interface Font {
    name: string;
    family: string;
}
export interface CommonProperties {
    fontSize: Length;
    color: string;
}
export type LengthUsageType = {
    mul: number;
    unit: LengthType;
    min?: number;
    max?: number;
};
export declare const LengthUsage: Record<string, LengthUsageType>;
export declare function convertLength(val: string, usage?: LengthUsageType): string;
export declare function convertBoolean(v: string, defaultValue?: boolean): boolean;
export declare function convertPercentage(val: string): number;
export declare function parseCommonProperty(elem: Element, props: CommonProperties, xml: XmlParser): boolean;
