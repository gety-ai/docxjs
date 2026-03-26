import { Part } from "../common/part";
import { FontDeclaration } from "./fonts";
export declare class FontTablePart extends Part {
    fonts: FontDeclaration[];
    parseXml(root: Element): void;
}
