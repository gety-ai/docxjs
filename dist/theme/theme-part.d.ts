import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DmlTheme } from "./theme";
export declare class ThemePart extends Part {
    theme: DmlTheme;
    constructor(pkg: OpenXmlPackage, path: string);
    parseXml(root: Element): void;
}
