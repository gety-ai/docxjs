import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { WmlSettings } from "./settings";
export declare class SettingsPart extends Part {
    settings: WmlSettings;
    constructor(pkg: OpenXmlPackage, path: string);
    parseXml(root: Element): void;
}
