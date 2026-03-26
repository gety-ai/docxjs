import { Part } from "../common/part";
import { CustomProperty } from "./custom-props";
export declare class CustomPropsPart extends Part {
    props: CustomProperty[];
    parseXml(root: Element): void;
}
