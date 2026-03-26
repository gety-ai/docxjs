import { Part } from "../common/part";
import { ExtendedPropsDeclaration } from "./extended-props";
export declare class ExtendedPropsPart extends Part {
    props: ExtendedPropsDeclaration;
    parseXml(root: Element): void;
}
