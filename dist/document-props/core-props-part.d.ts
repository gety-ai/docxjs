import { Part } from "../common/part";
import { CorePropsDeclaration } from "./core-props";
export declare class CorePropsPart extends Part {
    props: CorePropsDeclaration;
    parseXml(root: Element): void;
}
