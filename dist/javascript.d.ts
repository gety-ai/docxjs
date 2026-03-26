import { Length } from "./document/common";
import { ParagraphTab } from "./document/paragraph";
export declare function computePixelToPoint(container?: HTMLElement): number;
export declare function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint?: number): void;
