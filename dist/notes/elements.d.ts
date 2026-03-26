import { OpenXmlElementBase, DomType } from "../document/dom";
export declare abstract class WmlBaseNote implements OpenXmlElementBase {
    type: DomType;
    id: string;
    noteType: string;
}
export declare class WmlFootnote extends WmlBaseNote {
    type: DomType;
}
export declare class WmlEndnote extends WmlBaseNote {
    type: DomType;
}
