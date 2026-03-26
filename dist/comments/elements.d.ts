import { DomType, OpenXmlElementBase } from "../document/dom";
export declare class WmlComment extends OpenXmlElementBase {
    type: DomType;
    id: string;
    author: string;
    initials: string;
    date: string;
}
export declare class WmlCommentReference extends OpenXmlElementBase {
    id?: string;
    type: DomType;
    constructor(id?: string);
}
export declare class WmlCommentRangeStart extends OpenXmlElementBase {
    id?: string;
    type: DomType;
    constructor(id?: string);
}
export declare class WmlCommentRangeEnd extends OpenXmlElementBase {
    id?: string;
    type: DomType;
    constructor(id?: string);
}
