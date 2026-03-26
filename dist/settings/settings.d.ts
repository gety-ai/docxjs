import { Length } from "../document/common";
import { XmlParser } from "../parser/xml-parser";
export interface WmlSettings {
    defaultTabStop: Length;
    footnoteProps: NoteProperties;
    endnoteProps: NoteProperties;
    autoHyphenation: boolean;
}
export interface NoteProperties {
    nummeringFormat: string;
    defaultNoteIds: string[];
}
export declare function parseSettings(elem: Element, xml: XmlParser): WmlSettings;
export declare function parseNoteProperties(elem: Element, xml: XmlParser): NoteProperties;
