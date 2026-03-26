type TxmlChild = TxmlElementNode | TxmlTextNode;
declare class TxmlTextNode {
    textContent: string;
    nodeType: number;
    nodeName: string;
    localName: any;
    namespaceURI: any;
    childNodes: TxmlChild[];
    firstChild: any;
    firstElementChild: any;
    constructor(textContent: string);
}
declare class TxmlElementNode {
    nodeName: string;
    private namespaceMap;
    nodeType: number;
    localName: string;
    namespaceURI: string;
    childNodes: TxmlChild[];
    firstChild: TxmlChild;
    firstElementChild: TxmlElementNode;
    attributes: {
        name: string;
        localName: string;
        value: string;
    }[];
    constructor(nodeName: string, attributes: Record<string, string>, children: TxmlChild[], namespaceMap: Record<string, string>);
    get textContent(): string;
    lookupNamespaceURI(prefix: string): string;
}
declare class TxmlDocumentNode {
    childNodes: TxmlChild[];
    firstChild: TxmlChild;
    firstElementChild: TxmlElementNode;
    constructor(childNodes: TxmlChild[]);
}
export declare function parseXmlStringWithTxml(xmlString: string, trimXmlDeclaration?: boolean): TxmlDocumentNode;
export {};
