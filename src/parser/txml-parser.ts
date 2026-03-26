import { parse as parseTxml } from 'txml/txml';

function removeUTF8BOM(data: string) {
    return data.charCodeAt(0) === 0xFEFF ? data.substring(1) : data;
}

type TxmlChild = TxmlElementNode | TxmlTextNode;

class TxmlTextNode {
    nodeType = 3;
    nodeName = "#text";
    localName = null;
    namespaceURI = null;
    childNodes: TxmlChild[] = [];
    firstChild = null;
    firstElementChild = null;

    constructor(public textContent: string) {
    }
}

class TxmlElementNode {
    nodeType = 1;
    localName: string;
    namespaceURI: string;
    childNodes: TxmlChild[];
    firstChild: TxmlChild = null;
    firstElementChild: TxmlElementNode = null;
    attributes: { name: string; localName: string; value: string }[];

    constructor(public nodeName: string, attributes: Record<string, string>, children: TxmlChild[], private namespaceMap: Record<string, string>) {
        const prefix = nodeName.includes(":") ? nodeName.split(":")[0] : "";
        this.localName = nodeName.split(":").pop();
        this.namespaceURI = namespaceMap[prefix] ?? null;
        this.attributes = Object.entries(attributes ?? {}).map(([name, value]) => ({
            name,
            localName: name.split(":").pop(),
            value
        }));
        this.childNodes = children;
        this.firstChild = children[0] ?? null;
        this.firstElementChild = children.find((child): child is TxmlElementNode => child.nodeType == 1) ?? null;
    }

    get textContent() {
        return this.childNodes.map(child => child.textContent ?? "").join("");
    }

    lookupNamespaceURI(prefix: string) {
        return this.namespaceMap[prefix ?? ""] ?? null;
    }
}

class TxmlDocumentNode {
    firstChild: TxmlChild = null;
    firstElementChild: TxmlElementNode = null;

    constructor(public childNodes: TxmlChild[]) {
        this.firstChild = childNodes[0] ?? null;
        this.firstElementChild = childNodes.find((child): child is TxmlElementNode => child.nodeType == 1) ?? null;
    }
}

export function parseXmlStringWithTxml(xmlString: string, trimXmlDeclaration: boolean = false) {
    if (trimXmlDeclaration)
        xmlString = xmlString.replace(/<[?].*[?]>/, "");

    xmlString = removeUTF8BOM(xmlString);

    const parsed = parseTxml(xmlString, { keepWhitespace: true });
    return new TxmlDocumentNode(adaptChildren(parsed));
}

function adaptChildren(children: any[], namespaceMap: Record<string, string> = {}): TxmlChild[] {
    return (children ?? []).map(child => adaptNode(child, namespaceMap));
}

function adaptNode(node: any, namespaceMap: Record<string, string>): TxmlChild {
    if (typeof node === "string") {
        return new TxmlTextNode(node);
    }

    const currentNamespaceMap = { ...namespaceMap, ...extractNamespaceMap(node.attributes) };
    return new TxmlElementNode(node.tagName, node.attributes, adaptChildren(node.children, currentNamespaceMap), currentNamespaceMap);
}

function extractNamespaceMap(attributes: Record<string, string>) {
    const result: Record<string, string> = {};

    for (const [name, value] of Object.entries(attributes ?? {})) {
        if (name === "xmlns") {
            result[""] = value;
        } else if (name.startsWith("xmlns:")) {
            result[name.substring(6)] = value;
        }
    }

    return result;
}
