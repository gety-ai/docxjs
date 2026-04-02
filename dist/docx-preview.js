/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */
(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports) :
    typeof define === 'function' && define.amd ? define(['exports'], factory) :
    (global = typeof globalThis !== 'undefined' ? globalThis : global || self, factory(global.docx = {}));
})(this, (function (exports) { 'use strict';

    var RelationshipTypes;
    (function (RelationshipTypes) {
        RelationshipTypes["OfficeDocument"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        RelationshipTypes["FontTable"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
        RelationshipTypes["Image"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        RelationshipTypes["Numbering"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        RelationshipTypes["Styles"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        RelationshipTypes["StylesWithEffects"] = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
        RelationshipTypes["Theme"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        RelationshipTypes["Settings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        RelationshipTypes["WebSettings"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
        RelationshipTypes["Hyperlink"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        RelationshipTypes["Footnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        RelationshipTypes["Endnotes"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
        RelationshipTypes["Footer"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        RelationshipTypes["Header"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        RelationshipTypes["ExtendedProperties"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        RelationshipTypes["CoreProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        RelationshipTypes["CustomProperties"] = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties";
        RelationshipTypes["Comments"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        RelationshipTypes["CommentsExtended"] = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
        RelationshipTypes["AltChunk"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
    })(RelationshipTypes || (RelationshipTypes = {}));
    function parseRelationships(root, xml) {
        return xml.elements(root).map(e => ({
            id: xml.attr(e, "Id"),
            type: xml.attr(e, "Type"),
            target: xml.attr(e, "Target"),
            targetMode: xml.attr(e, "TargetMode")
        }));
    }

    function escapeClassName(className) {
        return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
    }
    function encloseFontFamily(fontFamily) {
        return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
    }
    function splitPath(path) {
        let si = path.lastIndexOf('/') + 1;
        let folder = si == 0 ? "" : path.substring(0, si);
        let fileName = si == 0 ? path : path.substring(si);
        return [folder, fileName];
    }
    function resolvePath(path, base) {
        try {
            const prefix = "http://docx/";
            const url = new URL(path, prefix + base).toString();
            return url.substring(prefix.length);
        }
        catch {
            return `${base}${path}`;
        }
    }
    function keyBy(array, by) {
        return array.reduce((a, x) => {
            a[by(x)] = x;
            return a;
        }, {});
    }
    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = () => reject();
            reader.readAsDataURL(blob);
        });
    }
    function isObject(item) {
        return item && typeof item === 'object' && !Array.isArray(item);
    }
    function isString(item) {
        return typeof item === 'string' || item instanceof String;
    }
    function mergeDeep(target, ...sources) {
        if (!sources.length)
            return target;
        const source = sources.shift();
        if (isObject(target) && isObject(source)) {
            for (const key in source) {
                if (isObject(source[key])) {
                    const val = target[key] ?? (target[key] = {});
                    mergeDeep(val, source[key]);
                }
                else {
                    target[key] = source[key];
                }
            }
        }
        return mergeDeep(target, ...sources);
    }
    function asArray(val) {
        return Array.isArray(val) ? val : [val];
    }
    function clamp(val, min, max) {
        return min > val ? min : (max < val ? max : val);
    }

    const ns$1 = {
        wordml: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"};
    const LengthUsage = {
        Dxa: { mul: 0.05, unit: "pt" },
        Emu: { mul: 1 / 12700, unit: "pt" },
        FontSize: { mul: 0.5, unit: "pt" },
        Border: { mul: 0.125, unit: "pt", min: 0.25, max: 12 },
        Point: { mul: 1, unit: "pt" },
        Percent: { mul: 0.02, unit: "%" }};
    function convertLength(val, usage = LengthUsage.Dxa) {
        if (val == null || /.+(p[xt]|[%])$/.test(val)) {
            return val;
        }
        var num = parseInt(val) * usage.mul;
        if (usage.min && usage.max)
            num = clamp(num, usage.min, usage.max);
        return `${num.toFixed(2)}${usage.unit}`;
    }
    function convertBoolean(v, defaultValue = false) {
        switch (v) {
            case "1": return true;
            case "0": return false;
            case "on": return true;
            case "off": return false;
            case "true": return true;
            case "false": return false;
            default: return defaultValue;
        }
    }
    function parseCommonProperty(elem, props, xml) {
        if (elem.namespaceURI != ns$1.wordml)
            return false;
        switch (elem.localName) {
            case "color":
                props.color = xml.attr(elem, "val");
                break;
            case "sz":
                props.fontSize = xml.lengthAttr(elem, "val", LengthUsage.FontSize);
                break;
            default:
                return false;
        }
        return true;
    }

    function parseXmlString(xmlString, trimXmlDeclaration = false) {
        if (trimXmlDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        xmlString = removeUTF8BOM$1(xmlString);
        const result = new DOMParser().parseFromString(xmlString, "application/xml");
        const errorText = hasXmlParserError(result);
        if (errorText)
            throw new Error(errorText);
        return result;
    }
    function hasXmlParserError(doc) {
        return doc.getElementsByTagName("parsererror")[0]?.textContent;
    }
    function removeUTF8BOM$1(data) {
        return data.charCodeAt(0) === 0xFEFF ? data.substring(1) : data;
    }
    function serializeXmlString(elem) {
        return new XMLSerializer().serializeToString(elem);
    }
    class XmlParser {
        elements(elem, localName = null) {
            const result = [];
            const childNodes = getChildNodes(elem);
            for (let i = 0, l = childNodes.length; i < l; i++) {
                let c = childNodes[i];
                if (isElementNode(c) && (localName == null || getLocalName(c) == localName))
                    result.push(c);
            }
            return result;
        }
        element(elem, localName) {
            const childNodes = getChildNodes(elem);
            for (let i = 0, l = childNodes.length; i < l; i++) {
                let c = childNodes[i];
                if (isElementNode(c) && getLocalName(c) == localName)
                    return c;
            }
            return null;
        }
        elementAttr(elem, localName, attrLocalName) {
            var el = this.element(elem, localName);
            return el ? this.attr(el, attrLocalName) : undefined;
        }
        attrs(elem) {
            return getAttributes(elem);
        }
        attr(elem, localName) {
            const attributes = getAttributes(elem);
            for (let i = 0, l = attributes.length; i < l; i++) {
                let a = attributes[i];
                if (a.localName == localName)
                    return a.value;
            }
            return null;
        }
        intAttr(node, attrName, defaultValue = null) {
            var val = this.attr(node, attrName);
            return val ? parseInt(val) : defaultValue;
        }
        hexAttr(node, attrName, defaultValue = null) {
            var val = this.attr(node, attrName);
            return val ? parseInt(val, 16) : defaultValue;
        }
        floatAttr(node, attrName, defaultValue = null) {
            var val = this.attr(node, attrName);
            return val ? parseFloat(val) : defaultValue;
        }
        boolAttr(node, attrName, defaultValue = null) {
            return convertBoolean(this.attr(node, attrName), defaultValue);
        }
        lengthAttr(node, attrName, usage = LengthUsage.Dxa) {
            return convertLength(this.attr(node, attrName), usage);
        }
    }
    const globalXmlParser = new XmlParser();
    function getChildNodes(elem) {
        if (!elem?.childNodes)
            return [];
        if (Array.isArray(elem.childNodes))
            return elem.childNodes;
        return Array.from(elem.childNodes);
    }
    function getAttributes(elem) {
        if (!elem?.attributes)
            return [];
        if (Array.isArray(elem.attributes))
            return elem.attributes;
        return Array.from(elem.attributes);
    }
    function isElementNode(node) {
        return node?.nodeType == 1 || (typeof node?.localName == "string" && typeof node?.nodeName == "string");
    }
    function getLocalName(node) {
        return node?.localName ?? node?.nodeName?.split?.(":")?.pop?.();
    }

    class Part {
        constructor(_package, path) {
            this._package = _package;
            this.path = path;
        }
        async load() {
            this.rels = await this._package.loadRelationships(this.path);
            const xmlText = await this._package.load(this.path);
            const xmlDoc = this._package.parseXmlDocument(xmlText);
            if (this._package.options.keepOrigin) {
                this._xmlDocument = xmlDoc;
            }
            this.parseXml(xmlDoc.firstElementChild);
        }
        save() {
            this._package.update(this.path, serializeXmlString(this._xmlDocument));
        }
        parseXml(root) {
        }
    }

    const embedFontTypeMap = {
        embedRegular: 'regular',
        embedBold: 'bold',
        embedItalic: 'italic',
        embedBoldItalic: 'boldItalic',
    };
    function parseFonts(root, xml) {
        return xml.elements(root).map(el => parseFont(el, xml));
    }
    function parseFont(elem, xml) {
        let result = {
            name: xml.attr(elem, "name"),
            embedFontRefs: []
        };
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "family":
                    result.family = xml.attr(el, "val");
                    break;
                case "altName":
                    result.altName = xml.attr(el, "val");
                    break;
                case "embedRegular":
                case "embedBold":
                case "embedItalic":
                case "embedBoldItalic":
                    result.embedFontRefs.push(parseEmbedFontRef(el, xml));
                    break;
            }
        }
        return result;
    }
    function parseEmbedFontRef(elem, xml) {
        return {
            id: xml.attr(elem, "id"),
            key: xml.attr(elem, "fontKey"),
            type: embedFontTypeMap[elem.localName]
        };
    }

    class FontTablePart extends Part {
        parseXml(root) {
            this.fonts = parseFonts(root, this._package.xmlParser);
        }
    }

    // DEFLATE is a complex format; to read this code, you should probably check the RFC first:
    // https://tools.ietf.org/html/rfc1951
    // You may also wish to take a look at the guide I made about this program:
    // https://gist.github.com/101arrowz/253f31eb5abc3d9275ab943003ffecad
    // Some of the following code is similar to that of UZIP.js:
    // https://github.com/photopea/UZIP.js
    // However, the vast majority of the codebase has diverged from UZIP.js to increase performance and reduce bundle size.
    // Sometimes 0 will appear where -1 would be more appropriate. This is because using a uint
    // is better for memory in most engines (I *think*).

    // aliases for shorter compressed code (most minifers don't do this)
    var u8 = Uint8Array, u16 = Uint16Array, i32 = Int32Array;
    // fixed length extra bits
    var fleb = new u8([0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0, /* unused */ 0, 0, /* impossible */ 0]);
    // fixed distance extra bits
    var fdeb = new u8([0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, /* unused */ 0, 0]);
    // code length index map
    var clim = new u8([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]);
    // get base, reverse index map from extra bits
    var freb = function (eb, start) {
        var b = new u16(31);
        for (var i = 0; i < 31; ++i) {
            b[i] = start += 1 << eb[i - 1];
        }
        // numbers here are at max 18 bits
        var r = new i32(b[30]);
        for (var i = 1; i < 30; ++i) {
            for (var j = b[i]; j < b[i + 1]; ++j) {
                r[j] = ((j - b[i]) << 5) | i;
            }
        }
        return { b: b, r: r };
    };
    var _a = freb(fleb, 2), fl = _a.b, revfl = _a.r;
    // we can ignore the fact that the other numbers are wrong; they never happen anyway
    fl[28] = 258, revfl[258] = 28;
    var _b = freb(fdeb, 0), fd = _b.b, revfd = _b.r;
    // map of value to reverse (assuming 16 bits)
    var rev = new u16(32768);
    for (var i = 0; i < 32768; ++i) {
        // reverse table algorithm from SO
        var x = ((i & 0xAAAA) >> 1) | ((i & 0x5555) << 1);
        x = ((x & 0xCCCC) >> 2) | ((x & 0x3333) << 2);
        x = ((x & 0xF0F0) >> 4) | ((x & 0x0F0F) << 4);
        rev[i] = (((x & 0xFF00) >> 8) | ((x & 0x00FF) << 8)) >> 1;
    }
    // create huffman tree from u8 "map": index -> code length for code index
    // mb (max bits) must be at most 15
    // TODO: optimize/split up?
    var hMap = (function (cd, mb, r) {
        var s = cd.length;
        // index
        var i = 0;
        // u16 "map": index -> # of codes with bit length = index
        var l = new u16(mb);
        // length of cd must be 288 (total # of codes)
        for (; i < s; ++i) {
            if (cd[i])
                ++l[cd[i] - 1];
        }
        // u16 "map": index -> minimum code for bit length = index
        var le = new u16(mb);
        for (i = 1; i < mb; ++i) {
            le[i] = (le[i - 1] + l[i - 1]) << 1;
        }
        var co;
        if (r) {
            // u16 "map": index -> number of actual bits, symbol for code
            co = new u16(1 << mb);
            // bits to remove for reverser
            var rvb = 15 - mb;
            for (i = 0; i < s; ++i) {
                // ignore 0 lengths
                if (cd[i]) {
                    // num encoding both symbol and bits read
                    var sv = (i << 4) | cd[i];
                    // free bits
                    var r_1 = mb - cd[i];
                    // start value
                    var v = le[cd[i] - 1]++ << r_1;
                    // m is end value
                    for (var m = v | ((1 << r_1) - 1); v <= m; ++v) {
                        // every 16 bit value starting with the code yields the same result
                        co[rev[v] >> rvb] = sv;
                    }
                }
            }
        }
        else {
            co = new u16(s);
            for (i = 0; i < s; ++i) {
                if (cd[i]) {
                    co[i] = rev[le[cd[i] - 1]++] >> (15 - cd[i]);
                }
            }
        }
        return co;
    });
    // fixed length tree
    var flt = new u8(288);
    for (var i = 0; i < 144; ++i)
        flt[i] = 8;
    for (var i = 144; i < 256; ++i)
        flt[i] = 9;
    for (var i = 256; i < 280; ++i)
        flt[i] = 7;
    for (var i = 280; i < 288; ++i)
        flt[i] = 8;
    // fixed distance tree
    var fdt = new u8(32);
    for (var i = 0; i < 32; ++i)
        fdt[i] = 5;
    // fixed length map
    var flm = /*#__PURE__*/ hMap(flt, 9, 0), flrm = /*#__PURE__*/ hMap(flt, 9, 1);
    // fixed distance map
    var fdm = /*#__PURE__*/ hMap(fdt, 5, 0), fdrm = /*#__PURE__*/ hMap(fdt, 5, 1);
    // find max of array
    var max = function (a) {
        var m = a[0];
        for (var i = 1; i < a.length; ++i) {
            if (a[i] > m)
                m = a[i];
        }
        return m;
    };
    // read d, starting at bit p and mask with m
    var bits = function (d, p, m) {
        var o = (p / 8) | 0;
        return ((d[o] | (d[o + 1] << 8)) >> (p & 7)) & m;
    };
    // read d, starting at bit p continuing for at least 16 bits
    var bits16 = function (d, p) {
        var o = (p / 8) | 0;
        return ((d[o] | (d[o + 1] << 8) | (d[o + 2] << 16)) >> (p & 7));
    };
    // get end of byte
    var shft = function (p) { return ((p + 7) / 8) | 0; };
    // typed array slice - allows garbage collector to free original reference,
    // while being more compatible than .slice
    var slc = function (v, s, e) {
        if (s == null || s < 0)
            s = 0;
        if (e == null || e > v.length)
            e = v.length;
        // can't use .constructor in case user-supplied
        return new u8(v.subarray(s, e));
    };
    // error codes
    var ec = [
        'unexpected EOF',
        'invalid block type',
        'invalid length/literal',
        'invalid distance',
        'stream finished',
        'no stream handler',
        ,
        'no callback',
        'invalid UTF-8 data',
        'extra field too long',
        'date not in range 1980-2099',
        'filename too long',
        'stream finishing',
        'invalid zip data'
        // determined by unknown compression method
    ];
    var err = function (ind, msg, nt) {
        var e = new Error(msg || ec[ind]);
        e.code = ind;
        if (Error.captureStackTrace)
            Error.captureStackTrace(e, err);
        if (!nt)
            throw e;
        return e;
    };
    // expands raw DEFLATE data
    var inflt = function (dat, st, buf, dict) {
        // source length       dict length
        var sl = dat.length, dl = dict ? dict.length : 0;
        if (!sl || st.f && !st.l)
            return buf || new u8(0);
        var noBuf = !buf;
        // have to estimate size
        var resize = noBuf || st.i != 2;
        // no state
        var noSt = st.i;
        // Assumes roughly 33% compression ratio average
        if (noBuf)
            buf = new u8(sl * 3);
        // ensure buffer can fit at least l elements
        var cbuf = function (l) {
            var bl = buf.length;
            // need to increase size to fit
            if (l > bl) {
                // Double or set to necessary, whichever is greater
                var nbuf = new u8(Math.max(bl * 2, l));
                nbuf.set(buf);
                buf = nbuf;
            }
        };
        //  last chunk         bitpos           bytes
        var final = st.f || 0, pos = st.p || 0, bt = st.b || 0, lm = st.l, dm = st.d, lbt = st.m, dbt = st.n;
        // total bits
        var tbts = sl * 8;
        do {
            if (!lm) {
                // BFINAL - this is only 1 when last chunk is next
                final = bits(dat, pos, 1);
                // type: 0 = no compression, 1 = fixed huffman, 2 = dynamic huffman
                var type = bits(dat, pos + 1, 3);
                pos += 3;
                if (!type) {
                    // go to end of byte boundary
                    var s = shft(pos) + 4, l = dat[s - 4] | (dat[s - 3] << 8), t = s + l;
                    if (t > sl) {
                        if (noSt)
                            err(0);
                        break;
                    }
                    // ensure size
                    if (resize)
                        cbuf(bt + l);
                    // Copy over uncompressed data
                    buf.set(dat.subarray(s, t), bt);
                    // Get new bitpos, update byte count
                    st.b = bt += l, st.p = pos = t * 8, st.f = final;
                    continue;
                }
                else if (type == 1)
                    lm = flrm, dm = fdrm, lbt = 9, dbt = 5;
                else if (type == 2) {
                    //  literal                            lengths
                    var hLit = bits(dat, pos, 31) + 257, hcLen = bits(dat, pos + 10, 15) + 4;
                    var tl = hLit + bits(dat, pos + 5, 31) + 1;
                    pos += 14;
                    // length+distance tree
                    var ldt = new u8(tl);
                    // code length tree
                    var clt = new u8(19);
                    for (var i = 0; i < hcLen; ++i) {
                        // use index map to get real code
                        clt[clim[i]] = bits(dat, pos + i * 3, 7);
                    }
                    pos += hcLen * 3;
                    // code lengths bits
                    var clb = max(clt), clbmsk = (1 << clb) - 1;
                    // code lengths map
                    var clm = hMap(clt, clb, 1);
                    for (var i = 0; i < tl;) {
                        var r = clm[bits(dat, pos, clbmsk)];
                        // bits read
                        pos += r & 15;
                        // symbol
                        var s = r >> 4;
                        // code length to copy
                        if (s < 16) {
                            ldt[i++] = s;
                        }
                        else {
                            //  copy   count
                            var c = 0, n = 0;
                            if (s == 16)
                                n = 3 + bits(dat, pos, 3), pos += 2, c = ldt[i - 1];
                            else if (s == 17)
                                n = 3 + bits(dat, pos, 7), pos += 3;
                            else if (s == 18)
                                n = 11 + bits(dat, pos, 127), pos += 7;
                            while (n--)
                                ldt[i++] = c;
                        }
                    }
                    //    length tree                 distance tree
                    var lt = ldt.subarray(0, hLit), dt = ldt.subarray(hLit);
                    // max length bits
                    lbt = max(lt);
                    // max dist bits
                    dbt = max(dt);
                    lm = hMap(lt, lbt, 1);
                    dm = hMap(dt, dbt, 1);
                }
                else
                    err(1);
                if (pos > tbts) {
                    if (noSt)
                        err(0);
                    break;
                }
            }
            // Make sure the buffer can hold this + the largest possible addition
            // Maximum chunk size (practically, theoretically infinite) is 2^17
            if (resize)
                cbuf(bt + 131072);
            var lms = (1 << lbt) - 1, dms = (1 << dbt) - 1;
            var lpos = pos;
            for (;; lpos = pos) {
                // bits read, code
                var c = lm[bits16(dat, pos) & lms], sym = c >> 4;
                pos += c & 15;
                if (pos > tbts) {
                    if (noSt)
                        err(0);
                    break;
                }
                if (!c)
                    err(2);
                if (sym < 256)
                    buf[bt++] = sym;
                else if (sym == 256) {
                    lpos = pos, lm = null;
                    break;
                }
                else {
                    var add = sym - 254;
                    // no extra bits needed if less
                    if (sym > 264) {
                        // index
                        var i = sym - 257, b = fleb[i];
                        add = bits(dat, pos, (1 << b) - 1) + fl[i];
                        pos += b;
                    }
                    // dist
                    var d = dm[bits16(dat, pos) & dms], dsym = d >> 4;
                    if (!d)
                        err(3);
                    pos += d & 15;
                    var dt = fd[dsym];
                    if (dsym > 3) {
                        var b = fdeb[dsym];
                        dt += bits16(dat, pos) & (1 << b) - 1, pos += b;
                    }
                    if (pos > tbts) {
                        if (noSt)
                            err(0);
                        break;
                    }
                    if (resize)
                        cbuf(bt + 131072);
                    var end = bt + add;
                    if (bt < dt) {
                        var shift = dl - dt, dend = Math.min(dt, end);
                        if (shift + bt < 0)
                            err(3);
                        for (; bt < dend; ++bt)
                            buf[bt] = dict[shift + bt];
                    }
                    for (; bt < end; ++bt)
                        buf[bt] = buf[bt - dt];
                }
            }
            st.l = lm, st.p = lpos, st.b = bt, st.f = final;
            if (lm)
                final = 1, st.m = lbt, st.d = dm, st.n = dbt;
        } while (!final);
        // don't reallocate for streams or user buffers
        return bt != buf.length && noBuf ? slc(buf, 0, bt) : buf.subarray(0, bt);
    };
    // starting at p, write the minimum number of bits that can hold v to d
    var wbits = function (d, p, v) {
        v <<= p & 7;
        var o = (p / 8) | 0;
        d[o] |= v;
        d[o + 1] |= v >> 8;
    };
    // starting at p, write the minimum number of bits (>8) that can hold v to d
    var wbits16 = function (d, p, v) {
        v <<= p & 7;
        var o = (p / 8) | 0;
        d[o] |= v;
        d[o + 1] |= v >> 8;
        d[o + 2] |= v >> 16;
    };
    // creates code lengths from a frequency table
    var hTree = function (d, mb) {
        // Need extra info to make a tree
        var t = [];
        for (var i = 0; i < d.length; ++i) {
            if (d[i])
                t.push({ s: i, f: d[i] });
        }
        var s = t.length;
        var t2 = t.slice();
        if (!s)
            return { t: et, l: 0 };
        if (s == 1) {
            var v = new u8(t[0].s + 1);
            v[t[0].s] = 1;
            return { t: v, l: 1 };
        }
        t.sort(function (a, b) { return a.f - b.f; });
        // after i2 reaches last ind, will be stopped
        // freq must be greater than largest possible number of symbols
        t.push({ s: -1, f: 25001 });
        var l = t[0], r = t[1], i0 = 0, i1 = 1, i2 = 2;
        t[0] = { s: -1, f: l.f + r.f, l: l, r: r };
        // efficient algorithm from UZIP.js
        // i0 is lookbehind, i2 is lookahead - after processing two low-freq
        // symbols that combined have high freq, will start processing i2 (high-freq,
        // non-composite) symbols instead
        // see https://reddit.com/r/photopea/comments/ikekht/uzipjs_questions/
        while (i1 != s - 1) {
            l = t[t[i0].f < t[i2].f ? i0++ : i2++];
            r = t[i0 != i1 && t[i0].f < t[i2].f ? i0++ : i2++];
            t[i1++] = { s: -1, f: l.f + r.f, l: l, r: r };
        }
        var maxSym = t2[0].s;
        for (var i = 1; i < s; ++i) {
            if (t2[i].s > maxSym)
                maxSym = t2[i].s;
        }
        // code lengths
        var tr = new u16(maxSym + 1);
        // max bits in tree
        var mbt = ln(t[i1 - 1], tr, 0);
        if (mbt > mb) {
            // more algorithms from UZIP.js
            // TODO: find out how this code works (debt)
            //  ind    debt
            var i = 0, dt = 0;
            //    left            cost
            var lft = mbt - mb, cst = 1 << lft;
            t2.sort(function (a, b) { return tr[b.s] - tr[a.s] || a.f - b.f; });
            for (; i < s; ++i) {
                var i2_1 = t2[i].s;
                if (tr[i2_1] > mb) {
                    dt += cst - (1 << (mbt - tr[i2_1]));
                    tr[i2_1] = mb;
                }
                else
                    break;
            }
            dt >>= lft;
            while (dt > 0) {
                var i2_2 = t2[i].s;
                if (tr[i2_2] < mb)
                    dt -= 1 << (mb - tr[i2_2]++ - 1);
                else
                    ++i;
            }
            for (; i >= 0 && dt; --i) {
                var i2_3 = t2[i].s;
                if (tr[i2_3] == mb) {
                    --tr[i2_3];
                    ++dt;
                }
            }
            mbt = mb;
        }
        return { t: new u8(tr), l: mbt };
    };
    // get the max length and assign length codes
    var ln = function (n, l, d) {
        return n.s == -1
            ? Math.max(ln(n.l, l, d + 1), ln(n.r, l, d + 1))
            : (l[n.s] = d);
    };
    // length codes generation
    var lc = function (c) {
        var s = c.length;
        // Note that the semicolon was intentional
        while (s && !c[--s])
            ;
        var cl = new u16(++s);
        //  ind      num         streak
        var cli = 0, cln = c[0], cls = 1;
        var w = function (v) { cl[cli++] = v; };
        for (var i = 1; i <= s; ++i) {
            if (c[i] == cln && i != s)
                ++cls;
            else {
                if (!cln && cls > 2) {
                    for (; cls > 138; cls -= 138)
                        w(32754);
                    if (cls > 2) {
                        w(cls > 10 ? ((cls - 11) << 5) | 28690 : ((cls - 3) << 5) | 12305);
                        cls = 0;
                    }
                }
                else if (cls > 3) {
                    w(cln), --cls;
                    for (; cls > 6; cls -= 6)
                        w(8304);
                    if (cls > 2)
                        w(((cls - 3) << 5) | 8208), cls = 0;
                }
                while (cls--)
                    w(cln);
                cls = 1;
                cln = c[i];
            }
        }
        return { c: cl.subarray(0, cli), n: s };
    };
    // calculate the length of output from tree, code lengths
    var clen = function (cf, cl) {
        var l = 0;
        for (var i = 0; i < cl.length; ++i)
            l += cf[i] * cl[i];
        return l;
    };
    // writes a fixed block
    // returns the new bit pos
    var wfblk = function (out, pos, dat) {
        // no need to write 00 as type: TypedArray defaults to 0
        var s = dat.length;
        var o = shft(pos + 2);
        out[o] = s & 255;
        out[o + 1] = s >> 8;
        out[o + 2] = out[o] ^ 255;
        out[o + 3] = out[o + 1] ^ 255;
        for (var i = 0; i < s; ++i)
            out[o + i + 4] = dat[i];
        return (o + 4 + s) * 8;
    };
    // writes a block
    var wblk = function (dat, out, final, syms, lf, df, eb, li, bs, bl, p) {
        wbits(out, p++, final);
        ++lf[256];
        var _a = hTree(lf, 15), dlt = _a.t, mlb = _a.l;
        var _b = hTree(df, 15), ddt = _b.t, mdb = _b.l;
        var _c = lc(dlt), lclt = _c.c, nlc = _c.n;
        var _d = lc(ddt), lcdt = _d.c, ndc = _d.n;
        var lcfreq = new u16(19);
        for (var i = 0; i < lclt.length; ++i)
            ++lcfreq[lclt[i] & 31];
        for (var i = 0; i < lcdt.length; ++i)
            ++lcfreq[lcdt[i] & 31];
        var _e = hTree(lcfreq, 7), lct = _e.t, mlcb = _e.l;
        var nlcc = 19;
        for (; nlcc > 4 && !lct[clim[nlcc - 1]]; --nlcc)
            ;
        var flen = (bl + 5) << 3;
        var ftlen = clen(lf, flt) + clen(df, fdt) + eb;
        var dtlen = clen(lf, dlt) + clen(df, ddt) + eb + 14 + 3 * nlcc + clen(lcfreq, lct) + 2 * lcfreq[16] + 3 * lcfreq[17] + 7 * lcfreq[18];
        if (bs >= 0 && flen <= ftlen && flen <= dtlen)
            return wfblk(out, p, dat.subarray(bs, bs + bl));
        var lm, ll, dm, dl;
        wbits(out, p, 1 + (dtlen < ftlen)), p += 2;
        if (dtlen < ftlen) {
            lm = hMap(dlt, mlb, 0), ll = dlt, dm = hMap(ddt, mdb, 0), dl = ddt;
            var llm = hMap(lct, mlcb, 0);
            wbits(out, p, nlc - 257);
            wbits(out, p + 5, ndc - 1);
            wbits(out, p + 10, nlcc - 4);
            p += 14;
            for (var i = 0; i < nlcc; ++i)
                wbits(out, p + 3 * i, lct[clim[i]]);
            p += 3 * nlcc;
            var lcts = [lclt, lcdt];
            for (var it = 0; it < 2; ++it) {
                var clct = lcts[it];
                for (var i = 0; i < clct.length; ++i) {
                    var len = clct[i] & 31;
                    wbits(out, p, llm[len]), p += lct[len];
                    if (len > 15)
                        wbits(out, p, (clct[i] >> 5) & 127), p += clct[i] >> 12;
                }
            }
        }
        else {
            lm = flm, ll = flt, dm = fdm, dl = fdt;
        }
        for (var i = 0; i < li; ++i) {
            var sym = syms[i];
            if (sym > 255) {
                var len = (sym >> 18) & 31;
                wbits16(out, p, lm[len + 257]), p += ll[len + 257];
                if (len > 7)
                    wbits(out, p, (sym >> 23) & 31), p += fleb[len];
                var dst = sym & 31;
                wbits16(out, p, dm[dst]), p += dl[dst];
                if (dst > 3)
                    wbits16(out, p, (sym >> 5) & 8191), p += fdeb[dst];
            }
            else {
                wbits16(out, p, lm[sym]), p += ll[sym];
            }
        }
        wbits16(out, p, lm[256]);
        return p + ll[256];
    };
    // deflate options (nice << 13) | chain
    var deo = /*#__PURE__*/ new i32([65540, 131080, 131088, 131104, 262176, 1048704, 1048832, 2114560, 2117632]);
    // empty
    var et = /*#__PURE__*/ new u8(0);
    // compresses data into a raw DEFLATE buffer
    var dflt = function (dat, lvl, plvl, pre, post, st) {
        var s = st.z || dat.length;
        var o = new u8(pre + s + 5 * (1 + Math.ceil(s / 7000)) + post);
        // writing to this writes to the output buffer
        var w = o.subarray(pre, o.length - post);
        var lst = st.l;
        var pos = (st.r || 0) & 7;
        if (lvl) {
            if (pos)
                w[0] = st.r >> 3;
            var opt = deo[lvl - 1];
            var n = opt >> 13, c = opt & 8191;
            var msk_1 = (1 << plvl) - 1;
            //    prev 2-byte val map    curr 2-byte val map
            var prev = st.p || new u16(32768), head = st.h || new u16(msk_1 + 1);
            var bs1_1 = Math.ceil(plvl / 3), bs2_1 = 2 * bs1_1;
            var hsh = function (i) { return (dat[i] ^ (dat[i + 1] << bs1_1) ^ (dat[i + 2] << bs2_1)) & msk_1; };
            // 24576 is an arbitrary number of maximum symbols per block
            // 424 buffer for last block
            var syms = new i32(25000);
            // length/literal freq   distance freq
            var lf = new u16(288), df = new u16(32);
            //  l/lcnt  exbits  index          l/lind  waitdx          blkpos
            var lc_1 = 0, eb = 0, i = st.i || 0, li = 0, wi = st.w || 0, bs = 0;
            for (; i + 2 < s; ++i) {
                // hash value
                var hv = hsh(i);
                // index mod 32768    previous index mod
                var imod = i & 32767, pimod = head[hv];
                prev[imod] = pimod;
                head[hv] = imod;
                // We always should modify head and prev, but only add symbols if
                // this data is not yet processed ("wait" for wait index)
                if (wi <= i) {
                    // bytes remaining
                    var rem = s - i;
                    if ((lc_1 > 7000 || li > 24576) && (rem > 423 || !lst)) {
                        pos = wblk(dat, w, 0, syms, lf, df, eb, li, bs, i - bs, pos);
                        li = lc_1 = eb = 0, bs = i;
                        for (var j = 0; j < 286; ++j)
                            lf[j] = 0;
                        for (var j = 0; j < 30; ++j)
                            df[j] = 0;
                    }
                    //  len    dist   chain
                    var l = 2, d = 0, ch_1 = c, dif = imod - pimod & 32767;
                    if (rem > 2 && hv == hsh(i - dif)) {
                        var maxn = Math.min(n, rem) - 1;
                        var maxd = Math.min(32767, i);
                        // max possible length
                        // not capped at dif because decompressors implement "rolling" index population
                        var ml = Math.min(258, rem);
                        while (dif <= maxd && --ch_1 && imod != pimod) {
                            if (dat[i + l] == dat[i + l - dif]) {
                                var nl = 0;
                                for (; nl < ml && dat[i + nl] == dat[i + nl - dif]; ++nl)
                                    ;
                                if (nl > l) {
                                    l = nl, d = dif;
                                    // break out early when we reach "nice" (we are satisfied enough)
                                    if (nl > maxn)
                                        break;
                                    // now, find the rarest 2-byte sequence within this
                                    // length of literals and search for that instead.
                                    // Much faster than just using the start
                                    var mmd = Math.min(dif, nl - 2);
                                    var md = 0;
                                    for (var j = 0; j < mmd; ++j) {
                                        var ti = i - dif + j & 32767;
                                        var pti = prev[ti];
                                        var cd = ti - pti & 32767;
                                        if (cd > md)
                                            md = cd, pimod = ti;
                                    }
                                }
                            }
                            // check the previous match
                            imod = pimod, pimod = prev[imod];
                            dif += imod - pimod & 32767;
                        }
                    }
                    // d will be nonzero only when a match was found
                    if (d) {
                        // store both dist and len data in one int32
                        // Make sure this is recognized as a len/dist with 28th bit (2^28)
                        syms[li++] = 268435456 | (revfl[l] << 18) | revfd[d];
                        var lin = revfl[l] & 31, din = revfd[d] & 31;
                        eb += fleb[lin] + fdeb[din];
                        ++lf[257 + lin];
                        ++df[din];
                        wi = i + l;
                        ++lc_1;
                    }
                    else {
                        syms[li++] = dat[i];
                        ++lf[dat[i]];
                    }
                }
            }
            for (i = Math.max(i, wi); i < s; ++i) {
                syms[li++] = dat[i];
                ++lf[dat[i]];
            }
            pos = wblk(dat, w, lst, syms, lf, df, eb, li, bs, i - bs, pos);
            if (!lst) {
                st.r = (pos & 7) | w[(pos / 8) | 0] << 3;
                // shft(pos) now 1 less if pos & 7 != 0
                pos -= 7;
                st.h = head, st.p = prev, st.i = i, st.w = wi;
            }
        }
        else {
            for (var i = st.w || 0; i < s + lst; i += 65535) {
                // end
                var e = i + 65535;
                if (e >= s) {
                    // write final block
                    w[(pos / 8) | 0] = lst;
                    e = s;
                }
                pos = wfblk(w, pos + 1, dat.subarray(i, e));
            }
            st.i = s;
        }
        return slc(o, 0, pre + shft(pos) + post);
    };
    // CRC32 table
    var crct = /*#__PURE__*/ (function () {
        var t = new Int32Array(256);
        for (var i = 0; i < 256; ++i) {
            var c = i, k = 9;
            while (--k)
                c = ((c & 1) && -306674912) ^ (c >>> 1);
            t[i] = c;
        }
        return t;
    })();
    // CRC32
    var crc = function () {
        var c = -1;
        return {
            p: function (d) {
                // closures have awful performance
                var cr = c;
                for (var i = 0; i < d.length; ++i)
                    cr = crct[(cr & 255) ^ d[i]] ^ (cr >>> 8);
                c = cr;
            },
            d: function () { return ~c; }
        };
    };
    // deflate with opts
    var dopt = function (dat, opt, pre, post, st) {
        if (!st) {
            st = { l: 1 };
            if (opt.dictionary) {
                var dict = opt.dictionary.subarray(-32768);
                var newDat = new u8(dict.length + dat.length);
                newDat.set(dict);
                newDat.set(dat, dict.length);
                dat = newDat;
                st.w = dict.length;
            }
        }
        return dflt(dat, opt.level == null ? 6 : opt.level, opt.mem == null ? (st.l ? Math.ceil(Math.max(8, Math.min(13, Math.log(dat.length))) * 1.5) : 20) : (12 + opt.mem), pre, post, st);
    };
    // Walmart object spread
    var mrg = function (a, b) {
        var o = {};
        for (var k in a)
            o[k] = a[k];
        for (var k in b)
            o[k] = b[k];
        return o;
    };
    // read 2 bytes
    var b2 = function (d, b) { return d[b] | (d[b + 1] << 8); };
    // read 4 bytes
    var b4 = function (d, b) { return (d[b] | (d[b + 1] << 8) | (d[b + 2] << 16) | (d[b + 3] << 24)) >>> 0; };
    var b8 = function (d, b) { return b4(d, b) + (b4(d, b + 4) * 4294967296); };
    // write bytes
    var wbytes = function (d, b, v) {
        for (; v; ++b)
            d[b] = v, v >>>= 8;
    };
    /**
     * Compresses data with DEFLATE without any wrapper
     * @param data The data to compress
     * @param opts The compression options
     * @returns The deflated version of the data
     */
    function deflateSync(data, opts) {
        return dopt(data, opts || {}, 0, 0);
    }
    /**
     * Expands DEFLATE data with no wrapper
     * @param data The data to decompress
     * @param opts The decompression options
     * @returns The decompressed version of the data
     */
    function inflateSync(data, opts) {
        return inflt(data, { i: 2 }, opts && opts.out, opts && opts.dictionary);
    }
    // flatten a directory structure
    var fltn = function (d, p, t, o) {
        for (var k in d) {
            var val = d[k], n = p + k, op = o;
            if (Array.isArray(val))
                op = mrg(o, val[1]), val = val[0];
            if (val instanceof u8)
                t[n] = [val, op];
            else {
                t[n += '/'] = [new u8(0), op];
                fltn(val, n, t, o);
            }
        }
    };
    // text encoder
    var te = typeof TextEncoder != 'undefined' && /*#__PURE__*/ new TextEncoder();
    // text decoder
    var td = typeof TextDecoder != 'undefined' && /*#__PURE__*/ new TextDecoder();
    // text decoder stream
    var tds = 0;
    try {
        td.decode(et, { stream: true });
        tds = 1;
    }
    catch (e) { }
    // decode UTF8
    var dutf8 = function (d) {
        for (var r = '', i = 0;;) {
            var c = d[i++];
            var eb = (c > 127) + (c > 223) + (c > 239);
            if (i + eb > d.length)
                return { s: r, r: slc(d, i - 1) };
            if (!eb)
                r += String.fromCharCode(c);
            else if (eb == 3) {
                c = ((c & 15) << 18 | (d[i++] & 63) << 12 | (d[i++] & 63) << 6 | (d[i++] & 63)) - 65536,
                    r += String.fromCharCode(55296 | (c >> 10), 56320 | (c & 1023));
            }
            else if (eb & 1)
                r += String.fromCharCode((c & 31) << 6 | (d[i++] & 63));
            else
                r += String.fromCharCode((c & 15) << 12 | (d[i++] & 63) << 6 | (d[i++] & 63));
        }
    };
    /**
     * Converts a string into a Uint8Array for use with compression/decompression methods
     * @param str The string to encode
     * @param latin1 Whether or not to interpret the data as Latin-1. This should
     *               not need to be true unless decoding a binary string.
     * @returns The string encoded in UTF-8/Latin-1 binary
     */
    function strToU8(str, latin1) {
        var i; 
        if (te)
            return te.encode(str);
        var l = str.length;
        var ar = new u8(str.length + (str.length >> 1));
        var ai = 0;
        var w = function (v) { ar[ai++] = v; };
        for (var i = 0; i < l; ++i) {
            if (ai + 5 > ar.length) {
                var n = new u8(ai + 8 + ((l - i) << 1));
                n.set(ar);
                ar = n;
            }
            var c = str.charCodeAt(i);
            if (c < 128 || latin1)
                w(c);
            else if (c < 2048)
                w(192 | (c >> 6)), w(128 | (c & 63));
            else if (c > 55295 && c < 57344)
                c = 65536 + (c & 1023 << 10) | (str.charCodeAt(++i) & 1023),
                    w(240 | (c >> 18)), w(128 | ((c >> 12) & 63)), w(128 | ((c >> 6) & 63)), w(128 | (c & 63));
            else
                w(224 | (c >> 12)), w(128 | ((c >> 6) & 63)), w(128 | (c & 63));
        }
        return slc(ar, 0, ai);
    }
    /**
     * Converts a Uint8Array to a string
     * @param dat The data to decode to string
     * @param latin1 Whether or not to interpret the data as Latin-1. This should
     *               not need to be true unless encoding to binary string.
     * @returns The original UTF-8/Latin-1 string
     */
    function strFromU8(dat, latin1) {
        if (latin1) {
            var r = '';
            for (var i = 0; i < dat.length; i += 16384)
                r += String.fromCharCode.apply(null, dat.subarray(i, i + 16384));
            return r;
        }
        else if (td) {
            return td.decode(dat);
        }
        else {
            var _a = dutf8(dat), s = _a.s, r = _a.r;
            if (r.length)
                err(8);
            return s;
        }
    }
    // skip local zip header
    var slzh = function (d, b) { return b + 30 + b2(d, b + 26) + b2(d, b + 28); };
    // read zip header
    var zh = function (d, b, z) {
        var fnl = b2(d, b + 28), fn = strFromU8(d.subarray(b + 46, b + 46 + fnl), !(b2(d, b + 8) & 2048)), es = b + 46 + fnl, bs = b4(d, b + 20);
        var _a = z && bs == 4294967295 ? z64e(d, es) : [bs, b4(d, b + 24), b4(d, b + 42)], sc = _a[0], su = _a[1], off = _a[2];
        return [b2(d, b + 10), sc, su, fn, es + b2(d, b + 30) + b2(d, b + 32), off];
    };
    // read zip64 extra field
    var z64e = function (d, b) {
        for (; b2(d, b) != 1; b += 4 + b2(d, b + 2))
            ;
        return [b8(d, b + 12), b8(d, b + 4), b8(d, b + 20)];
    };
    // extra field length
    var exfl = function (ex) {
        var le = 0;
        if (ex) {
            for (var k in ex) {
                var l = ex[k].length;
                if (l > 65535)
                    err(9);
                le += l + 4;
            }
        }
        return le;
    };
    // write zip header
    var wzh = function (d, b, f, fn, u, c, ce, co) {
        var fl = fn.length, ex = f.extra, col = co && co.length;
        var exl = exfl(ex);
        wbytes(d, b, ce != null ? 0x2014B50 : 0x4034B50), b += 4;
        if (ce != null)
            d[b++] = 20, d[b++] = f.os;
        d[b] = 20, b += 2; // spec compliance? what's that?
        d[b++] = (f.flag << 1) | (c < 0 && 8), d[b++] = u && 8;
        d[b++] = f.compression & 255, d[b++] = f.compression >> 8;
        var dt = new Date(f.mtime == null ? Date.now() : f.mtime), y = dt.getFullYear() - 1980;
        if (y < 0 || y > 119)
            err(10);
        wbytes(d, b, (y << 25) | ((dt.getMonth() + 1) << 21) | (dt.getDate() << 16) | (dt.getHours() << 11) | (dt.getMinutes() << 5) | (dt.getSeconds() >> 1)), b += 4;
        if (c != -1) {
            wbytes(d, b, f.crc);
            wbytes(d, b + 4, c < 0 ? -c - 2 : c);
            wbytes(d, b + 8, f.size);
        }
        wbytes(d, b + 12, fl);
        wbytes(d, b + 14, exl), b += 16;
        if (ce != null) {
            wbytes(d, b, col);
            wbytes(d, b + 6, f.attrs);
            wbytes(d, b + 10, ce), b += 14;
        }
        d.set(fn, b);
        b += fl;
        if (exl) {
            for (var k in ex) {
                var exf = ex[k], l = exf.length;
                wbytes(d, b, +k);
                wbytes(d, b + 2, l);
                d.set(exf, b + 4), b += 4 + l;
            }
        }
        if (col)
            d.set(co, b), b += col;
        return b;
    };
    // write zip footer (end of central directory)
    var wzf = function (o, b, c, d, e) {
        wbytes(o, b, 0x6054B50); // skip disk
        wbytes(o, b + 8, c);
        wbytes(o, b + 10, c);
        wbytes(o, b + 12, d);
        wbytes(o, b + 16, e);
    };
    /**
     * Synchronously creates a ZIP file. Prefer using `zip` for better performance
     * with more than one file.
     * @param data The directory structure for the ZIP archive
     * @param opts The main options, merged with per-file options
     * @returns The generated ZIP archive
     */
    function zipSync(data, opts) {
        if (!opts)
            opts = {};
        var r = {};
        var files = [];
        fltn(data, '', r, opts);
        var o = 0;
        var tot = 0;
        for (var fn in r) {
            var _a = r[fn], file = _a[0], p = _a[1];
            var compression = p.level == 0 ? 0 : 8;
            var f = strToU8(fn), s = f.length;
            var com = p.comment, m = com && strToU8(com), ms = m && m.length;
            var exl = exfl(p.extra);
            if (s > 65535)
                err(11);
            var d = compression ? deflateSync(file, p) : file, l = d.length;
            var c = crc();
            c.p(file);
            files.push(mrg(p, {
                size: file.length,
                crc: c.d(),
                c: d,
                f: f,
                m: m,
                u: s != fn.length || (m && (com.length != ms)),
                o: o,
                compression: compression
            }));
            o += 30 + s + exl + l;
            tot += 76 + 2 * (s + exl) + (ms || 0) + l;
        }
        var out = new u8(tot + 22), oe = o, cdl = tot - o;
        for (var i = 0; i < files.length; ++i) {
            var f = files[i];
            wzh(out, f.o, f, f.f, f.u, f.c.length);
            var badd = 30 + f.f.length + exfl(f.extra);
            out.set(f.c, f.o + badd);
            wzh(out, o, f, f.f, f.u, f.c.length, f.o, f.m), o += 16 + badd + (f.m ? f.m.length : 0);
        }
        wzf(out, o, files.length, cdl, oe);
        return out;
    }
    /**
     * Synchronously decompresses a ZIP archive. Prefer using `unzip` for better
     * performance with more than one file.
     * @param data The raw compressed ZIP file
     * @param opts The ZIP extraction options
     * @returns The decompressed files
     */
    function unzipSync(data, opts) {
        var files = {};
        var e = data.length - 22;
        for (; b4(data, e) != 0x6054B50; --e) {
            if (!e || data.length - e > 65558)
                err(13);
        }
        var c = b2(data, e + 8);
        if (!c)
            return {};
        var o = b4(data, e + 16);
        var z = o == 4294967295 || c == 65535;
        if (z) {
            var ze = b4(data, e - 12);
            z = b4(data, ze) == 0x6064B50;
            if (z) {
                c = b4(data, ze + 32);
                o = b4(data, ze + 48);
            }
        }
        for (var i = 0; i < c; ++i) {
            var _a = zh(data, o, z), c_2 = _a[0], sc = _a[1], su = _a[2], fn = _a[3], no = _a[4], off = _a[5], b = slzh(data, off);
            o = no;
            {
                if (!c_2)
                    files[fn] = slc(data, b, b + sc);
                else if (c_2 == 8)
                    files[fn] = inflateSync(data.subarray(b, b + sc), { out: new u8(su) });
                else
                    err(14, 'unknown compression type ' + c_2);
            }
        }
        return files;
    }

    class OpenXmlPackage {
        constructor(_files, options) {
            this._files = _files;
            this.options = options;
            this.xmlParser = new XmlParser();
            this.decoder = new TextDecoder();
            this.encoder = new TextEncoder();
        }
        get(path) {
            const p = normalizePath$2(path);
            return this._files[p] ?? this._files[p.replace(/\//g, "\\")] ?? null;
        }
        update(path, content) {
            this._files[normalizePath$2(path)] = toUint8Array(content, this.encoder);
        }
        static async load(input, options) {
            const data = await inputToUint8Array$1(input);
            return new OpenXmlPackage(normalizeFiles$1(unzipSync(data)), options);
        }
        static fromFiles(files, options) {
            return new OpenXmlPackage(normalizeFiles$1(files), options);
        }
        save(type = "blob") {
            const zipped = zipSync(this._files);
            switch (type) {
                case "uint8array":
                    return Promise.resolve(zipped);
                case "arraybuffer":
                    return Promise.resolve(zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength));
                case "blob":
                default:
                    return Promise.resolve(new Blob([new Uint8Array(zipped)]));
            }
        }
        load(path, type = "string") {
            const file = this.get(path);
            if (!file)
                return Promise.resolve(null);
            switch (type) {
                case "uint8array":
                    return Promise.resolve(file.slice());
                case "arraybuffer":
                    return Promise.resolve(file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength));
                case "blob":
                    return Promise.resolve(new Blob([new Uint8Array(file)]));
                case "string":
                default:
                    return Promise.resolve(this.decoder.decode(file));
            }
        }
        async loadRelationships(path = null) {
            let relsPath = `_rels/.rels`;
            if (path != null) {
                const [f, fn] = splitPath(path);
                relsPath = `${f}_rels/${fn}.rels`;
            }
            const txt = await this.load(relsPath);
            return txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
        }
        parseXmlDocument(txt) {
            return parseXmlString(txt, this.options.trimXmlDeclaration);
        }
    }
    function normalizePath$2(path) {
        return path.startsWith('/') ? path.substr(1) : path;
    }
    function normalizeFiles$1(files) {
        const result = {};
        for (const [path, file] of Object.entries(files ?? {})) {
            result[normalizePath$2(path)] = file;
        }
        return result;
    }
    async function inputToUint8Array$1(input) {
        if (input instanceof Uint8Array)
            return input.slice();
        if (input instanceof ArrayBuffer)
            return new Uint8Array(input);
        if (ArrayBuffer.isView(input))
            return new Uint8Array(input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength));
        if (input instanceof Blob)
            return new Uint8Array(await input.arrayBuffer());
        if (typeof input?.arrayBuffer === "function")
            return new Uint8Array(await input.arrayBuffer());
        throw new Error("Unsupported input type for OpenXmlPackage.load");
    }
    function toUint8Array(value, encoder) {
        if (value instanceof Uint8Array)
            return value;
        if (value instanceof ArrayBuffer)
            return new Uint8Array(value);
        if (typeof value === "string")
            return encoder.encode(value);
        if (value instanceof Blob)
            throw new Error("Blob updates are not supported for in-memory OpenXmlPackage");
        return new Uint8Array(value);
    }

    class DocumentPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.body = this._documentParser.parseDocumentFile(root);
        }
    }

    function parseBorder(elem, xml) {
        return {
            type: xml.attr(elem, "val"),
            color: xml.attr(elem, "color"),
            size: xml.lengthAttr(elem, "sz", LengthUsage.Border),
            offset: xml.lengthAttr(elem, "space", LengthUsage.Point),
            frame: xml.boolAttr(elem, 'frame'),
            shadow: xml.boolAttr(elem, 'shadow')
        };
    }
    function parseBorders(elem, xml) {
        var result = {};
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "left":
                    result.left = parseBorder(e, xml);
                    break;
                case "top":
                    result.top = parseBorder(e, xml);
                    break;
                case "right":
                    result.right = parseBorder(e, xml);
                    break;
                case "bottom":
                    result.bottom = parseBorder(e, xml);
                    break;
            }
        }
        return result;
    }

    var SectionType;
    (function (SectionType) {
        SectionType["Continuous"] = "continuous";
        SectionType["NextPage"] = "nextPage";
        SectionType["NextColumn"] = "nextColumn";
        SectionType["EvenPage"] = "evenPage";
        SectionType["OddPage"] = "oddPage";
    })(SectionType || (SectionType = {}));
    function parseSectionProperties(elem, xml = globalXmlParser) {
        var section = {};
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "pgSz":
                    section.pageSize = {
                        width: xml.lengthAttr(e, "w"),
                        height: xml.lengthAttr(e, "h"),
                        orientation: xml.attr(e, "orient")
                    };
                    break;
                case "type":
                    section.type = xml.attr(e, "val");
                    break;
                case "pgMar":
                    section.pageMargins = {
                        left: xml.lengthAttr(e, "left"),
                        right: xml.lengthAttr(e, "right"),
                        top: xml.lengthAttr(e, "top"),
                        bottom: xml.lengthAttr(e, "bottom"),
                        header: xml.lengthAttr(e, "header"),
                        footer: xml.lengthAttr(e, "footer"),
                        gutter: xml.lengthAttr(e, "gutter"),
                    };
                    break;
                case "cols":
                    section.columns = parseColumns(e, xml);
                    break;
                case "headerReference":
                    (section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(e, xml));
                    break;
                case "footerReference":
                    (section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(e, xml));
                    break;
                case "titlePg":
                    section.titlePage = xml.boolAttr(e, "val", true);
                    break;
                case "pgBorders":
                    section.pageBorders = parseBorders(e, xml);
                    break;
                case "pgNumType":
                    section.pageNumber = parsePageNumber(e, xml);
                    break;
            }
        }
        return section;
    }
    function parseColumns(elem, xml) {
        return {
            numberOfColumns: xml.intAttr(elem, "num"),
            space: xml.lengthAttr(elem, "space"),
            separator: xml.boolAttr(elem, "sep"),
            equalWidth: xml.boolAttr(elem, "equalWidth", true),
            columns: xml.elements(elem, "col")
                .map(e => ({
                width: xml.lengthAttr(e, "w"),
                space: xml.lengthAttr(e, "space")
            }))
        };
    }
    function parsePageNumber(elem, xml) {
        return {
            chapSep: xml.attr(elem, "chapSep"),
            chapStyle: xml.attr(elem, "chapStyle"),
            format: xml.attr(elem, "fmt"),
            start: xml.intAttr(elem, "start")
        };
    }
    function parseFooterHeaderReference(elem, xml) {
        return {
            id: xml.attr(elem, "id"),
            type: xml.attr(elem, "type"),
        };
    }

    function parseLineSpacing(elem, xml) {
        return {
            before: xml.lengthAttr(elem, "before"),
            after: xml.lengthAttr(elem, "after"),
            line: xml.intAttr(elem, "line"),
            lineRule: xml.attr(elem, "lineRule")
        };
    }

    function parseRunProperties(elem, xml) {
        let result = {};
        for (let el of xml.elements(elem)) {
            parseRunProperty(el, result, xml);
        }
        return result;
    }
    function parseRunProperty(elem, props, xml) {
        if (parseCommonProperty(elem, props, xml))
            return true;
        return false;
    }

    function parseParagraphProperties(elem, xml) {
        let result = {};
        for (let el of xml.elements(elem)) {
            parseParagraphProperty(el, result, xml);
        }
        return result;
    }
    function parseParagraphProperty(elem, props, xml) {
        if (elem.namespaceURI != ns$1.wordml)
            return false;
        if (parseCommonProperty(elem, props, xml))
            return true;
        switch (elem.localName) {
            case "tabs":
                props.tabs = parseTabs(elem, xml);
                break;
            case "sectPr":
                props.sectionProps = parseSectionProperties(elem, xml);
                break;
            case "numPr":
                props.numbering = parseNumbering$1(elem, xml);
                break;
            case "spacing":
                props.lineSpacing = parseLineSpacing(elem, xml);
                return false;
            case "textAlignment":
                props.textAlignment = xml.attr(elem, "val");
                return false;
            case "keepLines":
                props.keepLines = xml.boolAttr(elem, "val", true);
                break;
            case "keepNext":
                props.keepNext = xml.boolAttr(elem, "val", true);
                break;
            case "pageBreakBefore":
                props.pageBreakBefore = xml.boolAttr(elem, "val", true);
                break;
            case "outlineLvl":
                props.outlineLevel = xml.intAttr(elem, "val");
                break;
            case "pStyle":
                props.styleName = xml.attr(elem, "val");
                break;
            case "rPr":
                props.runProps = parseRunProperties(elem, xml);
                break;
            default:
                return false;
        }
        return true;
    }
    function parseTabs(elem, xml) {
        return xml.elements(elem, "tab")
            .map(e => ({
            position: xml.lengthAttr(e, "pos"),
            leader: xml.attr(e, "leader"),
            style: xml.attr(e, "val")
        }));
    }
    function parseNumbering$1(elem, xml) {
        var result = {};
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "numId":
                    result.id = xml.attr(e, "val");
                    break;
                case "ilvl":
                    result.level = xml.intAttr(e, "val");
                    break;
            }
        }
        return result;
    }

    function parseNumberingPart(elem, xml) {
        let result = {
            numberings: [],
            abstractNumberings: [],
            bulletPictures: []
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "num":
                    result.numberings.push(parseNumbering(e, xml));
                    break;
                case "abstractNum":
                    result.abstractNumberings.push(parseAbstractNumbering(e, xml));
                    break;
                case "numPicBullet":
                    result.bulletPictures.push(parseNumberingBulletPicture(e, xml));
                    break;
            }
        }
        return result;
    }
    function parseNumbering(elem, xml) {
        let result = {
            id: xml.attr(elem, 'numId'),
            overrides: []
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "abstractNumId":
                    result.abstractId = xml.attr(e, "val");
                    break;
                case "lvlOverride":
                    result.overrides.push(parseNumberingLevelOverrride(e, xml));
                    break;
            }
        }
        return result;
    }
    function parseAbstractNumbering(elem, xml) {
        let result = {
            id: xml.attr(elem, 'abstractNumId'),
            levels: []
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "name":
                    result.name = xml.attr(e, "val");
                    break;
                case "multiLevelType":
                    result.multiLevelType = xml.attr(e, "val");
                    break;
                case "numStyleLink":
                    result.numberingStyleLink = xml.attr(e, "val");
                    break;
                case "styleLink":
                    result.styleLink = xml.attr(e, "val");
                    break;
                case "lvl":
                    result.levels.push(parseNumberingLevel(e, xml));
                    break;
            }
        }
        return result;
    }
    function parseNumberingLevel(elem, xml) {
        let result = {
            level: xml.intAttr(elem, 'ilvl')
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "start":
                    result.start = xml.attr(e, "val");
                    break;
                case "lvlRestart":
                    result.restart = xml.intAttr(e, "val");
                    break;
                case "numFmt":
                    result.format = xml.attr(e, "val");
                    break;
                case "lvlText":
                    result.text = xml.attr(e, "val");
                    break;
                case "lvlJc":
                    result.justification = xml.attr(e, "val");
                    break;
                case "lvlPicBulletId":
                    result.bulletPictureId = xml.attr(e, "val");
                    break;
                case "pStyle":
                    result.paragraphStyle = xml.attr(e, "val");
                    break;
                case "pPr":
                    result.paragraphProps = parseParagraphProperties(e, xml);
                    break;
                case "rPr":
                    result.runProps = parseRunProperties(e, xml);
                    break;
            }
        }
        return result;
    }
    function parseNumberingLevelOverrride(elem, xml) {
        let result = {
            level: xml.intAttr(elem, 'ilvl')
        };
        for (let e of xml.elements(elem)) {
            switch (e.localName) {
                case "startOverride":
                    result.start = xml.intAttr(e, "val");
                    break;
                case "lvl":
                    result.numberingLevel = parseNumberingLevel(e, xml);
                    break;
            }
        }
        return result;
    }
    function parseNumberingBulletPicture(elem, xml) {
        var pict = xml.element(elem, "pict");
        var shape = pict && xml.element(pict, "shape");
        var imagedata = shape && xml.element(shape, "imagedata");
        return imagedata ? {
            id: xml.attr(elem, "numPicBulletId"),
            referenceId: xml.attr(imagedata, "id"),
            style: xml.attr(shape, "style")
        } : null;
    }

    class NumberingPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            Object.assign(this, parseNumberingPart(root, this._package.xmlParser));
            this.domNumberings = this._documentParser.parseNumberingFile(root);
        }
    }

    class StylesPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.styles = this._documentParser.parseStylesFile(root);
        }
    }

    var DomType;
    (function (DomType) {
        DomType["Document"] = "document";
        DomType["Paragraph"] = "paragraph";
        DomType["Run"] = "run";
        DomType["Break"] = "break";
        DomType["NoBreakHyphen"] = "noBreakHyphen";
        DomType["Table"] = "table";
        DomType["Row"] = "row";
        DomType["Cell"] = "cell";
        DomType["Hyperlink"] = "hyperlink";
        DomType["SmartTag"] = "smartTag";
        DomType["Drawing"] = "drawing";
        DomType["Image"] = "image";
        DomType["Text"] = "text";
        DomType["Tab"] = "tab";
        DomType["Symbol"] = "symbol";
        DomType["BookmarkStart"] = "bookmarkStart";
        DomType["BookmarkEnd"] = "bookmarkEnd";
        DomType["Footer"] = "footer";
        DomType["Header"] = "header";
        DomType["FootnoteReference"] = "footnoteReference";
        DomType["EndnoteReference"] = "endnoteReference";
        DomType["Footnote"] = "footnote";
        DomType["Endnote"] = "endnote";
        DomType["SimpleField"] = "simpleField";
        DomType["ComplexField"] = "complexField";
        DomType["Instruction"] = "instruction";
        DomType["VmlPicture"] = "vmlPicture";
        DomType["MmlMath"] = "mmlMath";
        DomType["MmlMathParagraph"] = "mmlMathParagraph";
        DomType["MmlFraction"] = "mmlFraction";
        DomType["MmlFunction"] = "mmlFunction";
        DomType["MmlFunctionName"] = "mmlFunctionName";
        DomType["MmlNumerator"] = "mmlNumerator";
        DomType["MmlDenominator"] = "mmlDenominator";
        DomType["MmlRadical"] = "mmlRadical";
        DomType["MmlBase"] = "mmlBase";
        DomType["MmlDegree"] = "mmlDegree";
        DomType["MmlSuperscript"] = "mmlSuperscript";
        DomType["MmlSubscript"] = "mmlSubscript";
        DomType["MmlPreSubSuper"] = "mmlPreSubSuper";
        DomType["MmlSubArgument"] = "mmlSubArgument";
        DomType["MmlSuperArgument"] = "mmlSuperArgument";
        DomType["MmlNary"] = "mmlNary";
        DomType["MmlDelimiter"] = "mmlDelimiter";
        DomType["MmlRun"] = "mmlRun";
        DomType["MmlEquationArray"] = "mmlEquationArray";
        DomType["MmlLimit"] = "mmlLimit";
        DomType["MmlLimitLower"] = "mmlLimitLower";
        DomType["MmlMatrix"] = "mmlMatrix";
        DomType["MmlMatrixRow"] = "mmlMatrixRow";
        DomType["MmlBox"] = "mmlBox";
        DomType["MmlBar"] = "mmlBar";
        DomType["MmlGroupChar"] = "mmlGroupChar";
        DomType["VmlElement"] = "vmlElement";
        DomType["Inserted"] = "inserted";
        DomType["Deleted"] = "deleted";
        DomType["DeletedText"] = "deletedText";
        DomType["Comment"] = "comment";
        DomType["CommentReference"] = "commentReference";
        DomType["CommentRangeStart"] = "commentRangeStart";
        DomType["CommentRangeEnd"] = "commentRangeEnd";
        DomType["AltChunk"] = "altChunk";
    })(DomType || (DomType = {}));
    class OpenXmlElementBase {
        constructor() {
            this.children = [];
            this.cssStyle = {};
        }
    }

    class WmlHeader extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.Header;
        }
    }
    class WmlFooter extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.Footer;
        }
    }

    class BaseHeaderFooterPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.rootElement = this.createRootElement();
            this.rootElement.children = this._documentParser.parseBodyElements(root);
        }
    }
    class HeaderPart extends BaseHeaderFooterPart {
        createRootElement() {
            return new WmlHeader();
        }
    }
    class FooterPart extends BaseHeaderFooterPart {
        createRootElement() {
            return new WmlFooter();
        }
    }

    function parseExtendedProps(root, xmlParser) {
        const result = {};
        for (let el of xmlParser.elements(root)) {
            switch (el.localName) {
                case "Template":
                    result.template = el.textContent;
                    break;
                case "Pages":
                    result.pages = safeParseToInt(el.textContent);
                    break;
                case "Words":
                    result.words = safeParseToInt(el.textContent);
                    break;
                case "Characters":
                    result.characters = safeParseToInt(el.textContent);
                    break;
                case "Application":
                    result.application = el.textContent;
                    break;
                case "Lines":
                    result.lines = safeParseToInt(el.textContent);
                    break;
                case "Paragraphs":
                    result.paragraphs = safeParseToInt(el.textContent);
                    break;
                case "Company":
                    result.company = el.textContent;
                    break;
                case "AppVersion":
                    result.appVersion = el.textContent;
                    break;
            }
        }
        return result;
    }
    function safeParseToInt(value) {
        if (typeof value === 'undefined')
            return;
        return parseInt(value);
    }

    class ExtendedPropsPart extends Part {
        parseXml(root) {
            this.props = parseExtendedProps(root, this._package.xmlParser);
        }
    }

    function parseCoreProps(root, xmlParser) {
        const result = {};
        for (let el of xmlParser.elements(root)) {
            switch (el.localName) {
                case "title":
                    result.title = el.textContent;
                    break;
                case "description":
                    result.description = el.textContent;
                    break;
                case "subject":
                    result.subject = el.textContent;
                    break;
                case "creator":
                    result.creator = el.textContent;
                    break;
                case "keywords":
                    result.keywords = el.textContent;
                    break;
                case "language":
                    result.language = el.textContent;
                    break;
                case "lastModifiedBy":
                    result.lastModifiedBy = el.textContent;
                    break;
                case "revision":
                    el.textContent && (result.revision = parseInt(el.textContent));
                    break;
            }
        }
        return result;
    }

    class CorePropsPart extends Part {
        parseXml(root) {
            this.props = parseCoreProps(root, this._package.xmlParser);
        }
    }

    class DmlTheme {
    }
    function parseTheme(elem, xml) {
        var result = new DmlTheme();
        var themeElements = xml.element(elem, "themeElements");
        for (let el of xml.elements(themeElements)) {
            switch (el.localName) {
                case "clrScheme":
                    result.colorScheme = parseColorScheme(el, xml);
                    break;
                case "fontScheme":
                    result.fontScheme = parseFontScheme(el, xml);
                    break;
            }
        }
        return result;
    }
    function parseColorScheme(elem, xml) {
        var result = {
            name: xml.attr(elem, "name"),
            colors: {}
        };
        for (let el of xml.elements(elem)) {
            var srgbClr = xml.element(el, "srgbClr");
            var sysClr = xml.element(el, "sysClr");
            if (srgbClr) {
                result.colors[el.localName] = xml.attr(srgbClr, "val");
            }
            else if (sysClr) {
                result.colors[el.localName] = xml.attr(sysClr, "lastClr");
            }
        }
        return result;
    }
    function parseFontScheme(elem, xml) {
        var result = {
            name: xml.attr(elem, "name"),
        };
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "majorFont":
                    result.majorFont = parseFontInfo(el, xml);
                    break;
                case "minorFont":
                    result.minorFont = parseFontInfo(el, xml);
                    break;
            }
        }
        return result;
    }
    function parseFontInfo(elem, xml) {
        return {
            latinTypeface: xml.elementAttr(elem, "latin", "typeface"),
            eaTypeface: xml.elementAttr(elem, "ea", "typeface"),
            csTypeface: xml.elementAttr(elem, "cs", "typeface"),
        };
    }

    class ThemePart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
        parseXml(root) {
            this.theme = parseTheme(root, this._package.xmlParser);
        }
    }

    class WmlBaseNote {
    }
    class WmlFootnote extends WmlBaseNote {
        constructor() {
            super(...arguments);
            this.type = DomType.Footnote;
        }
    }
    class WmlEndnote extends WmlBaseNote {
        constructor() {
            super(...arguments);
            this.type = DomType.Endnote;
        }
    }

    class BaseNotePart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
    }
    class FootnotesPart extends BaseNotePart {
        constructor(pkg, path, parser) {
            super(pkg, path, parser);
        }
        parseXml(root) {
            this.notes = this._documentParser.parseNotes(root, "footnote", WmlFootnote);
        }
    }
    class EndnotesPart extends BaseNotePart {
        constructor(pkg, path, parser) {
            super(pkg, path, parser);
        }
        parseXml(root) {
            this.notes = this._documentParser.parseNotes(root, "endnote", WmlEndnote);
        }
    }

    function parseSettings(elem, xml) {
        var result = {};
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "defaultTabStop":
                    result.defaultTabStop = xml.lengthAttr(el, "val");
                    break;
                case "footnotePr":
                    result.footnoteProps = parseNoteProperties(el, xml);
                    break;
                case "endnotePr":
                    result.endnoteProps = parseNoteProperties(el, xml);
                    break;
                case "autoHyphenation":
                    result.autoHyphenation = xml.boolAttr(el, "val");
                    break;
            }
        }
        return result;
    }
    function parseNoteProperties(elem, xml) {
        var result = {
            defaultNoteIds: []
        };
        for (let el of xml.elements(elem)) {
            switch (el.localName) {
                case "numFmt":
                    result.nummeringFormat = xml.attr(el, "val");
                    break;
                case "footnote":
                case "endnote":
                    result.defaultNoteIds.push(xml.attr(el, "id"));
                    break;
            }
        }
        return result;
    }

    class SettingsPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
        }
        parseXml(root) {
            this.settings = parseSettings(root, this._package.xmlParser);
        }
    }

    function parseCustomProps(root, xml) {
        return xml.elements(root, "property").map(e => {
            const firstChild = e.firstChild;
            return {
                formatId: xml.attr(e, "fmtid"),
                name: xml.attr(e, "name"),
                type: firstChild.nodeName,
                value: firstChild.textContent
            };
        });
    }

    class CustomPropsPart extends Part {
        parseXml(root) {
            this.props = parseCustomProps(root, this._package.xmlParser);
        }
    }

    class CommentsPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.comments = this._documentParser.parseComments(root);
            this.commentMap = keyBy(this.comments, x => x.id);
        }
    }

    class CommentsExtendedPart extends Part {
        constructor(pkg, path) {
            super(pkg, path);
            this.comments = [];
        }
        parseXml(root) {
            const xml = this._package.xmlParser;
            for (let el of xml.elements(root, "commentEx")) {
                this.comments.push({
                    paraId: xml.attr(el, 'paraId'),
                    paraIdParent: xml.attr(el, 'paraIdParent'),
                    done: xml.boolAttr(el, 'done')
                });
            }
            this.commentMap = keyBy(this.comments, x => x.paraId);
        }
    }

    async function parseDocumentInWorker(data, options) {
        const workerUrl = resolveWorkerUrl(options.workerUrl);
        if (!workerUrl || typeof Worker === "undefined") {
            return null;
        }
        const client = new ParserWorkerClient(workerUrl);
        try {
            const buffer = await toArrayBuffer$1(data);
            const parsed = await client.parse(buffer, options);
            return {
                parsed,
                package: new WorkerSessionPackage(client, parsed.sessionId, options)
            };
        }
        catch (error) {
            client.terminate();
            throw error;
        }
    }
    class ParserWorkerClient {
        constructor(workerUrl) {
            this.nextRequestId = 1;
            this.pending = new Map();
            this.worker = new Worker(workerUrl);
            this.worker.onmessage = event => this.handleMessage(event.data);
            this.worker.onerror = event => this.failAll(event.error ?? new Error(event.message));
        }
        async parse(buffer, options) {
            const response = await this.request({
                type: "parse",
                buffer,
                options: prepareWorkerOptions(options)
            }, [buffer]);
            return response.payload;
        }
        async loadResource(sessionId, path, outputType) {
            const response = await this.request({
                type: "load-resource",
                sessionId,
                path,
                outputType
            });
            if (response.type === "null")
                return null;
            if (outputType === "string")
                return response.value;
            const buffer = response.value;
            switch (outputType) {
                case "uint8array":
                    return new Uint8Array(buffer);
                case "blob":
                    return new Blob([new Uint8Array(buffer)]);
                case "arraybuffer":
                default:
                    return buffer;
            }
        }
        async save(sessionId, outputType) {
            const response = await this.request({
                type: "save",
                sessionId,
                outputType
            });
            const buffer = response.value;
            switch (outputType) {
                case "uint8array":
                    return new Uint8Array(buffer);
                case "blob":
                    return new Blob([new Uint8Array(buffer)]);
                case "arraybuffer":
                default:
                    return buffer;
            }
        }
        async dispose(sessionId) {
            if (!this.worker)
                return;
            try {
                await this.request({
                    type: "dispose",
                    sessionId
                });
            }
            finally {
                this.terminate();
            }
        }
        terminate() {
            if (!this.worker)
                return;
            this.worker.terminate();
            this.worker = null;
            this.failAll(new Error("Parser worker terminated"));
        }
        request(message, transfer = []) {
            const requestId = this.nextRequestId++;
            return new Promise((resolve, reject) => {
                this.pending.set(requestId, {
                    resolve: resolve,
                    reject
                });
                this.worker.postMessage({
                    ...message,
                    requestId
                }, transfer);
            });
        }
        handleMessage(message) {
            const pending = this.pending.get(message.requestId);
            if (!pending)
                return;
            this.pending.delete(message.requestId);
            if (message.type === "error") {
                pending.reject(new Error(message.error ?? "Unknown parser worker error"));
            }
            else {
                pending.resolve(message);
            }
        }
        failAll(error) {
            if (this.pending.size === 0)
                return;
            const pending = Array.from(this.pending.values());
            this.pending.clear();
            for (const entry of pending) {
                entry.reject(error);
            }
        }
    }
    class WorkerSessionPackage {
        constructor(client, sessionId, options) {
            this.client = client;
            this.sessionId = sessionId;
            this.options = options;
            this.xmlParser = new XmlParser();
        }
        get(path) {
            return path ? true : null;
        }
        update() {
            throw new Error("DOCX: update() is not supported for worker-backed packages");
        }
        load(path, type = "string") {
            return this.client.loadResource(this.sessionId, normalizePath$1(path), type);
        }
        save(type = "blob") {
            return this.client.save(this.sessionId, type);
        }
        async loadRelationships(path = null) {
            let relsPath = `_rels/.rels`;
            if (path != null) {
                const [folder, fileName] = splitPath(path);
                relsPath = `${folder}_rels/${fileName}.rels`;
            }
            const text = await this.load(relsPath, "string");
            return text ? parseRelationships(this.parseXmlDocument(text).firstElementChild, this.xmlParser) : null;
        }
        parseXmlDocument(text) {
            return parseXmlString(text, this.options.trimXmlDeclaration);
        }
        dispose() {
            return this.client.dispose(this.sessionId);
        }
    }
    function prepareWorkerOptions(options) {
        return JSON.parse(JSON.stringify({
            ...options,
            workerUrl: undefined
        }));
    }
    function resolveWorkerUrl(explicitUrl) {
        if (explicitUrl)
            return explicitUrl instanceof URL ? explicitUrl.href : explicitUrl.toString();
        if (typeof document === "undefined")
            return null;
        const scripts = Array.from(document.scripts ?? []).reverse();
        for (const script of scripts) {
            const src = script.src;
            if (!src)
                continue;
            if (/docx-preview(?:\.min)?\.js(?:\?.*)?$/i.test(src)) {
                return src.replace(/docx-preview(?:\.min)?\.js(?:\?.*)?$/i, "docx-preview-worker.js");
            }
        }
        return null;
    }
    async function toArrayBuffer$1(data) {
        if (data instanceof ArrayBuffer)
            return data;
        if (data instanceof Uint8Array)
            return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
        if (data && typeof data.arrayBuffer === "function")
            return await data.arrayBuffer();
        throw new Error("Unsupported input type for parser worker");
    }
    function normalizePath$1(path) {
        return path.startsWith("/") ? path.substring(1) : path;
    }

    const topLevelRels$1 = [
        { type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
        { type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
        { type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
        { type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
    ];
    class WordDocument {
        constructor() {
            this._objectUrls = new Set();
            this._snapshotPages = null;
            this.parts = [];
            this.partsMap = {};
            this.pages = null;
        }
        static async load(blob, parser, options) {
            var d = new WordDocument();
            d._options = options;
            d._parser = parser;
            if (options.useWorkerParser) {
                try {
                    const workerResult = await parseDocumentInWorker(blob, options);
                    if (workerResult) {
                        d._package = workerResult.package;
                        d.applySerializedDocument(workerResult.parsed);
                        return d;
                    }
                }
                catch (error) {
                    if (options.debug) {
                        console.warn("DOCX: Worker parser failed, falling back to main thread", error);
                    }
                }
            }
            d._package = await OpenXmlPackage.load(blob, options);
            d.rels = await d._package.loadRelationships();
            await Promise.all(topLevelRels$1.map(rel => {
                const r = d.rels.find(x => x.type === rel.type) ?? rel;
                return d.loadRelationshipPart(r.target, r.type);
            }));
            return d;
        }
        static fromSnapshot(snapshot, options) {
            const document = new WordDocument();
            document._options = options;
            document._package = OpenXmlPackage.fromFiles(snapshotFilesToMap(snapshot.files), options);
            document.applySerializedDocument({
                sessionId: "snapshot",
                rels: snapshot.rels,
                parts: materializeSnapshotParts(snapshot.parts),
                rolePaths: snapshot.rolePaths
            });
            document.pages = snapshot.pages;
            document._snapshotPages = snapshot.pages;
            return document;
        }
        preparePageForRender(page) {
            return this._snapshotPages ? cloneSerializable$1(page) : page;
        }
        applySerializedDocument(data) {
            this.rels = data.rels;
            this.parts = data.parts;
            this.partsMap = keyBy(this.parts, x => x.path);
            this.documentPart = data.rolePaths.documentPart ? this.partsMap[data.rolePaths.documentPart] : null;
            this.fontTablePart = data.rolePaths.fontTablePart ? this.partsMap[data.rolePaths.fontTablePart] : null;
            this.numberingPart = data.rolePaths.numberingPart ? this.partsMap[data.rolePaths.numberingPart] : null;
            this.stylesPart = data.rolePaths.stylesPart ? this.partsMap[data.rolePaths.stylesPart] : null;
            this.footnotesPart = data.rolePaths.footnotesPart ? this.partsMap[data.rolePaths.footnotesPart] : null;
            this.endnotesPart = data.rolePaths.endnotesPart ? this.partsMap[data.rolePaths.endnotesPart] : null;
            this.themePart = data.rolePaths.themePart ? this.partsMap[data.rolePaths.themePart] : null;
            this.corePropsPart = data.rolePaths.corePropsPart ? this.partsMap[data.rolePaths.corePropsPart] : null;
            this.extendedPropsPart = data.rolePaths.extendedPropsPart ? this.partsMap[data.rolePaths.extendedPropsPart] : null;
            this.customPropsPart = data.rolePaths.customPropsPart ? this.partsMap[data.rolePaths.customPropsPart] : null;
            this.settingsPart = data.rolePaths.settingsPart ? this.partsMap[data.rolePaths.settingsPart] : null;
            this.commentsPart = data.rolePaths.commentsPart ? this.partsMap[data.rolePaths.commentsPart] : null;
            this.commentsExtendedPart = data.rolePaths.commentsExtendedPart ? this.partsMap[data.rolePaths.commentsExtendedPart] : null;
            if (this.commentsPart?.comments && !this.commentsPart.commentMap) {
                this.commentsPart.commentMap = keyBy(this.commentsPart.comments, x => x.id);
            }
            if (this.commentsExtendedPart?.comments && !this.commentsExtendedPart.commentMap) {
                this.commentsExtendedPart.commentMap = keyBy(this.commentsExtendedPart.comments, x => x.paraId);
            }
        }
        save(type = "blob") {
            return this._package.save(type);
        }
        async dispose() {
            for (const url of this._objectUrls) {
                URL.revokeObjectURL(url);
            }
            this._objectUrls.clear();
            if (typeof this._package?.dispose === "function") {
                await this._package.dispose();
            }
        }
        async loadRelationshipPart(path, type) {
            if (this.partsMap[path])
                return this.partsMap[path];
            if (!this._package.get(path))
                return null;
            let part = null;
            const pkg = this._package;
            switch (type) {
                case RelationshipTypes.OfficeDocument:
                    this.documentPart = part = new DocumentPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.FontTable:
                    this.fontTablePart = part = new FontTablePart(pkg, path);
                    break;
                case RelationshipTypes.Numbering:
                    this.numberingPart = part = new NumberingPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.Styles:
                    this.stylesPart = part = new StylesPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.Theme:
                    this.themePart = part = new ThemePart(pkg, path);
                    break;
                case RelationshipTypes.Footnotes:
                    this.footnotesPart = part = new FootnotesPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.Endnotes:
                    this.endnotesPart = part = new EndnotesPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.Footer:
                    part = new FooterPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.Header:
                    part = new HeaderPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.CoreProperties:
                    this.corePropsPart = part = new CorePropsPart(pkg, path);
                    break;
                case RelationshipTypes.ExtendedProperties:
                    this.extendedPropsPart = part = new ExtendedPropsPart(pkg, path);
                    break;
                case RelationshipTypes.CustomProperties:
                    part = new CustomPropsPart(pkg, path);
                    break;
                case RelationshipTypes.Settings:
                    this.settingsPart = part = new SettingsPart(pkg, path);
                    break;
                case RelationshipTypes.Comments:
                    this.commentsPart = part = new CommentsPart(pkg, path, this._parser);
                    break;
                case RelationshipTypes.CommentsExtended:
                    this.commentsExtendedPart = part = new CommentsExtendedPart(pkg, path);
                    break;
            }
            if (part == null)
                return Promise.resolve(null);
            this.partsMap[path] = part;
            this.parts.push(part);
            await part.load();
            if (part.rels?.length > 0) {
                const [folder] = splitPath(part.path);
                await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
            }
            return part;
        }
        async loadDocumentImage(id, part) {
            const x = await this.loadResource(part ?? this.documentPart, id, "blob");
            return this.blobToURL(x);
        }
        async loadNumberingImage(id) {
            const x = await this.loadResource(this.numberingPart, id, "blob");
            return this.blobToURL(x);
        }
        async loadFont(id, key) {
            const x = await this.loadResource(this.fontTablePart, id, "uint8array");
            return x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x;
        }
        async loadAltChunk(id, part) {
            return await this.loadResource(part ?? this.documentPart, id, "string");
        }
        blobToURL(blob) {
            if (!blob)
                return null;
            if (this._options.useBase64URL) {
                return blobToBase64(blob);
            }
            const url = URL.createObjectURL(blob);
            this._objectUrls.add(url);
            return url;
        }
        findPartByRelId(id, basePart = null) {
            var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
            const folder = basePart ? splitPath(basePart.path)[0] : '';
            return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
        }
        getPathById(part, id) {
            const rel = part.rels.find(x => x.id == id);
            const [folder] = splitPath(part.path);
            return rel ? resolvePath(rel.target, folder) : null;
        }
        loadResource(part, id, outputType) {
            const path = this.getPathById(part, id);
            return path ? this._package.load(path, outputType) : Promise.resolve(null);
        }
    }
    function deobfuscate(data, guidKey) {
        const len = 16;
        const trimmed = guidKey.replace(/{|}|-/g, "");
        const numbers = new Array(len);
        for (let i = 0; i < len; i++)
            numbers[len - i - 1] = parseInt(trimmed.substring(i * 2, i * 2 + 2), 16);
        for (let i = 0; i < 32; i++)
            data[i] = data[i] ^ numbers[i % len];
        return data;
    }
    function snapshotFilesToMap(files) {
        const result = {};
        for (const file of files ?? []) {
            result[file.path] = new Uint8Array(file.buffer);
        }
        return result;
    }
    function materializeSnapshotParts(parts) {
        return (parts ?? []).map(part => {
            const materialized = { ...part };
            switch (part.kind) {
                case "styles":
                    materialized.styles = cloneSerializable$1(part.styles);
                    break;
                case "header":
                case "footer":
                    materialized.rootElement = cloneSerializable$1(part.rootElement);
                    break;
                case "footnotes":
                case "endnotes":
                    materialized.notes = cloneSerializable$1(part.notes);
                    break;
            }
            return materialized;
        });
    }
    function cloneSerializable$1(value) {
        if (typeof structuredClone === "function") {
            return structuredClone(value);
        }
        if (Array.isArray(value)) {
            return value.map(item => cloneSerializable$1(item));
        }
        if (!value || typeof value !== "object") {
            return value;
        }
        const result = {};
        for (const [key, entry] of Object.entries(value)) {
            result[key] = cloneSerializable$1(entry);
        }
        return result;
    }

    function parseBookmarkStart(elem, xml) {
        return {
            type: DomType.BookmarkStart,
            id: xml.attr(elem, "id"),
            name: xml.attr(elem, "name"),
            colFirst: xml.intAttr(elem, "colFirst"),
            colLast: xml.intAttr(elem, "colLast")
        };
    }
    function parseBookmarkEnd(elem, xml) {
        return {
            type: DomType.BookmarkEnd,
            id: xml.attr(elem, "id")
        };
    }

    class VmlElement extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.VmlElement;
            this.attrs = {};
        }
    }
    function parseVmlElement(elem, parser) {
        var result = new VmlElement();
        switch (elem.localName) {
            case "rect":
                result.tagName = "rect";
                Object.assign(result.attrs, { width: '100%', height: '100%' });
                break;
            case "oval":
                result.tagName = "ellipse";
                Object.assign(result.attrs, { cx: "50%", cy: "50%", rx: "50%", ry: "50%" });
                break;
            case "line":
                result.tagName = "line";
                break;
            case "shape":
                result.tagName = "g";
                break;
            case "textbox":
                result.tagName = "foreignObject";
                Object.assign(result.attrs, { width: '100%', height: '100%' });
                break;
            default:
                return null;
        }
        for (const at of globalXmlParser.attrs(elem)) {
            switch (at.localName) {
                case "style":
                    result.cssStyleText = at.value;
                    break;
                case "fillcolor":
                    result.attrs.fill = at.value;
                    break;
                case "from":
                    const [x1, y1] = parsePoint(at.value);
                    Object.assign(result.attrs, { x1, y1 });
                    break;
                case "to":
                    const [x2, y2] = parsePoint(at.value);
                    Object.assign(result.attrs, { x2, y2 });
                    break;
            }
        }
        for (const el of globalXmlParser.elements(elem)) {
            switch (el.localName) {
                case "stroke":
                    Object.assign(result.attrs, parseStroke(el));
                    break;
                case "fill":
                    Object.assign(result.attrs, parseFill());
                    break;
                case "imagedata":
                    result.tagName = "image";
                    Object.assign(result.attrs, { width: '100%', height: '100%' });
                    result.imageHref = {
                        id: globalXmlParser.attr(el, "id"),
                        title: globalXmlParser.attr(el, "title"),
                    };
                    break;
                case "txbxContent":
                    result.children.push(...parser.parseBodyElements(el));
                    break;
                default:
                    const child = parseVmlElement(el, parser);
                    child && result.children.push(child);
                    break;
            }
        }
        return result;
    }
    function parseStroke(el) {
        return {
            'stroke': globalXmlParser.attr(el, "color"),
            'stroke-width': globalXmlParser.lengthAttr(el, "weight", LengthUsage.Emu) ?? '1px'
        };
    }
    function parseFill(el) {
        return {};
    }
    function parsePoint(val) {
        return val.split(",");
    }

    class WmlComment extends OpenXmlElementBase {
        constructor() {
            super(...arguments);
            this.type = DomType.Comment;
        }
    }
    class WmlCommentReference extends OpenXmlElementBase {
        constructor(id) {
            super();
            this.id = id;
            this.type = DomType.CommentReference;
        }
    }
    class WmlCommentRangeStart extends OpenXmlElementBase {
        constructor(id) {
            super();
            this.id = id;
            this.type = DomType.CommentRangeStart;
        }
    }
    class WmlCommentRangeEnd extends OpenXmlElementBase {
        constructor(id) {
            super();
            this.id = id;
            this.type = DomType.CommentRangeEnd;
        }
    }

    var autos = {
        shd: "inherit",
        color: "black",
        borderColor: "black",
        highlight: "transparent"
    };
    const supportedNamespaceURIs = [];
    const mmlTagMap = {
        "oMath": DomType.MmlMath,
        "oMathPara": DomType.MmlMathParagraph,
        "f": DomType.MmlFraction,
        "func": DomType.MmlFunction,
        "fName": DomType.MmlFunctionName,
        "num": DomType.MmlNumerator,
        "den": DomType.MmlDenominator,
        "rad": DomType.MmlRadical,
        "deg": DomType.MmlDegree,
        "e": DomType.MmlBase,
        "sSup": DomType.MmlSuperscript,
        "sSub": DomType.MmlSubscript,
        "sPre": DomType.MmlPreSubSuper,
        "sup": DomType.MmlSuperArgument,
        "sub": DomType.MmlSubArgument,
        "d": DomType.MmlDelimiter,
        "nary": DomType.MmlNary,
        "eqArr": DomType.MmlEquationArray,
        "lim": DomType.MmlLimit,
        "limLow": DomType.MmlLimitLower,
        "m": DomType.MmlMatrix,
        "mr": DomType.MmlMatrixRow,
        "box": DomType.MmlBox,
        "bar": DomType.MmlBar,
        "groupChr": DomType.MmlGroupChar
    };
    class DocumentParser {
        constructor(options) {
            this.options = {
                ignoreWidth: false,
                debug: false,
                ...options
            };
        }
        parseNotes(xmlDoc, elemName, elemClass) {
            var result = [];
            for (let el of globalXmlParser.elements(xmlDoc, elemName)) {
                const node = new elemClass();
                node.id = globalXmlParser.attr(el, "id");
                node.noteType = globalXmlParser.attr(el, "type");
                node.children = this.parseBodyElements(el);
                result.push(node);
            }
            return result;
        }
        parseComments(xmlDoc) {
            var result = [];
            for (let el of globalXmlParser.elements(xmlDoc, "comment")) {
                const item = new WmlComment();
                item.id = globalXmlParser.attr(el, "id");
                item.author = globalXmlParser.attr(el, "author");
                item.initials = globalXmlParser.attr(el, "initials");
                item.date = globalXmlParser.attr(el, "date");
                item.children = this.parseBodyElements(el);
                result.push(item);
            }
            return result;
        }
        parseDocumentFile(xmlDoc) {
            var xbody = globalXmlParser.element(xmlDoc, "body");
            var background = globalXmlParser.element(xmlDoc, "background");
            var sectPr = globalXmlParser.element(xbody, "sectPr");
            return {
                type: DomType.Document,
                children: this.parseBodyElements(xbody),
                props: sectPr ? parseSectionProperties(sectPr, globalXmlParser) : {},
                cssStyle: background ? this.parseBackground(background) : {},
            };
        }
        parseBackground(elem) {
            var result = {};
            var color = xmlUtil.colorAttr(elem, "color");
            if (color) {
                result["background-color"] = color;
            }
            return result;
        }
        parseBodyElements(element) {
            var children = [];
            for (const elem of globalXmlParser.elements(element)) {
                switch (elem.localName) {
                    case "p":
                        children.push(this.parseParagraph(elem));
                        break;
                    case "altChunk":
                        children.push(this.parseAltChunk(elem));
                        break;
                    case "tbl":
                        children.push(this.parseTable(elem));
                        break;
                    case "sdt":
                        children.push(...this.parseSdt(elem, e => this.parseBodyElements(e)));
                        break;
                }
            }
            return children;
        }
        parseStylesFile(xstyles) {
            var result = [];
            for (const n of globalXmlParser.elements(xstyles)) {
                switch (n.localName) {
                    case "style":
                        result.push(this.parseStyle(n));
                        break;
                    case "docDefaults":
                        result.push(this.parseDefaultStyles(n));
                        break;
                }
            }
            return result;
        }
        parseDefaultStyles(node) {
            var result = {
                id: null,
                name: null,
                target: null,
                basedOn: null,
                styles: []
            };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "rPrDefault":
                        var rPr = globalXmlParser.element(c, "rPr");
                        if (rPr)
                            result.styles.push({
                                target: "span",
                                values: this.parseDefaultProperties(rPr, {})
                            });
                        break;
                    case "pPrDefault":
                        var pPr = globalXmlParser.element(c, "pPr");
                        if (pPr)
                            result.styles.push({
                                target: "p",
                                values: this.parseDefaultProperties(pPr, {})
                            });
                        break;
                }
            }
            return result;
        }
        parseStyle(node) {
            var result = {
                id: globalXmlParser.attr(node, "styleId"),
                isDefault: globalXmlParser.boolAttr(node, "default"),
                name: null,
                target: null,
                basedOn: null,
                styles: [],
                linked: null
            };
            switch (globalXmlParser.attr(node, "type")) {
                case "paragraph":
                    result.target = "p";
                    break;
                case "table":
                    result.target = "table";
                    break;
                case "character":
                    result.target = "span";
                    break;
            }
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "basedOn":
                        result.basedOn = globalXmlParser.attr(n, "val");
                        break;
                    case "name":
                        result.name = globalXmlParser.attr(n, "val");
                        break;
                    case "link":
                        result.linked = globalXmlParser.attr(n, "val");
                        break;
                    case "next":
                        result.next = globalXmlParser.attr(n, "val");
                        break;
                    case "aliases":
                        result.aliases = globalXmlParser.attr(n, "val").split(",");
                        break;
                    case "pPr":
                        result.styles.push({
                            target: "p",
                            values: this.parseDefaultProperties(n, {})
                        });
                        result.paragraphProps = parseParagraphProperties(n, globalXmlParser);
                        break;
                    case "rPr":
                        result.styles.push({
                            target: "span",
                            values: this.parseDefaultProperties(n, {})
                        });
                        result.runProps = parseRunProperties(n, globalXmlParser);
                        break;
                    case "tblPr":
                    case "tcPr":
                        result.styles.push({
                            target: "td",
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                    case "tblStylePr":
                        for (let s of this.parseTableStyle(n))
                            result.styles.push(s);
                        break;
                    case "rsid":
                    case "qFormat":
                    case "hidden":
                    case "semiHidden":
                    case "unhideWhenUsed":
                    case "autoRedefine":
                    case "uiPriority":
                        break;
                    default:
                        this.options.debug && console.warn(`DOCX: Unknown style element: ${n.localName}`);
                }
            }
            return result;
        }
        parseTableStyle(node) {
            var result = [];
            var type = globalXmlParser.attr(node, "type");
            var selector = "";
            var modificator = "";
            switch (type) {
                case "firstRow":
                    modificator = ".first-row";
                    selector = "tr.first-row td";
                    break;
                case "lastRow":
                    modificator = ".last-row";
                    selector = "tr.last-row td";
                    break;
                case "firstCol":
                    modificator = ".first-col";
                    selector = "td.first-col";
                    break;
                case "lastCol":
                    modificator = ".last-col";
                    selector = "td.last-col";
                    break;
                case "band1Vert":
                    modificator = ":not(.no-vband)";
                    selector = "td.odd-col";
                    break;
                case "band2Vert":
                    modificator = ":not(.no-vband)";
                    selector = "td.even-col";
                    break;
                case "band1Horz":
                    modificator = ":not(.no-hband)";
                    selector = "tr.odd-row";
                    break;
                case "band2Horz":
                    modificator = ":not(.no-hband)";
                    selector = "tr.even-row";
                    break;
                default: return [];
            }
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "pPr":
                        result.push({
                            target: `${selector} p`,
                            mod: modificator,
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                    case "rPr":
                        result.push({
                            target: `${selector} span`,
                            mod: modificator,
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                    case "tblPr":
                    case "tcPr":
                        result.push({
                            target: selector,
                            mod: modificator,
                            values: this.parseDefaultProperties(n, {})
                        });
                        break;
                }
            }
            return result;
        }
        parseNumberingFile(node) {
            var result = [];
            var mapping = {};
            var bullets = [];
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "abstractNum":
                        this.parseAbstractNumbering(n, bullets)
                            .forEach(x => result.push(x));
                        break;
                    case "numPicBullet":
                        bullets.push(this.parseNumberingPicBullet(n));
                        break;
                    case "num":
                        var numId = globalXmlParser.attr(n, "numId");
                        var abstractNumId = globalXmlParser.elementAttr(n, "abstractNumId", "val");
                        mapping[abstractNumId] = numId;
                        break;
                }
            }
            result.forEach(x => x.id = mapping[x.id]);
            return result;
        }
        parseNumberingPicBullet(elem) {
            var pict = globalXmlParser.element(elem, "pict");
            var shape = pict && globalXmlParser.element(pict, "shape");
            var imagedata = shape && globalXmlParser.element(shape, "imagedata");
            return imagedata ? {
                id: globalXmlParser.intAttr(elem, "numPicBulletId"),
                src: globalXmlParser.attr(imagedata, "id"),
                style: globalXmlParser.attr(shape, "style")
            } : null;
        }
        parseAbstractNumbering(node, bullets) {
            var result = [];
            var id = globalXmlParser.attr(node, "abstractNumId");
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "lvl":
                        result.push(this.parseNumberingLevel(id, n, bullets));
                        break;
                }
            }
            return result;
        }
        parseNumberingLevel(id, node, bullets) {
            var result = {
                id: id,
                level: globalXmlParser.intAttr(node, "ilvl"),
                start: 1,
                pStyleName: undefined,
                pStyle: {},
                rStyle: {},
                suff: "tab"
            };
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "start":
                        result.start = globalXmlParser.intAttr(n, "val");
                        break;
                    case "pPr":
                        this.parseDefaultProperties(n, result.pStyle);
                        break;
                    case "rPr":
                        this.parseDefaultProperties(n, result.rStyle);
                        break;
                    case "lvlPicBulletId":
                        var bulletId = globalXmlParser.intAttr(n, "val");
                        result.bullet = bullets.find(x => x?.id == bulletId);
                        break;
                    case "lvlText":
                        result.levelText = globalXmlParser.attr(n, "val");
                        break;
                    case "pStyle":
                        result.pStyleName = globalXmlParser.attr(n, "val");
                        break;
                    case "numFmt":
                        result.format = globalXmlParser.attr(n, "val");
                        break;
                    case "suff":
                        result.suff = globalXmlParser.attr(n, "val");
                        break;
                }
            }
            return result;
        }
        parseSdt(node, parser) {
            const sdtContent = globalXmlParser.element(node, "sdtContent");
            return sdtContent ? parser(sdtContent) : [];
        }
        parseInserted(node, parentParser) {
            return {
                type: DomType.Inserted,
                children: parentParser(node)?.children ?? []
            };
        }
        parseDeleted(node, parentParser) {
            return {
                type: DomType.Deleted,
                children: parentParser(node)?.children ?? []
            };
        }
        parseAltChunk(node) {
            return { type: DomType.AltChunk, children: [], id: globalXmlParser.attr(node, "id") };
        }
        parseParagraph(node) {
            var result = { type: DomType.Paragraph, children: [] };
            for (let el of globalXmlParser.elements(node)) {
                switch (el.localName) {
                    case "pPr":
                        this.parseParagraphProperties(el, result);
                        break;
                    case "r":
                        result.children.push(this.parseRun(el, result));
                        break;
                    case "hyperlink":
                        result.children.push(this.parseHyperlink(el, result));
                        break;
                    case "smartTag":
                        result.children.push(this.parseSmartTag(el, result));
                        break;
                    case "bookmarkStart":
                        result.children.push(parseBookmarkStart(el, globalXmlParser));
                        break;
                    case "bookmarkEnd":
                        result.children.push(parseBookmarkEnd(el, globalXmlParser));
                        break;
                    case "commentRangeStart":
                        result.children.push(new WmlCommentRangeStart(globalXmlParser.attr(el, "id")));
                        break;
                    case "commentRangeEnd":
                        result.children.push(new WmlCommentRangeEnd(globalXmlParser.attr(el, "id")));
                        break;
                    case "oMath":
                    case "oMathPara":
                        result.children.push(this.parseMathElement(el));
                        break;
                    case "sdt":
                        result.children.push(...this.parseSdt(el, e => this.parseParagraph(e).children));
                        break;
                    case "ins":
                        result.children.push(this.parseInserted(el, e => this.parseParagraph(e)));
                        break;
                    case "del":
                        result.children.push(this.parseDeleted(el, e => this.parseParagraph(e)));
                        break;
                }
            }
            return result;
        }
        parseParagraphProperties(elem, paragraph) {
            this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, c => {
                if (parseParagraphProperty(c, paragraph, globalXmlParser))
                    return true;
                switch (c.localName) {
                    case "pStyle":
                        paragraph.styleName = globalXmlParser.attr(c, "val");
                        break;
                    case "cnfStyle":
                        paragraph.className = values.classNameOfCnfStyle(c);
                        break;
                    case "framePr":
                        this.parseFrame(c, paragraph);
                        break;
                    case "rPr":
                        break;
                    default:
                        return false;
                }
                return true;
            });
        }
        parseFrame(node, paragraph) {
            var dropCap = globalXmlParser.attr(node, "dropCap");
            if (dropCap == "drop")
                paragraph.cssStyle["float"] = "left";
        }
        parseHyperlink(node, parent) {
            var result = { type: DomType.Hyperlink, parent: parent, children: [] };
            result.anchor = globalXmlParser.attr(node, "anchor");
            result.id = globalXmlParser.attr(node, "id");
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c, result));
                        break;
                }
            }
            return result;
        }
        parseSmartTag(node, parent) {
            var result = { type: DomType.SmartTag, parent, children: [] };
            var uri = globalXmlParser.attr(node, "uri");
            var element = globalXmlParser.attr(node, "element");
            if (uri)
                result.uri = uri;
            if (element)
                result.element = element;
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "r":
                        result.children.push(this.parseRun(c, result));
                        break;
                }
            }
            return result;
        }
        parseRun(node, parent) {
            var result = { type: DomType.Run, parent: parent, children: [] };
            for (let c of globalXmlParser.elements(node)) {
                c = this.checkAlternateContent(c);
                switch (c.localName) {
                    case "t":
                        result.children.push({
                            type: DomType.Text,
                            text: c.textContent
                        });
                        break;
                    case "delText":
                        result.children.push({
                            type: DomType.DeletedText,
                            text: c.textContent
                        });
                        break;
                    case "commentReference":
                        result.children.push(new WmlCommentReference(globalXmlParser.attr(c, "id")));
                        break;
                    case "fldSimple":
                        result.children.push({
                            type: DomType.SimpleField,
                            instruction: globalXmlParser.attr(c, "instr"),
                            lock: globalXmlParser.boolAttr(c, "lock", false),
                            dirty: globalXmlParser.boolAttr(c, "dirty", false)
                        });
                        break;
                    case "instrText":
                        result.fieldRun = true;
                        result.children.push({
                            type: DomType.Instruction,
                            text: c.textContent
                        });
                        break;
                    case "fldChar":
                        result.fieldRun = true;
                        result.children.push({
                            type: DomType.ComplexField,
                            charType: globalXmlParser.attr(c, "fldCharType"),
                            lock: globalXmlParser.boolAttr(c, "lock", false),
                            dirty: globalXmlParser.boolAttr(c, "dirty", false)
                        });
                        break;
                    case "noBreakHyphen":
                        result.children.push({ type: DomType.NoBreakHyphen });
                        break;
                    case "br":
                        result.children.push({
                            type: DomType.Break,
                            break: globalXmlParser.attr(c, "type") || "textWrapping"
                        });
                        break;
                    case "lastRenderedPageBreak":
                        result.children.push({
                            type: DomType.Break,
                            break: "lastRenderedPageBreak"
                        });
                        break;
                    case "sym":
                        result.children.push({
                            type: DomType.Symbol,
                            font: encloseFontFamily(globalXmlParser.attr(c, "font")),
                            char: globalXmlParser.attr(c, "char")
                        });
                        break;
                    case "tab":
                        result.children.push({ type: DomType.Tab });
                        break;
                    case "footnoteReference":
                        result.children.push({
                            type: DomType.FootnoteReference,
                            id: globalXmlParser.attr(c, "id")
                        });
                        break;
                    case "endnoteReference":
                        result.children.push({
                            type: DomType.EndnoteReference,
                            id: globalXmlParser.attr(c, "id")
                        });
                        break;
                    case "drawing":
                        let d = this.parseDrawing(c);
                        if (d)
                            result.children = [d];
                        break;
                    case "pict":
                        result.children.push(this.parseVmlPicture(c));
                        break;
                    case "rPr":
                        this.parseRunProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseMathElement(elem) {
            const propsTag = `${elem.localName}Pr`;
            const result = { type: mmlTagMap[elem.localName], children: [] };
            for (const el of globalXmlParser.elements(elem)) {
                const childType = mmlTagMap[el.localName];
                if (childType) {
                    result.children.push(this.parseMathElement(el));
                }
                else if (el.localName == "r") {
                    var run = this.parseRun(el);
                    run.type = DomType.MmlRun;
                    result.children.push(run);
                }
                else if (el.localName == propsTag) {
                    result.props = this.parseMathProperies(el);
                }
            }
            return result;
        }
        parseMathProperies(elem) {
            const result = {};
            for (const el of globalXmlParser.elements(elem)) {
                switch (el.localName) {
                    case "chr":
                        result.char = globalXmlParser.attr(el, "val");
                        break;
                    case "vertJc":
                        result.verticalJustification = globalXmlParser.attr(el, "val");
                        break;
                    case "pos":
                        result.position = globalXmlParser.attr(el, "val");
                        break;
                    case "degHide":
                        result.hideDegree = globalXmlParser.boolAttr(el, "val");
                        break;
                    case "begChr":
                        result.beginChar = globalXmlParser.attr(el, "val");
                        break;
                    case "endChr":
                        result.endChar = globalXmlParser.attr(el, "val");
                        break;
                }
            }
            return result;
        }
        parseRunProperties(elem, run) {
            this.parseDefaultProperties(elem, run.cssStyle = {}, null, c => {
                switch (c.localName) {
                    case "rStyle":
                        run.styleName = globalXmlParser.attr(c, "val");
                        break;
                    case "vertAlign":
                        run.verticalAlign = values.valueOfVertAlign(c, true);
                        break;
                    default:
                        return false;
                }
                return true;
            });
        }
        parseVmlPicture(elem) {
            const result = { type: DomType.VmlPicture, children: [] };
            for (const el of globalXmlParser.elements(elem)) {
                const child = parseVmlElement(el, this);
                child && result.children.push(child);
            }
            return result;
        }
        checkAlternateContent(elem) {
            if (elem.localName != 'AlternateContent')
                return elem;
            var choice = globalXmlParser.element(elem, "Choice");
            if (choice) {
                var requires = globalXmlParser.attr(choice, "Requires");
                var namespaceURI = elem.lookupNamespaceURI(requires);
                if (supportedNamespaceURIs.includes(namespaceURI))
                    return choice.firstElementChild;
            }
            return globalXmlParser.element(elem, "Fallback")?.firstElementChild;
        }
        parseDrawing(node) {
            for (var n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "inline":
                    case "anchor":
                        return this.parseDrawingWrapper(n);
                }
            }
        }
        parseDrawingWrapper(node) {
            var result = { type: DomType.Drawing, children: [], cssStyle: {} };
            var isAnchor = node.localName == "anchor";
            let wrapType = null;
            let simplePos = globalXmlParser.boolAttr(node, "simplePos");
            let behindDoc = globalXmlParser.boolAttr(node, "behindDoc");
            let docPrId = null;
            let posX = { relative: "page", align: "left", offset: "0" };
            let posY = { relative: "page", align: "top", offset: "0" };
            for (var n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "simplePos":
                        if (simplePos) {
                            posX.offset = globalXmlParser.lengthAttr(n, "x", LengthUsage.Emu);
                            posY.offset = globalXmlParser.lengthAttr(n, "y", LengthUsage.Emu);
                        }
                        break;
                    case "extent":
                        result.cssStyle["width"] = globalXmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
                        result.cssStyle["height"] = globalXmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
                        break;
                    case "docPr":
                        docPrId = globalXmlParser.attr(n, "id");
                        break;
                    case "positionH":
                    case "positionV":
                        if (!simplePos) {
                            let pos = n.localName == "positionH" ? posX : posY;
                            var alignNode = globalXmlParser.element(n, "align");
                            var offsetNode = globalXmlParser.element(n, "posOffset");
                            pos.relative = globalXmlParser.attr(n, "relativeFrom") ?? pos.relative;
                            if (alignNode)
                                pos.align = alignNode.textContent;
                            if (offsetNode)
                                pos.offset = convertLength(offsetNode.textContent, LengthUsage.Emu);
                        }
                        break;
                    case "wrapTopAndBottom":
                        wrapType = "wrapTopAndBottom";
                        break;
                    case "wrapNone":
                        wrapType = "wrapNone";
                        break;
                    case "graphic":
                        var g = this.parseGraphic(n);
                        if (g) {
                            if (docPrId && g.type === DomType.Image) {
                                g.docPrId = docPrId;
                            }
                            result.children.push(g);
                        }
                        break;
                }
            }
            if (wrapType == "wrapTopAndBottom") {
                result.cssStyle['display'] = 'block';
                if (posX.align) {
                    result.cssStyle['text-align'] = posX.align;
                    result.cssStyle['width'] = "100%";
                }
            }
            else if (wrapType == "wrapNone") {
                result.cssStyle['display'] = 'block';
                result.cssStyle['position'] = 'relative';
                result.cssStyle["width"] = "0px";
                result.cssStyle["height"] = "0px";
                if (posX.offset)
                    result.cssStyle["left"] = posX.offset;
                if (posY.offset)
                    result.cssStyle["top"] = posY.offset;
            }
            else if (isAnchor && (posX.align == 'left' || posX.align == 'right')) {
                result.cssStyle["float"] = posX.align;
            }
            result.props = {
                ...result.props,
                drawingAnchor: {
                    isAnchor,
                    wrapType,
                    behindDoc,
                    posX: { ...posX },
                    posY: { ...posY }
                }
            };
            return result;
        }
        parseGraphic(elem) {
            var graphicData = globalXmlParser.element(elem, "graphicData");
            for (let n of globalXmlParser.elements(graphicData)) {
                switch (n.localName) {
                    case "pic":
                        return this.parsePicture(n);
                }
            }
            return null;
        }
        parsePicture(elem) {
            var result = { type: DomType.Image, src: "", cssStyle: {} };
            var blipFill = globalXmlParser.element(elem, "blipFill");
            var blip = globalXmlParser.element(blipFill, "blip");
            var srcRect = globalXmlParser.element(blipFill, "srcRect");
            result.src = globalXmlParser.attr(blip, "embed");
            if (srcRect) {
                result.srcRect = [
                    globalXmlParser.intAttr(srcRect, "l", 0) / 100000,
                    globalXmlParser.intAttr(srcRect, "t", 0) / 100000,
                    globalXmlParser.intAttr(srcRect, "r", 0) / 100000,
                    globalXmlParser.intAttr(srcRect, "b", 0) / 100000,
                ];
            }
            var spPr = globalXmlParser.element(elem, "spPr");
            var xfrm = globalXmlParser.element(spPr, "xfrm");
            result.cssStyle["position"] = "relative";
            if (xfrm) {
                result.rotation = globalXmlParser.intAttr(xfrm, "rot", 0) / 60000;
                for (var n of globalXmlParser.elements(xfrm)) {
                    switch (n.localName) {
                        case "ext":
                            result.cssStyle["width"] = globalXmlParser.lengthAttr(n, "cx", LengthUsage.Emu);
                            result.cssStyle["height"] = globalXmlParser.lengthAttr(n, "cy", LengthUsage.Emu);
                            break;
                        case "off":
                            result.cssStyle["left"] = globalXmlParser.lengthAttr(n, "x", LengthUsage.Emu);
                            result.cssStyle["top"] = globalXmlParser.lengthAttr(n, "y", LengthUsage.Emu);
                            break;
                    }
                }
            }
            return result;
        }
        parseTable(node) {
            var result = { type: DomType.Table, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "tr":
                        result.children.push(this.parseTableRow(c));
                        break;
                    case "tblGrid":
                        result.columns = this.parseTableColumns(c);
                        break;
                    case "tblPr":
                        this.parseTableProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseTableColumns(node) {
            var result = [];
            for (const n of globalXmlParser.elements(node)) {
                switch (n.localName) {
                    case "gridCol":
                        result.push({ width: globalXmlParser.lengthAttr(n, "w") });
                        break;
                }
            }
            return result;
        }
        parseTableProperties(elem, table) {
            table.cssStyle = {};
            table.cellStyle = {};
            this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, c => {
                switch (c.localName) {
                    case "tblStyle":
                        table.styleName = globalXmlParser.attr(c, "val");
                        break;
                    case "tblLook":
                        table.className = values.classNameOftblLook(c);
                        break;
                    case "tblpPr":
                        this.parseTablePosition(c, table);
                        break;
                    case "tblStyleColBandSize":
                        table.colBandSize = globalXmlParser.intAttr(c, "val");
                        break;
                    case "tblStyleRowBandSize":
                        table.rowBandSize = globalXmlParser.intAttr(c, "val");
                        break;
                    case "hidden":
                        table.cssStyle["display"] = "none";
                        break;
                    default:
                        return false;
                }
                return true;
            });
            switch (table.cssStyle["text-align"]) {
                case "center":
                    delete table.cssStyle["text-align"];
                    table.cssStyle["margin-left"] = "auto";
                    table.cssStyle["margin-right"] = "auto";
                    break;
                case "right":
                    delete table.cssStyle["text-align"];
                    table.cssStyle["margin-left"] = "auto";
                    break;
            }
        }
        parseTablePosition(node, table) {
            var topFromText = globalXmlParser.lengthAttr(node, "topFromText");
            var bottomFromText = globalXmlParser.lengthAttr(node, "bottomFromText");
            var rightFromText = globalXmlParser.lengthAttr(node, "rightFromText");
            var leftFromText = globalXmlParser.lengthAttr(node, "leftFromText");
            table.cssStyle["float"] = 'left';
            table.cssStyle["margin-bottom"] = values.addSize(table.cssStyle["margin-bottom"], bottomFromText);
            table.cssStyle["margin-left"] = values.addSize(table.cssStyle["margin-left"], leftFromText);
            table.cssStyle["margin-right"] = values.addSize(table.cssStyle["margin-right"], rightFromText);
            table.cssStyle["margin-top"] = values.addSize(table.cssStyle["margin-top"], topFromText);
        }
        parseTableRow(node) {
            var result = { type: DomType.Row, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "tc":
                        result.children.push(this.parseTableCell(c));
                        break;
                    case "trPr":
                    case "tblPrEx":
                        this.parseTableRowProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseTableRowProperties(elem, row) {
            row.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
                switch (c.localName) {
                    case "cnfStyle":
                        row.className = values.classNameOfCnfStyle(c);
                        break;
                    case "tblHeader":
                        row.isHeader = globalXmlParser.boolAttr(c, "val");
                        break;
                    case "gridBefore":
                        row.gridBefore = globalXmlParser.intAttr(c, "val");
                        break;
                    case "gridAfter":
                        row.gridAfter = globalXmlParser.intAttr(c, "val");
                        break;
                    default:
                        return false;
                }
                return true;
            });
        }
        parseTableCell(node) {
            var result = { type: DomType.Cell, children: [] };
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "tbl":
                        result.children.push(this.parseTable(c));
                        break;
                    case "p":
                        result.children.push(this.parseParagraph(c));
                        break;
                    case "tcPr":
                        this.parseTableCellProperties(c, result);
                        break;
                }
            }
            return result;
        }
        parseTableCellProperties(elem, cell) {
            cell.cssStyle = this.parseDefaultProperties(elem, {}, null, c => {
                switch (c.localName) {
                    case "gridSpan":
                        cell.span = globalXmlParser.intAttr(c, "val", null);
                        break;
                    case "vMerge":
                        cell.verticalMerge = globalXmlParser.attr(c, "val") ?? "continue";
                        break;
                    case "cnfStyle":
                        cell.className = values.classNameOfCnfStyle(c);
                        break;
                    default:
                        return false;
                }
                return true;
            });
            this.parseTableCellVerticalText(elem, cell);
        }
        parseTableCellVerticalText(elem, cell) {
            const directionMap = {
                "btLr": {
                    writingMode: "vertical-rl",
                    transform: "rotate(180deg)"
                },
                "lrTb": {
                    writingMode: "vertical-lr",
                    transform: "none"
                },
                "tbRl": {
                    writingMode: "vertical-rl",
                    transform: "none"
                }
            };
            for (const c of globalXmlParser.elements(elem)) {
                if (c.localName === "textDirection") {
                    const direction = globalXmlParser.attr(c, "val");
                    const style = directionMap[direction] || { writingMode: "horizontal-tb" };
                    cell.cssStyle["writing-mode"] = style.writingMode;
                    cell.cssStyle["transform"] = style.transform;
                }
            }
        }
        parseDefaultProperties(elem, style = null, childStyle = null, handler = null) {
            style = style || {};
            for (const c of globalXmlParser.elements(elem)) {
                if (handler?.(c))
                    continue;
                switch (c.localName) {
                    case "jc":
                        style["text-align"] = values.valueOfJc(c);
                        break;
                    case "textAlignment":
                        style["vertical-align"] = values.valueOfTextAlignment(c);
                        break;
                    case "color":
                        style["color"] = xmlUtil.colorAttr(c, "val", null, autos.color);
                        break;
                    case "sz":
                        style["font-size"] = style["min-height"] = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
                        break;
                    case "shd":
                        style["background-color"] = xmlUtil.colorAttr(c, "fill", null, autos.shd);
                        break;
                    case "highlight":
                        style["background-color"] = xmlUtil.colorAttr(c, "val", null, autos.highlight);
                        break;
                    case "vertAlign":
                        break;
                    case "position":
                        style.verticalAlign = globalXmlParser.lengthAttr(c, "val", LengthUsage.FontSize);
                        break;
                    case "tcW":
                        if (this.options.ignoreWidth)
                            break;
                    case "tblW":
                        style["width"] = values.valueOfSize(c, "w");
                        break;
                    case "trHeight":
                        this.parseTrHeight(c, style);
                        break;
                    case "strike":
                        style["text-decoration"] = globalXmlParser.boolAttr(c, "val", true) ? "line-through" : "none";
                        break;
                    case "b":
                        style["font-weight"] = globalXmlParser.boolAttr(c, "val", true) ? "bold" : "normal";
                        break;
                    case "i":
                        style["font-style"] = globalXmlParser.boolAttr(c, "val", true) ? "italic" : "normal";
                        break;
                    case "caps":
                        style["text-transform"] = globalXmlParser.boolAttr(c, "val", true) ? "uppercase" : "none";
                        break;
                    case "smallCaps":
                        style["font-variant"] = globalXmlParser.boolAttr(c, "val", true) ? "small-caps" : "none";
                        break;
                    case "u":
                        this.parseUnderline(c, style);
                        break;
                    case "ind":
                    case "tblInd":
                        this.parseIndentation(c, style);
                        break;
                    case "rFonts":
                        this.parseFont(c, style);
                        break;
                    case "tblBorders":
                        this.parseBorderProperties(c, childStyle || style);
                        break;
                    case "tblCellSpacing":
                        style["border-spacing"] = values.valueOfMargin(c);
                        style["border-collapse"] = "separate";
                        break;
                    case "pBdr":
                        this.parseBorderProperties(c, style);
                        break;
                    case "bdr":
                        style["border"] = values.valueOfBorder(c);
                        break;
                    case "tcBorders":
                        this.parseBorderProperties(c, style);
                        break;
                    case "vanish":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["display"] = "none";
                        break;
                    case "kern":
                        break;
                    case "noWrap":
                        break;
                    case "tblCellMar":
                    case "tcMar":
                        this.parseMarginProperties(c, childStyle || style);
                        break;
                    case "tblLayout":
                        style["table-layout"] = values.valueOfTblLayout(c);
                        break;
                    case "vAlign":
                        style["vertical-align"] = values.valueOfTextAlignment(c);
                        break;
                    case "spacing":
                        if (elem.localName == "pPr")
                            this.parseSpacing(c, style);
                        break;
                    case "wordWrap":
                        if (globalXmlParser.boolAttr(c, "val"))
                            style["overflow-wrap"] = "break-word";
                        break;
                    case "suppressAutoHyphens":
                        style["hyphens"] = globalXmlParser.boolAttr(c, "val", true) ? "none" : "auto";
                        break;
                    case "lang":
                        style["$lang"] = globalXmlParser.attr(c, "val");
                        break;
                    case "rtl":
                    case "bidi":
                        if (globalXmlParser.boolAttr(c, "val", true))
                            style["direction"] = "rtl";
                        break;
                    case "bCs":
                    case "iCs":
                    case "szCs":
                    case "tabs":
                    case "outlineLvl":
                    case "contextualSpacing":
                    case "tblStyleColBandSize":
                    case "tblStyleRowBandSize":
                    case "webHidden":
                    case "pageBreakBefore":
                    case "suppressLineNumbers":
                    case "keepLines":
                    case "keepNext":
                    case "widowControl":
                    case "bidi":
                    case "rtl":
                    case "noProof":
                        break;
                    default:
                        if (this.options.debug)
                            console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`);
                        break;
                }
            }
            return style;
        }
        parseUnderline(node, style) {
            var val = globalXmlParser.attr(node, "val");
            if (val == null)
                return;
            switch (val) {
                case "dash":
                case "dashDotDotHeavy":
                case "dashDotHeavy":
                case "dashedHeavy":
                case "dashLong":
                case "dashLongHeavy":
                case "dotDash":
                case "dotDotDash":
                    style["text-decoration"] = "underline dashed";
                    break;
                case "dotted":
                case "dottedHeavy":
                    style["text-decoration"] = "underline dotted";
                    break;
                case "double":
                    style["text-decoration"] = "underline double";
                    break;
                case "single":
                case "thick":
                    style["text-decoration"] = "underline";
                    break;
                case "wave":
                case "wavyDouble":
                case "wavyHeavy":
                    style["text-decoration"] = "underline wavy";
                    break;
                case "words":
                    style["text-decoration"] = "underline";
                    break;
                case "none":
                    style["text-decoration"] = "none";
                    break;
            }
            var col = xmlUtil.colorAttr(node, "color");
            if (col)
                style["text-decoration-color"] = col;
        }
        parseFont(node, style) {
            var ascii = globalXmlParser.attr(node, "ascii");
            var asciiTheme = values.themeValue(node, "asciiTheme");
            var eastAsia = globalXmlParser.attr(node, "eastAsia");
            var fonts = [ascii, asciiTheme, eastAsia].filter(x => x).map(x => encloseFontFamily(x));
            if (fonts.length > 0)
                style["font-family"] = [...new Set(fonts)].join(', ');
        }
        parseIndentation(node, style) {
            var firstLine = globalXmlParser.lengthAttr(node, "firstLine");
            var hanging = globalXmlParser.lengthAttr(node, "hanging");
            var left = globalXmlParser.lengthAttr(node, "left");
            var start = globalXmlParser.lengthAttr(node, "start");
            var right = globalXmlParser.lengthAttr(node, "right");
            var end = globalXmlParser.lengthAttr(node, "end");
            if (firstLine)
                style["text-indent"] = firstLine;
            if (hanging)
                style["text-indent"] = `-${hanging}`;
            if (left || start)
                style["margin-inline-start"] = left || start;
            if (right || end)
                style["margin-inline-end"] = right || end;
        }
        parseSpacing(node, style) {
            var before = globalXmlParser.lengthAttr(node, "before");
            var after = globalXmlParser.lengthAttr(node, "after");
            var line = globalXmlParser.intAttr(node, "line", null);
            var lineRule = globalXmlParser.attr(node, "lineRule");
            if (before)
                style["margin-top"] = before;
            if (after)
                style["margin-bottom"] = after;
            if (line !== null) {
                switch (lineRule) {
                    case "auto":
                        style["line-height"] = `${(line / 240).toFixed(2)}`;
                        break;
                    case "atLeast":
                        style["line-height"] = `calc(100% + ${line / 20}pt)`;
                        break;
                    default:
                        style["line-height"] = style["min-height"] = `${line / 20}pt`;
                        break;
                }
            }
        }
        parseMarginProperties(node, output) {
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "left":
                        output["padding-left"] = values.valueOfMargin(c);
                        break;
                    case "right":
                        output["padding-right"] = values.valueOfMargin(c);
                        break;
                    case "top":
                        output["padding-top"] = values.valueOfMargin(c);
                        break;
                    case "bottom":
                        output["padding-bottom"] = values.valueOfMargin(c);
                        break;
                }
            }
        }
        parseTrHeight(node, output) {
            switch (globalXmlParser.attr(node, "hRule")) {
                case "exact":
                    output["height"] = globalXmlParser.lengthAttr(node, "val");
                    break;
                case "atLeast":
                default:
                    output["height"] = globalXmlParser.lengthAttr(node, "val");
                    break;
            }
        }
        parseBorderProperties(node, output) {
            for (const c of globalXmlParser.elements(node)) {
                switch (c.localName) {
                    case "start":
                    case "left":
                        output["border-left"] = values.valueOfBorder(c);
                        break;
                    case "end":
                    case "right":
                        output["border-right"] = values.valueOfBorder(c);
                        break;
                    case "top":
                        output["border-top"] = values.valueOfBorder(c);
                        break;
                    case "bottom":
                        output["border-bottom"] = values.valueOfBorder(c);
                        break;
                }
            }
        }
    }
    const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];
    class xmlUtil {
        static colorAttr(node, attrName, defValue = null, autoColor = 'black') {
            var v = globalXmlParser.attr(node, attrName);
            if (v) {
                if (v == "auto") {
                    return autoColor;
                }
                else if (knownColors.includes(v)) {
                    return v;
                }
                return `#${v}`;
            }
            var themeColor = globalXmlParser.attr(node, "themeColor");
            return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
        }
    }
    class values {
        static themeValue(c, attr) {
            var val = globalXmlParser.attr(c, attr);
            return val ? `var(--docx-${val}-font)` : null;
        }
        static valueOfSize(c, attr) {
            var type = LengthUsage.Dxa;
            switch (globalXmlParser.attr(c, "type")) {
                case "dxa": break;
                case "pct":
                    type = LengthUsage.Percent;
                    break;
                case "auto": return "auto";
            }
            return globalXmlParser.lengthAttr(c, attr, type);
        }
        static valueOfMargin(c) {
            return globalXmlParser.lengthAttr(c, "w");
        }
        static valueOfBorder(c) {
            var type = values.parseBorderType(globalXmlParser.attr(c, "val"));
            if (type == "none")
                return "none";
            var color = xmlUtil.colorAttr(c, "color");
            var size = globalXmlParser.lengthAttr(c, "sz", LengthUsage.Border);
            return `${size} ${type} ${color == "auto" ? autos.borderColor : color}`;
        }
        static parseBorderType(type) {
            switch (type) {
                case "single": return "solid";
                case "dashDotStroked": return "solid";
                case "dashed": return "dashed";
                case "dashSmallGap": return "dashed";
                case "dotDash": return "dotted";
                case "dotDotDash": return "dotted";
                case "dotted": return "dotted";
                case "double": return "double";
                case "doubleWave": return "double";
                case "inset": return "inset";
                case "nil": return "none";
                case "none": return "none";
                case "outset": return "outset";
                case "thick": return "solid";
                case "thickThinLargeGap": return "solid";
                case "thickThinMediumGap": return "solid";
                case "thickThinSmallGap": return "solid";
                case "thinThickLargeGap": return "solid";
                case "thinThickMediumGap": return "solid";
                case "thinThickSmallGap": return "solid";
                case "thinThickThinLargeGap": return "solid";
                case "thinThickThinMediumGap": return "solid";
                case "thinThickThinSmallGap": return "solid";
                case "threeDEmboss": return "solid";
                case "threeDEngrave": return "solid";
                case "triple": return "double";
                case "wave": return "solid";
            }
            return 'solid';
        }
        static valueOfTblLayout(c) {
            var type = globalXmlParser.attr(c, "val");
            return type == "fixed" ? "fixed" : "auto";
        }
        static classNameOfCnfStyle(c) {
            const val = globalXmlParser.attr(c, "val");
            const classes = [
                'first-row', 'last-row', 'first-col', 'last-col',
                'odd-col', 'even-col', 'odd-row', 'even-row',
                'ne-cell', 'nw-cell', 'se-cell', 'sw-cell'
            ];
            return classes.filter((_, i) => val[i] == '1').join(' ');
        }
        static valueOfJc(c) {
            var type = globalXmlParser.attr(c, "val");
            switch (type) {
                case "start":
                case "left": return "left";
                case "center": return "center";
                case "end":
                case "right": return "right";
                case "both": return "justify";
            }
            return type;
        }
        static valueOfVertAlign(c, asTagName = false) {
            var type = globalXmlParser.attr(c, "val");
            switch (type) {
                case "subscript": return "sub";
                case "superscript": return asTagName ? "sup" : "super";
            }
            return asTagName ? null : type;
        }
        static valueOfTextAlignment(c) {
            var type = globalXmlParser.attr(c, "val");
            switch (type) {
                case "auto":
                case "baseline": return "baseline";
                case "top": return "top";
                case "center": return "middle";
                case "bottom": return "bottom";
            }
            return type;
        }
        static addSize(a, b) {
            if (a == null)
                return b;
            if (b == null)
                return a;
            return `calc(${a} + ${b})`;
        }
        static classNameOftblLook(c) {
            const val = globalXmlParser.hexAttr(c, "val", 0);
            let className = "";
            if (globalXmlParser.boolAttr(c, "firstRow") || (val & 0x0020))
                className += " first-row";
            if (globalXmlParser.boolAttr(c, "lastRow") || (val & 0x0040))
                className += " last-row";
            if (globalXmlParser.boolAttr(c, "firstColumn") || (val & 0x0080))
                className += " first-col";
            if (globalXmlParser.boolAttr(c, "lastColumn") || (val & 0x0100))
                className += " last-col";
            if (globalXmlParser.boolAttr(c, "noHBand") || (val & 0x0200))
                className += " no-hband";
            if (globalXmlParser.boolAttr(c, "noVBand") || (val & 0x0400))
                className += " no-vband";
            return className.trim();
        }
    }

    const defaultTab = { pos: 0, leader: "none", style: "left" };
    const maxTabs = 50;
    function computePixelToPoint(container = document.body) {
        const temp = document.createElement("div");
        temp.style.width = '100pt';
        container.appendChild(temp);
        const result = 100 / temp.offsetWidth;
        container.removeChild(temp);
        return result;
    }
    function updateTabStop(elem, tabs, defaultTabSize, pixelToPoint = 72 / 96) {
        const p = elem.closest("p");
        const ebb = elem.getBoundingClientRect();
        const pbb = p.getBoundingClientRect();
        const pcs = getComputedStyle(p);
        const tabStops = tabs?.length > 0 ? tabs.map(t => ({
            pos: lengthToPoint(t.position),
            leader: t.leader,
            style: t.style
        })).sort((a, b) => a.pos - b.pos) : [defaultTab];
        const lastTab = tabStops[tabStops.length - 1];
        const pWidthPt = pbb.width * pixelToPoint;
        const size = lengthToPoint(defaultTabSize);
        let pos = lastTab.pos + size;
        if (pos < pWidthPt) {
            for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) {
                tabStops.push({ ...defaultTab, pos: pos });
            }
        }
        const marginLeft = parseFloat(pcs.marginLeft);
        const pOffset = pbb.left + marginLeft;
        const left = (ebb.left - pOffset) * pixelToPoint;
        const tab = tabStops.find(t => t.style != "clear" && t.pos > left);
        if (tab == null)
            return;
        let width = 1;
        if (tab.style == "right" || tab.style == "center") {
            const tabStops = Array.from(p.querySelectorAll(`.${elem.className}`));
            const nextIdx = tabStops.indexOf(elem) + 1;
            const range = document.createRange();
            range.setStart(elem, 1);
            if (nextIdx < tabStops.length) {
                range.setEndBefore(tabStops[nextIdx]);
            }
            else {
                range.setEndAfter(p);
            }
            const mul = tab.style == "center" ? 0.5 : 1;
            const nextBB = range.getBoundingClientRect();
            const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);
            width = tab.pos - offset * pixelToPoint;
        }
        else {
            width = tab.pos - left;
        }
        elem.innerHTML = "&nbsp;";
        elem.style.textDecoration = "inherit";
        elem.style.wordSpacing = `${width.toFixed(0)}pt`;
        switch (tab.leader) {
            case "dot":
            case "middleDot":
                elem.style.textDecoration = "underline";
                elem.style.textDecorationStyle = "dotted";
                break;
            case "hyphen":
            case "heavy":
            case "underscore":
                elem.style.textDecoration = "underline";
                break;
        }
    }
    function lengthToPoint(length) {
        return parseFloat(length);
    }

    function memo(getDeps, fn, opts) {
      let deps = opts.initialDeps ?? [];
      let result;
      let isInitial = true;
      function memoizedFunction() {
        var _a, _b, _c;
        let depTime;
        if (opts.key && ((_a = opts.debug) == null ? void 0 : _a.call(opts))) depTime = Date.now();
        const newDeps = getDeps();
        const depsChanged = newDeps.length !== deps.length || newDeps.some((dep, index) => deps[index] !== dep);
        if (!depsChanged) {
          return result;
        }
        deps = newDeps;
        let resultTime;
        if (opts.key && ((_b = opts.debug) == null ? void 0 : _b.call(opts))) resultTime = Date.now();
        result = fn(...newDeps);
        if (opts.key && ((_c = opts.debug) == null ? void 0 : _c.call(opts))) {
          const depEndTime = Math.round((Date.now() - depTime) * 100) / 100;
          const resultEndTime = Math.round((Date.now() - resultTime) * 100) / 100;
          const resultFpsPercentage = resultEndTime / 16;
          const pad = (str, num) => {
            str = String(str);
            while (str.length < num) {
              str = " " + str;
            }
            return str;
          };
          console.info(
            `%c⏱ ${pad(resultEndTime, 5)} /${pad(depEndTime, 5)} ms`,
            `
            font-size: .6rem;
            font-weight: bold;
            color: hsl(${Math.max(
          0,
          Math.min(120 - 120 * resultFpsPercentage, 120)
        )}deg 100% 31%);`,
            opts == null ? void 0 : opts.key
          );
        }
        if ((opts == null ? void 0 : opts.onChange) && !(isInitial && opts.skipInitialOnChange)) {
          opts.onChange(result);
        }
        isInitial = false;
        return result;
      }
      memoizedFunction.updateDeps = (newDeps) => {
        deps = newDeps;
      };
      return memoizedFunction;
    }
    function notUndefined(value, msg) {
      if (value === void 0) {
        throw new Error(`Unexpected undefined${""}`);
      } else {
        return value;
      }
    }
    const approxEqual = (a, b) => Math.abs(a - b) < 1.01;
    const debounce = (targetWindow, fn, ms) => {
      let timeoutId;
      return function(...args) {
        targetWindow.clearTimeout(timeoutId);
        timeoutId = targetWindow.setTimeout(() => fn.apply(this, args), ms);
      };
    };

    const getRect = (element) => {
      const { offsetWidth, offsetHeight } = element;
      return { width: offsetWidth, height: offsetHeight };
    };
    const defaultKeyExtractor = (index) => index;
    const defaultRangeExtractor = (range) => {
      const start = Math.max(range.startIndex - range.overscan, 0);
      const end = Math.min(range.endIndex + range.overscan, range.count - 1);
      const arr = [];
      for (let i = start; i <= end; i++) {
        arr.push(i);
      }
      return arr;
    };
    const observeElementRect = (instance, cb) => {
      const element = instance.scrollElement;
      if (!element) {
        return;
      }
      const targetWindow = instance.targetWindow;
      if (!targetWindow) {
        return;
      }
      const handler = (rect) => {
        const { width, height } = rect;
        cb({ width: Math.round(width), height: Math.round(height) });
      };
      handler(getRect(element));
      if (!targetWindow.ResizeObserver) {
        return () => {
        };
      }
      const observer = new targetWindow.ResizeObserver((entries) => {
        const run = () => {
          const entry = entries[0];
          if (entry == null ? void 0 : entry.borderBoxSize) {
            const box = entry.borderBoxSize[0];
            if (box) {
              handler({ width: box.inlineSize, height: box.blockSize });
              return;
            }
          }
          handler(getRect(element));
        };
        instance.options.useAnimationFrameWithResizeObserver ? requestAnimationFrame(run) : run();
      });
      observer.observe(element, { box: "border-box" });
      return () => {
        observer.unobserve(element);
      };
    };
    const addEventListenerOptions = {
      passive: true
    };
    const supportsScrollend = typeof window == "undefined" ? true : "onscrollend" in window;
    const observeElementOffset = (instance, cb) => {
      const element = instance.scrollElement;
      if (!element) {
        return;
      }
      const targetWindow = instance.targetWindow;
      if (!targetWindow) {
        return;
      }
      let offset = 0;
      const fallback = instance.options.useScrollendEvent && supportsScrollend ? () => void 0 : debounce(
        targetWindow,
        () => {
          cb(offset, false);
        },
        instance.options.isScrollingResetDelay
      );
      const createHandler = (isScrolling) => () => {
        const { horizontal, isRtl } = instance.options;
        offset = horizontal ? element["scrollLeft"] * (isRtl && -1 || 1) : element["scrollTop"];
        fallback();
        cb(offset, isScrolling);
      };
      const handler = createHandler(true);
      const endHandler = createHandler(false);
      element.addEventListener("scroll", handler, addEventListenerOptions);
      const registerScrollendEvent = instance.options.useScrollendEvent && supportsScrollend;
      if (registerScrollendEvent) {
        element.addEventListener("scrollend", endHandler, addEventListenerOptions);
      }
      return () => {
        element.removeEventListener("scroll", handler);
        if (registerScrollendEvent) {
          element.removeEventListener("scrollend", endHandler);
        }
      };
    };
    const measureElement = (element, entry, instance) => {
      if (entry == null ? void 0 : entry.borderBoxSize) {
        const box = entry.borderBoxSize[0];
        if (box) {
          const size = Math.round(
            box[instance.options.horizontal ? "inlineSize" : "blockSize"]
          );
          return size;
        }
      }
      return element[instance.options.horizontal ? "offsetWidth" : "offsetHeight"];
    };
    const elementScroll = (offset, {
      adjustments = 0,
      behavior
    }, instance) => {
      var _a, _b;
      const toOffset = offset + adjustments;
      (_b = (_a = instance.scrollElement) == null ? void 0 : _a.scrollTo) == null ? void 0 : _b.call(_a, {
        [instance.options.horizontal ? "left" : "top"]: toOffset,
        behavior
      });
    };
    class Virtualizer {
      constructor(opts) {
        this.unsubs = [];
        this.scrollElement = null;
        this.targetWindow = null;
        this.isScrolling = false;
        this.scrollState = null;
        this.measurementsCache = [];
        this.itemSizeCache = /* @__PURE__ */ new Map();
        this.laneAssignments = /* @__PURE__ */ new Map();
        this.pendingMeasuredCacheIndexes = [];
        this.prevLanes = void 0;
        this.lanesChangedFlag = false;
        this.lanesSettling = false;
        this.scrollRect = null;
        this.scrollOffset = null;
        this.scrollDirection = null;
        this.scrollAdjustments = 0;
        this.elementsCache = /* @__PURE__ */ new Map();
        this.now = () => {
          var _a, _b, _c;
          return ((_c = (_b = (_a = this.targetWindow) == null ? void 0 : _a.performance) == null ? void 0 : _b.now) == null ? void 0 : _c.call(_b)) ?? Date.now();
        };
        this.observer = /* @__PURE__ */ (() => {
          let _ro = null;
          const get = () => {
            if (_ro) {
              return _ro;
            }
            if (!this.targetWindow || !this.targetWindow.ResizeObserver) {
              return null;
            }
            return _ro = new this.targetWindow.ResizeObserver((entries) => {
              entries.forEach((entry) => {
                const run = () => {
                  const node = entry.target;
                  const index = this.indexFromElement(node);
                  if (!node.isConnected) {
                    this.observer.unobserve(node);
                    return;
                  }
                  if (this.shouldMeasureDuringScroll(index)) {
                    this.resizeItem(
                      index,
                      this.options.measureElement(node, entry, this)
                    );
                  }
                };
                this.options.useAnimationFrameWithResizeObserver ? requestAnimationFrame(run) : run();
              });
            });
          };
          return {
            disconnect: () => {
              var _a;
              (_a = get()) == null ? void 0 : _a.disconnect();
              _ro = null;
            },
            observe: (target) => {
              var _a;
              return (_a = get()) == null ? void 0 : _a.observe(target, { box: "border-box" });
            },
            unobserve: (target) => {
              var _a;
              return (_a = get()) == null ? void 0 : _a.unobserve(target);
            }
          };
        })();
        this.range = null;
        this.setOptions = (opts2) => {
          Object.entries(opts2).forEach(([key, value]) => {
            if (typeof value === "undefined") delete opts2[key];
          });
          this.options = {
            debug: false,
            initialOffset: 0,
            overscan: 1,
            paddingStart: 0,
            paddingEnd: 0,
            scrollPaddingStart: 0,
            scrollPaddingEnd: 0,
            horizontal: false,
            getItemKey: defaultKeyExtractor,
            rangeExtractor: defaultRangeExtractor,
            onChange: () => {
            },
            measureElement,
            initialRect: { width: 0, height: 0 },
            scrollMargin: 0,
            gap: 0,
            indexAttribute: "data-index",
            initialMeasurementsCache: [],
            lanes: 1,
            isScrollingResetDelay: 150,
            enabled: true,
            isRtl: false,
            useScrollendEvent: false,
            useAnimationFrameWithResizeObserver: false,
            ...opts2
          };
        };
        this.notify = (sync) => {
          var _a, _b;
          (_b = (_a = this.options).onChange) == null ? void 0 : _b.call(_a, this, sync);
        };
        this.maybeNotify = memo(
          () => {
            this.calculateRange();
            return [
              this.isScrolling,
              this.range ? this.range.startIndex : null,
              this.range ? this.range.endIndex : null
            ];
          },
          (isScrolling) => {
            this.notify(isScrolling);
          },
          {
            key: "production" !== "production",
            debug: () => this.options.debug,
            initialDeps: [
              this.isScrolling,
              this.range ? this.range.startIndex : null,
              this.range ? this.range.endIndex : null
            ]
          }
        );
        this.cleanup = () => {
          this.unsubs.filter(Boolean).forEach((d) => d());
          this.unsubs = [];
          this.observer.disconnect();
          if (this.rafId != null && this.targetWindow) {
            this.targetWindow.cancelAnimationFrame(this.rafId);
            this.rafId = null;
          }
          this.scrollState = null;
          this.scrollElement = null;
          this.targetWindow = null;
        };
        this._didMount = () => {
          return () => {
            this.cleanup();
          };
        };
        this._willUpdate = () => {
          var _a;
          const scrollElement = this.options.enabled ? this.options.getScrollElement() : null;
          if (this.scrollElement !== scrollElement) {
            this.cleanup();
            if (!scrollElement) {
              this.maybeNotify();
              return;
            }
            this.scrollElement = scrollElement;
            if (this.scrollElement && "ownerDocument" in this.scrollElement) {
              this.targetWindow = this.scrollElement.ownerDocument.defaultView;
            } else {
              this.targetWindow = ((_a = this.scrollElement) == null ? void 0 : _a.window) ?? null;
            }
            this.elementsCache.forEach((cached) => {
              this.observer.observe(cached);
            });
            this.unsubs.push(
              this.options.observeElementRect(this, (rect) => {
                this.scrollRect = rect;
                this.maybeNotify();
              })
            );
            this.unsubs.push(
              this.options.observeElementOffset(this, (offset, isScrolling) => {
                this.scrollAdjustments = 0;
                this.scrollDirection = isScrolling ? this.getScrollOffset() < offset ? "forward" : "backward" : null;
                this.scrollOffset = offset;
                this.isScrolling = isScrolling;
                if (this.scrollState) {
                  this.scheduleScrollReconcile();
                }
                this.maybeNotify();
              })
            );
            this._scrollToOffset(this.getScrollOffset(), {
              adjustments: void 0,
              behavior: void 0
            });
          }
        };
        this.rafId = null;
        this.getSize = () => {
          if (!this.options.enabled) {
            this.scrollRect = null;
            return 0;
          }
          this.scrollRect = this.scrollRect ?? this.options.initialRect;
          return this.scrollRect[this.options.horizontal ? "width" : "height"];
        };
        this.getScrollOffset = () => {
          if (!this.options.enabled) {
            this.scrollOffset = null;
            return 0;
          }
          this.scrollOffset = this.scrollOffset ?? (typeof this.options.initialOffset === "function" ? this.options.initialOffset() : this.options.initialOffset);
          return this.scrollOffset;
        };
        this.getFurthestMeasurement = (measurements, index) => {
          const furthestMeasurementsFound = /* @__PURE__ */ new Map();
          const furthestMeasurements = /* @__PURE__ */ new Map();
          for (let m = index - 1; m >= 0; m--) {
            const measurement = measurements[m];
            if (furthestMeasurementsFound.has(measurement.lane)) {
              continue;
            }
            const previousFurthestMeasurement = furthestMeasurements.get(
              measurement.lane
            );
            if (previousFurthestMeasurement == null || measurement.end > previousFurthestMeasurement.end) {
              furthestMeasurements.set(measurement.lane, measurement);
            } else if (measurement.end < previousFurthestMeasurement.end) {
              furthestMeasurementsFound.set(measurement.lane, true);
            }
            if (furthestMeasurementsFound.size === this.options.lanes) {
              break;
            }
          }
          return furthestMeasurements.size === this.options.lanes ? Array.from(furthestMeasurements.values()).sort((a, b) => {
            if (a.end === b.end) {
              return a.index - b.index;
            }
            return a.end - b.end;
          })[0] : void 0;
        };
        this.getMeasurementOptions = memo(
          () => [
            this.options.count,
            this.options.paddingStart,
            this.options.scrollMargin,
            this.options.getItemKey,
            this.options.enabled,
            this.options.lanes
          ],
          (count, paddingStart, scrollMargin, getItemKey, enabled, lanes) => {
            const lanesChanged = this.prevLanes !== void 0 && this.prevLanes !== lanes;
            if (lanesChanged) {
              this.lanesChangedFlag = true;
            }
            this.prevLanes = lanes;
            this.pendingMeasuredCacheIndexes = [];
            return {
              count,
              paddingStart,
              scrollMargin,
              getItemKey,
              enabled,
              lanes
            };
          },
          {
            key: false
          }
        );
        this.getMeasurements = memo(
          () => [this.getMeasurementOptions(), this.itemSizeCache],
          ({ count, paddingStart, scrollMargin, getItemKey, enabled, lanes }, itemSizeCache) => {
            if (!enabled) {
              this.measurementsCache = [];
              this.itemSizeCache.clear();
              this.laneAssignments.clear();
              return [];
            }
            if (this.laneAssignments.size > count) {
              for (const index of this.laneAssignments.keys()) {
                if (index >= count) {
                  this.laneAssignments.delete(index);
                }
              }
            }
            if (this.lanesChangedFlag) {
              this.lanesChangedFlag = false;
              this.lanesSettling = true;
              this.measurementsCache = [];
              this.itemSizeCache.clear();
              this.laneAssignments.clear();
              this.pendingMeasuredCacheIndexes = [];
            }
            if (this.measurementsCache.length === 0 && !this.lanesSettling) {
              this.measurementsCache = this.options.initialMeasurementsCache;
              this.measurementsCache.forEach((item) => {
                this.itemSizeCache.set(item.key, item.size);
              });
            }
            const min = this.lanesSettling ? 0 : this.pendingMeasuredCacheIndexes.length > 0 ? Math.min(...this.pendingMeasuredCacheIndexes) : 0;
            this.pendingMeasuredCacheIndexes = [];
            if (this.lanesSettling && this.measurementsCache.length === count) {
              this.lanesSettling = false;
            }
            const measurements = this.measurementsCache.slice(0, min);
            const laneLastIndex = new Array(lanes).fill(
              void 0
            );
            for (let m = 0; m < min; m++) {
              const item = measurements[m];
              if (item) {
                laneLastIndex[item.lane] = m;
              }
            }
            for (let i = min; i < count; i++) {
              const key = getItemKey(i);
              const cachedLane = this.laneAssignments.get(i);
              let lane;
              let start;
              if (cachedLane !== void 0 && this.options.lanes > 1) {
                lane = cachedLane;
                const prevIndex = laneLastIndex[lane];
                const prevInLane = prevIndex !== void 0 ? measurements[prevIndex] : void 0;
                start = prevInLane ? prevInLane.end + this.options.gap : paddingStart + scrollMargin;
              } else {
                const furthestMeasurement = this.options.lanes === 1 ? measurements[i - 1] : this.getFurthestMeasurement(measurements, i);
                start = furthestMeasurement ? furthestMeasurement.end + this.options.gap : paddingStart + scrollMargin;
                lane = furthestMeasurement ? furthestMeasurement.lane : i % this.options.lanes;
                if (this.options.lanes > 1) {
                  this.laneAssignments.set(i, lane);
                }
              }
              const measuredSize = itemSizeCache.get(key);
              const size = typeof measuredSize === "number" ? measuredSize : this.options.estimateSize(i);
              const end = start + size;
              measurements[i] = {
                index: i,
                start,
                size,
                end,
                key,
                lane
              };
              laneLastIndex[lane] = i;
            }
            this.measurementsCache = measurements;
            return measurements;
          },
          {
            key: "production" !== "production",
            debug: () => this.options.debug
          }
        );
        this.calculateRange = memo(
          () => [
            this.getMeasurements(),
            this.getSize(),
            this.getScrollOffset(),
            this.options.lanes
          ],
          (measurements, outerSize, scrollOffset, lanes) => {
            return this.range = measurements.length > 0 && outerSize > 0 ? calculateRange({
              measurements,
              outerSize,
              scrollOffset,
              lanes
            }) : null;
          },
          {
            key: "production" !== "production",
            debug: () => this.options.debug
          }
        );
        this.getVirtualIndexes = memo(
          () => {
            let startIndex = null;
            let endIndex = null;
            const range = this.calculateRange();
            if (range) {
              startIndex = range.startIndex;
              endIndex = range.endIndex;
            }
            this.maybeNotify.updateDeps([this.isScrolling, startIndex, endIndex]);
            return [
              this.options.rangeExtractor,
              this.options.overscan,
              this.options.count,
              startIndex,
              endIndex
            ];
          },
          (rangeExtractor, overscan, count, startIndex, endIndex) => {
            return startIndex === null || endIndex === null ? [] : rangeExtractor({
              startIndex,
              endIndex,
              overscan,
              count
            });
          },
          {
            key: "production" !== "production",
            debug: () => this.options.debug
          }
        );
        this.indexFromElement = (node) => {
          const attributeName = this.options.indexAttribute;
          const indexStr = node.getAttribute(attributeName);
          if (!indexStr) {
            console.warn(
              `Missing attribute name '${attributeName}={index}' on measured element.`
            );
            return -1;
          }
          return parseInt(indexStr, 10);
        };
        this.shouldMeasureDuringScroll = (index) => {
          var _a;
          if (!this.scrollState || this.scrollState.behavior !== "smooth") {
            return true;
          }
          const scrollIndex = this.scrollState.index ?? ((_a = this.getVirtualItemForOffset(this.scrollState.lastTargetOffset)) == null ? void 0 : _a.index);
          if (scrollIndex !== void 0 && this.range) {
            const bufferSize = Math.max(
              this.options.overscan,
              Math.ceil((this.range.endIndex - this.range.startIndex) / 2)
            );
            const minIndex = Math.max(0, scrollIndex - bufferSize);
            const maxIndex = Math.min(
              this.options.count - 1,
              scrollIndex + bufferSize
            );
            return index >= minIndex && index <= maxIndex;
          }
          return true;
        };
        this.measureElement = (node) => {
          if (!node) {
            this.elementsCache.forEach((cached, key2) => {
              if (!cached.isConnected) {
                this.observer.unobserve(cached);
                this.elementsCache.delete(key2);
              }
            });
            return;
          }
          const index = this.indexFromElement(node);
          const key = this.options.getItemKey(index);
          const prevNode = this.elementsCache.get(key);
          if (prevNode !== node) {
            if (prevNode) {
              this.observer.unobserve(prevNode);
            }
            this.observer.observe(node);
            this.elementsCache.set(key, node);
          }
          if ((!this.isScrolling || this.scrollState) && this.shouldMeasureDuringScroll(index)) {
            this.resizeItem(index, this.options.measureElement(node, void 0, this));
          }
        };
        this.resizeItem = (index, size) => {
          var _a;
          const item = this.measurementsCache[index];
          if (!item) return;
          const itemSize = this.itemSizeCache.get(item.key) ?? item.size;
          const delta = size - itemSize;
          if (delta !== 0) {
            if (((_a = this.scrollState) == null ? void 0 : _a.behavior) !== "smooth" && (this.shouldAdjustScrollPositionOnItemSizeChange !== void 0 ? this.shouldAdjustScrollPositionOnItemSizeChange(item, delta, this) : item.start < this.getScrollOffset() + this.scrollAdjustments)) {
              this._scrollToOffset(this.getScrollOffset(), {
                adjustments: this.scrollAdjustments += delta,
                behavior: void 0
              });
            }
            this.pendingMeasuredCacheIndexes.push(item.index);
            this.itemSizeCache = new Map(this.itemSizeCache.set(item.key, size));
            this.notify(false);
          }
        };
        this.getVirtualItems = memo(
          () => [this.getVirtualIndexes(), this.getMeasurements()],
          (indexes, measurements) => {
            const virtualItems = [];
            for (let k = 0, len = indexes.length; k < len; k++) {
              const i = indexes[k];
              const measurement = measurements[i];
              virtualItems.push(measurement);
            }
            return virtualItems;
          },
          {
            key: "production" !== "production",
            debug: () => this.options.debug
          }
        );
        this.getVirtualItemForOffset = (offset) => {
          const measurements = this.getMeasurements();
          if (measurements.length === 0) {
            return void 0;
          }
          return notUndefined(
            measurements[findNearestBinarySearch(
              0,
              measurements.length - 1,
              (index) => notUndefined(measurements[index]).start,
              offset
            )]
          );
        };
        this.getMaxScrollOffset = () => {
          if (!this.scrollElement) return 0;
          if ("scrollHeight" in this.scrollElement) {
            return this.options.horizontal ? this.scrollElement.scrollWidth - this.scrollElement.clientWidth : this.scrollElement.scrollHeight - this.scrollElement.clientHeight;
          } else {
            const doc = this.scrollElement.document.documentElement;
            return this.options.horizontal ? doc.scrollWidth - this.scrollElement.innerWidth : doc.scrollHeight - this.scrollElement.innerHeight;
          }
        };
        this.getOffsetForAlignment = (toOffset, align, itemSize = 0) => {
          if (!this.scrollElement) return 0;
          const size = this.getSize();
          const scrollOffset = this.getScrollOffset();
          if (align === "auto") {
            align = toOffset >= scrollOffset + size ? "end" : "start";
          }
          if (align === "center") {
            toOffset += (itemSize - size) / 2;
          } else if (align === "end") {
            toOffset -= size;
          }
          const maxOffset = this.getMaxScrollOffset();
          return Math.max(Math.min(maxOffset, toOffset), 0);
        };
        this.getOffsetForIndex = (index, align = "auto") => {
          index = Math.max(0, Math.min(index, this.options.count - 1));
          const size = this.getSize();
          const scrollOffset = this.getScrollOffset();
          const item = this.measurementsCache[index];
          if (!item) return;
          if (align === "auto") {
            if (item.end >= scrollOffset + size - this.options.scrollPaddingEnd) {
              align = "end";
            } else if (item.start <= scrollOffset + this.options.scrollPaddingStart) {
              align = "start";
            } else {
              return [scrollOffset, align];
            }
          }
          if (align === "end" && index === this.options.count - 1) {
            return [this.getMaxScrollOffset(), align];
          }
          const toOffset = align === "end" ? item.end + this.options.scrollPaddingEnd : item.start - this.options.scrollPaddingStart;
          return [
            this.getOffsetForAlignment(toOffset, align, item.size),
            align
          ];
        };
        this.scrollToOffset = (toOffset, { align = "start", behavior = "auto" } = {}) => {
          const offset = this.getOffsetForAlignment(toOffset, align);
          const now = this.now();
          this.scrollState = {
            index: null,
            align,
            behavior,
            startedAt: now,
            lastTargetOffset: offset,
            stableFrames: 0
          };
          this._scrollToOffset(offset, { adjustments: void 0, behavior });
          this.scheduleScrollReconcile();
        };
        this.scrollToIndex = (index, {
          align: initialAlign = "auto",
          behavior = "auto"
        } = {}) => {
          index = Math.max(0, Math.min(index, this.options.count - 1));
          const offsetInfo = this.getOffsetForIndex(index, initialAlign);
          if (!offsetInfo) {
            return;
          }
          const [offset, align] = offsetInfo;
          const now = this.now();
          this.scrollState = {
            index,
            align,
            behavior,
            startedAt: now,
            lastTargetOffset: offset,
            stableFrames: 0
          };
          this._scrollToOffset(offset, { adjustments: void 0, behavior });
          this.scheduleScrollReconcile();
        };
        this.scrollBy = (delta, { behavior = "auto" } = {}) => {
          const offset = this.getScrollOffset() + delta;
          const now = this.now();
          this.scrollState = {
            index: null,
            align: "start",
            behavior,
            startedAt: now,
            lastTargetOffset: offset,
            stableFrames: 0
          };
          this._scrollToOffset(offset, { adjustments: void 0, behavior });
          this.scheduleScrollReconcile();
        };
        this.getTotalSize = () => {
          var _a;
          const measurements = this.getMeasurements();
          let end;
          if (measurements.length === 0) {
            end = this.options.paddingStart;
          } else if (this.options.lanes === 1) {
            end = ((_a = measurements[measurements.length - 1]) == null ? void 0 : _a.end) ?? 0;
          } else {
            const endByLane = Array(this.options.lanes).fill(null);
            let endIndex = measurements.length - 1;
            while (endIndex >= 0 && endByLane.some((val) => val === null)) {
              const item = measurements[endIndex];
              if (endByLane[item.lane] === null) {
                endByLane[item.lane] = item.end;
              }
              endIndex--;
            }
            end = Math.max(...endByLane.filter((val) => val !== null));
          }
          return Math.max(
            end - this.options.scrollMargin + this.options.paddingEnd,
            0
          );
        };
        this._scrollToOffset = (offset, {
          adjustments,
          behavior
        }) => {
          this.options.scrollToFn(offset, { behavior, adjustments }, this);
        };
        this.measure = () => {
          this.itemSizeCache = /* @__PURE__ */ new Map();
          this.laneAssignments = /* @__PURE__ */ new Map();
          this.notify(false);
        };
        this.setOptions(opts);
      }
      scheduleScrollReconcile() {
        if (!this.targetWindow) {
          this.scrollState = null;
          return;
        }
        if (this.rafId != null) return;
        this.rafId = this.targetWindow.requestAnimationFrame(() => {
          this.rafId = null;
          this.reconcileScroll();
        });
      }
      reconcileScroll() {
        if (!this.scrollState) return;
        const el = this.scrollElement;
        if (!el) return;
        const MAX_RECONCILE_MS = 5e3;
        if (this.now() - this.scrollState.startedAt > MAX_RECONCILE_MS) {
          this.scrollState = null;
          return;
        }
        const offsetInfo = this.scrollState.index != null ? this.getOffsetForIndex(this.scrollState.index, this.scrollState.align) : void 0;
        const targetOffset = offsetInfo ? offsetInfo[0] : this.scrollState.lastTargetOffset;
        const STABLE_FRAMES = 1;
        const targetChanged = targetOffset !== this.scrollState.lastTargetOffset;
        if (!targetChanged && approxEqual(targetOffset, this.getScrollOffset())) {
          this.scrollState.stableFrames++;
          if (this.scrollState.stableFrames >= STABLE_FRAMES) {
            this.scrollState = null;
            return;
          }
        } else {
          this.scrollState.stableFrames = 0;
          if (targetChanged) {
            this.scrollState.lastTargetOffset = targetOffset;
            this.scrollState.behavior = "auto";
            this._scrollToOffset(targetOffset, {
              adjustments: void 0,
              behavior: "auto"
            });
          }
        }
        this.scheduleScrollReconcile();
      }
    }
    const findNearestBinarySearch = (low, high, getCurrentValue, value) => {
      while (low <= high) {
        const middle = (low + high) / 2 | 0;
        const currentValue = getCurrentValue(middle);
        if (currentValue < value) {
          low = middle + 1;
        } else if (currentValue > value) {
          high = middle - 1;
        } else {
          return middle;
        }
      }
      if (low > 0) {
        return low - 1;
      } else {
        return 0;
      }
    };
    function calculateRange({
      measurements,
      outerSize,
      scrollOffset,
      lanes
    }) {
      const lastIndex = measurements.length - 1;
      const getOffset = (index) => measurements[index].start;
      if (measurements.length <= lanes) {
        return {
          startIndex: 0,
          endIndex: lastIndex
        };
      }
      let startIndex = findNearestBinarySearch(
        0,
        lastIndex,
        getOffset,
        scrollOffset
      );
      let endIndex = startIndex;
      if (lanes === 1) {
        while (endIndex < lastIndex && measurements[endIndex].end < scrollOffset + outerSize) {
          endIndex++;
        }
      } else if (lanes > 1) {
        const endPerLane = Array(lanes).fill(0);
        while (endIndex < lastIndex && endPerLane.some((pos) => pos < scrollOffset + outerSize)) {
          const item = measurements[endIndex];
          endPerLane[item.lane] = item.end;
          endIndex++;
        }
        const startPerLane = Array(lanes).fill(scrollOffset + outerSize);
        while (startIndex >= 0 && startPerLane.some((pos) => pos >= scrollOffset)) {
          const item = measurements[startIndex];
          startPerLane[item.lane] = item.start;
          startIndex--;
        }
        startIndex = Math.max(0, startIndex - startIndex % lanes);
        endIndex = Math.min(lastIndex, endIndex + (lanes - 1 - endIndex % lanes));
      }
      return { startIndex, endIndex };
    }

    class VirtualizedRenderer {
        constructor(options) {
            this.options = options;
            this.cleanup = null;
            this.elementCache = new Map();
            this.pendingMeasureIndices = new Set();
            this.frameId = 0;
            this.pendingMeasureFrameId = 0;
            this.idleMeasureTimeoutId = 0;
            this.idleMeasureFrameId = 0;
            this.lastWindowSignature = null;
            this.isScrolling = false;
            this.contentElement = this.createContentElement();
            this.virtualizer = new Virtualizer({
                count: options.items.length,
                getScrollElement: () => options.scrollElement,
                estimateSize: index => options.items[index]?.estimatedSize ?? 0,
                scrollToFn: elementScroll,
                observeElementRect,
                observeElementOffset,
                overscan: options.overscan,
                onChange: (_, sync) => {
                    this.isScrolling = sync;
                    this.scheduleSync();
                    this.scheduleIdleMeasurement(sync ? 120 : 0);
                }
            });
            this.virtualizer.shouldAdjustScrollPositionOnItemSizeChange = () => false;
        }
        mount() {
            if (!this.contentElement.isConnected) {
                this.options.hostElement.replaceChildren(this.contentElement);
            }
            this.cleanup = this.virtualizer._didMount();
            this.virtualizer._willUpdate();
            this.sync();
            this.scheduleIdleMeasurement(0);
        }
        destroy() {
            if (this.frameId) {
                cancelAnimationFrame(this.frameId);
                this.frameId = 0;
            }
            if (this.pendingMeasureFrameId) {
                cancelAnimationFrame(this.pendingMeasureFrameId);
                this.pendingMeasureFrameId = 0;
            }
            if (this.idleMeasureTimeoutId) {
                clearTimeout(this.idleMeasureTimeoutId);
                this.idleMeasureTimeoutId = 0;
            }
            if (this.idleMeasureFrameId) {
                cancelAnimationFrame(this.idleMeasureFrameId);
                this.idleMeasureFrameId = 0;
            }
            this.cleanup?.();
            this.cleanup = null;
            this.options.hostElement.replaceChildren();
            this.contentElement.replaceChildren();
            this.elementCache.clear();
            this.pendingMeasureIndices.clear();
            this.lastWindowSignature = null;
        }
        get hostElement() {
            return this.options.hostElement;
        }
        createContentElement() {
            const element = this.options.document.createElement("div");
            element.style.position = "relative";
            element.style.width = "100%";
            element.style.minWidth = "0";
            element.style.boxSizing = "border-box";
            return element;
        }
        scheduleSync() {
            if (this.frameId) {
                return;
            }
            this.frameId = requestAnimationFrame(() => {
                this.frameId = 0;
                this.sync();
            });
        }
        scheduleIdleMeasurement(delay) {
            if (this.idleMeasureTimeoutId) {
                clearTimeout(this.idleMeasureTimeoutId);
                this.idleMeasureTimeoutId = 0;
            }
            if (this.idleMeasureFrameId) {
                cancelAnimationFrame(this.idleMeasureFrameId);
                this.idleMeasureFrameId = 0;
            }
            this.idleMeasureTimeoutId = window.setTimeout(() => {
                this.idleMeasureTimeoutId = 0;
                this.idleMeasureFrameId = requestAnimationFrame(() => {
                    this.idleMeasureFrameId = 0;
                    if (this.isScrolling) {
                        this.scheduleIdleMeasurement(120);
                        return;
                    }
                    this.measureMountedItems();
                });
            }, delay);
        }
        schedulePendingMeasurement(indices) {
            for (const index of indices) {
                this.pendingMeasureIndices.add(index);
            }
            if (!this.pendingMeasureIndices.size || this.pendingMeasureFrameId) {
                return;
            }
            this.pendingMeasureFrameId = requestAnimationFrame(() => {
                this.pendingMeasureFrameId = 0;
                const indices = Array.from(this.pendingMeasureIndices);
                this.pendingMeasureIndices.clear();
                this.measureIndices(indices);
            });
        }
        sync() {
            const virtualItems = this.virtualizer.getVirtualItems();
            const totalSize = this.virtualizer.getTotalSize();
            const nextIndices = new Set();
            const addedIndices = [];
            const removedIndices = [];
            this.contentElement.style.height = `${Math.max(0, totalSize)}px`;
            for (const item of virtualItems) {
                const existing = this.elementCache.get(item.index);
                const element = existing ?? this.options.renderItem(item.index);
                if (!existing) {
                    this.prepareItemElement(element);
                    this.elementCache.set(item.index, element);
                    this.contentElement.appendChild(element);
                    addedIndices.push(item.index);
                }
                else if (!element.isConnected || element.parentElement !== this.contentElement) {
                    this.contentElement.appendChild(element);
                }
                this.positionItemElement(element, item.start);
                nextIndices.add(item.index);
            }
            for (const [index, element] of this.elementCache.entries()) {
                if (nextIndices.has(index)) {
                    continue;
                }
                element.remove();
                this.elementCache.delete(index);
                this.pendingMeasureIndices.delete(index);
                removedIndices.push(index);
            }
            if (addedIndices.length) {
                this.schedulePendingMeasurement(addedIndices);
            }
            this.emitWindowChange(virtualItems, addedIndices, removedIndices);
            this.options.onRendered?.();
        }
        prepareItemElement(element) {
            element.style.position = "absolute";
            element.style.left = "0";
            element.style.right = this.options.centerItems ? "0" : "";
            element.style.marginLeft = this.options.centerItems ? "auto" : "";
            element.style.marginRight = this.options.centerItems ? "auto" : "";
            element.style.marginBottom = "0";
            element.style.boxSizing = "border-box";
        }
        positionItemElement(element, start) {
            element.style.top = `${Math.max(0, Math.round(start))}px`;
        }
        measureMountedItems() {
            this.measureIndices(this.elementCache.keys());
        }
        measureIndices(indices) {
            for (const index of indices) {
                const element = this.elementCache.get(index);
                if (!element?.isConnected) {
                    continue;
                }
                const size = Math.ceil(element.getBoundingClientRect().height) + (this.options.itemGap ?? 0);
                if (size > 0) {
                    this.virtualizer.resizeItem(index, size);
                }
            }
        }
        emitWindowChange(virtualItems, addedIndices, removedIndices) {
            const indices = virtualItems.map(item => item.index);
            const signature = indices.join(",");
            if (signature === this.lastWindowSignature) {
                return;
            }
            this.lastWindowSignature = signature;
            if (!indices.length) {
                return;
            }
            this.options.onWindowChange?.({
                startIndex: indices[0],
                endIndex: indices[indices.length - 1],
                indices,
                addedIndices,
                removedIndices,
                items: indices.map(index => ({
                    index,
                    element: this.elementCache.get(index)
                })),
                isScrolling: this.isScrolling
            });
        }
        getMountedItems() {
            return Array.from(this.elementCache.entries()).map(([index, element]) => ({ index, element }));
        }
        findMountedItem(index) {
            return this.elementCache.get(index) ?? null;
        }
        scrollToIndex(index, options = {}) {
            this.virtualizer.scrollToIndex(index, {
                align: mapBlockToAlign(options.block),
                behavior: options.behavior
            });
        }
    }
    function mapBlockToAlign(block) {
        switch (block) {
            case "center":
                return "center";
            case "end":
                return "end";
            case "nearest":
                return "auto";
            case "start":
            default:
                return "start";
        }
    }

    class DocumentPager {
        constructor(options, styleMap = {}) {
            this.options = options;
            this.styleMap = styleMap;
        }
        static createStyleMap(styles, options) {
            const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);
            const className = options.className ?? "docx";
            for (const style of styles.filter(x => x.basedOn)) {
                const baseStyle = stylesMap[style.basedOn];
                if (baseStyle) {
                    style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
                    style.runProps = mergeDeep(style.runProps, baseStyle.runProps);
                    for (const baseValues of baseStyle.styles) {
                        const styleValues = style.styles.find(x => x.target == baseValues.target);
                        if (styleValues) {
                            copyStyleProperties(baseValues.values, styleValues.values);
                        }
                        else {
                            style.styles.push({ ...baseValues, values: { ...baseValues.values } });
                        }
                    }
                }
                else if (options.debug) {
                    console.warn(`Can't find base style ${style.basedOn}`);
                }
            }
            for (const style of styles) {
                style.cssName = processStyleName(className, style.id);
            }
            return stylesMap;
        }
        buildPages(document) {
            const result = [];
            const allEndnoteIds = [];
            const sections = this.splitBySection(document.children, document.props);
            const pages = this.groupByPageBreaks(sections);
            let prevProps = null;
            for (let i = 0, l = pages.length; i < l; i++) {
                const pageSections = pages[i];
                const pageProps = pageSections[0].sectProps;
                let footerProps = pageProps;
                const initialEndnoteIds = allEndnoteIds.slice();
                for (const sect of pageSections) {
                    this.collectEndnoteIds(sect.elements, allEndnoteIds);
                    footerProps = sect.sectProps;
                }
                result.push({
                    pageIndex: i,
                    key: i,
                    sections: pageSections,
                    pageProps,
                    footerProps,
                    headerRefs: pageProps?.headerRefs,
                    footerRefs: footerProps?.footerRefs,
                    firstOfSection: prevProps != pageProps,
                    initialEndnoteIds,
                    estimatedHeight: this.estimatePageHeight(pageProps),
                    isLastPage: i == l - 1
                });
                prevProps = footerProps;
            }
            return result;
        }
        findStyle(styleName) {
            return styleName && this.styleMap?.[styleName];
        }
        isPageBreakElement(elem) {
            if (elem.type != DomType.Break)
                return false;
            if (elem.break == "lastRenderedPageBreak")
                return !this.options.ignoreLastRenderedPageBreak;
            return elem.break == "page";
        }
        isPageBreakSection(prev, next) {
            if (!prev || !next)
                return false;
            return prev.pageSize?.orientation != next.pageSize?.orientation
                || prev.pageSize?.width != next.pageSize?.width
                || prev.pageSize?.height != next.pageSize?.height;
        }
        splitBySection(elements, defaultProps) {
            let current = { sectProps: null, elements: [], pageBreak: false };
            const result = [current];
            for (const elem of elements) {
                if (elem.type == DomType.Paragraph) {
                    const style = this.findStyle(elem.styleName);
                    if (style?.paragraphProps?.pageBreakBefore && current.elements.length > 0) {
                        current.pageBreak = true;
                        current = { sectProps: null, elements: [], pageBreak: false };
                        result.push(current);
                    }
                }
                current.elements.push(elem);
                if (elem.type != DomType.Paragraph)
                    continue;
                const paragraph = elem;
                const sectProps = paragraph.sectionProps;
                let paragraphBreakIndex = -1;
                let runBreakIndex = -1;
                if (this.options.breakPages && paragraph.children) {
                    paragraphBreakIndex = paragraph.children.findIndex(run => {
                        runBreakIndex = run.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
                        return runBreakIndex != -1;
                    });
                }
                if (sectProps || paragraphBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current.pageBreak = paragraphBreakIndex != -1;
                    current = { sectProps: null, elements: [], pageBreak: false };
                    result.push(current);
                }
                if (paragraphBreakIndex == -1)
                    continue;
                const breakRun = paragraph.children[paragraphBreakIndex];
                const splitRun = runBreakIndex < breakRun.children.length - 1;
                if (paragraphBreakIndex < paragraph.children.length - 1 || splitRun) {
                    const children = elem.children;
                    const newParagraph = { ...elem, children: children.slice(paragraphBreakIndex) };
                    elem.children = children.slice(0, paragraphBreakIndex);
                    current.elements.push(newParagraph);
                    if (splitRun) {
                        const runChildren = breakRun.children;
                        const newRun = { ...breakRun, children: runChildren.slice(0, runBreakIndex) };
                        elem.children.push(newRun);
                        breakRun.children = runChildren.slice(runBreakIndex);
                    }
                }
            }
            let currentSectProps = null;
            for (let i = result.length - 1; i >= 0; i--) {
                if (result[i].sectProps == null) {
                    result[i].sectProps = currentSectProps ?? defaultProps;
                }
                else {
                    currentSectProps = result[i].sectProps;
                }
            }
            return this.coalesceEmptySections(this.resolveSectionProps(result));
        }
        resolveSectionProps(sections) {
            let previous = null;
            for (const section of sections) {
                if (previous) {
                    section.sectProps = this.mergeSectionProps(previous, section.sectProps);
                }
                previous = section.sectProps;
            }
            return sections;
        }
        mergeSectionProps(base, override) {
            if (!base)
                return override;
            if (!override)
                return base;
            if (base === override)
                return base;
            return {
                ...base,
                ...override,
                type: override.type ?? base.type,
                pageSize: override.pageSize ?? base.pageSize,
                pageMargins: override.pageMargins ?? base.pageMargins,
                pageBorders: override.pageBorders ?? base.pageBorders,
                pageNumber: override.pageNumber ?? base.pageNumber,
                columns: override.columns,
                footerRefs: override.footerRefs ?? base.footerRefs,
                headerRefs: override.headerRefs ?? base.headerRefs,
                titlePage: override.titlePage ?? base.titlePage,
            };
        }
        coalesceEmptySections(sections) {
            const result = [];
            for (let i = 0; i < sections.length; i++) {
                const section = sections[i];
                const next = sections[i + 1];
                if (next && !section.pageBreak && !this.sectionHasVisibleContent(section) && !this.sectionForcesStandalonePage(section)) {
                    next.elements = [...section.elements, ...next.elements];
                    next.sectProps = this.mergeSectionProps(section.sectProps, next.sectProps);
                    next.pageBreak = section.pageBreak || next.pageBreak;
                    continue;
                }
                result.push(section);
            }
            return result;
        }
        sectionForcesStandalonePage(section) {
            switch (section.sectProps?.type) {
                case SectionType.EvenPage:
                case SectionType.OddPage:
                    return true;
                default:
                    return false;
            }
        }
        sectionHasVisibleContent(section) {
            return section.elements?.some(element => this.elementHasVisibleContent(element)) ?? false;
        }
        elementHasVisibleContent(element) {
            if (!element)
                return false;
            switch (element.type) {
                case DomType.Text:
                case DomType.DeletedText:
                    return !!element.text?.trim();
                case DomType.Image:
                case DomType.Drawing:
                case DomType.Table:
                case DomType.Symbol:
                case DomType.Tab:
                case DomType.NoBreakHyphen:
                case DomType.FootnoteReference:
                case DomType.EndnoteReference:
                case DomType.CommentReference:
                case DomType.AltChunk:
                    return true;
            }
            return element.children?.some(child => this.elementHasVisibleContent(child)) ?? false;
        }
        groupByPageBreaks(sections) {
            let current = [];
            let prev;
            const result = [current];
            for (const section of sections) {
                current.push(section);
                if (this.options.ignoreLastRenderedPageBreak || section.pageBreak || this.isPageBreakSection(prev, section.sectProps)) {
                    result.push(current = []);
                }
                prev = section.sectProps;
            }
            return result.filter(x => x.length > 0);
        }
        collectEndnoteIds(elements, output) {
            if (!elements)
                return;
            for (const element of elements) {
                if (element.type == DomType.EndnoteReference) {
                    output.push(element.id);
                }
                if (element.children?.length) {
                    this.collectEndnoteIds(element.children, output);
                }
            }
        }
        estimatePageHeight(props) {
            const defaultPageHeight = 1122;
            const pageHeight = parseSizeToPixels$1(props?.pageSize?.height) ?? defaultPageHeight;
            return pageHeight + (this.options.inWrapper ? 30 : 0);
        }
    }
    function processStyleName(className, styleName) {
        return styleName ? `${className}_${escapeClassName(styleName)}` : className;
    }
    function copyStyleProperties(input, output, attrs = null) {
        if (!input)
            return output;
        if (output == null) {
            output = {};
        }
        if (attrs == null) {
            attrs = Object.getOwnPropertyNames(input);
        }
        for (const key of attrs) {
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key)) {
                output[key] = input[key];
            }
        }
        return output;
    }
    function parseSizeToPixels$1(value) {
        if (!value)
            return null;
        if (value.endsWith("px"))
            return parseFloat(value);
        if (value.endsWith("pt"))
            return parseFloat(value) * (96 / 72);
        if (value.endsWith("cm"))
            return parseFloat(value) * (96 / 2.54);
        if (value.endsWith("mm"))
            return parseFloat(value) * (96 / 25.4);
        if (value.endsWith("in"))
            return parseFloat(value) * 96;
        return parseFloat(value);
    }

    const ns = {
        svg: "http://www.w3.org/2000/svg",
        mathML: "http://www.w3.org/1998/Math/MathML"
    };
    class HtmlRenderer {
        constructor(htmlDocument) {
            this.htmlDocument = htmlDocument;
            this.className = "docx";
            this.styleMap = {};
            this.currentPart = null;
            this.currentSectionProps = null;
            this.tableVerticalMerges = [];
            this.currentVerticalMerge = null;
            this.tableCellPositions = [];
            this.currentCellPosition = null;
            this.footnoteMap = {};
            this.endnoteMap = {};
            this.currentEndnoteIds = [];
            this.usedHederFooterParts = [];
            this.currentTabs = [];
            this.commentMap = {};
            this.tasks = [];
            this.postRenderTasks = [];
            this.pageVirtualizer = null;
            this.lastMountedWindowSignature = null;
        }
        async render(document, bodyContainer, styleContainer = null, options) {
            this.document = document;
            this.options = options;
            this.className = options.className;
            this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
            this.styleMap = null;
            this.tasks = [];
            this.postRenderTasks = [];
            this.currentTabs = [];
            this.currentEndnoteIds = [];
            this.commentMap = {};
            this.footnoteMap = {};
            this.endnoteMap = {};
            this.usedHederFooterParts = [];
            this.currentSectionProps = null;
            this.pageVirtualizer?.destroy();
            this.pageVirtualizer = null;
            if (this.options.renderComments && globalThis.Highlight) {
                this.commentHighlight = new Highlight();
            }
            styleContainer = styleContainer || bodyContainer;
            removeAllElements(styleContainer);
            removeAllElements(bodyContainer);
            styleContainer.appendChild(this.createComment("docxjs library predefined styles"));
            styleContainer.appendChild(this.renderDefaultStyle());
            if (document.themePart) {
                styleContainer.appendChild(this.createComment("docxjs document theme values"));
                this.renderTheme(document.themePart, styleContainer);
            }
            if (document.stylesPart != null) {
                this.styleMap = this.processStyles(document.stylesPart.styles);
                styleContainer.appendChild(this.createComment("docxjs document styles"));
                styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
            }
            if (document.numberingPart) {
                this.prodessNumberings(document.numberingPart.domNumberings);
                styleContainer.appendChild(this.createComment("docxjs document numbering styles"));
                styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
            }
            if (document.footnotesPart) {
                this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
            }
            if (document.endnotesPart) {
                this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
            }
            if (document.settingsPart) {
                this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
            }
            if (!options.ignoreFonts && document.fontTablePart)
                this.renderFontTable(document.fontTablePart, styleContainer);
            const pages = document.pages ?? new DocumentPager(this.options, this.styleMap).buildPages(document.documentPart.body);
            const scrollElement = this.resolveVirtualScrollElement(bodyContainer, pages);
            let bodyHost = bodyContainer;
            if (scrollElement) {
                bodyHost = this.options.inWrapper ? this.renderWrapper([]) : this.createElement("div");
                bodyHost.dataset.docxPageCount = `${pages.length}`;
                bodyHost.dataset.docxVirtualized = "true";
                if (this.options.inWrapper || bodyHost !== bodyContainer) {
                    bodyContainer.appendChild(bodyHost);
                }
                this.pageVirtualizer = new VirtualizedRenderer({
                    document: this.htmlDocument,
                    hostElement: bodyHost,
                    scrollElement,
                    items: pages.map(page => ({
                        key: page.key,
                        estimatedSize: page.estimatedHeight
                    })),
                    overscan: this.options.virtualizePagesOverscan,
                    itemGap: this.options.inWrapper ? 30 : 0,
                    centerItems: this.options.inWrapper,
                    renderItem: index => this.renderPage(document.preparePageForRender(pages[index]), document.documentPart.body),
                    onRendered: () => {
                        this.flushPostRenderTasks();
                        this.refreshTabStops();
                    },
                    onWindowChange: payload => {
                        this.emitMountedPageWindowChange(pages, payload);
                    }
                });
                this.pageVirtualizer.mount();
            }
            else {
                var sectionElements = pages.map(page => this.renderPage(document.preparePageForRender(page), document.documentPart.body));
                bodyHost.dataset.docxPageCount = `${pages.length}`;
                if (this.options.inWrapper) {
                    bodyContainer.appendChild(this.renderWrapper(sectionElements));
                }
                else {
                    appendChildren(bodyContainer, sectionElements);
                }
                this.emitMountedPageWindowChange(pages, {
                    startIndex: pages[0]?.pageIndex ?? 0,
                    endIndex: pages[pages.length - 1]?.pageIndex ?? 0,
                    indices: pages.map((_, index) => index),
                    addedIndices: pages.map((_, index) => index),
                    removedIndices: [],
                    items: sectionElements.map((element, index) => ({
                        index,
                        element
                    })),
                    isScrolling: false
                });
            }
            if (this.commentHighlight && options.renderComments) {
                CSS.highlights.set(`${this.className}-comments`, this.commentHighlight);
            }
            this.flushPostRenderTasks();
            await Promise.allSettled(this.tasks);
            this.refreshTabStops();
            return this.createRenderHandle(document, bodyContainer, styleContainer, bodyHost, pages);
        }
        renderTheme(themePart, styleContainer) {
            const variables = {};
            const fontScheme = themePart.theme?.fontScheme;
            if (fontScheme) {
                if (fontScheme.majorFont) {
                    variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
                }
                if (fontScheme.minorFont) {
                    variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
                }
            }
            const colorScheme = themePart.theme?.colorScheme;
            if (colorScheme) {
                for (let [k, v] of Object.entries(colorScheme.colors)) {
                    variables[`--docx-${k}-color`] = `#${v}`;
                }
            }
            const cssText = this.styleToString(`.${this.className}`, variables);
            styleContainer.appendChild(this.createStyleElement(cssText));
        }
        renderFontTable(fontsPart, styleContainer) {
            for (let f of fontsPart.fonts) {
                for (let ref of f.embedFontRefs) {
                    this.tasks.push(this.document.loadFont(ref.id, ref.key).then(fontData => {
                        const cssValues = {
                            'font-family': encloseFontFamily(f.name),
                            'src': `url(${fontData})`
                        };
                        if (ref.type == "bold" || ref.type == "boldItalic") {
                            cssValues['font-weight'] = 'bold';
                        }
                        if (ref.type == "italic" || ref.type == "boldItalic") {
                            cssValues['font-style'] = 'italic';
                        }
                        const cssText = this.styleToString("@font-face", cssValues);
                        styleContainer.appendChild(this.createComment(`docxjs ${f.name} font`));
                        styleContainer.appendChild(this.createStyleElement(cssText));
                    }));
                }
            }
        }
        processStyleName(className) {
            return className ? `${this.className}_${escapeClassName(className)}` : this.className;
        }
        processStyles(styles) {
            return DocumentPager.createStyleMap(styles, this.options);
        }
        prodessNumberings(numberings) {
            for (let num of numberings.filter(n => n.pStyleName)) {
                const style = this.findStyle(num.pStyleName);
                if (style?.paragraphProps?.numbering) {
                    style.paragraphProps.numbering.level = num.level;
                }
            }
        }
        processElement(element) {
            if (element.children) {
                for (var e of element.children) {
                    e.parent = element;
                    if (e.type == DomType.Table) {
                        this.processTable(e);
                    }
                    else {
                        this.processElement(e);
                    }
                }
            }
        }
        processTable(table) {
            for (var r of table.children) {
                for (var c of r.children) {
                    c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
                        "border-left", "border-right", "border-top", "border-bottom",
                        "padding-left", "padding-right", "padding-top", "padding-bottom"
                    ]);
                    this.inheritTableCellPadding(table.cellStyle, c.cssStyle);
                    this.processElement(c);
                }
            }
        }
        inheritTableCellPadding(input, output) {
            if (!input || !output)
                return;
            for (const key of ["padding-left", "padding-right", "padding-top", "padding-bottom"]) {
                if (input[key] != null && shouldInheritPadding(output[key])) {
                    output[key] = input[key];
                }
            }
        }
        copyStyleProperties(input, output, attrs = null) {
            if (!input)
                return output;
            if (output == null)
                output = {};
            if (attrs == null)
                attrs = Object.getOwnPropertyNames(input);
            for (var key of attrs) {
                if (input.hasOwnProperty(key) && (!output.hasOwnProperty(key) || output[key] == null || output[key] === ""))
                    output[key] = input[key];
            }
            return output;
        }
        createPageElement(className, props) {
            var elem = this.createElement("section", { className });
            if (props) {
                if (props.pageMargins) {
                    elem.style.paddingLeft = props.pageMargins.left;
                    elem.style.paddingRight = props.pageMargins.right;
                    elem.style.paddingTop = props.pageMargins.top;
                    elem.style.paddingBottom = props.pageMargins.bottom;
                }
                if (props.pageSize) {
                    if (!this.options.ignoreWidth)
                        elem.style.width = props.pageSize.width;
                    if (!this.options.ignoreHeight)
                        elem.style.minHeight = props.pageSize.height;
                }
            }
            return elem;
        }
        createSectionContent(props) {
            var elem = this.createElement("article");
            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = props.columns.space;
                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
            return elem;
        }
        buildPages(document) {
            return new DocumentPager(this.options, this.styleMap).buildPages(document);
        }
        renderPage(page, document) {
            this.currentFootnoteIds = [];
            this.currentEndnoteIds = page.initialEndnoteIds.slice();
            const pageElement = this.createPageElement(this.className, page.pageProps);
            pageElement.dataset.index = `${page.pageIndex}`;
            this.renderStyleValues(document.cssStyle, pageElement);
            this.options.renderHeaders && this.renderHeaderFooter(page.pageProps.headerRefs, page.pageProps, page.pageIndex, page.firstOfSection, pageElement);
            for (const sect of page.sections) {
                const contentElement = this.createSectionContent(sect.sectProps);
                sect.elements.forEach(element => {
                    if (element.type == DomType.Table) {
                        this.processTable(element);
                    }
                    else {
                        this.processElement(element);
                    }
                });
                if (this.options.mergeAdjacent) {
                    sect.elements.forEach(element => this.ensureOptimizedTree(element));
                }
                this.currentSectionProps = sect.sectProps;
                this.renderSectionElements(sect, contentElement);
                pageElement.appendChild(contentElement);
            }
            if (this.options.renderFootnotes) {
                this.renderNotes(this.currentFootnoteIds, this.footnoteMap, pageElement);
            }
            if (this.options.renderEndnotes && page.isLastPage) {
                this.renderNotes(this.currentEndnoteIds, this.endnoteMap, pageElement);
            }
            this.options.renderFooters && this.renderHeaderFooter(page.footerProps.footerRefs, page.footerProps, page.pageIndex, page.firstOfSection, pageElement);
            return pageElement;
        }
        renderHeaderFooter(refs, props, page, firstOfSection, into) {
            if (!refs)
                return;
            var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
                ?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
                ?? refs.find(x => x.type == "default");
            var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart);
            if (part) {
                const previousPart = this.currentPart;
                const previousSectionProps = this.currentSectionProps;
                this.currentPart = part;
                this.currentSectionProps = props;
                if (!this.usedHederFooterParts.includes(part.path)) {
                    this.processElement(part.rootElement);
                    this.usedHederFooterParts.push(part.path);
                }
                if (this.options.mergeAdjacent) {
                    this.ensureOptimizedTree(part.rootElement);
                }
                const [el] = this.renderElements([part.rootElement], into);
                if (props?.pageMargins) {
                    if (part.rootElement.type === DomType.Header) {
                        el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
                        el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
                    }
                    else if (part.rootElement.type === DomType.Footer) {
                        el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
                        el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
                    }
                }
                this.currentPart = previousPart;
                this.currentSectionProps = previousSectionProps;
            }
        }
        isPageBreakElement(elem) {
            if (elem.type != DomType.Break)
                return false;
            if (elem.break == "lastRenderedPageBreak")
                return !this.options.ignoreLastRenderedPageBreak;
            return elem.break == "page";
        }
        isPageBreakSection(prev, next) {
            if (!prev)
                return false;
            if (!next)
                return false;
            return prev.pageSize?.orientation != next.pageSize?.orientation
                || prev.pageSize?.width != next.pageSize?.width
                || prev.pageSize?.height != next.pageSize?.height;
        }
        splitBySection(elements, defaultProps) {
            var current = { sectProps: null, elements: [], pageBreak: false };
            var result = [current];
            for (let elem of elements) {
                if (elem.type == DomType.Paragraph) {
                    const s = this.findStyle(elem.styleName);
                    if (s?.paragraphProps?.pageBreakBefore && current.elements.length > 0) {
                        current.pageBreak = true;
                        current = { sectProps: null, elements: [], pageBreak: false };
                        result.push(current);
                    }
                }
                current.elements.push(elem);
                if (elem.type == DomType.Paragraph) {
                    const p = elem;
                    var sectProps = p.sectionProps;
                    var pBreakIndex = -1;
                    var rBreakIndex = -1;
                    if (this.options.breakPages && p.children) {
                        pBreakIndex = p.children.findIndex(r => {
                            rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
                            return rBreakIndex != -1;
                        });
                    }
                    if (sectProps || pBreakIndex != -1) {
                        current.sectProps = sectProps;
                        current.pageBreak = pBreakIndex != -1;
                        current = { sectProps: null, elements: [], pageBreak: false };
                        result.push(current);
                    }
                    if (pBreakIndex != -1) {
                        let breakRun = p.children[pBreakIndex];
                        let splitRun = rBreakIndex < breakRun.children.length - 1;
                        if (pBreakIndex < p.children.length - 1 || splitRun) {
                            var children = elem.children;
                            var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
                            elem.children = children.slice(0, pBreakIndex);
                            current.elements.push(newParagraph);
                            if (splitRun) {
                                let runChildren = breakRun.children;
                                let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
                                elem.children.push(newRun);
                                breakRun.children = runChildren.slice(rBreakIndex);
                            }
                        }
                    }
                }
            }
            let currentSectProps = null;
            for (let i = result.length - 1; i >= 0; i--) {
                if (result[i].sectProps == null) {
                    result[i].sectProps = currentSectProps ?? defaultProps;
                }
                else {
                    currentSectProps = result[i].sectProps;
                }
            }
            return this.coalesceEmptySections(this.resolveSectionProps(result));
        }
        resolveSectionProps(sections) {
            let previous = null;
            for (const section of sections) {
                if (previous) {
                    section.sectProps = this.mergeSectionProps(previous, section.sectProps);
                }
                previous = section.sectProps;
            }
            return sections;
        }
        mergeSectionProps(base, override) {
            if (!base)
                return override;
            if (!override)
                return base;
            if (base === override)
                return base;
            return {
                ...base,
                ...override,
                type: override.type ?? base.type,
                pageSize: override.pageSize ?? base.pageSize,
                pageMargins: override.pageMargins ?? base.pageMargins,
                pageBorders: override.pageBorders ?? base.pageBorders,
                pageNumber: override.pageNumber ?? base.pageNumber,
                columns: override.columns,
                footerRefs: override.footerRefs ?? base.footerRefs,
                headerRefs: override.headerRefs ?? base.headerRefs,
                titlePage: override.titlePage ?? base.titlePage,
            };
        }
        coalesceEmptySections(sections) {
            const result = [];
            for (let i = 0; i < sections.length; i++) {
                const section = sections[i];
                const next = sections[i + 1];
                if (next && !section.pageBreak && !this.sectionHasVisibleContent(section) && !this.sectionForcesStandalonePage(section)) {
                    next.elements = [...section.elements, ...next.elements];
                    next.sectProps = this.mergeSectionProps(section.sectProps, next.sectProps);
                    next.pageBreak = section.pageBreak || next.pageBreak;
                    continue;
                }
                result.push(section);
            }
            return result;
        }
        sectionForcesStandalonePage(section) {
            switch (section.sectProps?.type) {
                case SectionType.EvenPage:
                case SectionType.OddPage:
                    return true;
                default:
                    return false;
            }
        }
        sectionHasVisibleContent(section) {
            return section.elements?.some(element => this.elementHasVisibleContent(element)) ?? false;
        }
        elementHasVisibleContent(element) {
            if (!element)
                return false;
            switch (element.type) {
                case DomType.Text:
                case DomType.DeletedText:
                    return !!element.text?.trim();
                case DomType.Image:
                case DomType.Drawing:
                case DomType.Table:
                case DomType.Symbol:
                case DomType.Tab:
                case DomType.NoBreakHyphen:
                case DomType.FootnoteReference:
                case DomType.EndnoteReference:
                case DomType.CommentReference:
                case DomType.AltChunk:
                    return true;
            }
            return element.children?.some(child => this.elementHasVisibleContent(child)) ?? false;
        }
        groupByPageBreaks(sections) {
            let current = [];
            let prev;
            const result = [current];
            for (let s of sections) {
                current.push(s);
                if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
                    result.push(current = []);
                prev = s.sectProps;
            }
            return result.filter(x => x.length > 0);
        }
        renderWrapper(children) {
            return this.createElement("div", { className: `${this.className}-wrapper` }, children);
        }
        renderDefaultStyle() {
            var c = this.className;
            var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
            if (this.options.hideWrapperOnPrint) {
                wrapperStyle = `@media not print { ${wrapperStyle} }`;
            }
            var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
`;
            if (this.options.renderComments) {
                styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`;
            }
            return this.createStyleElement(styleText);
        }
        renderNumbering(numberings, styleContainer) {
            var styleText = "";
            var resetCounters = [];
            for (var num of numberings) {
                var selector = `p.${this.numberingClass(num.id, num.level)}`;
                var listStyleType = "none";
                if (num.bullet) {
                    let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();
                    styleText += this.styleToString(`${selector}:before`, {
                        "content": "' '",
                        "display": "inline-block",
                        "background": `var(${valiable})`
                    }, num.bullet.style);
                    this.tasks.push(this.document.loadNumberingImage(num.bullet.src).then(data => {
                        var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
                        styleContainer.appendChild(this.createStyleElement(text));
                    }));
                }
                else if (num.levelText) {
                    let counter = this.numberingCounter(num.id, num.level);
                    const counterReset = counter + " " + (num.start - 1);
                    if (num.level > 0) {
                        styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                            "counter-set": counterReset
                        });
                    }
                    resetCounters.push(counterReset);
                    styleText += this.styleToString(`${selector}:before`, {
                        "content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
                        "counter-increment": counter,
                        ...num.rStyle,
                    });
                }
                else {
                    listStyleType = this.numFormatToCssValue(num.format);
                }
                styleText += this.styleToString(selector, {
                    "display": "list-item",
                    "list-style-position": "inside",
                    "list-style-type": listStyleType,
                    ...num.pStyle
                });
            }
            if (resetCounters.length > 0) {
                styleText += this.styleToString(this.rootSelector, {
                    "counter-reset": resetCounters.join(" ")
                });
            }
            return this.createStyleElement(styleText);
        }
        renderStyles(styles) {
            var styleText = "";
            const stylesMap = this.styleMap;
            const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);
            for (const style of styles) {
                var subStyles = style.styles;
                if (style.linked) {
                    var linkedStyle = style.linked && stylesMap[style.linked];
                    if (linkedStyle)
                        subStyles = subStyles.concat(linkedStyle.styles);
                    else if (this.options.debug)
                        console.warn(`Can't find linked style ${style.linked}`);
                }
                for (const subStyle of subStyles) {
                    var selector = `${style.target ?? ''}.${style.cssName}`;
                    if (style.target != subStyle.target)
                        selector += ` ${subStyle.target}`;
                    if (defautStyles[style.target] == style)
                        selector = `.${this.className} ${style.target}, ` + selector;
                    styleText += this.styleToString(selector, subStyle.values);
                }
            }
            return this.createStyleElement(styleText);
        }
        renderNotes(noteIds, notesMap, into) {
            var notes = noteIds.map(id => notesMap[id]).filter(x => x);
            if (notes.length > 0) {
                notes.forEach(note => {
                    this.processElement(note);
                    if (this.options.mergeAdjacent) {
                        this.ensureOptimizedTree(note);
                    }
                });
                var result = this.createElement("ol", null, this.renderElements(notes));
                into.appendChild(result);
            }
        }
        renderElement(elem) {
            switch (elem.type) {
                case DomType.Paragraph:
                    return this.renderParagraph(elem);
                case DomType.BookmarkStart:
                    return this.renderBookmarkStart(elem);
                case DomType.BookmarkEnd:
                    return null;
                case DomType.Run:
                    return this.renderRun(elem);
                case DomType.Table:
                    return this.renderTable(elem);
                case DomType.Row:
                    return this.renderTableRow(elem);
                case DomType.Cell:
                    return this.renderTableCell(elem);
                case DomType.Hyperlink:
                    return this.renderHyperlink(elem);
                case DomType.SmartTag:
                    return this.renderSmartTag(elem);
                case DomType.Drawing:
                    return this.renderDrawing(elem);
                case DomType.Image:
                    return this.renderImage(elem);
                case DomType.Text:
                    return this.renderText(elem);
                case DomType.Text:
                    return this.renderText(elem);
                case DomType.DeletedText:
                    return this.renderDeletedText(elem);
                case DomType.Tab:
                    return this.renderTab(elem);
                case DomType.Symbol:
                    return this.renderSymbol(elem);
                case DomType.Break:
                    return this.renderBreak(elem);
                case DomType.Footer:
                    return this.renderContainer(elem, "footer");
                case DomType.Header:
                    return this.renderContainer(elem, "header");
                case DomType.Footnote:
                case DomType.Endnote:
                    return this.renderContainer(elem, "li");
                case DomType.FootnoteReference:
                    return this.renderFootnoteReference(elem);
                case DomType.EndnoteReference:
                    return this.renderEndnoteReference(elem);
                case DomType.NoBreakHyphen:
                    return this.createElement("wbr");
                case DomType.VmlPicture:
                    return this.renderVmlPicture(elem);
                case DomType.VmlElement:
                    return this.renderVmlElement(elem);
                case DomType.MmlMath:
                    return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
                case DomType.MmlMathParagraph:
                    return this.renderContainer(elem, "span");
                case DomType.MmlFraction:
                    return this.renderContainerNS(elem, ns.mathML, "mfrac");
                case DomType.MmlBase:
                    return this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
                case DomType.MmlNumerator:
                case DomType.MmlDenominator:
                case DomType.MmlFunction:
                case DomType.MmlLimit:
                case DomType.MmlBox:
                    return this.renderContainerNS(elem, ns.mathML, "mrow");
                case DomType.MmlGroupChar:
                    return this.renderMmlGroupChar(elem);
                case DomType.MmlLimitLower:
                    return this.renderContainerNS(elem, ns.mathML, "munder");
                case DomType.MmlMatrix:
                    return this.renderContainerNS(elem, ns.mathML, "mtable");
                case DomType.MmlMatrixRow:
                    return this.renderContainerNS(elem, ns.mathML, "mtr");
                case DomType.MmlRadical:
                    return this.renderMmlRadical(elem);
                case DomType.MmlSuperscript:
                    return this.renderContainerNS(elem, ns.mathML, "msup");
                case DomType.MmlSubscript:
                    return this.renderContainerNS(elem, ns.mathML, "msub");
                case DomType.MmlDegree:
                case DomType.MmlSuperArgument:
                case DomType.MmlSubArgument:
                    return this.renderContainerNS(elem, ns.mathML, "mn");
                case DomType.MmlFunctionName:
                    return this.renderContainerNS(elem, ns.mathML, "ms");
                case DomType.MmlDelimiter:
                    return this.renderMmlDelimiter(elem);
                case DomType.MmlRun:
                    return this.renderMmlRun(elem);
                case DomType.MmlNary:
                    return this.renderMmlNary(elem);
                case DomType.MmlPreSubSuper:
                    return this.renderMmlPreSubSuper(elem);
                case DomType.MmlBar:
                    return this.renderMmlBar(elem);
                case DomType.MmlEquationArray:
                    return this.renderMllList(elem);
                case DomType.Inserted:
                    return this.renderInserted(elem);
                case DomType.Deleted:
                    return this.renderDeleted(elem);
                case DomType.CommentRangeStart:
                    return this.renderCommentRangeStart(elem);
                case DomType.CommentRangeEnd:
                    return this.renderCommentRangeEnd(elem);
                case DomType.CommentReference:
                    return this.renderCommentReference(elem);
                case DomType.AltChunk:
                    return this.renderAltChunk(elem);
            }
            return null;
        }
        renderElements(elems, into) {
            if (elems == null)
                return null;
            var result = elems.flatMap(e => e ? this.renderElement(e) : null).filter(e => e != null);
            if (into)
                appendChildren(into, result);
            return result;
        }
        renderContainer(elem, tagName, props) {
            return this.createElement(tagName, props, this.renderElements(elem.children));
        }
        renderContainerNS(elem, ns, tagName, props) {
            return this.createElementNS(ns, tagName, props, this.renderElements(elem.children));
        }
        renderParagraph(elem) {
            if (elem.sectionProps && !this.elementHasVisibleContent(elem)) {
                return null;
            }
            var result = this.renderContainer(elem, "p");
            const style = this.findStyle(elem.styleName);
            elem.tabs ?? (elem.tabs = style?.paragraphProps?.tabs);
            this.renderClass(elem, result);
            this.renderStyleValues(elem.cssStyle, result);
            this.renderCommonProperties(result.style, elem);
            const numbering = elem.numbering ?? style?.paragraphProps?.numbering;
            if (numbering) {
                result.classList.add(this.numberingClass(numbering.id, numbering.level));
            }
            this.normalizeRenderedDrawingParagraph(result);
            return result;
        }
        renderRunProperties(style, props) {
            this.renderCommonProperties(style, props);
        }
        renderCommonProperties(style, props) {
            if (props == null)
                return;
            if (props.color) {
                style["color"] = props.color;
            }
            if (props.fontSize) {
                style["font-size"] = props.fontSize;
            }
        }
        renderHyperlink(elem) {
            var result = this.renderContainer(elem, "a");
            this.renderStyleValues(elem.cssStyle, result);
            let href = '';
            if (elem.id) {
                const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
                href = rel?.target ?? href;
            }
            if (elem.anchor) {
                href += `#${elem.anchor}`;
            }
            result.href = href;
            return result;
        }
        renderSmartTag(elem) {
            return this.renderContainer(elem, "span");
        }
        renderCommentRangeStart(commentStart) {
            if (!this.options.renderComments)
                return null;
            const rng = new Range();
            this.commentHighlight?.add(rng);
            const result = this.createComment(`start of comment #${commentStart.id}`);
            this.later(() => rng.setStart(result, 0));
            this.commentMap[commentStart.id] = rng;
            return result;
        }
        renderCommentRangeEnd(commentEnd) {
            if (!this.options.renderComments)
                return null;
            const rng = this.commentMap[commentEnd.id];
            const result = this.createComment(`end of comment #${commentEnd.id}`);
            this.later(() => rng?.setEnd(result, 0));
            return result;
        }
        renderCommentReference(commentRef) {
            if (!this.options.renderComments)
                return null;
            var comment = this.document.commentsPart?.commentMap[commentRef.id];
            if (!comment)
                return null;
            const frg = new DocumentFragment();
            const commentRefEl = this.createElement("span", { className: `${this.className}-comment-ref` }, ['💬']);
            const commentsContainerEl = this.createElement("div", { className: `${this.className}-comment-popover` });
            this.renderCommentContent(comment, commentsContainerEl);
            frg.appendChild(this.createComment(`comment #${comment.id} by ${comment.author} on ${comment.date}`));
            frg.appendChild(commentRefEl);
            frg.appendChild(commentsContainerEl);
            return frg;
        }
        renderAltChunk(elem) {
            if (!this.options.renderAltChunks)
                return null;
            var result = this.createElement("iframe");
            this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
                result.srcdoc = x;
            }));
            return result;
        }
        renderCommentContent(comment, container) {
            container.appendChild(this.createElement('div', { className: `${this.className}-comment-author` }, [comment.author]));
            container.appendChild(this.createElement('div', { className: `${this.className}-comment-date` }, [new Date(comment.date).toLocaleString()]));
            this.renderElements(comment.children, container);
        }
        renderDrawing(elem) {
            var result = this.renderContainer(elem, "div");
            const anchor = elem.props?.drawingAnchor;
            const pageMargins = this.currentSectionProps?.pageMargins;
            const isPageRelativeAnchor = anchor?.wrapType == "wrapNone"
                && pageMargins
                && (anchor?.posX?.relative == "page" || anchor?.posY?.relative == "page");
            result.style.display = "inline-block";
            result.style.textIndent = "0px";
            this.renderStyleValues(elem.cssStyle, result);
            if (isPageRelativeAnchor) {
                result.style.display = "block";
                result.style.position = "absolute";
                if (anchor.posX?.relative == "page" && anchor.posX.offset) {
                    result.style.left = `calc(${anchor.posX.offset} - ${pageMargins.left ?? "0px"})`;
                }
                if (anchor.posY?.relative == "page" && anchor.posY.offset) {
                    result.style.top = `calc(${anchor.posY.offset} - ${pageMargins.top ?? "0px"})`;
                }
            }
            else if (!result.style.position) {
                result.style.position = "relative";
            }
            return result;
        }
        renderImage(elem) {
            let result = this.createElement("img");
            let transform = elem.cssStyle?.transform;
            this.renderStyleValues(elem.cssStyle, result);
            if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
                var [left, top, right, bottom] = elem.srcRect;
                transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
                result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
            }
            if (elem.rotation)
                transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;
            result.style.transform = transform?.trim();
            if (this.document) {
                this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
                    result.src = x;
                    if (this.options.onImageRendered && elem.docPrId) {
                        result.onload = () => {
                            this.options.onImageRendered({
                                docPrId: elem.docPrId,
                                imgEl: result,
                                naturalSize: { width: result.naturalWidth, height: result.naturalHeight }
                            });
                        };
                    }
                }));
            }
            return result;
        }
        renderText(elem) {
            return this.htmlDocument.createTextNode(elem.text);
        }
        renderDeletedText(elem) {
            return this.options.renderChanges ? this.renderText(elem) : null;
        }
        renderBreak(elem) {
            if (elem.break == "column") {
                const result = this.createElement("span");
                result.style.display = "block";
                result.style.height = "0";
                result.style.breakAfter = "column";
                return result;
            }
            if (elem.break == "textWrapping") {
                return this.createElement("br");
            }
            return null;
        }
        renderSectionElements(section, contentElement) {
            if (this.shouldRenderManualColumns(section)) {
                this.renderManualColumns(section, contentElement);
                return;
            }
            this.renderElements(section.elements, contentElement);
        }
        shouldRenderManualColumns(section) {
            const columns = section.sectProps?.columns;
            if (!columns || columns.numberOfColumns <= 1 || columns.equalWidth)
                return false;
            if (!columns.columns?.length || columns.columns.length < columns.numberOfColumns)
                return false;
            return this.hasColumnBreak(section.elements);
        }
        hasColumnBreak(elements) {
            return elements?.some(element => this.elementHasBreak(element, "column")) ?? false;
        }
        elementHasBreak(element, type) {
            if (!element)
                return false;
            if (element.type == DomType.Break) {
                return element.break == type;
            }
            return element.children?.some(child => this.elementHasBreak(child, type)) ?? false;
        }
        renderManualColumns(section, contentElement) {
            const columns = section.sectProps.columns;
            const groups = this.splitSectionByColumnBreaks(section.elements);
            const columnCount = Math.max(columns.numberOfColumns, groups.length);
            contentElement.style.columnCount = "";
            contentElement.style.columnGap = "";
            contentElement.style.columnRule = "";
            contentElement.style.display = "flex";
            contentElement.style.alignItems = "flex-start";
            contentElement.style.gap = columns.space ?? "0";
            for (let i = 0; i < columnCount; i++) {
                const column = this.createElement("div");
                const columnProps = columns.columns?.[i];
                column.style.boxSizing = "border-box";
                column.style.minWidth = "0";
                column.style.flex = "0 0 auto";
                column.style.width = columnProps?.width ?? this.getEqualColumnWidth(columns);
                if (!columnProps?.width) {
                    column.style.flex = "1 1 0";
                }
                if (columns.separator && i > 0) {
                    column.style.borderLeft = "1px solid black";
                    column.style.paddingLeft = columns.space ?? "0";
                }
                this.renderElements(groups[i] ?? [], column);
                contentElement.appendChild(column);
            }
        }
        getEqualColumnWidth(columns) {
            const count = Math.max(columns?.numberOfColumns ?? 1, 1);
            return `calc((100% - (${columns?.space ?? "0"} * ${count - 1})) / ${count})`;
        }
        splitSectionByColumnBreaks(elements) {
            const result = [[]];
            let current = result[0];
            for (const element of elements) {
                const fragments = this.splitElementByColumnBreaks(element);
                for (let i = 0; i < fragments.length; i++) {
                    current.push(...fragments[i]);
                    if (i < fragments.length - 1) {
                        current = [];
                        result.push(current);
                    }
                }
            }
            return result;
        }
        splitElementByColumnBreaks(element) {
            if (element?.type != DomType.Paragraph || !element.children?.length) {
                return [[element]];
            }
            const fragments = this.splitParagraphChildrenByColumnBreaks(element.children);
            if (fragments.length == 1) {
                return [[element]];
            }
            return fragments.map(children => {
                if (!children.length)
                    return [];
                return [{ ...element, children }];
            });
        }
        splitParagraphChildrenByColumnBreaks(children) {
            const result = [[]];
            let current = result[0];
            for (const child of children) {
                const fragments = this.splitChildByColumnBreaks(child);
                for (let i = 0; i < fragments.length; i++) {
                    const fragment = fragments[i];
                    if (fragment) {
                        current.push(fragment);
                    }
                    if (i < fragments.length - 1) {
                        current = [];
                        result.push(current);
                    }
                }
            }
            return result;
        }
        splitChildByColumnBreaks(child) {
            if (!child?.children?.length) {
                return [child];
            }
            const result = [];
            let current = [];
            let hasBreak = false;
            for (const nestedChild of child.children) {
                if (this.isColumnBreakElement(nestedChild)) {
                    result.push(current.length ? { ...child, children: current } : null);
                    current = [];
                    hasBreak = true;
                    continue;
                }
                current.push(nestedChild);
            }
            if (!hasBreak) {
                return [child];
            }
            result.push(current.length ? { ...child, children: current } : null);
            return result;
        }
        isColumnBreakElement(elem) {
            return elem?.type == DomType.Break && elem.break == "column";
        }
        renderInserted(elem) {
            if (this.options.renderChanges)
                return this.renderContainer(elem, "ins");
            return this.renderElements(elem.children);
        }
        renderDeleted(elem) {
            if (this.options.renderChanges)
                return this.renderContainer(elem, "del");
            return null;
        }
        renderSymbol(elem) {
            var span = this.createElement("span");
            span.style.fontFamily = elem.font;
            span.innerHTML = `&#x${elem.char};`;
            return span;
        }
        renderFootnoteReference(elem) {
            var result = this.createElement("sup");
            this.currentFootnoteIds.push(elem.id);
            result.textContent = `${this.currentFootnoteIds.length}`;
            return result;
        }
        renderEndnoteReference(elem) {
            var result = this.createElement("sup");
            this.currentEndnoteIds.push(elem.id);
            result.textContent = `${this.currentEndnoteIds.length}`;
            return result;
        }
        renderTab(elem) {
            var tabSpan = this.createElement("span");
            tabSpan.innerHTML = "&emsp;";
            if (this.options.experimental) {
                tabSpan.className = this.tabStopClass();
                var stops = findParent(elem, DomType.Paragraph)?.tabs;
                this.currentTabs.push({ stops, span: tabSpan });
            }
            return tabSpan;
        }
        renderBookmarkStart(elem) {
            return this.createElement("span", { id: elem.name });
        }
        renderRun(elem) {
            if (elem.fieldRun)
                return null;
            const result = this.createElement("span");
            const style = elem.cssStyle ? { ...elem.cssStyle } : null;
            const drawingOnlyOffset = this.extractDrawingOnlyRunOffset(style, elem);
            if (elem.id)
                result.id = elem.id;
            this.renderClass(elem, result);
            this.renderStyleValues(style, result);
            if (drawingOnlyOffset) {
                result.style.display = "inline-block";
                result.style.lineHeight = "0";
                result.style.position = "relative";
                result.style.top = drawingOnlyOffset;
                result.style.verticalAlign = "top";
            }
            if (elem.verticalAlign) {
                const wrapper = this.createElement(elem.verticalAlign);
                this.renderElements(elem.children, wrapper);
                result.appendChild(wrapper);
            }
            else {
                this.renderElements(elem.children, result);
            }
            return result;
        }
        extractDrawingOnlyRunOffset(style, elem) {
            const verticalAlign = style?.["vertical-align"];
            if (!verticalAlign || !this.isDrawingOnlyRun(elem) || !/^[-\d.]+(?:pt|px)$/.test(verticalAlign)) {
                return null;
            }
            delete style["vertical-align"];
            return verticalAlign;
        }
        isDrawingOnlyRun(elem) {
            if (!elem.children?.length) {
                return false;
            }
            return elem.children.every(child => this.isDrawingOnlyInlineElement(child));
        }
        isDrawingOnlyInlineElement(elem) {
            switch (elem?.type) {
                case DomType.Drawing:
                case DomType.Image:
                    return true;
                case DomType.Run:
                case DomType.Hyperlink:
                case DomType.SmartTag:
                    return elem.children?.every(child => this.isDrawingOnlyInlineElement(child)) ?? false;
                default:
                    return false;
            }
        }
        normalizeRenderedDrawingParagraph(paragraph) {
            if (!paragraph || paragraph.childElementCount != 1 || paragraph.textContent?.trim()) {
                return;
            }
            if (!paragraph.querySelector("img, svg")) {
                return;
            }
            paragraph.style.lineHeight = "0";
            paragraph.style.minHeight = "0";
            paragraph.style.fontSize = "0";
        }
        renderTable(elem) {
            let result = this.createElement("table");
            this.tableCellPositions.push(this.currentCellPosition);
            this.tableVerticalMerges.push(this.currentVerticalMerge);
            this.currentVerticalMerge = {};
            this.currentCellPosition = { col: 0, row: 0 };
            if (elem.columns)
                result.appendChild(this.renderTableColumns(elem.columns));
            this.renderClass(elem, result);
            this.renderElements(elem.children, result);
            this.renderStyleValues(elem.cssStyle, result);
            this.currentVerticalMerge = this.tableVerticalMerges.pop();
            this.currentCellPosition = this.tableCellPositions.pop();
            return result;
        }
        renderTableColumns(columns) {
            let result = this.createElement("colgroup");
            for (let col of columns) {
                let colElem = this.createElement("col");
                if (col.width)
                    colElem.style.width = col.width;
                result.appendChild(colElem);
            }
            return result;
        }
        renderTableRow(elem) {
            let result = this.createElement("tr");
            this.currentCellPosition.col = 0;
            if (elem.gridBefore)
                result.appendChild(this.renderTableCellPlaceholder(elem.gridBefore));
            this.renderClass(elem, result);
            this.renderElements(elem.children, result);
            this.renderStyleValues(elem.cssStyle, result);
            if (elem.gridAfter)
                result.appendChild(this.renderTableCellPlaceholder(elem.gridAfter));
            this.currentCellPosition.row++;
            return result;
        }
        renderTableCellPlaceholder(colSpan) {
            const result = this.createElement("td", { colSpan });
            result.style['border'] = 'none';
            return result;
        }
        renderTableCell(elem) {
            let result = this.renderContainer(elem, "td");
            const key = this.currentCellPosition.col;
            if (elem.verticalMerge) {
                if (elem.verticalMerge == "restart") {
                    this.currentVerticalMerge[key] = result;
                    result.rowSpan = 1;
                }
                else if (this.currentVerticalMerge[key]) {
                    this.currentVerticalMerge[key].rowSpan += 1;
                    result.style.display = "none";
                }
            }
            else {
                this.currentVerticalMerge[key] = null;
            }
            this.renderClass(elem, result);
            this.renderStyleValues(elem.cssStyle, result);
            if (elem.span)
                result.colSpan = elem.span;
            this.currentCellPosition.col += result.colSpan;
            return result;
        }
        renderVmlPicture(elem) {
            return this.renderContainer(elem, "div");
        }
        renderVmlElement(elem) {
            var container = this.createSvgElement("svg");
            container.setAttribute("style", elem.cssStyleText);
            const result = this.renderVmlChildElement(elem);
            if (elem.imageHref?.id) {
                this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
                    .then(x => result.setAttribute("href", x)));
            }
            container.appendChild(result);
            requestAnimationFrame(() => {
                const bb = container.firstElementChild.getBBox();
                container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
                container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
            });
            return container;
        }
        renderVmlChildElement(elem) {
            const result = this.createSvgElement(elem.tagName);
            Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));
            for (let child of elem.children) {
                if (child.type == DomType.VmlElement) {
                    result.appendChild(this.renderVmlChildElement(child));
                }
                else {
                    result.appendChild(...asArray(this.renderElement(child)));
                }
            }
            return result;
        }
        renderMmlRadical(elem) {
            const base = elem.children.find(el => el.type == DomType.MmlBase);
            if (elem.props?.hideDegree) {
                return this.createElementNS(ns.mathML, "msqrt", null, this.renderElements([base]));
            }
            const degree = elem.children.find(el => el.type == DomType.MmlDegree);
            return this.createElementNS(ns.mathML, "mroot", null, this.renderElements([base, degree]));
        }
        renderMmlDelimiter(elem) {
            const children = [];
            children.push(this.createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
            children.push(...this.renderElements(elem.children));
            children.push(this.createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));
            return this.createElementNS(ns.mathML, "mrow", null, children);
        }
        renderMmlNary(elem) {
            const children = [];
            const grouped = keyBy(elem.children, x => x.type);
            const sup = grouped[DomType.MmlSuperArgument];
            const sub = grouped[DomType.MmlSubArgument];
            const supElem = sup ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
            const subElem = sub ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
            const charElem = this.createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);
            if (supElem || subElem) {
                children.push(this.createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
            }
            else if (supElem) {
                children.push(this.createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
            }
            else if (subElem) {
                children.push(this.createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
            }
            else {
                children.push(charElem);
            }
            children.push(...this.renderElements(grouped[DomType.MmlBase].children));
            return this.createElementNS(ns.mathML, "mrow", null, children);
        }
        renderMmlPreSubSuper(elem) {
            const children = [];
            const grouped = keyBy(elem.children, x => x.type);
            const sup = grouped[DomType.MmlSuperArgument];
            const sub = grouped[DomType.MmlSubArgument];
            const supElem = sup ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
            const subElem = sub ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
            const stubElem = this.createElementNS(ns.mathML, "mo", null);
            children.push(this.createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));
            children.push(...this.renderElements(grouped[DomType.MmlBase].children));
            return this.createElementNS(ns.mathML, "mrow", null, children);
        }
        renderMmlGroupChar(elem) {
            const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
            const result = this.renderContainerNS(elem, ns.mathML, tagName);
            if (elem.props.char) {
                result.appendChild(this.createElementNS(ns.mathML, "mo", null, [elem.props.char]));
            }
            return result;
        }
        renderMmlBar(elem) {
            const result = this.renderContainerNS(elem, ns.mathML, "mrow");
            switch (elem.props.position) {
                case "top":
                    result.style.textDecoration = "overline";
                    break;
                case "bottom":
                    result.style.textDecoration = "underline";
                    break;
            }
            return result;
        }
        renderMmlRun(elem) {
            const result = this.createElementNS(ns.mathML, "ms", null, this.renderElements(elem.children));
            this.renderClass(elem, result);
            this.renderStyleValues(elem.cssStyle, result);
            return result;
        }
        renderMllList(elem) {
            const result = this.createElementNS(ns.mathML, "mtable");
            this.renderClass(elem, result);
            this.renderStyleValues(elem.cssStyle, result);
            for (let child of this.renderElements(elem.children)) {
                result.appendChild(this.createElementNS(ns.mathML, "mtr", null, [
                    this.createElementNS(ns.mathML, "mtd", null, [child])
                ]));
            }
            return result;
        }
        renderStyleValues(style, ouput) {
            if (!style)
                return;
            for (let k in style) {
                if (k.startsWith("$")) {
                    ouput.setAttribute(k.slice(1), style[k]);
                }
                else {
                    ouput.style[k] = style[k];
                }
            }
        }
        renderClass(input, ouput) {
            if (input.className)
                ouput.className = input.className;
            if (input.styleName)
                ouput.classList.add(this.processStyleName(input.styleName));
        }
        findStyle(styleName) {
            return styleName && this.styleMap?.[styleName];
        }
        numberingClass(id, lvl) {
            return `${this.className}-num-${id}-${lvl}`;
        }
        tabStopClass() {
            return `${this.className}-tab-stop`;
        }
        styleToString(selectors, values, cssText = null) {
            let result = `${selectors} {\r\n`;
            for (const key in values) {
                if (key.startsWith('$'))
                    continue;
                result += `  ${key}: ${values[key]};\r\n`;
            }
            if (cssText)
                result += cssText;
            return result + "}\r\n";
        }
        numberingCounter(id, lvl) {
            return `${this.className}-num-${id}-${lvl}`;
        }
        levelTextToContent(text, suff, id, numformat) {
            const suffMap = {
                "tab": "\\9",
                "space": "\\a0",
            };
            var result = text.replace(/%\d*/g, s => {
                let lvl = parseInt(s.substring(1), 10) - 1;
                return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
            });
            return `"${result}${suffMap[suff] ?? ""}"`;
        }
        numFormatToCssValue(format) {
            var mapping = {
                none: "none",
                bullet: "disc",
                decimal: "decimal",
                lowerLetter: "lower-alpha",
                upperLetter: "upper-alpha",
                lowerRoman: "lower-roman",
                upperRoman: "upper-roman",
                decimalZero: "decimal-leading-zero",
                aiueo: "katakana",
                aiueoFullWidth: "katakana",
                chineseCounting: "simp-chinese-informal",
                chineseCountingThousand: "simp-chinese-informal",
                chineseLegalSimplified: "simp-chinese-formal",
                chosung: "hangul-consonant",
                ideographDigital: "cjk-ideographic",
                ideographTraditional: "cjk-heavenly-stem",
                ideographLegalTraditional: "trad-chinese-formal",
                ideographZodiac: "cjk-earthly-branch",
                iroha: "katakana-iroha",
                irohaFullWidth: "katakana-iroha",
                japaneseCounting: "japanese-informal",
                japaneseDigitalTenThousand: "cjk-decimal",
                japaneseLegal: "japanese-formal",
                thaiNumbers: "thai",
                koreanCounting: "korean-hangul-formal",
                koreanDigital: "korean-hangul-formal",
                koreanDigital2: "korean-hanja-informal",
                hebrew1: "hebrew",
                hebrew2: "hebrew",
                hindiNumbers: "devanagari",
                ganada: "hangul",
                taiwaneseCounting: "cjk-ideographic",
                taiwaneseCountingThousand: "cjk-ideographic",
                taiwaneseDigital: "cjk-decimal",
            };
            return mapping[format] ?? format;
        }
        refreshTabStops() {
            if (!this.options.experimental)
                return;
            setTimeout(() => {
                const pixelToPoint = computePixelToPoint();
                for (let tab of this.currentTabs) {
                    updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
                }
            }, 500);
        }
        createElementNS(ns, tagName, props, children) {
            var result = ns ? this.htmlDocument.createElementNS(ns, tagName) : this.htmlDocument.createElement(tagName);
            Object.assign(result, props);
            children && appendChildren(result, children);
            return result;
        }
        createElement(tagName, props, children) {
            return this.createElementNS(undefined, tagName, props, children);
        }
        createSvgElement(tagName, props, children) {
            return this.createElementNS(ns.svg, tagName, props, children);
        }
        createStyleElement(cssText) {
            return this.createElement("style", { innerHTML: cssText });
        }
        createComment(text) {
            return this.htmlDocument.createComment(text);
        }
        later(func) {
            this.postRenderTasks.push(func);
        }
        flushPostRenderTasks(fromIndex = 0) {
            if (fromIndex >= this.postRenderTasks.length)
                return;
            const tasks = this.postRenderTasks.splice(fromIndex);
            tasks.forEach(task => task());
        }
        resolveVirtualScrollElement(bodyContainer, pages) {
            if (!this.options.virtualizePages || pages.length < 2 || this.options.renderComments)
                return null;
            return findScrollableElement(bodyContainer, this.htmlDocument);
        }
        createRenderHandle(document, bodyContainer, styleContainer, bodyHost, pages) {
            const pageIndexMap = new Map(pages.map((page, index) => [page.pageIndex, index]));
            return {
                destroy: () => {
                    if (this.commentHighlight && this.options.renderComments) {
                        CSS.highlights.delete(`${this.className}-comments`);
                    }
                    this.pageVirtualizer?.destroy();
                    this.pageVirtualizer = null;
                    removeAllElements(bodyContainer);
                    if (styleContainer && styleContainer !== bodyContainer) {
                        removeAllElements(styleContainer);
                    }
                    void document.dispose();
                },
                getMountedPages: () => {
                    if (this.pageVirtualizer) {
                        return this.pageVirtualizer.getMountedItems().map(item => ({
                            pageIndex: pages[item.index].pageIndex,
                            element: item.element
                        }));
                    }
                    return this.getStaticMountedPages(bodyHost);
                },
                findMountedPage: (pageIndex) => {
                    if (this.pageVirtualizer) {
                        const virtualIndex = pageIndexMap.get(pageIndex);
                        return virtualIndex == null ? null : this.pageVirtualizer.findMountedItem(virtualIndex);
                    }
                    return bodyHost.querySelector(`section.${this.className}[data-index="${pageIndex}"]`);
                },
                scrollToPage: (pageIndex, options = {}) => {
                    const virtualIndex = pageIndexMap.get(pageIndex);
                    if (virtualIndex == null) {
                        return false;
                    }
                    if (this.pageVirtualizer) {
                        this.pageVirtualizer.scrollToIndex(virtualIndex, options);
                        return true;
                    }
                    const page = bodyHost.querySelector(`section.${this.className}[data-index="${pageIndex}"]`);
                    if (!page) {
                        return false;
                    }
                    page.scrollIntoView(options);
                    return true;
                }
            };
        }
        getStaticMountedPages(bodyHost) {
            return Array
                .from(bodyHost.querySelectorAll(`section.${this.className}[data-index]`))
                .map(element => ({
                pageIndex: Number(element.dataset.index),
                element: element
            }));
        }
        emitMountedPageWindowChange(pages, payload) {
            if (!this.options.onMountedPageWindowChange) {
                return;
            }
            const pageIndices = payload.indices.map(index => pages[index]?.pageIndex).filter(index => index != null);
            const signature = pageIndices.join(",");
            if (signature === this.lastMountedWindowSignature) {
                return;
            }
            this.lastMountedWindowSignature = signature;
            this.options.onMountedPageWindowChange({
                startPageIndex: pageIndices[0] ?? 0,
                endPageIndex: pageIndices[pageIndices.length - 1] ?? 0,
                pageIndices,
                addedPageIndices: payload.addedIndices.map(index => pages[index]?.pageIndex).filter(index => index != null),
                removedPageIndices: payload.removedIndices.map(index => pages[index]?.pageIndex).filter(index => index != null),
                pages: payload.items.map(item => ({
                    pageIndex: pages[item.index]?.pageIndex,
                    element: item.element
                })).filter(item => item.pageIndex != null),
                isScrolling: payload.isScrolling
            });
        }
        collectEndnoteIds(elements, output) {
            if (!elements)
                return;
            for (const element of elements) {
                if (element.type == DomType.EndnoteReference) {
                    output.push(element.id);
                }
                if (element.children?.length) {
                    this.collectEndnoteIds(element.children, output);
                }
            }
        }
        estimatePageHeight(props) {
            const defaultPageHeight = 1122;
            const pageHeight = parseSizeToPixels(props?.pageSize?.height) ?? defaultPageHeight;
            return pageHeight + (this.options.inWrapper ? 30 : 0);
        }
        optimizeChildren(children) {
            const result = [];
            for (const child of children) {
                const previous = result[result.length - 1];
                if (this.canMergeRuns(previous, child)) {
                    for (const grandChild of child.children ?? []) {
                        grandChild.parent = previous;
                        previous.children.push(grandChild);
                    }
                    continue;
                }
                if (this.canMergeText(previous, child)) {
                    previous.text += child.text;
                    continue;
                }
                result.push(child);
            }
            return result;
        }
        ensureOptimizedTree(element) {
            if (element.__optimizedRuns)
                return;
            if (element.children?.length) {
                element.children.forEach(child => this.ensureOptimizedTree(child));
                element.children = this.optimizeChildren(element.children);
                element.children.forEach(child => child.parent = element);
            }
            element.__optimizedRuns = true;
        }
        canMergeRuns(left, right) {
            if (!left || !right || left.type != DomType.Run || right.type != DomType.Run)
                return false;
            if (left.fieldRun || right.fieldRun || left.id || right.id)
                return false;
            if (left.verticalAlign != right.verticalAlign || left.className != right.className || left.styleName != right.styleName)
                return false;
            if (!sameStyleMap(left.cssStyle, right.cssStyle))
                return false;
            return this.hasSimpleInlineChildren(left) && this.hasSimpleInlineChildren(right);
        }
        canMergeText(left, right) {
            return left?.type == DomType.Text && right?.type == DomType.Text;
        }
        hasSimpleInlineChildren(run) {
            return (run.children ?? []).every(child => simpleInlineChildTypes.has(child.type));
        }
    }
    function removeAllElements(elem) {
        elem.innerHTML = '';
    }
    function appendChildren(elem, children) {
        const ownerDocument = elem.ownerDocument ?? document;
        children.forEach(c => elem.appendChild(isString(c) ? ownerDocument.createTextNode(c) : c));
    }
    function findParent(elem, type) {
        var parent = elem.parent;
        while (parent != null && parent.type != type)
            parent = parent.parent;
        return parent;
    }
    function findScrollableElement(elem, htmlDocument) {
        const defaultView = htmlDocument.defaultView;
        let current = elem;
        while (current) {
            const style = defaultView?.getComputedStyle(current);
            const overflowY = style?.overflowY ?? '';
            const overflow = style?.overflow ?? '';
            if (/(auto|scroll|overlay)/.test(`${overflowY} ${overflow}`)) {
                return current;
            }
            current = current.parentElement;
        }
        return null;
    }
    function parseSizeToPixels(value) {
        if (!value)
            return null;
        const match = /^(-?\d*\.?\d+)(px|pt|pc|in|cm|mm|q)?$/i.exec(value.trim());
        if (!match)
            return null;
        const amount = parseFloat(match[1]);
        const unit = (match[2] ?? "px").toLowerCase();
        switch (unit) {
            case "pt": return amount * 96 / 72;
            case "pc": return amount * 16;
            case "in": return amount * 96;
            case "cm": return amount * 96 / 2.54;
            case "mm": return amount * 96 / 25.4;
            case "q": return amount * 96 / 101.6;
            default: return amount;
        }
    }
    function shouldInheritPadding(value) {
        if (value == null || value === "")
            return true;
        return /^0(?:\.0+)?(?:px|pt|pc|in|cm|mm|q)?$/i.test(value.trim());
    }
    const simpleInlineChildTypes = new Set([
        DomType.Text,
        DomType.DeletedText,
        DomType.Break,
        DomType.Tab,
        DomType.NoBreakHyphen,
        DomType.Symbol,
        DomType.FootnoteReference,
        DomType.EndnoteReference,
    ]);
    function sameStyleMap(left, right) {
        const leftKeys = Object.keys(left ?? {});
        const rightKeys = Object.keys(right ?? {});
        if (leftKeys.length != rightKeys.length)
            return false;
        return leftKeys.every(key => left[key] == right[key]);
    }

    const defaultOptions = {
        ignoreHeight: false,
        ignoreWidth: false,
        ignoreFonts: false,
        breakPages: true,
        debug: false,
        experimental: false,
        className: "docx",
        inWrapper: true,
        hideWrapperOnPrint: false,
        trimXmlDeclaration: true,
        ignoreLastRenderedPageBreak: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true,
        useBase64URL: false,
        renderChanges: false,
        renderComments: false,
        renderAltChunks: true,
        virtualizePages: false,
        virtualizePagesOverscan: 2,
        useWorkerParser: false,
        mergeAdjacent: false
    };

    // ==ClosureCompiler==
    // @output_file_name default.js
    // @compilation_level SIMPLE_OPTIMIZATIONS
    // ==/ClosureCompiler==
    // module.exports = {
    //     parse: parse,
    //     simplify: simplify,
    //     simplifyLostLess: simplifyLostLess,
    //     filter: filter,
    //     stringify: stringify,
    //     toContentString: toContentString,
    //     getElementById: getElementById,
    //     getElementsByClassName: getElementsByClassName,
    //     transformStream: transformStream,
    // };

    /**
     * @author: Tobias Nickel
     * @created: 06.04.2015
     * I needed a small xmlparser chat can be used in a worker.
     */

    /**
     * @typedef tNode 
     * @property {string} tagName 
     * @property {object} attributes
     * @property {(tNode|string)[]} children 
     **/

    /**
     * @typedef TParseOptions
     * @property {number} [pos]
     * @property {string[]} [noChildNodes]
     * @property {boolean} [setPos]
     * @property {boolean} [keepComments] 
     * @property {boolean} [keepWhitespace]
     * @property {boolean} [simplify]
     * @property {(a: tNode, b: tNode) => boolean} [filter]
     */

    /**
     * parseXML / html into a DOM Object. with no validation and some failur tolerance
     * @param {string} S your XML to parse
     * @param {TParseOptions} [options]  all other options:
     * @return {(tNode | string)[]}
     */
    function parse(S, options) {
        "txml";
        options = options || {};

        var pos = options.pos || 0;
        var keepComments = !!options.keepComments;
        var keepWhitespace = !!options.keepWhitespace;

        var openBracket = "<";
        var openBracketCC = "<".charCodeAt(0);
        var closeBracket = ">";
        var closeBracketCC = ">".charCodeAt(0);
        var minusCC = "-".charCodeAt(0);
        var slashCC = "/".charCodeAt(0);
        var exclamationCC = '!'.charCodeAt(0);
        var singleQuoteCC = "'".charCodeAt(0);
        var doubleQuoteCC = '"'.charCodeAt(0);
        var openCornerBracketCC = '['.charCodeAt(0);
        var closeCornerBracketCC = ']'.charCodeAt(0);


        /**
         * parsing a list of entries
         */
        function parseChildren(tagName) {
            var children = [];
            while (S[pos]) {
                if (S.charCodeAt(pos) == openBracketCC) {
                    if (S.charCodeAt(pos + 1) === slashCC) {
                        var closeStart = pos + 2;
                        pos = S.indexOf(closeBracket, pos);

                        var closeTag = S.substring(closeStart, pos);
                        if (closeTag.indexOf(tagName) == -1) {
                            var parsedText = S.substring(0, pos).split('\n');
                            throw new Error(
                                'Unexpected close tag\nLine: ' + (parsedText.length - 1) +
                                '\nColumn: ' + (parsedText[parsedText.length - 1].length + 1) +
                                '\nChar: ' + S[pos]
                            );
                        }

                        if (pos + 1) pos += 1;

                        return children;
                    } else if (S.charCodeAt(pos + 1) === exclamationCC) {
                        if (S.charCodeAt(pos + 2) == minusCC) {
                            //comment support
                            const startCommentPos = pos;
                            while (pos !== -1 && !(S.charCodeAt(pos) === closeBracketCC && S.charCodeAt(pos - 1) == minusCC && S.charCodeAt(pos - 2) == minusCC && pos != -1)) {
                                pos = S.indexOf(closeBracket, pos + 1);
                            }
                            if (pos === -1) {
                                pos = S.length;
                            }
                            if (keepComments) {
                                children.push(S.substring(startCommentPos, pos + 1));
                            }
                        } else if (
                            S.charCodeAt(pos + 2) === openCornerBracketCC &&
                            S.charCodeAt(pos + 8) === openCornerBracketCC &&
                            S.substr(pos + 3, 5).toLowerCase() === 'cdata'
                        ) {
                            // cdata
                            var cdataEndIndex = S.indexOf(']]>', pos);
                            if (cdataEndIndex == -1) {
                                children.push(S.substr(pos + 9));
                                pos = S.length;
                            } else {
                                children.push(S.substring(pos + 9, cdataEndIndex));
                                pos = cdataEndIndex + 3;
                            }
                            continue;
                        } else {
                            // doctypesupport
                            const startDoctype = pos + 1;
                            pos += 2;
                            var encapsuled = false;
                            while ((S.charCodeAt(pos) !== closeBracketCC || encapsuled === true) && S[pos]) {
                                if (S.charCodeAt(pos) === openCornerBracketCC) {
                                    encapsuled = true;
                                } else if (encapsuled === true && S.charCodeAt(pos) === closeCornerBracketCC) {
                                    encapsuled = false;
                                }
                                pos++;
                            }
                            children.push(S.substring(startDoctype, pos));
                        }
                        pos++;
                        continue;
                    }
                    var node = parseNode();
                    children.push(node);
                    if (node.tagName[0] === '?') {
                        children.push(...node.children);
                        node.children = [];
                    }
                } else {
                    var text = parseText();
                    if (keepWhitespace) {
                        if (text.length > 0) {
                            children.push(text);
                        }
                    } else {
                        var trimmed = text.trim();
                        if (trimmed.length > 0) {
                            children.push(trimmed);
                        }
                    }
                    pos++;
                }
            }
            return children;
        }

        /**
         *    returns the text outside of texts until the first '<'
         */
        function parseText() {
            var start = pos;
            pos = S.indexOf(openBracket, pos) - 1;
            if (pos === -2)
                pos = S.length;
            return S.slice(start, pos + 1);
        }
        /**
         *    returns text until the first nonAlphabetic letter
         */
        var nameSpacer = '\r\n\t>/= ';

        function parseName() {
            var start = pos;
            while (nameSpacer.indexOf(S[pos]) === -1 && S[pos]) {
                pos++;
            }
            return S.slice(start, pos);
        }
        /**
         *    is parsing a node, including tagName, Attributes and its children,
         * to parse children it uses the parseChildren again, that makes the parsing recursive
         */
        var NoChildNodes = options.noChildNodes || ['img', 'br', 'input', 'meta', 'link', 'hr'];

        function parseNode() {
            pos++;
            const tagName = parseName();
            const attributes = {};
            let children = [];

            // parsing attributes
            while (S.charCodeAt(pos) !== closeBracketCC && S[pos]) {
                var c = S.charCodeAt(pos);
                if ((c > 64 && c < 91) || (c > 96 && c < 123)) {
                    //if('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'.indexOf(S[pos])!==-1 ){
                    var name = parseName();
                    // search beginning of the string
                    var code = S.charCodeAt(pos);
                    while (code && code !== singleQuoteCC && code !== doubleQuoteCC && !((code > 64 && code < 91) || (code > 96 && code < 123)) && code !== closeBracketCC) {
                        pos++;
                        code = S.charCodeAt(pos);
                    }
                    if (code === singleQuoteCC || code === doubleQuoteCC) {
                        var value = parseString();
                        if (pos === -1) {
                            return {
                                tagName,
                                attributes,
                                children,
                            };
                        }
                    } else {
                        value = null;
                        pos--;
                    }
                    attributes[name] = value;
                }
                pos++;
            }
            // optional parsing of children
            if (S.charCodeAt(pos - 1) !== slashCC) {
                if (tagName == "script") {
                    var start = pos + 1;
                    pos = S.indexOf('</script>', pos);
                    children = [S.slice(start, pos)];
                    pos += 9;
                } else if (tagName == "style") {
                    var start = pos + 1;
                    pos = S.indexOf('</style>', pos);
                    children = [S.slice(start, pos)];
                    pos += 8;
                } else if (NoChildNodes.indexOf(tagName) === -1) {
                    pos++;
                    children = parseChildren(tagName);
                } else {
                    pos++;
                }
            } else {
                pos++;
            }
            return {
                tagName,
                attributes,
                children,
            };
        }

        /**
         *    is parsing a string, that starts with a char and with the same usually  ' or "
         */

        function parseString() {
            var startChar = S[pos];
            var startpos = pos + 1;
            pos = S.indexOf(startChar, startpos);
            return S.slice(startpos, pos);
        }

        /**
         *
         */
        function findElements() {
            var r = new RegExp('\\s' + options.attrName + '\\s*=[\'"]' + options.attrValue + '[\'"]').exec(S);
            if (r) {
                return r.index;
            } else {
                return -1;
            }
        }

        var out = null;
        if (options.attrValue !== undefined) {
            options.attrName = options.attrName || 'id';
            var out = [];

            while ((pos = findElements()) !== -1) {
                pos = S.lastIndexOf('<', pos);
                if (pos !== -1) {
                    out.push(parseNode());
                }
                S = S.substr(pos);
                pos = 0;
            }
        } else if (options.parseNode) {
            out = parseNode();
        } else {
            out = parseChildren('');
        }

        if (options.filter) {
            out = filter(out, options.filter);
        }

        if (options.simplify) {
            return simplify(Array.isArray(out) ? out : [out]);
        }

        if (options.setPos) {
            out.pos = pos;
        }

        return out;
    }

    /**
     * transform the DomObject to an object that is like the object of PHP`s simple_xmp_load_*() methods.
     * this format helps you to write that is more likely to keep your program working, even if there a small changes in the XML schema.
     * be aware, that it is not possible to reproduce the original xml from a simplified version, because the order of elements is not saved.
     * therefore your program will be more flexible and easier to read.
     *
     * @param {tNode[]} children the childrenList
     */
    function simplify(children) {
        var out = {};
        if (!children.length) {
            return '';
        }

        if (children.length === 1 && typeof children[0] == 'string') {
            return children[0];
        }
        // map each object
        children.forEach(function(child) {
            if (typeof child !== 'object') {
                return;
            }
            if (!out[child.tagName])
                out[child.tagName] = [];
            var kids = simplify(child.children);
            out[child.tagName].push(kids);
            if (Object.keys(child.attributes).length && typeof kids !== 'string') {
                kids._attributes = child.attributes;
            }
        });

        for (var i in out) {
            if (out[i].length == 1) {
                out[i] = out[i][0];
            }
        }

        return out;
    }
    /**
     * behaves the same way as Array.filter, if the filter method return true, the element is in the resultList
     * @params children{Array} the children of a node
     * @param f{function} the filter method
     */
    function filter(children, f, dept = 0, path = '') {
        var out = [];
        children.forEach(function(child, i) {
            if (typeof(child) === 'object' && f(child, i, dept, path)) out.push(child);
            if (child.children) {
                var kids = filter(child.children, f, dept + 1, (path ? path + '.' : '') + i + '.' + child.tagName);
                out = out.concat(kids);
            }
        });
        return out;
    }

    function removeUTF8BOM(data) {
        return data.charCodeAt(0) === 0xFEFF ? data.substring(1) : data;
    }
    class TxmlTextNode {
        constructor(textContent) {
            this.textContent = textContent;
            this.nodeType = 3;
            this.nodeName = "#text";
            this.localName = null;
            this.namespaceURI = null;
            this.childNodes = [];
            this.firstChild = null;
            this.firstElementChild = null;
        }
    }
    class TxmlElementNode {
        constructor(nodeName, attributes, children, namespaceMap) {
            this.nodeName = nodeName;
            this.namespaceMap = namespaceMap;
            this.nodeType = 1;
            this.firstChild = null;
            this.firstElementChild = null;
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
            this.firstElementChild = children.find((child) => child.nodeType == 1) ?? null;
        }
        get textContent() {
            return this.childNodes.map(child => child.textContent ?? "").join("");
        }
        lookupNamespaceURI(prefix) {
            return this.namespaceMap[prefix ?? ""] ?? null;
        }
    }
    class TxmlDocumentNode {
        constructor(childNodes) {
            this.childNodes = childNodes;
            this.firstChild = null;
            this.firstElementChild = null;
            this.firstChild = childNodes[0] ?? null;
            this.firstElementChild = childNodes.find((child) => child.nodeType == 1) ?? null;
        }
    }
    function parseXmlStringWithTxml(xmlString, trimXmlDeclaration = false) {
        if (trimXmlDeclaration)
            xmlString = xmlString.replace(/<[?].*[?]>/, "");
        xmlString = removeUTF8BOM(xmlString);
        const parsed = parse(xmlString, { keepWhitespace: true });
        return new TxmlDocumentNode(adaptChildren(parsed));
    }
    function adaptChildren(children, namespaceMap = {}) {
        return (children ?? []).map(child => adaptNode(child, namespaceMap));
    }
    function adaptNode(node, namespaceMap) {
        if (typeof node === "string") {
            return new TxmlTextNode(node);
        }
        const currentNamespaceMap = { ...namespaceMap, ...extractNamespaceMap(node.attributes) };
        return new TxmlElementNode(node.tagName, node.attributes, adaptChildren(node.children, currentNamespaceMap), currentNamespaceMap);
    }
    function extractNamespaceMap(attributes) {
        const result = {};
        for (const [name, value] of Object.entries(attributes ?? {})) {
            if (name === "xmlns") {
                result[""] = value;
            }
            else if (name.startsWith("xmlns:")) {
                result[name.substring(6)] = value;
            }
        }
        return result;
    }

    const topLevelRels = [
        { type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
        { type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
        { type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
        { type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
    ];
    async function parseToSnapshot(data, options) {
        const normalized = normalizeParseOptions(options);
        const builder = new SnapshotBuilder(normalized);
        return await builder.build(data);
    }
    function collectSnapshotTransferables(snapshot) {
        return (snapshot.files ?? []).map(file => file.buffer);
    }
    async function renderSnapshot(snapshot, bodyContainer, styleContainer, options) {
        validateSnapshot(snapshot);
        const normalized = normalizeRenderOptions(options);
        validateRenderOptions(snapshot, normalized);
        const document = WordDocument.fromSnapshot(snapshot, normalized);
        const renderer = new HtmlRenderer(bodyContainer.ownerDocument ?? window.document);
        return await renderer.render(document, bodyContainer, styleContainer, normalized);
    }
    class SnapshotOpenXmlPackage {
        constructor(files, options) {
            this.files = files;
            this.options = options;
            this.xmlParser = new XmlParser();
            this.decoder = new TextDecoder();
        }
        get(path) {
            const normalized = normalizePath(path);
            return this.files[normalized] ?? this.files[normalized.replace(/\//g, "\\")] ?? null;
        }
        async load(path, type = "string") {
            const file = this.get(path);
            if (!file)
                return null;
            switch (type) {
                case "uint8array":
                    return file.slice();
                case "arraybuffer":
                    return toArrayBuffer(file);
                case "blob":
                    return new Blob([file.slice()]);
                case "string":
                default:
                    return this.decoder.decode(file);
            }
        }
        async loadRelationships(path = null) {
            let relsPath = `_rels/.rels`;
            if (path != null) {
                const [folder, fileName] = splitPath(path);
                relsPath = `${folder}_rels/${fileName}.rels`;
            }
            const text = await this.load(relsPath, "string");
            return text ? parseRelationships(this.parseXmlDocument(text).firstElementChild, this.xmlParser) : null;
        }
        parseXmlDocument(text) {
            return parseXmlStringWithTxml(text, this.options.trimXmlDeclaration);
        }
    }
    class SnapshotBuilder {
        constructor(options) {
            this.options = options;
            this.rels = [];
            this.parts = [];
            this.partsMap = {};
            this.rolePaths = {};
            this.documentPart = null;
            this.stylesPart = null;
            this.parser = new DocumentParser(options);
        }
        async build(data) {
            const files = unzipSync(await inputToUint8Array(data));
            this.package = new SnapshotOpenXmlPackage(normalizeFiles(files), this.options);
            this.rels = await this.package.loadRelationships();
            await Promise.all(topLevelRels.map(rel => {
                const relationship = this.rels.find(x => x.type === rel.type) ?? rel;
                return this.loadRelationshipPart(relationship.target, relationship.type);
            }));
            const pagerOptions = {
                breakPages: this.options.breakPages,
                className: defaultOptions.className,
                debug: this.options.debug,
                ignoreLastRenderedPageBreak: this.options.ignoreLastRenderedPageBreak,
                inWrapper: this.options.inWrapper
            };
            const styleMap = this.stylesPart?.styles
                ? DocumentPager.createStyleMap(cloneSerializable(this.stylesPart.styles), pagerOptions)
                : {};
            const pages = new DocumentPager(pagerOptions, styleMap).buildPages(this.documentPart.body);
            return {
                meta: {
                    version: 1,
                    pageCount: pages.length,
                    parseOptions: this.options
                },
                files: this.collectSnapshotFiles(),
                rels: this.rels,
                parts: this.parts.map(part => this.serializePart(part)).filter(Boolean),
                rolePaths: this.rolePaths,
                pages
            };
        }
        async loadRelationshipPart(path, type) {
            const normalizedPath = normalizePath(path);
            if (this.partsMap[normalizedPath])
                return this.partsMap[normalizedPath];
            if (!this.package.get(normalizedPath))
                return null;
            let part = null;
            let partKind = null;
            switch (type) {
                case RelationshipTypes.OfficeDocument:
                    this.documentPart = part = new DocumentPart(this.package, normalizedPath, this.parser);
                    partKind = "document";
                    this.rolePaths.documentPart = normalizedPath;
                    break;
                case RelationshipTypes.FontTable:
                    part = new FontTablePart(this.package, normalizedPath);
                    partKind = "fontTable";
                    this.rolePaths.fontTablePart = normalizedPath;
                    break;
                case RelationshipTypes.Numbering:
                    part = new NumberingPart(this.package, normalizedPath, this.parser);
                    partKind = "numbering";
                    this.rolePaths.numberingPart = normalizedPath;
                    break;
                case RelationshipTypes.Styles:
                    this.stylesPart = part = new StylesPart(this.package, normalizedPath, this.parser);
                    partKind = "styles";
                    this.rolePaths.stylesPart = normalizedPath;
                    break;
                case RelationshipTypes.Theme:
                    part = new ThemePart(this.package, normalizedPath);
                    partKind = "theme";
                    this.rolePaths.themePart = normalizedPath;
                    break;
                case RelationshipTypes.Footnotes:
                    part = new FootnotesPart(this.package, normalizedPath, this.parser);
                    partKind = "footnotes";
                    this.rolePaths.footnotesPart = normalizedPath;
                    break;
                case RelationshipTypes.Endnotes:
                    part = new EndnotesPart(this.package, normalizedPath, this.parser);
                    partKind = "endnotes";
                    this.rolePaths.endnotesPart = normalizedPath;
                    break;
                case RelationshipTypes.Footer:
                    part = new FooterPart(this.package, normalizedPath, this.parser);
                    partKind = "footer";
                    break;
                case RelationshipTypes.Header:
                    part = new HeaderPart(this.package, normalizedPath, this.parser);
                    partKind = "header";
                    break;
                case RelationshipTypes.CoreProperties:
                    part = new CorePropsPart(this.package, normalizedPath);
                    partKind = "coreProps";
                    this.rolePaths.corePropsPart = normalizedPath;
                    break;
                case RelationshipTypes.ExtendedProperties:
                    part = new ExtendedPropsPart(this.package, normalizedPath);
                    partKind = "extendedProps";
                    this.rolePaths.extendedPropsPart = normalizedPath;
                    break;
                case RelationshipTypes.CustomProperties:
                    part = new CustomPropsPart(this.package, normalizedPath);
                    partKind = "customProps";
                    this.rolePaths.customPropsPart = normalizedPath;
                    break;
                case RelationshipTypes.Settings:
                    part = new SettingsPart(this.package, normalizedPath);
                    partKind = "settings";
                    this.rolePaths.settingsPart = normalizedPath;
                    break;
                case RelationshipTypes.Comments:
                    part = new CommentsPart(this.package, normalizedPath, this.parser);
                    partKind = "comments";
                    this.rolePaths.commentsPart = normalizedPath;
                    break;
                case RelationshipTypes.CommentsExtended:
                    part = new CommentsExtendedPart(this.package, normalizedPath);
                    partKind = "commentsExtended";
                    this.rolePaths.commentsExtendedPart = normalizedPath;
                    break;
            }
            if (!part)
                return null;
            part.__kind = partKind;
            this.partsMap[normalizedPath] = part;
            this.parts.push(part);
            await part.load();
            if (part.rels?.length > 0) {
                const [folder] = splitPath(part.path);
                await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
            }
            return part;
        }
        collectSnapshotFiles() {
            const serializedPaths = new Set(this.parts.map(part => normalizePath(part.path)));
            const resourcePaths = new Set();
            for (const part of this.parts) {
                const [folder] = splitPath(part.path);
                for (const rel of part.rels ?? []) {
                    if (rel.targetMode === "External")
                        continue;
                    const resolvedPath = normalizePath(resolvePath(rel.target, folder));
                    if (!this.package.get(resolvedPath) || serializedPaths.has(resolvedPath))
                        continue;
                    resourcePaths.add(resolvedPath);
                }
            }
            return Array.from(resourcePaths).map(path => ({
                path,
                buffer: toArrayBuffer(this.package.get(path))
            }));
        }
        serializePart(part) {
            const base = {
                kind: part.__kind,
                path: part.path,
                rels: part.rels
            };
            switch (part.__kind) {
                case "document":
                    return {
                        ...base,
                        body: serializeDocumentBody(part.body)
                    };
                case "fontTable":
                    return { ...base, fonts: part.fonts };
                case "numbering":
                    return {
                        ...base,
                        numberings: part.numberings,
                        abstractNumberings: part.abstractNumberings,
                        bulletPictures: part.bulletPictures,
                        domNumberings: part.domNumberings
                    };
                case "styles":
                    return { ...base, styles: part.styles };
                case "theme":
                    return { ...base, theme: part.theme };
                case "footnotes":
                case "endnotes":
                    return { ...base, notes: part.notes };
                case "header":
                case "footer":
                    return { ...base, rootElement: part.rootElement };
                case "coreProps":
                case "extendedProps":
                case "customProps":
                    return { ...base, props: part.props };
                case "settings":
                    return { ...base, settings: part.settings };
                case "comments":
                    return { ...base, comments: part.comments };
                case "commentsExtended":
                    return { ...base, comments: part.comments };
                default:
                    return null;
            }
        }
    }
    function normalizeParseOptions(options) {
        return {
            breakPages: options?.breakPages ?? defaultOptions.breakPages,
            debug: options?.debug ?? defaultOptions.debug,
            ignoreLastRenderedPageBreak: options?.ignoreLastRenderedPageBreak ?? defaultOptions.ignoreLastRenderedPageBreak,
            inWrapper: options?.inWrapper ?? defaultOptions.inWrapper,
            trimXmlDeclaration: options?.trimXmlDeclaration ?? defaultOptions.trimXmlDeclaration,
        };
    }
    function normalizeRenderOptions(options) {
        return {
            ...defaultOptions,
            ...options,
            useWorkerParser: false,
            workerUrl: undefined
        };
    }
    function validateSnapshot(snapshot) {
        if (!snapshot?.meta)
            throw new Error("DOCX: Invalid snapshot payload");
        if (snapshot.meta.version !== 1)
            throw new Error(`DOCX: Unsupported snapshot version ${snapshot.meta.version}`);
    }
    function validateRenderOptions(snapshot, options) {
        const parseOptions = snapshot.meta.parseOptions;
        if (options.breakPages !== parseOptions.breakPages) {
            throw new Error("DOCX: renderSnapshot() received breakPages that does not match snapshot parse options");
        }
        if (options.ignoreLastRenderedPageBreak !== parseOptions.ignoreLastRenderedPageBreak) {
            throw new Error("DOCX: renderSnapshot() received ignoreLastRenderedPageBreak that does not match snapshot parse options");
        }
    }
    function serializeDocumentBody(body) {
        if (!body)
            return null;
        return {
            ...body,
            children: []
        };
    }
    function normalizePath(path) {
        return path.startsWith("/") ? path.substring(1) : path;
    }
    function normalizeFiles(files) {
        const result = {};
        for (const [path, file] of Object.entries(files ?? {})) {
            result[normalizePath(path)] = file;
        }
        return result;
    }
    function toArrayBuffer(data) {
        return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
    }
    async function inputToUint8Array(data) {
        if (data instanceof Uint8Array)
            return data.slice();
        if (data instanceof ArrayBuffer)
            return new Uint8Array(data);
        if (ArrayBuffer.isView(data))
            return new Uint8Array(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength));
        if (typeof data?.arrayBuffer === "function")
            return new Uint8Array(await data.arrayBuffer());
        throw new Error("Unsupported input type for parseToSnapshot");
    }
    function cloneSerializable(value) {
        if (typeof structuredClone === "function") {
            return structuredClone(value);
        }
        if (Array.isArray(value)) {
            return value.map(item => cloneSerializable(item));
        }
        if (!value || typeof value !== "object") {
            return value;
        }
        const result = {};
        for (const [key, entry] of Object.entries(value)) {
            result[key] = cloneSerializable(entry);
        }
        return result;
    }

    function parseAsync(data, userOptions) {
        const ops = { ...defaultOptions, ...userOptions };
        return WordDocument.load(data, new DocumentParser(ops), ops);
    }
    async function renderDocument(document, bodyContainer, styleContainer, userOptions) {
        const ops = { ...defaultOptions, ...userOptions };
        const renderer = new HtmlRenderer(bodyContainer.ownerDocument ?? window.document);
        return await renderer.render(document, bodyContainer, styleContainer, ops);
    }
    async function renderAsync(data, bodyContainer, styleContainer, userOptions) {
        const doc = await parseAsync(data, userOptions);
        await renderDocument(doc, bodyContainer, styleContainer, userOptions);
        return doc;
    }

    exports.collectSnapshotTransferables = collectSnapshotTransferables;
    exports.defaultOptions = defaultOptions;
    exports.parseAsync = parseAsync;
    exports.parseToSnapshot = parseToSnapshot;
    exports.renderAsync = renderAsync;
    exports.renderDocument = renderDocument;
    exports.renderSnapshot = renderSnapshot;

}));
//# sourceMappingURL=docx-preview.js.map
