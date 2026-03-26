/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */
(function () {
    'use strict';

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
    var _b = freb(fdeb, 0), fd = _b.b;
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
    var flrm = /*#__PURE__*/ hMap(flt, 9, 1);
    // fixed distance map
    var fdrm = /*#__PURE__*/ hMap(fdt, 5, 1);
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
    // empty
    var et = /*#__PURE__*/ new u8(0);
    // read 2 bytes
    var b2 = function (d, b) { return d[b] | (d[b + 1] << 8); };
    // read 4 bytes
    var b4 = function (d, b) { return (d[b] | (d[b + 1] << 8) | (d[b + 2] << 16) | (d[b + 3] << 24)) >>> 0; };
    var b8 = function (d, b) { return b4(d, b) + (b4(d, b + 4) * 4294967296); };
    /**
     * Expands DEFLATE data with no wrapper
     * @param data The data to decompress
     * @param opts The decompression options
     * @returns The decompressed version of the data
     */
    function inflateSync(data, opts) {
        return inflt(data, { i: 2 }, opts && opts.out, opts && opts.dictionary);
    }
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
    function clamp(val, min, max) {
        return min > val ? min : (max < val ? max : val);
    }

    const ns = {
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
        if (elem.namespaceURI != ns.wordml)
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

    class DocumentPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.body = this._documentParser.parseDocumentFile(root);
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
        if (elem.namespaceURI != ns.wordml)
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

    class StylesPart extends Part {
        constructor(pkg, path, parser) {
            super(pkg, path);
            this._documentParser = parser;
        }
        parseXml(root) {
            this.styles = this._documentParser.parseStylesFile(root);
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

    const topLevelRels = [
        { type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
        { type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
        { type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
        { type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
    ];
    class WorkerOpenXmlPackage {
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
    class WorkerDocumentLoader {
        constructor(options) {
            this.options = options;
            this.rels = [];
            this.parts = [];
            this.partsMap = {};
            this.rolePaths = {};
            this.parser = new DocumentParser(options);
        }
        async load(buffer) {
            const files = unzipSync(new Uint8Array(buffer));
            this.package = new WorkerOpenXmlPackage(normalizeFiles(files), this.options);
            this.rels = await this.package.loadRelationships();
            await Promise.all(topLevelRels.map(rel => {
                const relationship = this.rels.find(x => x.type === rel.type) ?? rel;
                return this.loadRelationshipPart(relationship.target, relationship.type);
            }));
            return {
                rels: this.rels,
                parts: this.parts.map(part => this.serializePart(part)).filter(Boolean),
                rolePaths: this.rolePaths,
                files: Object.entries(this.package.files).map(([path, file]) => ({
                    path,
                    buffer: toArrayBuffer(file)
                }))
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
                    part = new DocumentPart(this.package, normalizedPath, this.parser);
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
                    part = new StylesPart(this.package, normalizedPath, this.parser);
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
        serializePart(part) {
            const base = {
                kind: part.__kind,
                path: part.path,
                rels: part.rels
            };
            switch (part.__kind) {
                case "document":
                    return { ...base, body: part.body };
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
                    return {
                        ...base,
                        comments: part.comments,
                        commentMap: part.commentMap
                    };
                case "commentsExtended":
                    return {
                        ...base,
                        comments: part.comments,
                        commentMap: part.commentMap
                    };
                default:
                    return null;
            }
        }
    }
    self.onmessage = async (event) => {
        if (event.data?.type !== "parse")
            return;
        try {
            const loader = new WorkerDocumentLoader(event.data.options);
            const payload = await loader.load(event.data.buffer);
            const transfer = payload.files.map((file) => file.buffer);
            self.postMessage({ type: "parsed", payload }, transfer);
        }
        catch (error) {
            self.postMessage({
                type: "error",
                error: error?.stack || error?.message || String(error)
            });
        }
    };
    function normalizeFiles(files) {
        const result = {};
        for (const [path, file] of Object.entries(files)) {
            result[normalizePath(path)] = file;
        }
        return result;
    }
    function normalizePath(path) {
        return path.startsWith("/") ? path.substring(1) : path;
    }
    function toArrayBuffer(file) {
        return file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
    }

})();
//# sourceMappingURL=docx-preview-worker.js.map
