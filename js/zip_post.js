/* L7.5 订舱单导出后处理（零依赖）：xlsx(zip) 层注入 ExcelJS 不支持的 OOXML 部件。
 *
 * 为什么需要它：ExcelJS 写 xlsx 时会丢失订舱单模板的「必备要素」——
 *   ① ActiveX 复选框控件链（GEODIS 6 组 / KLN 13 组）：
 *      xl/activeX/activeX*.bin|xml(+rels)、sheet <controls> 块、xl/ctrlProps/ctrlProp*.xml、
 *      drawing 里的 <xdr:sp macro="[0]!组合66_Click">、VML 的 _x0000_t201 形状、
 *      [Content_Types].xml 里的 activeX/ctrlProp Override
 *   ② 打印缩放总开关 <pageSetUpPr fitToPage="1"/>（ExcelJS 只写 pageSetup@fitToWidth，
 *      缺总开关时 Excel 不缩放 → A4 塞不进一页宽）
 *
 * 设计约束（用户铁律）：只作用于订舱单 kind==='booking'；失败必须【原样返回】不阻断导出；
 *   不引入第三方依赖（自写 zip 读写，浏览器用 CompressionStream，Node 用 zlib）。
 *
 * 算法要点（踩过的坑，勿轻易改）：
 *   - sheet 文件名会被 ExcelJS 重编号（源 sheet1.xml → 输出 sheet7.xml）→
 *     必须按 workbook.xml 的 sheet 顺序 + workbook.xml.rels 建立「源↔输出」映射，禁止硬编码。
 *   - rId 必须动态分配 maxRid(outRels)+1 起 —— 源模板 rId 与输出 rId 编号不一致，直接追加会撞车
 *     （实测出现重复 rId4：drawing 与 activeX 抢同一个 id）。
 *   - KLN 的控件包在 <mc:AlternateContent> 里 → 必须【整块搬运】，元素级提取 <xdr:sp> 会丢外层
 *     包装导致 "unmatched closing tag: xdr:wsDr"。
 *   - ExcelJS 输出缺 vmlDrawing 的 .rels（源模板有 EMF 关系）→ 需补回。
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.zipPost = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ───────────────────────── zip 读写（零依赖） ─────────────────────────

  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /** 解压/压缩后端：浏览器 CompressionStream，Node zlib。 */
  function makeCodec() {
    if (typeof DecompressionStream !== 'undefined' && typeof CompressionStream !== 'undefined') {
      return {
        kin: 'web',
        inflate: function (u8) {
          return new Promise(function (res, rej) {
            try {
              var ds = new DecompressionStream('deflate-raw');
              var w = ds.writable.getWriter(); w.write(u8); w.close();
              var chunks = [], total = 0, r = ds.readable.getReader();
              (function pump() {
                r.read().then(function (d) {
                  if (d.done) {
                    var out = new Uint8Array(total), off = 0;
                    chunks.forEach(function (c) { out.set(c, off); off += c.length; });
                    res(out); return;
                  }
                  chunks.push(d.value); total += d.value.length; pump();
                }, rej);
              })();
            } catch (e) { rej(e); }
          });
        },
        deflate: function (u8) {
          return new Promise(function (res, rej) {
            try {
              var cs = new CompressionStream('deflate-raw');
              var w = cs.writable.getWriter(); w.write(u8); w.close();
              var chunks = [], total = 0, r = cs.readable.getReader();
              (function pump() {
                r.read().then(function (d) {
                  if (d.done) {
                    var out = new Uint8Array(total), off = 0;
                    chunks.forEach(function (c) { out.set(c, off); off += c.length; });
                    res(out); return;
                  }
                  chunks.push(d.value); total += d.value.length; pump();
                }, rej);
              })();
            } catch (e) { rej(e); }
          });
        }
      };
    }
    try {
      var zlib = require('zlib');
      return {
        kin: 'node',
        inflate: function (u8) { return Promise.resolve(new Uint8Array(zlib.inflateRawSync(Buffer.from(u8)))); },
        deflate: function (u8) { return Promise.resolve(new Uint8Array(zlib.deflateRawSync(Buffer.from(u8), { level: 6 }))); }
      };
    } catch (e) { return null; }
  }

  var te = new TextEncoder();
  var td = new TextDecoder();

  /** 解析 zip → { entries: {name: e}, codec } ; e = {method, raw, _u8} */
  async function unzip(u8) {
    var codec = makeCodec();
    if (!codec) throw new Error('ZIP_CODEC_UNAVAILABLE');
    var dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    var eocd = -1;
    for (var i = u8.length - 22; i >= 0 && i >= u8.length - 65558; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('ZIP_EOCD_NOT_FOUND');
    var cnt = dv.getUint16(eocd + 10, true);
    var cdOff = dv.getUint32(eocd + 16, true);

    var entries = {}, p = cdOff;
    for (var n = 0; n < cnt; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      var method = dv.getUint16(p + 10, true);
      var csize = dv.getUint32(p + 20, true);
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var cmtLen = dv.getUint16(p + 32, true);
      var lho = dv.getUint32(p + 42, true);
      var name = td.decode(u8.subarray(p + 46, p + 46 + nameLen));
      var lNameLen = dv.getUint16(lho + 26, true);
      var lExtraLen = dv.getUint16(lho + 28, true);
      var dataStart = lho + 30 + lNameLen + lExtraLen;
      entries[name] = { method: method, raw: u8.subarray(dataStart, dataStart + csize), _u8: null, _codec: codec };
      p += 46 + nameLen + extraLen + cmtLen;
    }
    return { entries: entries, codec: codec };
  }

  async function readEntry(entries, name) {
    var e = entries[name];
    if (!e) return null;
    if (e._u8) return td.decode(e._u8);
    if (e.method === 0) { e._u8 = e.raw; return td.decode(e.raw); }
    e._u8 = await e._codec.inflate(e.raw);
    return td.decode(e._u8);
  }
  async function readEntryBytes(entries, name) {
    var e = entries[name];
    if (!e) return null;
    if (e._u8) return e._u8;
    if (e.method === 0) { e._u8 = e.raw; return e.raw; }
    e._u8 = await e._codec.inflate(e.raw);
    return e._u8;
  }
  function setText(entries, name, str) { entries[name] = { method: 8, raw: null, _u8: te.encode(str), _codec: entries.__codec }; }
  function setBytes(entries, name, u8) { entries[name] = { method: 8, raw: null, _u8: u8, _codec: entries.__codec }; }

  /** 打包回 zip（全部 deflate；UTF-8 标志；固定时间戳保证可复现） */
  async function rezip(entries, codec) {
    var names = Object.keys(entries).filter(function (k) { return k.indexOf('__') !== 0; });
    var locals = [], centrals = [], offset = 0;
    for (var i = 0; i < names.length; i++) {
      var name = names[i], e = entries[name];
      var data = e._u8 || (e.method === 0 ? e.raw : await codec.inflate(e.raw));
      var comp = await codec.deflate(data);
      var nameU8 = te.encode(name);
      var crc = crc32(data);
      var lh = new Uint8Array(30 + nameU8.length), ldv = new DataView(lh.buffer);
      ldv.setUint32(0, 0x04034b50, true); ldv.setUint16(4, 20, true); ldv.setUint16(6, 0x0800, true);
      ldv.setUint16(8, 8, true); ldv.setUint16(10, 0, true); ldv.setUint16(12, 0x2821, true);
      ldv.setUint32(14, crc, true); ldv.setUint32(18, comp.length, true); ldv.setUint32(22, data.length, true);
      ldv.setUint16(26, nameU8.length, true); ldv.setUint16(28, 0, true);
      lh.set(nameU8, 30);
      locals.push(lh, comp);

      var ch = new Uint8Array(46 + nameU8.length), cdv = new DataView(ch.buffer);
      cdv.setUint32(0, 0x02014b50, true); cdv.setUint16(4, 20, true); cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0x0800, true); cdv.setUint16(10, 8, true); cdv.setUint16(12, 0, true);
      cdv.setUint16(14, 0x2821, true); cdv.setUint32(16, crc, true); cdv.setUint32(20, comp.length, true);
      cdv.setUint32(24, data.length, true); cdv.setUint16(28, nameU8.length, true);
      cdv.setUint16(30, 0, true); cdv.setUint16(32, 0, true); cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true); cdv.setUint32(38, 0, true); cdv.setUint32(42, offset, true);
      ch.set(nameU8, 46);
      centrals.push(ch);
      offset += lh.length + comp.length;
    }
    var cdSize = centrals.reduce(function (a, b) { return a + b.length; }, 0);
    var eocd = new Uint8Array(22), edv = new DataView(eocd.buffer);
    edv.setUint32(0, 0x06054b50, true); edv.setUint16(4, 0, true); edv.setUint16(6, 0, true);
    edv.setUint16(8, names.length, true); edv.setUint16(10, names.length, true);
    edv.setUint32(12, cdSize, true); edv.setUint32(16, offset, true); edv.setUint16(20, 0, true);
    var parts = locals.concat(centrals, [eocd]);
    var total = parts.reduce(function (a, b) { return a + b.length; }, 0);
    var out = new Uint8Array(total), o = 0;
    parts.forEach(function (p2) { out.set(p2, o); o += p2.length; });
    return out;
  }

  // ───────────────────────── OOXML 辅助 ─────────────────────────

  /** workbook.xml 的 sheet 顺序 → 实际 xml 部件路径（经 workbook.xml.rels 解析）。 */
  async function sheetOrder(entries) {
    var wb = await readEntry(entries, 'xl/workbook.xml');
    var rels = await readEntry(entries, 'xl/_rels/workbook.xml.rels');
    if (!wb || !rels) return [];
    var rid2 = {};
    (rels.match(/<Relationship[^>]*\/>/g) || []).forEach(function (r) {
      var id = (r.match(/Id="([^"]+)"/) || [])[1];
      var tg = (r.match(/Target="([^"]+)"/) || [])[1];
      if (id && tg && /worksheets\//.test(tg)) rid2[id] = 'xl/' + tg.replace(/^\.\.\//, '').replace(/^\/?xl\//i, '');
    });
    var out = [];
    (wb.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (s) {
      var rid = (s.match(/r:id="([^"]+)"/) || [])[1];
      var name = (s.match(/name="([^"]*)"/) || [])[1];
      if (rid && rid2[rid]) out.push({ name: name, path: rid2[rid] });
    });
    return out;
  }

  /** 提取含控件的完整块（整块搬运，勿做元素级提取）。 */
  function extractControlBlocks(drawingXml) {
    var blocks = [], m;
    var a = /<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g;
    while ((m = a.exec(drawingXml))) if (/<xdr:sp[\s>]/.test(m[0])) blocks.push(m[0]);
    if (blocks.length) return { blocks: blocks, style: 'altContent' };
    var t = /<xdr:twoCellAnchor\b[\s\S]*?<\/xdr:twoCellAnchor>/g;
    while ((m = t.exec(drawingXml))) if (/<xdr:sp[\s>]/.test(m[0]) && /macro=/.test(m[0])) blocks.push(m[0]);
    return { blocks: blocks, style: 'twoCellAnchor' };
  }

  function maxRid(xml) {
    var mx = 0, m;
    var r1 = /[rR]:id="rId(\d+)"/g, r2 = /Id="rId(\d+)"/g;
    while ((m = r1.exec(xml))) { var n = parseInt(m[1], 10); if (n > mx) mx = n; }
    while ((m = r2.exec(xml))) { var n2 = parseInt(m[1], 10); if (n2 > mx) mx = n2; }
    return mx;
  }
  function esc(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  // ───────────────────────── 后处理主体 ─────────────────────────

  /**
   * 订舱单导出后处理。
   * @param {ArrayBuffer|Uint8Array} buffer   ExcelJS 写出的 xlsx
   * @param {ArrayBuffer|Uint8Array} [srcBuffer]  源模板 xlsx（提供才做 ActiveX 注入）
   * @returns {Promise<Uint8Array>} 处理后 xlsx；任何异常 → 原样返回入参（不阻断导出）
   */
  async function postProcessBooking(buffer, srcBuffer) {
    var u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    try {
      var out = await unzip(u8);
      var E = out.entries; E.__codec = out.codec;
      var codec = out.codec;
      var rep = { fitSheets: 0, parts: 0, ct: 0, ctrl: 0, renamed: 0, sp: 0, vmlShape: 0, vmlRels: 0 };

      // ── A) A4 一页宽：注入 <pageSetUpPr fitToPage="1"/>（真正的总开关）
      for (var k in E) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(k)) continue;
        var s = await readEntry(E, k);
        if (s === null) continue;
        if (!/<pageSetUpPr\b[^>]*fitToPage="1"/.test(s)) {
          if (/<pageSetUpPr\b[^>]*\/>/.test(s)) {
            s = s.replace(/<pageSetUpPr\b[^>]*\/>/, function (mm) {
              return /fitToPage=/.test(mm) ? mm.replace(/fitToPage="[^"]*"/, 'fitToPage="1"') : mm.replace('/>', ' fitToPage="1"/>');
            });
          } else if (/<sheetPr\b[^>]*\/>/.test(s)) {
            s = s.replace(/<sheetPr\b([^>]*)\/>/, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
          } else if (/<sheetPr\b[^>]*>/.test(s)) {
            s = s.replace(/(<sheetPr\b[^>]*>)/, '$1<pageSetUpPr fitToPage="1"/>');
          } else {
            s = s.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
          }
        }
        if (/<pageSetup\b[^>]*\/>/.test(s)) {
          s = s.replace(/<pageSetup\b[^>]*\/>/, function (mm) {
            var r = mm;
            r = /fitToWidth=/.test(r) ? r.replace(/fitToWidth="[^"]*"/, 'fitToWidth="1"') : r.replace('/>', ' fitToWidth="1"/>');
            if (/fitToHeight=/.test(r)) r = r.replace(/fitToHeight="[^"]*"/, 'fitToHeight="0"');
            return r;
          });
        } else {
          s = s.replace(/(<\/worksheet>)/, '<pageSetup fitToWidth="1" fitToHeight="0"/></worksheet>');
        }
        setText(E, k, s);
        rep.fitSheets++;
      }

      // ── B) ActiveX 复选框链注入（需源模板）
      if (srcBuffer) {
        var srcU8 = srcBuffer instanceof Uint8Array ? srcBuffer : new Uint8Array(srcBuffer);
        var srcEntries = (await unzip(srcU8)).entries;
        srcEntries.__codec = codec;

        // B1) 独立部件：activeX / ctrlProps
        for (var sn in srcEntries) {
          if (/^xl\/(activeX|ctrlProps)\//.test(sn)) {
            var b = await readEntryBytes(srcEntries, sn);
            if (b) { setBytes(E, sn, b); rep.parts++; }
          }
        }

        // B2) [Content_Types].xml 补 activeX/ctrlProp Override
        try {
          var ctS = await readEntry(srcEntries, '[Content_Types].xml');
          var ctO = await readEntry(E, '[Content_Types].xml');
          if (ctS && ctO) {
            var ovs = ctS.match(/<Override[^>]*(activeX|ctrlProp)[^>]*\/>/g) || [];
            if (ovs.length && !/activeX/i.test(ctO)) {
              setText(E, '[Content_Types].xml', ctO.replace('</Types>', ovs.join('') + '</Types>'));
              rep.ct = ovs.length;
            }
          }
        } catch (eCT) {}

        // B3) 按 sheet 顺序映射，注入 <controls> + rels
        var os = await sheetOrder(srcEntries), oo = await sheetOrder(E);
        for (var si = 0; si < Math.min(os.length, oo.length); si++) {
          var sPath = os[si].path, oPath = oo[si].path;
          if (!E[oPath]) continue;
          var sXml = await readEntry(srcEntries, sPath);
          var oXml = await readEntry(E, oPath);
          if (!sXml || !oXml) continue;
          if (/<controls>/.test(oXml)) continue;              // 已有 → 不重复注入
          var ctlM = sXml.match(/<controls>[\s\S]*?<\/controls>/);
          if (!ctlM) continue;

          var sRelPath = sPath.replace(/([^/]+)$/, '_rels/$1.rels');
          var oRelPath = oPath.replace(/([^/]+)$/, '_rels/$1.rels');
          var rS = await readEntry(srcEntries, sRelPath);
          if (rS === null) continue;
          var rO = await readEntry(E, oRelPath);
          if (rO === null) rO = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

          // 只搬 control/ctrlProp 关系
          var ctlRels = (rS.match(/<Relationship[^>]*\/>/g) || []).filter(function (r) {
            return /relationships\/(control|ctrlProp)"/.test(r);
          });
          if (!ctlRels.length) continue;

          var next = maxRid(rO) + 1, map = {}, newRels = [];
          ctlRels.forEach(function (r) {
            var oldId = (r.match(/Id="(rId\d+)"/) || [])[1];
            if (!oldId) return;
            var newId = 'rId' + (next++);
            map[oldId] = newId;
            newRels.push(r.replace(/Id="rId\d+"/, 'Id="' + newId + '"'));
          });
          rep.renamed += newRels.length;

          // <controls> 块里的 r:id 同步改写
          var newCtl = ctlM[0];
          Object.keys(map).forEach(function (o) {
            newCtl = newCtl.split('r:id="' + o + '"').join('r:id="' + map[o] + '"');
          });
          rep.ctrl += (newCtl.match(/<control\b/g) || []).length;

          // 注入位置：legacyDrawing 之后，否则 </worksheet> 之前
          var oOut = /<legacyDrawing [^>]*\/>/.test(oXml)
            ? oXml.replace(/(<legacyDrawing [^>]*\/>)/, '$1' + newCtl)
            : oXml.replace('</worksheet>', newCtl + '</worksheet>');
          setText(E, oPath, oOut);
          setText(E, oRelPath, rO.replace('</Relationships>', newRels.join('') + '</Relationships>'));
        }

        // B4) drawing：整块搬运控件（mc:AlternateContent / twoCellAnchor）
        var drawingNames = [];
        for (var dn in srcEntries) if (/^xl\/drawings\/drawing\d+\.xml$/.test(dn)) drawingNames.push(dn);
        // 输出 drawing 名称可能不同 —— 按顺序配对
        var outDrawings = [];
        for (var on2 in E) if (/^xl\/drawings\/drawing\d+\.xml$/.test(on2)) outDrawings.push(on2);
        outDrawings.sort();
        for (var di = 0; di < drawingNames.length; di++) {
          var dSrc = await readEntry(srcEntries, drawingNames[di]);
          if (!dSrc) continue;
          var ex = extractControlBlocks(dSrc);
          if (!ex.blocks.length) continue;
          var dTarget = E[drawingNames[di]] ? drawingNames[di] : outDrawings[0];
          if (!dTarget) continue;
          var dOut = await readEntry(E, dTarget);
          if (dOut === null) continue;
          if (/<xdr:sp[\s>]/.test(dOut)) continue;   // 已有控件 → 跳过
          setText(E, dTarget, dOut.replace('</xdr:wsDr>', ex.blocks.join('') + '</xdr:wsDr>'));
          rep.sp += ex.blocks.length;
        }

        // B5) VML：补 _x0000_t201 形状 + 补回 vml 的 .rels
        var vmlSrc = [];
        for (var vn in srcEntries) if (/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(vn)) vmlSrc.push(vn);
        vmlSrc.sort();
        var vmlOut = [];
        for (var vo in E) if (/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(vo)) vmlOut.push(vo);
        vmlOut.sort();
        for (var vi = 0; vi < vmlSrc.length; vi++) {
          var vSrc = vmlSrc[vi];
          var vTarget = E[vSrc] ? vSrc : vmlOut[0];
          if (!vTarget) continue;
          var vS = await readEntry(srcEntries, vSrc);
          var vO = await readEntry(E, vTarget);
          if (vS === null || vO === null) continue;
          if (vO.indexOf('_x0000_t201') === -1) {
            var st = (vS.match(/<v:shapetype[^>]*_x0000_t201[\s\S]*?<\/v:shapetype>/) || [])[0] || '';
            var shapes = vS.match(/<v:shape[^>]*type="#_x0000_t201"[\s\S]*?<\/v:shape>/g) || [];
            var add = st + shapes.join('');
            if (add) {
              setText(E, vTarget, vO.replace(/<\/xml>\s*$/, add + '</xml>'));
              rep.vmlShape += shapes.length;
            }
          }
          // 补回 vml 的 .rels（ExcelJS 输出缺失，源有 EMF 关系）
          var vRel = vSrc.replace(/([^/]+)$/, '_rels/$1.rels');
          if (srcEntries[vRel] && !E[vRel]) {
            var vrb = await readEntryBytes(srcEntries, vRel);
            if (vrb) { setBytes(E, vRel, vrb); rep.vmlRels++; }
          }
        }

        // B6) 补回 media（LOGO 等）—— ExcelJS 已保留，但源有新增图时补齐
        for (var mn in srcEntries) {
          if (/^xl\/media\//.test(mn) && !E[mn]) {
            var mb = await readEntryBytes(srcEntries, mn);
            if (mb) setBytes(E, mn, mb);
          }
        }
      }

      E.__report = rep;
      return await rezip(E, codec);
    } catch (e) {
      try { if (typeof console !== 'undefined') console.warn('[zipPost] 订舱单后处理失败，已回退原始文件：', e && e.message); } catch (e2) {}
      return u8;
    }
  }

  return { postProcessBooking: postProcessBooking, _unzip: unzip, _rezip: rezip };
});
