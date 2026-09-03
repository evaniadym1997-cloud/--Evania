/* ============================================================
 * hayloft-comic 一键导出 PPT（真实文本框式 .pptx）
 * 依赖：assets/vendor/pptxgen.bundle.js（window.PptxGenJS）
 *
 * 设计要点：
 *  - 只导出 .deck 中带 data-title 的真实静态讲评 slide；
 *    跳过 JS 动态模板页 / 历史目录 / 生成器相关页 / 重复速查页
 *  - 中文一律华文楷体(STKaiti)，英文/数字一律 Times New Roman；
 *    混排文本按字符切 run（pptxgenjs addText runs 数组）
 *  - 字号分层：主标题 34/封面36、要点 26、原文引文 17、长文 15
 *  - 防溢出：文本先按字体/字宽切成“单行 chunk”，逐 chunk 垂直排布，
 *    放不下即拆新页；逐格估算高度，不依赖 PPT 端自动折行兜底
 *  - 图片页：读取 DOM 真实 src -> dataURL 等比 contain 放入
 *  - file:// 协议无法可靠读取图片时给出友好中文提示
 * ============================================================ */
(function () {
  'use strict';

  function main() {
  /* ---------------- 基础常量 ---------------- */
  var CN_FONT = '华文楷体';            // STKaiti
  var EN_FONT = 'Times New Roman';
  var SLIDE_W = 13.333, SLIDE_H = 7.5; // 16:9 英寸
  var TEXT_LEFT = 0.7;                  // 正文/标题左边距
  var TEXT_W = 11.93;                   // 正文可用宽度
  var HEADER_BOTTOM = 1.68;             // 标题区结束 y
  var BODY_BOTTOM = 7.06;               // 正文区底部
  var COL = {
    title: '7A3B12', kicker: '9A8069', head: '8A4A1A',
    body: '2B2620', red: 'C00000', grey: '77716A', cream: 'FDF6EC', accent: 'C98B4B'
  };
  // 每类文字的默认字号（pt）与行距倍率
  var KIND_FONT = {
    title: 34, coverTitle: 36, kicker: 15, kicker2: 20, head: 26, bullet: 26,
    quote: 17, passage: 15, tip: 17, guide: 17, credit: 13
  };

  /* ---------------- 字符工具 ---------------- */
  function isCJKChar(ch) {
    var c = ch.codePointAt(0);
    return (c >= 0x3400 && c <= 0x4DBF) || (c >= 0x4E00 && c <= 0x9FFF) ||
      (c >= 0x2E80 && c <= 0x2EFF) || (c >= 0x3000 && c <= 0x303F) ||
      (c >= 0x2F00 && c <= 0x2FDF) || (c >= 0xF900 && c <= 0xFAFF) ||
      (c >= 0xFF00 && c <= 0xFFEF) || (c >= 0x2018 && c <= 0x201F) ||
      (c >= 0x2014 && c <= 0x2015) || c === 0x2026 || c === 0x3000;
  }
  // 去掉 emoji / 装饰符号（保留 ①-⑨、·、引号等教学字符）
  function cleanText(s) {
    return String(s)
      .replace(/\uFEFF/g, '')
      .replace(/[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{2190}-\u{21FF}\u{FE0F}\u{2764}\u{2705}\u{1F1E6}-\u{1F1FF}]/gu, '')
      .replace(/[ \t\u00A0]+/g, ' ')
      .replace(/^\s+|\s+$/g, '');
  }
  // 段内 seg 文本：统一空白但保留跨标签边界空格（供排版/字符级 run 使用）
  function segText(s) {
    return String(s)
      .replace(/\uFEFF/g, '')
      .replace(/[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{2190}-\u{21FF}\u{FE0F}\u{2764}\u{2705}\u{1F1E6}-\u{1F1FF}]/gu, '')
      .replace(/[ \t\u00A0\r\n]+/g, ' ');
  }
  // 估算字符宽（单位：em；CJK=1，TNR 拉丁约 0.52，空格/符号按常规）
  function chEm(ch) {
    if (ch === ' ' || ch === '\u00A0') return 0.28;
    if (isCJKChar(ch)) return 1.0;
    if (/[0-9]/.test(ch)) return 0.53;
    if (/[A-Za-z]/.test(ch)) return 0.52;
    if (/[,.;:!?'")(\[\]{}<>/\\|@#$%^&*_+\-=~`]/.test(ch)) return 0.28;
    return 0.6;
  }

  /* ---------------- DOM -> 富文本 ---------------- */
  var SKIP_SEL = '.notes,.teacher-note,.teacher-note-label,.cell-hint,.p-tag,' +
    '.panel-badge,.pill-row,.emotion-tag,.story-card,.btn,button,input,select,textarea,script,style' +
    ',.tf-hint,.tf-note,.jump-hint';
  // story-card 按图片页整体处理，故文本提取时跳过其内部

  function isSkipEl(el) {
    if (!el || el.nodeType !== 1) return false;
    if (el.closest && el.closest('.story-card')) return true;
    try { return el.matches(SKIP_SEL); } catch (e) { return false; }
  }

  // 递归收集富文本：{t,bold,red}，遇 <br> 或块级元素边界插入 {br:true}
  var BLOCK_TAGS = { P: 1, DIV: 1, LI: 1, UL: 1, OL: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, TD: 1, TR: 1, TABLE: 1, BLOCKQUOTE: 1 };
  function richOf(rootEl) {
    var out = [];
    function isBlockTag(node) {
      return node && node.nodeType === 1 && BLOCK_TAGS[node.tagName] === 1;
    }
    function walk(node, bold, red, italic) {
      if (node.nodeType === 3) { // text
        var v = node.nodeValue;
        if (v && v.trim()) out.push({ t: v, bold: bold, red: red, italic: italic });
        return;
      }
      if (node.nodeType !== 1) return;
      if (isSkipEl(node)) return;
      var tag = node.tagName;
      if (tag === 'IMG') return;
      var nb = bold || tag === 'B' || tag === 'STRONG';
      var ni = italic || (node.classList && node.classList.contains('en'));
      if (tag === 'BR') { out.push({ br: true }); return; }
      if (tag === 'INPUT') return;
      var kids = node.childNodes;
      var firstRun = true;
      for (var i = 0; i < kids.length; i++) {
        var kid = kids[i];
        // 块级兄弟之间自动断段，避免 model-passage 等多段容器被无分隔拼接
        if (isBlockTag(kid) && kid !== rootEl && !firstRun && out.length && !(out[out.length - 1].br)) {
          out.push({ br: true });
        }
        walk(kid, nb, red, ni);
        if (isBlockTag(kid) && kid !== rootEl && out.length && !(out[out.length - 1].br)) {
          out.push({ br: true });
        }
        if (isBlockTag(kid)) firstRun = false;
      }
    }
    for (var i = 0; i < rootEl.childNodes.length; i++) walk(rootEl.childNodes[i], false, false, false);
    return out;
  }
  // rich -> 段落数组（按 br 拆段）
  function parasOf(rich) {
    var paras = [[]];
    for (var i = 0; i < rich.length; i++) {
      var s = rich[i];
      if (s.br) { paras.push([]); }
      else { paras[paras.length - 1].push(s); }
    }
    // 去空段 & 合并相邻同型
    var res = [];
    for (var j = 0; j < paras.length; j++) {
      var cur = paras[j].map(function (x) { return { seg: x, txt: segText(x.t) }; })
        .filter(function (x) { return x.txt.replace(/ /g, '').length > 0; });
      if (!cur.length) continue;
      if (cur.length) { cur[0].txt = cur[0].txt.replace(/^ +/, ''); cur[cur.length - 1].txt = cur[cur.length - 1].txt.replace(/ +$/, ''); }
      cur = cur.filter(function (x) { return x.txt.length > 0; });
      if (!cur.length) continue;
      var merged = [];
      for (var k = 0; k < cur.length; k++) {
        var last = merged[merged.length - 1];
        var segT = cur[k].txt;
        if (last && last.bold === cur[k].seg.bold && last.red === cur[k].seg.red && last.italic === cur[k].seg.italic) {
          last.t += segT;
        } else {
          merged.push({ t: segT, bold: !!cur[k].seg.bold, red: !!cur[k].seg.red, italic: !!cur[k].seg.italic });
        }
      }
      res.push(merged);
    }
    return res;
  }
  function paraTextEm(para) {
    var em = 0;
    for (var i = 0; i < para.length; i++) {
      var seg = para[i], txt = segText(seg.t);
      for (var j = 0; j < txt.length; j++) em += chEm(txt[j]);
    }
    return em;
  }
  // 段宽估算（单位 em）
  function paraEm(para) {
    var em = 0;
    for (var i = 0; i < para.length; i++) {
      var txt = segText(para[i].t);
      for (var j = 0; j < txt.length; j++) em += chEm(txt[j]);
    }
    return em;
  }
  function runFontFace(t) {
    // 单个字符判定：CJK / 全角标点走中文字体，其余走英文字体
    if (!t) return EN_FONT;
    var c = t.codePointAt(0);
    return (c >= 0x3400 && c <= 0x4DBF) || (c >= 0x4E00 && c <= 0x9FFF) ||
      (c >= 0x2E80 && c <= 0x2EFF) || (c >= 0x3000 && c <= 0x303F) ||
      (c >= 0x2F00 && c <= 0x2FDF) || (c >= 0xF900 && c <= 0xFAFF) ||
      (c >= 0xFF00 && c <= 0xFFEF) || (c >= 0x2018 && c <= 0x201F) ||
      (c >= 0x2014 && c <= 0x2015) || c === 0x2026 || c === 0x3000
      ? CN_FONT : EN_FONT;
  }
  // 单段（para）按其总 em 决定字号层
  function kindForPara(para, tag, cardCtx) {
    var em = paraEm(para);
    var hasLong = para.some(function (s) { return cleanText(s.t).length > 40; });
    if (tag && /^TH$/i.test(tag)) return 'quote';
    if (tag && /^H3$/i.test(tag)) return 'head';
    if (cardCtx && cardCtx.modelPassage) return 'passage';
    if (cardCtx && cardCtx.cardPara) return 'bullet';
    if (em > 150) return 'passage';
    return 'bullet';
  }

  /* ---------------- 原子选择 / 分类 ---------------- */
  // 页面里出现的“原子容器”，其内部有 h3/p/span，但不应再被外层泛化 p 重复收集
  var ATOM_SEL = '.page-kicker,.kicker,.dim,.scene-desc,.crit-card,.model-passage,.cell-table,.mindmap,.mini-flow,.check-list,.guide-cell,h1,h2,h3,p';
  function collectAtoms(sectionEl) {
    // querySelectorAll 本身按文档序返回；过滤掉被外层原子容器覆盖的内层节点即可保序
    var all = Array.prototype.slice.call(sectionEl.querySelectorAll(ATOM_SEL));
    var keep = [];
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      var inside = el.parentElement;
      var nested = false;
      while (inside && inside !== sectionEl) {
        if (inside.matches && inside.matches(ATOM_SEL) && inside !== el) { nested = true; break; }
        inside = inside.parentElement;
      }
      if (!nested) keep.push(el);
    }
    return keep;
  }

  function isDimHint(el) {
    var t = el.textContent || '';
    return /按\s*[←→]|翻页|T 切主题|F 全屏|全屏|涂鸦|点击标题|点击.*展开|Image Prompt|键盘|鼠标/.test(t);
  }

  // 解析一个文本型 section 为行块列表 lines
  // line: {kind, para, sep:'afterBlock'|'afterPara'|''}
  function parseTextSection(sectionEl) {
    var blocks = [];
    var titleEl = sectionEl.querySelector('h1,h2');
    var titlePara = titleEl ? parasOf(richOf(titleEl))[0] || [] : null;
    var titleKind = titleEl && titleEl.tagName === 'H1' ? 'coverTitle' : 'title';

    // kicker：取 section 直属第一个 page-kicker / kicker
    var kickEl = sectionEl.querySelector('.page-kicker,.kicker');
    var kickPara = kickEl ? parasOf(richOf(kickEl))[0] || [] : null;

    var atoms = collectAtoms(sectionEl);
    for (var i = 0; i < atoms.length; i++) {
      var el = atoms[i];
      if (isSkipEl(el)) continue;
      var tag = el.tagName;
      var cls = (typeof el.className === 'string') ? el.className : (el.className && el.className.baseVal) || '';
      var clsHas = function (c) { return cls.split(/\s+/).indexOf(c) >= 0; };
      if (el === titleEl || (titleEl && titleEl.contains(el) && tag !== 'H1' && tag !== 'H2')) continue;
      if (el === kickEl) continue;

      if (el.matches('.page-kicker,.kicker')) { /* 已跳过第一个，若有第二个 kicker 作导语处理 */
        var kp = parasOf(richOf(el))[0];
        if (kp) blocks.push({ kind: 'kicker2', para: kp });
        continue;
      }
      if (el.matches('.dim')) {
        if (isDimHint(el)) continue;
        var dimParas = parasOf(richOf(el));
        for (var d = 0; d < dimParas.length; d++) blocks.push({ kind: 'tip', para: dimParas[d] });
        continue;
      }
      if (el.matches('.scene-desc')) {
        var sdParas = parasOf(richOf(el));
        for (var s2 = 0; s2 < sdParas.length; s2++) {
          var em = paraEm(sdParas[s2]);
          blocks.push({ kind: em <= 34 ? 'tip' : 'passage', para: sdParas[s2] });
        }
        continue;
      }
      if (el.matches('.crit-card')) {
        var cardHeadEl = el.querySelector(':scope > h3, h3');
        var cardHead = cardHeadEl ? (parasOf(richOf(cardHeadEl))[0] || null) : null;
        if (cardHead) blocks.push({ kind: 'head', para: cardHead, cardStart: true });
        var kids = el.children;
        for (var c = 0; c < kids.length; c++) {
          var kid = kids[c];
          if (kid === cardHeadEl || isSkipEl(kid)) continue;
          var kt = kid.tagName;
          if (kt === 'P' || kt === 'DIV' || kt === 'UL' || kt === 'OL' || /^H[3-6]$/.test(kt)) {
            var sub = kt === 'UL' || kt === 'OL' ? Array.prototype.slice.call(kid.querySelectorAll('li')) : [kid];
            for (var u = 0; u < sub.length; u++) {
              var sp = parasOf(richOf(sub[u]));
              for (var s3 = 0; s3 < sp.length; s3++) {
                // 长原文/整段范文降为 passage(15)，短要点维持 26
                blocks.push({ kind: paraEm(sp[s3]) > 62 ? 'passage' : 'bullet', para: sp[s3] });
              }
            }
          }
        }
        continue;
      }
      if (el.matches('.model-passage')) {
        var mp = parasOf(richOf(el));
        for (var m = 0; m < mp.length; m++) blocks.push({ kind: 'passage', para: mp[m] });
        continue;
      }
      if (el.matches('.cell-table')) {
        // 表头长引文（quote），表体行要点（bullet）
        var thead = el.querySelector('thead');
        if (thead) {
          var ths = thead.querySelectorAll('th');
          for (var h1 = 0; h1 < ths.length; h1++) {
            var thP = parasOf(richOf(ths[h1]));
            // 长引文（整句原文）-> quote 17pt；短列标题（如“冲突 Conflict”）-> head 26pt
            var thRaw = cleanText(ths[h1].textContent || '');
            var thIsLong = (ths[h1].hasAttribute && ths[h1].hasAttribute('colspan') && thRaw.length > 20) || thRaw.length > 32;
            for (var s4 = 0; s4 < thP.length; s4++) blocks.push({ kind: thIsLong ? 'quote' : 'head', para: thP[s4] });
          }
        }
        var rows = el.querySelectorAll('tbody tr');
        for (var r = 0; r < rows.length; r++) {
          var tds = rows[r].querySelectorAll(':scope > td');
          if (tds.length === 0) continue;
          var merged = tds.length === 1 || (tds[0].querySelector('.cell-body') === null && tds.length > 1);
          if (merged) {
            // 行 = 左标签(加粗) + 右正文，合并成一个要点
            var rowPara = [];
            var leftP = parasOf(richOf(tds[0]))[0] || null;
            var rightP = tds.length > 1 ? parasOf(richOf(tds[1]))[0] || null : null;
            if (leftP) {
              for (var l0 = 0; l0 < leftP.length; l0++) rowPara.push(leftP[l0]);
            }
            if (rightP) {
              for (var l1 = 0; l1 < rightP.length; l1++) rowPara.push({ t: ' ', bold: false, red: false });
              for (var l2 = 0; l2 < rightP.length; l2++) rowPara.push(rightP[l2]);
            }
            if (rowPara.length) blocks.push({ kind: 'bullet', para: rowPara });
          } else {
            for (var c2 = 0; c2 < tds.length; c2++) {
              var cellPs = parasOf(richOf(tds[c2]));
              for (var s5 = 0; s5 < cellPs.length; s5++) blocks.push({ kind: 'bullet', para: cellPs[s5] });
            }
          }
        }
        continue;
      }
      if (el.matches('.mindmap')) {
        // hub 与 branch 可能交错出现，必须按 DOM 顺序整体遍历
        var mbAll = el.querySelectorAll('.hub, .branch');
        for (var mb = 0; mb < mbAll.length; mb++) {
          var mbIsHub = mbAll[mb].classList && mbAll[mb].classList.contains('hub');
          var mbParas = parasOf(richOf(mbAll[mb]));
          for (var s7 = 0; s7 < mbParas.length; s7++) blocks.push({ kind: mbIsHub ? 'head' : 'bullet', para: mbParas[s7] });
        }
        continue;
      }
      if (el.matches('.mini-flow')) {
        var fls = el.querySelectorAll('.flow-line');
        for (var fl = 0; fl < fls.length; fl++) {
          var fp = parasOf(richOf(fls[fl]));
          for (var s8 = 0; s8 < fp.length; s8++) blocks.push({ kind: 'quote', para: fp[s8] });
        }
        continue;
      }
      if (el.matches('.check-list,ul,ol')) {
        var lis = el.querySelectorAll('li');
        for (var li = 0; li < lis.length; li++) {
          var lp = parasOf(richOf(lis[li]));
          for (var s9 = 0; s9 < lp.length; s9++) blocks.push({ kind: 'bullet', para: lp[s9] });
        }
        continue;
      }
      if (el.matches('.guide-cell')) {
        var gp = parasOf(richOf(el));
        for (var g = 0; g < gp.length; g++) blocks.push({ kind: 'guide', para: gp[g] });
        continue;
      }
      // 泛化 h/p 兜底
      if (/^H[1-6]$/.test(tag)) {
        var hPara = parasOf(richOf(el))[0];
        if (hPara) blocks.push({ kind: tag === 'H1' || tag === 'H2' ? 'title' : 'head', para: hPara });
        continue;
      }
      if (tag === 'P' || tag === 'DIV') {
        var gps = parasOf(richOf(el));
        for (var s0 = 0; s0 < gps.length; s0++) blocks.push({ kind: 'bullet', para: gps[s0] });
        continue;
      }
      // 其它未知元素：提取纯文本并入 guide 层
      var tRaw = cleanText(el.textContent || '');
      if (tRaw) blocks.push({ kind: 'guide', para: [{ t: tRaw, bold: false, red: false }] });
    }
    return {
      titleKind: titleKind,
      titlePara: titlePara,
      kickPara: kickPara,
      blocks: blocks
    };
  }

  /* ---------------- 单行 chunk 切分 ---------------- */
  // 给定一个 para 与字号，切成多行（每行 em 不超 capEm，带 6% 余量）
  function chunkPara(para, fontSize, widthIn) {
    var capEm = ((widthIn * 72) - 8) / fontSize * 0.94;
    var lines = [];
    var cur = [];
    var acc = 0;
    for (var i = 0; i < para.length; i++) {
      var seg = para[i], txt = segText(seg.t);
      if (!txt) continue;
      for (var j = 0; j < txt.length; j++) {
        var w = chEm(txt[j]);
        if (acc + w > capEm && cur.length > 0) {
          lines.push(cur);
          cur = [];
          acc = 0;
        }
        cur.push({ t: txt[j], bold: seg.bold, red: seg.red, italic: seg.italic });
        acc += w;
      }
    }
    if (cur.length) lines.push(cur);
    if (!lines.length) lines.push([{ t: '', bold: false, red: false, italic: false }]);
    return lines;
  }

  /* ---------------- 页面排版 ---------------- */
  function toPptRuns(para, fontSize, color, forceBold) {
    // 按“字体族/加粗/颜色/斜体”逐字符归类切 run，确保混排段内中英文各自命中正确字体
    var runs = [];
    var buf = '', bufFace = null, bufBold = false, bufColor = null, bufItalic = false;
    function flush() {
      if (buf) {
        runs.push({ text: buf, options: { fontFace: bufFace, fontSize: fontSize, bold: bufBold, color: bufColor, italic: bufItalic } });
        buf = '';
      }
    }
    function emit(ch, face, bold, colr, ital) {
      if (face !== bufFace || bold !== bufBold || colr !== bufColor || ital !== bufItalic) { flush(); bufFace = face; bufBold = bold; bufColor = colr; bufItalic = ital; }
      buf += ch;
    }
    for (var i = 0; i < para.length; i++) {
      var seg = para[i];
      var txt = segText(seg.t);
      if (!txt) continue;
      var bold = !!(seg.bold || seg.red || forceBold);
      var colr = seg.red ? COL.red : (color || COL.body);
      var ital = !!seg.italic;
      var chars = Array.from(txt);
      for (var j = 0; j < chars.length; j++) emit(chars[j], runFontFace(chars[j]), bold, colr, ital);
    }
    flush();
    return runs;
  }

  // 纯字符串 -> 逐字符字体 run（等价于 toPptRuns 的单段包装）
  function segRuns(str, fontSize, color, bold) {
    return toPptRuns([{ t: str, bold: !!bold, red: false }], fontSize, color, !!bold);
  }
  // 单行视觉高度（inch），含少量字面间距
  function lineHeightIn(fontSize) { return fontSize * 1.34 / 72; }

  // 把 block 序列排版为页面数组
  // pageItem: {titlePara,titleKind,kickPara,lines:[{kind,para}]}
  function layoutPages(textSection, maxBulletsPerPage) {
    maxBulletsPerPage = maxBulletsPerPage || 6;
    var pages = [];
    var lines = [];
    var title = textSection.titlePara || [{ t: '', bold: false, red: false }];
    var titleKind = textSection.titleKind || 'title';
    var kick = textSection.kickPara || null;

    // 预生成每 block 的 chunk 行
    var chunkLines = [];
    for (var i = 0; i < textSection.blocks.length; i++) {
      var bk = textSection.blocks[i];
      var fsize = KIND_FONT[bk.kind] || 26;
      if (bk.kind === 'head' && bk.cardStart) fsize = KIND_FONT.head;
      if (bk.kind === 'title') fsize = KIND_FONT.title;
      var chunks = chunkPara(bk.para, fsize, TEXT_W);
      // passage 段落允许更紧凑；chunk 都带 kind
      var paraGap = (bk.kind === 'bullet' || bk.kind === 'head') ? 0.045 : 0.03;
      for (var c = 0; c < chunks.length; c++) {
        chunkLines.push({ kind: bk.kind, para: chunks[c], gap: (c === chunks.length - 1 ? paraGap : 0) });
      }
      chunkLines.push({ kind: 'sep', sep: (bk.kind === 'head' || bk.kind === 'cardStart') ? 0.16 : 0.12 });
    }
    // 需要把块边界保留成块与块间的空隙
    var items = [];
    for (var j = 0; j < chunkLines.length; j++) {
      var cl = chunkLines[j];
      if (cl.kind === 'sep') {
        if (items.length && items[items.length - 1].sep === undefined) items[items.length - 1].sep = cl.sep;
        continue;
      }
      items.push({ kind: cl.kind, para: cl.para, sep: cl.gap });
    }

    var curPage = { titlePara: title, titleKind: titleKind, kickPara: kick, items: [], bulletCount: 0 };
    var curH = 0;
    var maxH = 6.92 - HEADER_BOTTOM; // 页脚上方留白，防压叠

    // “要点行”上限只约束 bullet/head/tip/guide；quote/passage 长文按高度自然排页
    function countKind(kind) {
      return (kind === 'bullet' || kind === 'head' || kind === 'tip' || kind === 'guide' || kind === 'kicker2');
    }

    function closePage() {
      if (curPage.items.length) pages.push(curPage);
      curPage = { titlePara: title, titleKind: titleKind, kickPara: kick, items: [], bulletCount: 0 };
      curH = 0;
    }

    for (var m = 0; m < items.length; m++) {
      var it = items[m];
      var itH = lineHeightIn(KIND_FONT[it.kind] || 26) + (it.sep || 0);
      var isBullet = countKind(it.kind);
      // 若单行都超高（几乎不可能），按逐行拆页
      if (itH > maxH) itH = maxH;
      if ((curH + itH > maxH + 0.01) || (isBullet && curPage.bulletCount >= maxBulletsPerPage && curPage.items.length > 0)) {
        closePage();
      }
      curPage.items.push(it);
      curH += itH;
      if (isBullet) curPage.bulletCount++;
    }
    closePage();
    return pages;
  }

  /* ---------------- 图片页数据 ---------------- */
  function panelOfSection(sectionEl) {
    var cards = Array.prototype.slice.call(sectionEl.querySelectorAll('.story-card'));
    var out = [];
    for (var i = 0; i < cards.length; i++) {
      var img = cards[i].querySelector('img');
      if (!img) continue;
      var labelEl = cards[i].querySelector('.p-label');
      var noteEl = cards[i].querySelector('.p-note');
      var tagEl = cards[i].querySelector('.p-tag');
      out.push({
        src: img.getAttribute('src') || '',
        tag: tagEl ? cleanText(tagEl.textContent) : '',
        label: labelEl ? cleanText(labelEl.textContent) : '',
        note: noteEl ? cleanText(noteEl.textContent) : ''
      });
    }
    return out;
  }

  /* ---------------- 渲染 pptx ---------------- */
  function addHeader(slide, textSection, pageIndex, total) {
    slide.background = { color: COL.cream };
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: 0.085, fill: { color: COL.accent }, line: { type: 'none' } });
    if (pageIndex === undefined) return;
    var kick = textSection && textSection.kickPara;
    var kickY = 0.3;
    if (kick && kick.length) {
      slide.addText(toPptRuns(kick, KIND_FONT.kicker, COL.kicker), {
        x: TEXT_LEFT, y: kickY, w: TEXT_W, h: 0.42,
        align: 'left', valign: 'middle', margin: 0
      });
    }
    var titlePara = (textSection && textSection.titlePara) || [];
    if (titlePara.length) {
      var tf = KIND_FONT[textSection.titleKind] || 34;
      var titleEm = paraEm(titlePara);
      var capEmT = (((TEXT_W * 72) - 8) / tf) * 0.94;
      var titleLines = Math.max(1, Math.ceil(titleEm / capEmT));
      // 标题若需换行则降至 30pt，保证仍处 30-36 档且不挤压正文
      if (titleLines > 1 && tf > 30) { tf = 30; titleLines = Math.max(1, Math.ceil(titleEm / (((TEXT_W * 72) - 8) / tf * 0.94))); }
      var titleBoxH = 0.30 + titleLines * lineHeightIn(tf) * 0.9;
      slide.addText(toPptRuns(titlePara, tf, COL.title), {
        x: TEXT_LEFT, y: 0.66, w: TEXT_W, h: titleBoxH,
        align: 'left', valign: 'top', margin: 0
      });
      return Math.min(1.68, 0.66 + titleBoxH + 0.08);
    }
    return HEADER_BOTTOM;
  }

  function addFooter(slide, idx, total, label) {
    var txt = label || ('Hayloft 讲评 · ' + idx + ' / ' + total);
    slide.addText(segRuns(txt, 12, COL.grey, false), {
      x: 10.7, y: 7.13, w: 2.45, h: 0.3, align: 'right', margin: 0
    });
  }

  function renderTextSlide(pptx, textSection, items, idx, total) {
    var slide = pptx.addSlide();
    var bodyTop = addHeader(slide, textSection, idx, total);
    if (!(bodyTop > 0)) bodyTop = HEADER_BOTTOM;
    var y = bodyTop;
    var maxY = 6.92;
    // 防大量留白：内容不满页时按行等额伸展行高（单行附加 <=0.42in，不改变字号）
    var needH = 0;
    for (var n0 = 0; n0 < items.length; n0++) {
      needH += lineHeightIn(KIND_FONT[items[n0].kind] || 26) + (items[n0].sep || 0);
    }
    var extraPer = items.length ? (maxY - y - needH) / items.length : 0;
    if (extraPer < 0) extraPer = 0;
    if (extraPer > 0.42) extraPer = 0.42;
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var fsize = KIND_FONT[it.kind] || 26;
      var color = COL.body;
      if (it.kind === 'head') color = COL.head;
      if (it.kind === 'kicker2') color = COL.kicker;
      if (it.kind === 'tip') color = COL.head;
      if (it.kind === 'quote' || it.kind === 'guide') color = COL.body;
      var h = lineHeightIn(fsize) + extraPer;
      var runs = toPptRuns(it.para, fsize, color, it.kind === 'head');
      var opts = {
        x: TEXT_LEFT,
        y: y, w: TEXT_W - (it.kind === 'head' ? 0.12 : 0),
        h: h + 0.02,
        align: 'left', valign: 'middle',
        margin: 0,
        shrinkText: false,
        lineSpacingMultiple: 1.0
      };
      if (it.kind === 'quote') { color = COL.body; opts.x = TEXT_LEFT + 0.06; opts.w = TEXT_W - 0.3; }
      try {
        slide.addText(runs, opts);
      } catch (e) {
        // 兜底：异常时整段按普通正文输出
        var txt = it.para.map(function (s) { return segText(s.t); }).join('');
        var fallRuns = segRuns(txt, fsize, color, it.kind === 'head');
        var fopts = { x: opts.x, y: opts.y, w: opts.w, h: opts.h, align: 'left', valign: 'middle', margin: 0 };
        slide.addText(fallRuns, fopts);
      }
      y += h + (it.sep || 0);
      if (y > maxY) { y = maxY; }
    }
    addFooter(slide, idx, total);
    return slide;
  }

  function addImageContain(slide, dataURL, box) {
    // 先拿图片尺寸
    var im = new Image();
    // 由于是 dataURL，同步等待不可行；用 Promise 在外面处理，这里直接返回比例
    return new Promise(function (resolve, reject) {
      im.onload = function () {
        var iw = im.naturalWidth, ih = im.naturalHeight;
        var br = Math.min(box.w / iw, box.h / ih);
        var w = iw * br, h = ih * br;
        var x = box.x + (box.w - w) / 2;
        var y = box.y + (box.h - h) / 2;
        slide.addImage({ data: dataURL, x: x, y: y, w: w, h: h });
        resolve();
      };
      im.onerror = function () { reject(new Error('图片 dataURL 校验失败')); };
      im.src = dataURL;
    });
  }

  async function renderPanelSlide(pptx, card, idx, total) {
    var slide = pptx.addSlide();
    slide.background = { color: COL.cream };
    slide.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: 0.085, fill: { color: COL.accent }, line: { type: 'none' } });
    if (card && card.overview) { return await renderOverviewSlide(slide, card, idx, total); }
    // kicker
    slide.addText(segRuns('分镜漫画 · 逐格放大 · ' + (idx) + ' / ' + total, 15, COL.kicker, false), {
      x: TEXT_LEFT, y: 0.3, w: TEXT_W, h: 0.4, margin: 0
    });
    var titleText = (card.tag ? card.tag + ' · ' : '') + card.label;
    // 标题长则缩字号
    var tf = 32;
    if (titleText.length > 26) tf = 28;
    slide.addText(segRuns(titleText, tf, COL.title, false), {
      x: TEXT_LEFT, y: 0.66, w: TEXT_W, h: 0.72, margin: 0, valign: 'middle'
    });
    var note = card.note || '';
    var noteH = 0;
    if (note) {
      var noteLines = Math.max(1, Math.ceil((note.length * 1.0) / 30));
      noteH = noteLines * lineHeightIn(17) + 0.12;
    }
    var box = { x: 0.95, y: 1.62, w: SLIDE_W - 1.9, h: 5.05 - noteH };
    // 图片：不依赖 dataURL 加载失败
    if (card.dataURL) {
      try {
        await addImageContain(slide, card.dataURL, box);
      } catch (e) {
        slide.addShape('rect', { x: box.x, y: box.y, w: box.w, h: box.h, fill: { color: 'EFE2CE' }, line: { type: 'none' } });
        slide.addText(segRuns('（图片加载失败：' + card.src + '）', 15, COL.red, false), {
          x: box.x, y: box.y, w: box.w, h: 0.8, margin: 0
        });
      }
    } else {
      slide.addShape('rect', { x: box.x, y: box.y, w: box.w, h: box.h, fill: { color: 'EFE2CE' }, line: { type: 'none' } });
      slide.addText(segRuns('（当前 file:// 下图片不可读取：' + (card.src || '') + '\n请用本地服务器打开后重试，文字内容不受影响）', 15, COL.grey, false), {
        x: box.x + 0.2, y: box.y + 0.2, w: box.w - 0.4, h: 1.2, margin: 0
      });
    }
    if (note) {
      var nruns = segRuns(note, 17, COL.body, false);
      slide.addText(nruns, {
        x: TEXT_LEFT, y: 7.5 - noteH - 0.22, w: TEXT_W, h: noteH,
        align: 'left', valign: 'top', margin: 0
      });
    }
    addFooter(slide, idx, total);
    return slide;
  }

  /* 3x5 速查总览页：标题 + 15 张等比缩略图 + 标签 */
  async function renderOverviewSlide(slide, card, idx, total) {
    slide.addText(segRuns('分镜漫画 · 15 格速查', 32, COL.title, false), {
      x: TEXT_LEFT, y: 0.32, w: TEXT_W, h: 0.66, margin: 0, valign: 'middle'
    });
    slide.addText(segRuns('Overview · ' + card.cards.length + ' frames', 15, COL.kicker, false), {
      x: TEXT_LEFT, y: 1.0, w: TEXT_W, h: 0.4, margin: 0
    });
    var cols = 5, rows = Math.ceil(card.cards.length / cols);
    var areaX = 0.7, areaY = 1.55, areaW = SLIDE_W - 1.4, areaH = 7.0 - 1.55;
    var cw = areaW / cols, ch = areaH / rows;
    var labelH = 0.32, imgInset = 0.06;
    for (var i = 0; i < card.cards.length; i++) {
      var cc = card.cards[i];
      var cx = areaX + (i % cols) * cw, cy = areaY + Math.floor(i / cols) * ch;
      if (cc.dataURL) {
        var box = { x: cx + imgInset, y: cy + 0.05, w: cw - imgInset * 2, h: ch - labelH - 0.14 };
        await addImageContain(slide, cc.dataURL, box);
      } else {
        slide.addShape('rect', { x: cx + imgInset, y: cy + 0.05, w: cw - imgInset * 2, h: ch - labelH - 0.14, fill: { color: 'EFE2CE' }, line: { type: 'none' } });
      }
      var tagTxt = (cc.tag || ('PANEL ' + (i + 1)));
      slide.addText(segRuns(tagTxt, 12, COL.grey, false), {
        x: cx, y: cy + ch - labelH + 0.04, w: cw, h: labelH - 0.04,
        align: 'center', valign: 'middle', margin: 0
      });
    }
    addFooter(slide, idx, total);
    return slide;
  }

  /* ---------------- 入口 ---------------- */
  function sectionTitle(s) {
    return (s.getAttribute('data-title') || '').trim();
  }
  var DYNAMIC_TITLE_RE = /^(历史目录|板块·|生成故事板|Panel \d|角色与设定|四段式|逐段讲评|范文参考|任务总览|PDF 板块)/;
  function isDynamicTitle(t) { return DYNAMIC_TITLE_RE.test(t); }
  function isPanelSection(t) { return /^分镜漫画 · PANEL/.test(t); }
  function isOverviewSection(t) { return /15格速查|速查/.test(t); }

  function buildPlan() {
    var sections = Array.prototype.slice.call(document.querySelectorAll('.deck > section.slide[data-title]'));
    var plan = { textSections: [], panelCards: [] };
    for (var i = 0; i < sections.length; i++) {
      var t = sectionTitle(sections[i]);
      if (!t || isDynamicTitle(t)) continue;
      if (isOverviewSection(t)) {
        // “15格速查”总览页导出为一张 3x5 缩略图速查页
        var ovCards = panelOfSection(sections[i]);
        if (ovCards.length) plan.panelCards.push({ overview: true, title: t, cards: ovCards });
        continue;
      }
      if (isPanelSection(t)) {
        var cards = panelOfSection(sections[i]);
        plan.panelCards = plan.panelCards.concat(cards);
        continue;
      }
      plan.textSections.push(parseTextSection(sections[i]));
    }
    return plan;
  }

  function statusBox() {
    var box = document.getElementById('hc-export-status');
    if (box) return box;
    box = document.createElement('div');
    box.id = 'hc-export-status';
    box.style.cssText = 'position:fixed;right:18px;top:132px;z-index:2147483002;max-width:300px;background:rgba(255,252,245,.96);border:1px solid #e6c9a6;border-radius:12px;padding:10px 14px;font:13px/1.5 -apple-system,"PingFang SC","Microsoft YaHei",sans-serif;color:#4a2c10;box-shadow:0 6px 18px rgba(90,50,10,.18);white-space:pre-wrap;word-break:break-word;';
    document.body.appendChild(box);
    return box;
  }
  function setStatus(msg) {
    var b = statusBox();
    b.textContent = msg;
  }

  function createButton() {
    if (document.getElementById('hc-export-btn')) return;
    var b = document.createElement('button');
    b.id = 'hc-export-btn';
    b.textContent = '导出 PPT';
    b.title = '把本页所有讲评 slide 导出为 .pptx（真实文本框）';
    b.style.cssText = 'position:fixed;right:18px;top:82px;z-index:2147483001;font:600 15px/1 -apple-system,"PingFang SC","Microsoft YaHei",sans-serif;color:#fff;background:linear-gradient(135deg,#d99a5b,#b06a2c);border:0;border-radius:14px;padding:13px 18px;cursor:pointer;box-shadow:0 5px 16px rgba(160,90,20,.35);';
    b.addEventListener('mouseenter', function () { b.style.filter = 'brightness(1.06)'; });
    b.addEventListener('mouseleave', function () { b.style.filter = ''; });
    b.addEventListener('click', onExportClick);
    document.body.appendChild(b);
  }

  function disableUI(busyMsg) {
    var b = document.getElementById('hc-export-btn');
    if (b) { b.disabled = true; b.style.opacity = '0.62'; }
    setStatus(busyMsg || '正在准备导出…');
  }
  function enableUI() {
    var b = document.getElementById('hc-export-btn');
    if (b) { b.disabled = false; b.style.opacity = '1'; }
  }

  async function loadImageData(src) {
    var abs = new URL(src, document.baseURI).href;
    var img = new Image();
    img.crossOrigin = 'anonymous';
    await new Promise(function (resolve, reject) {
      img.onload = resolve;
      img.onerror = function () { reject(new Error('无法加载图片：' + src)); };
      img.src = abs;
    });
    var cv = document.createElement('canvas');
    cv.width = img.naturalWidth;
    cv.height = img.naturalHeight;
    var ctx = cv.getContext('2d');
    ctx.drawImage(img, 0, 0);
    return cv.toDataURL('image/png');
  }

  var FILE_PROTO_HINT = '检测到当前页面以 file:// 协议打开，浏览器禁止脚本跨协议读取图片。\n若图片导出失败，请在 hayloft-comic 目录运行：\n  python3 -m http.server 8000\n然后访问 http://localhost:8000/ 再点“导出 PPT”。\n（文字内容仍会正常导出）';

  async function onExportClick() {
    var isFile = window.location && location.protocol === 'file:';
    if (isFile) setStatus('提示：当前为 file:// 打开方式，图片读取可能受限，文字仍会正常导出。\n如需完整图片，请用本地服务器打开（见导出完成后提示）。');
    disableUI('正在扫描讲评页面…');
    var plan;
    try { plan = buildPlan(); }
    catch (e) { enableUI(); setStatus('解析页面失败：' + e.message); return; }

    var totalCards = plan.panelCards.length;
    var totalTextPages = 0;
    var textPageGroups = plan.textSections.map(function (sec) {
      var pages = layoutPages(sec, 6);
      totalTextPages += pages.length;
      return pages;
    });
    var totalSlides = totalTextPages + totalCards;
    var label = '导出中：共 ' + totalTextPages + ' 页讲评 + ' + totalCards + ' 页漫画';

    // 预加载所有漫画图片（速查总览页含 15 张子图，单独计数）
    var failedImgs = [];
    var totalImgs = 0, imgLoaded = 0;
    for (var ii0 = 0; ii0 < plan.panelCards.length; ii0++) {
      if (plan.panelCards[ii0].overview) totalImgs += plan.panelCards[ii0].cards.length;
      else totalImgs += 1;
    }
    async function loadOne(owner, c) {
      imgLoaded++;
      setStatus('正在加载漫画图片 ' + imgLoaded + ' / ' + totalImgs + ' …');
      if (!c.src) return;
      try { c.dataURL = await loadImageData(c.src); }
      catch (e) { failedImgs.push(c.src); c.loadError = true; }
    }
    for (var i = 0; i < plan.panelCards.length; i++) {
      var card = plan.panelCards[i];
      if (card.overview) {
        for (var j0 = 0; j0 < card.cards.length; j0++) await loadOne(card, card.cards[j0]);
      } else {
        await loadOne(card, card);
      }
    }
    if (isFile) {
      // file:// 下即使 loadError 也继续，但给提示；若成功也不打扰
    }
    if (failedImgs.length && isFile) {
      setStatus(FILE_PROTO_HINT + '\n失败图片 ' + failedImgs.length + ' 张将留空占位。');
    } else if (failedImgs.length) {
      setStatus('有 ' + failedImgs.length + ' 张图片加载失败，将以占位框展示。\n' + failedImgs.slice(0, 5).join('\n'));
    }

    var pptx = new window.PptxGenJS();
    pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
    pptx.layout = 'WIDE';
    pptx.author = 'Marvis File Agent';
    pptx.company = '';
    pptx.title = 'hayloft-comic 讲评';

    var slideIdx = 0;
    // 文本页
    for (var g = 0; g < textPageGroups.length; g++) {
      var sec = plan.textSections[g];
      var pages = textPageGroups[g];
      for (var p = 0; p < pages.length; p++) {
        slideIdx++;
        setStatus(label + '\n正在排布文字页 ' + slideIdx + ' / ' + totalSlides + ' …');
        renderTextSlide(pptx, sec, pages[p].items, slideIdx, totalSlides);
      }
    }
    // 图片页
    for (var k = 0; k < totalCards; k++) {
      slideIdx++;
      setStatus(label + '\n正在写入漫画页 ' + slideIdx + ' / ' + totalSlides + ' …');
      try { await renderPanelSlide(pptx, plan.panelCards[k], slideIdx, totalSlides); }
      catch (e) { /* 单页失败不影响整体 */ }
    }

    setStatus('正在生成 .pptx 并触发下载…');
    try {
      await pptx.writeFile({ fileName: 'hayloft-comic讲评.pptx' });
      var doneMsg = '已导出 ' + totalSlides + ' 页：hayloft-comic讲评.pptx';
      if (isFile) doneMsg += '\n（file:// 下图片加载受限；需要完整图片请用 python3 -m http.server 8000 启动本地服务器后重试）';
      if (failedImgs.length) doneMsg += '\n注意：' + failedImgs.length + ' 张图片未能加载（占位处理）。';
      setStatus(doneMsg);
    } catch (e) {
      setStatus('生成 PPT 失败：' + e.message + (isFile ? '\n' + FILE_PROTO_HINT : ''));
    }
    enableUI();
  }

  /* ---------------- 自检 ---------------- */
  function selfCheck() {
    var okLib = !!(window.PptxGenJS);
    var sections = document.querySelectorAll('.deck > section.slide[data-title]');
    var titles = [];
    for (var i = 0; i < sections.length; i++) titles.push(sectionTitle(sections[i]));
    return { okLib: okLib, slideCount: sections.length, titles: titles };
  }

  function init() {
    createButton();
    var r = selfCheck();
    console.log('[export-pptx] PptxGenJS:', r.okLib ? '已加载' : '缺失', '| 静态 slide 数:', r.slideCount);
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
  } // --- end main() ---

  // 本地 assets/vendor 已内置；若本地加载异常则尝试 CDN 兜底一次
  if (window.PptxGenJS) { main(); return; }
  if (window.__hcPptxCdnTried) { console.warn('[export-pptx] window.PptxGenJS 未加载'); return; }
  window.__hcPptxCdnTried = true;
  var __sc = document.createElement('script');
  __sc.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3/dist/pptxgen.bundle.js';
  __sc.onload = function () { main(); };
  __sc.onerror = function () { console.warn('[export-pptx] 本地 vendor 与 CDN 兜底均加载失败'); };
  (document.head || document.documentElement).appendChild(__sc);
})();
