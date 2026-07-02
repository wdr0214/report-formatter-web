/* global JSZip */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const DEFAULT_TEMPLATE_NAME = "汇报材料模板.docx";
const DEFAULT_TEMPLATE_URL = `./${encodeURIComponent(DEFAULT_TEMPLATE_NAME)}`;

const state = {
  templateFile: null,
  selectedFiles: [],
  results: [],
};

const el = {
  chooseTemplate: document.getElementById("chooseTemplate"),
  templateInput: document.getElementById("templateInput"),
  templateName: document.getElementById("templateName"),
  chooseFiles: document.getElementById("chooseFiles"),
  chooseFolder: document.getElementById("chooseFolder"),
  fileInput: document.getElementById("fileInput"),
  folderInput: document.getElementById("folderInput"),
  fileRows: document.getElementById("fileRows"),
  formatButton: document.getElementById("formatButton"),
  previewButton: document.getElementById("previewButton"),
  downloadButton: document.getElementById("downloadButton"),
  previewBox: document.getElementById("previewBox"),
  runtimeStatus: document.getElementById("runtimeStatus"),
};

const rulesDefault = () => ({
  mainTitle: { font: "宋体", size: "44", bold: true },
  unitTitle: { font: "楷体_GB2312", size: "36", bold: false },
  level1: { font: "黑体", size: "32", bold: true },
  level2: { font: "楷体_GB2312", size: "32", bold: false },
  level3: { font: "仿宋_GB2312", size: "32", bold: true },
  level4: { font: "仿宋_GB2312", size: "32", bold: false },
  body: { font: "仿宋_GB2312", size: "32", bold: false },
  firstLine: "640",
  line: "600",
});

const regex = {
  level1: /^[一二三四五六七八九十]+、/,
  level2: /^（[一二三四五六七八九十]+）/,
  level3: /^\d+[\.\．]/,
  level4: /^（\d+）/,
  level3Space: /^(\s*\d+[\.\．])\s+/,
  caption: /^[图表](\s*\d+([\-\.－—]\d+)*|\s+\S+)/,
  chapterTitle: /^第[一二三四五六七八九十百零〇]+章(\s|　)*\S*/,
  header: /(总经理办公会|总办会|党委会|董事会|战略与投资委员会|汇报材料|会议文件)/,
  repeatedPunctuation: /([，。！？；：,.!?;:])\1+/g,
  multiSpace: / {2,}/g,
  chineseSpace: /([\u4e00-\u9fff]) +([\u4e00-\u9fff])/g,
};

el.chooseTemplate.addEventListener("click", () => el.templateInput.click());
el.chooseFiles.addEventListener("click", () => el.fileInput.click());
el.chooseFolder.addEventListener("click", () => el.folderInput.click());
el.formatButton.addEventListener("click", processSelectedFiles);
el.previewButton.addEventListener("click", showPreview);
el.downloadButton.addEventListener("click", downloadResults);

el.templateInput.addEventListener("change", () => {
  state.templateFile = el.templateInput.files[0] || null;
  el.templateName.textContent = state.templateFile ? `已更换模板：${state.templateFile.name}` : `默认模板：${DEFAULT_TEMPLATE_NAME}`;
});

el.fileInput.addEventListener("change", () => {
  addFiles(Array.from(el.fileInput.files || []));
  el.fileInput.value = "";
});

el.folderInput.addEventListener("change", () => {
  addFiles(Array.from(el.folderInput.files || []));
  el.folderInput.value = "";
});

function addFiles(files) {
  const known = new Set(state.selectedFiles.map((item) => item.pathKey));
  for (const file of files) {
    const pathKey = file.webkitRelativePath || file.name;
    if (!known.has(pathKey) && /\.(docx|doc|wps)$/i.test(file.name)) {
      state.selectedFiles.push({ file, pathKey, status: "待处理" });
      known.add(pathKey);
    }
  }
  state.results = [];
  renderFileRows();
}

function renderFileRows() {
  el.fileRows.innerHTML = "";
  state.selectedFiles.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td></td><td></td><td></td><td></td>`;
    tr.children[0].textContent = item.file.name;
    tr.children[1].textContent = item.pathKey;
    tr.children[2].textContent = item.status;
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "small danger";
    removeButton.textContent = "删除";
    removeButton.addEventListener("click", () => {
      state.selectedFiles.splice(index, 1);
      state.results = [];
      renderFileRows();
      el.previewBox.textContent = state.selectedFiles.length ? "文件列表已更新，请重新点击“调整格式”。" : "已清空文件列表。";
    });
    tr.children[3].appendChild(removeButton);
    el.fileRows.appendChild(tr);
  });
}

async function processSelectedFiles() {
  if (!window.JSZip) {
    alert("未加载 JSZip，无法处理 docx。请检查网络或将 JSZip 文件放到本项目中。");
    return;
  }
  if (!state.selectedFiles.length) {
    alert("请先选择待处理文件。");
    return;
  }

  setBusy(true);
  state.results = [];
  try {
    const rules = await loadSelectedTemplateRules();
    for (const item of state.selectedFiles) {
      item.status = "处理中";
      renderFileRows();
      const result = await processOneFile(item.file, rules);
      state.results.push(result);
      item.status = result.success ? "成功" : `不支持：请转为 docx`;
      renderFileRows();
    }
    showPreview();
  } catch (error) {
    alert(`处理失败：${error.message || error}`);
  } finally {
    setBusy(false);
  }
}

async function loadSelectedTemplateRules() {
  if (state.templateFile) return loadTemplateRules(state.templateFile.name, await state.templateFile.arrayBuffer());
  try {
    const response = await fetch(DEFAULT_TEMPLATE_URL, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return loadTemplateRules(DEFAULT_TEMPLATE_NAME, await response.arrayBuffer());
  } catch (error) {
    throw new Error(`未能读取同目录下的默认模板“${DEFAULT_TEMPLATE_NAME}”。请用“更换模板”手动选择模板，或通过 start-local-server.py 启动后再打开页面。`);
  }
}

function setBusy(busy) {
  el.formatButton.disabled = busy;
  el.previewButton.disabled = busy;
  el.downloadButton.disabled = busy;
  el.runtimeStatus.textContent = busy ? "正在本机浏览器内处理..." : "文件仅在本机浏览器内处理，不上传服务器。";
}

async function processOneFile(file, rules) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "doc" || ext === "wps") {
    return {
      file,
      success: false,
      message: "不支持直接处理该格式，请先用 WPS/Word 另存为 docx。",
      previewText: "",
    };
  }
  if (ext !== "docx") {
    return { file, success: false, message: "不支持的文件格式。", previewText: "" };
  }

  try {
    const input = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(input);
    const documentEntry = zip.file("word/document.xml");
    if (!documentEntry) throw new Error("不是有效的 docx 文件。");
    const documentXml = await documentEntry.async("string");
    const parser = new DOMParser();
    const doc = parser.parseFromString(documentXml, "application/xml");
    const body = firstByTag(doc, "body");
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    const tocRanges = findTocRanges(paragraphs);

    let nonEmptyIndex = 0;
    let foundHeader = false;
    let mainTitleApplied = false;
    let linesAfterMainTitle = -1;
    const stats = { spaces: 0, punctuation: 0, symbols: 0 };

    paragraphs.forEach((p, index) => {
      const text = getParagraphText(p);
      if (!text.trim()) return;
      if (isInRanges(index, tocRanges) || isFigureOrTableCaption(p, text)) return;

      nonEmptyIndex += 1;
      if (!mainTitleApplied && isFrontHeader(text, nonEmptyIndex)) {
        foundHeader = true;
        return;
      }

      const isMain = !mainTitleApplied && isMainTitle(p, text, nonEmptyIndex, foundHeader);
      if (isMain) {
        mainTitleApplied = true;
        linesAfterMainTitle = 0;
      }

      let onlyParentheses = false;
      if (mainTitleApplied && !isMain && linesAfterMainTitle >= 0 && linesAfterMainTitle < 3) {
        linesAfterMainTitle += 1;
        onlyParentheses = isCentered(p);
      }

      const fixed = fixParagraphText(text, p, onlyParentheses);
      stats.spaces += fixed.spaces;
      stats.punctuation += fixed.punctuation;
      stats.symbols += fixed.symbols;
      if (fixed.text !== text) setParagraphTextPreserveRuns(p, fixed.text);
    });

    nonEmptyIndex = 0;
    foundHeader = false;
    mainTitleApplied = false;
    linesAfterMainTitle = -1;
    paragraphs.forEach((p, index) => {
      const text = getParagraphText(p);
      if (!text.trim()) return;
      nonEmptyIndex += 1;
      if (!mainTitleApplied && isFrontHeader(text, nonEmptyIndex)) {
        foundHeader = true;
        return;
      }
      if (isInRanges(index, tocRanges) || isFigureOrTableCaption(p, text) || isCenteredChapterTitle(p, text) || isInTable(p)) {
        return;
      }
      if (mainTitleApplied && linesAfterMainTitle >= 0 && linesAfterMainTitle < 3) {
        linesAfterMainTitle += 1;
        if (isCentered(p)) return;
      }

      const kind = detectKind(p, text, nonEmptyIndex, foundHeader, mainTitleApplied);
      applyParagraphFormat(p, kind, rules);
      if (kind === "mainTitle") {
        mainTitleApplied = true;
        linesAfterMainTitle = 0;
      }
    });

    await ensureTrackRevisions(zip);
    const serializer = new XMLSerializer();
    zip.file("word/document.xml", serializer.serializeToString(doc));
    const blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    return {
      file,
      success: true,
      blob,
      outputName: makeOutputName(file.name),
      stats,
      previewText: paragraphs.map(getParagraphText).filter(Boolean).slice(0, 80).join("\n"),
      message: "处理成功；浏览器版可能无法生成完整 Office 可视修订痕迹，部分修改为直接应用。",
    };
  } catch (error) {
    return { file, success: false, message: error.message || String(error), previewText: "" };
  }
}

async function loadTemplateRules(name, arrayBuffer) {
  if (!/\.docx$/i.test(name)) return rulesDefault();
  const rules = rulesDefault();
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const entry = zip.file("word/document.xml");
    if (!entry) return rules;
    const xml = await entry.async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of paragraphs) {
      const text = getParagraphText(p).trim();
      if (!text) continue;
      const style = readFirstRunStyle(p) || rules.body;
      const pPr = child(p, "pPr");
      if (pPr) {
        const ind = child(pPr, "ind");
        const spacing = child(pPr, "spacing");
        if (ind?.getAttributeNS(W_NS, "firstLine")) rules.firstLine = ind.getAttributeNS(W_NS, "firstLine");
        if (spacing?.getAttributeNS(W_NS, "line")) rules.line = spacing.getAttributeNS(W_NS, "line");
      }
      if (text.includes("关于") && text.includes("汇报")) rules.mainTitle = style;
      else if (text.includes("单位")) rules.unitTitle = style;
      else if (regex.level1.test(text)) rules.level1 = style;
      else if (regex.level2.test(text)) rules.level2 = style;
      else if (regex.level3.test(text)) rules.level3 = style;
      else if (regex.level4.test(text)) rules.level4 = { ...style, bold: false };
      else if (text.includes("正文")) rules.body = { ...style, bold: false };
    }
    rules.level4 = { ...rules.level4, bold: false };
  } catch {
    return rules;
  }
  return rules;
}

function fixParagraphText(text, p, onlyParentheses) {
  let result = text;
  const stats = { spaces: 0, punctuation: 0, symbols: 0, text };
  if (onlyParentheses) {
    const normalized = normalizeParentheses(result);
    stats.symbols += normalized.count;
    result = normalized.text;
    stats.text = result;
    return stats;
  }

  if (!isCenteredChapterTitle(p, text)) {
    stats.spaces += countMatches(result, regex.multiSpace);
    stats.spaces += countMatches(result, regex.chineseSpace);
    stats.spaces += countMatches(result, regex.level3Space);
    result = result.replace(regex.level3Space, "$1");
    result = result.replace(regex.chineseSpace, "$1$2");
    result = result.replace(regex.multiSpace, " ");
  }
  stats.punctuation += countMatches(result, regex.repeatedPunctuation);
  result = result.replace(regex.repeatedPunctuation, "$1");
  const normalized = normalizeSymbols(result);
  stats.symbols += normalized.count;
  result = normalized.text;
  stats.text = result;
  return stats;
}

function countMatches(text, pattern) {
  const flags = pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`;
  const copy = new RegExp(pattern.source, flags);
  return Array.from(text.matchAll(copy)).length;
}

function normalizeSymbols(text) {
  const chars = Array.from(text);
  const quoteIndexes = [];
  let count = 0;
  chars.forEach((ch, index) => {
    if (ch === "(") {
      chars[index] = "（";
      count += 1;
    } else if (ch === ")") {
      chars[index] = "）";
      count += 1;
    } else if (ch === "\"") {
      quoteIndexes.push(index);
    }
  });
  const paired = quoteIndexes.length - (quoteIndexes.length % 2);
  for (let i = 0; i < paired; i += 2) {
    chars[quoteIndexes[i]] = "“";
    chars[quoteIndexes[i + 1]] = "”";
    count += 2;
  }
  return { text: chars.join(""), count };
}

function normalizeParentheses(text) {
  const chars = Array.from(text);
  let count = 0;
  chars.forEach((ch, index) => {
    if (ch === "(") {
      chars[index] = "（";
      count += 1;
    } else if (ch === ")") {
      chars[index] = "）";
      count += 1;
    }
  });
  return { text: chars.join(""), count };
}

function detectKind(p, text, index, foundHeader, mainTitleApplied) {
  const stripped = text.trimStart();
  if (!mainTitleApplied && isMainTitle(p, text, index, foundHeader)) return "mainTitle";
  if (regex.level1.test(stripped)) return "level1";
  if (regex.level2.test(stripped)) return "level2";
  if (regex.level3.test(stripped)) return isInlineTitle(stripped) ? "level3Inline" : "level3";
  if (regex.level4.test(stripped)) return isInlineTitle(stripped) ? "level4Inline" : "level4";
  return "body";
}

function applyParagraphFormat(p, kind, rules) {
  const map = {
    mainTitle: [rules.mainTitle, "center", "0", false],
    level1: [rules.level1, "both", rules.firstLine, false],
    level2: [rules.level2, "both", rules.firstLine, false],
    level3: [rules.level3, "both", rules.firstLine, false],
    level4: [rules.level4, "both", rules.firstLine, false],
    body: [rules.body, "both", rules.firstLine, true],
  };
  if (kind === "level3Inline") return applyInlineParagraph(p, rules.body, rules.level3, rules);
  if (kind === "level4Inline") return applyInlineParagraph(p, rules.body, rules.level4, rules);
  const [style, align, firstLine, keepBold] = map[kind] || map.body;
  applyWholeParagraph(p, style, align, firstLine, rules.line, keepBold);
}

function applyWholeParagraph(p, style, align, firstLine, line, keepBold) {
  applyParagraphProps(p, align, firstLine, line);
  Array.from(p.getElementsByTagNameNS(W_NS, "r")).forEach((run) => applyRunStyle(run, style, keepBold));
}

function applyInlineParagraph(p, bodyStyle, titleStyle, rules) {
  applyParagraphProps(p, "both", rules.firstLine, rules.line);
  const text = getParagraphText(p);
  const end = findInlineTitleEnd(text.trimStart());
  const leading = text.length - text.trimStart().length;
  setParagraphSingleRun(p, text);
  const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
  runs.forEach((run) => applyRunStyle(run, bodyStyle, true));
  if (end < 0 || !runs.length) return;
  splitInlineTitle(p, leading + end + 1, titleStyle, bodyStyle);
}

function splitInlineTitle(p, titleLength, titleStyle, bodyStyle) {
  const text = getParagraphText(p);
  if (titleLength <= 0 || titleLength >= text.length) return;
  Array.from(p.getElementsByTagNameNS(W_NS, "r")).forEach((run) => run.parentNode.removeChild(run));
  const titleRun = p.ownerDocument.createElementNS(W_NS, "w:r");
  const titleText = p.ownerDocument.createElementNS(W_NS, "w:t");
  titleText.textContent = text.slice(0, titleLength);
  titleRun.appendChild(titleText);
  p.appendChild(titleRun);
  applyRunStyle(titleRun, titleStyle, false);

  const bodyRun = p.ownerDocument.createElementNS(W_NS, "w:r");
  const bodyText = p.ownerDocument.createElementNS(W_NS, "w:t");
  bodyText.textContent = text.slice(titleLength);
  bodyRun.appendChild(bodyText);
  p.appendChild(bodyRun);
  applyRunStyle(bodyRun, bodyStyle, true);
}

function applyParagraphProps(p, align, firstLine, line) {
  const pPr = ensureChild(p, "pPr", true);
  const jc = ensureChild(pPr, "jc");
  jc.setAttributeNS(W_NS, "w:val", align);
  const ind = ensureChild(pPr, "ind");
  ind.setAttributeNS(W_NS, "w:left", "0");
  if (firstLine === "0") ind.removeAttributeNS(W_NS, "firstLine");
  else ind.setAttributeNS(W_NS, "w:firstLine", firstLine);
  const spacing = ensureChild(pPr, "spacing");
  spacing.setAttributeNS(W_NS, "w:before", "0");
  spacing.setAttributeNS(W_NS, "w:after", "0");
  spacing.setAttributeNS(W_NS, "w:line", line);
  spacing.setAttributeNS(W_NS, "w:lineRule", "exact");
}

function applyRunStyle(run, style, keepBold) {
  const rPr = ensureChild(run, "rPr", true);
  const fonts = ensureChild(rPr, "rFonts");
  fonts.setAttributeNS(W_NS, "w:eastAsia", style.font);
  fonts.setAttributeNS(W_NS, "w:ascii", style.font);
  fonts.setAttributeNS(W_NS, "w:hAnsi", style.font);
  const sz = ensureChild(rPr, "sz");
  sz.setAttributeNS(W_NS, "w:val", style.size);
  const szCs = ensureChild(rPr, "szCs");
  szCs.setAttributeNS(W_NS, "w:val", style.size);
  if (!keepBold || style.bold) {
    let b = child(rPr, "b");
    if (style.bold) {
      if (!b) b = ensureChild(rPr, "b");
      b.removeAttributeNS(W_NS, "val");
    } else if (b) {
      rPr.removeChild(b);
    }
    const bCs = child(rPr, "bCs");
    if (!style.bold && bCs) rPr.removeChild(bCs);
  }
}

function getParagraphText(p) {
  return Array.from(p.getElementsByTagNameNS(W_NS, "t")).map((node) => node.textContent || "").join("");
}

function setParagraphTextPreserveRuns(p, text) {
  const textNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));
  if (!textNodes.length) return setParagraphSingleRun(p, text);
  const lengths = textNodes.map((node) => (node.textContent || "").length);
  let position = 0;
  textNodes.forEach((node, index) => {
    let piece;
    if (index === textNodes.length - 1) piece = text.slice(position);
    else {
      piece = text.slice(position, position + lengths[index]);
      position += lengths[index];
    }
    node.textContent = piece;
    preserveSpace(node, piece);
  });
}

function setParagraphSingleRun(p, text) {
  let runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
  if (!runs.length) {
    const run = p.ownerDocument.createElementNS(W_NS, "w:r");
    p.appendChild(run);
    runs = [run];
  }
  let firstText = null;
  runs.forEach((run) => {
    Array.from(run.getElementsByTagNameNS(W_NS, "t")).forEach((t) => {
      if (!firstText) firstText = t;
      else t.textContent = "";
    });
  });
  if (!firstText) {
    firstText = p.ownerDocument.createElementNS(W_NS, "w:t");
    runs[0].appendChild(firstText);
  }
  firstText.textContent = text;
  preserveSpace(firstText, text);
}

function preserveSpace(node, text) {
  if (text.startsWith(" ") || text.endsWith(" ") || text.includes("  ")) node.setAttributeNS(XML_NS, "xml:space", "preserve");
  else node.removeAttributeNS(XML_NS, "space");
}

function readFirstRunStyle(p) {
  const run = child(p, "r");
  const rPr = run ? child(run, "rPr") : null;
  if (!rPr) return null;
  const fonts = child(rPr, "rFonts");
  const sz = child(rPr, "sz");
  return {
    font: fonts?.getAttributeNS(W_NS, "eastAsia") || fonts?.getAttributeNS(W_NS, "ascii") || "仿宋_GB2312",
    size: sz?.getAttributeNS(W_NS, "val") || "32",
    bold: Boolean(child(rPr, "b")),
  };
}

function isFigureOrTableCaption(p, text) {
  return isCentered(p) && regex.caption.test(text.trimStart());
}

function isCenteredChapterTitle(p, text) {
  return isCentered(p) && regex.chapterTitle.test(text.trim());
}

function isFrontHeader(text, index) {
  return index <= 4 && regex.header.test(text);
}

function isMainTitle(p, text, index, foundHeader) {
  const stripped = text.trimStart();
  if (foundHeader && isCentered(p) && stripped.length <= 80) return true;
  return !foundHeader && index <= 2 && isCentered(p) && stripped.length <= 40;
}

function isCentered(p) {
  const jc = p.getElementsByTagNameNS(W_NS, "jc")[0];
  return jc?.getAttributeNS(W_NS, "val") === "center";
}

function isInTable(p) {
  let node = p.parentNode;
  while (node) {
    if (node.localName === "tbl" && node.namespaceURI === W_NS) return true;
    node = node.parentNode;
  }
  return false;
}

function findTocRanges(paragraphs) {
  const ranges = [];
  let inToc = false;
  let start = -1;
  paragraphs.forEach((p, index) => {
    const instr = Array.from(p.getElementsByTagNameNS(W_NS, "instrText")).map((node) => node.textContent || "").join("");
    const fldTypes = Array.from(p.getElementsByTagNameNS(W_NS, "fldChar")).map((node) => node.getAttributeNS(W_NS, "fldCharType"));
    if (instr.includes("TOC")) {
      inToc = true;
      if (start < 0) start = index;
    }
    if (inToc && fldTypes.includes("end")) {
      ranges.push([start, index]);
      inToc = false;
      start = -1;
    }
  });
  if (inToc && start >= 0) ranges.push([start, Math.min(paragraphs.length - 1, start + 80)]);
  return ranges;
}

function isInRanges(index, ranges) {
  return ranges.some(([start, end]) => index >= start && index <= end);
}

function isInlineTitle(text) {
  const end = findInlineTitleEnd(text);
  return end > 0 && end < text.length - 1;
}

function findInlineTitleEnd(text) {
  const indexes = [text.indexOf("。"), text.indexOf("："), text.indexOf(":")].filter((item) => item >= 0);
  return indexes.length ? Math.min(...indexes) : -1;
}

async function ensureTrackRevisions(zip) {
  const entry = zip.file("word/settings.xml");
  if (!entry) return;
  const xml = await entry.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (!doc.getElementsByTagNameNS(W_NS, "trackRevisions").length) {
    doc.documentElement.appendChild(doc.createElementNS(W_NS, "w:trackRevisions"));
    zip.file("word/settings.xml", new XMLSerializer().serializeToString(doc));
  }
}

function child(parent, localName) {
  return Array.from(parent.childNodes).find((node) => node.nodeType === 1 && node.localName === localName && node.namespaceURI === W_NS) || null;
}

function firstByTag(doc, localName) {
  return doc.getElementsByTagNameNS(W_NS, localName)[0] || null;
}

function ensureChild(parent, localName, insertFirst = false) {
  let existing = child(parent, localName);
  if (existing) return existing;
  const node = parent.ownerDocument.createElementNS(W_NS, `w:${localName}`);
  if (insertFirst && parent.firstChild) parent.insertBefore(node, parent.firstChild);
  else parent.appendChild(node);
  return node;
}

function makeOutputName(name) {
  return name.replace(/\.docx$/i, "_格式整理.docx");
}

function showPreview() {
  if (!state.results.length) {
    el.previewBox.textContent = "还没有处理结果。";
    return;
  }
  el.previewBox.textContent = state.results.map((result) => {
    const lines = [
      `文件：${result.file.name}`,
      `状态：${result.success ? "成功" : "失败"}`,
      `说明：${result.message}`,
    ];
    if (result.stats) {
      lines.push(`多余空格修订：${result.stats.spaces}`);
      lines.push(`重复标点修订：${result.stats.punctuation}`);
      lines.push(`英文符号转中文修订：${result.stats.symbols}`);
    }
    if (result.previewText) {
      lines.push("");
      lines.push(result.previewText.slice(0, 2000));
    }
    return lines.join("\n");
  }).join("\n\n------------------------------\n\n");
}

async function downloadResults() {
  const successful = state.results.filter((result) => result.success && result.blob);
  if (!successful.length) {
    alert("没有可下载的处理结果。");
    return;
  }
  if (successful.length === 1) {
    triggerDownload(successful[0].blob, successful[0].outputName);
    return;
  }
  const zip = new JSZip();
  successful.forEach((result) => zip.file(result.outputName, result.blob));
  const blob = await zip.generateAsync({ type: "blob" });
  triggerDownload(blob, "汇报材料格式整理结果.zip");
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
