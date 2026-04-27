(() => {
  "use strict";

  const STORAGE_KEY = "outliner.projects.v1";
  const LEGACY_KEYS = ["toc-builder.v2", "toc-builder.v1"];

  const root = document.getElementById("root");
  const paper = document.getElementById("paper");
  const maxLevelsInput = document.getElementById("maxLevels");
  const statusEl = document.getElementById("status");
  const tocTitle = document.getElementById("toc-title");
  const tocSub = document.getElementById("toc-sub");
  const sidebar = document.getElementById("sidebar");
  const sidebarToggle = document.getElementById("sidebarToggle");
  const projectListEl = document.getElementById("projectList");
  const newProjectBtn = document.getElementById("newProjectBtn");
  const activeProjectName = document.getElementById("activeProjectName");
  const projectSearch = document.getElementById("projectSearch");
  const sidebarClose = document.getElementById("sidebarClose");
  const sidebarBackdrop = document.getElementById("sidebarBackdrop");

  let store = { projects: [], currentId: null };
  let searchQuery = "";

  const placeholders = [
    "Chapter",
    "Section",
    "Subsection",
    "Topic",
    "Subtopic",
    "Point",
    "Sub-point",
    "Detail",
    "Note",
    "Reference",
  ];

  const maxLevels = () => {
    const v = parseInt(maxLevelsInput.value, 10);
    if (isNaN(v) || v < 1) return 1;
    if (v > 10) return 10;
    return v;
  };

  let draggedLi = null;

  // ---------- Node creation ----------
  function createNode(level, html = "", plain = false) {
    const li = document.createElement("li");
    li.className = "node";
    li.dataset.level = level;
    if (plain) li.dataset.plain = "true";

    const row = document.createElement("div");
    row.className = "row";

    // Drag handle
    const handle = document.createElement("div");
    handle.className = "drag-handle";
    handle.innerHTML = "&#x2807;&#x2807;"; // braille dots as grip
    handle.title = "Drag to move";
    handle.addEventListener("mousedown", () => { li.draggable = true; });
    handle.addEventListener("mouseup", () => { li.draggable = false; });

    // Toggle
    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "toggle hidden";
    toggle.textContent = "▼";
    toggle.title = "Expand / collapse";
    toggle.tabIndex = -1;
    toggle.addEventListener("click", () => toggleCollapse(li));

    // Number badge (outline numbering, color by level)
    const badge = document.createElement("span");
    badge.className = "level-badge level-" + level;
    badge.textContent = "";

    // Editable text
    const txt = document.createElement("div");
    txt.className = "text";
    txt.contentEditable = "true";
    txt.spellcheck = false;
    txt.dataset.placeholder = plain ? "Text" : (placeholders[level - 1] || "Level " + level);
    txt.innerHTML = html;
    updateEmptyClass(txt);
    txt.addEventListener("input", () => { updateEmptyClass(txt); scheduleSave(); });
    txt.addEventListener("focus", () => row.classList.add("focused"));
    txt.addEventListener("blur", () => row.classList.remove("focused"));
    txt.addEventListener("keydown", (e) => handleKey(e, li));
    txt.addEventListener("paste", (e) => {
      e.preventDefault();
      const text = (e.clipboardData || window.clipboardData).getData("text");
      document.execCommand("insertText", false, text);
    });

    // Actions: sibling first (+), then child, bold, delete.
    // Plain text cells skip the Add Item / Add Sub Item buttons since their
    // numbering semantics are intentionally different.
    const actions = document.createElement("div");
    actions.className = "actions";
    if (!plain) {
      actions.appendChild(makeActionBtn("btn-sibling", "+", "Add item (Ctrl+Enter)", () => addSibling(li)));
      actions.appendChild(makeActionBtn("btn-child", "↳", "Add sub-item", () => addChild(li)));
    }
    actions.appendChild(makeActionBtn("btn-bold", "B", "Bold (Ctrl+B)", () => { txt.focus(); document.execCommand("bold"); }));
    actions.appendChild(makeActionBtn("btn-bullet-dash", "-", "Dash bullet (does not nest into hollow)", () => insertListAtLineStart(txt, "dash")));
    actions.appendChild(makeActionBtn("btn-bullet-dark", "●", "Indented bullet (filled)", () => insertListAtLineStart(txt, "black")));
    actions.appendChild(makeActionBtn("btn-bullet-light", "○", "Indented bullet (hollow)", () => insertListAtLineStart(txt, "white")));
    actions.appendChild(makeActionBtn("btn-num1", "1.", "Numbered list (1., 2.)", () => {
      txt.focus();
      const ctx = getLineContext(txt);
      const token = ctx ? findNextNumber(txt, ctx, "num1") : "1";
      insertListAtLineStart(txt, "num1", token);
    }));
    actions.appendChild(makeActionBtn("btn-num2", "a.", "Numbered sub-list (a., b.)", () => {
      txt.focus();
      const ctx = getLineContext(txt);
      const token = ctx ? findNextNumber(txt, ctx, "num2") : "a";
      insertListAtLineStart(txt, "num2", token);
    }));
    actions.appendChild(makeActionBtn("btn-del", "×", "Delete (Ctrl+Del)", () => removeNode(li)));

    row.append(handle, toggle, badge, txt, actions);
    li.appendChild(row);

    wireDragAndDrop(li, row);
    refreshChildButton(li);
    return li;
  }

  function makeActionBtn(cls, label, title, handler) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = cls;
    b.innerHTML = label;
    b.title = title;
    b.tabIndex = -1;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", () => { handler(); scheduleSave(); });
    return b;
  }

  function updateEmptyClass(txt) {
    const hasContent = txt.textContent.trim().length > 0;
    txt.classList.toggle("is-empty", !hasContent);
  }

  // ---------- Tree operations ----------
  function addChild(li) {
    const level = parseInt(li.dataset.level, 10);
    if (level >= maxLevels()) return;
    let ul = li.querySelector(":scope > ul.tree");
    if (!ul) {
      ul = document.createElement("ul");
      ul.className = "tree";
      li.appendChild(ul);
    }
    const child = createNode(level + 1);
    ul.appendChild(child);
    li.classList.remove("collapsed");
    refreshToggle(li);
    renumber();
    focusNode(child);
    updateEmpty();
  }

  function addSibling(li) {
    const level = parseInt(li.dataset.level, 10);
    const sib = createNode(level);
    li.after(sib);
    renumber();
    focusNode(sib);
    updateEmpty();
  }

  function addSiblingAbove(li) {
    const level = parseInt(li.dataset.level, 10);
    const sib = createNode(level);
    li.before(sib);
    renumber();
    focusNode(sib);
    updateEmpty();
  }

  // 'b' adds a plain text cell directly below the hovered cell.
  //   - On a numbered cell: nest as the first child (one level deeper).
  //   - On a plain text cell: add as the next sibling at the same level
  //     (so chains of text cells stay aligned and don't drift right).
  // Falls back to a same-level sibling when already at the maximum depth.
  function addPlainSibling(li) {
    const level = parseInt(li.dataset.level, 10);
    const isPlain = li.dataset.plain === "true";
    if (isPlain || level >= maxLevels()) {
      const sib = createNode(level, "", true);
      li.after(sib);
      renumber();
      focusNode(sib);
      updateEmpty();
      return;
    }
    let childUl = li.querySelector(":scope > ul.tree");
    if (!childUl) {
      childUl = document.createElement("ul");
      childUl.className = "tree";
      li.appendChild(childUl);
    }
    const sib = createNode(level + 1, "", true);
    childUl.insertBefore(sib, childUl.firstChild);
    li.classList.remove("collapsed");
    refreshToggle(li);
    renumber();
    focusNode(sib);
    updateEmpty();
  }

  // 'a' adds a plain text cell as the previous sibling at the same level as
  // the hovered cell, so it stays horizontally aligned with the hovered cell.
  function addPlainSiblingAbove(li) {
    const level = parseInt(li.dataset.level, 10);
    const sib = createNode(level, "", true);
    li.before(sib);
    renumber();
    focusNode(sib);
    updateEmpty();
  }

  function removeNode(li) {
    const parentLi = li.parentElement.closest("li.node");
    const next = li.nextElementSibling || li.previousElementSibling || parentLi;
    li.remove();
    if (parentLi) refreshToggle(parentLi);
    renumber();
    if (next) focusNode(next);
    updateEmpty();
  }

  function toggleCollapse(li) {
    const ul = li.querySelector(":scope > ul.tree");
    if (!ul || !ul.children.length) return;
    li.classList.toggle("collapsed");
  }

  function refreshToggle(li) {
    const ul = li.querySelector(":scope > ul.tree");
    const toggle = li.querySelector(":scope > .row > .toggle");
    const hasChildren = ul && ul.children.length > 0;
    toggle.classList.toggle("hidden", !hasChildren);
    if (!hasChildren && ul) ul.remove();
  }

  function refreshChildButton(li) {
    const level = parseInt(li.dataset.level, 10);
    const btn = li.querySelector(":scope > .row > .actions > .btn-child");
    if (btn) btn.disabled = level >= maxLevels();
  }

  // Next li.node in document order: first visible child, else next sibling,
  // else first ancestor's next sibling. Returns null if this is the last cell.
  function getNextCell(li) {
    const childUl = li.querySelector(":scope > ul.tree");
    if (childUl && childUl.children.length && !li.classList.contains("collapsed")) {
      return childUl.children[0];
    }
    let cur = li;
    while (cur) {
      const sib = cur.nextElementSibling;
      if (sib && sib.matches("li.node")) return sib;
      const parentLi = cur.parentElement ? cur.parentElement.closest("li.node") : null;
      if (!parentLi) return null;
      cur = parentLi;
    }
    return null;
  }

  // Previous li.node in document order: previous sibling's deepest visible
  // descendant, else the parent li. Returns null if this is the first cell.
  function getPrevCell(li) {
    const sib = li.previousElementSibling;
    if (sib && sib.matches("li.node")) {
      let cur = sib;
      while (true) {
        const childUl = cur.querySelector(":scope > ul.tree");
        if (childUl && childUl.children.length && !cur.classList.contains("collapsed")) {
          cur = childUl.children[childUl.children.length - 1];
        } else break;
      }
      return cur;
    }
    return li.parentElement ? li.parentElement.closest("li.node") : null;
  }

  function focusNode(li) {
    const t = li.querySelector(":scope > .row > .text");
    if (!t) return;
    t.focus();
    const range = document.createRange();
    range.selectNodeContents(t);
    range.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function updateEmpty() {
    paper.classList.toggle("empty", root.children.length === 0);
  }

  // ---------- Outline numbering ----------
  function renumber() {
    function walk(ul, prefix) {
      let counter = 0;
      [...ul.children].forEach((li) => {
        const isPlain = li.dataset.plain === "true";
        const badge = li.querySelector(":scope > .row > .level-badge");
        const child = li.querySelector(":scope > ul.tree");
        if (isPlain) {
          // Badge is visibility:hidden via CSS, but its width still occupies
          // space. Set the text to the parent's prefix so the hidden badge
          // matches the parent's badge width and the plain cell's text lines
          // up with the parent's text after the -32px outdent.
          if (badge) badge.textContent = prefix;
          if (child) walk(child, prefix);
        } else {
          counter++;
          const num = prefix ? prefix + "." + counter : String(counter);
          if (badge) badge.textContent = num;
          if (child) walk(child, num);
        }
      });
    }
    walk(root, "");
  }

  // ---------- Relevel (after moving to a different depth) ----------
  function relevel(li, newLevel) {
    li.dataset.level = newLevel;
    const badge = li.querySelector(":scope > .row > .level-badge");
    if (badge) badge.className = "level-badge level-" + newLevel;
    const txt = li.querySelector(":scope > .row > .text");
    if (txt) txt.dataset.placeholder = placeholders[newLevel - 1] || "Level " + newLevel;
    refreshChildButton(li);
    const ul = li.querySelector(":scope > ul.tree");
    if (ul) [...ul.children].forEach((c) => relevel(c, newLevel + 1));
  }

  function subtreeDepth(li) {
    let max = 1;
    const ul = li.querySelector(":scope > ul.tree");
    if (ul) {
      [...ul.children].forEach((c) => { max = Math.max(max, 1 + subtreeDepth(c)); });
    }
    return max;
  }

  // ---------- Drag & drop ----------
  function wireDragAndDrop(li, row) {
    li.addEventListener("dragstart", (e) => {
      if (!li.draggable) return;
      draggedLi = li;
      e.dataTransfer.effectAllowed = "move";
      try { e.dataTransfer.setData("text/plain", "node"); } catch {}
      setTimeout(() => li.classList.add("dragging"), 0);
    });
    li.addEventListener("dragend", () => {
      li.draggable = false;
      li.classList.remove("dragging");
      clearDropIndicators();
      draggedLi = null;
    });

    row.addEventListener("dragover", (e) => {
      if (!draggedLi || draggedLi === li || draggedLi.contains(li)) return;
      const pos = computeDropPosition(row, e.clientY);
      if (!isDropAllowed(draggedLi, li, pos)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      setIndicator(row, pos);
    });
    row.addEventListener("dragleave", (e) => {
      if (!row.contains(e.relatedTarget)) clearIndicator(row);
    });
    row.addEventListener("drop", (e) => {
      if (!draggedLi || draggedLi === li || draggedLi.contains(li)) return;
      const pos = computeDropPosition(row, e.clientY);
      if (!isDropAllowed(draggedLi, li, pos)) return;
      e.preventDefault();
      performDrop(draggedLi, li, pos);
      clearDropIndicators();
      scheduleSave();
    });
  }

  function computeDropPosition(row, clientY) {
    const rect = row.getBoundingClientRect();
    const y = clientY - rect.top;
    const h = rect.height;
    if (y < h * 0.28) return "before";
    if (y > h * 0.72) return "after";
    return "inside";
  }

  function isDropAllowed(dragged, target, pos) {
    const targetLevel = parseInt(target.dataset.level, 10);
    const depth = subtreeDepth(dragged);
    let newRootLevel;
    if (pos === "inside") newRootLevel = targetLevel + 1;
    else newRootLevel = targetLevel;
    return (newRootLevel + depth - 1) <= maxLevels();
  }

  function setIndicator(row, pos) {
    row.classList.remove("drop-before", "drop-after", "drop-inside");
    row.classList.add("drop-" + pos);
  }
  function clearIndicator(row) {
    row.classList.remove("drop-before", "drop-after", "drop-inside");
  }
  function clearDropIndicators() {
    document.querySelectorAll(".drop-before,.drop-after,.drop-inside")
      .forEach((el) => el.classList.remove("drop-before", "drop-after", "drop-inside"));
  }

  function performDrop(dragged, target, pos) {
    const targetLevel = parseInt(target.dataset.level, 10);
    if (pos === "inside") {
      let ul = target.querySelector(":scope > ul.tree");
      if (!ul) {
        ul = document.createElement("ul");
        ul.className = "tree";
        target.appendChild(ul);
      }
      ul.appendChild(dragged);
      target.classList.remove("collapsed");
      relevel(dragged, targetLevel + 1);
      refreshToggle(target);
    } else {
      const parent = target.parentElement; // ul
      if (pos === "before") parent.insertBefore(dragged, target);
      else parent.insertBefore(dragged, target.nextElementSibling);
      relevel(dragged, targetLevel);
    }
    // Refresh old parent (source may have become empty)
    // (since we don't track old parent explicitly, just refresh all toggles)
    root.querySelectorAll("li.node").forEach(refreshToggle);
    renumber();
  }

  // ---------- Bullet helpers ----------
  // Black bullet: "  ●  " (indent 2)
  // White bullet: "      ○  " (indent 6 = black's 2 + nesting stride 4).
  // ○ wrapped in <span class="wb"> so it can be shrunk to match ● visually.
  const BLACK_PREFIX = "  ●  ";        // 5 chars in flat text
  const WHITE_PREFIX = "      ○  ";    // 9 chars in flat text
  const DASH_PREFIX  = "  -  ";        // initial dash (indent = 2); extra "  " pairs may precede
  const BLACK_LEN = 5;
  const WHITE_LEN = 9;
  // Dash line: any 2-space pairs of indent (>=1), then "-  ". Indent grows/shrinks via Tab / Shift+Tab.
  const DASH_RE = /^((?:  )+)-  /;

  function matchBullet(line) {
    if (line.startsWith(WHITE_PREFIX)) return { type: "white", prefix: WHITE_PREFIX, len: WHITE_LEN };
    if (line.startsWith(BLACK_PREFIX)) return { type: "black", prefix: BLACK_PREFIX, len: BLACK_LEN };
    const dm = line.match(DASH_RE);
    if (dm) return { type: "dash", prefix: dm[0], len: dm[0].length, indent: dm[1].length };
    return null;
  }

  // Numbered lists
  // Level 1: "  1.  " (2sp indent + arabic digits + ". " + 2sp)
  // Level 2: "      a.  " (6sp indent + lowercase letter + ". " + 2sp)
  const NUM1_RE = /^  (\d+)\.  /;
  const NUM2_RE = /^      ([a-z])\.  /;

  function matchNumber(line) {
    let m;
    if ((m = line.match(NUM2_RE))) return { type: "num2", prefix: m[0], len: m[0].length, token: m[1] };
    if ((m = line.match(NUM1_RE))) return { type: "num1", prefix: m[0], len: m[0].length, token: m[1] };
    return null;
  }

  function matchList(line) {
    return matchBullet(line) || matchNumber(line);
  }

  function nextToken(type, token) {
    if (type === "num1") return String(parseInt(token, 10) + 1);
    if (type === "num2") {
      if (token >= "z") return "z";
      return String.fromCharCode(token.charCodeAt(0) + 1);
    }
    return token;
  }

  // Walk lines above the caret to find the next number/letter to use when a
  // user nests/un-nests into a numbered list. Falls back to "1"/"a".
  function findNextNumber(txt, ctx, type) {
    const flat = getFlatText(txt);
    const priorBlock = flat.slice(0, Math.max(0, ctx.lineStart - 1));
    const lines = priorBlock.split("\n");
    for (let i = lines.length - 1; i >= 0; i--) {
      if (type === "num1") {
        const m = lines[i].match(NUM1_RE);
        if (m) return String(parseInt(m[1], 10) + 1);
      } else {
        const m = lines[i].match(NUM2_RE);
        if (m) {
          const n = String.fromCharCode(m[1].charCodeAt(0) + 1);
          return n > "z" ? "z" : n;
        }
      }
    }
    return type === "num1" ? "1" : "a";
  }

  function writeNumber(type, token) {
    const prefix = type === "num1" ? `  ${token}.  ` : `      ${token}.  `;
    // execCommand path keeps native undo/redo working and replaces any active selection atomically.
    document.execCommand("insertText", false, prefix);
  }

  function writeListItem(kind, token) {
    if (kind === "black" || kind === "white" || kind === "dash") writeBullet(kind);
    else writeNumber(kind, token);
  }

  // Inserts the bullet prefix via execCommand("insertText") so native undo/redo
  // treats the whole prefix as one atomic entry. insertHTML with nested spans
  // produced partial/ghost undo states, so ○ is inserted as plain text here.
  function writeBullet(type, customPrefix) {
    let prefix;
    if (type === "white") prefix = "      ○  ";
    else if (type === "dash") prefix = customPrefix || "  -  ";
    else prefix = "  ●  ";
    document.execCommand("insertText", false, prefix);
  }

  function getFlatText(el) {
    let text = "";
    (function walk(node) {
      for (const child of node.childNodes) {
        if (child.nodeType === 3) text += child.nodeValue;
        else if (child.nodeName === "BR") text += "\n";
        else if (child.nodeType === 1) walk(child);
      }
    })(el);
    return text;
  }

  function getCaretOffset(el) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return -1;
    const range = sel.getRangeAt(0);
    const container = range.startContainer;
    if (container !== el && !el.contains(container)) return -1;
    const pre = document.createRange();
    pre.selectNodeContents(el);
    pre.setEnd(container, range.startOffset);
    const frag = pre.cloneContents();
    let offset = 0;
    (function walk(node) {
      for (const child of node.childNodes) {
        if (child.nodeType === 3) offset += child.nodeValue.length;
        else if (child.nodeName === "BR") offset += 1;
        else if (child.nodeType === 1) walk(child);
      }
    })(frag);
    return offset;
  }

  function getLineContext(el) {
    const flat = getFlatText(el);
    const off = getCaretOffset(el);
    if (off < 0) return null;
    const lineStart = flat.lastIndexOf("\n", off - 1) + 1;
    const nextNL = flat.indexOf("\n", off);
    const lineEnd = nextNL === -1 ? flat.length : nextNL;
    return {
      off,
      lineStart,
      lineEnd,
      lineBefore: flat.slice(lineStart, off),
      line: flat.slice(lineStart, lineEnd),
    };
  }

  function setCaretAtOffset(el, target) {
    let remaining = target;
    let done = false;
    (function walk(node) {
      if (done) return;
      for (const child of node.childNodes) {
        if (done) return;
        if (child.nodeType === 3) {
          const len = child.nodeValue.length;
          if (remaining <= len) {
            const r = document.createRange();
            r.setStart(child, remaining);
            r.collapse(true);
            const s = window.getSelection();
            s.removeAllRanges();
            s.addRange(r);
            done = true;
            return;
          }
          remaining -= len;
        } else if (child.nodeName === "BR") {
          if (remaining === 0) {
            const r = document.createRange();
            r.setStartBefore(child);
            r.collapse(true);
            const s = window.getSelection();
            s.removeAllRanges();
            s.addRange(r);
            done = true;
            return;
          }
          remaining -= 1;
        } else if (child.nodeType === 1) {
          walk(child);
        }
      }
    })(el);
    if (!done) {
      const r = document.createRange();
      r.selectNodeContents(el);
      r.collapse(false);
      const s = window.getSelection();
      s.removeAllRanges();
      s.addRange(r);
    }
  }

  function selectOffsets(el, start, end) {
    setCaretAtOffset(el, start);
    const sel = window.getSelection();
    const a = sel.anchorNode, ao = sel.anchorOffset;
    setCaretAtOffset(el, end);
    const b = sel.anchorNode, bo = sel.anchorOffset;
    const r = document.createRange();
    r.setStart(a, ao);
    r.setEnd(b, bo);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  // Returns {start, end} as flat-text offsets for the current selection inside `el`.
  function getSelectionOffsets(el) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    if (range.startContainer !== el && !el.contains(range.startContainer)) return null;
    function offsetOf(container, offset) {
      const pre = document.createRange();
      pre.selectNodeContents(el);
      pre.setEnd(container, offset);
      const frag = pre.cloneContents();
      let n = 0;
      (function walk(node) {
        for (const c of node.childNodes) {
          if (c.nodeType === 3) n += c.nodeValue.length;
          else if (c.nodeName === "BR") n += 1;
          else if (c.nodeType === 1) walk(c);
        }
      })(frag);
      return n;
    }
    const start = offsetOf(range.startContainer, range.startOffset);
    const end = offsetOf(range.endContainer, range.endOffset);
    return { start: Math.min(start, end), end: Math.max(start, end) };
  }

  function insertListAtLineStart(txt, kind, token) {
    txt.focus();
    const offs = getSelectionOffsets(txt);
    const flat = getFlatText(txt);
    if (!offs) { writeListItem(kind, token); return; }

    const firstLineStart = flat.lastIndexOf("\n", offs.start - 1) + 1;
    const nlAfterEnd = flat.indexOf("\n", Math.max(offs.end - 1, firstLineStart));
    const lastLineEnd = nlAfterEnd === -1 ? flat.length : nlAfterEnd;

    // Collect lines covered by the selection, top-down.
    const lines = [];
    let p = firstLineStart;
    while (p <= lastLineEnd) {
      const nl = flat.indexOf("\n", p);
      const lineEnd = nl === -1 ? flat.length : nl;
      lines.push({ start: p, end: lineEnd, text: flat.slice(p, lineEnd) });
      if (nl === -1 || lineEnd >= lastLineEnd) break;
      p = nl + 1;
    }

    // Single-line fast path (preserves existing behavior, including caret placement).
    if (lines.length <= 1) {
      const ln = lines[0] || { start: firstLineStart, end: lastLineEnd, text: "" };
      const existing = matchList(ln.text);
      if (existing) selectOffsets(txt, ln.start, ln.start + existing.len);
      else setCaretAtOffset(txt, ln.start);
      writeListItem(kind, token);
      return;
    }

    // Multi-line: apply bottom-up so earlier offsets remain valid.
    for (let i = lines.length - 1; i >= 0; i--) {
      const ln = lines[i];
      let t = token;
      if (kind === "num1") {
        t = String(parseInt(token || "1", 10) + i);
      } else if (kind === "num2") {
        const base = (token || "a").charCodeAt(0);
        t = String.fromCharCode(Math.min(base + i, "z".charCodeAt(0)));
      }
      const existing = matchList(ln.text);
      if (existing) selectOffsets(txt, ln.start, ln.start + existing.len);
      else setCaretAtOffset(txt, ln.start);
      writeListItem(kind, t);
    }
  }

  // Indent/outdent logic for a single line (used by Tab / Shift+Tab, per-line in multi-line selections).
  function indentLine(txt, ln) {
    const m = matchList(ln.text);
    if (m && m.type === "dash") {
      setCaretAtOffset(txt, ln.start);
      document.execCommand("insertText", false, "  ");
    } else if (m && m.type === "black") {
      selectOffsets(txt, ln.start, ln.start + m.len);
      writeBullet("white");
    } else if (m && m.type === "num1") {
      selectOffsets(txt, ln.start, ln.start + m.len);
      writeNumber("num2", "a");
    } else if (m && (m.type === "white" || m.type === "num2")) {
      // Already at max nest in its chain; no-op for multi-line indent.
      return;
    } else {
      // Non-list line: prepend 2 spaces at line start (multi-line indent convention).
      setCaretAtOffset(txt, ln.start);
      document.execCommand("insertText", false, "  ");
    }
  }

  function outdentLine(txt, ln) {
    const m = matchList(ln.text);
    if (m && m.type === "dash") {
      if (m.indent > 2) {
        selectOffsets(txt, ln.start, ln.start + 2);
        document.execCommand("delete");
      } else {
        selectOffsets(txt, ln.start, ln.start + m.len);
        document.execCommand("delete");
      }
    } else if (m && m.type === "white") {
      selectOffsets(txt, ln.start, ln.start + m.len);
      writeBullet("black");
    } else if (m && m.type === "black") {
      selectOffsets(txt, ln.start, ln.start + m.len);
      document.execCommand("delete");
    } else if (m && m.type === "num2") {
      selectOffsets(txt, ln.start, ln.start + m.len);
      writeNumber("num1", "1");
    } else if (m && m.type === "num1") {
      selectOffsets(txt, ln.start, ln.start + m.len);
      document.execCommand("delete");
    } else if (ln.text.startsWith("  ")) {
      selectOffsets(txt, ln.start, ln.start + 2);
      document.execCommand("delete");
    }
  }

  // ---------- Keyboard shortcuts ----------
  function handleKey(e, li) {
    // Bold
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "b") {
      e.preventDefault();
      document.execCommand("bold");
      scheduleSave();
      return;
    }
    // Italic (bonus)
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "i") {
      e.preventDefault();
      document.execCommand("italic");
      scheduleSave();
      return;
    }
    // Ctrl+Enter -> add sibling
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      addSibling(li);
      scheduleSave();
      return;
    }
    // Arrow Up/Down -> move caret to prev/next cell in document order
    if (e.key === "ArrowUp" && !e.ctrlKey && !e.metaKey && !e.shiftKey && !e.altKey) {
      const prev = getPrevCell(li);
      if (prev) {
        e.preventDefault();
        focusNode(prev);
      }
      return;
    }
    if (e.key === "ArrowDown" && !e.ctrlKey && !e.metaKey && !e.shiftKey && !e.altKey) {
      const next = getNextCell(li);
      if (next) {
        e.preventDefault();
        focusNode(next);
      }
      return;
    }
    // Shift+Enter -> jump to next cell in document order, or create a sibling if there is none
    if (e.key === "Enter" && e.shiftKey) {
      e.preventDefault();
      const next = getNextCell(li);
      if (next) {
        focusNode(next);
      } else {
        addSibling(li);
        scheduleSave();
      }
      return;
    }
    // Plain Enter -> newline inside text (with list continuation)
    if (e.key === "Enter" && !e.shiftKey) {
      const txt = li.querySelector(":scope > .row > .text");
      const ctx = txt ? getLineContext(txt) : null;
      const m = ctx ? matchList(ctx.line) : null;
      if (m) {
        e.preventDefault();
        if (ctx.line === m.prefix) {
          // Empty list item -> remove the prefix
          selectOffsets(txt, ctx.lineStart, ctx.lineStart + m.len);
          document.execCommand("delete");
        } else if (ctx.lineBefore.length >= m.len) {
          document.execCommand("insertLineBreak");
          if (m.type === "dash") {
            writeBullet("dash", m.prefix);
          } else if (m.type === "black" || m.type === "white") {
            writeBullet(m.type);
          } else {
            writeNumber(m.type, nextToken(m.type, m.token));
          }
        } else {
          document.execCommand("insertLineBreak");
        }
        scheduleSave();
        return;
      }
      e.preventDefault();
      document.execCommand("insertLineBreak");
      scheduleSave();
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      const txt = li.querySelector(":scope > .row > .text");
      if (!txt) return;
      const offs = getSelectionOffsets(txt);
      const flat = getFlatText(txt);
      const spansMultipleLines =
        offs && offs.start !== offs.end &&
        flat.indexOf("\n", offs.start) !== -1 &&
        flat.indexOf("\n", offs.start) < offs.end;

      if (spansMultipleLines) {
        const firstLineStart = flat.lastIndexOf("\n", offs.start - 1) + 1;
        const nlAfterEnd = flat.indexOf("\n", Math.max(offs.end - 1, firstLineStart));
        const lastLineEnd = nlAfterEnd === -1 ? flat.length : nlAfterEnd;
        const lines = [];
        let p = firstLineStart;
        while (p <= lastLineEnd) {
          const nl = flat.indexOf("\n", p);
          const lineEnd = nl === -1 ? flat.length : nl;
          lines.push({ start: p, end: lineEnd, text: flat.slice(p, lineEnd) });
          if (nl === -1 || lineEnd >= lastLineEnd) break;
          p = nl + 1;
        }
        // Bottom-up so offsets above stay valid after each mutation.
        for (let i = lines.length - 1; i >= 0; i--) {
          if (e.shiftKey) outdentLine(txt, lines[i]);
          else indentLine(txt, lines[i]);
        }
        scheduleSave();
        return;
      }

      // Single-line path
      const ctx = getLineContext(txt);
      const ln = ctx ? { start: ctx.lineStart, end: ctx.lineEnd, text: ctx.line } : null;
      if (!ln) {
        if (!e.shiftKey) document.execCommand("insertText", false, "  ");
        scheduleSave();
        return;
      }
      if (e.shiftKey) {
        outdentLine(txt, ln);
      } else {
        const m = matchList(ln.text);
        if (!m || m.type === "white" || m.type === "num2") {
          // Non-list line or already-max-nested prefix: insert 2 spaces at the current caret.
          document.execCommand("insertText", false, "  ");
        } else if (m.type === "num1") {
          // 1. -> a. continues numbering from any earlier a./b./... above this line
          selectOffsets(txt, ln.start, ln.start + m.len);
          writeNumber("num2", findNextNumber(txt, { lineStart: ln.start }, "num2"));
        } else {
          indentLine(txt, ln);
        }
      }
      scheduleSave();
      return;
    }
    if (e.key === "Delete" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      removeNode(li);
      scheduleSave();
      return;
    }
  }

  // ---------- Sanitization ----------
  const ALLOWED_TAGS = new Set(["B", "STRONG", "I", "EM", "BR", "SPAN"]);
  function sanitizeHTML(html) {
    const div = document.createElement("div");
    div.innerHTML = html;
    function walk(node) {
      [...node.childNodes].forEach((child) => {
        if (child.nodeType === 1) {
          if (!ALLOWED_TAGS.has(child.tagName)) {
            while (child.firstChild) node.insertBefore(child.firstChild, child);
            node.removeChild(child);
            return;
          }
          if (child.tagName === "SPAN") {
            const cls = child.getAttribute("class");
            [...child.attributes].forEach((a) => child.removeAttribute(a.name));
            if (cls === "wb") {
              child.setAttribute("class", "wb");
              walk(child);
            } else {
              while (child.firstChild) node.insertBefore(child.firstChild, child);
              node.removeChild(child);
            }
            return;
          }
          [...child.attributes].forEach((a) => child.removeAttribute(a.name));
          walk(child);
        } else if (child.nodeType !== 3) {
          child.remove();
        }
      });
    }
    walk(div);
    return div.innerHTML;
  }

  // ---------- Serialization ----------
  function serialize(ul) {
    const items = [];
    ul.querySelectorAll(":scope > li.node").forEach((li) => {
      const txt = li.querySelector(":scope > .row > .text");
      const childUl = li.querySelector(":scope > ul.tree");
      items.push({
        html: sanitizeHTML(txt.innerHTML),
        collapsed: li.classList.contains("collapsed"),
        plain: li.dataset.plain === "true" || undefined,
        children: childUl ? serialize(childUl) : [],
      });
    });
    return items;
  }

  function deserialize(items, parentUl, level) {
    items.forEach((item) => {
      const html = item.html != null ? item.html : (item.text || "");
      const li = createNode(level, sanitizeHTML(html), !!item.plain);
      parentUl.appendChild(li);
      if (item.children && item.children.length) {
        const ul = document.createElement("ul");
        ul.className = "tree";
        li.appendChild(ul);
        deserialize(item.children, ul, level + 1);
        refreshToggle(li);
        if (item.collapsed) li.classList.add("collapsed");
      }
    });
  }

  function buildState() {
    return {
      title: tocTitle.textContent,
      subtitle: tocSub.textContent,
      maxLevels: maxLevels(),
      items: serialize(root),
    };
  }

  function loadState(data) {
    root.innerHTML = "";
    tocTitle.textContent = data.title || "";
    tocSub.textContent = data.subtitle || "";
    if (data.maxLevels) maxLevelsInput.value = data.maxLevels;
    deserialize(data.items || [], root, 1);
    renumber();
    updateEmpty();
  }

  // ---------- Project store ----------
  function uid() {
    return "p_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
  }

  function persistStore() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
    } catch (err) {
      setStatus("Save failed: " + err.message);
    }
  }

  function getCurrentProject() {
    return store.projects.find((p) => p.id === store.currentId) || null;
  }

  function deriveProjectName(data) {
    const t = (data.title || "").trim();
    if (t) return t;
    const firstItem = (data.items && data.items[0]) || null;
    if (firstItem) {
      const tmp = document.createElement("div");
      tmp.innerHTML = firstItem.html || "";
      const txt = tmp.textContent.trim();
      if (txt) return txt.slice(0, 60);
    }
    return "Untitled";
  }

  function createProject(seedData) {
    const now = Date.now();
    const data = seedData || { title: "", subtitle: "", maxLevels: 10, items: [] };
    const p = {
      id: uid(),
      name: deriveProjectName(data),
      createdAt: now,
      updatedAt: now,
      data,
    };
    store.projects.push(p);
    store.currentId = p.id;
    persistStore();
    return p;
  }

  function switchProject(id) {
    if (store.currentId === id) return;
    const p = store.projects.find((x) => x.id === id);
    if (!p) return;
    flushSave();
    store.currentId = id;
    persistStore();
    loadState(p.data);
    renderSidebar();
    updateActiveTitle();
  }

  function renameProject(id, name) {
    const p = store.projects.find((x) => x.id === id);
    if (!p) return;
    p.name = name.trim() || "Untitled";
    p.updatedAt = Date.now();
    persistStore();
    renderSidebar();
    if (p.id === store.currentId) updateActiveTitle();
  }

  function deleteProject(id) {
    const idx = store.projects.findIndex((x) => x.id === id);
    if (idx === -1) return;
    store.projects.splice(idx, 1);
    if (store.currentId === id) {
      if (store.projects.length === 0) {
        const p = createProject();
        loadState(p.data);
      } else {
        const next = store.projects[Math.min(idx, store.projects.length - 1)];
        store.currentId = next.id;
        loadState(next.data);
      }
    }
    persistStore();
    renderSidebar();
    updateActiveTitle();
  }

  function updateActiveTitle() {
    const p = getCurrentProject();
    activeProjectName.textContent = p ? p.name : "Plan, outline, present.";
  }

  // ---------- Auto-save ----------
  let saveTimer = null;
  function doSave() {
    try {
      const p = getCurrentProject();
      if (!p) return;
      p.data = buildState();
      p.updatedAt = Date.now();
      const derived = deriveProjectName(p.data);
      const nameChanged = p.name !== derived;
      const looksAuto = !p._manualName;
      if (looksAuto && nameChanged) {
        p.name = derived;
        renderSidebar();
        updateActiveTitle();
      } else {
        renderSidebarMeta(p.id);
      }
      persistStore();
      flash("Saved");
    } catch (err) {
      setStatus("Auto-save failed: " + err.message);
    }
  }
  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(doSave, 300);
  }
  function flushSave() {
    if (saveTimer) {
      clearTimeout(saveTimer);
      saveTimer = null;
      doSave();
    }
  }
  window.addEventListener("beforeunload", flushSave);

  function loadStore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.projects)) {
          store = parsed;
          return true;
        }
      }
    } catch {}
    // Migrate from legacy single-document key
    for (const key of LEGACY_KEYS) {
      try {
        const raw = localStorage.getItem(key);
        if (!raw) continue;
        const data = JSON.parse(raw);
        const p = createProject(data);
        p.name = deriveProjectName(data);
        persistStore();
        return true;
      } catch {}
    }
    return false;
  }

  // ---------- Sidebar rendering ----------
  // Full date + time in Jakarta (WIB / Asia/Jakarta) so the user can see
  // exactly when a project was last edited. Format: "27 Apr 2026, 14:30".
  function formatTime(ts) {
    const d = new Date(ts);
    return d.toLocaleString("en-GB", {
      timeZone: "Asia/Jakarta",
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    });
  }

  // Natural, case-insensitive name sort (A-Z, 0-9 like Windows Explorer).
  const projectCollator = new Intl.Collator(undefined, {
    numeric: true,
    sensitivity: "base",
  });

  function renderSidebar() {
    const q = searchQuery.trim().toLowerCase();
    const sorted = [...store.projects].sort((a, b) =>
      projectCollator.compare(a.name || "", b.name || "")
    );
    const filtered = q ? sorted.filter((p) => p.name.toLowerCase().includes(q)) : sorted;
    projectListEl.innerHTML = "";
    if (filtered.length === 0) {
      const empty = document.createElement("div");
      empty.className = "project-empty";
      empty.textContent = q ? "No matches." : "No projects yet.";
      projectListEl.appendChild(empty);
      return;
    }
    filtered.forEach((p) => {
      projectListEl.appendChild(renderProjectItem(p));
    });
  }

  function renderProjectItem(p) {
    const item = document.createElement("div");
    item.className = "project-item" + (p.id === store.currentId ? " active" : "");
    item.dataset.id = p.id;

    const info = document.createElement("div");
    info.className = "project-info";
    const name = document.createElement("div");
    name.className = "project-name";
    name.textContent = p.name;
    const meta = document.createElement("div");
    meta.className = "project-meta";
    meta.textContent = formatTime(p.updatedAt);
    info.append(name, meta);

    const actions = document.createElement("div");
    actions.className = "project-actions";
    const renameBtn = document.createElement("button");
    renameBtn.className = "act-rename";
    renameBtn.title = "Rename";
    renameBtn.innerHTML = "&#9998;";
    renameBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      startRename(item, p);
    });
    const delBtn = document.createElement("button");
    delBtn.className = "act-del";
    delBtn.title = "Delete";
    delBtn.innerHTML = "&times;";
    delBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (confirm("Delete project \"" + p.name + "\"? This cannot be undone.")) {
        deleteProject(p.id);
      }
    });
    actions.append(renameBtn, delBtn);

    item.append(info, actions);
    item.addEventListener("click", () => switchProject(p.id));
    item.addEventListener("dblclick", (e) => {
      e.preventDefault();
      startRename(item, p);
    });
    return item;
  }

  function startRename(item, p) {
    const nameEl = item.querySelector(".project-name");
    if (!nameEl) return;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "project-name-input";
    input.value = p.name;
    nameEl.replaceWith(input);
    input.focus();
    input.select();
    const commit = (save) => {
      if (save) {
        p._manualName = true;
        renameProject(p.id, input.value);
      } else {
        renderSidebar();
      }
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); commit(true); }
      else if (e.key === "Escape") { e.preventDefault(); commit(false); }
    });
    input.addEventListener("blur", () => commit(true));
    input.addEventListener("click", (e) => e.stopPropagation());
  }

  function renderSidebarMeta(id) {
    // Light update of timestamp for a single item without full rerender (cheap path)
    const el = projectListEl.querySelector('.project-item[data-id="' + id + '"] .project-meta');
    const p = store.projects.find((x) => x.id === id);
    if (el && p) el.textContent = formatTime(p.updatedAt);
  }

  // ---------- Status ----------
  let flashTimer = null;
  function flash(msg) {
    statusEl.textContent = msg;
    statusEl.classList.add("flash");
    clearTimeout(flashTimer);
    flashTimer = setTimeout(() => {
      statusEl.classList.remove("flash");
      statusEl.textContent = "Ready";
    }, 1200);
  }
  function setStatus(msg) { statusEl.textContent = msg; }

  // ---------- Toolbar ----------
  document.getElementById("addRoot").addEventListener("click", () => {
    const li = createNode(1);
    root.appendChild(li);
    renumber();
    focusNode(li);
    updateEmpty();
    scheduleSave();
  });
  document.getElementById("expandAll").addEventListener("click", () => {
    root.querySelectorAll("li.node").forEach((li) => li.classList.remove("collapsed"));
    scheduleSave();
  });
  document.getElementById("collapseAll").addEventListener("click", () => {
    root.querySelectorAll("li.node").forEach((li) => {
      const ul = li.querySelector(":scope > ul.tree");
      if (ul && ul.children.length) li.classList.add("collapsed");
    });
    scheduleSave();
  });
  document.getElementById("saveBtn").addEventListener("click", () => {
    const data = buildState();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const safe = (data.title || "outliner").replace(/[^\w\-]+/g, "_").slice(0, 60) || "outliner";
    a.href = url;
    a.download = safe + ".json";
    a.click();
    URL.revokeObjectURL(url);
    flash("Exported JSON");
  });
  document.getElementById("loadFile").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        // Load into a brand-new project so we don't overwrite the current one
        const p = createProject(data);
        const baseName = file.name.replace(/\.json$/i, "");
        if (baseName) { p.name = baseName; p._manualName = true; }
        loadState(p.data);
        renderSidebar();
        updateActiveTitle();
        persistStore();
        flash("Loaded as new project");
      } catch (err) {
        alert("Invalid JSON: " + err.message);
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  });
  document.getElementById("exportPng").addEventListener("click", async () => {
    if (typeof html2canvas === "undefined") {
      alert("PNG library not loaded. Check your internet connection.");
      return;
    }
    document.body.classList.add("exporting");
    setStatus("Rendering PNG...");
    try {
      const canvas = await html2canvas(paper, { backgroundColor: "#ffffff", scale: 2 });
      const link = document.createElement("a");
      const safe = (tocTitle.textContent || "outliner").replace(/[^\w\-]+/g, "_").slice(0, 60) || "outliner";
      link.download = safe + ".png";
      link.href = canvas.toDataURL("image/png");
      link.click();
      flash("PNG exported");
    } catch (err) {
      alert("Export failed: " + err.message);
      setStatus("Ready");
    } finally {
      document.body.classList.remove("exporting");
    }
  });
  document.getElementById("clearBtn").addEventListener("click", () => {
    if (!confirm("Clear everything? This cannot be undone.")) return;
    root.innerHTML = "";
    tocTitle.textContent = "";
    tocSub.textContent = "";
    const li = createNode(1);
    root.appendChild(li);
    renumber();
    focusNode(li);
    updateEmpty();
    scheduleSave();
  });
  maxLevelsInput.addEventListener("change", () => {
    root.querySelectorAll("li.node").forEach((li) => refreshChildButton(li));
    scheduleSave();
  });
  tocTitle.addEventListener("input", scheduleSave);
  tocSub.addEventListener("input", scheduleSave);

  // ---------- Sidebar wiring ----------
  newProjectBtn.addEventListener("click", () => {
    flushSave();
    const p = createProject();
    loadState(p.data);
    root.appendChild(createNode(1));
    renumber();
    updateEmpty();
    renderSidebar();
    updateActiveTitle();
    scheduleSave();
    tocTitle.focus();
  });
  const isMobile = () => window.matchMedia("(max-width: 900px)").matches;
  function setSidebarOpen(open) {
    sidebar.classList.toggle("collapsed", !open);
    sidebarBackdrop.classList.toggle("show", open && isMobile());
    if (!isMobile()) {
      try { localStorage.setItem("outliner.sidebarCollapsed", open ? "0" : "1"); } catch {}
    }
  }
  sidebarToggle.addEventListener("click", () => {
    setSidebarOpen(sidebar.classList.contains("collapsed"));
  });
  sidebarClose.addEventListener("click", () => setSidebarOpen(false));
  sidebarBackdrop.addEventListener("click", () => setSidebarOpen(false));
  window.addEventListener("resize", () => {
    if (!isMobile()) sidebarBackdrop.classList.remove("show");
  });
  projectListEl.addEventListener("click", (e) => {
    if (e.target.closest(".project-item") && isMobile()) {
      setSidebarOpen(false);
    }
  });
  projectSearch.addEventListener("input", () => {
    searchQuery = projectSearch.value;
    renderSidebar();
  });

  // ---------- Hover shortcuts (not in text-edit mode) ----------
  // a = add sibling above, b = add sibling below, dd (double-d) = delete the hovered cell.
  let hoveredLi = null;
  let lastDPressAt = 0;
  const DD_WINDOW_MS = 500;

  root.addEventListener("mouseover", (e) => {
    const li = e.target.closest("li.node");
    if (li && root.contains(li)) hoveredLi = li;
  });
  root.addEventListener("mouseout", (e) => {
    if (!e.relatedTarget || !root.contains(e.relatedTarget)) hoveredLi = null;
  });

  function isEditingText() {
    const a = document.activeElement;
    if (!a) return false;
    if (a.tagName === "INPUT" || a.tagName === "TEXTAREA") return true;
    if (a.isContentEditable) return true;
    return false;
  }

  function blurActive() {
    const a = document.activeElement;
    if (a && typeof a.blur === "function") a.blur();
  }

  document.addEventListener("keydown", (e) => {
    if (isEditingText()) return;
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (!hoveredLi || !root.contains(hoveredLi)) return;
    const k = e.key.toLowerCase();
    if (k === "a") {
      e.preventDefault();
      addPlainSiblingAbove(hoveredLi);
      blurActive();
      scheduleSave();
    } else if (k === "b") {
      e.preventDefault();
      addPlainSibling(hoveredLi);
      blurActive();
      scheduleSave();
    } else if (k === "d") {
      e.preventDefault();
      const now = Date.now();
      if (now - lastDPressAt <= DD_WINDOW_MS) {
        lastDPressAt = 0;
        const toDelete = hoveredLi;
        removeNode(toDelete);
        blurActive();
        scheduleSave();
      } else {
        lastDPressAt = now;
      }
    } else {
      lastDPressAt = 0;
    }
  });

  // ---------- Init ----------
  loadStore();
  let current = getCurrentProject();
  if (!current) {
    current = createProject();
  }
  loadState(current.data);
  if (root.children.length === 0) {
    root.appendChild(createNode(1));
    scheduleSave();
  }
  renumber();
  updateEmpty();
  renderSidebar();
  updateActiveTitle();
  if (isMobile()) {
    sidebar.classList.add("collapsed");
  } else {
    try {
      if (localStorage.getItem("outliner.sidebarCollapsed") === "1") {
        sidebar.classList.add("collapsed");
      }
    } catch {}
  }
})();
