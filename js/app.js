(() => {
  "use strict";

  const STORAGE_KEY = "toc-builder.v2";

  const root = document.getElementById("root");
  const paper = document.getElementById("paper");
  const maxLevelsInput = document.getElementById("maxLevels");
  const statusEl = document.getElementById("status");
  const tocTitle = document.getElementById("toc-title");
  const tocSub = document.getElementById("toc-sub");

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
  function createNode(level, html = "") {
    const li = document.createElement("li");
    li.className = "node";
    li.dataset.level = level;

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
    txt.dataset.placeholder = placeholders[level - 1] || "Level " + level;
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

    // Actions: sibling first (+), then child, bold, delete
    const actions = document.createElement("div");
    actions.className = "actions";
    actions.appendChild(makeActionBtn("btn-sibling", "+", "Add item (Ctrl+Enter)", () => addSibling(li)));
    actions.appendChild(makeActionBtn("btn-child", "↳", "Add sub-item (Tab)", () => addChild(li)));
    actions.appendChild(makeActionBtn("btn-bold", "B", "Bold (Ctrl+B)", () => { txt.focus(); document.execCommand("bold"); }));
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
      [...ul.children].forEach((li, i) => {
        const num = prefix ? prefix + "." + (i + 1) : String(i + 1);
        const badge = li.querySelector(":scope > .row > .level-badge");
        if (badge) badge.textContent = num;
        const child = li.querySelector(":scope > ul.tree");
        if (child) walk(child, num);
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
    // Plain Enter -> newline inside text
    if (e.key === "Enter") {
      e.preventDefault();
      document.execCommand("insertLineBreak");
      scheduleSave();
      return;
    }
    if (e.key === "Tab" && !e.shiftKey) {
      e.preventDefault();
      addChild(li);
      scheduleSave();
      return;
    }
    if (e.key === "Tab" && e.shiftKey) {
      e.preventDefault();
      toggleCollapse(li);
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
  const ALLOWED_TAGS = new Set(["B", "STRONG", "I", "EM", "BR"]);
  function sanitizeHTML(html) {
    const div = document.createElement("div");
    div.innerHTML = html;
    function walk(node) {
      [...node.childNodes].forEach((child) => {
        if (child.nodeType === 1) {
          if (!ALLOWED_TAGS.has(child.tagName)) {
            while (child.firstChild) node.insertBefore(child.firstChild, child);
            node.removeChild(child);
          } else {
            [...child.attributes].forEach((a) => child.removeAttribute(a.name));
            walk(child);
          }
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
        children: childUl ? serialize(childUl) : [],
      });
    });
    return items;
  }

  function deserialize(items, parentUl, level) {
    items.forEach((item) => {
      const html = item.html != null ? item.html : (item.text || "");
      const li = createNode(level, sanitizeHTML(html));
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

  // ---------- Auto-save ----------
  let saveTimer = null;
  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(buildState()));
        flash("Saved");
      } catch (err) {
        setStatus("Auto-save failed: " + err.message);
      }
    }, 300);
  }

  function restore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY)
        || localStorage.getItem("toc-builder.v1"); // migrate
      if (!raw) return false;
      loadState(JSON.parse(raw));
      return true;
    } catch {
      return false;
    }
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
    const safe = (data.title || "toc").replace(/[^\w\-]+/g, "_").slice(0, 60) || "toc";
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
        loadState(JSON.parse(reader.result));
        scheduleSave();
        flash("Loaded");
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
      const safe = (tocTitle.textContent || "toc").replace(/[^\w\-]+/g, "_").slice(0, 60) || "toc";
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

  // ---------- Init ----------
  if (!restore()) {
    root.appendChild(createNode(1));
  }
  renumber();
  updateEmpty();
})();
