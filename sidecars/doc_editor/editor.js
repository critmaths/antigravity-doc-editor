/**
 * Antigravity Document Editor - Client Side WYSIWYG & Sidecar Controller
 * Supports Word (.docx) & OpenDocument (.odt) documents with live disk saving.
 */

(function () {
  'use strict';

  // API Base resolution (runs on same sidecar port)
  const API_BASE = window.location.origin;

  // Editor State
  let currentDocPath = '';
  let currentDocTitle = 'Untitled Document';
  let currentDocFormat = 'odt';
  let isDirty = false;
  let isSaving = false;
  let modalConfirmCallback = null;

  // DOM Elements
  const docTitleInput = document.getElementById('doc-title');
  const docFormatBadge = document.getElementById('doc-format-badge');
  const docSelector = document.getElementById('doc-selector');
  const saveStatus = document.getElementById('save-status');
  const btnNew = document.getElementById('btn-new');
  const btnSave = document.getElementById('btn-save');
  const btnSaveAs = document.getElementById('btn-save-as');
  const saveAsMenu = document.getElementById('save-as-menu');
  const editorCanvas = document.getElementById('editor-canvas');

  // Toolbar Elements
  const btnUndo = document.getElementById('btn-undo');
  const btnRedo = document.getElementById('btn-redo');
  const headingSelect = document.getElementById('heading-select');
  const fontFamilySelect = document.getElementById('font-family-select');
  const fontSizeSelect = document.getElementById('font-size-select');
  const btnBold = document.getElementById('btn-bold');
  const btnItalic = document.getElementById('btn-italic');
  const btnUnderline = document.getElementById('btn-underline');
  const btnStrike = document.getElementById('btn-strike');
  const colorPicker = document.getElementById('color-picker');
  const highlightPicker = document.getElementById('highlight-picker');
  const btnAlignLeft = document.getElementById('btn-align-left');
  const btnAlignCenter = document.getElementById('btn-align-center');
  const btnAlignRight = document.getElementById('btn-align-right');
  const btnAlignJustify = document.getElementById('btn-align-justify');
  const btnListUl = document.getElementById('btn-list-ul');
  const btnListOl = document.getElementById('btn-list-ol');
  const btnLink = document.getElementById('btn-link');
  const btnTable = document.getElementById('btn-table');
  const btnClearFormat = document.getElementById('btn-clear-format');

  // Modal Elements
  const modalOverlay = document.getElementById('modal-overlay');
  const modalTitle = document.getElementById('modal-title');
  const modalBody = document.getElementById('modal-body');
  const modalConfirm = document.getElementById('modal-confirm');
  const modalCancel = document.getElementById('modal-cancel');

  // ===========================================================================
  // Modal System (Custom In-Page HTML Dialog)
  // ===========================================================================
  function openModal(options) {
    modalTitle.textContent = options.title || 'Dialog';
    modalBody.innerHTML = options.bodyHtml || '';
    modalConfirmCallback = options.onConfirm || null;
    modalOverlay.classList.remove('hidden');

    const firstInput = modalBody.querySelector('input, select');
    if (firstInput) {
      setTimeout(() => firstInput.focus(), 50);
    }
  }

  function closeModal() {
    modalOverlay.classList.add('hidden');
    modalBody.innerHTML = '';
    modalConfirmCallback = null;
  }

  modalConfirm.addEventListener('click', () => {
    if (modalConfirmCallback) {
      const shouldClose = modalConfirmCallback();
      if (shouldClose !== false) {
        closeModal();
      }
    } else {
      closeModal();
    }
  });

  modalCancel.addEventListener('click', closeModal);
  modalOverlay.addEventListener('click', (e) => {
    if (e.target === modalOverlay) closeModal();
  });

  // ===========================================================================
  // Status & State Helpers
  // ===========================================================================
  function setDirty(dirty) {
    isDirty = dirty;
    if (dirty) {
      saveStatus.textContent = 'Unsaved changes';
      saveStatus.className = 'save-status dirty';
    } else {
      saveStatus.textContent = 'Saved';
      saveStatus.className = 'save-status';
    }
  }

  function updateFormatBadge(format) {
    currentDocFormat = (format || 'odt').toLowerCase();
    docFormatBadge.textContent = currentDocFormat.toUpperCase();
    docFormatBadge.className = `format-badge format-${currentDocFormat}`;
  }

  // ===========================================================================
  // Document Operations (Load, Save, Create, Convert)
  // ===========================================================================
  async function loadDocument(filePath) {
    if (!filePath) return;
    try {
      saveStatus.textContent = 'Loading...';
      saveStatus.className = 'save-status';
      const encoded = encodeURIComponent(filePath);
      const res = await fetch(`${API_BASE}/api/document?path=${encoded}`);
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Failed to load document (${res.status})`);
      }
      const data = await res.json();
      currentDocPath = data.path;
      currentDocTitle = data.title || 'Untitled Document';
      docTitleInput.value = currentDocTitle;
      updateFormatBadge(data.format);
      editorCanvas.innerHTML = data.html || '<p></p>';
      setDirty(false);

      // Update URL query string without reloading
      const url = new URL(window.location.href);
      url.searchParams.set('file', currentDocPath);
      window.history.replaceState({}, '', url.toString());

      // Update selector dropdown if needed
      if (docSelector) {
        docSelector.value = currentDocPath;
      }
    } catch (err) {
      alert(`Error loading document: ${err.message}`);
      saveStatus.textContent = 'Load error';
    }
  }

  async function saveDocument() {
    if (isSaving) return;
    if (!currentDocPath) {
      createNewDocument();
      return;
    }

    try {
      isSaving = true;
      saveStatus.textContent = 'Saving...';
      saveStatus.className = 'save-status saving';

      const payload = {
        path: currentDocPath,
        html: editorCanvas.innerHTML
      };

      const res = await fetch(`${API_BASE}/api/document`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Failed to save (${res.status})`);
      }

      const result = await res.json();
      setDirty(false);
      saveStatus.textContent = 'Saved to disk';
      setTimeout(() => {
        if (!isDirty) saveStatus.textContent = 'Saved';
      }, 2500);
    } catch (err) {
      alert(`Save failed: ${err.message}`);
      saveStatus.textContent = 'Save failed';
      saveStatus.className = 'save-status dirty';
    } finally {
      isSaving = false;
    }
  }

  function createNewDocument() {
    const defaultDir = currentDocPath ? currentDocPath.substring(0, currentDocPath.lastIndexOf('\\') + 1) : '';
    const bodyHtml = `
      <div>
        <label for="new-doc-name">Document Title / File Name</label>
        <input type="text" id="new-doc-name" value="New Document" style="width: 100%; margin-top: 6px;">
      </div>
      <div>
        <label for="new-doc-format">Format</label>
        <select id="new-doc-format" style="width: 100%; margin-top: 6px; padding: 6px; background: #0f172a; color: #fff; border: 1px solid #334155; border-radius: 4px;">
          <option value="docx" ${currentDocFormat === 'docx' ? 'selected' : ''}>Microsoft Word (.docx)</option>
          <option value="odt" ${currentDocFormat === 'odt' ? 'selected' : ''}>OpenDocument Text (.odt)</option>
        </select>
      </div>
    `;

    openModal({
      title: 'Create New Document',
      bodyHtml: bodyHtml,
      onConfirm: async () => {
        const nameInput = document.getElementById('new-doc-name');
        const formatSelect = document.getElementById('new-doc-format');
        const title = (nameInput.value || 'New Document').trim();
        const ext = formatSelect.value;
        const sanitized = title.replace(/[^a-zA-Z0-9_\- ]/g, '').trim() || 'Document';
        const fileName = `${sanitized}.${ext}`;
        const newPath = defaultDir ? `${defaultDir}${fileName}` : fileName;

        try {
          const res = await fetch(`${API_BASE}/api/document`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              createNew: true,
              path: newPath,
              title: title,
              html: `<h1>${title}</h1><p>Start writing here...</p>`
            })
          });
          if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.error || 'Failed to create new document');
          }
          await loadDocumentsList();
          await loadDocument(newPath);
        } catch (err) {
          alert(`Creation failed: ${err.message}`);
        }
      }
    });
  }

  function convertDocument(targetFormat) {
    const ext = targetFormat.toLowerCase();
    const currentName = currentDocPath ? currentDocPath.substring(currentDocPath.lastIndexOf('\\') + 1) : 'document';
    const baseName = currentName.substring(0, currentName.lastIndexOf('.')) || currentName;
    const defaultDir = currentDocPath ? currentDocPath.substring(0, currentDocPath.lastIndexOf('\\') + 1) : '';
    const suggestedTarget = `${defaultDir}${baseName}_exported.${ext}`;

    const bodyHtml = `
      <div>
        <label for="export-target-path">Target File Path (${ext.toUpperCase()})</label>
        <input type="text" id="export-target-path" value="${suggestedTarget}" style="width: 100%; margin-top: 6px;">
      </div>
    `;

    openModal({
      title: `Export as ${ext.toUpperCase()}`,
      bodyHtml: bodyHtml,
      onConfirm: async () => {
        const targetInput = document.getElementById('export-target-path');
        const targetPath = targetInput.value.trim();
        if (!targetPath) return false;

        try {
          saveStatus.textContent = 'Exporting...';
          const res = await fetch(`${API_BASE}/api/convert`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              sourcePath: currentDocPath,
              targetPath: targetPath,
              format: ext,
              html: editorCanvas.innerHTML
            })
          });

          if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.error || 'Failed to convert document');
          }

          saveStatus.textContent = 'Exported!';
          if (ext === 'docx' || ext === 'odt') {
            await loadDocumentsList();
            await loadDocument(targetPath);
          } else {
            alert(`Document successfully exported to: ${targetPath}`);
          }
        } catch (err) {
          alert(`Export failed: ${err.message}`);
        }
      }
    });
  }

  async function loadDocumentsList() {
    try {
      const res = await fetch(`${API_BASE}/api/documents`);
      if (!res.ok) return;
      const data = await res.json();
      docSelector.innerHTML = '<option value="">-- Choose document --</option>';
      if (data.files && Array.isArray(data.files)) {
        data.files.forEach(f => {
          const opt = document.createElement('option');
          opt.value = f.path;
          opt.textContent = `[${f.format.toUpperCase()}] ${f.name}`;
          docSelector.appendChild(opt);
        });
      }
      if (currentDocPath) {
        docSelector.value = currentDocPath;
      }
    } catch {}
  }

  // ===========================================================================
  // WYSIWYG Formatting Helpers
  // ===========================================================================
  function formatText(cmd, val = null) {
    editorCanvas.focus();
    document.execCommand(cmd, false, val);
    setDirty(true);
    updateToolbarState();
  }

  function insertLink() {
    const selectedText = window.getSelection().toString();
    const bodyHtml = `
      <div>
        <label for="link-text">Display Text</label>
        <input type="text" id="link-text" value="${selectedText}" style="width: 100%; margin-top: 6px;">
      </div>
      <div>
        <label for="link-url">Target URL</label>
        <input type="text" id="link-url" placeholder="https://example.com" style="width: 100%; margin-top: 6px;">
      </div>
    `;

    openModal({
      title: 'Insert Link',
      bodyHtml: bodyHtml,
      onConfirm: () => {
        const textVal = document.getElementById('link-text').value.trim();
        const urlVal = document.getElementById('link-url').value.trim();
        if (!urlVal) return false;

        editorCanvas.focus();
        if (textVal && textVal !== selectedText) {
          const linkHtml = `<a href="${urlVal}">${textVal}</a>`;
          document.execCommand('insertHTML', false, linkHtml);
        } else {
          document.execCommand('createLink', false, urlVal);
        }
        setDirty(true);
      }
    });
  }

  function insertTable() {
    const bodyHtml = `
      <div style="display: flex; gap: 12px;">
        <div style="flex: 1;">
          <label for="table-rows">Rows</label>
          <input type="number" id="table-rows" value="3" min="1" max="20" style="width: 100%; margin-top: 6px;">
        </div>
        <div style="flex: 1;">
          <label for="table-cols">Columns</label>
          <input type="number" id="table-cols" value="3" min="1" max="10" style="width: 100%; margin-top: 6px;">
        </div>
      </div>
    `;

    openModal({
      title: 'Insert Table',
      bodyHtml: bodyHtml,
      onConfirm: () => {
        const rows = parseInt(document.getElementById('table-rows').value, 10) || 3;
        const cols = parseInt(document.getElementById('table-cols').value, 10) || 3;

        let tableHtml = '<table class="doc-table"><thead><tr>';
        for (let c = 1; c <= cols; c++) {
          tableHtml += `<th>Header ${c}</th>`;
        }
        tableHtml += '</tr></thead><tbody>';
        for (let r = 1; r <= rows; r++) {
          tableHtml += '<tr>';
          for (let c = 1; c <= cols; c++) {
            tableHtml += `<td>Data ${r}-${c}</td>`;
          }
          tableHtml += '</tr>';
        }
        tableHtml += '</tbody></table><p></p>';

        editorCanvas.focus();
        document.execCommand('insertHTML', false, tableHtml);
        setDirty(true);
      }
    });
  }

  function updateToolbarState() {
    try {
      btnBold.classList.toggle('active', document.queryCommandState('bold'));
      btnItalic.classList.toggle('active', document.queryCommandState('italic'));
      btnUnderline.classList.toggle('active', document.queryCommandState('underline'));
      btnStrike.classList.toggle('active', document.queryCommandState('strikeThrough'));
      btnAlignLeft.classList.toggle('active', document.queryCommandState('justifyLeft'));
      btnAlignCenter.classList.toggle('active', document.queryCommandState('justifyCenter'));
      btnAlignRight.classList.toggle('active', document.queryCommandState('justifyRight'));
      btnAlignJustify.classList.toggle('active', document.queryCommandState('justifyFull'));
      btnListUl.classList.toggle('active', document.queryCommandState('insertUnorderedList'));
      btnListOl.classList.toggle('active', document.queryCommandState('insertOrderedList'));
    } catch {}
  }

  // ===========================================================================
  // Event Bindings
  // ===========================================================================
  editorCanvas.addEventListener('input', () => setDirty(true));
  editorCanvas.addEventListener('keyup', updateToolbarState);
  editorCanvas.addEventListener('mouseup', updateToolbarState);

  // Document title rename
  docTitleInput.addEventListener('input', () => {
    currentDocTitle = docTitleInput.value;
    setDirty(true);
  });

  // Action buttons
  btnNew.addEventListener('click', createNewDocument);
  btnSave.addEventListener('click', saveDocument);

  // Export dropdown
  btnSaveAs.addEventListener('click', (e) => {
    e.stopPropagation();
    saveAsMenu.classList.toggle('hidden');
  });

  document.addEventListener('click', () => {
    saveAsMenu.classList.add('hidden');
  });

  saveAsMenu.querySelectorAll('[data-export]').forEach(btn => {
    btn.addEventListener('click', () => {
      const format = btn.getAttribute('data-export');
      convertDocument(format);
      saveAsMenu.classList.add('hidden');
    });
  });

  // Document switcher
  docSelector.addEventListener('change', (e) => {
    const val = e.target.value;
    if (val && val !== currentDocPath) {
      if (isDirty) {
        openModal({
          title: 'Unsaved Changes',
          bodyHtml: '<p>You have unsaved changes in the current document. Switching will discard unsaved edits. Proceed?</p>',
          onConfirm: () => loadDocument(val)
        });
      } else {
        loadDocument(val);
      }
    }
  });

  // Formatting Toolbar Buttons
  btnUndo.addEventListener('click', () => formatText('undo'));
  btnRedo.addEventListener('click', () => formatText('redo'));

  headingSelect.addEventListener('change', (e) => {
    const val = e.target.value;
    if (val === 'p') formatText('formatBlock', '<p>');
    else if (val === 'blockquote') formatText('formatBlock', '<blockquote>');
    else formatText('formatBlock', `<${val}>`);
  });

  fontFamilySelect.addEventListener('change', (e) => {
    formatText('fontName', e.target.value);
  });

  fontSizeSelect.addEventListener('change', (e) => {
    // Map pt to font size command 1-7
    const pt = parseInt(e.target.value, 10);
    let sizeIndex = 3;
    if (pt <= 9) sizeIndex = 1;
    else if (pt <= 10) sizeIndex = 2;
    else if (pt <= 12) sizeIndex = 3;
    else if (pt <= 14) sizeIndex = 4;
    else if (pt <= 18) sizeIndex = 5;
    else if (pt <= 24) sizeIndex = 6;
    else sizeIndex = 7;
    formatText('fontSize', sizeIndex);
  });

  btnBold.addEventListener('click', () => formatText('bold'));
  btnItalic.addEventListener('click', () => formatText('italic'));
  btnUnderline.addEventListener('click', () => formatText('underline'));
  btnStrike.addEventListener('click', () => formatText('strikeThrough'));

  colorPicker.addEventListener('input', (e) => {
    formatText('foreColor', e.target.value);
  });

  highlightPicker.addEventListener('input', (e) => {
    formatText('hiliteColor', e.target.value);
  });

  btnAlignLeft.addEventListener('click', () => formatText('justifyLeft'));
  btnAlignCenter.addEventListener('click', () => formatText('justifyCenter'));
  btnAlignRight.addEventListener('click', () => formatText('justifyRight'));
  btnAlignJustify.addEventListener('click', () => formatText('justifyFull'));

  btnListUl.addEventListener('click', () => formatText('insertUnorderedList'));
  btnListOl.addEventListener('click', () => formatText('insertOrderedList'));
  btnLink.addEventListener('click', insertLink);
  btnTable.addEventListener('click', insertTable);
  btnClearFormat.addEventListener('click', () => formatText('removeFormat'));

  // Global Keyboard Shortcuts
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      saveDocument();
    }
  });

  // ===========================================================================
  // Initialization
  // ===========================================================================
  async function init() {
    await loadDocumentsList();

    const params = new URLSearchParams(window.location.search);
    const requestedFile = params.get('file') || params.get('path');
    if (requestedFile) {
      await loadDocument(requestedFile);
    } else if (docSelector.options.length > 1) {
      await loadDocument(docSelector.options[1].value);
    } else {
      editorCanvas.innerHTML = '<h1>Welcome to Document Editor</h1><p>Open a document from the menu or click <b>New</b> to create one.</p>';
      setDirty(false);
    }
  }

  // Export functions to global scope for testing and debugging
  window.loadDocument = loadDocument;
  window.saveDocument = saveDocument;
  window.convertDocument = convertDocument;
  window.createNewDocument = createNewDocument;
  window.openModal = openModal;
  window.closeModal = closeModal;
  window.insertLink = insertLink;
  window.insertTable = insertTable;
  window.updateToolbarState = updateToolbarState;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
