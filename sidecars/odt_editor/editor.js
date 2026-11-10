// Antigravity ODT WYSIWYG Editor Client
(() => {
  let currentFilePath = null;
  let isDirty = false;

  const docTitleEl = document.getElementById('doc-title');
  const docSelectorEl = document.getElementById('doc-selector');
  const editorCanvasEl = document.getElementById('editor-canvas');
  const saveStatusEl = document.getElementById('save-status');
  const footerPathEl = document.getElementById('footer-path');
  const btnSaveEl = document.getElementById('btn-save');
  const toastEl = document.getElementById('toast');

  const headingSelect = document.getElementById('heading-select');
  const fontFamilySelect = document.getElementById('font-family-select');
  const fontSizeSelect = document.getElementById('font-size-select');
  const colorPicker = document.getElementById('color-picker');
  const highlightPicker = document.getElementById('highlight-picker');

  function normalizePath(p) {
    if (!p) return '';
    return p.replace(/\\/g, '/').toLowerCase();
  }

  function showToast(message, isError = false) {
    toastEl.textContent = message;
    toastEl.style.backgroundColor = isError ? '#991b1b' : '#065f46';
    toastEl.classList.add('show');
    setTimeout(() => {
      toastEl.classList.remove('show');
    }, 2500);
  }

  function setDirty(state) {
    isDirty = state;
    if (state) {
      saveStatusEl.textContent = '• Unsaved changes';
      saveStatusEl.classList.add('dirty');
    } else {
      saveStatusEl.textContent = 'Saved';
      saveStatusEl.classList.remove('dirty');
    }
  }

  function execCommand(cmd, val = null) {
    document.execCommand(cmd, false, val);
    editorCanvasEl.focus();
    setDirty(true);
  }

  function formatText(cmd, val) {
    execCommand(cmd, val);
  }

  async function loadDocument(filePath) {
    try {
      saveStatusEl.textContent = 'Loading...';
      const encoded = encodeURIComponent(filePath);
      const resp = await fetch(`/api/document?path=${encoded}`);
      if (!resp.ok) {
        throw new Error(`Failed to load: HTTP ${resp.status}`);
      }
      const data = await resp.json();
      if (data.error) {
        throw new Error(data.error);
      }

      currentFilePath = filePath;
      docTitleEl.textContent = data.title || 'Document';
      editorCanvasEl.innerHTML = data.html || '<p></p>';
      footerPathEl.textContent = filePath;
      setDirty(false);

      // Match in selector dropdown
      const normalizedCurrent = normalizePath(filePath);
      for (let i = 0; i < docSelectorEl.options.length; i++) {
        if (normalizePath(docSelectorEl.options[i].value) === normalizedCurrent) {
          docSelectorEl.selectedIndex = i;
          break;
        }
      }
    } catch (err) {
      console.error('Error loading document:', err);
      showToast(`Error: ${err.message}`, true);
      saveStatusEl.textContent = 'Error loading';
    }
  }

  async function saveDocument() {
    if (!currentFilePath) {
      showToast('No active document to save', true);
      return;
    }

    try {
      btnSaveEl.disabled = true;
      btnSaveEl.textContent = 'Saving...';
      const htmlContent = editorCanvasEl.innerHTML;

      const resp = await fetch('/api/document', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          path: currentFilePath,
          html: htmlContent
        })
      });

      const res = await resp.json();
      if (!resp.ok || !res.success) {
        throw new Error(res.error || 'Failed to save document');
      }

      setDirty(false);
      showToast('Document saved successfully to disk!');
    } catch (err) {
      console.error('Save failed:', err);
      showToast(`Save failed: ${err.message}`, true);
    } finally {
      btnSaveEl.disabled = false;
      btnSaveEl.innerHTML = '<span class="btn-icon">💾</span> Save';
    }
  }

  async function fetchDocumentsList() {
    try {
      const resp = await fetch('/api/documents');
      if (!resp.ok) return [];
      const data = await resp.json();
      const files = Array.isArray(data.files) ? data.files : [];

      docSelectorEl.innerHTML = '';
      if (files.length === 0) {
        const opt = document.createElement('option');
        opt.textContent = 'No .odt files found';
        docSelectorEl.appendChild(opt);
        return [];
      }

      files.forEach(f => {
        const opt = document.createElement('option');
        opt.value = f.path;
        opt.textContent = f.name;
        if (currentFilePath && normalizePath(f.path) === normalizePath(currentFilePath)) {
          opt.selected = true;
        }
        docSelectorEl.appendChild(opt);
      });

      return files;
    } catch (e) {
      console.error('Failed to list documents:', e);
      return [];
    }
  }

  // Event Listeners for Toolbar
  document.querySelectorAll('.btn-tool[data-cmd]').forEach(btn => {
    btn.addEventListener('click', () => {
      const cmd = btn.getAttribute('data-cmd');
      execCommand(cmd);
    });
  });

  headingSelect.addEventListener('change', () => {
    const val = headingSelect.value;
    if (val === 'p') {
      execCommand('formatBlock', '<p>');
    } else {
      execCommand('formatBlock', `<${val}>`);
    }
  });

  fontFamilySelect.addEventListener('change', () => {
    execCommand('fontName', fontFamilySelect.value);
  });

  fontSizeSelect.addEventListener('change', () => {
    execCommand('fontSize', fontSizeSelect.value);
  });

  colorPicker.addEventListener('input', () => {
    execCommand('foreColor', colorPicker.value);
  });

  highlightPicker.addEventListener('input', () => {
    execCommand('hiliteColor', highlightPicker.value);
  });

  btnSaveEl.addEventListener('click', saveDocument);

  docSelectorEl.addEventListener('change', () => {
    const selected = docSelectorEl.value;
    if (selected && selected !== currentFilePath) {
      if (isDirty) {
        if (!confirm('You have unsaved changes. Discard and switch document?')) {
          docSelectorEl.value = currentFilePath;
          return;
        }
      }
      loadDocument(selected);
    }
  });

  editorCanvasEl.addEventListener('input', () => {
    setDirty(true);
  });

  // Keyboard Shortcuts: Ctrl+S to save
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      saveDocument();
    }
  });

  // Main initialization routine
  async function init() {
    const urlParams = new URLSearchParams(window.location.search);
    const fileParam = urlParams.get('file');

    const files = await fetchDocumentsList();

    if (fileParam) {
      await loadDocument(fileParam);
    } else if (files.length > 0) {
      await loadDocument(files[0].path);
    }
  }

  init();
})();
