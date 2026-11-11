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
  let knownDirectories = [];
  let defaultWorkingDir = '';

  // Helpers
  function getPathParts(filePath) {
    if (!filePath) return { dir: '', name: '', base: '', ext: '' };
    const normalized = filePath.replace(/\\/g, '/');
    const lastSlash = normalized.lastIndexOf('/');
    const dir = lastSlash >= 0 ? filePath.substring(0, lastSlash + 1) : '';
    const name = lastSlash >= 0 ? filePath.substring(lastSlash + 1) : filePath;
    const lastDot = name.lastIndexOf('.');
    const base = lastDot >= 0 ? name.substring(0, lastDot) : name;
    const ext = lastDot >= 0 ? name.substring(lastDot + 1).toLowerCase() : '';
    return { dir, name, base, ext };
  }

  // DOM Elements
  const docTitleInput = document.getElementById('doc-title');
  const docFormatBadge = document.getElementById('doc-format-badge');
  const docSelector = document.getElementById('doc-selector');
  const saveStatus = document.getElementById('save-status');
  const btnOpen = document.getElementById('btn-open');
  const fileImportInput = document.getElementById('file-import-input');
  const folderImportInput = document.getElementById('folder-import-input');
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
    modalConfirm.textContent = options.confirmText || 'Confirm';
    if (options.hideCancel) {
      modalCancel.style.display = 'none';
    } else {
      modalCancel.style.display = '';
    }
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
    modalCancel.style.display = '';
    modalConfirm.textContent = 'Confirm';
    modalConfirm.disabled = false;
  }

  function showInfoModal(title, bodyHtml, isError = false) {
    openModal({
      title: title,
      bodyHtml: `<div style="font-size: 13px; line-height: 1.6; color: ${isError ? '#ef4444' : '#e2e8f0'};">${bodyHtml}</div>`,
      confirmText: 'OK',
      hideCancel: true
    });
  }

  modalConfirm.addEventListener('click', async () => {
    if (modalConfirmCallback) {
      modalConfirm.disabled = true;
      try {
        const shouldClose = await modalConfirmCallback();
        if (shouldClose !== false) {
          closeModal();
        }
      } catch (err) {
        console.error('Modal action error:', err);
      } finally {
        modalConfirm.disabled = false;
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
      showInfoModal('Load Error', `<p style="color: #ef4444;">Error loading document: ${err.message}</p>`, true);
      saveStatus.textContent = 'Load error';
      saveStatus.className = 'save-status dirty';
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
      showInfoModal('Save Error', `<p style="color: #ef4444;">Failed to save document: ${err.message}</p>`, true);
      saveStatus.textContent = 'Save failed';
      saveStatus.className = 'save-status dirty';
    } finally {
      isSaving = false;
    }
  }

  function setupInlineFolderBrowser(container, folderInput, updatePreview) {
    const toggleBtn = container.querySelector('.btn-browse-folder');
    const pillsContainer = container.querySelector('.folder-quick-pills');
    const explorerEl = container.querySelector('.folder-explorer-box');
    const subfoldersList = container.querySelector('.subfolders-list');
    const upBtn = container.querySelector('.btn-explorer-up');
    const currentPathLabel = container.querySelector('.explorer-current-path');
    let parentPath = '';

    async function loadDirectory(path) {
      if (!explorerEl || explorerEl.classList.contains('hidden')) return;
      subfoldersList.innerHTML = '<div style="padding: 10px; color: var(--text-muted); font-size: 12px;">Loading folders...</div>';

      try {
        const q = path ? `?path=${encodeURIComponent(path)}` : '';
        const res = await fetch(`${API_BASE}/api/directories${q}`);
        if (!res.ok) return;
        const data = await res.json();
        parentPath = data.parent || '';

        if (currentPathLabel) currentPathLabel.textContent = data.current || path;
        if (upBtn) {
          upBtn.disabled = !data.parent;
          upBtn.style.opacity = data.parent ? '1' : '0.4';
        }

        // Render quick location pills if container empty
        if (pillsContainer && pillsContainer.children.length === 0) {
          const quicks = (data.quickLocations || []).concat(data.drives || []);
          quicks.forEach(q => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn-action dir-pill';
            btn.style.cssText = 'font-size: 11px; padding: 3px 8px; border-radius: 12px; margin-right: 4px; margin-bottom: 4px; cursor: pointer;';
            btn.textContent = (q.name.startsWith('C:') ? '💾 ' : '📁 ') + q.name;
            btn.addEventListener('click', () => {
              const p = q.path.endsWith('\\') ? q.path : q.path + '\\';
              folderInput.value = p;
              updatePreview();
              loadDirectory(q.path);
            });
            pillsContainer.appendChild(btn);
          });
        }

        // Render subfolders
        subfoldersList.innerHTML = '';
        const subs = data.subdirectories || [];
        if (subs.length === 0) {
          subfoldersList.innerHTML = '<div style="padding: 10px; color: var(--text-muted); font-size: 12px;">No subfolders in this directory</div>';
        } else {
          subs.forEach(s => {
            const item = document.createElement('div');
            item.style.cssText = 'display: flex; align-items: center; gap: 8px; padding: 6px 10px; border-radius: 4px; cursor: pointer; transition: background 0.1s ease; user-select: none;';
            item.innerHTML = `<span style="font-size: 14px;">📁</span><span style="font-size: 12px; color: #f8fafc; font-weight: 500;">${s.name}</span>`;
            item.addEventListener('mouseenter', () => item.style.background = '#334155');
            item.addEventListener('mouseleave', () => item.style.background = 'transparent');
            item.addEventListener('click', () => {
              const p = s.path.endsWith('\\') ? s.path : s.path + '\\';
              folderInput.value = p;
              updatePreview();
              loadDirectory(s.path);
            });
            subfoldersList.appendChild(item);
          });
        }
      } catch (err) {
        subfoldersList.innerHTML = `<div style="padding: 10px; color: #ef4444; font-size: 12px;">Error: ${err.message}</div>`;
      }
    }

    if (toggleBtn) {
      toggleBtn.addEventListener('click', () => {
        explorerEl.classList.toggle('hidden');
        if (!explorerEl.classList.contains('hidden')) {
          toggleBtn.textContent = '▲ Close Folders';
          loadDirectory(folderInput.value.trim());
        } else {
          toggleBtn.textContent = '📁 Browse Folders';
        }
      });
    }

    if (upBtn) {
      upBtn.addEventListener('click', () => {
        if (parentPath) {
          const p = parentPath.endsWith('\\') ? parentPath : parentPath + '\\';
          folderInput.value = p;
          updatePreview();
          loadDirectory(parentPath);
        }
      });
    }

    // Pre-populate quick location pills
    fetch(`${API_BASE}/api/directories`)
      .then(r => r.json())
      .then(data => {
        if (pillsContainer && pillsContainer.children.length === 0 && data.quickLocations) {
          const quicks = (data.quickLocations || []).concat(data.drives || []);
          quicks.forEach(q => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn-action dir-pill';
            btn.style.cssText = 'font-size: 11px; padding: 3px 8px; border-radius: 12px; margin-right: 4px; margin-bottom: 4px; cursor: pointer;';
            btn.textContent = (q.name.startsWith('C:') ? '💾 ' : '📁 ') + q.name;
            btn.addEventListener('click', () => {
              const p = q.path.endsWith('\\') ? q.path : q.path + '\\';
              folderInput.value = p;
              updatePreview();
              if (!explorerEl.classList.contains('hidden')) {
                loadDirectory(q.path);
              }
            });
            pillsContainer.appendChild(btn);
          });
        }
      }).catch(() => {});
  }

  function openDocumentModal() {
    // Collect discovered files for quick-opening
    let recentPills = '';
    if (docSelector && docSelector.options.length > 1) {
      const items = [];
      for (let i = 1; i < docSelector.options.length && i <= 8; i++) {
        const opt = docSelector.options[i];
        items.push(`<button type="button" class="btn-action open-recent-pill" data-path="${opt.value}" style="font-size: 11px; padding: 4px 8px; border-radius: 4px; text-align: left; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 100%; cursor: pointer;">${opt.textContent}</button>`);
      }
      if (items.length > 0) {
        recentPills = `
          <div>
            <label style="font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px;">Recent Workspace Documents</label>
            <div style="display: flex; flex-direction: column; gap: 4px; margin-top: 6px; max-height: 140px; overflow-y: auto;">
              ${items.join('')}
            </div>
          </div>
          <div style="text-align: center; color: var(--text-muted); font-size: 11px; letter-spacing: 0.5px;">— OR —</div>
        `;
      }
    }

    const bodyHtml = `
      <div style="display: flex; flex-direction: column; gap: 14px;">
        <div>
          <button type="button" id="modal-btn-browse" class="btn-action primary" style="width: 100%; justify-content: center; padding: 12px; font-size: 14px; font-weight: 600; cursor: pointer;">
            📂 Choose File from Computer (.docx, .odt, .md, .txt)
          </button>
        </div>
        ${recentPills}
        <div>
          <label for="open-doc-path">File Path on Disk</label>
          <input type="text" id="open-doc-path" placeholder="C:\\path\\to\\document.docx" style="width: 100%; margin-top: 6px;">
        </div>
      </div>
    `;

    openModal({
      title: 'Open / Import Document',
      bodyHtml: bodyHtml,
      onConfirm: async () => {
        const pathInput = document.getElementById('open-doc-path');
        const inputPath = pathInput ? pathInput.value.trim() : '';
        if (inputPath) {
          await loadDocument(inputPath);
          return true;
        }
        return false;
      }
    });

    const browseBtn = document.getElementById('modal-btn-browse');
    if (browseBtn) {
      browseBtn.addEventListener('click', () => {
        if (fileImportInput) {
          fileImportInput.value = '';
          fileImportInput.click();
        }
      });
    }

    document.querySelectorAll('.open-recent-pill').forEach(btn => {
      btn.addEventListener('click', async () => {
        const p = btn.getAttribute('data-path');
        if (p) {
          closeModal();
          await loadDocument(p);
        }
      });
    });
  }

  function createNewDocument() {
    const parts = getPathParts(currentDocPath);
    const defaultDir = parts.dir || defaultWorkingDir || '';

    const bodyHtml = `
      <div id="new-dialog-container" style="display: flex; flex-direction: column; gap: 12px;">
        <div>
          <label for="new-doc-folder">Destination Folder</label>
          <div style="display: flex; gap: 6px; margin-top: 6px;">
            <input type="text" id="new-doc-folder" value="${defaultDir}" style="flex: 1;">
            <button type="button" class="btn-action primary btn-browse-folder" style="white-space: nowrap; cursor: pointer;">📁 Browse Folders</button>
          </div>
          <div class="folder-quick-pills" style="margin-top: 8px; display: flex; flex-wrap: wrap;"></div>
          <!-- Folder Explorer -->
          <div class="folder-explorer-box hidden" style="margin-top: 8px; background: #0f172a; border: 1px solid var(--border-color); border-radius: 6px; padding: 8px;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px;">
              <span class="explorer-current-path" style="font-size: 11px; color: var(--text-muted); font-family: monospace; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 320px;"></span>
              <button type="button" class="btn-action btn-explorer-up" style="font-size: 11px; padding: 2px 8px; cursor: pointer;">⬆ Up</button>
            </div>
            <div class="subfolders-list" style="max-height: 140px; overflow-y: auto; display: flex; flex-direction: column; gap: 2px;"></div>
          </div>
        </div>
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
        <div style="font-size: 11px; color: var(--text-muted); word-break: break-all;" id="new-preview-path">
          Full path: ${defaultDir}New Document.${currentDocFormat === 'docx' ? 'docx' : 'odt'}
        </div>
      </div>
    `;

    openModal({
      title: 'Create New Document',
      bodyHtml: bodyHtml,
      confirmText: 'Create',
      onConfirm: async () => {
        const folderInput = document.getElementById('new-doc-folder');
        const nameInput = document.getElementById('new-doc-name');
        const formatSelect = document.getElementById('new-doc-format');
        let folder = folderInput ? folderInput.value.trim() : '';
        const title = (nameInput.value || 'New Document').trim();
        const ext = formatSelect.value;
        const sanitized = title.replace(/[^a-zA-Z0-9_\- ]/g, '').trim() || 'Document';
        const fileName = `${sanitized}.${ext}`;
        if (folder && !folder.endsWith('\\') && !folder.endsWith('/')) folder += '\\';
        const newPath = folder ? folder + fileName : fileName;

        try {
          saveStatus.textContent = 'Creating...';
          saveStatus.className = 'save-status saving';
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
          saveStatus.textContent = 'Created';
          saveStatus.className = 'save-status';
          await loadDocumentsList();
          await loadDocument(newPath);
          return true;
        } catch (err) {
          saveStatus.textContent = 'Creation failed';
          saveStatus.className = 'save-status dirty';
          showInfoModal('Creation Failed', `<p style="color: #ef4444;">Failed to create document: ${err.message}</p>`, true);
          return false;
        }
      }
    });

    const container = document.getElementById('new-dialog-container');
    const folderInput = document.getElementById('new-doc-folder');
    const nameInput = document.getElementById('new-doc-name');
    const formatSelect = document.getElementById('new-doc-format');
    const previewEl = document.getElementById('new-preview-path');
    const updatePreview = () => {
      let f = folderInput.value.trim();
      let n = (nameInput.value || 'Document').trim().replace(/[^a-zA-Z0-9_\- ]/g, '') || 'Document';
      let e = formatSelect.value;
      if (f && !f.endsWith('\\') && !f.endsWith('/')) f += '\\';
      if (previewEl) previewEl.textContent = `Full path: ${f}${n}.${e}`;
    };
    folderInput.addEventListener('input', updatePreview);
    nameInput.addEventListener('input', updatePreview);
    formatSelect.addEventListener('change', updatePreview);

    setupInlineFolderBrowser(container, folderInput, updatePreview);
  }

  function convertDocument(targetFormat) {
    const ext = targetFormat.toLowerCase();
    const parts = getPathParts(currentDocPath);
    const titleClean = (docTitleInput && docTitleInput.value ? docTitleInput.value : currentDocTitle)
      .trim().replace(/[^a-zA-Z0-9_\- ]/g, '');
    const baseName = titleClean || parts.base || 'document';
    const defaultDir = parts.dir || defaultWorkingDir || '';
    const suggestedFileName = `${baseName}_exported.${ext}`;

    const formatLabels = {
      docx: 'Word Document (.docx)',
      odt: 'OpenDocument Text (.odt)',
      md: 'Markdown (.md)',
      html: 'HTML Webpage (.html)',
      txt: 'Plain Text (.txt)'
    };
    const formatLabel = formatLabels[ext] || `.${ext.toUpperCase()}`;

    const bodyHtml = `
      <div id="export-dialog-container" style="display: flex; flex-direction: column; gap: 12px;">
        <div>
          <label for="export-folder-path">Destination Folder</label>
          <div style="display: flex; gap: 6px; margin-top: 6px;">
            <input type="text" id="export-folder-path" value="${defaultDir}" style="flex: 1;">
            <button type="button" class="btn-action primary btn-browse-folder" style="white-space: nowrap; cursor: pointer;">📁 Browse Folders</button>
          </div>
          <div class="folder-quick-pills" style="margin-top: 8px; display: flex; flex-wrap: wrap;"></div>
          <!-- Folder Explorer -->
          <div class="folder-explorer-box hidden" style="margin-top: 8px; background: #0f172a; border: 1px solid var(--border-color); border-radius: 6px; padding: 8px;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px;">
              <span class="explorer-current-path" style="font-size: 11px; color: var(--text-muted); font-family: monospace; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 320px;"></span>
              <button type="button" class="btn-action btn-explorer-up" style="font-size: 11px; padding: 2px 8px; cursor: pointer;">⬆ Up</button>
            </div>
            <div class="subfolders-list" style="max-height: 140px; overflow-y: auto; display: flex; flex-direction: column; gap: 2px;"></div>
          </div>
        </div>
        <div>
          <label for="export-file-name">File Name</label>
          <input type="text" id="export-file-name" value="${suggestedFileName}" style="width: 100%; margin-top: 6px;">
        </div>
        <div style="font-size: 11px; color: var(--text-muted); word-break: break-all;" id="export-preview-path">
          Full path: ${defaultDir}${suggestedFileName}
        </div>
      </div>
    `;

    openModal({
      title: `Export as ${formatLabel}`,
      bodyHtml: bodyHtml,
      confirmText: 'Export',
      onConfirm: async () => {
        const folderInput = document.getElementById('export-folder-path');
        const fileInput = document.getElementById('export-file-name');
        let folder = (folderInput ? folderInput.value.trim() : '');
        let fileName = (fileInput ? fileInput.value.trim() : '');
        if (!fileName) {
          showInfoModal('Missing File Name', '<p style="color: #ef4444;">Please provide a file name for the export.</p>', true);
          return false;
        }
        if (!fileName.toLowerCase().endsWith('.' + ext)) fileName += '.' + ext;
        if (folder && !folder.endsWith('\\') && !folder.endsWith('/')) folder += '\\';
        const targetPath = folder ? folder + fileName : fileName;

        try {
          saveStatus.textContent = 'Exporting...';
          saveStatus.className = 'save-status saving';
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
            throw new Error(err.error || `Server responded with ${res.status}`);
          }

          saveStatus.textContent = 'Exported!';
          saveStatus.className = 'save-status';
          setTimeout(() => {
            if (!isDirty) saveStatus.textContent = 'Saved';
          }, 3000);

          if (ext === 'docx' || ext === 'odt') {
            await loadDocumentsList();
            await loadDocument(targetPath);
          } else {
            showInfoModal(
              'Export Successful',
              `<div style="display: flex; flex-direction: column; gap: 8px;">
                <p>Document exported successfully as <strong>${formatLabel}</strong>.</p>
                <div style="background: #0f172a; padding: 8px 10px; border-radius: 4px; font-family: monospace; font-size: 11px; color: #38bdf8; word-break: break-all;">
                  ${targetPath}
                </div>
              </div>`,
              false
            );
          }
          return true;
        } catch (err) {
          saveStatus.textContent = 'Export failed';
          saveStatus.className = 'save-status dirty';
          showInfoModal(
            'Export Failed',
            `<p style="color: #ef4444;">Failed to export document: ${err.message}</p>`,
            true
          );
          return false;
        }
      }
    });

    const container = document.getElementById('export-dialog-container');
    const folderInput = document.getElementById('export-folder-path');
    const fileInput = document.getElementById('export-file-name');
    const previewEl = document.getElementById('export-preview-path');
    const updatePreview = () => {
      let f = folderInput.value.trim();
      let n = fileInput.value.trim();
      if (f && !f.endsWith('\\') && !f.endsWith('/')) f += '\\';
      if (previewEl) previewEl.textContent = `Full path: ${f}${n}`;
    };
    folderInput.addEventListener('input', updatePreview);
    fileInput.addEventListener('input', updatePreview);

    setupInlineFolderBrowser(container, folderInput, updatePreview);
  }

  async function loadDocumentsList() {
    try {
      const res = await fetch(`${API_BASE}/api/documents`);
      if (!res.ok) return;
      const data = await res.json();
      docSelector.innerHTML = '<option value="">-- Choose document --</option>';
      if (data.files && Array.isArray(data.files)) {
        const dirSet = new Set();
        data.files.forEach(f => {
          if (f.path) {
            const lastSlash = f.path.lastIndexOf('\\');
            if (lastSlash > 0) dirSet.add(f.path.substring(0, lastSlash + 1));
          }
          const opt = document.createElement('option');
          opt.value = f.path;
          opt.textContent = `[${f.format.toUpperCase()}] ${f.name}`;
          docSelector.appendChild(opt);
        });
        knownDirectories = Array.from(dirSet);
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
  if (btnOpen) btnOpen.addEventListener('click', openDocumentModal);
  btnNew.addEventListener('click', createNewDocument);
  btnSave.addEventListener('click', saveDocument);

  // File import input
  if (fileImportInput) {
    fileImportInput.addEventListener('change', async (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;

      closeModal();

      saveStatus.textContent = 'Importing...';
      saveStatus.className = 'save-status saving';

      // 1. Direct path if present (Electron)
      if (file.path && (file.path.toLowerCase().endsWith('.docx') || file.path.toLowerCase().endsWith('.odt'))) {
        await loadDocument(file.path);
        return;
      }

      // 2. Markdown or plain text import
      const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
      if (ext === '.md' || ext === '.markdown' || ext === '.txt') {
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const text = reader.result;
            const lines = text.split(/\r?\n\r?\n/);
            const html = lines.map(line => {
              const trimmed = line.trim();
              if (trimmed.startsWith('# ')) return `<h1>${trimmed.slice(2)}</h1>`;
              if (trimmed.startsWith('## ')) return `<h2>${trimmed.slice(3)}</h2>`;
              if (trimmed.startsWith('### ')) return `<h3>${trimmed.slice(4)}</h3>`;
              if (trimmed.startsWith('> ')) return `<blockquote>${trimmed.slice(2)}</blockquote>`;
              return `<p>${trimmed.replace(/\n/g, '<br/>')}</p>`;
            }).join('\n');

            editorCanvas.innerHTML = html;
            currentDocTitle = file.name.replace(/\.[^/.]+$/, '');
            docTitleInput.value = currentDocTitle;
            currentDocPath = file.path || '';
            updateFormatBadge(ext.replace('.', ''));
            setDirty(true);
            saveStatus.textContent = 'Loaded';
            saveStatus.className = 'save-status';
          } catch (err) {
            saveStatus.textContent = 'Import failed';
            saveStatus.className = 'save-status dirty';
            openModal({ title: 'Import Failed', bodyHtml: `<p style="color: #ef4444;">${err.message}</p>` });
          }
        };
        reader.readAsText(file);
        return;
      }

      // 3. Binary .docx or .odt via native FileReader.readAsDataURL (fast, non-blocking)
      const reader = new FileReader();
      reader.onload = async () => {
        try {
          const dataUrl = reader.result;
          const base64 = dataUrl.indexOf(',') >= 0 ? dataUrl.split(',')[1] : dataUrl;

          const res = await fetch(`${API_BASE}/api/import`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename: file.name, base64: base64 })
          });

          if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.error || `Server responded with status ${res.status}`);
          }

          const doc = await res.json();
          currentDocPath = doc.path;
          currentDocTitle = doc.title || file.name;
          docTitleInput.value = currentDocTitle;
          updateFormatBadge(doc.format);
          editorCanvas.innerHTML = doc.html || '<p></p>';
          setDirty(false);
          saveStatus.textContent = 'Loaded';
          saveStatus.className = 'save-status';
          await loadDocumentsList();
        } catch (err) {
          saveStatus.textContent = 'Import failed';
          saveStatus.className = 'save-status dirty';
          openModal({
            title: 'Import Failed',
            bodyHtml: `<p style="color: #ef4444;">Failed to import document: ${err.message}</p>`
          });
        }
      };
      reader.onerror = () => {
        saveStatus.textContent = 'Read failed';
        saveStatus.className = 'save-status dirty';
        openModal({
          title: 'Read Error',
          bodyHtml: '<p style="color: #ef4444;">Could not read the selected file from disk.</p>'
        });
      };
      reader.readAsDataURL(file);
    });
  }

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
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
      e.preventDefault();
      openDocumentModal();
    }
  });

  // ===========================================================================
  // Initialization
  // ===========================================================================
  async function init() {
    fetch(`${API_BASE}/api/directories`)
      .then(r => r.json())
      .then(d => {
        if (d && d.current) {
          defaultWorkingDir = d.current.endsWith('\\') ? d.current : d.current + '\\';
        }
      }).catch(() => {});

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
  window.openDocumentModal = openDocumentModal;
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
