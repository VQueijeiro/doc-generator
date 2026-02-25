// ========== STATE ==========
let manuals = [];
let currentManualId = null;
let templates = { manual: null, desarrollo: null }; // persistent per type
let deleteTargetId = null;
let quillInstances = {};

const MANUALS_KEY = 'docgen_manuals';
const CURRENT_KEY = 'docgen_current';
const TEMPLATES_KEY = 'docgen_templates';
const DOC_TYPES = {
    manual: { label: 'Manual de Uso', icon: '📘', defaultTitle: 'Manual de Usuario' },
    desarrollo: { label: 'Doc. de Desarrollo', icon: '⚙️', defaultTitle: 'Documento de Desarrollo' }
};

const $ = (s) => document.querySelector(s);
const sectionsContainer = $('#sectionsContainer');
const emptyState = $('#emptyState');
const sectionCount = $('#sectionCount');
const docTitle = $('#docTitle');

// ========== INIT ==========
document.addEventListener('DOMContentLoaded', () => {
    // Initialize theme
    const savedTheme = localStorage.getItem('docgen_theme');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
        document.documentElement.setAttribute('data-theme', 'dark');
    }

    loadState();
    if (manuals.length === 0) createNewManual('manual', false);
    loadManual(currentManualId || manuals[0].id);
    renderSidebar();
    renderSections();
    setupEventListeners();
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
});

// ========== EVENT LISTENERS ==========
function setupEventListeners() {
    $('#sidebarToggle').addEventListener('click', toggleSidebar);
    $('#sidebarClose').addEventListener('click', closeSidebar);
    $('#sidebarOverlay').addEventListener('click', closeSidebar);
    $('#newManualBtn').addEventListener('click', () => showModal('newDocModal'));
    $('#closeNewDocModal').addEventListener('click', () => hideModal('newDocModal'));

    // Doc type selection
    document.querySelectorAll('.doc-type-card').forEach(card => {
        card.addEventListener('click', () => {
            const type = card.dataset.type;
            hideModal('newDocModal');
            createNewManual(type, true);
            closeSidebar();
        });
    });

    // Theme toggle
    const themeToggle = $('#themeToggle');
    const themeIcon = $('#themeIcon');
    if (document.documentElement.getAttribute('data-theme') === 'dark') {
        themeToggle.classList.add('active');
        themeIcon.textContent = '🌙';
    }
    themeToggle.addEventListener('click', () => {
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        if (isDark) {
            document.documentElement.removeAttribute('data-theme');
            themeToggle.classList.remove('active');
            themeIcon.textContent = '☀️';
            localStorage.setItem('docgen_theme', 'light');
        } else {
            document.documentElement.setAttribute('data-theme', 'dark');
            themeToggle.classList.add('active');
            themeIcon.textContent = '🌙';
            localStorage.setItem('docgen_theme', 'dark');
        }
    });

    $('#addSectionBtn').addEventListener('click', addSection);
    $('#settingsBtn').addEventListener('click', openTemplateModal);
    $('#saveTemplate').addEventListener('click', saveTemplate);
    $('#closeTemplateModal').addEventListener('click', () => hideModal('templateModal'));
    $('#previewBtn').addEventListener('click', showPreview);
    $('#exportWordBtn').addEventListener('click', exportToWord);
    $('#exportWord').addEventListener('click', exportToWord);
    $('#closePreview').addEventListener('click', () => hideModal('previewModal'));
    $('#copyMarkdown').addEventListener('click', copyMarkdown);
    $('#confirmDelete').addEventListener('click', confirmDeleteManual);
    $('#cancelDelete').addEventListener('click', () => hideModal('deleteModal'));

    docTitle.addEventListener('input', () => {
        const manual = getCurrentManual();
        if (manual) { manual.title = docTitle.value; manual.updatedAt = Date.now(); saveAll(); renderSidebar(); }
    });

    // Logo upload
    $('#logoUpload').addEventListener('click', (e) => {
        if (e.target.id === 'removeLogo') return;
        const input = document.createElement('input');
        input.type = 'file'; input.accept = 'image/*';
        input.addEventListener('change', () => { if (input.files[0]) loadLogo(input.files[0]); });
        input.click();
    });
    $('#removeLogo').addEventListener('click', (e) => {
        e.stopPropagation();
        const tmpl = getTemplateForCurrent();
        tmpl.logo = null;
        $('#logoPreview').style.display = 'none';
        $('#logoPlaceholder').style.display = '';
        $('#removeLogo').style.display = 'none';
    });

    // Global paste
    document.addEventListener('paste', (e) => {
        const el = document.activeElement;
        if (el.closest('.ql-editor') || el.classList.contains('section-title-input') ||
            el.id === 'docTitle' || el.id === 'companyName' ||
            el.id === 'docSubtitle' || el.id === 'footerText') return;
        for (const item of e.clipboardData.items) {
            if (item.type.startsWith('image/')) {
                e.preventDefault();
                const sections = getCurrentSections();
                if (sections.length === 0 || sections[sections.length - 1].image) addSection();
                handleImageFile(item.getAsFile(), getCurrentSections().length - 1);
                break;
            }
        }
    });
}

// ========== SIDEBAR ==========
function toggleSidebar() { $('#sidebar').classList.toggle('open'); $('#sidebarOverlay').classList.toggle('active'); }
function closeSidebar() { $('#sidebar').classList.remove('open'); $('#sidebarOverlay').classList.remove('active'); }

function renderSidebar() {
    const list = $('#sidebarList');
    const sorted = [...manuals].sort((a, b) => b.updatedAt - a.updatedAt);
    list.innerHTML = sorted.map(m => {
        const isActive = m.id === currentManualId;
        const date = new Date(m.updatedAt).toLocaleDateString('es-AR', { day: '2-digit', month: 'short' });
        const type = DOC_TYPES[m.docType] || DOC_TYPES.manual;
        return `<div class="sidebar-item ${isActive ? 'active' : ''}" data-id="${m.id}">
            <div class="sidebar-item-info">
                <span class="sidebar-item-title">${escapeHtml(m.title || 'Sin título')}</span>
                <span class="sidebar-item-sections">
                    <span class="sidebar-item-type ${m.docType || 'manual'}">${type.label}</span> · ${m.sections.length} sec · ${date}
                </span>
            </div>
            <button class="sidebar-item-delete" data-delete="${m.id}" title="Eliminar">🗑</button>
        </div>`;
    }).join('');

    list.querySelectorAll('.sidebar-item').forEach(item => {
        item.addEventListener('click', (e) => {
            if (e.target.closest('.sidebar-item-delete')) return;
            if (item.dataset.id !== currentManualId) {
                loadManual(item.dataset.id); renderSections(); renderSidebar(); closeSidebar();
            }
        });
    });
    list.querySelectorAll('.sidebar-item-delete').forEach(btn => {
        btn.addEventListener('click', (e) => { e.stopPropagation(); requestDeleteManual(btn.dataset.delete); });
    });
}

// ========== MANUAL CRUD ==========
function createNewManual(docType, switchTo) {
    const type = DOC_TYPES[docType] || DOC_TYPES.manual;
    const manual = {
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title: type.defaultTitle,
        docType: docType,
        sections: [],
        createdAt: Date.now(),
        updatedAt: Date.now()
    };
    manuals.push(manual);
    if (switchTo) { loadManual(manual.id); renderSections(); }
    saveAll(); renderSidebar();
    if (switchTo) toast('Nuevo documento creado');
}

function loadManual(id) {
    quillInstances = {};
    currentManualId = id;
    const manual = getCurrentManual();
    if (manual) {
        docTitle.value = manual.title;
        updateDocTypeBadge();
    }
    localStorage.setItem(CURRENT_KEY, id);
}

function getCurrentManual() { return manuals.find(m => m.id === currentManualId); }
function getCurrentSections() { const m = getCurrentManual(); return m ? m.sections : []; }

function updateDocTypeBadge() {
    const manual = getCurrentManual();
    const badge = $('#docTypeBadge');
    if (manual && DOC_TYPES[manual.docType]) {
        const type = DOC_TYPES[manual.docType];
        badge.textContent = type.label;
        badge.className = `doc-type-badge ${manual.docType}`;
    } else {
        badge.textContent = '';
        badge.className = 'doc-type-badge';
    }
}

function requestDeleteManual(id) {
    const m = manuals.find(m => m.id === id);
    if (!m) return;
    deleteTargetId = id;
    $('#deleteModalText').textContent = `¿Eliminar "${m.title}"?`;
    showModal('deleteModal');
}

function confirmDeleteManual() {
    if (!deleteTargetId) return;
    manuals = manuals.filter(m => m.id !== deleteTargetId);
    if (manuals.length === 0) createNewManual('manual', false);
    if (currentManualId === deleteTargetId) { loadManual(manuals[0].id); renderSections(); }
    deleteTargetId = null;
    saveAll(); renderSidebar(); hideModal('deleteModal'); toast('Documento eliminado');
}

// ========== TEMPLATE (per doc type) ==========
function getDefaultTemplate() {
    return { logo: null, companyName: '', authorName: '', version: '1', subtitle: '', footerText: '' };
}

function getTemplateForCurrent() {
    const manual = getCurrentManual();
    const type = manual ? manual.docType : 'manual';
    if (!templates[type]) templates[type] = getDefaultTemplate();
    return templates[type];
}

function loadLogo(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const tmpl = getTemplateForCurrent();
        tmpl.logo = e.target.result;
        $('#logoPreview').src = e.target.result;
        $('#logoPreview').style.display = '';
        $('#logoPlaceholder').style.display = 'none';
        $('#removeLogo').style.display = 'flex';
    };
    reader.readAsDataURL(file);
}

function openTemplateModal() {
    const tmpl = getTemplateForCurrent();
    $('#companyName').value = tmpl.companyName || '';
    $('#authorName').value = tmpl.authorName || '';
    $('#docVersion').value = tmpl.version || '1';
    $('#docSubtitle').value = tmpl.subtitle || '';
    $('#footerText').value = tmpl.footerText || '';
    if (tmpl.logo) {
        $('#logoPreview').src = tmpl.logo;
        $('#logoPreview').style.display = '';
        $('#logoPlaceholder').style.display = 'none';
        $('#removeLogo').style.display = 'flex';
    } else {
        $('#logoPreview').style.display = 'none';
        $('#logoPlaceholder').style.display = '';
        $('#removeLogo').style.display = 'none';
    }
    showModal('templateModal');
}

function saveTemplate() {
    const tmpl = getTemplateForCurrent();
    tmpl.companyName = $('#companyName').value.trim();
    tmpl.authorName = $('#authorName').value.trim();
    tmpl.version = $('#docVersion').value.trim() || '1';
    tmpl.subtitle = $('#docSubtitle').value.trim();
    tmpl.footerText = $('#footerText').value.trim();
    localStorage.setItem(TEMPLATES_KEY, JSON.stringify(templates));
    hideModal('templateModal');
    toast('Template guardado para este tipo de documento');
}

// ========== SECTIONS CRUD ==========
function addSection() {
    const manual = getCurrentManual();
    if (!manual) return;
    manual.sections.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title: '', image: null, imageType: null, html: '', text: ''
    });
    manual.updatedAt = Date.now();
    saveAll(); renderSections(); renderSidebar();
    setTimeout(() => {
        const cards = sectionsContainer.querySelectorAll('.section-card');
        if (cards.length) cards[cards.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
}

function removeSection(idx) {
    const manual = getCurrentManual();
    if (!manual) return;
    const editorId = `editor-${manual.sections[idx].id}`;
    delete quillInstances[editorId];
    manual.sections.splice(idx, 1);
    manual.updatedAt = Date.now();
    saveAll(); renderSections(); renderSidebar();
}

function moveSection(idx, dir) {
    const sections = getCurrentSections();
    const newIdx = idx + dir;
    if (newIdx < 0 || newIdx >= sections.length) return;
    [sections[idx], sections[newIdx]] = [sections[newIdx], sections[idx]];
    getCurrentManual().updatedAt = Date.now();
    quillInstances = {};
    saveAll(); renderSections();
}

// ========== RENDER SECTIONS ==========
function renderSections() {
    quillInstances = {};
    const sections = getCurrentSections();
    sectionsContainer.querySelectorAll('.section-card').forEach(c => c.remove());

    if (sections.length === 0) {
        emptyState.style.display = '';
    } else {
        emptyState.style.display = 'none';
        sections.forEach((s, i) => {
            sectionsContainer.appendChild(createSectionCard(s, i));
            initQuill(s, i);
        });
    }
    sectionCount.textContent = sections.length + (sections.length === 1 ? ' sección' : ' secciones');
}

function createSectionCard(section, idx) {
    const card = document.createElement('div');
    card.className = 'section-card';
    card.draggable = true;
    card.dataset.idx = idx;
    const editorId = `editor-${section.id}`;

    card.innerHTML = `
        <div class="section-header">
            <div class="section-header-left">
                <span class="drag-handle" title="Arrastrar para reordenar">⠿</span>
                <span class="section-number">Sección ${idx + 1}</span>
            </div>
            <div class="section-header-right">
                <button class="btn-icon move-up" title="Mover arriba">↑</button>
                <button class="btn-icon move-down" title="Mover abajo">↓</button>
                <button class="btn-icon btn-danger delete-section" title="Eliminar">🗑</button>
            </div>
        </div>
        <div class="section-body">
            <div class="section-left">
                <label class="context-label">Título de la sección</label>
                <input type="text" class="section-title-input" placeholder="Ej: Inicio de sesión" value="${escapeAttr(section.title)}">
                <div class="image-drop-zone ${section.image ? 'has-image' : ''}">
                    ${section.image
                        ? `<img src="${section.image}" alt="Captura"><button class="remove-image" title="Quitar imagen">✕</button>`
                        : `<span class="placeholder-icon">🖼</span><span class="placeholder-text">Arrastrá una imagen aquí,<br>hacé clic, o pegá con Ctrl+V</span>`
                    }
                </div>
            </div>
            <div class="section-right">
                <label class="context-label">Texto de la sección</label>
                <div class="editor-wrapper">
                    <div id="${editorId}"></div>
                </div>
            </div>
        </div>
    `;

    card.querySelector('.move-up').addEventListener('click', () => moveSection(idx, -1));
    card.querySelector('.move-down').addEventListener('click', () => moveSection(idx, 1));
    card.querySelector('.delete-section').addEventListener('click', () => removeSection(idx));

    const dropZone = card.querySelector('.image-drop-zone');
    setupDropZone(dropZone, idx);

    const removeBtn = card.querySelector('.remove-image');
    if (removeBtn) {
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            getCurrentSections()[idx].image = null;
            getCurrentSections()[idx].imageType = null;
            getCurrentManual().updatedAt = Date.now();
            saveAll(); renderSections();
        });
    }

    card.querySelector('.section-title-input').addEventListener('input', (e) => {
        getCurrentSections()[idx].title = e.target.value;
        getCurrentManual().updatedAt = Date.now();
        saveAll();
    });

    // Drag reorder
    card.addEventListener('dragstart', (e) => { card.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', idx.toString()); });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
    card.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; card.classList.add('drag-over'); });
    card.addEventListener('dragleave', () => card.classList.remove('drag-over'));
    card.addEventListener('drop', (e) => {
        e.preventDefault(); card.classList.remove('drag-over');
        const fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
        if (fromIdx !== idx && !isNaN(fromIdx)) {
            const sections = getCurrentSections();
            const [moved] = sections.splice(fromIdx, 1);
            sections.splice(idx, 0, moved);
            getCurrentManual().updatedAt = Date.now();
            saveAll(); renderSections();
        }
    });

    return card;
}

// ========== QUILL EDITOR ==========
function initQuill(section, idx) {
    const editorId = `editor-${section.id}`;
    const container = document.getElementById(editorId);
    if (!container) return;

    const quill = new Quill(container, {
        theme: 'snow',
        placeholder: 'Escribí o pegá el texto de esta sección...',
        modules: {
            toolbar: [
                [{ 'header': [1, 2, 3, false] }],
                ['bold', 'italic', 'underline', 'strike'],
                [{ 'color': [] }, { 'background': [] }],
                [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                [{ 'align': [] }],
                ['clean']
            ]
        }
    });

    // Load existing content
    if (section.html) {
        quill.root.innerHTML = section.html;
    }

    // Save on change (debounced)
    let saveTimer;
    quill.on('text-change', () => {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(() => {
            const sections = getCurrentSections();
            if (sections[idx]) {
                sections[idx].html = quill.root.innerHTML;
                sections[idx].text = quill.getText();
                getCurrentManual().updatedAt = Date.now();
                saveAll();
            }
        }, 500);
    });

    quillInstances[editorId] = quill;
}

// ========== IMAGE HANDLING ==========
function setupDropZone(zone, idx) {
    zone.addEventListener('click', () => {
        if (getCurrentSections()[idx].image) return;
        const input = document.createElement('input');
        input.type = 'file'; input.accept = 'image/*';
        input.addEventListener('change', () => { if (input.files[0]) handleImageFile(input.files[0], idx); });
        input.click();
    });
    zone.addEventListener('dragover', (e) => { e.preventDefault(); if (e.dataTransfer.types.includes('Files')) zone.classList.add('drag-hover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-hover'));
    zone.addEventListener('drop', (e) => {
        zone.classList.remove('drag-hover');
        if (e.dataTransfer.files.length && e.dataTransfer.files[0].type.startsWith('image/')) {
            e.preventDefault(); e.stopPropagation(); handleImageFile(e.dataTransfer.files[0], idx);
        }
    });
}

function handleImageFile(file, idx) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const sections = getCurrentSections();
        sections[idx].image = e.target.result;
        sections[idx].imageType = file.type;
        getCurrentManual().updatedAt = Date.now();
        saveAll(); renderSections(); renderSidebar();
    };
    reader.readAsDataURL(file);
}

// ========== WORD EXPORT ==========
async function exportToWord() {
    const sections = getCurrentSections();
    if (sections.length === 0) { toast('Agregá secciones primero'); return; }
    if (!window.docx) { toast('Error: la librería docx no se cargó. Recargá la página.'); return; }

    toast('Generando documento Word...');

    try {
        const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel,
                Header, Footer, PageNumber, AlignmentType,
                BorderStyle, PageBreak, Bookmark, InternalHyperlink,
                TableOfContents, Table, TableRow, TableCell, WidthType, VerticalAlign } = window.docx;

        const tmpl = getTemplateForCurrent();
        const headerChildren = [];
        if (tmpl.logo) {
            try {
                const logoData = await dataUrlToArrayBuffer(tmpl.logo);
                headerChildren.push(new Paragraph({ children: [new ImageRun({ data: logoData, transformation: { width: 120, height: 40 }, type: 'png' })], alignment: AlignmentType.LEFT }));
            } catch (e) { console.warn('Logo error:', e); }
        }
        if (tmpl.companyName) {
            headerChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.companyName, bold: true, size: 18, color: '666666' })], alignment: AlignmentType.LEFT }));
        }

        const footerChildren = [];
        if (tmpl.footerText) {
            footerChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.footerText, size: 16, color: '999999', italics: true })], alignment: AlignmentType.LEFT }));
        }
        footerChildren.push(new Paragraph({
            children: [
                new TextRun({ text: 'Página ', size: 16, color: '999999' }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '999999' }),
                new TextRun({ text: ' de ', size: 16, color: '999999' }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: '999999' })
            ], alignment: AlignmentType.RIGHT
        }));

        const docChildren = [];

        // Title page - header table (estilo portada con logo, título, versión, autor)
        const manual = getCurrentManual();
        const creationDate = new Date(manual.createdAt).toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const docVersion = tmpl.version || '1';
        const authorName = tmpl.authorName || '';

        const logoCellChildren = [];
        if (tmpl.logo) {
            try {
                const logoDataForTable = await dataUrlToArrayBuffer(tmpl.logo);
                logoCellChildren.push(new Paragraph({ children: [new ImageRun({ data: logoDataForTable, transformation: { width: 100, height: 35 }, type: 'png' })], alignment: AlignmentType.CENTER }));
            } catch(e) { console.warn('Logo table error:', e); }
        }
        if (tmpl.companyName) {
            logoCellChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.companyName, bold: true, size: 22, color: '002366' })], alignment: AlignmentType.CENTER, spacing: { before: 40 } }));
        }
        if (logoCellChildren.length === 0) {
            logoCellChildren.push(new Paragraph({ children: [new TextRun({ text: ' ' })] }));
        }

        const cBorder = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
        const cBorders = { top: cBorder, bottom: cBorder, left: cBorder, right: cBorder };

        const titleTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            rowSpan: 2,
                            width: { size: 30, type: WidthType.PERCENTAGE },
                            children: logoCellChildren,
                            verticalAlign: VerticalAlign.CENTER,
                            borders: cBorders,
                        }),
                        new TableCell({
                            columnSpan: 2,
                            children: [new Paragraph({ children: [new TextRun({ text: (docTitle.value || 'Documento').toUpperCase(), bold: true, size: 24, color: '002366' })], alignment: AlignmentType.LEFT })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: cBorders,
                        }),
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: `Versión ${docVersion}`, size: 20, color: '333333' })] })],
                            borders: cBorders,
                        }),
                        new TableCell({
                            children: [
                                new Paragraph({ children: [new TextRun({ text: authorName || '', size: 20, color: '333333' })] }),
                                new Paragraph({ children: [new TextRun({ text: `(${creationDate})`, size: 18, color: '666666' })] }),
                            ],
                            borders: cBorders,
                        }),
                    ]
                }),
            ]
        });

        docChildren.push(titleTable);
        docChildren.push(new Paragraph({ spacing: { before: 2000 } }));
        docChildren.push(new Paragraph({ children: [new TextRun({ text: docTitle.value || 'Documento', bold: true, size: 56, color: '002366' })], alignment: AlignmentType.CENTER }));
        if (tmpl.subtitle) docChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.subtitle, size: 28, color: '666666' })], alignment: AlignmentType.CENTER, spacing: { before: 200 } }));
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // TOC
        docChildren.push(new Paragraph({ children: [new TextRun({ text: 'Índice de Contenidos', bold: true, size: 32, color: '002366' })], spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: '002366' } } }));
        docChildren.push(new TableOfContents("Índice", { hyperlink: true, headingStyleRange: "1-2" }));
        sections.forEach((s, i) => {
            const title = s.title || `Sección ${i + 1}`;
            docChildren.push(new Paragraph({ children: [new InternalHyperlink({ anchor: `section_${i}`, children: [new TextRun({ text: `${i + 1}. ${title}`, style: 'Hyperlink', size: 22 })] })], spacing: { before: 100, after: 40 } }));
        });
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // Sections
        for (let i = 0; i < sections.length; i++) {
            const s = sections[i];
            const title = s.title || `Sección ${i + 1}`;
            docChildren.push(new Paragraph({ children: [new Bookmark({ id: `section_${i}`, children: [new TextRun({ text: `${i + 1}. ${title}`, bold: true, size: 28, color: '002366' })] })], heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: '002366' } } }));

            if (s.image) {
                try {
                    const imgData = await dataUrlToArrayBuffer(s.image);
                    const dims = await getImageDimensions(s.image);
                    const maxW = 500; let w = dims.width, h = dims.height;
                    if (w > maxW) { h = Math.round(h * (maxW / w)); w = maxW; }
                    docChildren.push(new Paragraph({ children: [new ImageRun({ data: imgData, transformation: { width: w, height: h }, type: s.imageType?.includes('png') ? 'png' : 'jpg' })], alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 } }));
                } catch (e) { console.warn('Image error section', i, e); }
            }

            // Parse HTML to docx paragraphs
            const textContent = s.text || '';
            if (textContent.trim()) {
                const htmlContent = s.html || '';
                const paragraphs = htmlToDocxParagraphs(htmlContent, TextRun, Paragraph);
                docChildren.push(...paragraphs);
            }

            if (i < sections.length - 1) docChildren.push(new Paragraph({ children: [new PageBreak()] }));
        }

        const doc = new Document({
            features: { updateFields: true },
            styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
            sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, pageNumbers: { start: 1 } } }, headers: headerChildren.length > 0 ? { default: new Header({ children: headerChildren }) } : undefined, footers: { default: new Footer({ children: footerChildren }) }, children: docChildren }]
        });

        const blob = await Packer.toBlob(doc);
        const fileName = (docTitle.value || 'documento').replace(/[^a-zA-Z0-9áéíóúñÁÉÍÓÚÑ\s-]/g, '').replace(/\s+/g, '_');
        window.saveAs(blob, `${fileName}.docx`);
        toast('Documento Word descargado');
    } catch (e) {
        console.error('Export error:', e);
        toast('Error al exportar: ' + e.message);
    }
}

// Convert Quill HTML to docx paragraphs
function htmlToDocxParagraphs(html, TextRun, Paragraph) {
    const paragraphs = [];
    const tmp = document.createElement('div');
    tmp.innerHTML = html;

    const blocks = tmp.querySelectorAll('p, h1, h2, h3, li, div');
    if (blocks.length === 0) {
        // Fallback: plain text
        const text = tmp.textContent || '';
        for (const line of text.split('\n')) {
            if (line.trim()) paragraphs.push(new Paragraph({ children: [new TextRun({ text: line, size: 22, font: 'Calibri' })], spacing: { before: 60, after: 60 } }));
        }
        return paragraphs;
    }

    blocks.forEach(block => {
        const runs = [];
        processNode(block, runs, TextRun);
        if (runs.length === 0) return;

        const opts = { children: runs, spacing: { before: 60, after: 60 } };
        const tag = block.tagName.toLowerCase();
        if (tag === 'h1') { opts.spacing = { before: 200, after: 100 }; }
        if (tag === 'h2') { opts.spacing = { before: 160, after: 80 }; }
        if (tag === 'h3') { opts.spacing = { before: 120, after: 60 }; }
        if (tag === 'li') {
            const parent = block.parentElement;
            if (parent && parent.tagName === 'OL') {
                opts.numbering = { reference: 'default-numbering', level: 0 };
            } else {
                opts.bullet = { level: 0 };
            }
        }

        paragraphs.push(new Paragraph(opts));
    });

    return paragraphs.length > 0 ? paragraphs : [new Paragraph({ children: [new TextRun({ text: tmp.textContent || '', size: 22, font: 'Calibri' })], spacing: { before: 60, after: 60 } })];
}

function processNode(node, runs, TextRun) {
    node.childNodes.forEach(child => {
        if (child.nodeType === 3) { // Text node
            const text = child.textContent;
            if (text) {
                const styles = getInheritedStyles(child);
                runs.push(new TextRun({ text, size: styles.size || 22, font: 'Calibri', bold: styles.bold, italics: styles.italic, underline: styles.underline ? {} : undefined, strike: styles.strike }));
            }
        } else if (child.nodeType === 1) { // Element
            processNode(child, runs, TextRun);
        }
    });
}

function getInheritedStyles(node) {
    const styles = { bold: false, italic: false, underline: false, strike: false, size: 22 };
    let el = node.parentElement;
    while (el) {
        const tag = el.tagName.toLowerCase();
        if (tag === 'strong' || tag === 'b') styles.bold = true;
        if (tag === 'em' || tag === 'i') styles.italic = true;
        if (tag === 'u') styles.underline = true;
        if (tag === 's' || tag === 'del' || tag === 'strike') styles.strike = true;
        if (tag === 'h1') { styles.bold = true; styles.size = 32; }
        if (tag === 'h2') { styles.bold = true; styles.size = 28; }
        if (tag === 'h3') { styles.bold = true; styles.size = 24; }
        if (el.classList.contains('ql-editor') || el.classList.contains('section-right')) break;
        el = el.parentElement;
    }
    return styles;
}

// ========== PREVIEW & MARKDOWN ==========
function showPreview() {
    const sections = getCurrentSections();
    if (sections.length === 0) { toast('Agregá secciones primero'); return; }
    const content = $('#previewContent');
    let html = `<h1>${escapeHtml(docTitle.value || 'Documento')}</h1>`;
    sections.forEach((s, i) => {
        const title = s.title || `Sección ${i + 1}`;
        html += `<div class="preview-section"><h2>${i + 1}. ${escapeHtml(title)}</h2>`;
        if (s.image) html += `<img src="${s.image}" alt="${escapeHtml(title)}">`;
        html += s.html ? `<div class="preview-text">${s.html}</div>` : `<div class="preview-text" style="color:var(--text-muted)">(Sin texto todavía)</div>`;
        html += `</div>`;
    });
    content.innerHTML = html;
    showModal('previewModal');
}

function copyMarkdown() {
    const sections = getCurrentSections();
    let md = `# ${docTitle.value || 'Documento'}\n\n`;
    sections.forEach((s, i) => {
        md += `## ${i + 1}. ${s.title || `Sección ${i + 1}`}\n\n`;
        if (s.text) md += s.text + '\n\n';
        md += '---\n\n';
    });
    navigator.clipboard.writeText(md).then(() => toast('Markdown copiado'));
}

// ========== PERSISTENCE ==========
function saveAll() {
    try {
        localStorage.setItem(MANUALS_KEY, JSON.stringify(manuals));
        localStorage.setItem(CURRENT_KEY, currentManualId);
    } catch (e) {
        console.warn('localStorage full');
        toast('Almacenamiento lleno. Exportá y eliminá documentos.');
    }
}

function loadState() {
    try {
        const saved = localStorage.getItem(MANUALS_KEY);
        if (saved) manuals = JSON.parse(saved);

        // Migrate old format
        if (manuals.length === 0) {
            const old = localStorage.getItem('docgen_sections');
            const oldTitle = localStorage.getItem('docgen_title');
            if (old) {
                const sections = JSON.parse(old);
                if (sections.length > 0) {
                    // Convert old generated field to html/text
                    sections.forEach(s => {
                        if (s.generated && !s.html) {
                            s.html = `<p>${escapeHtml(s.generated)}</p>`;
                            s.text = s.generated;
                        }
                    });
                    manuals.push({ id: 'migrated', title: oldTitle || 'Manual de Usuario', docType: 'manual', sections, createdAt: Date.now(), updatedAt: Date.now() });
                    localStorage.removeItem('docgen_sections');
                    localStorage.removeItem('docgen_title');
                }
            }
        }

        // Migrate manuals without docType
        manuals.forEach(m => { if (!m.docType) m.docType = 'manual'; });

        currentManualId = localStorage.getItem(CURRENT_KEY);
        const tmpl = localStorage.getItem(TEMPLATES_KEY);
        if (tmpl) templates = JSON.parse(tmpl);
        // Migrate old single template
        const oldTmpl = localStorage.getItem('docgen_template');
        if (oldTmpl && !templates.manual) {
            templates.manual = JSON.parse(oldTmpl);
            localStorage.removeItem('docgen_template');
            localStorage.setItem(TEMPLATES_KEY, JSON.stringify(templates));
        }
    } catch (e) {
        manuals = [];
    }
}

// ========== UTILS ==========
function showModal(id) { document.getElementById(id).classList.add('active'); }
function hideModal(id) { document.getElementById(id).classList.remove('active'); }
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 3000); }
function escapeHtml(str) { return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>'); }
function escapeAttr(str) { return (str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function dataUrlToArrayBuffer(dataUrl) {
    return new Promise((resolve) => { const b = atob(dataUrl.split(',')[1]); const u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); resolve(u.buffer); });
}
function getImageDimensions(dataUrl) {
    return new Promise((resolve) => { const img = new Image(); img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight }); img.onerror = () => resolve({ width: 400, height: 300 }); img.src = dataUrl; });
}
