// ========== MULTI-MANUAL STATE ==========
let manuals = [];       // Array of { id, title, sections, createdAt, updatedAt }
let currentManualId = null;
let templateConfig = { logo: null, companyName: '', subtitle: '', footerText: '' };
let deleteTargetId = null;

const MANUALS_KEY = 'docgen_manuals';
const CURRENT_KEY = 'docgen_current';
const TEMPLATE_KEY = 'docgen_template';

const $ = (s) => document.querySelector(s);
const sectionsContainer = $('#sectionsContainer');
const emptyState = $('#emptyState');
const sectionCount = $('#sectionCount');
const docTitle = $('#docTitle');

// ========== INIT ==========
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    if (manuals.length === 0) createNewManual(false);
    loadManual(currentManualId || manuals[0].id);
    renderSidebar();
    renderSections();
    setupEventListeners();
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
});

// ========== EVENT LISTENERS ==========
function setupEventListeners() {
    // Sidebar
    $('#sidebarToggle').addEventListener('click', toggleSidebar);
    $('#sidebarClose').addEventListener('click', closeSidebar);
    $('#sidebarOverlay').addEventListener('click', closeSidebar);
    $('#newManualBtn').addEventListener('click', () => { createNewManual(true); closeSidebar(); });

    // Header
    $('#addSectionBtn').addEventListener('click', addSection);
    $('#settingsBtn').addEventListener('click', openTemplateModal);
    $('#saveTemplate').addEventListener('click', saveTemplate);
    $('#closeTemplateModal').addEventListener('click', () => hideModal('templateModal'));
    $('#previewBtn').addEventListener('click', showPreview);
    $('#exportWordBtn').addEventListener('click', exportToWord);
    $('#exportWord').addEventListener('click', exportToWord);
    $('#closePreview').addEventListener('click', () => hideModal('previewModal'));
    $('#copyMarkdown').addEventListener('click', copyMarkdown);

    // Delete modal
    $('#confirmDelete').addEventListener('click', confirmDeleteManual);
    $('#cancelDelete').addEventListener('click', () => hideModal('deleteModal'));

    // Title change
    docTitle.addEventListener('input', () => {
        const manual = getCurrentManual();
        if (manual) {
            manual.title = docTitle.value;
            manual.updatedAt = Date.now();
            saveAll();
            renderSidebar();
        }
    });

    // Logo upload
    $('#logoUpload').addEventListener('click', (e) => {
        if (e.target.id === 'removeLogo') return;
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.addEventListener('change', () => { if (input.files[0]) loadLogo(input.files[0]); });
        input.click();
    });
    $('#removeLogo').addEventListener('click', (e) => {
        e.stopPropagation();
        templateConfig.logo = null;
        $('#logoPreview').style.display = 'none';
        $('#logoPlaceholder').style.display = '';
        $('#removeLogo').style.display = 'none';
    });

    // Template file
    $('#templateUploadZone').addEventListener('click', () => $('#templateFileInput').click());
    $('#templateFileInput').addEventListener('change', (e) => {
        if (e.target.files[0]) {
            $('#templateFileName').textContent = e.target.files[0].name;
            $('#templateUploadZone').classList.add('has-file');
        }
    });

    // Global paste
    document.addEventListener('paste', (e) => {
        const el = document.activeElement;
        if (el.classList.contains('context-input') || el.classList.contains('generated-text') ||
            el.classList.contains('section-title-input') || el.id === 'docTitle' ||
            el.id === 'companyName' || el.id === 'docSubtitle' || el.id === 'footerText') return;
        const items = e.clipboardData.items;
        for (const item of items) {
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
function toggleSidebar() {
    $('#sidebar').classList.toggle('open');
    $('#sidebarOverlay').classList.toggle('active');
}
function closeSidebar() {
    $('#sidebar').classList.remove('open');
    $('#sidebarOverlay').classList.remove('active');
}

function renderSidebar() {
    const list = $('#sidebarList');
    // Sort by updatedAt descending
    const sorted = [...manuals].sort((a, b) => b.updatedAt - a.updatedAt);
    list.innerHTML = sorted.map(m => {
        const isActive = m.id === currentManualId;
        const date = new Date(m.updatedAt).toLocaleDateString('es-AR', { day: '2-digit', month: 'short' });
        const secCount = m.sections.length;
        return `
            <div class="sidebar-item ${isActive ? 'active' : ''}" data-id="${m.id}">
                <div class="sidebar-item-info">
                    <span class="sidebar-item-title">${escapeHtml(m.title || 'Sin título')}</span>
                    <span class="sidebar-item-sections">${secCount} sección${secCount !== 1 ? 'es' : ''} · ${date}</span>
                </div>
                <button class="sidebar-item-delete" data-delete="${m.id}" title="Eliminar">🗑</button>
            </div>
        `;
    }).join('');

    // Click handlers
    list.querySelectorAll('.sidebar-item').forEach(item => {
        item.addEventListener('click', (e) => {
            if (e.target.closest('.sidebar-item-delete')) return;
            const id = item.dataset.id;
            if (id !== currentManualId) {
                loadManual(id);
                renderSections();
                renderSidebar();
                closeSidebar();
            }
        });
    });
    list.querySelectorAll('.sidebar-item-delete').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            requestDeleteManual(btn.dataset.delete);
        });
    });
}

// ========== MANUAL CRUD ==========
function createNewManual(switchTo) {
    const manual = {
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title: 'Manual de Usuario',
        sections: [],
        createdAt: Date.now(),
        updatedAt: Date.now()
    };
    manuals.push(manual);
    if (switchTo) {
        loadManual(manual.id);
        renderSections();
    }
    saveAll();
    renderSidebar();
    if (switchTo) toast('Nuevo manual creado');
}

function loadManual(id) {
    currentManualId = id;
    const manual = getCurrentManual();
    if (manual) {
        docTitle.value = manual.title;
    }
    localStorage.setItem(CURRENT_KEY, id);
}

function getCurrentManual() {
    return manuals.find(m => m.id === currentManualId);
}

function getCurrentSections() {
    const manual = getCurrentManual();
    return manual ? manual.sections : [];
}

function requestDeleteManual(id) {
    const manual = manuals.find(m => m.id === id);
    if (!manual) return;
    deleteTargetId = id;
    $('#deleteModalText').textContent = `¿Estás seguro de que querés eliminar "${manual.title}"?`;
    showModal('deleteModal');
}

function confirmDeleteManual() {
    if (!deleteTargetId) return;
    manuals = manuals.filter(m => m.id !== deleteTargetId);
    if (manuals.length === 0) createNewManual(false);
    if (currentManualId === deleteTargetId) {
        loadManual(manuals[0].id);
        renderSections();
    }
    deleteTargetId = null;
    saveAll();
    renderSidebar();
    hideModal('deleteModal');
    toast('Manual eliminado');
}

// ========== SECTIONS CRUD ==========
function addSection() {
    const manual = getCurrentManual();
    if (!manual) return;
    manual.sections.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title: '', image: null, imageType: null, context: '', generated: ''
    });
    manual.updatedAt = Date.now();
    saveAll();
    renderSections();
    renderSidebar();
    setTimeout(() => {
        const cards = sectionsContainer.querySelectorAll('.section-card');
        if (cards.length) cards[cards.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 50);
}

function removeSection(idx) {
    const manual = getCurrentManual();
    if (!manual) return;
    manual.sections.splice(idx, 1);
    manual.updatedAt = Date.now();
    saveAll();
    renderSections();
    renderSidebar();
}

function moveSection(idx, dir) {
    const sections = getCurrentSections();
    const newIdx = idx + dir;
    if (newIdx < 0 || newIdx >= sections.length) return;
    [sections[idx], sections[newIdx]] = [sections[newIdx], sections[idx]];
    getCurrentManual().updatedAt = Date.now();
    saveAll();
    renderSections();
}

// ========== RENDER SECTIONS ==========
function renderSections() {
    const sections = getCurrentSections();
    const fragment = document.createDocumentFragment();

    if (sections.length === 0) {
        emptyState.style.display = '';
    } else {
        emptyState.style.display = 'none';
        sections.forEach((s, i) => fragment.appendChild(createSectionCard(s, i)));
    }

    sectionsContainer.querySelectorAll('.section-card').forEach(c => c.remove());
    sectionsContainer.appendChild(fragment);
    sectionCount.textContent = sections.length + (sections.length === 1 ? ' sección' : ' secciones');
}

function createSectionCard(section, idx) {
    const card = document.createElement('div');
    card.className = 'section-card';
    card.draggable = true;
    card.dataset.idx = idx;

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
                <input type="text" class="section-title-input" placeholder="Ej: Inicio de sesión" data-idx="${idx}" value="${escapeAttr(section.title)}">
                <div class="image-drop-zone ${section.image ? 'has-image' : ''}" data-idx="${idx}">
                    ${section.image
                        ? `<img src="${section.image}" alt="Captura"><button class="remove-image" title="Quitar imagen">✕</button>`
                        : `<span class="placeholder-icon">🖼</span><span class="placeholder-text">Arrastrá una imagen aquí,<br>hacé clic, o pegá con Ctrl+V</span>`
                    }
                </div>
            </div>
            <div class="section-right">
                <div class="generated-label">
                    <label class="context-label">Texto de la sección</label>
                </div>
                <textarea class="generated-text" placeholder="Pegá acá el texto generado o escribí manualmente el contenido de esta sección del manual." data-idx="${idx}">${section.generated}</textarea>
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
            const sections = getCurrentSections();
            sections[idx].image = null;
            sections[idx].imageType = null;
            getCurrentManual().updatedAt = Date.now();
            saveAll();
            renderSections();
        });
    }

    card.querySelector('.section-title-input').addEventListener('input', (e) => {
        getCurrentSections()[idx].title = e.target.value;
        getCurrentManual().updatedAt = Date.now();
        saveAll();
    });

    card.querySelector('.generated-text').addEventListener('input', (e) => {
        getCurrentSections()[idx].generated = e.target.value;
        getCurrentManual().updatedAt = Date.now();
        saveAll();
    });

    dropZone.addEventListener('paste', (e) => {
        for (const item of e.clipboardData.items) {
            if (item.type.startsWith('image/')) { e.preventDefault(); handleImageFile(item.getAsFile(), idx); break; }
        }
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

// ========== TEMPLATE ==========
function loadLogo(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        templateConfig.logo = e.target.result;
        $('#logoPreview').src = e.target.result;
        $('#logoPreview').style.display = '';
        $('#logoPlaceholder').style.display = 'none';
        $('#removeLogo').style.display = 'flex';
    };
    reader.readAsDataURL(file);
}

function openTemplateModal() {
    $('#companyName').value = templateConfig.companyName;
    $('#docSubtitle').value = templateConfig.subtitle;
    $('#footerText').value = templateConfig.footerText;
    if (templateConfig.logo) {
        $('#logoPreview').src = templateConfig.logo;
        $('#logoPreview').style.display = '';
        $('#logoPlaceholder').style.display = 'none';
        $('#removeLogo').style.display = 'flex';
    }
    showModal('templateModal');
}

function saveTemplate() {
    templateConfig.companyName = $('#companyName').value.trim();
    templateConfig.subtitle = $('#docSubtitle').value.trim();
    templateConfig.footerText = $('#footerText').value.trim();
    localStorage.setItem(TEMPLATE_KEY, JSON.stringify(templateConfig));
    hideModal('templateModal');
    toast('Template guardado');
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
                TableOfContents } = window.docx;

        const headerChildren = [];
        if (templateConfig.logo) {
            try {
                const logoData = await dataUrlToArrayBuffer(templateConfig.logo);
                headerChildren.push(new Paragraph({ children: [new ImageRun({ data: logoData, transformation: { width: 120, height: 40 }, type: 'png' })], alignment: AlignmentType.LEFT }));
            } catch (e) { console.warn('Logo error:', e); }
        }
        if (templateConfig.companyName) {
            headerChildren.push(new Paragraph({ children: [new TextRun({ text: templateConfig.companyName, bold: true, size: 18, color: '666666' })], alignment: AlignmentType.LEFT }));
        }

        const footerChildren = [];
        if (templateConfig.footerText) {
            footerChildren.push(new Paragraph({ children: [new TextRun({ text: templateConfig.footerText, size: 16, color: '999999', italics: true })], alignment: AlignmentType.LEFT }));
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

        // Title page
        docChildren.push(new Paragraph({ spacing: { before: 3000 } }));
        docChildren.push(new Paragraph({ children: [new TextRun({ text: docTitle.value || 'Manual de Usuario', bold: true, size: 56, color: '2B2B7B' })], alignment: AlignmentType.CENTER }));
        if (templateConfig.subtitle) docChildren.push(new Paragraph({ children: [new TextRun({ text: templateConfig.subtitle, size: 28, color: '666666' })], alignment: AlignmentType.CENTER, spacing: { before: 200 } }));
        if (templateConfig.companyName) docChildren.push(new Paragraph({ children: [new TextRun({ text: templateConfig.companyName, size: 24, color: '999999' })], alignment: AlignmentType.CENTER, spacing: { before: 400 } }));
        docChildren.push(new Paragraph({ children: [new TextRun({ text: new Date().toLocaleDateString('es-AR', { year: 'numeric', month: 'long', day: 'numeric' }), size: 20, color: '999999' })], alignment: AlignmentType.CENTER, spacing: { before: 200 } }));
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // TOC
        docChildren.push(new Paragraph({ children: [new TextRun({ text: 'Índice de Contenidos', bold: true, size: 32, color: '2B2B7B' })], spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: '2B2B7B' } } }));
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
            docChildren.push(new Paragraph({ children: [new Bookmark({ id: `section_${i}`, children: [new TextRun({ text: `${i + 1}. ${title}`, bold: true, size: 28, color: '2B2B7B' })] })], heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: '2B2B7B' } } }));

            if (s.image) {
                try {
                    const imgData = await dataUrlToArrayBuffer(s.image);
                    const dims = await getImageDimensions(s.image);
                    const maxW = 500; let w = dims.width, h = dims.height;
                    if (w > maxW) { h = Math.round(h * (maxW / w)); w = maxW; }
                    docChildren.push(new Paragraph({ children: [new ImageRun({ data: imgData, transformation: { width: w, height: h }, type: s.imageType?.includes('png') ? 'png' : 'jpg' })], alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 } }));
                } catch (e) { console.warn('Image error section', i, e); }
            }

            if (s.generated) {
                for (const line of s.generated.split('\n')) {
                    docChildren.push(new Paragraph({ children: [new TextRun({ text: line, size: 22, font: 'Calibri' })], spacing: { before: 60, after: 60 } }));
                }
            }
            if (i < sections.length - 1) docChildren.push(new Paragraph({ children: [new PageBreak()] }));
        }

        const doc = new Document({
            features: { updateFields: true },
            styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
            sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, pageNumbers: { start: 1 } } }, headers: headerChildren.length > 0 ? { default: new Header({ children: headerChildren }) } : undefined, footers: { default: new Footer({ children: footerChildren }) }, children: docChildren }]
        });

        const blob = await Packer.toBlob(doc);
        const fileName = (docTitle.value || 'manual').replace(/[^a-zA-Z0-9áéíóúñÁÉÍÓÚÑ\s-]/g, '').replace(/\s+/g, '_');
        window.saveAs(blob, `${fileName}.docx`);
        toast('Documento Word descargado');
    } catch (e) {
        console.error('Export error:', e);
        toast('Error al exportar: ' + e.message);
    }
}

// ========== PREVIEW & EXPORT ==========
function showPreview() {
    const sections = getCurrentSections();
    if (sections.length === 0) { toast('Agregá secciones primero'); return; }
    const content = $('#previewContent');
    let html = `<h1>${escapeHtml(docTitle.value || 'Manual de Usuario')}</h1>`;
    sections.forEach((s, i) => {
        const title = s.title || `Sección ${i + 1}`;
        html += `<div class="preview-section"><h2>${i + 1}. ${escapeHtml(title)}</h2>`;
        if (s.image) html += `<img src="${s.image}" alt="${escapeHtml(title)}">`;
        html += s.generated ? `<div class="preview-text">${escapeHtml(s.generated)}</div>` : `<div class="preview-text" style="color:var(--text-muted)">(Sin texto todavía)</div>`;
        html += `</div>`;
    });
    content.innerHTML = html;
    showModal('previewModal');
}

function copyMarkdown() {
    const sections = getCurrentSections();
    let md = `# ${docTitle.value || 'Manual de Usuario'}\n\n`;
    sections.forEach((s, i) => {
        md += `## ${i + 1}. ${s.title || `Sección ${i + 1}`}\n\n`;
        if (s.generated) md += s.generated + '\n\n';
        md += '---\n\n';
    });
    navigator.clipboard.writeText(md).then(() => toast('Markdown copiado al portapapeles'));
}

// ========== PERSISTENCE ==========
function saveAll() {
    try {
        localStorage.setItem(MANUALS_KEY, JSON.stringify(manuals));
        localStorage.setItem(CURRENT_KEY, currentManualId);
    } catch (e) {
        console.warn('localStorage full');
        toast('Almacenamiento lleno. Exportá y eliminá manuales.');
    }
}

function loadState() {
    try {
        const saved = localStorage.getItem(MANUALS_KEY);
        if (saved) manuals = JSON.parse(saved);

        // Migrate from old format
        if (manuals.length === 0) {
            const oldSections = localStorage.getItem('docgen_sections');
            const oldTitle = localStorage.getItem('docgen_title');
            if (oldSections) {
                const sections = JSON.parse(oldSections);
                if (sections.length > 0) {
                    manuals.push({
                        id: 'migrated',
                        title: oldTitle || 'Manual de Usuario',
                        sections: sections,
                        createdAt: Date.now(),
                        updatedAt: Date.now()
                    });
                    localStorage.removeItem('docgen_sections');
                    localStorage.removeItem('docgen_title');
                }
            }
        }

        currentManualId = localStorage.getItem(CURRENT_KEY);
        const tmpl = localStorage.getItem(TEMPLATE_KEY);
        if (tmpl) templateConfig = JSON.parse(tmpl);
    } catch (e) {
        manuals = [];
    }
}

// ========== MODALS & UTILS ==========
function showModal(id) { document.getElementById(id).classList.add('active'); }
function hideModal(id) { document.getElementById(id).classList.remove('active'); }

function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 3000);
}

function escapeHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}
function escapeAttr(str) {
    return (str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function dataUrlToArrayBuffer(dataUrl) {
    return new Promise((resolve) => {
        const binary = atob(dataUrl.split(',')[1]);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        resolve(bytes.buffer);
    });
}
function getImageDimensions(dataUrl) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
        img.onerror = () => resolve({ width: 400, height: 300 });
        img.src = dataUrl;
    });
}
