/* global Quill */
// ========== CONSTANTS & STATE ==========
const MANUALS_KEY  = 'docgen_manuals';
const CURRENT_KEY  = 'docgen_current';
const TEMPLATES_KEY = 'docgen_templates';
const CLIENTS_KEY  = 'docgen_clients';
const THEME_KEY    = 'docgen_theme';

const DOC_TYPES = {
    manual:      { label: 'Manual de Uso',      icon: '📘', defaultTitle: 'Manual de Usuario' },
    desarrollo:  { label: 'Doc. de Desarrollo', icon: '⚙️', defaultTitle: 'Documento de Desarrollo' },
    minuta:      { label: 'Minuta de Reunión',  icon: '📋', defaultTitle: 'Minuta de Reunión' },
    cronograma:  { label: 'Cronograma',         icon: '📅', defaultTitle: 'Cronograma de Proyecto' }
};

let manuals          = [];
let currentManualId  = null;
let templates        = {};
let deleteTargetId   = null;
const quillInstances = {}; // editorId → Quill instance (never reassigned, keys deleted on cleanup)
let _saveTimer       = null;

const $  = (s) => document.querySelector(s);
const sectionsContainer = $('#sectionsContainer');
const emptyState        = $('#emptyState');
const sectionCount      = $('#sectionCount');
const docTitle          = $('#docTitle');

// ========== INDEXEDDB ==========
const DB_NAME     = 'docgen-db';
const DB_VERSION  = 1;
const IMG_STORE   = 'images';
let _db = null;

function openDB() {
    if (_db) return Promise.resolve(_db);
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = (e) => e.target.result.createObjectStore(IMG_STORE);
        req.onsuccess  = (e) => { _db = e.target.result; resolve(_db); };
        req.onerror    = (e) => reject(e.target.error);
    });
}

async function dbSave(key, value) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(IMG_STORE, 'readwrite');
        tx.objectStore(IMG_STORE).put(value, key);
        tx.oncomplete = resolve;
        tx.onerror    = (e) => reject(e.target.error);
    });
}

async function dbGet(key) {
    const db = await openDB();
    return new Promise((resolve) => {
        const req = db.transaction(IMG_STORE, 'readonly').objectStore(IMG_STORE).get(key);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror   = () => resolve(null);
    });
}

async function dbDelete(key) {
    if (!key) return;
    const db = await openDB();
    return new Promise((resolve) => {
        const tx = db.transaction(IMG_STORE, 'readwrite');
        tx.objectStore(IMG_STORE).delete(key);
        tx.oncomplete = resolve;
        tx.onerror    = resolve;
    });
}

// ========== INIT ==========
document.addEventListener('DOMContentLoaded', async () => {
    // Apply saved theme before render
    const savedTheme = localStorage.getItem(THEME_KEY);
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    if (savedTheme === 'dark' || (!savedTheme && prefersDark)) {
        document.documentElement.setAttribute('data-theme', 'dark');
    }

    await openDB().catch(() => console.warn('IndexedDB unavailable'));
    loadState();
    await migrateImagesToIndexedDB();
    if (manuals.length === 0) createNewManual('manual', false);
    loadManual(currentManualId || manuals[0].id);
    renderSidebar();
    renderSections();
    setupEventListeners();
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
});

// ========== MIGRATION: base64 in localStorage → IndexedDB ==========
async function migrateImagesToIndexedDB() {
    let dirty = false;
    for (const manual of manuals) {
        for (const s of manual.sections) {
            if (s.image && !s.imageKey) {
                const key = `${manual.id}_${s.id}`;
                try {
                    await dbSave(key, s.image);
                    s.imageKey = key;
                    delete s.image;
                    dirty = true;
                } catch (e) {
                    console.warn('Migration failed for section', s.id, e);
                }
            }
        }
    }
    if (dirty) saveMeta();
}

// ========== EVENT LISTENERS ==========
function setupEventListeners() {
    // Sidebar
    $('#sidebarToggle').addEventListener('click', toggleSidebar);
    $('#sidebarClose').addEventListener('click', closeSidebar);
    $('#sidebarOverlay').addEventListener('click', closeSidebar);
    $('#newManualBtn').addEventListener('click', () => showModal('newDocModal'));
    $('#closeNewDocModal').addEventListener('click', () => hideModal('newDocModal'));
    document.querySelectorAll('.doc-type-card').forEach(card => {
        card.addEventListener('click', () => {
            hideModal('newDocModal');
            createNewManual(card.dataset.type, true);
            closeSidebar();
        });
    });

    // Theme toggle
    const themeToggle = $('#themeToggle');
    const themeIcon   = $('#themeIcon');
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
            localStorage.setItem(THEME_KEY, 'light');
        } else {
            document.documentElement.setAttribute('data-theme', 'dark');
            themeToggle.classList.add('active');
            themeIcon.textContent = '🌙';
            localStorage.setItem(THEME_KEY, 'dark');
        }
    });

    // Sections – delegated click + drag on container
    $('#addSectionBtn').addEventListener('click', addSection);
    sectionsContainer.addEventListener('click',    handleContainerClick);
    sectionsContainer.addEventListener('dragover',  handleContainerDragOver);
    sectionsContainer.addEventListener('dragleave', handleContainerDragLeave);
    sectionsContainer.addEventListener('drop',      handleContainerDrop);

    // Template modal
    $('#settingsBtn').addEventListener('click', openTemplateModal);
    $('#saveTemplate').addEventListener('click', saveTemplate);
    $('#closeTemplateModal').addEventListener('click', () => hideModal('templateModal'));

    // Logo
    $('#logoUpload').addEventListener('click', (e) => {
        if (e.target.id === 'removeLogo') return;
        const input = document.createElement('input');
        input.type = 'file'; input.accept = 'image/*';
        input.addEventListener('change', () => { if (input.files[0]) loadLogo(input.files[0]); });
        input.click();
    });
    $('#removeLogo').addEventListener('click', (e) => {
        e.stopPropagation();
        getTemplateForCurrent().logo = null;
        $('#logoPreview').style.display = 'none';
        $('#logoPlaceholder').style.display = '';
        $('#removeLogo').style.display = 'none';
    });

    // Preview & export
    $('#previewBtn').addEventListener('click', showPreview);
    $('#exportWordBtn').addEventListener('click', exportToWord);
    $('#exportWord').addEventListener('click', exportToWord);
    $('#closePreview').addEventListener('click', () => hideModal('previewModal'));
    $('#copyMarkdown').addEventListener('click', copyMarkdown);

    // Delete confirm
    $('#confirmDelete').addEventListener('click', confirmDeleteManual);
    $('#cancelDelete').addEventListener('click', () => hideModal('deleteModal'));

    // Doc title
    docTitle.addEventListener('input', () => {
        const manual = getCurrentManual();
        if (!manual) return;
        manual.title = docTitle.value;
        manual.updatedAt = Date.now();
        scheduleSave();
        renderSidebar();
        if (manual.docType === 'minuta') updateMinutaSubject();
    });

    // Minuta panel
    $('#minutaClientSelect').addEventListener('change', (e) => {
        const manual = getCurrentManual();
        if (!manual) return;
        manual.minutaClient = e.target.value;
        manual.updatedAt = Date.now();
        scheduleSave();
        updateMinutaSubject();
    });
    $('#minutaDate').addEventListener('change', (e) => {
        const manual = getCurrentManual();
        if (!manual) return;
        manual.minutaDate = e.target.value;
        manual.updatedAt = Date.now();
        scheduleSave();
        updateMinutaSubject();
    });
    $('#manageClientsBtn').addEventListener('click', openClientModal);
    $('#addClientBtn').addEventListener('click', addClient);
    $('#newClientInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') addClient(); });
    $('#closeClientModal').addEventListener('click', () => hideModal('clientModal'));

    // Global paste → image
    document.addEventListener('paste', handleGlobalPaste);

    // Gantt / Cronograma
    $('#addMilestoneBtn').addEventListener('click', addMilestone);
    $('#exportGanttPdfBtn').addEventListener('click', exportGanttPdf);
}

// ========== DELEGATED SECTION CLICK ==========
function handleContainerClick(e) {
    const card = e.target.closest('.section-card');
    if (!card) return;
    const sectionId = card.dataset.sectionId;
    const manual    = getCurrentManual();
    if (!manual) return;
    const idx = manual.sections.findIndex(s => s.id === sectionId);
    if (idx === -1) return;

    if      (e.target.closest('.move-up'))       { moveSection(idx, -1); }
    else if (e.target.closest('.move-down'))      { moveSection(idx,  1); }
    else if (e.target.closest('.delete-section')) { removeSection(idx); }
    else if (e.target.closest('.remove-image'))   { e.stopPropagation(); removeImage(idx, sectionId, manual); }
    else if (e.target.closest('.image-drop-zone') && !manual.sections[idx].imageKey) {
        triggerImagePick(idx, sectionId);
    }
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
    const list   = $('#sidebarList');
    const sorted = [...manuals].sort((a, b) => b.updatedAt - a.updatedAt);
    list.innerHTML = sorted.map(m => {
        const isActive = m.id === currentManualId;
        const date = new Date(m.updatedAt).toLocaleDateString('es-AR', { day: '2-digit', month: 'short' });
        const type = DOC_TYPES[m.docType] || DOC_TYPES.manual;
        return `<div class="sidebar-item${isActive ? ' active' : ''}" data-id="${m.id}">
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
                loadManual(item.dataset.id);
                renderSections();
                renderSidebar();
                closeSidebar();
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
        id:           Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title:        type.defaultTitle,
        docType,
        sections:     [],
        milestones:   [],
        createdAt:    Date.now(),
        updatedAt:    Date.now(),
        minutaClient: '',
        minutaDate:   ''
    };
    manuals.push(manual);
    if (switchTo) { loadManual(manual.id); renderSections(); }
    saveMeta();
    renderSidebar();
    if (switchTo) toast('Nuevo documento creado');
}

function loadManual(id) {
    currentManualId = id;
    const manual = getCurrentManual();
    if (!manual) return;
    docTitle.value = manual.title;
    updateDocTypeBadge();
    updateMinutaPanel();
    updateGanttPanel();
    localStorage.setItem(CURRENT_KEY, id);
}

function getCurrentManual()   { return manuals.find(m => m.id === currentManualId) || null; }
function getCurrentSections() { const m = getCurrentManual(); return m ? m.sections : []; }

function updateDocTypeBadge() {
    const manual = getCurrentManual();
    const badge  = $('#docTypeBadge');
    if (manual && DOC_TYPES[manual.docType]) {
        badge.textContent = DOC_TYPES[manual.docType].label;
        badge.className   = `doc-type-badge ${manual.docType}`;
    } else {
        badge.textContent = '';
        badge.className   = 'doc-type-badge';
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
    const target = manuals.find(m => m.id === deleteTargetId);
    if (target) target.sections.forEach(s => { if (s.imageKey) dbDelete(s.imageKey); });
    manuals = manuals.filter(m => m.id !== deleteTargetId);
    if (manuals.length === 0) createNewManual('manual', false);
    if (currentManualId === deleteTargetId) { loadManual(manuals[0].id); renderSections(); }
    deleteTargetId = null;
    saveMeta();
    renderSidebar();
    hideModal('deleteModal');
    toast('Documento eliminado');
}

// ========== MINUTA PANEL ==========
function updateMinutaPanel() {
    const manual = getCurrentManual();
    const panel  = $('#minutaPanel');
    if (!manual || manual.docType !== 'minuta') { panel.style.display = 'none'; return; }
    panel.style.display = '';
    renderClientSelect();
    $('#minutaClientSelect').value = manual.minutaClient || '';
    const dateInput = $('#minutaDate');
    if (manual.minutaDate) {
        dateInput.value = manual.minutaDate;
    } else {
        dateInput.value = new Date().toISOString().slice(0, 10);
        manual.minutaDate = dateInput.value;
    }
    updateMinutaSubject();
}

function updateMinutaSubject() {
    const manual = getCurrentManual();
    if (!manual || manual.docType !== 'minuta') return;
    const client = manual.minutaClient || '';
    const date   = manual.minutaDate || '';
    const title  = manual.title || 'Minuta de Reunión';
    let subject  = 'Seleccioná un cliente para generar el asunto';
    if (client && date) {
        const [y, m, d] = date.split('-');
        subject = `[${client}] ${title} - ${d}/${m}/${y}`;
    } else if (client) {
        subject = `[${client}] ${title}`;
    }
    $('#minutaSubjectText').textContent = subject;
}

// ========== CLIENTS ==========
function getClients() {
    try { return JSON.parse(localStorage.getItem(CLIENTS_KEY) || '[]'); }
    catch { return []; }
}
function saveClients(clients) { localStorage.setItem(CLIENTS_KEY, JSON.stringify(clients)); }

function renderClientSelect() {
    const select  = $('#minutaClientSelect');
    const current = select.value;
    select.innerHTML = '<option value="">Seleccionar cliente...</option>' +
        getClients().map(c => `<option value="${escapeAttr(c)}">${escapeHtml(c)}</option>`).join('');
    if (current) select.value = current;
}

function openClientModal() { renderClientListModal(); showModal('clientModal'); }

function renderClientListModal() {
    const clients = getClients();
    const list    = $('#clientListModal');
    if (!clients.length) {
        list.innerHTML = '<p style="color:var(--text-muted);font-size:0.9rem;">No hay clientes todavía.</p>';
        return;
    }
    list.innerHTML = clients.map((c, i) =>
        `<div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);">
            <span>${escapeHtml(c)}</span>
            <button class="btn-icon btn-danger" data-idx="${i}" title="Eliminar">🗑</button>
        </div>`
    ).join('');
    list.querySelectorAll('[data-idx]').forEach(btn => {
        btn.addEventListener('click', () => {
            const clients = getClients();
            clients.splice(parseInt(btn.dataset.idx), 1);
            saveClients(clients);
            renderClientListModal();
            renderClientSelect();
        });
    });
}

function addClient() {
    const input = $('#newClientInput');
    const name  = input.value.trim();
    if (!name) return;
    const clients = getClients();
    if (clients.includes(name)) { toast('El cliente ya existe'); return; }
    clients.push(name);
    saveClients(clients);
    input.value = '';
    renderClientListModal();
    renderClientSelect();
    toast('Cliente agregado');
}

// ========== TEMPLATE ==========
function getDefaultTemplate() {
    return { logo: null, companyName: '', authorName: '', version: '1', subtitle: '', footerText: '' };
}
function getTemplateForCurrent() {
    const manual = getCurrentManual();
    const type   = manual ? manual.docType : 'manual';
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
    $('#authorName').value  = tmpl.authorName  || '';
    $('#docVersion').value  = tmpl.version     || '1';
    $('#docSubtitle').value = tmpl.subtitle    || '';
    $('#footerText').value  = tmpl.footerText  || '';
    if (tmpl.logo) {
        $('#logoPreview').src   = tmpl.logo;
        $('#logoPreview').style.display   = '';
        $('#logoPlaceholder').style.display = 'none';
        $('#removeLogo').style.display    = 'flex';
    } else {
        $('#logoPreview').style.display     = 'none';
        $('#logoPlaceholder').style.display = '';
        $('#removeLogo').style.display      = 'none';
    }
    showModal('templateModal');
}

function saveTemplate() {
    const tmpl = getTemplateForCurrent();
    tmpl.companyName = $('#companyName').value.trim();
    tmpl.authorName  = $('#authorName').value.trim();
    tmpl.version     = $('#docVersion').value.trim() || '1';
    tmpl.subtitle    = $('#docSubtitle').value.trim();
    tmpl.footerText  = $('#footerText').value.trim();
    localStorage.setItem(TEMPLATES_KEY, JSON.stringify(templates));
    hideModal('templateModal');
    toast('Template guardado para este tipo de documento');
}

// ========== SECTIONS CRUD ==========
function addSection() {
    const manual = getCurrentManual();
    if (!manual) return;
    const s = {
        id:        Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title:     '',
        imageKey:  null,
        imageType: null,
        html:      '',
        text:      ''
    };
    manual.sections.push(s);
    manual.updatedAt = Date.now();

    emptyState.style.display = 'none';
    const card = createSectionCard(s, manual.sections.length - 1);
    sectionsContainer.appendChild(card);
    initQuill(s);
    updateSectionCount();
    saveMeta();
    renderSidebar();
    setTimeout(() => card.scrollIntoView({ behavior: 'smooth', block: 'center' }), 50);
}

function removeSection(idx) {
    const manual = getCurrentManual();
    if (!manual) return;
    const s = manual.sections[idx];
    if (!s) return;
    if (s.imageKey) dbDelete(s.imageKey);
    delete quillInstances[`editor-${s.id}`];
    const card = sectionsContainer.querySelector(`.section-card[data-section-id="${s.id}"]`);
    if (card) card.remove();
    manual.sections.splice(idx, 1);
    manual.updatedAt = Date.now();
    if (!manual.sections.length) emptyState.style.display = '';
    updateSectionNumbers();
    saveMeta();
    renderSidebar();
}

function moveSection(idx, dir) {
    const sections = getCurrentSections();
    const newIdx   = idx + dir;
    if (newIdx < 0 || newIdx >= sections.length) return;
    [sections[idx], sections[newIdx]] = [sections[newIdx], sections[idx]];
    getCurrentManual().updatedAt = Date.now();
    reorderSectionCards();
    updateSectionNumbers();
    scheduleSave();
    renderSidebar();
}

function removeImage(idx, sectionId, manual) {
    const s = manual.sections[idx];
    if (!s) return;
    if (s.imageKey) { dbDelete(s.imageKey); s.imageKey = null; }
    s.imageType = null;
    manual.updatedAt = Date.now();
    const card = sectionsContainer.querySelector(`.section-card[data-section-id="${sectionId}"]`);
    if (card) {
        const zone = card.querySelector('.image-drop-zone');
        zone.classList.remove('has-image');
        zone.innerHTML = `<span class="placeholder-icon">🖼</span><span class="placeholder-text">Arrastrá una imagen aquí,<br>hacé clic, o pegá con Ctrl+V</span>`;
    }
    scheduleSave();
    renderSidebar();
}

// ========== SMART RENDER ==========
// Diffs the DOM against the current sections array.
// Creates cards only for new sections, removes cards for deleted ones,
// and reorders existing cards – preserving Quill instances throughout.
function renderSections() {
    if (isGanttMode()) return;
    const sections = getCurrentSections();

    // Map existing rendered cards by sectionId
    const existing = new Map();
    sectionsContainer.querySelectorAll('.section-card').forEach(c => existing.set(c.dataset.sectionId, c));

    const activeIds = new Set(sections.map(s => s.id));

    // Remove cards for sections that no longer exist
    existing.forEach((card, id) => {
        if (!activeIds.has(id)) {
            delete quillInstances[`editor-${id}`];
            card.remove();
        }
    });

    if (!sections.length) {
        emptyState.style.display = '';
        sectionCount.textContent = '0 secciones';
        return;
    }

    emptyState.style.display = 'none';

    // Add new cards + reorder all cards in section order
    sections.forEach((s, i) => {
        let card = existing.get(s.id);
        if (!card) {
            // New section: create card, load image async, init Quill
            card = createSectionCard(s, i);
            sectionsContainer.appendChild(card);
            loadImageIntoCard(s, card);
            initQuill(s);
        } else {
            // Update section number label only
            const numEl = card.querySelector('.section-number');
            if (numEl) numEl.textContent = `Sección ${i + 1}`;
        }
        // appendChild on existing node moves it to end → achieves reorder
        sectionsContainer.appendChild(card);
    });

    updateSectionCount();
}

function reorderSectionCards() {
    getCurrentSections().forEach(s => {
        const card = sectionsContainer.querySelector(`.section-card[data-section-id="${s.id}"]`);
        if (card) sectionsContainer.appendChild(card);
    });
}

function updateSectionNumbers() {
    getCurrentSections().forEach((s, i) => {
        const card  = sectionsContainer.querySelector(`.section-card[data-section-id="${s.id}"]`);
        const numEl = card && card.querySelector('.section-number');
        if (numEl) numEl.textContent = `Sección ${i + 1}`;
    });
    updateSectionCount();
}

function updateSectionCount() {
    const n = getCurrentSections().length;
    sectionCount.textContent = n + (n === 1 ? ' sección' : ' secciones');
}

function createSectionCard(section, idx) {
    const card     = document.createElement('div');
    const editorId = `editor-${section.id}`;
    card.className    = 'section-card';
    card.draggable    = true;
    card.dataset.sectionId = section.id;

    card.innerHTML = `
        <div class="section-header">
            <div class="section-header-left">
                <span class="drag-handle" title="Arrastrar para reordenar">⠿</span>
                <span class="section-number">Sección ${idx + 1}</span>
            </div>
            <div class="section-header-right">
                <button class="btn-icon move-up"       title="Mover arriba">↑</button>
                <button class="btn-icon move-down"     title="Mover abajo">↓</button>
                <button class="btn-icon btn-danger delete-section" title="Eliminar">🗑</button>
            </div>
        </div>
        <div class="section-body">
            <div class="section-left">
                <label class="context-label">Título de la sección</label>
                <input type="text" class="section-title-input" placeholder="Ej: Inicio de sesión" value="${escapeAttr(section.title)}">
                <div class="image-drop-zone">
                    <span class="placeholder-icon">🖼</span>
                    <span class="placeholder-text">Arrastrá una imagen aquí,<br>hacé clic, o pegá con Ctrl+V</span>
                </div>
            </div>
            <div class="section-right">
                <label class="context-label">Texto de la sección</label>
                <div class="editor-wrapper"><div id="${editorId}"></div></div>
            </div>
        </div>`;

    // Title input – looks up section by ID so idx doesn't go stale
    card.querySelector('.section-title-input').addEventListener('input', (e) => {
        const manual = getCurrentManual();
        if (!manual) return;
        const s = manual.sections.find(s => s.id === section.id);
        if (s) { s.title = e.target.value; manual.updatedAt = Date.now(); scheduleSave(); }
    });

    // Drag-and-drop reorder (card level)
    card.addEventListener('dragstart', (e) => {
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', section.id);
    });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));

    // Image drop zone – file drag only
    const zone = card.querySelector('.image-drop-zone');
    setupDropZoneDrag(zone, section.id);

    return card;
}

// Attaches file-drag listeners to a drop zone.
// Only stops propagation for Files so card-reorder drag can still bubble.
function setupDropZoneDrag(zone, sectionId) {
    zone.addEventListener('dragover', (e) => {
        if (!e.dataTransfer.types.includes('Files')) return;
        e.preventDefault();
        e.stopPropagation();
        zone.classList.add('drag-hover');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-hover'));
    zone.addEventListener('drop', (e) => {
        zone.classList.remove('drag-hover');
        if (e.dataTransfer.files.length && e.dataTransfer.files[0].type.startsWith('image/')) {
            e.preventDefault();
            e.stopPropagation();
            const manual = getCurrentManual();
            if (!manual) return;
            const idx = manual.sections.findIndex(s => s.id === sectionId);
            if (idx !== -1) handleImageFile(e.dataTransfer.files[0], idx, sectionId);
        }
    });
}

// ========== CONTAINER DRAG (card reorder) ==========
function handleContainerDragOver(e) {
    const card = e.target.closest('.section-card');
    if (!card) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    sectionsContainer.querySelectorAll('.section-card.drag-over').forEach(c => {
        if (c !== card) c.classList.remove('drag-over');
    });
    card.classList.add('drag-over');
}

function handleContainerDragLeave(e) {
    if (!sectionsContainer.contains(e.relatedTarget)) {
        sectionsContainer.querySelectorAll('.section-card.drag-over').forEach(c => c.classList.remove('drag-over'));
    }
}

function handleContainerDrop(e) {
    sectionsContainer.querySelectorAll('.section-card.drag-over').forEach(c => c.classList.remove('drag-over'));
    // File drops are handled by the zone – skip here
    if (e.dataTransfer.files.length) return;
    const targetCard = e.target.closest('.section-card');
    if (!targetCard) return;
    const fromId = e.dataTransfer.getData('text/plain');
    const toId   = targetCard.dataset.sectionId;
    if (!fromId || fromId === toId) return;
    const sections = getCurrentSections();
    const fromIdx  = sections.findIndex(s => s.id === fromId);
    const toIdx    = sections.findIndex(s => s.id === toId);
    if (fromIdx === -1 || toIdx === -1) return;
    e.preventDefault();
    const [moved] = sections.splice(fromIdx, 1);
    sections.splice(toIdx, 0, moved);
    getCurrentManual().updatedAt = Date.now();
    reorderSectionCards();
    updateSectionNumbers();
    scheduleSave();
    renderSidebar();
}

// ========== IMAGE HANDLING ==========
function triggerImagePick(idx, sectionId) {
    const input = document.createElement('input');
    input.type = 'file'; input.accept = 'image/*';
    input.addEventListener('change', () => { if (input.files[0]) handleImageFile(input.files[0], idx, sectionId); });
    input.click();
}

function handleImageFile(file, idx, sectionId) {
    const reader = new FileReader();
    reader.onload = async (e) => {
        const dataUrl = e.target.result;
        const manual  = getCurrentManual();
        if (!manual) return;
        const s = manual.sections[idx];
        if (!s) return;
        const key = `${manual.id}_${s.id}`;
        // Delete previous image if it had a different key
        if (s.imageKey && s.imageKey !== key) dbDelete(s.imageKey);
        try { await dbSave(key, dataUrl); } catch (err) { console.warn('dbSave failed:', err); }
        s.imageKey  = key;
        s.imageType = file.type;
        manual.updatedAt = Date.now();

        // Update the card's zone in place – no full re-render needed
        const card = sectionsContainer.querySelector(`.section-card[data-section-id="${sectionId}"]`);
        if (card) {
            const zone = card.querySelector('.image-drop-zone');
            zone.classList.add('has-image');
            zone.innerHTML = `<img src="${dataUrl}" alt="Captura"><button class="remove-image" title="Quitar imagen">✕</button>`;
        }
        saveMeta();
        renderSidebar();
    };
    reader.readAsDataURL(file);
}

async function loadImageIntoCard(section, card) {
    if (!section.imageKey) return;
    const dataUrl = await dbGet(section.imageKey);
    if (!dataUrl) return;
    const zone = card.querySelector('.image-drop-zone');
    if (!zone) return;
    zone.classList.add('has-image');
    zone.innerHTML = `<img src="${dataUrl}" alt="Captura"><button class="remove-image" title="Quitar imagen">✕</button>`;
}

function handleGlobalPaste(e) {
    const el = document.activeElement;
    if (el.closest('.ql-editor') || el.classList.contains('section-title-input') ||
        el.id === 'docTitle' || el.id === 'companyName' ||
        el.id === 'docSubtitle' || el.id === 'footerText') return;
    for (const item of e.clipboardData.items) {
        if (item.type.startsWith('image/')) {
            e.preventDefault();
            const sections = getCurrentSections();
            const last = sections[sections.length - 1];
            if (!sections.length || last.imageKey) addSection();
            const updated = getCurrentSections();
            const target  = updated[updated.length - 1];
            if (target) handleImageFile(item.getAsFile(), updated.length - 1, target.id);
            break;
        }
    }
}

// ========== QUILL EDITOR ==========
function initQuill(section) {
    const editorId  = `editor-${section.id}`;
    if (quillInstances[editorId]) return; // already alive – preserve it
    const container = document.getElementById(editorId);
    if (!container) return;

    const quill = new Quill(container, {
        theme: 'snow',
        placeholder: 'Escribí o pegá el texto de esta sección...',
        modules: {
            toolbar: [
                [{ header: [1, 2, 3, false] }],
                ['bold', 'italic', 'underline', 'strike'],
                [{ color: [] }, { background: [] }],
                [{ list: 'ordered' }, { list: 'bullet' }],
                [{ align: [] }],
                ['clean']
            ]
        }
    });

    if (section.html) quill.root.innerHTML = section.html;

    let saveTimer;
    quill.on('text-change', () => {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(() => {
            const manual = getCurrentManual();
            if (!manual) return;
            const s = manual.sections.find(s => s.id === section.id);
            if (s) { s.html = quill.root.innerHTML; s.text = quill.getText(); manual.updatedAt = Date.now(); saveMeta(); }
        }, 500);
    });

    quillInstances[editorId] = quill;
}

// ========== PERSISTENCE ==========
function scheduleSave() {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(saveMeta, 600);
}

// Saves metadata only – images live in IndexedDB, not localStorage
function saveMeta() {
    clearTimeout(_saveTimer);
    _saveTimer = null;
    try {
        const data = manuals.map(m => ({
            ...m,
            sections: m.sections.map(s => ({
                id:        s.id,
                title:     s.title,
                imageKey:  s.imageKey  || null,
                imageType: s.imageType || null,
                html:      s.html      || '',
                text:      s.text      || ''
            })),
            milestones: (m.milestones || []).map(ms => ({
                id:   ms.id,
                html: ms.html || '',
                text: ms.text || '',
                from: ms.from || '',
                to:   ms.to   || ''
            }))
        }));
        localStorage.setItem(MANUALS_KEY, JSON.stringify(data));
        localStorage.setItem(CURRENT_KEY, currentManualId);
    } catch (e) {
        console.warn('localStorage save failed:', e);
        toast('Almacenamiento lleno. Exportá y eliminá documentos.');
    }
}

function loadState() {
    try {
        const saved = localStorage.getItem(MANUALS_KEY);
        if (saved) manuals = JSON.parse(saved);

        // Migrate legacy single-manual format
        if (!manuals.length) {
            const old = localStorage.getItem('docgen_sections');
            if (old) {
                const sections = JSON.parse(old);
                sections.forEach(s => {
                    if (s.generated && !s.html) { s.html = `<p>${escapeHtml(s.generated)}</p>`; s.text = s.generated; }
                });
                if (sections.length) {
                    manuals.push({ id: 'migrated', title: localStorage.getItem('docgen_title') || 'Manual de Usuario', docType: 'manual', sections, createdAt: Date.now(), updatedAt: Date.now(), minutaClient: '', minutaDate: '' });
                    localStorage.removeItem('docgen_sections');
                    localStorage.removeItem('docgen_title');
                }
            }
        }

        // Ensure all manuals have current fields
        manuals.forEach(m => {
            if (!m.docType)                  m.docType      = 'manual';
            if (m.minutaClient === undefined) m.minutaClient = '';
            if (m.minutaDate   === undefined) m.minutaDate   = '';
            if (!m.milestones)               m.milestones   = [];
        });

        currentManualId = localStorage.getItem(CURRENT_KEY);

        const tmpl = localStorage.getItem(TEMPLATES_KEY);
        if (tmpl) templates = JSON.parse(tmpl);
        // Migrate legacy single template
        const oldTmpl = localStorage.getItem('docgen_template');
        if (oldTmpl && !templates.manual) {
            templates.manual = JSON.parse(oldTmpl);
            localStorage.removeItem('docgen_template');
            localStorage.setItem(TEMPLATES_KEY, JSON.stringify(templates));
        }
    } catch (e) {
        console.warn('loadState error:', e);
        manuals = [];
    }
}

// ========== WORD EXPORT ==========
// Default logo (Intiza) — used when no custom logo is set in template settings
const INTIZA_LOGO_B64 = 'data:image/jpeg;base64,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';

async function exportToWord() {
    const sections = getCurrentSections();
    if (!sections.length) { toast('Agregá secciones primero'); return; }
    if (!window.docx)     { toast('Error: la librería docx no se cargó. Recargá la página.'); return; }
    toast('Generando documento Word...');
    try {
        const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel,
                Header, Footer, PageNumber, AlignmentType, VerticalAlign,
                BorderStyle, PageBreak, Bookmark, InternalHyperlink,
                TableOfContents, Table, TableRow, TableCell, WidthType } = window.docx;

        const tmpl   = getTemplateForCurrent();
        const manual = getCurrentManual();
        const creationDate = new Date(manual.createdAt).toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });

        // ---- Template border (Intiza: #4472C4) ----
        const hBorder  = { style: BorderStyle.SINGLE, size: 8, color: '4472C4' };
        const hBorders = { top: hBorder, bottom: hBorder, left: hBorder, right: hBorder };

        // ---- Logo: custom upload or default Intiza logo ----
        const logoDataUrl = tmpl.logo || INTIZA_LOGO_B64;
        const logoType    = logoDataUrl.startsWith('data:image/png') ? 'png' : 'jpg';
        const logoBuffer  = await dataUrlToArrayBuffer(logoDataUrl);

        // ---- Header table: 3 cols, 2 rows (Intiza template) ----
        // Col widths: 2875 + 3798 + 2684 = 9357 DXA
        const headerTable = new Table({
            width: { size: 9357, type: WidthType.DXA },
            rows: [
                new TableRow({ children: [
                    new TableCell({
                        rowSpan: 2,
                        width: { size: 2875, type: WidthType.DXA },
                        borders: hBorders,
                        verticalAlign: VerticalAlign.CENTER,
                        children: [new Paragraph({
                            children: [new ImageRun({ data: logoBuffer, transformation: { width: 160, height: 65 }, type: logoType })],
                            alignment: AlignmentType.CENTER
                        })]
                    }),
                    new TableCell({
                        columnSpan: 2,
                        borders: hBorders,
                        children: [new Paragraph({
                            children: [new TextRun({ text: (docTitle.value || 'Documento').toUpperCase(), bold: true, size: 28, color: '002366' })],
                            alignment: AlignmentType.LEFT,
                            spacing: { before: 60, after: 60 }
                        })]
                    })
                ]}),
                new TableRow({ children: [
                    new TableCell({
                        width: { size: 3798, type: WidthType.DXA },
                        borders: hBorders,
                        children: [new Paragraph({
                            children: [new TextRun({ text: `Versión ${tmpl.version || '1'}`, size: 20, color: '333333' })],
                            alignment: AlignmentType.LEFT,
                            spacing: { before: 40, after: 40 }
                        })]
                    }),
                    new TableCell({
                        width: { size: 2684, type: WidthType.DXA },
                        borders: hBorders,
                        children: [new Paragraph({
                            children: [
                                ...(tmpl.authorName ? [new TextRun({ text: tmpl.authorName, size: 20, color: '333333' }), new TextRun({ text: '   ', size: 20 })] : []),
                                new TextRun({ text: `(${creationDate})`, size: 18, color: '888888' })
                            ],
                            alignment: AlignmentType.LEFT,
                            spacing: { before: 40, after: 40 }
                        })]
                    })
                ]})
            ]
        });

        // ---- Footer ----
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

        // ---- Cover page ----
        docChildren.push(new Paragraph({ spacing: { before: 3000 } }));
        docChildren.push(new Paragraph({
            children: [new TextRun({ text: docTitle.value || 'Documento', bold: true, size: 56, color: '002366' })],
            alignment: AlignmentType.CENTER
        }));
        if (tmpl.subtitle) {
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: tmpl.subtitle, size: 28, color: '666666' })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200 }
            }));
        }
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // ---- TOC ----
        docChildren.push(new Paragraph({ children: [new TextRun({ text: 'Índice de Contenidos', bold: true, size: 32, color: '002366' })], spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: '002366' } } }));
        docChildren.push(new TableOfContents('Índice', { hyperlink: true, headingStyleRange: '1-2' }));
        sections.forEach((s, i) => {
            const title = s.title || `Sección ${i + 1}`;
            docChildren.push(new Paragraph({ children: [new InternalHyperlink({ anchor: `section_${i}`, children: [new TextRun({ text: `${i + 1}. ${title}`, style: 'Hyperlink', size: 22 })] })], spacing: { before: 100, after: 40 } }));
        });
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // ---- Section content ----
        for (let i = 0; i < sections.length; i++) {
            const s     = sections[i];
            const title = s.title || `Sección ${i + 1}`;
            docChildren.push(new Paragraph({
                children: [new Bookmark({ id: `section_${i}`, children: [new TextRun({ text: `${i + 1}. ${title}`, bold: true, size: 28, color: '002366' })] })],
                heading:  HeadingLevel.HEADING_1,
                spacing:  { before: 400, after: 200 },
                border:   { bottom: { style: BorderStyle.SINGLE, size: 2, color: '002366' } }
            }));

            if (s.imageKey) {
                try {
                    const imgUrl = await dbGet(s.imageKey);
                    if (imgUrl) {
                        const imgData = await dataUrlToArrayBuffer(imgUrl);
                        const dims    = await getImageDimensions(imgUrl);
                        const maxW = 500;
                        let w = dims.width, h = dims.height;
                        if (w > maxW) { h = Math.round(h * (maxW / w)); w = maxW; }
                        docChildren.push(new Paragraph({
                            children: [new ImageRun({ data: imgData, transformation: { width: w, height: h }, type: s.imageType?.includes('png') ? 'png' : 'jpg' })],
                            alignment: AlignmentType.CENTER,
                            spacing:   { before: 200, after: 200 }
                        }));
                    }
                } catch (e) { console.warn('Image section', i, e); }
            }

            if (s.text?.trim()) docChildren.push(...htmlToDocxParagraphs(s.html || '', TextRun, Paragraph));
            if (i < sections.length - 1) docChildren.push(new Paragraph({ children: [new PageBreak()] }));
        }

        const doc = new Document({
            features: { updateFields: true },
            styles:   { default: { document: { run: { font: 'Calibri', size: 22 } } } },
            sections: [{
                properties: { page: { margin: { top: 1417, right: 1701, bottom: 1417, left: 1701 }, pageNumbers: { start: 1 } } },
                headers:    { default: new Header({ children: [headerTable] }) },
                footers:    { default: new Footer({ children: footerChildren }) },
                children:   docChildren
            }]
        });

        const blob     = await Packer.toBlob(doc);
        const fileName = (docTitle.value || 'documento').replace(/[^a-zA-Z0-9áéíóúñÁÉÍÓÚÑ\s-]/g, '').replace(/\s+/g, '_');
        window.saveAs(blob, `${fileName}.docx`);
        toast('Documento Word descargado');
    } catch (e) {
        console.error('Export error:', e);
        toast('Error al exportar: ' + e.message);
    }
}

// ========== HTML → DOCX PARAGRAPHS ==========
function htmlToDocxParagraphs(html, TextRun, Paragraph) {
    const paragraphs = [];
    const tmp        = document.createElement('div');
    tmp.innerHTML = html;
    const blocks  = tmp.querySelectorAll('p, h1, h2, h3, li, div');
    if (!blocks.length) {
        const text = tmp.textContent || '';
        text.split('\n').filter(l => l.trim()).forEach(l =>
            paragraphs.push(new Paragraph({ children: [new TextRun({ text: l, size: 22, font: 'Calibri' })], spacing: { before: 60, after: 60 } })));
        return paragraphs;
    }
    blocks.forEach(block => {
        const runs = [];
        processNode(block, runs, TextRun);
        if (!runs.length) return;
        const opts = { children: runs, spacing: { before: 60, after: 60 } };
        const tag  = block.tagName.toLowerCase();
        if (tag === 'h1') opts.spacing = { before: 200, after: 100 };
        if (tag === 'h2') opts.spacing = { before: 160, after: 80 };
        if (tag === 'h3') opts.spacing = { before: 120, after: 60 };
        if (tag === 'li') {
            if (block.parentElement?.tagName === 'OL') opts.numbering = { reference: 'default-numbering', level: 0 };
            else opts.bullet = { level: 0 };
        }
        paragraphs.push(new Paragraph(opts));
    });
    return paragraphs.length
        ? paragraphs
        : [new Paragraph({ children: [new TextRun({ text: tmp.textContent || '', size: 22, font: 'Calibri' })], spacing: { before: 60, after: 60 } })];
}

function processNode(node, runs, TextRun) {
    node.childNodes.forEach(child => {
        if (child.nodeType === 3) {
            const text = child.textContent;
            if (text) {
                const s = getInheritedStyles(child);
                runs.push(new TextRun({ text, size: s.size || 22, font: 'Calibri', bold: s.bold, italics: s.italic, underline: s.underline ? {} : undefined, strike: s.strike }));
            }
        } else if (child.nodeType === 1) {
            processNode(child, runs, TextRun);
        }
    });
}

function getInheritedStyles(node) {
    const s = { bold: false, italic: false, underline: false, strike: false, size: 22 };
    let el  = node.parentElement;
    while (el) {
        const tag = el.tagName.toLowerCase();
        if (tag === 'strong' || tag === 'b')           s.bold      = true;
        if (tag === 'em'     || tag === 'i')            s.italic    = true;
        if (tag === 'u')                                s.underline = true;
        if (tag === 's' || tag === 'del' || tag === 'strike') s.strike = true;
        if (tag === 'h1') { s.bold = true; s.size = 32; }
        if (tag === 'h2') { s.bold = true; s.size = 28; }
        if (tag === 'h3') { s.bold = true; s.size = 24; }
        if (el.classList.contains('ql-editor') || el.classList.contains('section-right')) break;
        el = el.parentElement;
    }
    return s;
}

// ========== PREVIEW & MARKDOWN ==========
async function showPreview() {
    const sections = getCurrentSections();
    if (!sections.length) { toast('Agregá secciones primero'); return; }
    const content = $('#previewContent');
    let html = `<h1>${escapeHtml(docTitle.value || 'Documento')}</h1>`;
    for (let i = 0; i < sections.length; i++) {
        const s     = sections[i];
        const title = s.title || `Sección ${i + 1}`;
        html += `<div class="preview-section"><h2>${i + 1}. ${escapeHtml(title)}</h2>`;
        if (s.imageKey) {
            const imgUrl = await dbGet(s.imageKey);
            if (imgUrl) html += `<img src="${imgUrl}" alt="${escapeHtml(title)}">`;
        }
        html += s.html
            ? `<div class="preview-text">${s.html}</div>`
            : `<div class="preview-text" style="color:var(--text-muted)">(Sin texto todavía)</div>`;
        html += `</div>`;
    }
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

// ========== MODAL UTILS ==========
function showModal(id) { document.getElementById(id).classList.add('active'); }
function hideModal(id) { document.getElementById(id).classList.remove('active'); }
function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 3000);
}

// ========== STRING UTILS ==========
function escapeHtml(str) {
    return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}
function escapeAttr(str) {
    return (str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function dataUrlToArrayBuffer(dataUrl) {
    return new Promise((resolve) => {
        const b = atob(dataUrl.split(',')[1]);
        const u = new Uint8Array(b.length);
        for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i);
        resolve(u.buffer);
    });
}
function getImageDimensions(dataUrl) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload  = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
        img.onerror = () => resolve({ width: 400, height: 300 });
        img.src = dataUrl;
    });
}

// ========== GANTT / CRONOGRAMA ==========
function isGanttMode() {
    const m = getCurrentManual();
    return !!(m && m.docType === 'cronograma');
}

function updateGanttPanel() {
    const isGantt = isGanttMode();
    const manual  = getCurrentManual();

    const secToolbar = $('#sectionsToolbar');
    if (secToolbar) secToolbar.style.display = isGantt ? 'none' : '';
    sectionsContainer.style.display = isGantt ? 'none' : '';
    emptyState.style.display = 'none'; // each mode manages its own empty states

    const gc = $('#ganttContainer');
    if (gc) gc.style.display = isGantt ? '' : 'none';

    if (!isGantt) {
        // Restore sections display
        updateSectionCount();
        const sections = getCurrentSections();
        if (!sections.length) emptyState.style.display = '';
        return;
    }

    if (!manual.milestones) manual.milestones = [];
    renderGanttMilestones();
    renderGanttChart();
    updateMilestoneCount();
}

function updateMilestoneCount() {
    const manual = getCurrentManual();
    const count = manual?.milestones?.length || 0;
    const el = $('#milestoneCount');
    if (el) el.textContent = count === 1 ? '1 hito' : `${count} hitos`;
}

function addMilestone() {
    const manual = getCurrentManual();
    if (!manual) return;
    if (!manual.milestones) manual.milestones = [];

    const today    = new Date();
    const nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
    const fmt = d => d.toISOString().slice(0, 10);

    const milestone = {
        id:   Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        html: '',
        text: '',
        from: fmt(today),
        to:   fmt(nextWeek)
    };
    manual.milestones.push(milestone);
    manual.updatedAt = Date.now();

    const container = $('#ganttMilestones');
    if (container) {
        const row = createMilestoneRow(milestone, manual.milestones.length);
        container.appendChild(row);
        initMilestoneQuill(milestone);
        setTimeout(() => row.scrollIntoView({ behavior: 'smooth', block: 'nearest' }), 50);
    }

    renderGanttChart();
    updateMilestoneCount();
    saveMeta();
}

function removeMilestone(id) {
    const manual = getCurrentManual();
    if (!manual || !manual.milestones) return;
    const idx = manual.milestones.findIndex(m => m.id === id);
    if (idx === -1) return;

    delete quillInstances[`editor-g-${id}`];
    const row = document.querySelector(`.milestone-row[data-id="${id}"]`);
    if (row) row.remove();

    manual.milestones.splice(idx, 1);
    manual.updatedAt = Date.now();

    // Renumber remaining rows
    document.querySelectorAll('#ganttMilestones .milestone-row').forEach((r, i) => {
        const numEl = r.querySelector('.milestone-num');
        if (numEl) numEl.textContent = i + 1;
    });

    renderGanttChart();
    updateMilestoneCount();
    saveMeta();
}

function createMilestoneRow(milestone, num) {
    const row = document.createElement('div');
    row.className = 'milestone-row';
    row.dataset.id = milestone.id;

    row.innerHTML = `
        <div class="milestone-num">${num}</div>
        <div class="milestone-editor-cell">
            <div id="editor-g-${milestone.id}" class="milestone-quill-wrap"></div>
        </div>
        <div class="milestone-dates">
            <label class="gantt-date-label">Desde</label>
            <input type="date" class="gantt-date-input milestone-from" value="${milestone.from || ''}">
            <label class="gantt-date-label">Hasta</label>
            <input type="date" class="gantt-date-input milestone-to" value="${milestone.to || ''}">
        </div>
        <button class="btn-icon milestone-delete" title="Eliminar hito">✕</button>
    `;

    row.querySelector('.milestone-from').addEventListener('change', (e) => {
        milestone.from = e.target.value;
        const m = getCurrentManual();
        if (m) m.updatedAt = Date.now();
        renderGanttChart();
        scheduleSave();
    });
    row.querySelector('.milestone-to').addEventListener('change', (e) => {
        milestone.to = e.target.value;
        const m = getCurrentManual();
        if (m) m.updatedAt = Date.now();
        renderGanttChart();
        scheduleSave();
    });
    row.querySelector('.milestone-delete').addEventListener('click', () => removeMilestone(milestone.id));

    return row;
}

function initMilestoneQuill(milestone) {
    const editorId = `editor-g-${milestone.id}`;
    if (quillInstances[editorId]) return;
    const container = document.getElementById(editorId);
    if (!container) return;

    const q = new Quill(container, {
        theme: 'snow',
        placeholder: 'Nombre del hito...',
        modules: {
            toolbar: [
                ['bold', 'italic', 'underline'],
                [{ color: [] }],
                ['clean']
            ]
        }
    });

    if (milestone.html) q.root.innerHTML = milestone.html;

    q.on('text-change', () => {
        milestone.html = q.root.innerHTML;
        milestone.text = q.getText().trim();
        const manual = getCurrentManual();
        if (manual) manual.updatedAt = Date.now();
        scheduleSave();
    });

    quillInstances[editorId] = q;
}

function renderGanttMilestones() {
    const manual    = getCurrentManual();
    const container = $('#ganttMilestones');
    if (!container) return;

    const milestones = manual?.milestones || [];

    const existing = new Map();
    container.querySelectorAll('.milestone-row').forEach(r => existing.set(r.dataset.id, r));

    const activeIds = new Set(milestones.map(m => m.id));
    existing.forEach((row, id) => {
        if (!activeIds.has(id)) { delete quillInstances[`editor-g-${id}`]; row.remove(); }
    });

    milestones.forEach((m, i) => {
        let row = existing.get(m.id);
        if (!row) {
            row = createMilestoneRow(m, i + 1);
            container.appendChild(row);
            initMilestoneQuill(m);
        } else {
            const numEl = row.querySelector('.milestone-num');
            if (numEl) numEl.textContent = i + 1;
            // Sync date inputs if changed externally
            const fi = row.querySelector('.milestone-from');
            const ti = row.querySelector('.milestone-to');
            if (fi && fi.value !== m.from) fi.value = m.from || '';
            if (ti && ti.value !== m.to)   ti.value = m.to   || '';
        }
        container.appendChild(row); // reorder
    });

    updateMilestoneCount();
}

function renderGanttChart() {
    const manual  = getCurrentManual();
    const chartEl = $('#ganttRight');
    if (!chartEl) return;

    const milestones = manual?.milestones || [];
    const withDates  = milestones.filter(m => m.from && m.to && m.from <= m.to);

    if (!withDates.length) {
        chartEl.innerHTML = '<div class="gantt-empty">Agregá hitos con fechas para ver el diagrama.</div>';
        return;
    }

    // Overall date range
    const allDates  = withDates.flatMap(m => [new Date(m.from), new Date(m.to)]);
    const minDate   = new Date(Math.min(...allDates));
    const maxDate   = new Date(Math.max(...allDates));

    // Snap start to Monday of minDate's week
    const startDate = new Date(minDate);
    const dow = startDate.getDay();
    startDate.setDate(startDate.getDate() - (dow === 0 ? 6 : dow - 1));

    // Snap end to Sunday of maxDate's week
    const endDate = new Date(maxDate);
    const edow = endDate.getDay();
    if (edow !== 0) endDate.setDate(endDate.getDate() + (7 - edow));

    const totalDays = Math.max(1, Math.round((endDate - startDate) / 864e5));
    const useMonths = totalDays > 84; // > 12 weeks → switch to monthly

    // Build periods
    const periods = [];
    if (useMonths) {
        const d = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
        while (d <= endDate) {
            const pStart = new Date(Math.max(d, startDate));
            const pEnd   = new Date(Math.min(new Date(d.getFullYear(), d.getMonth() + 1, 0), endDate));
            periods.push({ label: d.toLocaleDateString('es-AR', { month: 'short', year: '2-digit' }), start: pStart, end: pEnd });
            d.setMonth(d.getMonth() + 1);
        }
    } else {
        const d = new Date(startDate);
        while (d <= endDate) {
            const pStart = new Date(d);
            const pEnd   = new Date(Math.min(new Date(d.getTime() + 6 * 864e5), endDate));
            periods.push({ label: d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit' }), start: pStart, end: pEnd });
            d.setDate(d.getDate() + 7);
        }
    }

    const pct = (d) => Math.round((d - startDate) / (864e5 * totalDays) * 1000) / 10;

    // Header
    const headerCells = periods.map(p => {
        const w = Math.round(((p.end - p.start) / 864e5 + 1) / totalDays * 1000) / 10;
        return `<div class="gantt-hdr-cell" style="width:${w}%">${p.label}</div>`;
    }).join('');

    // Bar rows – one per milestone (including those without dates = empty row for alignment)
    const barRows = milestones.map(m => {
        if (!m.from || !m.to || m.from > m.to) return '<div class="gantt-bar-row"></div>';
        const left  = Math.max(0, pct(new Date(m.from)));
        const right = Math.min(100, pct(new Date(m.to)) + (1 / totalDays * 100));
        const width = Math.max(0.5, right - left);
        return `<div class="gantt-bar-row"><div class="gantt-bar" style="left:${left}%;width:${width}%"></div></div>`;
    }).join('');

    chartEl.innerHTML = `
        <div class="gantt-chart-inner">
            <div class="gantt-hdr-row">${headerCells}</div>
            <div class="gantt-bars">${barRows}</div>
        </div>
    `;
}

function exportGanttPdf() {
    window.print();
}
