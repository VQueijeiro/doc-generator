/* global Quill */
// ========== CONSTANTS & STATE ==========
const MANUALS_KEY  = 'docgen_manuals';
const CURRENT_KEY  = 'docgen_current';
const TEMPLATES_KEY = 'docgen_templates';
const CLIENTS_KEY  = 'docgen_clients';
const THEME_KEY    = 'docgen_theme';

const DOC_TYPES = {
    manual:     { label: 'Manual de Uso',      icon: '📘', defaultTitle: 'Manual de Usuario' },
    desarrollo: { label: 'Doc. de Desarrollo', icon: '⚙️', defaultTitle: 'Documento de Desarrollo' },
    minuta:     { label: 'Minuta de Reunión',  icon: '📋', defaultTitle: 'Minuta de Reunión' }
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
async function exportToWord() {
    const sections = getCurrentSections();
    if (!sections.length) { toast('Agregá secciones primero'); return; }
    if (!window.docx)     { toast('Error: la librería docx no se cargó. Recargá la página.'); return; }
    toast('Generando documento Word...');
    try {
        const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel,
                Header, Footer, PageNumber, AlignmentType,
                BorderStyle, PageBreak, Bookmark, InternalHyperlink,
                TableOfContents, Table, TableRow, TableCell, WidthType } = window.docx;

        const tmpl   = getTemplateForCurrent();
        const manual = getCurrentManual();
        const creationDate = new Date(manual.createdAt).toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });

        // ---- Header ----
        const headerChildren = [];
        if (tmpl.logo) {
            try {
                const d = await dataUrlToArrayBuffer(tmpl.logo);
                headerChildren.push(new Paragraph({ children: [new ImageRun({ data: d, transformation: { width: 120, height: 40 }, type: 'png' })], alignment: AlignmentType.LEFT }));
            } catch (e) { console.warn('Logo header:', e); }
        }
        if (tmpl.companyName) {
            headerChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.companyName, bold: true, size: 18, color: '666666' })], alignment: AlignmentType.LEFT }));
        }

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
        const cBorder  = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
        const cBorders = { top: cBorder, bottom: cBorder, left: cBorder, right: cBorder };

        // ---- Cover table (Word 2010 compatible: no rowSpan/columnSpan) ----
        const logoCellChildren = [];
        if (tmpl.logo) {
            try {
                const d = await dataUrlToArrayBuffer(tmpl.logo);
                logoCellChildren.push(new Paragraph({ children: [new ImageRun({ data: d, transformation: { width: 100, height: 35 }, type: 'png' })], alignment: AlignmentType.CENTER }));
            } catch (e) { console.warn('Logo cover:', e); }
        }
        if (tmpl.companyName) {
            logoCellChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.companyName, bold: true, size: 22, color: '002366' })], alignment: AlignmentType.CENTER, spacing: { before: 40 } }));
        }
        if (!logoCellChildren.length) logoCellChildren.push(new Paragraph({ children: [new TextRun({ text: ' ' })] }));

        const verRuns = [new TextRun({ text: `Versión ${tmpl.version || '1'}`, size: 20, color: '333333' })];
        if (tmpl.authorName) verRuns.push(new TextRun({ text: `    ${tmpl.authorName}`, size: 20, color: '333333', bold: true }));
        verRuns.push(new TextRun({ text: `    (${creationDate})`, size: 18, color: '888888' }));

        docChildren.push(new Table({
            width: { size: 9000, type: WidthType.DXA },
            rows: [
                new TableRow({ children: [
                    new TableCell({ width: { size: 2700, type: WidthType.DXA }, children: logoCellChildren, borders: cBorders }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (docTitle.value || 'Documento').toUpperCase(), bold: true, size: 26, color: '002366' })], alignment: AlignmentType.LEFT })], borders: cBorders })
                ]}),
                new TableRow({ children: [
                    new TableCell({ width: { size: 2700, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: ' ' })] })], borders: cBorders }),
                    new TableCell({ children: [new Paragraph({ children: verRuns })], borders: cBorders })
                ]})
            ]
        }));

        docChildren.push(new Paragraph({ spacing: { before: 2000 } }));
        docChildren.push(new Paragraph({ children: [new TextRun({ text: docTitle.value || 'Documento', bold: true, size: 56, color: '002366' })], alignment: AlignmentType.CENTER }));
        if (tmpl.subtitle) docChildren.push(new Paragraph({ children: [new TextRun({ text: tmpl.subtitle, size: 28, color: '666666' })], alignment: AlignmentType.CENTER, spacing: { before: 200 } }));
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
                properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, pageNumbers: { start: 1 } } },
                headers:    headerChildren.length ? { default: new Header({ children: headerChildren }) } : undefined,
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
