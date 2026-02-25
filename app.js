// State
let sections = [];
let apiKey = localStorage.getItem('docgen_apikey') || '';
const STORAGE_KEY = 'docgen_sections';
const TITLE_KEY = 'docgen_title';

// DOM
const $ = (s) => document.querySelector(s);
const sectionsContainer = $('#sectionsContainer');
const emptyState = $('#emptyState');
const sectionCount = $('#sectionCount');
const docTitle = $('#docTitle');

// Init
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    renderAll();
    setupEventListeners();
    checkApiKey();
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('sw.js');
    }
});

function setupEventListeners() {
    $('#addSectionBtn').addEventListener('click', addSection);
    $('#settingsBtn').addEventListener('click', () => showModal('apiKeyModal'));
    $('#saveApiKey').addEventListener('click', saveApiKey);
    $('#closeApiKeyModal').addEventListener('click', () => hideModal('apiKeyModal'));
    $('#previewBtn').addEventListener('click', showPreview);
    $('#generateAllBtn').addEventListener('click', generateAll);
    $('#closePreview').addEventListener('click', () => hideModal('previewModal'));
    $('#exportPdf').addEventListener('click', exportPdf);
    $('#copyMarkdown').addEventListener('click', copyMarkdown);
    docTitle.addEventListener('input', () => {
        localStorage.setItem(TITLE_KEY, docTitle.value);
    });

    // Global paste
    document.addEventListener('paste', (e) => {
        const activeEl = document.activeElement;
        if (activeEl.classList.contains('context-input') || activeEl.classList.contains('generated-text') ||
            activeEl.id === 'docTitle' || activeEl.id === 'apiKeyInput') return;

        const items = e.clipboardData.items;
        for (const item of items) {
            if (item.type.startsWith('image/')) {
                e.preventDefault();
                const file = item.getAsFile();
                // If no sections or last section has image, add new
                if (sections.length === 0 || sections[sections.length - 1].image) {
                    addSection();
                }
                const lastIdx = sections.length - 1;
                handleImageFile(file, lastIdx);
                break;
            }
        }
    });
}

// API Key
function checkApiKey() {
    if (!apiKey) {
        showModal('apiKeyModal');
    } else {
        $('#closeApiKeyModal').style.display = '';
    }
}

async function saveApiKey() {
    const key = $('#apiKeyInput').value.trim();
    if (!key) return;

    const status = $('#apiStatus');
    status.textContent = 'Validando...';
    status.className = 'api-status';

    try {
        const res = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'x-api-key': key,
                'anthropic-version': '2023-06-01',
                'anthropic-dangerous-direct-browser-access': 'true',
                'content-type': 'application/json'
            },
            body: JSON.stringify({
                model: 'claude-sonnet-4-5-20250929',
                max_tokens: 10,
                messages: [{ role: 'user', content: 'Di "OK"' }]
            })
        });

        if (res.ok) {
            apiKey = key;
            localStorage.setItem('docgen_apikey', key);
            status.textContent = '✓ API key válida y guardada';
            status.className = 'api-status success';
            $('#closeApiKeyModal').style.display = '';
            setTimeout(() => hideModal('apiKeyModal'), 1200);
        } else {
            const err = await res.json();
            status.textContent = '✗ Error: ' + (err.error?.message || 'Key inválida');
            status.className = 'api-status error';
        }
    } catch (e) {
        status.textContent = '✗ Error de conexión';
        status.className = 'api-status error';
    }
}

// Sections CRUD
function addSection() {
    sections.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        image: null,
        imageType: null,
        context: '',
        generated: ''
    });
    saveState();
    renderAll();
    // Scroll to new section
    setTimeout(() => {
        const cards = sectionsContainer.querySelectorAll('.section-card');
        if (cards.length) cards[cards.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 50);
}

function removeSection(idx) {
    sections.splice(idx, 1);
    saveState();
    renderAll();
}

function moveSection(idx, dir) {
    const newIdx = idx + dir;
    if (newIdx < 0 || newIdx >= sections.length) return;
    [sections[idx], sections[newIdx]] = [sections[newIdx], sections[idx]];
    saveState();
    renderAll();
}

// Render
function renderAll() {
    const fragment = document.createDocumentFragment();

    if (sections.length === 0) {
        emptyState.style.display = '';
    } else {
        emptyState.style.display = 'none';
        sections.forEach((s, i) => fragment.appendChild(createSectionCard(s, i)));
    }

    // Remove old cards
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
                <div class="image-drop-zone ${section.image ? 'has-image' : ''}" data-idx="${idx}">
                    ${section.image
                        ? `<img src="${section.image}" alt="Captura"><button class="remove-image" title="Quitar imagen">✕</button>`
                        : `<span class="placeholder-icon">🖼</span><span class="placeholder-text">Arrastrá una imagen aquí,<br>hacé clic, o pegá con Ctrl+V</span>`
                    }
                </div>
                <label class="context-label">Contexto (opcional)</label>
                <textarea class="context-input" placeholder="Ej: Esta pantalla muestra el formulario de alta de cliente..." data-idx="${idx}">${section.context}</textarea>
            </div>
            <div class="section-right">
                <div class="generated-label">
                    <label class="context-label">Texto generado</label>
                </div>
                <textarea class="generated-text" placeholder="El texto generado por IA aparecerá aquí. También podés escribir manualmente." data-idx="${idx}">${section.generated}</textarea>
                <button class="btn btn-primary generate-btn" data-idx="${idx}" ${!section.image ? 'disabled' : ''}>
                    ⚡ Generar texto
                </button>
            </div>
        </div>
    `;

    // Events
    card.querySelector('.move-up').addEventListener('click', () => moveSection(idx, -1));
    card.querySelector('.move-down').addEventListener('click', () => moveSection(idx, 1));
    card.querySelector('.delete-section').addEventListener('click', () => removeSection(idx));

    const dropZone = card.querySelector('.image-drop-zone');
    setupDropZone(dropZone, idx);

    const removeBtn = card.querySelector('.remove-image');
    if (removeBtn) {
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            sections[idx].image = null;
            sections[idx].imageType = null;
            saveState();
            renderAll();
        });
    }

    const contextInput = card.querySelector('.context-input');
    contextInput.addEventListener('input', () => {
        sections[idx].context = contextInput.value;
        saveState();
    });

    const generatedText = card.querySelector('.generated-text');
    generatedText.addEventListener('input', () => {
        sections[idx].generated = generatedText.value;
        saveState();
    });

    // Paste image into section
    const handlePaste = (e) => {
        const items = e.clipboardData.items;
        for (const item of items) {
            if (item.type.startsWith('image/')) {
                e.preventDefault();
                handleImageFile(item.getAsFile(), idx);
                break;
            }
        }
    };
    dropZone.addEventListener('paste', handlePaste);
    contextInput.addEventListener('paste', handlePaste);

    const genBtn = card.querySelector('.generate-btn');
    genBtn.addEventListener('click', () => generateForSection(idx));

    // Drag and drop reorder
    card.addEventListener('dragstart', (e) => {
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', idx.toString());
    });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
    card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        card.classList.add('drag-over');
    });
    card.addEventListener('dragleave', () => card.classList.remove('drag-over'));
    card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over');
        const fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
        const toIdx = idx;
        if (fromIdx !== toIdx && !isNaN(fromIdx)) {
            const [moved] = sections.splice(fromIdx, 1);
            sections.splice(toIdx, 0, moved);
            saveState();
            renderAll();
        }
    });

    return card;
}

// Image handling
function setupDropZone(zone, idx) {
    zone.addEventListener('click', () => {
        if (sections[idx].image) return;
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.addEventListener('change', () => {
            if (input.files[0]) handleImageFile(input.files[0], idx);
        });
        input.click();
    });

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        // Only handle file drops, not section reorder
        if (e.dataTransfer.types.includes('Files')) {
            zone.classList.add('drag-hover');
        }
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-hover'));
    zone.addEventListener('drop', (e) => {
        zone.classList.remove('drag-hover');
        if (e.dataTransfer.files.length && e.dataTransfer.files[0].type.startsWith('image/')) {
            e.preventDefault();
            e.stopPropagation();
            handleImageFile(e.dataTransfer.files[0], idx);
        }
    });
}

function handleImageFile(file, idx) {
    const reader = new FileReader();
    reader.onload = (e) => {
        sections[idx].image = e.target.result;
        sections[idx].imageType = file.type;
        saveState();
        renderAll();
    };
    reader.readAsDataURL(file);
}

// AI Generation
async function generateForSection(idx) {
    const section = sections[idx];
    if (!section.image) {
        toast('Agregá una imagen primero');
        return;
    }
    if (!apiKey) {
        showModal('apiKeyModal');
        return;
    }

    const btn = sectionsContainer.querySelectorAll('.generate-btn')[idx];
    btn.classList.add('loading');
    btn.textContent = '⏳ Generando...';

    try {
        // Extract base64 and media type from data URL
        const match = section.image.match(/^data:(image\/[^;]+);base64,(.+)$/);
        if (!match) throw new Error('Formato de imagen inválido');

        const mediaType = match[1];
        const base64Data = match[2];

        const userContent = [
            {
                type: 'image',
                source: {
                    type: 'base64',
                    media_type: mediaType,
                    data: base64Data
                }
            },
            {
                type: 'text',
                text: section.context
                    ? `Contexto adicional del usuario: "${section.context}"\n\nGenerá el texto para esta sección del manual de usuario basándote en la imagen y el contexto.`
                    : 'Generá el texto para esta sección del manual de usuario basándote en la imagen.'
            }
        ];

        const res = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'x-api-key': apiKey,
                'anthropic-version': '2023-06-01',
                'anthropic-dangerous-direct-browser-access': 'true',
                'content-type': 'application/json'
            },
            body: JSON.stringify({
                model: 'claude-sonnet-4-5-20250929',
                max_tokens: 1500,
                system: `Sos un redactor técnico profesional especializado en documentación de software. Tu tarea es generar texto para un manual de usuario corporativo.

Reglas:
- Escribí en español neutro/formal, en tercera persona o usando "el usuario"
- El texto debe ser instructivo: describí lo que se ve en la pantalla y los pasos que el usuario debe seguir
- Usá un tono profesional y claro, apto para un cliente empresarial
- Estructurá con pasos numerados cuando corresponda
- No uses markdown ni formato especial, solo texto plano con saltos de línea
- Mencioná los nombres exactos de botones, campos, menús y secciones visibles en la imagen
- Sé conciso pero completo: cubrí todos los elementos relevantes de la pantalla
- No agregues encabezados de sección ni títulos, solo el cuerpo del texto`,
                messages: [{ role: 'user', content: userContent }]
            })
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.error?.message || 'Error en la API');
        }

        const data = await res.json();
        const text = data.content[0].text;
        sections[idx].generated = text;
        saveState();
        renderAll();
        toast('Texto generado correctamente');

    } catch (e) {
        toast('Error: ' + e.message);
        console.error(e);
    } finally {
        btn.classList.remove('loading');
        btn.textContent = '⚡ Generar texto';
    }
}

async function generateAll() {
    const withImages = sections.filter(s => s.image && !s.generated);
    if (withImages.length === 0) {
        toast('No hay secciones pendientes de generar');
        return;
    }
    if (!apiKey) { showModal('apiKeyModal'); return; }

    const btn = $('#generateAllBtn');
    btn.disabled = true;
    btn.textContent = '⏳ Generando...';

    for (let i = 0; i < sections.length; i++) {
        if (sections[i].image && !sections[i].generated) {
            await generateForSection(i);
            // Small delay between calls
            await new Promise(r => setTimeout(r, 500));
        }
    }

    btn.disabled = false;
    btn.textContent = '⚡ Generar todo';
    toast('Generación completa');
}

// Preview & Export
function showPreview() {
    if (sections.length === 0) {
        toast('Agregá secciones primero');
        return;
    }

    const content = $('#previewContent');
    let html = `<h1>${docTitle.value || 'Manual de Usuario'}</h1>`;

    sections.forEach((s, i) => {
        html += `<div class="preview-section">`;
        html += `<h2>Sección ${i + 1}</h2>`;
        if (s.image) html += `<img src="${s.image}" alt="Sección ${i + 1}">`;
        if (s.generated) html += `<div class="preview-text">${escapeHtml(s.generated)}</div>`;
        else html += `<div class="preview-text" style="color:var(--text-muted)">(Sin texto generado)</div>`;
        html += `</div>`;
    });

    content.innerHTML = html;
    showModal('previewModal');
}

function exportPdf() {
    hideModal('previewModal');
    setTimeout(() => window.print(), 300);
}

function copyMarkdown() {
    let md = `# ${docTitle.value || 'Manual de Usuario'}\n\n`;
    sections.forEach((s, i) => {
        md += `## Sección ${i + 1}\n\n`;
        if (s.generated) md += s.generated + '\n\n';
        if (s.image) md += `![Sección ${i + 1}](imagen_seccion_${i + 1})\n\n`;
        md += '---\n\n';
    });

    navigator.clipboard.writeText(md).then(() => toast('Markdown copiado al portapapeles'));
}

// Persistence
function saveState() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(sections));
    } catch (e) {
        // localStorage full — likely too many large images
        console.warn('localStorage full, cannot save state');
        toast('Advertencia: almacenamiento lleno. Considerá exportar y limpiar secciones.');
    }
}

function loadState() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) sections = JSON.parse(saved);
        const title = localStorage.getItem(TITLE_KEY);
        if (title) docTitle.value = title;
    } catch (e) {
        sections = [];
    }
}

// Modals
function showModal(id) {
    document.getElementById(id).classList.add('active');
}
function hideModal(id) {
    document.getElementById(id).classList.remove('active');
}

// Utils
function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 3000);
}

function escapeHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}
