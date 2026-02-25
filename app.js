// State
let sections = [];
let templateConfig = {
    logo: null,
    companyName: '',
    subtitle: '',
    footerText: ''
};
const STORAGE_KEY = 'docgen_sections';
const TITLE_KEY = 'docgen_title';
const TEMPLATE_KEY = 'docgen_template';

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
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('sw.js');
    }
});

function setupEventListeners() {
    $('#addSectionBtn').addEventListener('click', addSection);
    $('#settingsBtn').addEventListener('click', openTemplateModal);
    $('#saveTemplate').addEventListener('click', saveTemplate);
    $('#closeTemplateModal').addEventListener('click', () => hideModal('templateModal'));
    $('#previewBtn').addEventListener('click', showPreview);
    $('#exportWordBtn').addEventListener('click', exportToWord);
    $('#exportWord').addEventListener('click', exportToWord);
    $('#closePreview').addEventListener('click', () => hideModal('previewModal'));
    $('#copyMarkdown').addEventListener('click', copyMarkdown);

    docTitle.addEventListener('input', () => {
        localStorage.setItem(TITLE_KEY, docTitle.value);
    });

    // Logo upload
    const logoUpload = $('#logoUpload');
    logoUpload.addEventListener('click', (e) => {
        if (e.target.id === 'removeLogo') return;
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.addEventListener('change', () => {
            if (input.files[0]) loadLogo(input.files[0]);
        });
        input.click();
    });
    $('#removeLogo').addEventListener('click', (e) => {
        e.stopPropagation();
        templateConfig.logo = null;
        $('#logoPreview').style.display = 'none';
        $('#logoPlaceholder').style.display = '';
        $('#removeLogo').style.display = 'none';
    });

    // Template file upload
    $('#templateUploadZone').addEventListener('click', () => $('#templateFileInput').click());
    $('#templateFileInput').addEventListener('change', (e) => {
        if (e.target.files[0]) {
            $('#templateFileName').textContent = e.target.files[0].name;
            $('#templateUploadZone').classList.add('has-file');
        }
    });

    // Global paste
    document.addEventListener('paste', (e) => {
        const activeEl = document.activeElement;
        if (activeEl.classList.contains('context-input') ||
            activeEl.classList.contains('generated-text') ||
            activeEl.classList.contains('section-title-input') ||
            activeEl.id === 'docTitle' ||
            activeEl.id === 'companyName' ||
            activeEl.id === 'docSubtitle' ||
            activeEl.id === 'footerText') return;

        const items = e.clipboardData.items;
        for (const item of items) {
            if (item.type.startsWith('image/')) {
                e.preventDefault();
                const file = item.getAsFile();
                if (sections.length === 0 || sections[sections.length - 1].image) {
                    addSection();
                }
                handleImageFile(file, sections.length - 1);
                break;
            }
        }
    });
}

// Logo
function loadLogo(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        templateConfig.logo = e.target.result;
        const preview = $('#logoPreview');
        preview.src = e.target.result;
        preview.style.display = '';
        $('#logoPlaceholder').style.display = 'none';
        $('#removeLogo').style.display = 'flex';
    };
    reader.readAsDataURL(file);
}

// Template modal
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

// Sections CRUD
function addSection() {
    sections.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        title: '',
        image: null,
        imageType: null,
        context: '',
        generated: ''
    });
    saveState();
    renderAll();
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

    const titleInput = card.querySelector('.section-title-input');
    titleInput.addEventListener('input', () => {
        sections[idx].title = titleInput.value;
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

// Word Export using docx.js
async function exportToWord() {
    if (sections.length === 0) {
        toast('Agregá secciones primero');
        return;
    }

    if (!window.docx) {
        toast('Error: la librería docx no se cargó. Recargá la página.');
        console.error('window.docx is undefined. Available globals:', Object.keys(window).filter(k => k.toLowerCase().includes('doc')));
        return;
    }

    toast('Generando documento Word...');

    try {
        const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel,
                Header, Footer, PageNumber, NumberFormat, AlignmentType,
                BorderStyle, PageBreak, Bookmark, InternalHyperlink,
                TableOfContents } = window.docx;

        // Prepare header children
        const headerChildren = [];

        if (templateConfig.logo) {
            try {
                const logoData = await dataUrlToArrayBuffer(templateConfig.logo);
                headerChildren.push(new Paragraph({
                    children: [
                        new ImageRun({
                            data: logoData,
                            transformation: { width: 120, height: 40 },
                            type: 'png'
                        })
                    ],
                    alignment: AlignmentType.LEFT
                }));
            } catch (e) {
                console.warn('Could not add logo to header:', e);
            }
        }

        if (templateConfig.companyName) {
            headerChildren.push(new Paragraph({
                children: [new TextRun({
                    text: templateConfig.companyName,
                    bold: true,
                    size: 18,
                    color: '666666'
                })],
                alignment: AlignmentType.LEFT
            }));
        }

        // Prepare footer
        const footerChildren = [];
        if (templateConfig.footerText) {
            footerChildren.push(new Paragraph({
                children: [new TextRun({
                    text: templateConfig.footerText,
                    size: 16,
                    color: '999999',
                    italics: true
                })],
                alignment: AlignmentType.LEFT
            }));
        }
        footerChildren.push(new Paragraph({
            children: [
                new TextRun({ text: 'Página ', size: 16, color: '999999' }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '999999' }),
                new TextRun({ text: ' de ', size: 16, color: '999999' }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: '999999' })
            ],
            alignment: AlignmentType.RIGHT
        }));

        // Build document children
        const docChildren = [];

        // Title page
        docChildren.push(new Paragraph({ spacing: { before: 3000 } }));
        docChildren.push(new Paragraph({
            children: [new TextRun({
                text: docTitle.value || 'Manual de Usuario',
                bold: true,
                size: 56,
                color: '2B2B7B'
            })],
            alignment: AlignmentType.CENTER
        }));

        if (templateConfig.subtitle) {
            docChildren.push(new Paragraph({
                children: [new TextRun({
                    text: templateConfig.subtitle,
                    size: 28,
                    color: '666666'
                })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200 }
            }));
        }

        if (templateConfig.companyName) {
            docChildren.push(new Paragraph({
                children: [new TextRun({
                    text: templateConfig.companyName,
                    size: 24,
                    color: '999999'
                })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 400 }
            }));
        }

        docChildren.push(new Paragraph({
            children: [new TextRun({
                text: new Date().toLocaleDateString('es-AR', { year: 'numeric', month: 'long', day: 'numeric' }),
                size: 20,
                color: '999999'
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 200 }
        }));

        // Page break after title
        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // Table of Contents - title
        docChildren.push(new Paragraph({
            children: [new TextRun({ text: 'Índice de Contenidos', bold: true, size: 32, color: '2B2B7B' })],
            spacing: { after: 300 },
            border: {
                bottom: { style: BorderStyle.SINGLE, size: 2, color: '2B2B7B' }
            }
        }));

        // Auto-generated TOC field (Word will populate with page numbers)
        docChildren.push(new TableOfContents("Índice", {
            hyperlink: true,
            headingStyleRange: "1-2",
        }));

        // Manual TOC entries with internal hyperlinks (visible before updating in Word)
        sections.forEach((s, i) => {
            const title = s.title || `Sección ${i + 1}`;
            const anchorId = `section_${i}`;
            docChildren.push(new Paragraph({
                children: [
                    new InternalHyperlink({
                        anchor: anchorId,
                        children: [
                            new TextRun({
                                text: `${i + 1}. ${title}`,
                                style: 'Hyperlink',
                                size: 22,
                            })
                        ]
                    })
                ],
                spacing: { before: 100, after: 40 }
            }));
        });

        docChildren.push(new Paragraph({ children: [new PageBreak()] }));

        // Sections content
        for (let i = 0; i < sections.length; i++) {
            const s = sections[i];
            const title = s.title || `Sección ${i + 1}`;

            // Section heading with bookmark anchor
            const anchorId = `section_${i}`;
            docChildren.push(new Paragraph({
                children: [
                    new Bookmark({
                        id: anchorId,
                        children: [
                            new TextRun({
                                text: `${i + 1}. ${title}`,
                                bold: true,
                                size: 28,
                                color: '2B2B7B'
                            })
                        ]
                    })
                ],
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 400, after: 200 },
                border: {
                    bottom: { style: BorderStyle.SINGLE, size: 2, color: '2B2B7B' }
                }
            }));

            // Image
            if (s.image) {
                try {
                    const imgData = await dataUrlToArrayBuffer(s.image);
                    const dims = await getImageDimensions(s.image);
                    // Scale to fit page width (max ~500pt = ~6.9 inches)
                    const maxWidth = 500;
                    let w = dims.width;
                    let h = dims.height;
                    if (w > maxWidth) {
                        h = Math.round(h * (maxWidth / w));
                        w = maxWidth;
                    }
                    docChildren.push(new Paragraph({
                        children: [new ImageRun({
                            data: imgData,
                            transformation: { width: w, height: h },
                            type: s.imageType?.includes('png') ? 'png' : 'jpg'
                        })],
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200, after: 200 }
                    }));
                } catch (e) {
                    console.warn('Could not add image for section', i, e);
                }
            }

            // Text content
            if (s.generated) {
                const lines = s.generated.split('\n');
                for (const line of lines) {
                    docChildren.push(new Paragraph({
                        children: [new TextRun({
                            text: line,
                            size: 22,
                            font: 'Calibri'
                        })],
                        spacing: { before: 60, after: 60 }
                    }));
                }
            }

            // Page break between sections (except last)
            if (i < sections.length - 1) {
                docChildren.push(new Paragraph({ children: [new PageBreak()] }));
            }
        }

        const doc = new Document({
            features: {
                updateFields: true
            },
            styles: {
                default: {
                    document: {
                        run: { font: 'Calibri', size: 22 }
                    }
                }
            },
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                        pageNumbers: { start: 1 }
                    }
                },
                headers: headerChildren.length > 0 ? {
                    default: new Header({ children: headerChildren })
                } : undefined,
                footers: {
                    default: new Footer({ children: footerChildren })
                },
                children: docChildren
            }]
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

// Preview
function showPreview() {
    if (sections.length === 0) {
        toast('Agregá secciones primero');
        return;
    }

    const content = $('#previewContent');
    let html = `<h1>${escapeHtml(docTitle.value || 'Manual de Usuario')}</h1>`;

    sections.forEach((s, i) => {
        const title = s.title || `Sección ${i + 1}`;
        html += `<div class="preview-section">`;
        html += `<h2>${i + 1}. ${escapeHtml(title)}</h2>`;
        if (s.image) html += `<img src="${s.image}" alt="${escapeHtml(title)}">`;
        if (s.generated) html += `<div class="preview-text">${escapeHtml(s.generated)}</div>`;
        else html += `<div class="preview-text" style="color:var(--text-muted)">(Sin texto todavía)</div>`;
        html += `</div>`;
    });

    content.innerHTML = html;
    showModal('previewModal');
}

function copyMarkdown() {
    let md = `# ${docTitle.value || 'Manual de Usuario'}\n\n`;
    sections.forEach((s, i) => {
        const title = s.title || `Sección ${i + 1}`;
        md += `## ${i + 1}. ${title}\n\n`;
        if (s.generated) md += s.generated + '\n\n';
        md += '---\n\n';
    });

    navigator.clipboard.writeText(md).then(() => toast('Markdown copiado al portapapeles'));
}

// Persistence
function saveState() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(sections));
    } catch (e) {
        console.warn('localStorage full');
        toast('Almacenamiento lleno. Exportá y eliminá secciones.');
    }
}

function loadState() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) sections = JSON.parse(saved);
        const title = localStorage.getItem(TITLE_KEY);
        if (title) docTitle.value = title;
        const tmpl = localStorage.getItem(TEMPLATE_KEY);
        if (tmpl) templateConfig = JSON.parse(tmpl);
    } catch (e) {
        sections = [];
    }
}

// Modals
function showModal(id) { document.getElementById(id).classList.add('active'); }
function hideModal(id) { document.getElementById(id).classList.remove('active'); }

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

function escapeAttr(str) {
    return (str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function dataUrlToArrayBuffer(dataUrl) {
    return new Promise((resolve) => {
        const binary = atob(dataUrl.split(',')[1]);
        const len = binary.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
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
