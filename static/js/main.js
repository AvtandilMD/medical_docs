// ===== DOM Elements =====
const navBtns = document.querySelectorAll('.nav-btn');
const tabContents = document.querySelectorAll('.tab-content');
const docTypeBtns = document.querySelectorAll('.doc-type-btn');
const form100 = document.getElementById('form100');
const medicalRecord = document.getElementById('medicalRecord');

const saveDocumentBtn = document.getElementById('saveDocument');
const printDocumentBtn = document.getElementById('printDocument');
const saveAsTemplateBtn = document.getElementById('saveAsTemplate');
const clearFormBtn = document.getElementById('clearForm');

const templateModal = document.getElementById('templateModal');
const filenameModal = document.getElementById('filenameModal');

const toast = document.getElementById('toast');
const toastMessage = document.getElementById('toastMessage');
const loadingOverlay = document.getElementById('loadingOverlay');
const templatesGrid = document.getElementById('templatesGrid');
const emptyTemplates = document.getElementById('emptyTemplates');

let currentDocType = 'medical_record'; // საწყისად სამედიცინო ჩანაწერი

// ===== Init =====
document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    initDocTypeSelection();
    initButtons();
    initSearch();
    initModals();
    setDefaultDate();
    initDateTimePickers(); // ← აქ ხდება picker-ების ჩართვა
    initCardNumberSync();
    initSignatureUploads(); // ← ხელმოწერების ატვირთვაც ჩართულია
    loadTemplates();
    loadSavedSignatures(); // ← შენახული ხელმოწერების ჩატვირთვა
});

// ===== Navigation (Documents / Templates) =====
function initNavigation() {
    navBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabId = btn.dataset.tab;
            navBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            tabContents.forEach(c => {
                c.classList.remove('active');
                if (c.id === `${tabId}Tab`) c.classList.add('active');
            });

            if (tabId === 'templates') loadTemplates();
        });
    });
}

// ===== DocType Selection (Sidebar) =====
function initDocTypeSelection() {
    // დარწმუნდით, რომ სწორი ფორმა ჩანს თავიდან
    if (currentDocType === 'medical_record') {
        if (medicalRecord) medicalRecord.style.display = 'block';
        // form100 არ ვმალავთ სრულად, რომ ერთ გვერდზე იყოს (როგორც გინდოდა)
        // ან თუ გინდა გადართვა იყოს:
        // if (form100) form100.style.display = 'none';
    }

    docTypeBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.type;
            currentDocType = type;

            docTypeBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            // სქროლი შესაბამის ფორმასთან
            if (type === 'form_100' && form100) {
                form100.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
            if (type === 'medical_record' && medicalRecord) {
                medicalRecord.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        });
    });
}

function setDefaultDate() {
    const dateInput = document.querySelector('[name="document_date"]');
    if (dateInput && !dateInput.value) {
        // დღევანდელი თარიღი Y-m-d ფორმატში
        dateInput.value = new Date().toISOString().split('T')[0];
    }
}

// ===== Flatpickr (კალენდარი) =====
function initDateTimePickers() {
    if (typeof flatpickr === 'undefined') {
        console.warn('flatpickr not loaded');
        return;
    }

    // ქართული ენა
    if (flatpickr.l10ns && flatpickr.l10ns.ka) {
        flatpickr.localize(flatpickr.l10ns.ka);
    }

    // 1. თარიღი + დრო (24-საათიანი)
    // კლასი: .datetime-24
    flatpickr('.datetime-24', {
        enableTime: true,
        time_24hr: true,
        dateFormat: 'Y-m-d H:i',
        altInput: true,
        altFormat: 'd.m.Y H:i',
        allowInput: true
    });

    // 2. მხოლოდ თარიღი (საათის გარეშე)
    // კლასი: .date-only
    flatpickr('.date-only', {
        enableTime: false,
        dateFormat: 'Y-m-d',
        altInput: true,
        altFormat: 'd.m.Y',
        allowInput: true
    });
}

// ===== Card Number -> Registration Number Sync =====
function initCardNumberSync() {
    const mrCardNumber = document.querySelector('#medicalRecordForm [name="card_number"]');
    const form100Reg   = document.querySelector('#form100Form [name="registration_number"]');

    if (!mrCardNumber || !form100Reg) return;

    form100Reg.addEventListener('input', () => {
        form100Reg.dataset.manualEdited = 'true';
    });

    mrCardNumber.addEventListener('input', () => {
        const val = mrCardNumber.value;
        if (!form100Reg.value || (form100Reg.dataset.autoFilled === 'true' && form100Reg.dataset.manualEdited !== 'true')) {
            form100Reg.value = val;
            form100Reg.dataset.autoFilled = 'true';
        }
    });
}

// ===== Signatures (ატვირთვა / გასუფთავება) =====
function initSignatureUploads() {
    // ფორმა #100
    const docInput = document.getElementById('doctorSigInput');
    if (docInput) docInput.addEventListener('change', e => handleSigUpload(e, 'doctor'));

    const stampInput = document.getElementById('stampInput');
    if (stampInput) stampInput.addEventListener('change', e => handleSigUpload(e, 'stamp'));

    const headInput = document.getElementById('headSigInput');
    if (headInput) headInput.addEventListener('change', e => handleSigUpload(e, 'head'));

    // სამედიცინო ჩანაწერი
    const mrDocInput = document.getElementById('mrDoctorSigInput');
    if (mrDocInput) mrDocInput.addEventListener('change', e => handleSigUpload(e, 'mrDoctor'));
}

function handleSigUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
        showToast('ფაილი ძალიან დიდია (მაქს 2MB)', 'error');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const base64 = e.target.result;

        // ელემენტების ID-ები
        const map = {
            'doctor': { preview: 'doctorSigPreview', data: 'doctorSigData' },
            'stamp': { preview: 'stampPreview', data: 'stampData' },
            'head': { preview: 'headSigPreview', data: 'headSigData' },
            'mrDoctor': { preview: 'mrDoctorSigPreview', data: 'mrDoctorSigData' }
        };

        const ids = map[type];
        if (ids) {
            const preview = document.getElementById(ids.preview);
            if (preview) preview.innerHTML = `<img src="${base64}" alt="sig" style="max-width:100%; max-height:100%;">`;

            const dataInput = document.getElementById(ids.data);
            if (dataInput) dataInput.value = base64;

            // სერვერზე შენახვა (სურვილისამებრ, რომ შემდეგ ჯერზეც დარჩეს)
            uploadSignatureToServer(file, type);
        }
        showToast('ხელმოწერა ატვირთულია!', 'success');
    };
    reader.readAsDataURL(file);
}

async function uploadSignatureToServer(file, type) {
    const formData = new FormData();
    formData.append('file', file);
    // ტიპების გაერთიანება სერვერისთვის (mrDoctor -> doctor)
    let serverType = type;
    if (type === 'mrDoctor') serverType = 'doctor';

    formData.append('type', serverType);

    try {
        await fetch('/api/upload-signature', {
            method: 'POST',
            body: formData
        });
    } catch (e) {
        console.error('Signature upload error:', e);
    }
}

function clearSignature(type) {
    const map = {
        'doctor': { preview: 'doctorSigPreview', data: 'doctorSigData', input: 'doctorSigInput' },
        'stamp': { preview: 'stampPreview', data: 'stampData', input: 'stampInput' },
        'head': { preview: 'headSigPreview', data: 'headSigData', input: 'headSigInput' },
        'mrDoctor': { preview: 'mrDoctorSigPreview', data: 'mrDoctorSigData', input: 'mrDoctorSigInput' }
    };

    const ids = map[type];
    if (!ids) return;

    const preview = document.getElementById(ids.preview);
    if (preview) preview.innerHTML = '<span>ატვირთეთ</span>';

    const data = document.getElementById(ids.data);
    if (data) data.value = '';

    const input = document.getElementById(ids.input);
    if (input) input.value = '';
}

async function loadSavedSignatures() {
    try {
        const response = await fetch('/api/get-signatures');
        const result = await response.json();
        if (result.success && result.signatures) {
            const s = result.signatures;
            // Form 100
            if (s.doctor) setSig('doctor', s.doctor);
            if (s.stamp) setSig('stamp', s.stamp);
            if (s.head) setSig('head', s.head);
            // MR
            if (s.doctor) setSig('mrDoctor', s.doctor);
        }
    } catch (e) { console.error(e); }
}

function setSig(type, base64) {
    const map = {
        'doctor': { preview: 'doctorSigPreview', data: 'doctorSigData' },
        'stamp': { preview: 'stampPreview', data: 'stampData' },
        'head': { preview: 'headSigPreview', data: 'headSigData' },
        'mrDoctor': { preview: 'mrDoctorSigPreview', data: 'mrDoctorSigData' }
    };
    const ids = map[type];
    if (ids) {
        const p = document.getElementById(ids.preview);
        if (p) p.innerHTML = `<img src="${base64}" style="max-width:100%; max-height:100%;">`;
        const d = document.getElementById(ids.data);
        if (d) d.value = base64;
    }
}

// ===== Buttons =====
function initButtons() {
    if (saveDocumentBtn) {
        saveDocumentBtn.addEventListener('click', () => {
            openFilenameModal();
        });
    }
    if (printDocumentBtn) {
        printDocumentBtn.addEventListener('click', handlePrint);
    }
    if (saveAsTemplateBtn) {
        saveAsTemplateBtn.addEventListener('click', openTemplateModal);
    }
    if (clearFormBtn) {
        clearFormBtn.addEventListener('click', () => {
            if (confirm('ნამდვილად გსურთ ფორმის გასუფთავება?')) {
                clearCurrentForm();
                showToast('ფორმა გასუფთავდა', 'success');
            }
        });
    }
}

// ===== Form Data =====
function getFormData() {
    const form = currentDocType === 'form_100'
        ? document.getElementById('form100Form')
        : document.getElementById('medicalRecordForm');

    const data = { document_type: currentDocType };
    if (!form) return data;

    const fd = new FormData(form);
    fd.forEach((val, key) => { data[key] = val; });
    return data;
}

// ===== Save =====
async function handleSave(filename) {
    showLoading();
    const data = getFormData();
    data.filename = filename;

    try {
        const resp = await fetch('/api/save-document', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await resp.json();
        hideLoading();

        if (!result.success) {
            showToast(`შეცდომა: ${result.error}`, 'error');
            return;
        }

        // შეტყობინება
        if (result.is_pdf) {
            showToast('PDF დოკუმენტი შენახულია!', 'success');
        } else {
            showToast('დოკუმენტი შენახულია (DOCX)!', 'warning');
        }

        // გადმოწერა
        const link = document.createElement('a');
        link.href = `/api/download/${result.filename}`;
        link.download = result.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

    } catch (e) {
        hideLoading();
        showToast(`შეცდომა: ${e.message}`, 'error');
    }
}

// ===== Print =====
async function handlePrint() {
    showLoading();
    const data = getFormData();

    try {
        const resp = await fetch('/api/print-document', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await resp.json();
        hideLoading();

        if (!result.success) {
            showToast(`შეცდომა: ${result.error}`, 'error');
            return;
        }

        if (result.is_pdf) {
            const printUrl = `/api/print-page/${result.filename}`;
            showToast('PDF მზადაა ბეჭდვისთვის', 'success');

            const win = window.open(
                printUrl,
                'PrintWindow_' + Date.now(),
                'width=900,height=700,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes'
            );
            if (!win) {
                const link = document.createElement('a');
                link.href = printUrl;
                link.target = '_blank';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
        } else {
            showToast('PDF ვერ შეიქმნა, იტვირთება DOCX.', 'warning');
            const link = document.createElement('a');
            link.href = `/api/download/${result.filename}`;
            link.download = result.filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

    } catch (e) {
        hideLoading();
        showToast(`შეცდომა: ${e.message}`, 'error');
    }
}

// ===== Templates =====
async function loadTemplates() {
    try {
        const resp = await fetch('/api/templates');
        const result = await resp.json();
        if (result.success) renderTemplates(result.templates);
    } catch (e) {
        console.error('Templates error', e);
    }
}

function renderTemplates(templates) {
    if (!templatesGrid) return;

    if (!templates || templates.length === 0) {
        templatesGrid.innerHTML = '';
        if (emptyTemplates) emptyTemplates.style.display = 'block';
        return;
    }
    if (emptyTemplates) emptyTemplates.style.display = 'none';

    templatesGrid.innerHTML = templates.map(t => `
        <div class="template-card" data-template-id="${t.id}">
            <div class="template-card-header">
                <div class="template-icon">
                    <i class="fas ${t.data.document_type === 'form_100' ? 'fa-file-alt' : 'fa-notes-medical'}"></i>
                </div>
                <div class="template-actions">
                    <button onclick="useTemplate('${t.id}')" title="გამოყენება">
                        <i class="fas fa-check"></i>
                    </button>
                    <button class="delete" onclick="deleteTemplate('${t.id}')" title="წაშლა">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </div>
            <h3>${t.name}</h3>
            <span class="template-type">
                ${t.data.document_type === 'form_100' ? 'ფორმა №100' : 'სამედიცინო ჩანაწერი'}
            </span>
        </div>
    `).join('');
}

async function saveTemplate(name) {
    const data = getFormData();
    data.template_name = name;

    try {
        const resp = await fetch('/api/templates', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await resp.json();
        if (result.success) {
            showToast('შაბლონი შენახულია!', 'success');
            loadTemplates();
        } else {
            showToast(`შეცდომა: ${result.error}`, 'error');
        }
    } catch (e) {
        showToast(`შეცდომა: ${e.message}`, 'error');
    }
}

async function useTemplate(templateId) {
    try {
        const resp = await fetch('/api/templates');
        const result = await resp.json();
        if (!result.success) return;

        const t = result.templates.find(x => x.id === templateId);
        if (!t) return;

        const docType = t.data.document_type || 'form_100';
        currentDocType = docType;

        const btn = document.querySelector(`.doc-type-btn[data-type="${docType}"]`);
        if (btn) btn.click();

        // მცირე შეყოვნება, რომ სქროლი და ტაბი შეიცვალოს
        setTimeout(() => {
            const form = docType === 'form_100'
                ? document.getElementById('form100Form')
                : document.getElementById('medicalRecordForm');
            if (!form) return;

            Object.keys(t.data).forEach(key => {
                if (['document_type', 'template_name', 'created'].includes(key)) return;
                const el = form.querySelector(`[name="${key}"]`);
                // Flatpickr-ის შემთხვევაში, setDate მეთოდია სასურველი, მაგრამ value-ც მუშაობს altInput-თან
                if (el) {
                    if (el._flatpickr) {
                        el._flatpickr.setDate(t.data[key]);
                    } else {
                        el.value = t.data[key];
                    }
                }
            });
            showToast('შაბლონი ჩაიტვირთა!', 'success');
        }, 100);

    } catch (e) {
        showToast(`შეცდომა: ${e.message}`, 'error');
    }
}

async function deleteTemplate(templateId) {
    if (!confirm('ნამდვილად გსურთ შაბლონის წაშლა?')) return;
    try {
        const resp = await fetch(`/api/templates/${templateId}`, { method: 'DELETE' });
        const result = await resp.json();
        if (result.success) {
            showToast('შაბლონი წაიშალა!', 'success');
            loadTemplates();
        } else {
            showToast(`შეცდომა: ${result.error}`, 'error');
        }
    } catch (e) {
        showToast(`შეცდომა: ${e.message}`, 'error');
    }
}

// ===== Modals =====
function initModals() {
    document.getElementById('closeModal')?.addEventListener('click', closeTemplateModal);
    document.getElementById('cancelTemplate')?.addEventListener('click', closeTemplateModal);
    document.getElementById('confirmSaveTemplate')?.addEventListener('click', () => {
        const name = document.getElementById('templateName')?.value.trim();
        if (!name) {
            showToast('შეიყვანეთ შაბლონის სახელი', 'error');
            return;
        }
        saveTemplate(name);
        closeTemplateModal();
    });

    document.getElementById('closeFilenameModal')?.addEventListener('click', closeFilenameModal);
    document.getElementById('cancelFilename')?.addEventListener('click', closeFilenameModal);
    document.getElementById('confirmFilename')?.addEventListener('click', () => {
        const filename = document.getElementById('docFilename')?.value.trim();
        if (!filename) {
            showToast('შეიყვანეთ ფაილის სახელი', 'error');
            return;
        }
        handleSave(filename);
        closeFilenameModal();
    });

    templateModal?.addEventListener('click', e => {
        if (e.target === templateModal) closeTemplateModal();
    });
    filenameModal?.addEventListener('click', e => {
        if (e.target === filenameModal) closeFilenameModal();
    });
}

function openTemplateModal() {
    if (templateModal) {
        templateModal.classList.add('active');
        const input = document.getElementById('templateName');
        if (input) { input.value = ''; input.focus(); }
    }
}

function closeTemplateModal() {
    if (templateModal) templateModal.classList.remove('active');
}

function openFilenameModal() {
    const patientName = currentDocType === 'form_100'
        ? document.getElementById('patient_name')?.value
        : document.getElementById('mr_patient_name')?.value;

    const today = new Date().toISOString().split('T')[0];
    const suggested = patientName
        ? `${patientName.replace(/\s+/g, '_')}_${today}`
        : `document_${today}`;

    const input = document.getElementById('docFilename');
    if (input) input.value = suggested;

    if (filenameModal) filenameModal.classList.add('active');
}

function closeFilenameModal() {
    if (filenameModal) filenameModal.classList.remove('active');
}

// ===== Utils =====
function clearCurrentForm() {
    const form = currentDocType === 'form_100'
        ? document.getElementById('form100Form')
        : document.getElementById('medicalRecordForm');
    if (form) form.reset();
}

function showToast(message, type = 'success') {
    if (!toast || !toastMessage) return;
    toast.className = `toast ${type}`;
    toastMessage.textContent = message;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}

function showLoading() {
    if (loadingOverlay) loadingOverlay.classList.add('active');
}

function hideLoading() {
    if (loadingOverlay) loadingOverlay.classList.remove('active');
}

// ===== Search System =====
function initSearch() {
    const searchInput = document.getElementById('patientSearch');
    const resultsBox = document.getElementById('searchResults');
    let debounceTimer;

    if (!searchInput || !resultsBox) return;

    searchInput.addEventListener('input', (e) => {
        clearTimeout(debounceTimer);
        const query = e.target.value.trim();

        if (query.length < 2) {
            resultsBox.style.display = 'none';
            return;
        }

        debounceTimer = setTimeout(() => performSearch(query), 300);
    });

    // დახურვა სხვაგან დაწკაპუნებისას
    document.addEventListener('click', (e) => {
        if (!searchInput.contains(e.target) && !resultsBox.contains(e.target)) {
            resultsBox.style.display = 'none';
        }
    });
}

async function performSearch(query) {
    const resultsBox = document.getElementById('searchResults');
    resultsBox.innerHTML = '<div style="padding:10px; text-align:center; color:#666;">ეძებს...</div>';
    resultsBox.style.display = 'block';

    try {
        const resp = await fetch(`/api/search-patients?q=${encodeURIComponent(query)}`);
        const data = await resp.json();

        if (!data.success || data.results.length === 0) {
            resultsBox.innerHTML = '<div style="padding:10px; text-align:center; color:#999;">ვერ მოიძებნა</div>';
            return;
        }

        resultsBox.innerHTML = data.results.map(item => {
            if (item.type === 'document') {
                return `
                    <div class="search-item" onclick="window.open('${item.path}', '_blank')">
                        <h4><i class="fas fa-file-word"></i> ${item.name}</h4>
                        <p>
                            <span class="search-tag tag-doc">DOCX</span>
                            <span>${item.date}</span>
                        </p>
                    </div>
                `;
            } else {
                return `
                    <div class="search-item" onclick="useTemplate('${item.id}')">
                        <h4><i class="fas fa-user-injured"></i> ${item.patient}</h4>
                        <p>
                            <span class="search-tag tag-tmpl">შაბლონი</span>
                            <span>${item.date}</span>
                        </p>
                    </div>
                `;
            }
        }).join('');

    } catch (e) {
        console.error(e);
        resultsBox.innerHTML = '<div style="padding:10px; text-align:center; color:red;">შეცდომა</div>';
    }
}
function copyToEHR() {
    const data = getFormData(); // ეს ფუნქცია უკვე გაქვთ, იღებს ყველა მონაცემს

    // ვირჩევთ მხოლოდ იმას, რაც გვჭირდება
    const exportData = {
        name: data.patient_name || "",
        personal_id: data.personal_id || "",
        diagnosis: data.main_diagnosis || "",
        anamnesis: data.anamnesis || ""
    };

    // ვაკოპირებთ ბუფერში
    navigator.clipboard.writeText(JSON.stringify(exportData)).then(() => {
        showToast("მონაცემები დაკოპირდა! ახლა გახსენით EHR და დააჭირეთ 'შევსებას'", "success");
        // სურვილისამებრ, ავტომატურად გახსნას საიტი:
        window.open('https://ehr.moh.gov.ge/index.php', '_blank');
    });
}

// ===== Global for templates =====
window.useTemplate = useTemplate;
window.deleteTemplate = deleteTemplate;
window.clearSignature = clearSignature;