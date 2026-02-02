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

let currentDocType = 'medical_record';
let currentAction = null;

// ===== Initialize =====
document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    initDocTypeSelection();
    initButtons();
    initModals();
    initSignatureUploads();
    loadTemplates();
    loadSavedSignatures();
    setDefaultDate();
    initCardNumberSync();
});

// ===== Navigation =====
function initNavigation() {
    navBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabId = btn.dataset.tab;
            navBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            tabContents.forEach(content => {
                content.classList.remove('active');
                if (content.id === `${tabId}Tab`) {
                    content.classList.add('active');
                }
            });
            if (tabId === 'templates') loadTemplates();
        });
    });
}

// ===== Document Type Selection =====
function initDocTypeSelection() {
    docTypeBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const type = btn.dataset.type;
            currentDocType = type;

            docTypeBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            if (type === 'form_100') {
                // თუ სამედიცინო ჩანაწერში არის შევსებული მონაცემი, გადავიტანოთ ფორმა #100-ში
                copyMRToForm100IfNeeded();
                form100.style.display = 'block';
                medicalRecord.style.display = 'none';
            } else {
                form100.style.display = 'none';
                medicalRecord.style.display = 'block';
            }
        });
    });
}

function setDefaultDate() {
    const today = new Date().toISOString().split('T')[0];
    const dateInput = document.querySelector('[name="document_date"]');
    if (dateInput) dateInput.value = today;
}

// ===== Buttons =====
function initButtons() {
    if (saveDocumentBtn) {
        saveDocumentBtn.addEventListener('click', () => {
            currentAction = 'save';
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
            if (confirm('გსურთ ფორმის გასუფთავება?')) {
                clearCurrentForm();
                showToast('გასუფთავდა', 'success');
            }
        });
    }
}

// ===== Form Data =====
function getFormData() {
    const form = currentDocType === 'form_100'
        ? document.getElementById('form100Form')
        : document.getElementById('medicalRecordForm');

    if (!form) return { document_type: currentDocType };

    const formData = new FormData(form);
    const data = { document_type: currentDocType };
    formData.forEach((value, key) => { data[key] = value; });
    return data;
}

// ===== Save =====
async function handleSave(filename) {
    showLoading();
    const data = getFormData();
    data.filename = filename;

    try {
        const response = await fetch('/api/save-document', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();

        if (result.success) {
            showToast('შეინახა!', 'success');
            const link = document.createElement('a');
            link.href = `/api/download/${result.filename}`;
            link.download = result.filename;
            link.click();
        } else {
            showToast(`შეცდომა: ${result.error}`, 'error');
        }
    } catch (error) {
        showToast(`შეცდომა: ${error.message}`, 'error');
    } finally {
        hideLoading();
    }
}

// ===== Print =====
async function handlePrint() {
    showLoading();
    const data = getFormData();

    try {
        const response = await fetch('/api/print-document', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();
        hideLoading();

        if (result.success) {
            const printUrl = `/api/print-page/${result.filename}`;
            showToast('PDF მზადაა!', 'success');

            const win = window.open(printUrl, '_blank');
            if (!win) {
                const link = document.createElement('a');
                link.href = printUrl;
                link.target = '_blank';
                link.click();
            }
        } else {
            showToast(`შეცდომა: ${result.error}`, 'error');
        }
    } catch (error) {
        hideLoading();
        showToast(`შეცდომა: ${error.message}`, 'error');
    }
}

// ===== Templates =====
async function loadTemplates() {
    try {
        const response = await fetch('/api/templates');
        const result = await response.json();
        if (result.success) renderTemplates(result.templates);
    } catch (error) {
        console.error('Templates error:', error);
    }
}

function renderTemplates(templates) {
    if (!templatesGrid) return;
    if (templates.length === 0) {
        templatesGrid.innerHTML = '';
        if (emptyTemplates) emptyTemplates.style.display = 'block';
        return;
    }
    if (emptyTemplates) emptyTemplates.style.display = 'none';

    templatesGrid.innerHTML = templates.map(t => `
        <div class="template-card">
            <div class="template-card-header">
                <div class="template-icon">
                    <i class="fas ${t.type === 'form_100' ? 'fa-file-alt' : 'fa-notes-medical'}"></i>
                </div>
                <div class="template-actions">
                    <button onclick="useTemplate('${t.id}')"><i class="fas fa-check"></i></button>
                    <button class="delete" onclick="deleteTemplate('${t.id}')"><i class="fas fa-trash"></i></button>
                </div>
            </div>
            <h3>${t.name}</h3>
            <span class="template-type">${t.type === 'form_100' ? 'ფორმა №100' : 'სამედიცინო ჩანაწერი'}</span>
        </div>
    `).join('');
}

async function saveTemplate(name) {
    const data = getFormData();
    data.template_name = name;

    try {
        const response = await fetch('/api/templates', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (result.success) {
            showToast('შაბლონი შეინახა!', 'success');
            loadTemplates();
        }
    } catch (error) {
        showToast(`შეცდომა: ${error.message}`, 'error');
    }
}

async function useTemplate(templateId) {
    try {
        const response = await fetch('/api/templates');
        const result = await response.json();
        if (result.success) {
            const template = result.templates.find(t => t.id === templateId);
            if (template) {
                document.querySelector('[data-tab="documents"]').click();
                document.querySelector(`[data-type="${template.data.document_type}"]`).click();
                setTimeout(() => {
                    fillFormWithData(template.data);
                    showToast('შაბლონი ჩაიტვირთა!', 'success');
                }, 100);
            }
        }
    } catch (error) {
        showToast(`შეცდომა: ${error.message}`, 'error');
    }
}

async function deleteTemplate(templateId) {
    if (!confirm('წავშალოთ შაბლონი?')) return;
    try {
        await fetch(`/api/templates/${templateId}`, { method: 'DELETE' });
        showToast('წაიშალა!', 'success');
        loadTemplates();
    } catch (error) {
        showToast(`შეცდომა: ${error.message}`, 'error');
    }
}

function fillFormWithData(data) {
    const form = data.document_type === 'form_100'
        ? document.getElementById('form100Form')
        : document.getElementById('medicalRecordForm');
    if (!form) return;
    Object.keys(data).forEach(key => {
        const input = form.querySelector(`[name="${key}"]`);
        if (input) input.value = data[key];
    });
}

function clearCurrentForm() {
    const form = currentDocType === 'form_100'
        ? document.getElementById('form100Form')
        : document.getElementById('medicalRecordForm');
    if (form) form.reset();
    setDefaultDate();
}
// ===== Medical Record card_number -> Form 100 registration_number =====
function initCardNumberSync() {
    // ვპოულობთ სამედიცინო ჩანაწერის ბარათის ნომრის ველს
    const mrCardNumber = document.querySelector('#medicalRecordForm [name="card_number"]');
    // ვპოულობთ ფორმა №100-ის რეგისტრაციის ველს
    const form100Reg = document.querySelector('#form100Form [name="registration_number"]');

    if (!mrCardNumber || !form100Reg) {
        // თუ რომელიმე არ არსებობს, არაფერს ვაკეთებთ
        return;
    }

    // თუ მომხმარებელი ხელით შეცვლის რეგისტრაციის ნომერს,
    // აღარ გადავაწეროთ ავტომატურად ბარათის ნომრის ცვლილებით
    form100Reg.addEventListener('input', () => {
        form100Reg.dataset.manualEdited = 'true';
    });

    // როდესაც ბარათის ნომერი იცვლება
    mrCardNumber.addEventListener('input', () => {
        const val = mrCardNumber.value;

        // თუ რეგისტრაციის № ცარიელია ან ადრე ავტომატურად იყო შევსებული,
        // ვანახლებთ. თუ ხელით შეცვალეს, აღარ ვეხებით.
        if (!form100Reg.value || form100Reg.dataset.autoFilled === 'true' && form100Reg.dataset.manualEdited !== 'true') {
            form100Reg.value = val;
            form100Reg.dataset.autoFilled = 'true';
        }
    });
}
// ===== Modals =====
function initModals() {
    document.getElementById('closeModal')?.addEventListener('click', closeTemplateModal);
    document.getElementById('cancelTemplate')?.addEventListener('click', closeTemplateModal);
    document.getElementById('confirmSaveTemplate')?.addEventListener('click', () => {
        const name = document.getElementById('templateName')?.value.trim();
        if (!name) { showToast('შეიყვანეთ სახელი', 'error'); return; }
        saveTemplate(name);
        closeTemplateModal();
    });

    document.getElementById('closeFilenameModal')?.addEventListener('click', closeFilenameModal);
    document.getElementById('cancelFilename')?.addEventListener('click', closeFilenameModal);
    document.getElementById('confirmFilename')?.addEventListener('click', () => {
        const filename = document.getElementById('docFilename')?.value.trim();
        if (!filename) { showToast('შეიყვანეთ სახელი', 'error'); return; }
        if (currentAction === 'save') handleSave(filename);
        closeFilenameModal();
    });

    templateModal?.addEventListener('click', e => { if (e.target === templateModal) closeTemplateModal(); });
    filenameModal?.addEventListener('click', e => { if (e.target === filenameModal) closeFilenameModal(); });
}

function openTemplateModal() {
    templateModal?.classList.add('active');
    document.getElementById('templateName').value = '';
}

function closeTemplateModal() {
    templateModal?.classList.remove('active');
}

function openFilenameModal() {
    const patientName = currentDocType === 'form_100'
        ? document.getElementById('patient_name')?.value
        : document.getElementById('mr_patient_name')?.value;
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('docFilename').value = patientName
        ? `${patientName.replace(/\s+/g, '_')}_${today}`
        : `document_${today}`;
    filenameModal?.classList.add('active');
}

function closeFilenameModal() {
    filenameModal?.classList.remove('active');
}

// ===== Signatures =====
function initSignatureUploads() {
    // ფორმა #100
    document.getElementById('doctorSigInput')?.addEventListener('change', e => handleSigUpload(e, 'doctor'));
    document.getElementById('stampInput')?.addEventListener('change', e => handleSigUpload(e, 'stamp'));
    document.getElementById('headSigInput')?.addEventListener('change', e => handleSigUpload(e, 'head'));

    // სამედიცინო ჩანაწერი
    document.getElementById('mrDoctorSigInput')?.addEventListener('change', e => handleSigUpload(e, 'mrDoctor'));
}

function handleSigUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) { showToast('ფაილი ძალიან დიდია', 'error'); return; }

    const reader = new FileReader();
    reader.onload = function(e) {
        const base64 = e.target.result;

        const previewMap = {
            'doctor': 'doctorSigPreview',
            'stamp': 'stampPreview',
            'head': 'headSigPreview',
            'mrDoctor': 'mrDoctorSigPreview'
        };
        const dataMap = {
            'doctor': 'doctorSigData',
            'stamp': 'stampData',
            'head': 'headSigData',
            'mrDoctor': 'mrDoctorSigData'
        };

        const preview = document.getElementById(previewMap[type]);
        if (preview) preview.innerHTML = `<img src="${base64}" alt="sig">`;

        const dataInput = document.getElementById(dataMap[type]);
        if (dataInput) dataInput.value = base64;

        showToast('ატვირთულია!', 'success');
    };
    reader.readAsDataURL(file);
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
            if (s.doctor) {
                document.getElementById('doctorSigPreview').innerHTML = `<img src="${s.doctor}">`;
                document.getElementById('doctorSigData').value = s.doctor;
            }
            if (s.stamp) {
                document.getElementById('stampPreview').innerHTML = `<img src="${s.stamp}">`;
                document.getElementById('stampData').value = s.stamp;
            }
            if (s.head) {
                document.getElementById('headSigPreview').innerHTML = `<img src="${s.head}">`;
                document.getElementById('headSigData').value = s.head;
            }
        }
    } catch (e) { console.error(e); }
}

// ===== UI =====
function showToast(message, type = 'success') {
    if (!toast || !toastMessage) return;
    toast.className = `toast ${type}`;
    toastMessage.textContent = message;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}

function showLoading() { loadingOverlay?.classList.add('active'); }
function hideLoading() { loadingOverlay?.classList.remove('active'); }

// ===== Global =====
window.useTemplate = useTemplate;
window.deleteTemplate = deleteTemplate;
window.clearSignature = clearSignature;

// ===== Medical Record -> Form 100 Auto-Copy =====
function copyMRToForm100IfNeeded() {
    const mrForm = document.getElementById('medicalRecordForm');
    const f100Form = document.getElementById('form100Form');
    if (!mrForm || !f100Form) return;

    const mrPatient = mrForm.querySelector('[name="patient_name"]');
    if (!mrPatient || !mrPatient.value.trim()) {
        // თუ პაციენტის სახელი არ არის შევსებული, ვთვლით, რომ ფორმა ცარიელია
        return;
    }

    copyMRToForm100();
    showToast('მონაცემები სამედიცინო ჩანაწერიდან ფორმა №100-ში გადმოწერილია', 'success');
}

function copyMRToForm100() {
    const mr = document.getElementById('medicalRecordForm');
    const f100 = document.getElementById('form100Form');
    if (!mr || !f100) return;

    // პაციენტის სახელი
    const mrPatient = mr.querySelector('[name="patient_name"]')?.value || '';
    if (mrPatient && f100.querySelector('[name="patient_name"]')) {
        f100.querySelector('[name="patient_name"]').value = mrPatient;
    }

    // ანამნეზი (თუ ორივე გაქვს, MR-ის ანამნეზი გადადის ფორმა100-ის ანამნეზში)
    const mrComplaints = mr.querySelector('[name="complaints"]')?.value || '';
    const mrAnam = mr.querySelector('[name="anamnesis"]')?.value || '';
    const anamTarget = f100.querySelector('[name="anamnesis"]');
    if (anamTarget) {
        if (mrAnam) anamTarget.value = mrAnam;
        else if (mrComplaints) anamTarget.value = mrComplaints;
    }

    // დიაგნოზი (ICD + აღწერა ან წინასწარი დიაგნოზი)
    const icd = mr.querySelector('[name="icd_code"]')?.value || '';
    const diagDesc = mr.querySelector('[name="diagnosis_description"]')?.value || '';
    const prelimDiag = mr.querySelector('[name="preliminary_diagnosis"]')?.value || '';
    const mainDiagTarget = f100.querySelector('[name="main_diagnosis"]');
    if (mainDiagTarget) {
        let text = '';
        if (diagDesc) text += diagDesc;
        if (icd) text += (text ? ' ' : '') + icd;
        if (!text && prelimDiag) text = prelimDiag;
        mainDiagTarget.value = text;
    }

    // ვიტალები -> მიღების ვიტალები ფორმა100-ში
    const t = mr.querySelector('[name="temperature"]')?.value || '';
    const bp = mr.querySelector('[name="blood_pressure"]')?.value || '';
    const hr = mr.querySelector('[name="heart_rate"]')?.value || '';
    const rr = mr.querySelector('[name="respiratory_rate"]')?.value || '';
    const spo2 = mr.querySelector('[name="spo2"]')?.value || '';

    if (f100.querySelector('[name="admission_temp"]') && t) f100.querySelector('[name="admission_temp"]').value = t;
    if (f100.querySelector('[name="admission_bp"]') && bp) f100.querySelector('[name="admission_bp"]').value = bp;
    if (f100.querySelector('[name="admission_hr"]') && hr) f100.querySelector('[name="admission_hr"]').value = hr;
    if (f100.querySelector('[name="admission_rr"]') && rr) f100.querySelector('[name="admission_rr"]').value = rr;
    if (f100.querySelector('[name="admission_spo2"]') && spo2) f100.querySelector('[name="admission_spo2"]').value = spo2;

    // ზოგადი მდგომარეობა -> form100-ის course_type ან discharge_condition (სურვილისამებრ)
    // მაგალითად, თუ MR-ში general_condition არის "საშუალო სიმძიმის", form100-ში შეგვიძლია ჩავსვათ discharge_condition-ში
    const mrGeneral = mr.querySelector('[name="general_condition"]')?.value || '';
    const dischargeCondTarget = f100.querySelector('[name="discharge_condition"]');
    if (mrGeneral && dischargeCondTarget && !dischargeCondTarget.value) {
        dischargeCondTarget.value = mrGeneral;
    }
}

// გლობალური რომ სხვა ადგილებიდანაც გამოიძახო (თუ დაგჭირდება)
window.copyMRToForm100 = copyMRToForm100;