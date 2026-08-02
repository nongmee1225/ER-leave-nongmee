const fs = require('fs');
const path = require('path');

const code = fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, 'Index.html'), 'utf8');

new Function(code);

const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)]
  .map(match => match[1])
  .join('\n');
new Function(scripts);

const checks = [
  ['unsupported server favicon disabled', !code.includes('APP_FAVICON_DATA_URL') && !code.includes('.setFaviconUrl(')],
  ['announcement defaults', code.includes('DEFAULT_ANNOUNCEMENT_CONTENT') && code.includes('AnnouncementUpdatedAt')],
  ['announcement backend', code.includes('function getAnnouncement()') && code.includes('function updateAnnouncement(')],
  ['announcement reader modal', html.includes('id="announcementModal"') && html.includes('openAnnouncementModal()')],
  ['announcement admin editor', html.includes('id="announcementAdminContent"') && html.includes("openPinModal('updateAnnouncement')")],
  ['home screen app metadata', html.includes('viewport-fit=cover') && html.includes('apple-mobile-web-app-capable') && html.includes('id="appManifestLink"') && html.includes('id="appTouchIcon"')],
  ['install app UI removed', !html.includes('id="installAppButton"') && !html.includes('id="installModal"') && !html.includes('function openInstallModal()')],
  ['system center backend', code.includes('function getSystemCenterSettings(') && code.includes('function updateSystemCenterSettings(')],
  ['leave type config schema', code.includes('function ensureLeaveTypeConfigSchema_(') && code.includes('function saveLeaveTypeConfig(')],
  ['fixed quota data retained', code.includes('function saveLeaveQuotaConfig(') && !html.includes('id="quotaConfigList"')],
  ['settings audit retained in background', code.includes('function auditSettingChange_(') && !html.includes('id="settingsAudit"')],
  ['leave template locked to standard', code.includes('function getDocumentSettings_(') && !html.includes('id="tplFont"') && !html.includes('พรีวิวแม่แบบ')],
  ['form data fn', code.includes('function getLeaveFormData')],
  ['PDF preview modal', html.includes('id="printModal"') && html.includes('openPrintableLeaveForm') && html.includes('exportCurrentLeavePdf')],
  ['pdf button', html.includes('openPrintableLeaveForm') && html.includes('ใบลา PDF')],
  ['print css', html.includes('@media print')],
  ['self signature upload', html.includes('id="mySignatureFile"') && html.includes('uploadMySignature')],
  ['new staff signature upload', html.includes('id="newStaffSignatureFile"') && html.includes('ใช้รูปนี้เป็นลายเซ็นเจ้าหน้าที่ใหม่')],
  ['admin staff signature upload', html.includes('id="sigfile-') && html.includes('uploadStaffSignature')],
  ['supervisor signature upload', html.includes('id="supervisorSignatureFile"') && html.includes('uploadSupervisorSignature')],
  ['approver signature upload', html.includes('id="approverSignatureFile"') && html.includes('uploadApproverSignature')],
  ['staff tab separated', html.includes("showTab('staff')") && html.includes('id="tab-staff"') && html.includes('arrangeStaffTab_()')],
  ['admin tab access gate', code.includes('function verifyAdminAccess(') && html.includes("openPinModal('unlockAdminTab')") && html.includes("pendingAction === 'unlockAdminTab'")],
  ['admin session lock', html.includes('function lockAdminTab(') && html.includes('ออกจากโหมดแอดมิน')],
  ['preview without PIN', html.includes('function previewRecordWithoutPin()') && html.includes('previewLeaveFormData(rec.staffName, rec.leaveType, rec.start, rec.end, rec.isHalfDay, rec.halfDayPeriod, rec.reason)') && !code.includes("if (!isAdmin && !isSup && !isSelf) return { success: false, message: 'PIN ไม่ถูกต้อง' };\n  if (!staffName || !leaveType")],
  ['simplified admin settings', html.includes('<h3>ข้อมูลหน่วยงาน</h3>') && !html.includes('จัดการประเภทลาและสิทธิ์') && !html.includes('ชื่อผู้แก้ไขการตั้งค่า') && !html.includes('id="sysFiscalYear"') && !html.includes('id="sysReminderDays"')],
  ['admin staff list reuses unlocked session', html.includes('function loadAllStaffForAdmin(') && html.includes('.getAllStaffList(lastAdminPin)') && !html.includes("openPinModal('loadStaff')") && !html.includes("pendingAction === 'loadStaff'")],
  ['preview toolbar separates unsaved preview from saved PDF export', html.includes('setPdfZoom(-0.1)') && html.includes('setPdfZoom(0.1)') && html.includes('id="btnExportPdf"') && html.includes("document.getElementById('btnExportPdf').hidden = true;") && html.includes("document.getElementById('btnExportPdf').hidden = false;") && !html.includes('window.print()')],
  ['private profile session clears', html.includes('function clearMyProfileSession(') && html.includes('5 * 60 * 1000') && html.includes("activeTabName === 'staff'")],
  ['calendar entries deduplicated', code.includes('seenCalendarEntries') && code.includes('dedupeKey')],
  ['calendar duplicate defense on both layers', code.includes('calendarStaffKey') && code.includes('calendarTypeKey') && html.includes('function uniqueCalendarEntries_(')],
  ['important action colors', html.includes('button.primary-action') && html.includes('button.accent-action')],
  ['no immediate LINE on save', !code.includes('บันทึกการลาใหม่') && !code.includes('retroactiveLabel = startDate < todayIso')],
  ['scheduled LINE reminder fixed at one day', code.includes('const reminderDays = 1;') && code.includes('if (todayStr >= start && todayStr <= end)')],
  ['Thai system dates', code.includes('function formatSystemDateTH_(') && html.includes('function formatSystemDateTH(') && html.includes("input.setAttribute('lang', 'th')")],
  ['phone and PIN stored as text', code.includes("setNumberFormat('@')") && code.includes('writePlainText_(sheet.getRange(addedRow, 6), personalPin)') && code.includes('writePlainText_(sheet.getRange(addedRow, 8), phone)')],
  ['official paragraph indentation', (html.match(/leave-html-flow-row leave-html-indent leave-html-identity-row/g) || []).length >= 2],
  ['clear notification states', html.includes('function inferToastType_(') && html.includes(".toast.error") && html.includes("role=\"status\" aria-live=\"polite\"")],
  ['desktop A4 layout retained', html.includes('.leave-html-date-range-row { flex-wrap: nowrap') && !html.includes('leave-html-duration') && html.includes('font-size: 12.5pt')],
  ['aligned form controls', html.includes('leave-html-check-option') && html.includes('leave-html-stat-head-main') && html.includes('grid-template-columns: 58px minmax(0, 1fr)') && html.includes('grid-template-columns: 14mm 55mm 17mm minmax(0, 1fr)')],
  ['small-screen layout isolated', html.includes('.compact-layout .leave-html-identity-row {') && html.includes('.compact-layout .leave-html-date-range-row {') && html.includes('leave-html-form compact-layout')],
  ['small-screen symmetric statistics columns', html.includes('.compact-layout.leave-html-form-sick .leave-html-stat-table th,') && html.includes('width: 25%;') && html.includes('width: 33.3333%;')],
  ['small-screen long text is fitted without truncation', html.includes('function fitCompactFormText_(') && html.includes('el.scrollWidth > el.clientWidth + 1') && html.includes('fitCompactFormText_(stage.firstElementChild)')],
  ['embedded TH Sarabun fallback for unsupported devices', html.includes("font-family: 'EmbeddedTHSarabunNew'") && html.includes("data:font/truetype;base64,") && html.includes("document.fonts.load('400 12.5pt EmbeddedTHSarabunNew')") && !html.includes('__THSARABUN_NEW_BASE64__')],
  ['desktop PDF downloads without Windows share dialog', html.includes('const isMobileDevice = /Android|iPhone|iPad|iPod/i.test') && html.includes('if (isMobileDevice && canShareFile)')],
  ['mobile saved PDF matches compact preview', html.includes('const useCompactExport = isCompactPdfViewport_();') && html.includes("(useCompactExport ? ' compact-layout' : '')") && html.includes('if (useCompactExport) fitCompactFormText_(stage.firstElementChild)')],
  ['PDF action and filename describe saved document', html.includes('hidden>บันทึก PDF</button>') && html.includes('function formatPdfFileDateRange_(') && html.includes("const safeLeaveType = String(record.leaveType") && html.includes("'ใบลา_' + safeLeaveType + '_' + safeName + '_' + shortDateRange")],
  ['last leave uses completed prior record', code.includes('if (en >= beforeStartDate) continue;') && code.includes('en > best.end') && code.includes('data[i][1] !== staffName || data[i][2] !== leaveType') && code.includes('data[i][0] === excludeRecordId')],
  ['leave reason vertically aligned', html.includes('.leave-html-reason-row {') && html.includes('align-items: center; min-height: 22pt') && !html.includes('class="leave-html-flow-row" style="margin:0">เนื่องจาก')],
  ['controlled dotted-line widths', html.includes('flex: 0 1 78mm; width: 78mm;') && html.includes('.leave-html-comment-line .leave-html-line.line-full { width: 70mm;') && html.includes('.leave-html-signature-row .leave-html-line { width: 48mm;')],
  ['supervisor blank lines keep writing width', html.includes('.leave-html-supervisor-field-row .leave-html-line.line-write { width: 48mm;') && !html.includes('.leave-html-line.line-content,\n    .leave-html-supervisor-field-row .leave-html-line.line-write { width: auto;')],
  ['automatic raster PDF disabled', !html.includes('if (res.recordId) createAndStorePdfForRecord_')]
  ,['double submit protection', html.includes('if (confirmBtn.disabled) return;') && html.includes('id="btnSubmitRecord"')]
  ,['inline field validation', html.includes('function markFieldError_(') && html.includes('field-error-text') && html.includes('aria-invalid')]
  ,['descriptive delete confirmation', html.includes('pendingDeleteDescription') && html.includes('pinModalDescription') && html.includes('ต้องการลบรายการ')]
  ,['Apps Script write lock', code.includes('function acquireWriteLock_()') && (code.match(/acquireWriteLock_\(\)/g) || []).length >= 8]
  ,['Bangkok timezone fixed', code.includes("const SYSTEM_TIMEZONE = 'Asia/Bangkok'") && !code.includes('Session.getScriptTimeZone()')]
  ,['calendar loading lock', html.includes('calendarLoading') && html.includes('id="calPrevBtn"') && html.includes('id="calNextBtn"')]
  ,['accessible focus and button size', html.includes('button:focus-visible') && html.includes('min-height:42px')]
  ,['Safari date controls stay inside their columns', html.includes('-webkit-appearance: none') && html.includes('input[type="date"]::-webkit-date-and-time-value') && html.includes('.profile-history-row .form-field { min-width: 0; width: 100%; max-width: 100%; overflow: hidden; }')]
  ,['personal mode banner', html.includes('id="myProfileActiveName"') && html.includes('personal-session-bar')]
  ,['record staff requires explicit selection', html.includes('<option value="" disabled selected>เลือกชื่อเจ้าหน้าที่</option>') && html.includes("sel.value = '';")]
  ,['automatic fiscal-year setup notice', code.includes('function getFiscalYearSetupStatus(') && html.includes('id="fiscalSetupNotice"') && html.includes('refreshFiscalYearSetupStatus_')]
  ,['vacation blocked until annual carry setup', html.includes("rec.leaveType === 'ลาพักผ่อน' && missingCarry") && code.includes('return found ? Number(found.vacationCarryDays) || 0 : 0;')]
  ,['carry setup follows personnel rights', code.includes('function getVacationAccumulationEligibility_(') && code.includes('noAccumulationCount') && html.includes('vacationAccumulationEligible')]
  ,['previous fiscal carry confirmation', html.includes('id="usePreviousCarry"') && html.includes('function applyPreviousVacationCarry_()') && code.includes('previousVacationCarryAvailable')]
  ,['carry maximum enforced', code.includes('safeCarryDays') && html.includes('vacationAccumMax')]
  ,['clear refresh action after changes', html.includes('id="dataRefreshBar"') && html.includes('function refreshAllData(') && (html.match(/showDataRefreshBar_\(/g) || []).length >= 4]
  ,['admin idle timeout', html.includes('ADMIN_IDLE_TIMEOUT_MS = 10 * 60 * 1000') && html.includes('function startAdminSessionTimer_()') && html.includes('lockAdminTab(true)')]
  ,['admin credentials cleared on lock', html.includes('adminTabUnlocked = false;') && html.includes('lastAdminPin = null;') && html.includes('clearTimeout(adminSessionTimer)')]
  ,['leave share available after save and in history', html.includes('id="shareLeaveModal"') && html.includes('class="detail-del line-share-btn"') && html.includes("prepareLeaveShare(\\'' + res.recordId") && (html.match(/prepareLeaveShare\(/g) || []).length >= 3]
  ,['leave share uses standard LINE and copy actions', html.includes('function openLineShare()') && html.includes('https://line.me/R/share?text=') && html.includes('function copyLeaveShareMessage()') && !html.includes('id="btnNativeShare"') && !html.includes('function shareLeaveWithDevice()')]
  ,['backdated share label uses Bangkok date', html.includes("timeZone: 'Asia/Bangkok'") && html.includes("' (บันทึกย้อนหลัง)'")]
  ,['LINE quota backend uses official endpoints', code.includes('function getLineQuotaStatus(') && code.includes('/v2/bot/message/quota') && code.includes('/v2/bot/message/quota/consumption')]
  ,['LINE quota admin status and thresholds', html.includes('id="lineQuotaCard"') && html.includes('function loadLineQuotaStatus(') && code.includes("percent >= 95") && code.includes("percent >= 80")]
  ,['smart duplicate leave response', code.includes('function leaveConflictResult_(') && code.includes("code: exact ? 'DUPLICATE_LEAVE' : 'OVERLAPPING_LEAVE'") && code.includes("return leaveConflictResult_(existing[i], startDate, endDate, 'duplicate')")]
  ,['smart duplicate leave actions', html.includes('id="duplicateLeaveModal"') && html.includes('เปิดดูรายการเดิม') && html.includes('แก้ไขรายการเดิม') && html.includes('function showDuplicateLeaveModal(') && html.includes('openPrintableLeaveForm(recordId)') && html.includes('openEditRecordModal(item.recordId')]
  ,['fresh isolated workspace initializer', code.includes('function initializeFreshBearLeaveSystem(confirmationToken)') && code.includes("confirmationToken !== 'CREATE_EMPTY_BEAR_SYSTEM'") && code.includes("['Staff', ['ID', 'Name', 'Active'") && code.includes("['AdminPin', '']")]
  ,['blank PIN never grants privileged access', code.includes('if (!inputPin) return false;') && code.includes('return !!adminPin && adminPin === inputPin;')]
  ,['legacy organization and personal data removed', !/(โรงพยาบาลขุนตาล|กายภาพบำบัด|เวชกรรมฟื้นฟู|นายชาญวิทย์|สมณะ|กนกพรรณ|ช่างจัด|วรญา|1S9tfaQ)/.test(code + html)]
  ,['new approval signature point is present', html.includes('function renderGroupHeadCommentBlock()') && html.includes('ความเห็นของผู้บังคับบัญชา (หัวหน้ากลุ่มงาน)')]
  ,['HR verifier identity excluded', !/(วรญา|มงคลจิรพร)/.test(html)]
];

for (const [name, ok] of checks) {
  if (!ok) throw new Error('QC failed: ' + name);
}

console.log('printable leave form QC OK');
