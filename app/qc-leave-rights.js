const fs = require('fs');
const vm = require('vm');
const path = require('path');

const source = fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8');
const match = source.match(/const LEAVE_QUOTA_SEED_ROWS = (\[[\s\S]*?\n\]);/);
if (!match) throw new Error('ไม่พบ LEAVE_QUOTA_SEED_ROWS');

const rows = vm.runInNewContext(match[1]);
const expected = [
  ['ข้าราชการพลเรือน', 10, 45, 60, 90, 15, 20],
  ['ลูกจ้างประจำ', 10, 45, 60, 90, 15, 20],
  ['พนักงานราชการ', 10, 10, 30, 90, 0, 5],
  ['พนักงานกระทรวงฯ (พกส.)', 10, 15, 45, 90, 15, 5],
  ['ลูกจ้างชั่วคราว (รายเดือน)', 10, 0, 15, 90, 0, 0],
  ['ลูกจ้างชั่วคราว (รายวัน)', 10, 0, 15, 0, 0, 0]
];
const leaveTypes = [
  'ลาพักผ่อน',
  'ลากิจส่วนตัว',
  'ลาป่วย',
  'ลาคลอดบุตร',
  'ลาไปช่วยเหลือภริยาที่คลอดบุตร'
];

for (const [personnelType, vacation, personal, sick, maternity, spouse, accumMax] of expected) {
  const quotas = [vacation, personal, sick, maternity, spouse];
  leaveTypes.forEach((leaveType, index) => {
    const row = rows.find(item => item[0] === personnelType && item[1] === leaveType);
    if (!row) throw new Error(`ขาดสิทธิ์ ${personnelType}: ${leaveType}`);
    if (Number(row[2]) !== quotas[index]) {
      throw new Error(`${personnelType}: ${leaveType} ต้องเป็น ${quotas[index]} แต่พบ ${row[2]}`);
    }
  });
  const vacationRow = rows.find(item => item[0] === personnelType && item[1] === 'ลาพักผ่อน');
  if ((Number(vacationRow[4]) || 0) !== accumMax) {
    throw new Error(`${personnelType}: วันสะสมสูงสุดต้องเป็น ${accumMax} แต่พบ ${vacationRow[4]}`);
  }
}

if (!source.includes('personnelType: row.personnelType, position: row.position, gender: row.gender')) {
  throw new Error('getInitialData ไม่ได้ส่งเพศไปคำนวณสรุป');
}
if (!source.includes("if (gender === 'ชาย') return 'ลาไปช่วยเหลือภริยาที่คลอดบุตร'")) {
  throw new Error('ไม่พบการจับคู่สิทธิ์ช่วยภริยาสำหรับเจ้าหน้าที่ชาย');
}

console.log('leave rights reference QC OK');
