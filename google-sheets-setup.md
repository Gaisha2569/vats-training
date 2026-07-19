# วิธีตั้งค่าเก็บผลผู้เรียนส่วนกลาง (Google Sheets)

ทำครั้งเดียว ใช้เวลาประมาณ 10 นาที ไม่มีค่าใช้จ่าย
เมื่อทำเสร็จ ผลของผู้เรียนทุกคนจะไหลมารวมที่ Google Sheet ของท่าน
และแดชบอร์ด Admin ในเว็บจะดึงข้อมูลจริงมาแสดงแทนข้อมูลตัวอย่าง

---

## ขั้นที่ 1 — สร้าง Google Sheet

1. เปิด [sheets.new](https://sheets.new) สร้างสเปรดชีตใหม่
2. ตั้งชื่อไฟล์ เช่น `ผลการอบรม VATS`

## ขั้นที่ 2 — ใส่สคริปต์

1. ในสเปรดชีต เมนู **ส่วนขยาย (Extensions)** → **Apps Script**
2. ลบโค้ดเดิมที่มีอยู่ทั้งหมด แล้ววางโค้ดข้างล่างนี้แทน
3. กดไอคอน **บันทึก** (💾)

```javascript
const SHEET_NAME = 'results';
const HEADERS = ['uid', 'ชื่อ-สกุล', 'หน่วยงาน', 'ความก้าวหน้า(%)',
                 'ก่อนเรียน', 'หลังเรียน', 'อัปเดตล่าสุด'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet();
    const rec = [d.uid, d.name, d.ward, d.progress, d.pre, d.post,
                 Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm')];
    const rows = sh.getDataRange().getValues();
    let found = 0;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === d.uid) { found = i + 1; break; }
    }
    if (found) sh.getRange(found, 1, 1, rec.length).setValues([rec]);
    else sh.appendRow(rec);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const rows = getSheet().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({
      uid: rows[i][0], name: rows[i][1], ward: rows[i][2],
      progress: rows[i][3], pre: rows[i][4], post: rows[i][5]
    });
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## ขั้นที่ 3 — เผยแพร่เป็น Web App

1. มุมขวาบน กด **ทำให้ใช้งานได้ (Deploy)** → **การทำให้ใช้งานได้ใหม่ (New deployment)**
2. กดไอคอนเฟือง ⚙️ ข้าง "เลือกประเภท" → เลือก **แอปเว็บ (Web app)**
3. ตั้งค่าดังนี้

   | ช่อง | เลือก |
   |---|---|
   | ดำเนินการในชื่อ (Execute as) | **ฉัน (Me)** |
   | ผู้ที่มีสิทธิ์เข้าถึง (Who has access) | **ทุกคน (Anyone)** |

4. กด **ทำให้ใช้งานได้** → ครั้งแรกจะให้กด **ให้สิทธิ์ (Authorize)**
   ถ้าขึ้นเตือน "Google ยังไม่ได้ยืนยันแอปนี้" ให้กด **ขั้นสูง** → **ไปที่ ... (ไม่ปลอดภัย)**
   เป็นเรื่องปกติเพราะเป็นสคริปต์ที่ท่านเขียนเอง
5. คัดลอก **URL ของแอปเว็บ** ที่ได้ (ลงท้ายด้วย `/exec`)

## ขั้นที่ 4 — ใส่ URL ลงในเว็บ

เปิดไฟล์ `config.js` แล้วแก้บรรทัด `API` เป็น URL ที่คัดลอกมา

```javascript
API: 'https://script.google.com/macros/s/AKfycb...../exec'
```

พร้อมกันนี้เปลี่ยนรหัสผ่าน Admin ด้วย

```javascript
ADMIN_PASS: 'รหัสผ่านที่ท่านตั้งเอง',
```

แล้ว push ขึ้น GitHub

```powershell
cd "C:\Users\Dragon\Desktop\Clip"
git add config.js
git commit -m "ตั้งค่าเก็บผลส่วนกลางและรหัสผ่าน Admin"
git push
```

---

## ทดสอบว่าใช้ได้

1. เปิดเว็บ → เลือก **ผู้เรียน** → กรอกชื่อทดสอบ
2. ดูวิดีโอสัก 1 หัวข้อ หรือทำแบบทดสอบ
3. กลับไปดู Google Sheet ควรมีแถวใหม่ขึ้นมาภายในไม่กี่วินาที
4. กลับหน้าแรก → เลือก **Admin** → ใส่รหัสผ่าน → แดชบอร์ดควรแสดงชื่อที่เพิ่งทดสอบ

---

## ข้อจำกัดที่ควรทราบ

เว็บนี้เป็น static site บน GitHub Pages ไม่มีเซิร์ฟเวอร์ของตัวเอง จึงมีข้อจำกัดดังนี้

- **รหัสผ่าน Admin อยู่ในไฟล์ `config.js` ซึ่งเปิดอ่านได้**
  ป้องกันคนทั่วไปที่กดเล่นได้ แต่กันคนที่รู้วิธีดูโค้ดไม่ได้
  อย่าใช้รหัสผ่านเดียวกับระบบอื่นของโรงพยาบาล

- **URL ของ Apps Script เปิดอ่านได้เช่นกัน**
  ผู้ที่ทราบ URL สามารถดูรายชื่อผู้เรียนหรือส่งข้อมูลปลอมเข้ามาได้
  เหมาะกับการอบรมภายในที่ความเสี่ยงต่ำ ไม่ควรเก็บข้อมูลที่เป็นความลับ

- **ไม่ควรกรอกข้อมูลที่ระบุตัวผู้ป่วย** ในระบบนี้

หากต้องการความปลอดภัยระดับใช้งานจริงทั้งโรงพยาบาล
ควรย้ายไปใช้ระบบที่มีการล็อกอินจริง เช่น Google Workspace ของหน่วยงาน
หรือ LMS ที่โรงพยาบาลใช้อยู่
