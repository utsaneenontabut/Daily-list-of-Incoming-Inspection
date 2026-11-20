# Daily list of Incoming Inspection

เว็บแอปแจกเป้าตรวจรับ (Incoming Inspection) — ไฟล์เดียว `index.html` ทำงานในเบราว์เซอร์ ไม่มีเซิร์ฟเวอร์

## ไฟล์ในโปรเจกต์
- `index.html` — ตัวแอป (ไม่มีข้อมูลบริษัทฝังอยู่ ข้อมูลทั้งหมดอัปโหลดผ่านหน้าเว็บ)
- `Code.gs` — โค้ด Google Apps Script สำหรับใช้ Google Sheet เป็นตัวกลางข้อมูลร่วมทีม (ไม่บังคับ)
- `docs/SETUP_GOOGLE_SHEET.md` — คู่มือตั้งค่า Google Sheet

## ใช้งานบน GitHub Pages
Settings > Pages > Branch `main` / root > Save แล้วเปิด `https://<user>.github.io/<repo>/`

## ความปลอดภัย
- **อย่าอัปโหลดไฟล์ข้อมูลจริง** (Export SAP, Phase, รายการตรวจสอบประจำวัน, ไฟล์ส่งต่องาน) ขึ้น repository นี้
- GitHub Pages ฟรีใช้ได้กับ repository สาธารณะเท่านั้น ใครก็เปิดหน้าเว็บได้ แต่ไม่เห็นข้อมูลบริษัท เพราะข้อมูลอยู่ในเบราว์เซอร์ของผู้ใช้หรือ Google Sheet ของบริษัท
- URL และ TOKEN ของ Apps Script ห้ามใส่ลงในไฟล์ใน repository ให้ผู้ใช้กรอกในแอปเอง
