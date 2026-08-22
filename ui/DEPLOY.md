# วิธี Deploy หน้า Video Creator (Web App)

## ขั้นตอน

1. เปิดโปรเจกต์ Google Apps Script ของระบบ (ที่มี `main.gs`, `config.gs`, `webapp.gs` ฯลฯ)
2. สร้างไฟล์ HTML ใหม่ ชื่อ **`VideoCreator`** (File → New → HTML)
3. คัดลอกเนื้อหาทั้งหมดจาก `ui/video-creator.html` ไปวาง แล้วบันทึก
4. คัดลอกโค้ดจาก `webapp.gs` ไปวางในไฟล์ `webapp.gs` ของโปรเจกต์
5. Deploy → **New deployment** → เลือกประเภท **Web app**
   - Execute as: **Me**
   - Who has access: **Only myself** (หรือ Anyone with Google account ถ้าจะให้ทีมใช้)
6. เปิด URL ที่ได้ — ครบทุกฟีเจอร์: ดึงสินค้าจาก Sheet, AI สร้างเรื่อง+บท 3 ซีน, บันทึกงาน

## หมายเหตุ

- เปิดไฟล์ `video-creator.html` ตรง ๆ ในเครื่องได้เหมือนเดิม แต่จะทำงานใน **โหมดตัวอย่าง (Demo)** ไม่ยิง API จริง
- ปุ่ม "เริ่มสร้าง" จะเขียนแถวใหม่ลงชีต `Control` โดยใส่สถานะ:
  - ปิด "ส่งเข้า LINE ให้แอดมินอนุมัติ" → สถานะ `PENDING` (เข้าคววาม `processAllPending` เดิมทันที)
  - เปิดอนุมัติ → สถานะ `PENDING_REVIEW` + ส่ง LINE แจ้งเตือน
- การสร้าง "วิดีโอจริงทุกซีน" ต้องต่อกับบริการ text-to-video ภายนอก (เช่น AZ Creator / API อื่น) — ปัจจุบันระบบจะเตรียมบท + แคปชัน + คิวโพสต์ให้ครบ ส่วนไฟล์วิดีโอให้ใส่ URL ในคอลัมน์ Media URL ของแถวนั้น
