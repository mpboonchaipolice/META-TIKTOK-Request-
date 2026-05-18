# ระบบออกหนังสือขอข้อมูล — TCSD2

ระบบสร้างหนังสือขอข้อมูล (Facebook / Instagram / TikTok) อัตโนมัติ จาก Template Word + รายการ URL ใน Excel
ทำงาน **100% ในเบราว์เซอร์** ไม่มีการส่งข้อมูลออกไปยัง server ใดๆ — เหมาะกับเอกสารราชการที่ต้องการความปลอดภัยของข้อมูล

## คุณสมบัติ

- ✅ Format หนังสือ **ไม่เพี้ยน** เพราะแก้แค่ตัวแปรใน .docx โดยตรง (ไม่ convert ผ่าน Google Docs)
- ✅ รัน **เลขหนังสืออัตโนมัติ** ทีละ 1 ต่อหนังสือ
- ✅ **Facebook/Instagram:** 4 URL ต่อหนังสือ · **TikTok:** 1 URL ต่อหนังสือ
- ✅ อัพโหลด Excel **ลากวาง** หรือคลิกเลือก
- ✅ ดาวน์โหลด **.docx รายฉบับ** หรือ **รวม ZIP**
- ✅ เก็บ Template ใน **IndexedDB** ของเบราว์เซอร์ — อัพโหลดครั้งเดียว ใช้ได้ตลอด
- ✅ ฟรี 100% — host บน GitHub Pages

## วิธี Deploy บน GitHub Pages

### 1. สร้าง Repository
1. ล็อกอิน GitHub → คลิก **New repository**
2. ตั้งชื่อ เช่น `letter-generator` (จะเป็น public หรือ private ก็ได้ — Pages รองรับทั้งคู่ใน plan ฟรี)
3. ไม่ต้อง init README (เพราะจะ push ทับ)

### 2. อัพโหลดไฟล์
**วิธีที่ 1 — ผ่านหน้าเว็บ (ง่ายสุด):**
1. ในหน้า repo เปล่า กด **uploading an existing file**
2. ลากไฟล์ 4 ไฟล์เข้าไป: `index.html`, `style.css`, `app.js`, `README.md`
3. Commit changes

**วิธีที่ 2 — ผ่าน git CLI:**
```bash
git clone https://github.com/USERNAME/letter-generator.git
cd letter-generator
# คัดลอก index.html, style.css, app.js, README.md เข้ามา
git add .
git commit -m "Initial commit"
git push origin main
```

### 3. เปิด GitHub Pages
1. ในหน้า repo → **Settings** → **Pages** (เมนูซ้าย)
2. ที่หัวข้อ **Build and deployment** → **Source** เลือก **Deploy from a branch**
3. **Branch:** เลือก `main` / `(root)` → **Save**
4. รอ 1-2 นาที (refresh หน้านี้จะมีกล่องสีเขียวบอก URL)
5. URL จะเป็น `https://USERNAME.github.io/letter-generator/`

### 4. ใช้งาน
1. เปิด URL ที่ได้
2. คลิก **จัดการ Template** → อัพโหลดไฟล์ template ทั้ง 3 (อัพโหลดครั้งเดียวพอ — ระบบจำไว้)
3. กรอกข้อมูล + อัพโหลด Excel
4. กด **ประมวลผล**

> ⚠️ **สำคัญ:** Template ที่อัพโหลดต้องเป็นไฟล์ที่ผมแก้ให้แล้ว (เวอร์ชัน "fixed templates" จากที่ส่งให้รอบก่อน) — เพราะ template ของเดิมมี Word "แบ่ง run" ทำให้ระบบ replace ตัวแปรไม่เจอ

## วิธีทำ PDF

หลังดาวน์โหลด .docx แล้ว:

| โปรแกรม | ขั้นตอน |
|---|---|
| **Microsoft Word** | File → Save As → เลือก *PDF* |
| **LibreOffice Writer** | File → Export as → Export as PDF |
| **Google Docs** | อัพโหลด → เปิด → File → Download → PDF Document (.pdf) *(แต่ format อาจเพี้ยน)* |
| **macOS Preview / Pages** | เปิดด้วย Pages → File → Export To → PDF |

**Tip สำหรับงานราชการ:** ใช้ Microsoft Word จะคงรูปแบบเป๊ะที่สุด เพราะเป็นโปรแกรมที่ใช้สร้าง template

## โครงสร้างโปรเจค

```
letter-generator/
├── index.html      # UI หลัก
├── style.css       # สไตล์
├── app.js          # Logic ทั้งหมด
└── README.md       # คู่มือ
```

## เทคโนโลยีที่ใช้

| Library | ใช้ทำอะไร |
|---|---|
| [PizZip](https://github.com/open-xmlformats/pizzip) | อ่าน/เขียน .docx (ซึ่งเป็น ZIP) |
| [SheetJS (xlsx)](https://sheetjs.com/) | อ่าน Excel |
| [JSZip](https://stuk.github.io/jszip/) | รวมไฟล์เป็น ZIP |
| [FileSaver.js](https://github.com/eligrey/FileSaver.js) | บันทึกไฟล์ลงเครื่อง |
| IndexedDB | เก็บ template ในเบราว์เซอร์ |

ทุก library โหลดจาก CDN ไม่ต้อง npm/build อะไร

## ตัวแปรใน Template

Template Word ต้องมีตัวแปรในรูปแบบ `<VarName>` ดังนี้:

| ตัวแปร | ความหมาย | ใช้ใน |
|---|---|---|
| `<LetterDate>` | วันที่ออกหนังสือ | ทุก template |
| `<LetterNo>` | เลขที่หนังสือ (รันอัตโนมัติ) | ทุก template |
| `<StartTime>` | วันเริ่มต้นของข้อมูลที่ขอ | ทุก template |
| `<URL1>` | URL บัญชีที่ 1 | Facebook/Instagram/TikTok |
| `<URL2>` | URL บัญชีที่ 2 | Facebook/Instagram |
| `<URL3>` | URL บัญชีที่ 3 | Facebook/Instagram |
| `<URL4>` | URL บัญชีที่ 4 | Facebook/Instagram |

> ⚠️ ตัวแปรต้องอยู่ใน Word "run" เดียว ไม่ถูกแบ่ง (rsid issue) — ถ้าแก้ template เองให้ลบทั้งบรรทัด แล้วพิมพ์ใหม่จากซ้ายไปขวารวดเดียวไม่หยุด

## ข้อมูลความเป็นส่วนตัว

ระบบนี้ทำงาน **client-side ทั้งหมด**:
- Template เก็บใน **IndexedDB** ของเบราว์เซอร์ผู้ใช้ (ไม่ได้อัพโหลดไป server)
- Excel ที่อัพโหลด → อ่านในเบราว์เซอร์ → ไม่ส่งออก
- การสร้าง .docx → ทำในเบราว์เซอร์ → ไม่ส่งออก
- ดาวน์โหลด → ตรงจาก memory ของเบราว์เซอร์

GitHub Pages ทำหน้าที่แค่ส่ง HTML/CSS/JS ให้ผู้ใช้ — ไม่เห็นข้อมูลใดๆ ของผู้ใช้เลย

## License

ภายในของ TCSD2 — ห้ามเผยแพร่ template ไฟล์
