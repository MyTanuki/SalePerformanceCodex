# GitHub Commit / Push Checklist

Repository: https://github.com/MyTanuki/SalePerformanceCodex.git

## Purpose

ใช้ไฟล์นี้เป็นขั้นตอนยืนยันก่อน Commit และ Push งานขึ้น GitHub เพื่อให้ตรวจสอบได้ว่าไฟล์ที่จะส่งขึ้น repository ถูกต้องครบถ้วน และไม่รวมไฟล์ข้อมูลหรือไฟล์ชั่วคราวที่ไม่เกี่ยวข้อง

## Current Workflow

1. ตรวจสอบสถานะไฟล์

```powershell
git status --short
```

2. ตรวจสอบไฟล์ที่เปลี่ยนแปลง

```powershell
git diff -- index.html app.js styles.css README.md
```

3. เลือก stage เฉพาะไฟล์ที่ต้องการส่งขึ้น GitHub

```powershell
git add index.html app.js styles.css README.md GITHUB_COMMIT_PUSH.md
```

ถ้าต้องการรวมไฟล์อื่น ให้ตรวจสอบ diff ก่อน แล้วจึงเพิ่มชื่อไฟล์แยกทีละไฟล์

4. ตรวจสอบรายการที่จะ commit

```powershell
git status --short
git diff --cached --stat
```

5. Commit

```powershell
git commit -m "Update sales performance dashboard"
```

6. Push ไปยัง GitHub

```powershell
git push origin main
```

## Files To Review Before Push

- `index.html` โครงสร้างหน้า Web App และ modal
- `app.js` logic dashboard, filter, revenue analysis, information check
- `styles.css` layout, theme, animation, table style
- `README.md` เอกสารประกอบ Web App
- `GITHUB_COMMIT_PUSH.md` ขั้นตอน commit และ push

## Files To Avoid Unless Intentionally Needed

- ไฟล์ `.csv` ข้อมูลจริงหรือข้อมูลลูกค้า
- `desktop.ini`
- ไฟล์ backup หรือไฟล์ชั่วคราว
- ไฟล์ที่ถูกลบโดยไม่ได้ตั้งใจ

## Confirmation Before Push

ก่อน push ควรยืนยัน 3 เรื่องนี้:

- ตรวจสอบแล้วว่าไม่มีข้อมูลลูกค้าในไฟล์ที่จะ commit โดยไม่ตั้งใจ
- ตรวจสอบแล้วว่าไฟล์ที่ถูกลบเป็นการลบที่ตั้งใจ
- ทดสอบ syntax หรือเปิด preview แล้วเท่าที่เครื่องทำได้

