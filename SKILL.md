# Si-Som EXCEL Skill (si-som-excel) v2.0

## Overview
Expert Excel Automation for "Don" (Krudony). High-precision, high-safety engine for complex educational documents.

## Core Protocols (กฎเหล็ก v2.0)
1. **DYNAMIC DISCOVERY (ห้ามเดา):**
   - ห้ามระบุแถว/คอลัมน์เริ่มต้นแบบตายตัว (Fixed Rows) เด็ดขาด!
   - ต้องสั่งแสกนหา "คำว่า เลขที่" หรือ "ชื่อ-สกุล" เพื่อหาจุดสตาร์ทของทุก Sheet ใหม่เสมอ!
2. **XLWINGS SUPREMACY (เครื่องยนต์หลัก):**
   - บน Windows ต้องใช้ scripts/write_excel_pro.py (xlwings) เพื่อป้องกันไฟล์พัง 100%
   - ใช้ Visible=True เพื่อรับมือกับ Protected View และการแจ้งเตือนของ Excel
3. **VERIFY TWICE, WRITE ONCE:**
   - รายงานพิกัด (เช่น "แถวเริ่ม 8, สิ้นสุด 19") ให้ดอนทราบก่อนลงมือ Bulk Write ทุกครั้ง
4. **THAI STANDARD:**
   - บังคับใช้ **TH SarabunPSK 16pt** และจัดกึ่งกลาง (Center) ทุุกเซลล์ที่แก้ไข
