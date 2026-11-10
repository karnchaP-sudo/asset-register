# ทะเบียนสินทรัพย์ (Fixed Asset Register)

เว็บแอปทะเบียนสินทรัพย์ถาวร: รหัสทรัพย์สิน (เช่น `OF2601/0001`), รูปถ่าย, ค่าเสื่อมราคาแบบเส้นตรง และพิมพ์ป้าย QR ตรงไปเครื่องพิมพ์ NIIMBOT ผ่าน Bluetooth

## ข้อมูล
- หน้าเว็บนี้ **ไม่มีข้อมูลบริษัท** ผู้ใช้แต่ละเครื่องกด "นำเข้าไฟล์ข้อมูล" (.json) เอง
- ข้อมูลและรูปเก็บใน `localStorage` ของเบราว์เซอร์เครื่องนั้นเท่านั้น ใช้ปุ่ม "ดาวน์โหลดไฟล์สำรอง" เพื่อสำรองหรือย้ายเครื่อง

## พิมพ์ป้าย NIIMBOT
- ใช้ได้กับ Chrome หรือ Edge บนคอมพิวเตอร์หรือ Android (Web Bluetooth) iPhone/iPad ใช้ไม่ได้
- ทดสอบออกแบบสำหรับรุ่น B4 ม้วน 50 × 30 มม. (ตั้งค่าแบบย่อ 50 × 15 มม. ได้)

## ไลบรารีที่ใช้
- [niimbluelib](https://github.com/MultiMote/niimbluelib) 0.47.1 (MIT) – `vendor/niimbluelib-0.47.1.min.js`
- [qrcode-generator](https://github.com/kazuhikoarase/qrcode-generator) 1.4.4 (MIT) – `vendor/qrcode-generator-1.4.4.js`
