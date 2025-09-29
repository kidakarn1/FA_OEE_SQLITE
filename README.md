# 🏭 FA System (Factory Automation)

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)]()
[![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)]()
[![License](https://img.shields.io/badge/license-proprietary-red.svg)]()

> **FA System (Factory Automation)** is a production monitoring platform designed for real-time efficiency tracking, loss & defect management, and automated FA Tag generation.

---

## 📑 Table of Contents
1. [Overview (English)](#overview-english)
2. [Features – English](#features--english)
3. [Overview (ภาษาไทย)](#ภาพรวม--ภาษาไทย)
4. [คุณสมบัติ – ภาษาไทย](#คุณสมบัติ--ภาษาไทย)
5. [Summary](#-summary)

---

## 🌍 Overview (English)

**FA System** is developed to support real-time production monitoring, improve efficiency, and ensure accurate data collection on the shop floor.  

It is designed for manufacturing businesses that need to:  
- 📌 Track production in real time  
- 📌 Monitor OEE (**Availability, Performance, Quality**)  
- 📌 Manage production losses and defects effectively  
- 📌 Configure production line parameters and equipment (Printer, Scanner, Counter, etc.)  
- 📌 Generate and reprint FA Tags automatically when production is completed  

With **FA System**, both shop floor workers and supervisors can work more efficiently, while management gains accurate insights into production performance.  

---

## 🚀 Features – English

### 1️⃣ Menu **Start Production**
_For shop floor workers_
- ✅ Employee card scan before production start  
- ✅ Auto-generate **Plan / Lot / WI** from Web Manage FA  
- ✅ Scan **TAG RM** to record Lot, PO, Supplier for traceability  
- ✅ **Production Dashboard (OEE):**
  - **Availability (A):** Time Loss, Top 3 Loss (excluding Loss X & P1)  
  - **Quality (Q):** NG quantity, Top 3 NG Codes, Defect logging (NC/NG, FG/Child Part, Supplier)  
  - **Performance (P):** STD ACT (Target), Actual (Real output), recalculated every 2 minutes  
  - **OEE Formula:** `OEE = A × P × Q`  
  - Hourly **Target vs Actual** table  
- ✅ **Additional Functions:** Record Defect/Loss, Reprint Tags, Change Worker, INC/DEC Production QTY, Close Production with summary  
- ✅ **Special Line Functions:** Multi-plan run, Cross-line plan, QR Product scan, Auto Start at shift time  

---

### 2️⃣ Menu **Admin**
_For Leaders / Supervisors / Unit Chiefs_
- ✅ Reprint Tags (Production, Defect, Label)  
- ✅ Switch production line  
- ✅ Access production history and Reprint logs  

---

### 3️⃣ Menu **System (Config)**
_For Admins / Engineers_
- ✅ Configure Line: **PD, Line, Printer, Scanner, Counter QTY, Power Lamp**  
- ✅ Set Production Parameters: **Delay, Cavity, OEE thresholds, Q-policy** (include/exclude Part No. from Q)  
- ✅ Enable / disable **Auto Start by Shift**  
- ✅ Device Testing: Printer, Scanner, Lamp, Counter  

---

### 4️⃣ Menu **Exit**
- ✅ Shutdown system  
- ✅ Restart computer  
- ✅ Safety check before exit (pending jobs, FA Tags not printed)  

---

## 🇹🇭 ภาพรวม (ภาษาไทย)

**FA System (Factory Automation)** ถูกพัฒนาขึ้นเพื่อรองรับการติดตามการผลิตแบบเรียลไทม์ เพิ่มประสิทธิภาพ และช่วยให้การเก็บข้อมูลหน้างานมีความถูกต้องแม่นยำ  

ออกแบบมาสำหรับธุรกิจการผลิตที่ต้องการ:  
- 📌 ติดตามการผลิตแบบเรียลไทม์  
- 📌 ตรวจสอบ OEE (**Availability, Performance, Quality**)  
- 📌 จัดการการสูญเสีย (Loss) และงานเสีย (Defect)  
- 📌 ตั้งค่าอุปกรณ์และพารามิเตอร์ของไลน์การผลิต (Printer, Scanner, Counter ฯลฯ)  
- 📌 สร้างและพิมพ์ซ้ำ FA Tag อัตโนมัติเมื่อการผลิตเสร็จสิ้น  

---

## 🚀 คุณสมบัติ – ภาษาไทย

### 1️⃣ เมนู **Start Production**
_สำหรับพนักงานหน้างาน (Worker)_
- ✅ สแกนบัตรพนักงานก่อนเริ่มงาน  
- ✅ ระบบ Auto-generate **Plan / Lot / WI** จาก Web Manage FA  
- ✅ สแกน **TAG RM** เพื่อบันทึก Lot, PO, Supplier สำหรับการ Traceability  
- ✅ **หน้าจอ Production Dashboard (OEE):**
  - **Availability (A):** เวลา Loss, Top 3 Loss (ไม่รวม Loss X & P1)  
  - **Quality (Q):** จำนวน NG, Top 3 NG Code, บันทึก Defect (NC/NG, FG/Child Part, Supplier)  
  - **Performance (P):** STD ACT (เป้าหมาย), Actual (เดินงานจริง), คำนวณใหม่ทุก 2 นาที  
  - **สูตร OEE:** `OEE = A × P × Q`  
  - ตาราง **เป้าหมาย vs ผลผลิตจริง** รายชั่วโมง  
- ✅ **ฟังก์ชันเพิ่มเติม:** บันทึก Defect/Loss, Reprint Tag, เปลี่ยนพนักงาน, เพิ่ม/ลดยอดการผลิต, ปิดงานพร้อมสรุปผล  
- ✅ **ฟังก์ชันพิเศษสำหรับไลน์:** เดินหลายแผนพร้อมกัน, เลือกแผนจากไลน์อื่น, Scan QR Product, Auto Start ตามเวลา Shift  

---

### 2️⃣ เมนู **Admin**
_สำหรับ Leader / Supervisor / Unit Chief_
- ✅ Reprint Tag (Production, Defect, Label)  
- ✅ เปลี่ยนไลน์การผลิต  
- ✅ ดูประวัติการผลิตและ Reprint Logs  

---

### 3️⃣ เมนู **System (Config)**
_สำหรับ Admin / Engineer_
- ✅ ตั้งค่าไลน์: **PD, Line, Printer, Scanner, Counter QTY, Power Lamp**  
- ✅ ตั้งค่าพารามิเตอร์: **Delay, Cavity, OEE Thresholds, นโยบาย Q** (เลือก Part ที่นับ/ไม่นับ Q)  
- ✅ เปิด/ปิด **Auto Start ตาม Shift**  
- ✅ ทดสอบอุปกรณ์: Printer, Scanner, Lamp, Counter  

---

### 4️⃣ เมนู **Exit**
- ✅ ปิดระบบ (Shutdown)  
- ✅ Restart คอมพิวเตอร์  
- ✅ ตรวจสอบก่อนปิดงาน (เช่น งานค้าง, FA Tag ยังพิมพ์ไม่ครบ)  

---

## 📊 Summary
**FA System** provides a professional solution for:  
- Real-time production monitoring  
- OEE tracking (**A / P / Q**)  
- Loss & Defect control  
- Production line configuration  
- Automatic FA Tag generation  



## ⚙️ Installation

### Prerequisites
- 🖥️ Windows 
- 🗄️ Database: SQL SERVER / MySql
- 🔧 Git

### Steps
1. Clone the repository
   ```bash
   git clone https://github.com/your-username/fa-system.git
   cd fa-system
