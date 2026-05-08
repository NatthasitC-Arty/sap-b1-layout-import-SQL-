# SAP B1 Crystal Layout Batch Import (HANA)

เครื่องมือสำหรับ **import ไฟล์ Crystal Reports (.rpt) จำนวนมาก** เข้า SAP Business One **บน SAP HANA** แบบ batch ผ่าน SQL Direct (INSERT/UPDATE ตรงเข้า table `RDOC`)

---

## โครงสร้าง

```
sap-b1-layout-import-SQL-/
│
├── 🔑 _settings.bat                ← shared HANA connection (gitignored)
├── 🔑 _settings.bat.example        ← template
│
├── 🔧 .bat (double-click ได้)
│   ├── TestConnect.bat             ← ทดสอบ HANA
│   ├── RunImport.bat               ← import จาก Excel
│   ├── RunRollback.bat             ← เลือก layout ลบทีละตัว
│   └── RunDeleteNonSystem.bat      ← ลบ non-system ทั้งหมด
│
├── ⚙️ PowerShell scripts
│   ├── DB-HANA.ps1                 ← HANA plugin (Sap.Data.Hana wrapper)
│   ├── Generate-MapExcel.ps1       ← สแกน folder → สร้าง RPT_Import_Map.xlsx
│   ├── Test-SQLConnect.ps1         ← logic test connection
│   ├── Backup-RDOC.ps1             ← backup table RDOC
│   ├── Import_SQL_Direct.ps1       ← ⭐ logic import
│   ├── Rollback-BySelection.ps1    ← logic rollback
│   └── Delete-NonSystemLayouts.ps1 ← logic delete non-system
│
├── 📊 Config/RPT_Import_Map.xlsx   ← mapping (auto-generated)
└── 📋 Import_SQL_Log.txt           ← log (gitignored)
```

---

## Requirements

| รายการ | Version |
|--------|---------|
| Windows | 10 / 11 / Server 2016+ |
| PowerShell | 5.1 (built-in) |
| **SAP HANA Client** | 2.0+ (`Sap.Data.Hana.dll` ต้องลง — ดาวน์โหลดจาก SAP Marketplace) |
| SAP Business One | v10.0 บน HANA |
| HANA User | สิทธิ์ INSERT/UPDATE/DELETE ที่ schema ของ company DB (table `RDOC`, `RITM`, `RDC1`, `RCON`, `DFLT_PRNTING`) |

> ✅ **ไม่ต้อง** ลง SAP DI API / Service Layer — script คุยกับ HANA ตรงๆ ผ่าน .NET data provider

---

## Setup (ครั้งแรก)

### 1. ลง SAP HANA Client

ดาวน์โหลดจาก SAP Marketplace → `SAP HANA CLIENT 2.0` (ฟรี ~250 MB) — ติดตั้งไปที่ default path

หลังลงเสร็จ ต้องมีไฟล์ `Sap.Data.Hana.v4.5.dll` ที่:
```
C:\Program Files\SAP\hdbclient\dotnetv45\Sap.Data.Hana.v4.5.dll
```

### 2. สร้าง `_settings.bat`

```cmd
copy _settings.bat.example _settings.bat
notepad _settings.bat
```

แก้ 5 บรรทัด:
```bat
set SERVER=10.10.10.109:30015                  ← HANA host:port
set COMPANYDB=SBO_ENCONFUND_TRAINING           ← HANA schema = SAP B1 company name
set DBUSER=SYSTEM
set DBPASSWORD=YourPasswordHere
set RPTROOT=C:\GitHub\Enconfund\FORM           ← root folder ที่มีไฟล์ .rpt
```

> 🔍 หา schema name: ใน HANA Studio รัน `SELECT SCHEMA_NAME FROM SYS.SCHEMAS WHERE SCHEMA_NAME LIKE 'SBO%'`

### 3. ทดสอบ connection

ดับเบิลคลิก `TestConnect.bat` — ต้องเห็น:
```
[1/4] Pinging 10.10.10.109 ...    Ping: OK
[2/4] Testing TCP 30015 ...        TCP 30015: OPEN
[3/4] Connecting to HANA ...       HANA Login: OK (server 2.00.071.x)
[4/4] Counting layouts in RDOC ... Total: 662  Crystal: 197
READY TO IMPORT
```

---

## Mapping File (Excel)

`Config/RPT_Import_Map.xlsx` sheet `RPT_MAP` — schema:

| Col | Field | ตัวอย่าง |
|-----|-------|---------|
| A | No | 1 |
| B | **DocCode** | `INV10004` (8 ตัว) |
| C | **TypeCode** | `INV1` (4 ตัว = DocCode 4 ตัวแรก) |
| D | **RPT_FileName** | `INV10004__AR INVOICE Enconfund (1).rpt` |
| E | **RPT_Folder** | `INV10004__AR INVOICE Enconfund (1)` |
| F | LayoutName | `AR INVOICE Enconfund (1)` (ไปอยู่ใน `RDOC.DocName`) |
| G | Note | (ใส่อะไรก็ได้) |

### Auto-generate Excel

ถ้า .rpt files มี folder structure `<8-char DocCode>__<DocName>` (รูปแบบที่ extract tool สร้าง) — generate Excel อัตโนมัติได้เลย:

```powershell
.\Scripts\Generate-MapExcel.ps1 -RptRoot C:\GitHub\Enconfund\FORM
```

จะสแกนทุก folder ภายใต้ `-RptRoot`, parse ชื่อ folder, สร้าง row data ใส่ `Config\RPT_Import_Map.xlsx`

---

## Workflow

```
1. TestConnect.bat                  → READY TO IMPORT
2. .\Scripts\Backup-RDOC.ps1 ...    → backup table
3. RunImport.bat (MODE=-DryRun)     → preview
4. RunImport.bat (MODE=)            → run จริง
5. Verify ใน SAP B1 Client
```

### RunImport.bat options

```bat
set AUTHOR=manager       ← ใส่ใน RDOC.Author เฉพาะ INSERT (UPDATE คงเดิม)
set MODE=                ← ว่าง = real run, -DryRun = preview only
set ONDUP=Update         ← Update / Skip
```

**ONDUP behavior** (dedup โดย DocCode PK):
- `Update` = ถ้า DocCode มีอยู่แล้ว → UPDATE Template/RptHash/DocName
- `Skip`   = ถ้ามีอยู่แล้ว → ข้าม
- ถ้า existing row คือ `Author='System'` → **บังคับ SKIP เสมอ** (กันไม่ให้ overwrite system layout)

---

## Critical schema facts

| Table | Role | Key |
|-------|------|-----|
| `RDOC` | Layout metadata + binary template | `DocCode` PK |
| `RITM` | Layout items / line items | `(DocCode, ItemNum)` |
| `RDC1` | Secondary metadata | — |
| `RCON` | Conditions | — |
| `DFLT_PRNTING` | Per-user default layout per ObjectType | — |

- **DocCode format**: `<TypeCode><4-digit seq>` เช่น `INV20003` (8 chars total)
- **TypeCode = 4 chars แรก** ของ DocCode → ผูกกับ form ใน SAP B1
- **Author tag for system layouts = `'System'`** (ไม่ใช่ `'-System-'`)
- **`RDOC.Template` เก็บ raw .rpt bytes** (OLE Compound, magic `D0 CF 11 E0 A1 B1 1A E1`)

---

## HANA-specific notes

### Connection string

```
Server=<host>:<port>;UserID=<user>;Password=<pwd>;CurrentSchema=<schema>;CommunicationTimeout=15000;
```

`CurrentSchema` คือ company DB name (เช่น `SBO_ENCONFUND_TRAINING`) — script ใช้ value นี้เป็น schema-qualified prefix สำหรับทุก SQL (`"SBO_ENCONFUND_TRAINING"."RDOC"`)

### Parameter binding

HANA's `HanaCommand` รองรับเฉพาะ **positional `?` parameters**. Script ใน repo นี้ **เขียน SQL ด้วย `@name` placeholders** (อ่านง่าย) แล้ว `DB-HANA.ps1` แปลเป็น `?` ก่อน execute (ดู `Set-DBCommandText` / `Submit-DBParams`)

### SQL dialect

| MSSQL | HANA |
|-------|------|
| `GETDATE()` | `CURRENT_TIMESTAMP` |
| `ISNULL(x, y)` | `IFNULL(x, y)` |
| `DATALENGTH(blob)` | `LENGTH(blob)` |
| `[dbo].[RDOC]` | `"SBO_..."."RDOC"` |
| `SqlDbType.Image` | `HanaDbType.Blob` |
| `@param` | `?` (positional) |

---

## Troubleshooting

### `SAP HANA .NET data provider not found`
- ลง SAP HANA Client จาก SAP Marketplace
- ตรวจ: `Test-Path "C:\Program Files\SAP\hdbclient\dotnetv45\Sap.Data.Hana.v4.5.dll"` ต้อง `True`

### `cannot find schema 'SBO_...'`
- Schema name ผิด — รัน `SELECT SCHEMA_NAME FROM SYS.SCHEMAS WHERE SCHEMA_NAME LIKE 'SBO%'` ใน HANA Studio
- Note: schema name มัก uppercase แต่ HANA case-sensitive ใน double quotes

### `authentication failed`
- DBUSER / DBPASSWORD ผิด
- Login ด้วย user เดียวกันใน HANA Studio ยืนยัน

### Import ผ่าน แต่ SAP B1 เปิด layout ไม่ได้
- **Crystal Runtime version ต่างกัน** — re-save .rpt ด้วย CR Designer version เก่ากว่า
- **Print Preview "ODBC -2028"** = layout default pointer ใน `DFLT_PRNTING` ชี้ไป DocCode ที่ลบแล้ว → cleanup orphan rows

---

## ข้อควรระวัง

⚠️ **SAP ไม่ Support การ INSERT ตรงเข้า RDOC** — ถ้าเกิดปัญหา SAP support อาจไม่ช่วย → **Backup ทุกครั้งก่อน import**

⚠️ **DocCode ที่ extract มาจาก source DB อาจชนกับ system layout** — script จะ SKIP รายการที่ `Author='System'` อัตโนมัติเพื่อกันความเสี่ยง

⚠️ **Password ใน `_settings.bat`** เก็บเป็น plain text — `.gitignore` exclude ไว้แล้ว แต่อย่าเก็บไฟล์นี้ไว้ใน cloud public

---

## License / Support

Internal tool — Enconfund / SDA Consult Team
Contact: consult@sala-daeng.com
