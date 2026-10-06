# คู่มือการใช้งาน Visio MCP Server สำหรับทีม

คู่มือนี้สำหรับ system engineer ที่ใช้ Claude ช่วยวาง diagram ใน Microsoft Visio เช่น network diagram,
system architecture, rack diagram และ HCI/virtualization diagram โดยไม่ต้องลากวางเอง

> ใช้ได้เฉพาะ **Windows ที่ติดตั้ง Microsoft Visio** (ทดสอบกับ Visio 2013 ขึ้นไป) เซิร์ฟเวอร์ควบคุม Visio
> ผ่าน COM จึงไม่ทำงานบน Mac/Linux และไม่ใช้กับ Visio for the web

---

## 1. ทำอะไรได้บ้าง

| งาน | Tool ที่ใช้ |
|---|---|
| วาด diagram ทั้งใบจากคำสั่งเดียว (จัดตำแหน่ง ต่อสาย ใส่กรอบกลุ่มให้เอง) | `build_diagram` |
| วาง rack 42U แล้วใส่ server/switch ตามเลข U (กันซ้อนทับ/เกินช่อง) | `add_rack`, `add_rack_device`, `list_rack` |
| ใช้ไอคอนอุปกรณ์จริงจาก stencil (HPE, Dell, Juniper, Aruba, Cisco, Synology, VMware ...) | `list_stencils`, `search_stencil_shapes`, `add_stencil_shape` |
| ส่งออกเป็น PNG / PDF / SVG ไว้แปะเอกสาร | `export_diagram` |
| ปรับแต่งเอง ทีละ shape (สี เส้น ลูกศร ข้อความ) | `add_shape`, `connect_shapes`, `add_text`, `delete_shape` |

ตัวอย่างที่เห็นผลทันทีอยู่ที่โฟลเดอร์ [`examples/`](../examples)

---

## 2. ติดตั้ง (ทำครั้งเดียวต่อเครื่อง)

ต้องมี Python 3.10 ขึ้นไป

```powershell
git clone https://github.com/tongkhemmathat-dot/Office-Visio-MCP-Server.git
cd Office-Visio-MCP-Server
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
```

ตรวจว่าใช้ได้ (จะเปิด Visio ขึ้นมาแล้วปิดเอง):

```powershell
.\.venv\Scripts\pip install pytest
.\.venv\Scripts\python -m pytest -q
```

ถ้าเครื่องไม่มี Visio ชุดทดสอบส่วนที่ต้องใช้ Visio จะถูกข้ามโดยอัตโนมัติ

## 3. เชื่อมกับ Claude

**Claude Code** (แทน `<โฟลเดอร์>` ด้วย path ที่ clone ไว้):

```powershell
claude mcp add visio_server -- <โฟลเดอร์>\.venv\Scripts\python.exe <โฟลเดอร์>\visio_mcp_server\visio_server.py
```

**Claude Desktop** เพิ่มในไฟล์ `%APPDATA%\Claude\claude_desktop_config.json` แล้วเปิดแอปใหม่:

```json
{
  "mcpServers": {
    "visio_server": {
      "command": "<โฟลเดอร์>\\.venv\\Scripts\\python.exe",
      "args": ["<โฟลเดอร์>\\visio_mcp_server\\visio_server.py"]
    }
  }
}
```

> ไฟล์ `mcp-config.json` ใน repo เป็นตัวอย่างที่ชี้ path ของเครื่องผู้พัฒนา ต้องแก้ path ก่อนใช้

เริ่ม session ใหม่แล้วถาม Claude ว่า "list the Visio tools" ถ้าเห็น tool 18 ตัวแสดงว่าพร้อมใช้

**ข้อควรรู้เรื่อง Visio ที่เปิดอยู่**
- ถ้าคุณเปิด Visio ไว้แล้ว เซิร์ฟเวอร์จะต่อเข้ากับตัวนั้นและ **ไม่ปิด Visio ของคุณ** (ปิดเฉพาะเอกสารที่มันเปิดเอง)
- ถ้า Visio ยังไม่เปิด เซิร์ฟเวอร์จะเปิดให้และปิดเมื่อจบ session
- เซิร์ฟเวอร์บันทึกไฟล์ทุกครั้งที่แก้ จึงไม่ควรเปิดไฟล์เดียวกันแก้ด้วยมือพร้อมกัน

---

## 4. Stencil (ไอคอนอุปกรณ์)

เซิร์ฟเวอร์ค้นหา stencil ใน 3 ที่ (รวมโฟลเดอร์ย่อย): `Visio Content` ของ Visio, `Documents\My Shapes` และ
path ที่ตั้งใน Visio ดังนั้นวางไฟล์ `.vssx`/`.vss` ไว้ใน **`Documents\My Shapes`** แล้วเรียกใช้ด้วยชื่อไฟล์ได้เลย
(เช่น `HPE-ProLiant-DL`, `Juniper EX Series`, `Dell-Racks`)

| ยี่ห้อ | แหล่งดาวน์โหลด | ข้อควรระวัง |
|---|---|---|
| HPE (ProLiant, rack, Aruba), Dell, Dell EMC | VisioCafe (visiocafe.com) | ใช้สร้างเอกสารภายในเท่านั้น ห้ามแจกจ่ายต่อและห้ามแก้ไข shape |
| Cisco | cisco.com → Brand Center → Network Topology Icons | ใช้ได้ฟรี ห้ามแก้ไข |
| Juniper | juniper.net → Icons and Stencils | ห้ามใช้โปรโมตสินค้ายี่ห้ออื่น |
| Synology | GitHub `bhdicaire/visioStencils` | ชุดรวมจากชุมชน |
| Sangfor | Sangfor Support Portal (ต้องล็อกอิน) | ไม่มีให้โหลดสาธารณะ |

- **ห้าม commit stencil ขึ้น git** (ลิขสิทธิ์ของเจ้าของ) โฟลเดอร์ `My Shapes` อยู่นอก repo อยู่แล้ว
- stencil ของ VisioCafe วาดด้วย **หน่วยจริง** (เช่น server กว้าง 19 นิ้ว) ดูหัวข้อ rack ด้านล่าง
- ค้นหา shape: บอก Claude ว่า "หา DL380 Gen11 ใน stencil HPE-ProLiant" (`search_stencil_shapes`)
- ชื่อไฟล์ซ้ำกัน (เช่นมีทั้ง `.vss` และ `.vssx`) ระบบเลือก `.vssx` ให้ ถ้าชื่อกำกวมจะแจ้งรายการให้เลือก

---

## 5. วิธีสั่งงานผ่าน Claude

บอกเป็นภาษาธรรมดาได้เลย Claude จะเลือก tool เอง ตัวอย่าง:

**วาด network diagram**
> วาด network diagram ลง `D:\Diagrams\site-a.vsdx`: Internet → Firewall → Core Switch 2 ตัว →
> เซิร์ฟเวอร์ 3 เครื่อง ใช้ไอคอนจาก stencil ของ Visio ใส่กรอบ "HCI cluster" รอบเซิร์ฟเวอร์ แล้ว export เป็น PNG ด้วย

**rack diagram**
> สร้าง rack Dell 42U ชื่อ "Rack A01" ใส่ HPE DL380 Gen10 5 เครื่อง (2U) ที่ U30–U39 และ Juniper EX4300
> 2 ตัวที่ U41–U42 แล้วบอกว่าเหลือกี่ U

**แก้ไฟล์เดิม**
> เปิด `D:\Diagrams\site-a.vsdx` เปลี่ยนเส้นระหว่าง core1 กับ core2 เป็นสีแดงเส้นประ แล้วเพิ่มป้าย "LAG"

คำแนะนำเวลาสั่ง
- ระบุ **path เต็มของไฟล์** ทุกครั้ง (ถ้าใส่แค่ชื่อไฟล์ จะไปอยู่ที่โฟลเดอร์ Documents)
- บอกชื่อ stencil/shape ที่ต้องการถ้ารู้ ถ้าไม่รู้ให้ Claude ค้นหาให้ก่อน
- ถ้า diagram ใหญ่ ให้ใช้ `build_diagram` (สั่งครั้งเดียวทั้งใบ) แทนวางทีละ shape จะเร็วและเรียบร้อยกว่ามาก

---

## 6. `build_diagram` — วาดทั้งใบจาก spec

`spec` คือ JSON ที่บอกว่ามีอะไร ต่อกับอะไร อยู่ชั้นไหน เซิร์ฟเวอร์จัดตำแหน่งให้พอดีหน้ากระดาษ

| ฟิลด์ | ความหมาย |
|---|---|
| `title`, `subtitle` | หัวเรื่อง (ไม่ใส่ก็ได้) |
| `direction` | `top-down` (tier 0 อยู่บนสุด, ค่าเริ่มต้น) หรือ `left-right` |
| `icon_size` | ขนาดด้านที่ยาวที่สุดของไอคอน เป็นนิ้ว (ค่าเริ่มต้น 1.0) |
| `page` | `{"width": 11.69, "height": 8.27}` ค่าเริ่มต้น A4 แนวนอน |
| `nodes[]` | `id` (ห้ามซ้ำ), `label`, `tier` (0 = แถวแรก), `stencil` + `master` (ใส่คู่กัน ไม่ใส่ = กล่องสี่เหลี่ยม), `fill_color`, `text_color`, `icon_size`, `width`, `height` |
| `links[]` | `from`, `to` (id ของ node), `label`, `type` (`straight`/`curved`/`dynamic`), `color`, `pattern` (`solid`/`dash`/`dot`/`dashdot`), `weight`, `arrow` (`none`/`end`/`begin`/`both`) |
| `groups[]` | `label`, `nodes` (รายการ id), `color` → กรอบประเส้นล้อมกลุ่ม |

ตัวอย่างขั้นต่ำ:

```json
{
  "title": "Office network",
  "nodes": [
    {"id": "fw", "label": "Firewall", "tier": 0, "stencil": "PERIPH_U", "master": "Firewall"},
    {"id": "sw", "label": "Core switch", "tier": 1, "stencil": "PERIPH_U", "master": "Switch"},
    {"id": "nas", "label": "NAS", "tier": 2}
  ],
  "links": [{"from": "fw", "to": "sw"}, {"from": "sw", "to": "nas", "label": "10GbE"}]
}
```

- สีเป็นรูป `#RRGGBB` หรือชื่อ (`red`, `blue`, `green`, `orange`, `teal`, `purple`, `gray` ...)
- ถ้าข้อมูลผิด (id ซ้ำ, link ไป node ที่ไม่มี, ชนิดเส้นไม่ถูก) จะได้ข้อความแจ้ง **ก่อนวาดอะไรลงไฟล์**
- ใส่ `replace_existing: true` เพื่อล้างหน้าก่อนวาดใหม่ (ค่าเริ่มต้นคือวาดต่อจากของเดิม)
- รันจาก command line ได้โดยไม่ผ่าน Claude (เหมาะกับ CI/สคริปต์):

```powershell
.\.venv\Scripts\python examples\build_from_spec.py examples\hci_network.json out\hci.vsdx --png --pdf
```

---

## 7. Rack diagram (วางอุปกรณ์ตามเลข U)

ลำดับการเรียก:

1. **`add_rack`** — วางตู้ rack ได้ `rack_id` กลับมา
2. **`add_rack_device`** — ใส่อุปกรณ์ด้วย `rack_id`, stencil, shape, `u_position` (U1 = ล่างสุด) และ `u_size`
3. **`list_rack`** — ดูว่าอะไรอยู่ที่ U ไหน เหลือช่องว่างตรงไหน กี่ U

ตัวอย่างค่า: `add_rack(x=30, y=10, name="Rack A01")` → ได้ rack Dell 4220 (42U) ทางซ้ายของหน้า

หลักการที่ต้องรู้
- **พิกัดเป็นนิ้วจริง** stencil ของ HPE/Dell/Juniper วาดด้วยขนาดจริง (19 นิ้ว, 1U = 1.75 นิ้ว) `add_rack`
  จึงตั้งสเกลหน้าเป็น 1:12 ให้ (A4 แนวนอน = 140 × 99 นิ้วจริง) ถ้าวาด shape อื่นในหน้าเดียวกันด้วยมือ ต้องใช้หน่วยเดียวกัน
- เซิร์ฟเวอร์จะ **ปฏิเสธ** ถ้าอุปกรณ์เกินตู้ (เช่น U42–43 ในตู้ 42U) หรือทับกับอุปกรณ์เดิม และบอกว่าทับกับตัวไหน
- ข้อมูลตำแหน่งเก็บในตัวไฟล์ `.vsdx` เอง ปิดแล้วเปิดใหม่ `list_rack` ยังอ่านได้
- `u_size` ใช้คำนวณที่ว่างเท่านั้น ขนาดภาพมาจาก stencil เอง (DL380 = 2U, สวิตช์ 1U) ระบุให้ตรงกับของจริง
- ใช้ตู้อื่นได้โดยระบุ `stencil` และ `master` ของ frame เอง แล้วปรับ `u1_offset` (ระยะจากขอบล่างของ frame ถึงฐาน U1)
  ค่าเริ่มต้น 4.0 ใช้กับ Dell 4220 ถ้าใช้ frame อื่นแล้วอุปกรณ์เหลื่อม ให้ปรับค่านี้

---

## 8. รายการ Tool ทั้งหมด

| Tool | หน้าที่ |
|---|---|
| `create_visio_file` | สร้างไฟล์ใหม่ (A4 แนวนอน ขนาดกำหนดเอง) หรือจาก template |
| `open_visio_file` / `close_document` | เปิด / ปิด (บันทึกก่อนปิดเป็นค่าเริ่มต้น) |
| `set_page_setup` | ตั้งขนาดหน้าและสเกล |
| `build_diagram` | วาดทั้งใบจาก spec |
| `add_rack` / `add_rack_device` / `list_rack` | rack diagram |
| `list_stencils` | ดู stencil ในเครื่อง (กรองด้วยชื่อได้) |
| `list_stencil_masters` | ดู shape ใน stencil |
| `search_stencil_shapes` | ค้นหา shape ตามชื่อ ภายใน stencil ที่ชื่อไฟล์ตรง |
| `add_stencil_shape` | วาง shape จาก stencil (ใส่สี/ฟอนต์ได้ การใส่สีพื้นจะทำให้ไอคอนแบนไม่มีเงา 3 มิติ) |
| `add_shape` | สี่เหลี่ยม วงกลม วงรี เส้น พร้อมสไตล์ |
| `connect_shapes` | ต่อ shape สองตัว (ใส่สี ลูกศร ป้ายได้) |
| `add_text`, `list_shapes`, `delete_shape` | ใส่ข้อความ / ดูรายการ / ลบ |
| `export_diagram` | ส่งออก `.png` `.jpg` `.gif` `.bmp` `.svg` `.pdf` (ดูจากนามสกุลไฟล์) |

---

## 9. แนวปฏิบัติของทีม (แนะนำ)

- **ตั้งชื่อไฟล์**: `<site>-<ประเภท>-<yyyymmdd>.vsdx` เช่น `bkk-hci-network-20261006.vsdx` และเก็บ PNG/PDF คู่กัน
- **Tier สม่ำเสมอ**: Internet/WAN → Firewall → Core/Distribution → Server/Compute → Storage/Backup
- **สีเส้น**: เทาเข้ม = เครือข่ายธุรกิจ/จัดการ, เขียวน้ำเงิน เส้นประ = storage (SAN/vSAN), แดง = เส้นที่ต้องเน้น
  (เช่น link ที่ยังไม่ redundant)
- **ชื่ออุปกรณ์** ในป้ายให้ตรงกับชื่อใน CMDB/โฮสต์เนมจริง และใส่รุ่น + U ใน rack diagram
- **อย่าใส่ข้อมูลลับในป้าย** (รหัสผ่าน, key) และระวัง IP/ชื่อภายในเมื่อส่งไฟล์ออกนอกองค์กร
- งานที่ต้องแก้ซ้ำบ่อย ให้เก็บ spec JSON ไว้ใน git แล้วสร้าง diagram ใหม่จาก spec แทนการแก้ `.vsdx` ด้วยมือ
- ตรวจ diagram ที่ Claude สร้างทุกครั้งก่อนส่งต่อ: ตำแหน่ง U ใน rack เป็นการคำนวณจาก stencil ไม่ได้ตรวจกับรูยึดจริง

---

## 10. แก้ปัญหา

| อาการ | สาเหตุ / วิธีแก้ |
|---|---|
| ไม่เห็น tool ใน Claude | ยังไม่ได้เริ่ม session ใหม่หลังลงทะเบียน หรือ path ผิด ตรวจด้วย `claude mcp list` |
| `Microsoft Visio is not installed` | เครื่องไม่มี Visio (หรือเป็น Visio for the web) |
| `Stencil not found` | วางไฟล์ไว้ใน `Documents\My Shapes` หรือยัง ลอง `list_stencils` ดูชื่อจริง |
| `Stencil '...' is ambiguous` | ชื่อตรงหลายไฟล์ ใช้ชื่อเต็มหรือ path เต็ม |
| `Master '...' not found` | ชื่อ shape ไม่ตรง ใช้ `list_stencil_masters` / `search_stencil_shapes` |
| `... overlaps '...'` / `does not fit` | เลข U ทับหรือเกินตู้ เรียก `list_rack` ดูช่องว่าง |
| งานล้นหน้ากระดาษ | หน้าตั้งเป็นแนวนอนขนาดกำหนดเองแล้ว ถ้ายังล้น ใช้ `set_page_setup` ขยายหน้า หรือแบ่งหลายไฟล์ |
| ไอคอนใหญ่/เล็กผิดปกติ | stencil แบบหน่วยจริง: ใช้ `add_rack` (ตั้งสเกลให้) หรือ `build_diagram` (ปรับขนาดไอคอนให้) |
| ป้ายบนเส้นมีพื้นหลังเทา | ข้อจำกัดที่ทราบ ยังไม่ได้แก้ |
| `Cell is guarded` | เซิร์ฟเวอร์ข้ามให้เอง ถ้าเห็นใน log แสดงว่า shape นั้นล็อกบางส่วน |
| Visio ค้าง/ไม่ตอบสนอง | ปิด Visio ทั้งหมดแล้วลองใหม่ ตรวจว่าไม่มีกล่องโต้ตอบเปิดค้างอยู่ |

---

## 11. สำหรับผู้พัฒนา

```
visio_mcp_server/
  visio_server.py   tool ทั้งหมด + การคุยกับ Visio (COM)
  diagram.py        ตรวจ spec, จัด layout, คำนวณ rack (ไม่แตะ Visio ทดสอบได้เลย)
  styles.py         แปลงสี/เส้น และใส่สไตล์ให้ shape
tests/              pytest (unit + e2e ที่ใช้ Visio จริง ข้ามเองถ้าไม่มี Visio)
examples/           ตัวอย่าง spec และ build_from_spec.py
```

- เพิ่ม tool ใหม่: เขียนฟังก์ชัน `async def` ใน `visio_server.py` ติด decorator `@tool("ชื่อการกระทำ")`
  error จะถูกแปลงเป็นข้อความ `Error <การกระทำ>: <เหตุผล>` ให้อัตโนมัติ
- แยกส่วนที่เป็นตรรกะล้วนไว้ใน `diagram.py`/`styles.py` แล้วเขียน unit test เพื่อให้ CI รันได้โดยไม่ต้องมี Visio
- รันทดสอบ: `pytest` (ส่วน e2e จะเปิด Visio จริง ปิดงานที่ยังไม่บันทึกก่อน)
