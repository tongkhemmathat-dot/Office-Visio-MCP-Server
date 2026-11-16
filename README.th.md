# Visio MCP Server (ภาษาไทย)

[English](README.md) | **ไทย**

MCP server ที่ให้ AI (หรือ MCP client ตัวไหนก็ได้) สร้างและแก้ไขไดอะแกรม Microsoft Visio ผ่าน COM automation
ของ Visio: เพิ่มเชป เชื่อมเส้น ใช้ stencil สร้างไดอะแกรมเครือข่ายทั้งภาพจาก JSON spec เดียว ติดตั้งอุปกรณ์ลงแร็ค
ส่งออกเป็น PNG/PDF/SVG และแปลงไฟล์ SVG เป็น `.vsdx` ที่แก้ไขต่อได้

![demo](./public/demo.gif)

## สารบัญ

1. [สิ่งที่ต้องมี](#สิ่งที่ต้องมี)
2. [การติดตั้ง](#การติดตั้ง)
3. [เชื่อมต่อกับ client](#เชื่อมต่อกับ-client)
4. [หลักการทำงาน (อ่านครั้งเดียว)](#หลักการทำงาน-อ่านครั้งเดียว)
5. [เริ่มใช้งานเร็ว](#เริ่มใช้งานเร็ว)
6. [รายการ tool](#รายการ-tool)
7. [ตัวอย่างการใช้งาน](#ตัวอย่างการใช้งาน)
8. [SVG เป็น Visio แบบละเอียด](#svg-เป็น-visio-แบบละเอียด)
9. [แก้ปัญหา](#แก้ปัญหา)
10. [สำหรับนักพัฒนา](#สำหรับนักพัฒนา)

เอกสารอื่น: [คู่มือทีม](docs/TEAM_GUIDE.md) (prompt, stencil, แร็ค), [สถานะงาน SVG](docs/SVG_IMPORT_TODO.md),
[บันทึกการพัฒนา](docs/DEVELOPMENT_NOTES.md), [changelog](CHANGELOG.md) และตัวอย่าง spec ใน [`examples/`](examples)

## สิ่งที่ต้องมี

- Windows (server ควบคุม Visio ผ่าน COM)
- Microsoft Visio แบบติดตั้งในเครื่อง (Standard หรือ Professional) เคยใช้กับ Visio 2013 ขึ้นไป
- Python 3.12 ขึ้นไป
- แพ็กเกจ Python (ลงให้อัตโนมัติด้วย `pip install -e .`):
  `mcp[cli]<2`, `pywin32`, `Pillow` (วัดความกว้างข้อความตอนนำเข้า SVG), `resvg-py` (แปลงไอคอน SVG เป็นรูป)

## การติดตั้ง

```powershell
git clone https://github.com/tongkhemmathat-dot/Office-Visio-MCP-Server.git
cd Office-Visio-MCP-Server
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

ลองเริ่ม server (จะรอ MCP client ทาง stdin จึงดูเหมือนค้าง กด Ctrl+C เพื่อหยุด):

```powershell
python -m visio_mcp_server.visio_server
# หรือหลัง pip install -e .
visio_mcp_server
```

ถ้าไม่ได้ติดตั้ง Visio server จะแจ้งข้อความแล้วปิดตัวเอง

## เชื่อมต่อกับ client

### Claude Code

```powershell
claude mcp add visio -- D:\path\to\repo\.venv\Scripts\python.exe -m visio_mcp_server.visio_server
```

### Claude Desktop หรือ client ที่ใช้ไฟล์ตั้งค่า JSON

```json
{
  "mcpServers": {
    "visio": {
      "command": "D:\\path\\to\\repo\\.venv\\Scripts\\python.exe",
      "args": ["-m", "visio_mcp_server.visio_server"]
    }
  }
}
```

### ไม่ต้อง clone (uvx)

```json
{
  "mcpServers": {
    "visio": {
      "command": "uvx",
      "args": ["--from", "office-visio-mcp-server", "visio_mcp_server"]
    }
  }
}
```

ให้ใช้ Python ที่ลงแพ็กเกจไว้ (`.venv` ของโปรเจกต์) รีสตาร์ท client หลังแก้ config แล้วลองสั่งว่า
*"แสดงรายการ Visio stencil ในเครื่องนี้"* เพื่อเช็กว่าเชื่อมต่อได้

## หลักการทำงาน (อ่านครั้งเดียว)

- **ใช้ Visio ตัวเดียวร่วมกัน** server จะเกาะเข้ากับ Visio ที่เปิดอยู่แล้ว (`GetActiveObject`) ถ้าไม่มีจะเปิดเอง
  และปิดตอน server จบ แต่จะไม่ปิด Visio ที่ตัวเองไม่ได้เปิด ไดอะแกรมที่สร้างจะโผล่ในหน้าต่าง Visio ของคุณ
  จึงควรเซฟหรือปิดงานที่ยังไม่ได้เซฟก่อนทดลอง
- **ระบุไฟล์ด้วย path** ทุก tool รับ `file_path` ถ้าใส่แค่ชื่อไฟล์ (`demo.vsdx`) จะอยู่ในโฟลเดอร์ **Documents**
  tool ส่วนใหญ่เซฟไฟล์ให้หลังแก้ไขทุกครั้ง
- **หน่วยเป็นนิ้ว** จุดเริ่มอยู่ที่มุมล่างซ้าย แกน y ชี้ขึ้น หน้าเริ่มต้นคือ A4 แนวนอน 11.69 x 8.27 นิ้ว
- **shape ID เชื่อมแต่ละขั้นตอน** `add_shape` กับ `add_stencil_shape` คืนค่า shape ID ใช้ต่อกับ `connect_shapes`,
  `add_text`, `delete_shape` ถ้าลืม ใช้ `list_shapes` ดูใหม่
- **สี** ใส่ได้ทั้ง `#RRGGBB` และชื่อ (`red`, `blue`, `teal` ฯลฯ) ลายเส้นมี `solid`, `dash`, `dot`, `dashdot`
  `line_weight` และ `font_size` เป็นพอยต์
- **เทสต์ที่รัน Visio จริง** จะสร้างเอกสารใน Visio ของคุณ ส่วนเทสต์ที่แค่แยกวิเคราะห์ไม่แตะ Visio

## เริ่มใช้งานเร็ว

สั่ง AI เป็นภาษาธรรมดา เช่น

> สร้าง `C:\Diagrams\hello.vsdx` เพิ่มสี่เหลี่ยม "Start" ที่ (2, 4) และวงกลม "End" ที่ (6, 4)
> เชื่อมด้วยลูกศร แล้วส่งออกเป็น PNG

เบื้องหลัง AI จะเรียก `create_visio_file`, `add_shape` สองครั้ง, `connect_shapes` และ `export_diagram`
ส่วนด้านล่างแสดงอาร์กิวเมนต์ของแต่ละ tool

## รายการ tool

อาร์กิวเมนต์เป็น JSON ทั้งหมด ตัวที่เป็น optional ไม่ต้องใส่ก็ได้

| กลุ่ม | Tool |
|---|---|
| ไฟล์ | `create_visio_file`, `open_visio_file`, `close_document`, `set_page_setup`, `export_diagram` |
| เชป | `add_shape`, `add_text`, `connect_shapes`, `list_shapes`, `delete_shape` |
| Stencil | `list_stencils`, `list_stencil_masters`, `search_stencil_shapes`, `add_stencil_shape` |
| ทั้งไดอะแกรม | `build_diagram` |
| แร็ค | `add_rack`, `add_rack_device`, `list_rack` |
| SVG | `svg_to_visio` |

### create_visio_file

สร้างไดอะแกรมใหม่ เอกสารเปล่าจะได้ขนาดหน้าคงที่ Visio จึงไม่รีเซ็ตเป็นขนาดกระดาษของเครื่องพิมพ์

```json
{
  "save_path": "C:\\Diagrams\\MyDiagram.vsdx",
  "template_path": "template .vstx/.vst (ไม่บังคับ)",
  "page_width": 11.69,
  "page_height": 8.27
}
```

`page_width/page_height` ไม่มีผลเมื่อใช้ template

### open_visio_file / close_document

```json
{ "file_path": "C:\\Diagrams\\MyDiagram.vsdx" }
```

```json
{ "file_path": "C:\\Diagrams\\MyDiagram.vsdx", "save_changes": true }
```

### set_page_setup

ตั้งขนาดหน้าแรก และสเกลวาด (ไม่บังคับ)

```json
{ "file_path": "C:\\Diagrams\\rack.vsdx", "width": 11.69, "height": 8.27, "drawing_scale": 12 }
```

`drawing_scale` คือจำนวนนิ้วจริงต่อหนึ่งนิ้วบนหน้า (12 คือ 1:12) stencil จาก VisioCafe (HPE, Dell ฯลฯ) วาดด้วยขนาดจริง
จึงต้องใช้หน้าที่ตั้งสเกล หลังตั้งแล้วพิกัดที่ส่งให้ tool เป็นนิ้วจริง (drawing inch)

### add_shape

```json
{
  "file_path": "C:\\Diagrams\\flow.vsdx",
  "shape_type": "Rectangle",
  "x": 2.0, "y": 4.0, "width": 1.5, "height": 0.8,
  "text": "Start",
  "fill_color": "#DBEAFE",
  "line_color": "#2563EB",
  "line_weight": 1.5,
  "line_pattern": "solid",
  "text_color": "#1E293B",
  "font_size": 12,
  "bold": true,
  "rounding": 0.1
}
```

ตัวอย่าง `shape_type`: `Rectangle`, `Circle`, `Line` ถ้าชนิดไม่รู้จักจะใช้สี่เหลี่ยมแทน `x, y` คือจุดกึ่งกลางเชป
`rounding` คือรัศมีมุมเป็นนิ้ว คืนค่า shape ID

### connect_shapes

```json
{
  "file_path": "C:\\Diagrams\\flow.vsdx",
  "shape1_id": 1,
  "shape2_id": 2,
  "connector_type": "Straight",
  "line_color": "#475569",
  "line_weight": 2,
  "line_pattern": "dash",
  "begin_arrow": false,
  "end_arrow": true,
  "label": "HTTPS",
  "text_color": "#475569",
  "font_size": 10
}
```

`connector_type` มี `Dynamic` (เดินเส้นหักมุมฉากอัตโนมัติ), `Straight` และ `Curved` สองแบบหลังลากจากกึ่งกลางถึงกึ่งกลาง

### add_text

```json
{ "file_path": "C:\\Diagrams\\flow.vsdx", "shape_id": 1, "text": "Start" }
```

### list_shapes / delete_shape

```json
{ "file_path": "C:\\Diagrams\\flow.vsdx" }
```

```json
{ "file_path": "C:\\Diagrams\\flow.vsdx", "shape_id": 3 }
```

### list_stencils / list_stencil_masters / search_stencil_shapes

Stencil คือคลังเชปสำเร็จรูป ("master") เช่น เซิร์ฟเวอร์ สวิตช์ ไฟร์วอลล์ คลาวด์ แร็ค

```json
{ "query": "server" }
```

```json
{ "stencil": "SERVER_U" }
```

```json
{ "query": "DL380", "stencil": "HPE-ProLiant", "limit": 20 }
```

`list_stencils` แสดงไฟล์ที่พบ `list_stencil_masters` แสดงเชปใน stencil หนึ่ง (ชื่อหรือ path เต็ม) และ
`search_stencil_shapes` ค้นชื่อเชปในทุก stencil ที่ชื่อไฟล์มีคำใน `stencil`

### add_stencil_shape

```json
{
  "file_path": "C:\\Diagrams\\net.vsdx",
  "stencil": "SERVER_U",
  "master": "Database server",
  "x": 4.0, "y": 5.0,
  "text": "DB-01",
  "width": 1.2, "height": 1.2,
  "fill_color": "#E2E8F0",
  "font_size": 11,
  "bold": true
}
```

ตำแหน่งคือจุดกึ่งกลาง ถ้าใส่สีพื้น จะแทนที่เงา 3 มิติของ stencil ด้วยสีเรียบ คืนค่า shape ID

### build_diagram

วาดทั้งไดอะแกรมจาก spec เดียว node ถูกจัดเป็นชั้น (tier) เป็นแถว หรือเป็นคอลัมน์เมื่อ `left-right`
ให้พอดีหน้า ดู[ตัวอย่าง B](#b-ไดอะแกรมเครือข่ายจาก-spec-เดียว)

```json
{
  "file_path": "C:\\Diagrams\\net.vsdx",
  "replace_existing": false,
  "spec": {
    "title": "หัวเรื่อง", "subtitle": "บรรทัดรองใต้หัวเรื่อง",
    "page": { "width": 11.69, "height": 8.27 },
    "direction": "top-down",
    "icon_size": 1.0,
    "nodes": [ { "id": "fw", "label": "Firewall", "tier": 1, "stencil": "PERIPH_U", "master": "Firewall" } ],
    "links": [ { "from": "fw", "to": "sw", "label": "10GbE", "type": "straight", "arrow": "end" } ],
    "groups": [ { "label": "Core", "nodes": ["fw", "sw"], "color": "#3B82F6" } ]
  }
}
```

| ฟิลด์ | ความหมาย |
|---|---|
| `direction` | `top-down` (tier 0 อยู่บนสุด ค่าเริ่มต้น) หรือ `left-right` |
| `nodes[].tier` | 0 คือแถว/คอลัมน์แรก node ที่ tier เดียวกันอยู่แถวเดียวกัน |
| `nodes[].stencil`, `master` | ไอคอน (ไม่บังคับ) ถ้าไม่ใส่ node จะเป็นกล่องมีป้ายชื่อ |
| `nodes[].fill_color`, `text_color`, `width`, `height`, `icon_size` | สไตล์/ขนาดรายตัว (ไม่บังคับ) |
| `links[].type` | `straight` (ค่าเริ่มต้น), `curved`, `dynamic` |
| `links[].arrow` | `none` (ค่าเริ่มต้น), `end`, `begin`, `both` |
| `links[].color`, `pattern`, `weight`, `label` | สไตล์เส้นและข้อความ (ไม่บังคับ) |
| `groups` | กรอบเส้นประหลัง node ที่ระบุ พร้อมป้ายชื่อตัวพิมพ์ใหญ่ |
| `replace_existing` | `true` ล้างหน้าก่อน ค่าเริ่มต้นคือเพิ่มต่อจากของเดิม |

คืนค่า JSON ที่มี shape ID ของทุก node

### add_rack / add_rack_device / list_rack

stencil แร็ควาดด้วยนิ้วจริง `add_rack` จึงตั้งสเกลหน้าเป็น 1:12 (A4 แนวนอนกว้างประมาณ 140 x 99 นิ้ว) และพิกัดเป็นนิ้วจริง

```json
{ "file_path": "C:\\Diagrams\\rack.vsdx", "x": 30, "y": 10, "name": "Rack A01" }
```

คืนค่า shape ID ของแร็ค ตัวเลือกอื่น: `stencil` (ค่าเริ่มต้น `Dell-Racks`), `master` (`4220 Rack Frame`),
`u_height` (42), `u1_offset` (4.0), `drawing_scale` (12)

```json
{
  "file_path": "C:\\Diagrams\\rack.vsdx",
  "rack_id": 1,
  "stencil": "HPE-ProLiant-DL",
  "master": "DL380 Gen10 SFF front",
  "u_position": 38,
  "u_size": 2,
  "label": "app-01",
  "callout": true
}
```

`u_position` คือ U ต่ำสุดที่อุปกรณ์ใช้ (U1 อยู่ล่างสุด) ถ้าซ้อนกับของเดิมหรืออยู่นอกแร็คจะถูกปฏิเสธ
`callout` วาดเส้นชี้และป้าย "app-01 (U38-39)"

```json
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1 }
```

`list_rack` คืนรายการอุปกรณ์ ช่วง U ที่ว่าง และจำนวน U ที่ใช้/ว่าง

### export_diagram

```json
{ "file_path": "C:\\Diagrams\\net.vsdx", "output_path": "C:\\Diagrams\\net.png" }
```

นามสกุลกำหนดรูปแบบ: `.png .jpg .jpeg .gif .bmp .svg .pdf` ถ้าหน้ากว้างมากและมีรูปเยอะ ให้ส่งออก `.pdf` หรือ `.svg`
แทนบิตแมป

### svg_to_visio

```json
{
  "svg_path": "C:\\Diagrams\\architecture.svg",
  "output_path": "C:\\Diagrams\\architecture.vsdx",
  "export_png": "C:\\Diagrams\\architecture.check.png",
  "group_shapes": true
}
```

`export_png` และ `group_shapes` ไม่บังคับ (จัดกลุ่มเปิดไว้โดยค่าเริ่มต้น) รายละเอียดดู
[SVG เป็น Visio แบบละเอียด](#svg-เป็น-visio-แบบละเอียด)

## ตัวอย่างการใช้งาน

### A. Flowchart ทีละเชป

ลำดับการเรียกของ AI (ย่อ path):

```json
// 1. create_visio_file
{ "save_path": "C:\\Diagrams\\FlowChart.vsdx" }

// 2. add_shape  -> ได้ ID 1
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "shape_type": "Rectangle",
  "x": 2, "y": 4, "width": 1.5, "height": 0.8, "text": "Start", "fill_color": "#DBEAFE" }

// 3. add_shape  -> ได้ ID 2
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "shape_type": "Circle",
  "x": 6, "y": 4, "width": 1, "height": 1, "text": "End", "fill_color": "#DCFCE7" }

// 4. connect_shapes
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "shape1_id": 1, "shape2_id": 2,
  "connector_type": "Straight", "end_arrow": true, "label": "next" }

// 5. export_diagram
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "output_path": "C:\\Diagrams\\FlowChart.png" }
```

### B. ไดอะแกรมเครือข่ายจาก spec เดียว

`build_diagram` ใช้การเรียกครั้งเดียว นี่คือ [`examples/three_tier_app.json`](examples/three_tier_app.json) แบบย่อ:

```json
{
  "title": "Three-tier Web Application",
  "direction": "left-right",
  "nodes": [
    {"id": "users", "label": "Users",         "tier": 0, "fill_color": "#E2E8F0"},
    {"id": "lb",    "label": "Load Balancer", "tier": 1, "fill_color": "#FEF3C7"},
    {"id": "web1",  "label": "Web 1", "tier": 2, "stencil": "SERVER_U", "master": "Web server"},
    {"id": "web2",  "label": "Web 2", "tier": 2, "stencil": "SERVER_U", "master": "Web server"},
    {"id": "db",    "label": "Database", "tier": 3, "stencil": "SERVER_U", "master": "Database server"}
  ],
  "links": [
    {"from": "users", "to": "lb",   "arrow": "end", "label": "HTTPS"},
    {"from": "lb",    "to": "web1", "arrow": "end"},
    {"from": "lb",    "to": "web2", "arrow": "end"},
    {"from": "web1",  "to": "db",   "arrow": "end"},
    {"from": "web2",  "to": "db",   "arrow": "end"}
  ],
  "groups": [ {"label": "Web tier", "nodes": ["web1", "web2"], "color": "#16A34A"} ]
}
```

หาชื่อ stencil และ master: สั่ง `list_stencils` พร้อม query (`server`) แล้ว `list_stencil_masters` ของตัวที่ต้องการ
(`SERVER_U` มี Web server, Database server ฯลฯ) หรือใช้ `search_stencil_shapes` เมื่อรู้ชื่อรุ่นแค่บางส่วน

### C. ไดอะแกรมแร็ค

```json
// 1. add_rack -> ได้ rack_id
{ "file_path": "C:\\Diagrams\\rack.vsdx", "x": 30, "y": 10, "name": "Rack A01" }

// 2. add_rack_device (ทำซ้ำสำหรับแต่ละอุปกรณ์)
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1,
  "stencil": "Juniper EX Series", "master": "EX4300-24T Front", "u_position": 40 }
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1,
  "stencil": "HPE-ProLiant-DL", "master": "DL380 Gen10 SFF front",
  "u_position": 38, "u_size": 2, "label": "app-01" }

// 3. list_rack: ดูว่าติดตั้งอะไรแล้วและ U ไหนว่าง
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1 }
```

### D. SVG เป็น Visio

สั่งว่า *"แปลง `C:\Diagrams\architecture.svg` เป็น `C:\Diagrams\architecture.vsdx` และส่งออก PNG ไว้เทียบ"*
tool ที่ถูกเรียก:

```json
{
  "svg_path": "C:\\Diagrams\\architecture.svg",
  "output_path": "C:\\Diagrams\\architecture.vsdx",
  "export_png": "C:\\Diagrams\\architecture.check.png"
}
```

ผลลัพธ์ตัวอย่าง (ตัวเลขขึ้นกับ SVG):

```
Created C:\Diagrams\architecture.vsdx (13.33 x 7.50 in): 150 lines, 31 pictures, 60 texts. Grouped 12 SVG group(s). Skipped or approximated 2 thing(s): filter, gradient. Exported to: C:\Diagrams\architecture.check.png
```

เปิด `.vsdx` ใน Visio: ทุกองค์ประกอบเป็นเชปจริงที่แก้ไขได้ และตั้งชื่อตาม `id` ใน SVG

### E. รันโดยไม่ใช้ AI client

สคริปต์ใน `examples/` เรียกฟังก์ชันเดียวกันโดยตรง (ใช้ `.venv` ของโปรเจกต์):

```powershell
# JSON spec -> .vsdx (+ PNG และ PDF)
python examples\build_from_spec.py examples\hci_network.json out\hci.vsdx --png --pdf

# SVG -> .vsdx พร้อมภาพเทียบ Edge กับ Visio ต่อไฟล์
python examples\svg_compare.py out my1.svg my2.svg
```

`svg_compare.py` เขียน `.vsdx`, `.visio.png`, `.edge.png` และ `.side.png` (สองภาพวางคู่กัน) ให้ตัดสินจากภาพวางคู่
คะแนนความต่างที่พิมพ์ออกมาเป็นแค่ค่าคร่าวๆ

## SVG เป็น Visio แบบละเอียด

**เป้าหมาย:** ใส่ SVG แล้วได้ไฟล์ Visio ที่แก้ไขต่อได้ ตำแหน่ง สี ข้อความ และลำดับชั้นเหมือนต้นฉบับ
แต่ละองค์ประกอบของ SVG กลายเป็นเชป Visio ที่เลือก เปลี่ยนสไตล์ และย้ายได้

### แปลงอะไรได้บ้าง

| SVG | ผลใน Visio |
|---|---|
| `<g>`, transform (translate, scale, rotate, skew, matrix), `viewBox` | คำนวณตำแหน่งให้แล้ว เชปที่มาจาก `<g>` เดียวกันและอยู่ติดกันจะรวมเป็นกลุ่ม Visio |
| `rect` (มี `rx/ry`), `circle`, `ellipse` | เชปสี่เหลี่ยม (มุมมน) / วงรี |
| `line`, `polyline`, `polygon` | เส้นพร้อมลายเส้นประและหัวลูกศร |
| `path` (เส้นตรง โค้ง ส่วนโค้ง) | เส้นโค้งถูกแปลงเป็นเส้นตรงสั้นๆ `path` ที่มีหลาย subpath เป็นเชปเดียว เติมสีแบบ even-odd รูจึงยังเป็นรู |
| `text`, `tspan` | เชปข้อความ เลือกฟอนต์ ขนาด ตัวหนา การจัดชิด และวัดความกว้างต่อข้อความ |
| ป้าย HTML ใน `<foreignObject>` ของ draw.io | อ่านเป็นข้อความ พร้อมตัดบรรทัด การจัดชิด และ bullet ไม่วาดข้อความ "Text is not SVG" |
| `<use>`, `<symbol>`, `<switch>` | ขยายให้ครบ `switch` เลือกกิ่งเดียวกับที่เบราว์เซอร์เลือก |
| `<style>`, class, `#id`, `style=""`, การสืบทอด | คำนวณก่อนวาด |
| opacity, `fill-opacity`, `stroke-opacity` | ความโปร่งใสของสีพื้น เส้น และข้อความ |
| หัวลูกศรจาก `<marker>` | ชนิดและขนาดหัวลูกศรตาม marker |
| `<image>` PNG/JPEG, SVG data URI, ไฟล์ข้าง SVG | เป็นรูป ไอคอน SVG ที่ฝังอยู่ถูกแปลงเป็นรูปด้วย `resvg-py` (Edge แบบ headless ใช้เป็นตัวสำรอง) |

### การตั้งชื่อ เลเยอร์ กลุ่ม

- เชปตั้งชื่อตาม `id` ของ SVG (ทำให้ไม่ซ้ำกัน)
- เลเยอร์: `Shapes`, `Connectors`, `Labels`, `Icons` เปิด/ปิดหรือล็อกได้ใน Layer Properties ของ Visio
- เมื่อ `group_shapes: true` (ค่าเริ่มต้น) เชปที่ติดกันจาก `<g>` เดียวกันจะถูกรวมเป็นกลุ่ม ชื่อกลุ่มคือ id ของ `<g>`
  ตั้งเป็น `false` ถ้าต้องการรายการเชปแบบแบน

### ที่ยังไม่รองรับ (แจ้งในผลลัพธ์)

- `clip-path`, `mask`, `filter`, `<pattern>`, `<svg>` ซ้อนใน `<svg>`, `<symbol>` ที่มี `viewBox` ของตัวเอง
- gradient ใช้สีเฉลี่ยแทน
- รูปที่ถูกหมุนจะวาดไม่หมุน และไม่ใช้ `preserveAspectRatio` ของรูป
- `textLength`, `dx/dy` รายตัวอักษร และ `<tspan>` ที่สไตล์ต่างกันในบรรทัดเดียว
- เส้นเชื่อมเป็นเส้นธรรมดา ไม่ได้ glue เข้ากับเชปปลายทาง
- การเติมสีแบบ nonzero จะวาดเป็น even-odd (ต่างกันเฉพาะ subpath ซ้อนกันที่ทิศทางเดียวกัน)

### ความเร็วและเคล็ดลับ

- เวลาเพิ่มตามจำนวนองค์ประกอบ ไดอะแกรมที่มีเส้นราว 200 เส้น 16 path และ 14 รูปใช้ประมาณ 90 วินาที
  เพราะเขียนลง Visio ผ่าน COM ทีละ cell
- ถ้าต้องการความเหมือนสูง ใส่ `export_png` แล้วเทียบกับ SVG
- ต้องการผลที่แบนและเร็วขึ้น ใช้ `group_shapes: false`
- ฟอนต์: ชื่อฟอนต์ถูกจับคู่กับฟอนต์ที่ติดตั้ง (serif เป็น Times New Roman, sans-serif เป็น Arial,
  monospace เป็น Consolas, Helvetica เป็น Arial, system-ui เป็น Segoe UI) ติดตั้งฟอนต์ที่ SVG ใช้จะได้ข้อความพอดีที่สุด

## แก้ปัญหา

| อาการ | วิธีแก้ |
|---|---|
| server ปิดตัวทันที | ไม่ได้ติดตั้ง Visio หรือ COM ไม่ได้ลงทะเบียน ลองเปิด Visio เองหนึ่งครั้ง |
| `Visio could not render this page as .png` | หน้ากว้างเกินไป ส่งออก `.pdf` หรือ `.svg` แทน |
| COM error หรือ tool ค้าง | ปิดกล่องโต้ตอบใน Visio (dialog ที่ค้างอยู่ขวาง automation) ดู Task Manager ว่ามี `VISIO.EXE` ที่เราเปิดแล้วค้างหรือไม่ แล้วปิดเฉพาะตัวนั้น |
| ชนิดเชปไม่รู้จัก | จะใช้สี่เหลี่ยมแทน เช็กตัวสะกด หรือใช้ master จาก stencil |
| หา stencil หรือ master ไม่เจอ | ใช้ `list_stencils`, `list_stencil_masters` หรือ `search_stencil_shapes` ดูชื่อที่ถูกต้อง |
| เชป stencil เล็กหรือใหญ่เกินไป | เชปเหล่านี้ใช้ขนาดจริง ใช้ `set_page_setup` พร้อม `drawing_scale` (หรือ `add_rack` ที่ตั้งให้แล้ว) |
| ไดอะแกรมที่เปิดอยู่ถูกเปลี่ยน | server ทำงานใน Visio ที่คุณเปิดอยู่ เซฟและปิดงานของคุณก่อนทดลอง |
| ข้อความใน SVG ตัดบรรทัดหรือตำแหน่งต่าง | ติดตั้งฟอนต์ของ SVG หรือเทียบด้วย `examples/svg_compare.py` |
| ผลจาก SVG ดูแบน | gradient, filter และ mask ยังไม่รองรับ ดูรายการข้างบน |
| `svg_to_visio` แจ้งว่า output ต้องลงท้าย `.vsdx` | แก้ `output_path` |

## สำหรับนักพัฒนา

```powershell
pip install -e ".[test]"
python -m pytest tests --deselect tests/test_e2e_visio.py   # เร็ว ไม่แตะ Visio
python -m pytest tests                                       # รันครบ ต้องมี Visio ปิด Visio ของคุณก่อน
```

การแยกวิเคราะห์สไตล์ การจัด tier การคำนวณแร็ค และการอ่าน SVG เทสต์ได้โดยไม่ต้องมี Visio ส่วนเทสต์ e2e
สร้างเอกสารใน Visio จริง

โครงสร้างโค้ด:

| ไฟล์ | หน้าที่ |
|---|---|
| `visio_mcp_server/visio_server.py` | MCP tool และการเรียก Visio COM |
| `visio_mcp_server/styles.py` | แยกวิเคราะห์สี เส้น และสไตล์ข้อความ |
| `visio_mcp_server/diagram.py` | ตรวจ spec จัด tier และคำนวณแร็ค (ไม่ต้องใช้ Visio) |
| `visio_mcp_server/svg_import.py` | อ่าน SVG เป็นองค์ประกอบพื้นฐาน (rect, ellipse, polyline, path, text, image) |
| `visio_mcp_server/svg_html.py` | อ่านป้าย HTML ของ draw.io |
| `visio_mcp_server/svg_text.py` | หาฟอนต์และวัดความกว้างข้อความ |
| `examples/` | spec และสคริปต์ที่รันได้โดยไม่ต้องใช้ AI client |

ดู [CHANGELOG.md](CHANGELOG.md) และ [docs/DEVELOPMENT_NOTES.md](docs/DEVELOPMENT_NOTES.md) สำหรับข้อควรระวังเรื่อง
Visio automation ที่อยู่เบื้องหลังโค้ด

## ลิขสิทธิ์

MIT ดูไฟล์ LICENSE
