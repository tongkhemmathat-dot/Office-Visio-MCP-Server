# svg_to_visio: สถานะงานและสิ่งที่ยังไม่ได้ทำ

อัปเดตล่าสุด: 2026-11-16 · branch `main`

## ทำเสร็จและ commit แล้ว

| Commit | เนื้อหา |
|---|---|
| `398b2d5` | tool `svg_to_visio` และตัวอ่าน SVG `svg_import.py` |
| `882ee87` | group/transform/viewBox, circle, ellipse, line, polygon, path, rx, CSS, opacity, `<use>`/`<symbol>` |
| `3e87961` | path หลาย subpath เป็นเชปเดียว (เจาะรูได้), `<switch>`, `examples/svg_compare.py` |
| `4c69a3e` | ฟอนต์ต่อข้อความ และวัดความกว้างข้อความจริงด้วย PIL (`svg_text.py`) |
| `7f2af4c` | อ่านข้อความ HTML ใน `<foreignObject>` ของ draw.io (`svg_html.py`) |
| `7642e2a` | ขนาดและชนิดหัวลูกศรตาม `<marker>` |
| `5b03f61` | แปลงไอคอน SVG เป็น PNG ด้วย `resvg-py` (Edge เป็น fallback) |
| `584c712` | ตั้งชื่อเชปตาม SVG `id`, แยกเลเยอร์ (Shapes/Connectors/Labels/Icons), จัดกลุ่มตาม `<g>` |

เทสต์ทั้งหมด 134 ตัวผ่าน (`python -m pytest tests`, ส่วน e2e ต้องมี Visio)
**commit หลัง `9feeba3` ยังไม่ได้ push ไป origin** (ตรวจด้วย `git status -sb`)

## ข้อควรระวังเมื่อทดสอบ

- **อย่าสั่ง `taskkill /IM VISIO.EXE`** มันปิด Visio ทุกตัวรวมของผู้ใช้ ดู process ด้วย `Get-Process VISIO` (ชื่อหน้าต่างบอกไฟล์ที่เปิด)
  และอย่าใช้ `tasklist | grep -i visio` เพราะตรงกับ `AppProvisioningPlugin.exe` ด้วย
- server จะ attach เข้า Visio ที่เปิดอยู่ (`GetActiveObject`) การทดสอบจึงสร้างเอกสารในหน้าต่างของผู้ใช้
  ให้ผู้ใช้ปิด Visio ก่อน หรือปิดเฉพาะ PID ของตัวที่เราเปิดเอง
- ไฟล์ขนาดใหญ่ช้า: ชุด Kubernetes (201 polyline + 16 path + 14 รูป) ใช้ราว 90 วินาที เพราะเขียนทีละแถวผ่าน COM

## จงใจไม่ทำ (และเหตุผล)

- **gradient จริง**: ตั้ง `FillGradientEnabled = TRUE` แล้วเชปกลายเป็นสีขาวทั้งเชป (cell เป็น `THEMEVAL()` ต้องตั้ง
  `FillGradientDir/Angle` และสี stop ให้ครบ อาจต้องใช้ section `FillGradient`) ตอนนี้ใช้สีเฉลี่ยของ stop
- **glue เส้นเชื่อมเข้ากับเชป**: polyline ที่สร้างเป็นเชป 2 มิติ ไม่ใช่ connector 1 มิติ จึงยัง glue ไม่ได้
  ทำได้เฉพาะเส้นตรง 2 จุด (`DrawLine` + `GlueTo`) ถ้าจะทำ ควรทำเฉพาะเส้นที่ปลายชนกรอบของเชปอื่น
- **จับคู่ไอคอนกับ stencil**: ไอคอนที่ฝังใน SVG เป็นไอคอนทางการอยู่แล้ว การแทนที่ด้วย PNG จากคลังได้คุณภาพต่ำลง
  และจับคู่จากชื่อได้แค่ 15 จาก 26 (ผิด 1: Config → AppConfig) ถ้าจะทำจริงควรสร้าง `.vssx` จากชุดไอคอนทางการ
  (กฎของผู้ใช้: **ให้หาโหลดไอคอนจริงมาใช้ ไม่วาดเอง ถ้าไม่มีค่อยวาด**) แหล่งในเครื่อง:
  `D:\Claude_app\pptmcp\Office-PowerPoint-MCP-Server\icons\` (aws 821 ไฟล์, k8s, docker, arcgis, visio)

## ข้อจำกัดที่ยังมี

- Visio เติมแบบ even-odd เสมอ ต่างจาก SVG nonzero เฉพาะกรณี subpath ซ้อนทิศทางเดียวกัน
- clip-path, mask, filter, `<pattern>`, nested `<svg>`, `<symbol>` ที่มี viewBox: ข้ามและรายงานในผลลัพธ์
- รูปที่ถูกหมุนจะวาดไม่หมุน และ `preserveAspectRatio` ของ `<image>` ยังไม่รองรับ (ยืดให้เต็มกรอบ)
- ข้อความ: `textLength`, `dx/dy` รายตัวอักษร, `<tspan>` ที่ต่างสไตล์ในบรรทัดเดียวกัน ยังไม่รองรับ
- หัวลูกศร: ขนาดเลือกจากตารางประมาณการ (`ARROW_SIZE_LIMITS`) ยังไม่ได้วัดกับ Visio จริงละเอียด
- คะแนนใน `examples/svg_compare.py` ไม่แม่น (ภาพ export ของ Visio ขนาดไม่เท่า Edge) ให้ดู `*.side.png`

## ชุดทดสอบนอก repo

ตัวอย่างที่ใช้: UML จาก Wikimedia Commons, draw.io/Inkscape จาก kubernetes/website และ prometheus, Graphviz (`@viz-js/viz`),
matplotlib อยู่ใน scratchpad ของ session (หายได้) Wikimedia จำกัดอัตรา (HTTP 429) ให้โหลดทีละน้อยและเว้นระยะ
อย่าเก็บไฟล์ภายนอกที่มีลิขสิทธิ์ใน repo ถ้าจะทำชุดถาวรให้สร้าง fixture เล็กๆ เองใน `tests/fixtures/`

## วิธีรันซ้ำ

```
python -m pytest tests --deselect tests/test_e2e_visio.py     # เฉพาะ unit (ไม่แตะ Visio)
python -m pytest tests                                         # ทั้งหมด (ต้องปิด Visio ก่อน)
python examples/svg_compare.py <out_dir> a.svg b.svg           # แปลงแล้วเทียบกับ Edge
```
