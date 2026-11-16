"""Build fda_system_on_cloud.vsdx (4 pages) by driving Visio over COM.

Usage: python build_fda_cloud.py <out_dir>
Writes fda_system_on_cloud.vsdx, fda_system_on_cloud.pdf and page PNGs into out_dir.
"""
import os
import sys

import win32com.client as w

OUT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else ".")
os.makedirs(OUT, exist_ok=True)
SRC = os.path.join(OUT, "FDA_SystemDesign_07092026.vsdx")  # the team's original (page 1)
VSDX = os.path.join(OUT, "FDA_SystemDesign_07092026_cloud.vsdx")

# All pages are A3 landscape. The original file's page is far larger (701.6 x 496.1 mm) than
# its drawing, so page 1 is resized to A3 and the drawing is moved into it.
W, H = 16.54, 11.69
K = 1.0
FONT = "Sarabun"
NAVY, ORANGE, GRAY, RED = "#1F3864", "#ED7D31", "#595959", "#C00000"
SECURE, APPZ, DMZ = "#FFF2CC", "#FCE4D6", "#DDEBF7"
AWS_BG, NIPA_BG = "#EDE7F6", "#FFF4E6"
GREEN = "#5FA846"
WHITE = "#FFFFFF"

STENCIL_DIR = r"C:\Program Files (x86)\Microsoft Office\Office15\Visio Content\1033"

app = w.Dispatch("Visio.Application")
app.Visible = False
app.AlertResponse = 7
doc = app.Documents.OpenEx(SRC, 1 | 64)  # untitled copy: the original stays untouched
stencils = {}


def stencil(name):
    if name not in stencils:
        p = os.path.join(STENCIL_DIR, name + ".VSSX")
        if not os.path.exists(p):
            p = os.path.join(STENCIL_DIR, name + ".vss")
        stencils[name] = app.Documents.OpenEx(p, 64)
    return stencils[name]


AWS_ICON_DIR = r"D:\Claude_app\pptmcp\Office-PowerPoint-MCP-Server\icons\aws"
# master name -> PNG file name (AWS Architecture Icons, as packaged by awslabs/aws-icons-for-plantuml)
AWS_ICONS = {
    "AWS Cloud": "AWSCloud", "Virtual Private Cloud": "VirtualPrivateCloud",
    "Amazon EC2": "EC2", "Amazon EBS": "ElasticBlockStore",
    "Amazon S3": "SimpleStorageService", "AWS IAM": "IdentityandAccessManagement",
    "AWS KMS": "KeyManagementService", "Amazon CloudWatch": "CloudWatch",
    "Internet Gateway": "VPCInternetGateway",
    "Network ACL": "VPCNetworkAccessControlList",
}


def build_aws_stencil():
    """Make AWS_Icons.vssx (one master per icon) by cloning a stock stencil and replacing masters."""
    base = os.path.join(STENCIL_DIR, "SERVER_M.VSSX")
    st = app.Documents.OpenEx(base, 1 | 64)  # open as an untitled copy
    for i in range(st.Masters.Count, 0, -1):
        st.Masters.Item(i).Delete()
    tmp = app.Documents.Add("")
    pg = tmp.Pages.Item(1)
    for name, png in AWS_ICONS.items():
        sh = pg.Import(os.path.join(AWS_ICON_DIR, png + ".png"))
        m = st.Drop(sh, 0, 0)
        m.Name = name
        sh.Delete()
    tmp.Close()
    st.SaveAs(os.path.join(OUT, "AWS_Icons.vssx"))
    stencils["AWS_Icons"] = st


def rgb(hexcol):
    h = hexcol.lstrip("#")
    return "RGB(%d,%d,%d)" % (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


_fid = []


def set_font_id():
    if not _fid:
        for i in range(1, doc.Fonts.Count + 1):
            f = doc.Fonts.Item(i)
            if f.Name == FONT:
                _fid.append(f.ID)
                break
        else:
            raise RuntimeError("font %s not in Visio font table" % FONT)
    return _fid[0]


class Pg:
    """Top-left based coordinate helper for one page."""

    def __init__(self, page, name):
        self.p = page
        self.p.Name = name
        s = page.PageSheet
        s.CellsU("PageWidth").FormulaU = "%s in" % W
        s.CellsU("PageHeight").FormulaU = "%s in" % H
        s.CellsU("PrintPageOrientation").FormulaU = "2"
        if page.Shapes.Count == 0:  # new page: same thin border as the original drawing
            b = page.DrawRectangle(0.1, 0.1, W - 0.1, H - 0.1)
            b.CellsU("FillPattern").FormulaU = "0"
            b.CellsU("LineWeight").FormulaU = "0.75 pt"
        self.layers = {
            "fda": page.Layers.Add("อย. (ตามรูปต้นฉบับ)"),
            "team": page.Layers.Add("ทีมเพิ่ม (สมมติ/รอยืนยัน)"),
            "cloud": page.Layers.Add("Cloud"),
        }

    def y(self, top):
        return H - top * K

    def _style(self, sh, fill=None, line=None, dash=False, weight=1.0, text_color=None,
               size=11, bold=False, halign=1, valign=1, margin=0.06, layer="fda"):
        if fill:
            sh.CellsU("FillForegnd").FormulaU = rgb(fill)
            sh.CellsU("FillPattern").FormulaU = "1"
        else:
            sh.CellsU("FillPattern").FormulaU = "0"
        if line:
            sh.CellsU("LineColor").FormulaU = rgb(line)
            sh.CellsU("LinePattern").FormulaU = "2" if dash else "1"
            sh.CellsU("LineWeight").FormulaU = "%s pt" % (weight * K)
        else:
            sh.CellsU("LinePattern").FormulaU = "0"
        sh.CellsU("Char.Font").FormulaU = str(set_font_id())
        sh.CellsU("Char.Size").FormulaU = "%s pt" % (size * K)
        sh.CellsU("Char.ComplexScriptFont").FormulaU = str(set_font_id())
        sh.CellsU("Char.ComplexScriptSize").FormulaU = "%s pt" % (size * K)
        sh.CellsU("Char.Color").FormulaU = rgb(text_color or "#000000")
        sh.CellsU("Char.Style").FormulaU = "1" if bold else "0"
        sh.CellsU("Para.HorzAlign").FormulaU = str(halign)
        sh.CellsU("VerticalAlign").FormulaU = str(valign)
        for c in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            sh.CellsU(c).FormulaU = "%s in" % (margin * K)
        self.layers[layer].Add(sh, 0)

    def box(self, x, top, w_, h_, text="", fill=None, line=None, dash=False, weight=1.0,
            text_color=None, size=11, bold=False, halign=1, valign=1, rounding=0, layer="fda",
            margin=0.06):
        sh = self.p.DrawRectangle(x * K, self.y(top + h_), (x + w_) * K, self.y(top))
        if rounding:
            sh.CellsU("Rounding").FormulaU = "%s in" % (rounding * K)
        sh.Text = text
        self._style(sh, fill, line, dash, weight, text_color, size, bold, halign, valign,
                    margin, layer)
        return sh

    def text(self, x, top, w_, h_, text, size=11, color=GRAY, bold=False, halign=0, valign=0,
             layer="fda"):
        return self.box(x, top, w_, h_, text, size=size, text_color=color, bold=bold,
                        halign=halign, valign=valign, layer=layer, margin=0.02)

    def icon(self, sten, master, cx, cy, size, layer="fda"):
        m = stencil(sten).Masters.ItemU(master)
        sh = self.p.Drop(m, cx * K, self.y(cy))
        wv, hv = sh.CellsU("Width").Result(0), sh.CellsU("Height").Result(0)
        k = size * K / max(wv, hv)
        sh.CellsU("Width").FormulaU = "%s in" % (wv * k)
        sh.CellsU("Height").FormulaU = "%s in" % (hv * k)
        sh.Text = ""
        sh.CellsU("LocPinX").FormulaU = "Width*0.5"
        sh.CellsU("LocPinY").FormulaU = "Height*0.5"
        sh.CellsU("PinX").FormulaU = "%s in" % (cx * K)
        sh.CellsU("PinY").FormulaU = "%s in" % self.y(cy)
        self.layers[layer].Add(sh, 0)
        return sh

    def conn(self, a, b, dash=False, color=NAVY, weight=1.5, label="", layer="fda"):
        c = self.p.Drop(app.ConnectorToolDataObject, 0, 0)
        c.CellsU("BeginX").GlueTo(a.CellsU("PinX"))
        c.CellsU("EndX").GlueTo(b.CellsU("PinX"))
        c.CellsU("ConLineRouteExt").FormulaU = "1"
        c.CellsU("ShapeRouteStyle").FormulaU = "2"
        c.CellsU("BeginArrow").FormulaU = "0"
        c.CellsU("EndArrow").FormulaU = "0"
        c.CellsU("LineColor").FormulaU = rgb(color)
        c.CellsU("LineWeight").FormulaU = "%s pt" % (weight * K)
        c.CellsU("LinePattern").FormulaU = "2" if dash else "1"
        if label:
            c.Text = label
            c.CellsU("Char.Font").FormulaU = str(set_font_id())
            c.CellsU("Char.Size").FormulaU = "%s pt" % (10 * K)
            c.CellsU("Char.ComplexScriptFont").FormulaU = str(set_font_id())
            c.CellsU("Char.ComplexScriptSize").FormulaU = "%s pt" % (10 * K)
            c.CellsU("Char.Color").FormulaU = rgb("#000000")
            c.CellsU("TextBkgndTrans").FormulaU = "100%"
        self.layers[layer].Add(c, 0)
        return c

    def props(self, sh, d):
        for k, v in d.items():
            sh.AddNamedRow(243, k, 0)
            sh.CellsU("Prop." + k).FormulaU = '"%s"' % v.replace('"', "'")

    # -- common page furniture ------------------------------------------------
    def furniture(self, title, subtitle, n):
        self.box(0.5, 0.35, 0.12, 0.55, "", fill=NAVY)
        self.text(0.7, 0.3, 13.0, 0.5, title, size=24, color=NAVY, bold=True, valign=1)
        self.text(0.7, 0.8, 13.5, 0.3, subtitle, size=12, color=GRAY, valign=1)
        self.text(15.0, 11.2, 1.0, 0.3, "%d / 4" % n, size=11, color=GRAY, halign=2)

    def legend(self, x, top, wd=5.6):
        self.box(x, top, wd, 1.0, "", fill=WHITE, line="#BFBFBF", weight=0.75)
        items = [("solid", "เส้นทึบ = ตามภาพของ อย."),
                 ("dash", "เส้นประ = ทีมเพิ่ม (สมมติ/รอยืนยัน)"),
                 ("tag", None)]
        for i, (kind, txt) in enumerate(items):
            yy = top + 0.17 + i * 0.3
            if kind == "tag":
                self.text(x + 0.15, yy - 0.12, 0.9, 0.26, "[Existing]", size=10, color=RED,
                          bold=True, valign=1)
                self.text(x + 1.05, yy - 0.12, 4.4, 0.26, "= อุปกรณ์ที่ อย. มีอยู่แล้ว",
                          size=10, valign=1)
            else:
                ln = self.p.DrawLine((x + 0.15) * K, self.y(yy), (x + 0.95) * K, self.y(yy))
                ln.CellsU("LineColor").FormulaU = rgb(NAVY)
                ln.CellsU("LineWeight").FormulaU = "%s pt" % (1.5 * K)
                ln.CellsU("LinePattern").FormulaU = "2" if kind == "dash" else "1"
                self.layers["fda"].Add(ln, 0)
                self.text(x + 1.05, yy - 0.12, 4.4, 0.26, txt, size=10, valign=1)


def new_page(first=False):
    return doc.Pages.Item(1) if first else doc.Pages.Add()


SPEC = {
    "web": dict(name="Web Server", os="Windows Server 2025", cpu="8 vCPU", ram="32 GB",
                disk="500 GB", sw="Web Customized Application", head="#1F3864",
                note="*FDA Provide VM*", master="Web server"),
    "gis": dict(name="GIS Server", os="Windows Server 2025", cpu="8 vCPU", ram="32 GB",
                disk="500 GB", sw="ArcGIS Enterprise Standard", head="#8FBF6A",
                note="*FDA Provide VM*", master="Application server"),
    "db": dict(name="Database Server", os="Windows Server 2025", cpu="8 vCPU", ram="32 GB",
               disk="500 GB", sw="MS SQL Server", head="#3F8F4F",
               note="*FDA Provide VM and RDBMS*", master="Database server"),
}


# =============================================================================
# Page 1 - FDA prototype
# =============================================================================
def page1():
    """Page 1 is the team's original drawing, kept as is; only small additions are made."""
    P = Pg(new_page(True), "1 ต้นแบบ อย.")
    P.p.CreateSelection(0, 0).Move(0.4, -0.6, "in")  # drawing frame -> top-left of the A3 page
    P.text(0.95, 0.4, 12.0, 0.3,
           "VM 3 เครื่อง 24 vCPU / 96 GB RAM / 1.5 TB disk ใน 3 โซน "
           "(ตัวเลขเป็นของ อย. ยังไม่ใช่ผล sizing ของทีม)", size=12, color=GRAY, valign=1)
    P.legend(0.95, 9.75)
    P.box(7.0, 9.75, 4.6, 0.7,
          "หมายเหตุ (ข้อเสนอทีม): ใช้ Windows 11 Pro หรือ Enterprise แทน Home "
          "(GIS Desktop)", fill=WHITE, line=NAVY, dash=True, size=10, halign=0, margin=0.08,
          layer="team")
    P.text(15.0, 11.2, 1.0, 0.3, "1 / 4", size=11, color=GRAY, halign=2)
    return P


# =============================================================================
# Pages 2 and 3 - cloud options
# =============================================================================
def cloud_page(first, name, title, cloud):
    aws = cloud == "aws"
    P = Pg(new_page(False), name)
    P.furniture(title, "ที่เก็บภาพเริ่มต้น 1 TB ยังไม่ระบุอัตราการเติบโต (รอยืนยัน)",
                2 if aws else 3)
    bg = AWS_BG if aws else NIPA_BG
    edge = "#7E57C2" if aws else ORANGE
    cl_title = ("AWS Cloud > Region ap-southeast-7 (Thailand)" if aws
                else "NIPA Cloud Space (ที่ตั้งศูนย์ข้อมูล (รอยืนยัน))")
    cf = P.box(0.5, 1.4, 11.7, 7.95, "", fill=bg, line=edge, dash=True, weight=2, layer="cloud")
    P.text(1.2 if aws else 0.65, 1.45, 8.0, 0.3, cl_title, size=13, color=NAVY, bold=True,
           layer="cloud")
    P.box(0.75, 1.9, 10.45, 4.55, "", line=edge, dash=True, weight=1.5, layer="cloud")
    P.text(1.4 if aws else 0.9, 1.93, 6.0, 0.3,
           "VPC (สมมติ)" if aws else "Network (VPC) (สมมติ)", size=12,
           color=NAVY, bold=True, layer="cloud")
    if aws:
        P.icon("AWS_Icons", "AWS Cloud", 0.9, 1.6, 0.4, layer="cloud")
        P.icon("AWS_Icons", "Virtual Private Cloud", 1.12, 2.08, 0.34, layer="cloud")
    zx = {"db": 1.0, "gis": 4.4, "web": 7.8}
    zone_info = {"db": ("Secure subnet" if aws else "Secure", SECURE),
                 "gis": ("Application subnet" if aws else "Application", APPZ),
                 "web": ("DMZ subnet" if aws else "DMZ", DMZ)}
    zone = {}
    cards = {}
    res = ("EC2 m7i.2xlarge + EBS\n(m7i มีใน ap-southeast-7 (รอยืนยัน))" if aws
           else "Compute apa.2xlarge.v2 + Block Storage\n(สเปก flavor (รอยืนยัน))")
    for k in ("db", "gis", "web"):
        nm, col = zone_info[k]
        zone[k] = P.box(zx[k], 2.35, 3.1, 3.95, "", fill=col, line="#BFBFBF", weight=0.75)
        P.text(zx[k] + 0.1, 2.38, 2.9, 0.3, nm, size=12, color=NAVY, bold=True)
        card = P.box(zx[k] + 0.15, 3.3, 2.8, 2.85,
                     "%s\nWindows Server 2025\n8 vCPU, 32 GB RAM, 500 GB\n%s" % (
                         SPEC[k]["name"], res),
                     fill=WHITE, line="#BFBFBF", weight=0.75, size=11, valign=2, margin=0.08)
        cards[k] = card
        if aws:
            P.icon("AWS_Icons", "Amazon EC2", zx[k] + 1.05, 3.85, 0.9, layer="cloud")
            P.icon("AWS_Icons", "Amazon EBS", zx[k] + 2.05, 3.85, 0.9, layer="cloud")
        else:
            P.icon("SERVER_M", SPEC[k]["master"], zx[k] + 1.55, 3.85, 1.0)
        P.props(card, {"OS": "Windows Server 2025", "vCPU": "8 vCPU", "RAM": "32 GB",
                       "Disk": "500 GB gp3", "Software": SPEC[k]["sw"],
                       "Status": "ข้อเสนอทีม (สมมติ/รอยืนยัน)"})
    P.conn(cards["db"], cards["gis"])
    P.conn(cards["gis"], cards["web"])
    # storage
    if aws:
        store = P.box(5.7, 6.85, 3.6, 1.0, "Amazon S3 (field-photos)\n1 TB เริ่มต้น",
                      fill=WHITE, line=edge, weight=1.25, size=12, bold=True, rounding=0.1,
                      layer="cloud")
        smalls = [("IAM", "IAM\n(สิทธิ์ราย user/app)"), ("KMS", "KMS\n(เข้ารหัสข้อมูลที่เก็บ)"),
                  ("CW", "CloudWatch\n(ตรวจสอบการใช้งาน)")]
        sx = [1.0, 4.2, 7.4]
    else:
        store = P.box(5.7, 6.85, 3.6, 1.0, "Object Storage (field-photos)\n1 TB เริ่มต้น",
                      fill=WHITE, line=edge, weight=1.25, size=12, bold=True, rounding=0.1,
                      layer="cloud")
        smalls = [("API", "S3-compatible\nAPI endpoint"), ("KEY", "Access key + policy\n(ราย app)")]
        sx = [1.0, 4.2]
    small_icons = {"IAM": "AWS IAM", "KMS": "AWS KMS", "CW": "Amazon CloudWatch"}
    if aws:
        store.CellsU("LeftMargin").FormulaU = "0.9 in"
        P.icon("AWS_Icons", "Amazon S3", 6.25, 7.35, 0.7, layer="cloud")
    for (kk, txt), x in zip(smalls, sx):
        sb = P.box(x, 8.1, 2.9, 0.95, txt, fill=WHITE, line=edge, weight=1.0, size=11,
                   rounding=0.1, layer="cloud")
        if aws:
            sb.CellsU("LeftMargin").FormulaU = "0.8 in"
            P.icon("AWS_Icons", small_icons[kk], x + 0.45, 8.575, 0.6, layer="cloud")
    P.conn(store, cards["gis"], dash=True, label="(สมมติ)", layer="team")
    P.conn(store, cards["web"], dash=True, label="(สมมติ)", layer="team")
    # edge boxes straddling cloud right edge
    e1 = ("Internet Gateway\n(สมมติ)" if aws else "External IP\n(สมมติ)")
    e2 = ("Security Group / NACL\n(สมมติ)" if aws else "Firewall / security group\n(รอยืนยัน)")
    P.box(11.5, 2.7, 1.5, 0.95, e1, fill=WHITE, line=edge, dash=True, weight=1.25, size=10,
          rounding=0.08, layer="cloud", valign=2 if aws else 1)
    P.box(11.5, 5.0, 1.5, 0.95, e2, fill=WHITE, line=edge, dash=True, weight=1.25, size=10,
          rounding=0.08, layer="cloud", valign=2 if aws else 1)
    if aws:
        P.icon("AWS_Icons", "Internet Gateway", 12.25, 3.05, 0.45, layer="cloud")
        P.icon("AWS_Icons", "Network ACL", 12.25, 5.35, 0.45, layer="cloud")
    # outside right
    inet = P.icon("NETLOC_M", "Cloud", 14.9, 4.33, 1.2)
    P.text(15.55, 4.2, 0.95, 0.3, "Internet", size=12, color="#000000", bold=True, halign=0)
    user = P.icon("PERIPH_M", "User", 14.9, 2.4, 0.7)
    P.text(13.2, 2.25, 1.3, 0.3, "Web User", size=12, color="#000000", bold=True, halign=2)
    surv = P.icon("PERIPH_M", "Smart phone", 14.9, 6.95, 0.7, layer="team")
    P.text(13.4, 7.35, 3.0, 0.5, "ทีมสำรวจภาคสนาม (มือถือ)", size=12, color="#000000",
           bold=True, halign=1, layer="team")
    P.conn(inet, user)
    P.conn(inet, zone["web"])
    P.conn(surv, inet, dash=True, label="HTTPS (สมมติ)", layer="team")
    # office box
    off = P.box(2.5, 9.75, 7.8, 1.3, "", fill=WHITE, line=GRAY, weight=1.25)
    P.text(2.6, 9.78, 7.6, 0.3, "สำนักงาน อย. [Existing]", size=12, color=RED, bold=True,
           halign=1, valign=1)
    for cx, mst, sten, lab in ((3.8, "Switch", "PERIPH_M", "Network Switch"),
                               (6.4, "Firewall", "PERIPH_M", "Firewall"),
                               (9.0, "PC", "COMPS_M", "GIS Desktop")):
        P.icon(sten, mst, cx, 10.4, 0.5)
        P.text(cx - 1.0, 10.68, 2.0, 0.26, lab, size=10, color="#000000", halign=1)
    P.conn(off, cf, dash=True, label="(รอยืนยัน)", layer="team")
    P.legend(10.5, 9.75)
    return P


# =============================================================================
# Page 2 - AWS, drawn to the AWS architecture-diagram conventions
# (group colours/borders: AWS Cloud, Region, VPC, public/private subnet, security
#  group, corporate data center; service icon with its label underneath)
# =============================================================================
AWS_DARK = "#232F3E"; AWS_REGION = "#00A4A6"; AWS_VPC = "#8C4FFF"
AWS_PUB = "#7AA116"; AWS_PUB_BG = "#F2F6E8"; AWS_PRIV = "#00A4A6"; AWS_PRIV_BG = "#E6F6F7"
AWS_SG = "#DD344C"; AWS_DC = "#5A6C86"


def aws_page():
    P = Pg(new_page(False), "2 ตัวเลือก A AWS")
    P.furniture("ตัวเลือก A: AWS (Asia Pacific (Thailand) ap-southeast-7)",
                "ผังตามรูปแบบ AWS Architecture Diagram - ขอให้ทีม AWS ยืนยันว่า EC2 m7i.2xlarge "
                "ใช้งานได้ใน ap-southeast-7", 2)
    L = "cloud"
    # groups
    cf = P.box(0.5, 1.4, 12.9, 8.0, "", line=AWS_DARK, weight=1.5, layer=L)
    P.icon("AWS_Icons", "AWS Cloud", 0.8, 1.7, 0.6, layer=L)
    P.text(1.2, 1.45, 5.0, 0.5, "AWS Cloud", size=13, color=AWS_DARK, bold=True, valign=1,
           layer=L)
    P.box(0.8, 2.0, 12.3, 7.2, "", line=AWS_REGION, dash=True, weight=1.5, layer=L)
    P.text(1.0, 2.05, 8.0, 0.35, "Region: Asia Pacific (Thailand) ap-southeast-7", size=12,
           color=AWS_REGION, bold=True, valign=1, layer=L)
    P.box(1.0, 2.6, 11.6, 3.9, "", line=AWS_VPC, weight=1.5, layer=L)
    P.icon("AWS_Icons", "Virtual Private Cloud", 1.3, 2.88, 0.45, layer=L)
    P.text(1.65, 2.65, 5.0, 0.4, "VPC (สมมติ)", size=12, color=AWS_VPC, bold=True, valign=1,
           layer=L)
    tiers = [("db", 1.2, "Private subnet - Secure (สมมติ)", AWS_PRIV, AWS_PRIV_BG),
             ("gis", 4.9, "Private subnet - Application (สมมติ)", AWS_PRIV, AWS_PRIV_BG),
             ("web", 8.6, "Public subnet - DMZ (สมมติ)", AWS_PUB, AWS_PUB_BG)]
    sg = {}
    for k, x, nm, lc, bgc in tiers:
        P.box(x, 3.2, 3.4, 3.15, "", fill=bgc, line=lc, weight=1.25, layer=L)
        P.text(x + 0.1, 3.23, 3.2, 0.3, nm, size=10, color=lc, bold=True, valign=1, layer=L)
        s = P.box(x + 0.3, 3.65, 2.8, 2.55, "", line=AWS_SG, dash=True, weight=1.25, layer=L)
        sg[k] = s
        P.text(x + 0.4, 3.68, 2.6, 0.25, "Security group (สมมติ)", size=9, color=AWS_SG,
               valign=1, layer=L)
        cx = x + 1.7
        P.icon("AWS_Icons", "Amazon EC2", cx - 0.65, 4.5, 0.75, layer=L)
        P.icon("AWS_Icons", "Amazon EBS", cx + 0.65, 4.5, 0.75, layer=L)
        P.text(cx - 1.4, 4.9, 1.5, 0.5, "%s\nm7i.2xlarge" % SPEC[k]["name"], size=9,
               color="#000000", halign=1, layer=L)
        P.text(cx - 0.0, 4.9, 1.3, 0.5, "EBS gp3\n500 GB", size=9, color="#000000",
               halign=1, layer=L)
        P.text(x + 0.4, 5.45, 2.6, 0.7,
               "Windows Server 2025\n8 vCPU / 32 GiB RAM\n%s" % SPEC[k]["sw"], size=9,
               color=GRAY, halign=1, layer=L)
        P.props(s, {"OS": "Windows Server 2025", "InstanceType": "m7i.2xlarge",
                    "vCPU": "8", "RAM": "32 GiB", "Disk": "500 GB gp3",
                    "Software": SPEC[k]["sw"],
                    "Status": "m7i ใน ap-southeast-7 (รอยืนยัน); ข้อเสนอทีม (สมมติ)"})
    P.conn(sg["db"], sg["gis"])
    P.conn(sg["gis"], sg["web"])
    # internet gateway on the VPC edge
    igw = P.icon("AWS_Icons", "Internet Gateway", 12.6, 4.5, 0.6, layer=L)
    P.box(12.15, 4.85, 0.9, 0.5, "Internet\ngateway (สมมติ)", fill=WHITE, size=9,
          text_color="#000000", halign=1, layer=L, margin=0.01)
    # regional services
    svc = [("AWS IAM", 2.3, "IAM\n(สิทธิ์ราย user/app) (สมมติ)"),
           ("AWS KMS", 5.0, "AWS KMS\n(เข้ารหัสข้อมูลที่เก็บ) (สมมติ)"),
           ("Amazon S3", 8.45, "Amazon S3 field-photos\n1 TB เริ่มต้น (รอยืนยัน อัตราเติบโต)"),
           ("Amazon CloudWatch", 11.5, "Amazon CloudWatch\n(ตรวจสอบการใช้งาน) (สมมติ)")]
    s3 = None
    for m, cx, lab in svc:
        ic = P.icon("AWS_Icons", m, cx, 7.55, 0.8, layer=L)
        P.text(cx - 1.4, 8.0, 2.8, 0.6, lab, size=9, color="#000000", halign=1, layer=L)
        if m == "Amazon S3":
            s3 = ic
    P.conn(s3, sg["gis"], dash=True, layer="team")
    P.conn(s3, sg["web"], dash=True, layer="team")
    # outside AWS
    inet = P.icon("NETLOC_M", "Cloud", 15.0, 4.5, 1.2)
    P.text(15.65, 4.35, 0.85, 0.3, "Internet", size=12, color="#000000", bold=True, halign=0)
    user = P.icon("PERIPH_M", "User", 15.0, 2.6, 0.7)
    P.text(13.5, 2.45, 1.1, 0.3, "Web User", size=12, color="#000000", bold=True, halign=2)
    surv = P.icon("PERIPH_M", "Smart phone", 15.0, 6.6, 0.7, layer="team")
    P.text(13.9, 7.0, 2.3, 0.5, "ทีมสำรวจภาคสนาม\n(มือถือ)", size=11, color="#000000",
           bold=True, halign=1, layer="team")
    P.conn(inet, user)
    P.conn(inet, igw, dash=True, layer="team")
    P.conn(igw, sg["web"], dash=True, layer="team")
    P.conn(surv, inet, dash=True, label="HTTPS (สมมติ)", layer="team")
    # corporate data center
    off = P.box(2.5, 9.75, 7.8, 1.5, "", fill=WHITE, line=AWS_DC, weight=1.25)
    P.text(2.6, 9.78, 7.6, 0.3, "Corporate data center: สำนักงาน อย. [Existing]", size=12,
           color=RED, bold=True, halign=1, valign=1)
    for cx, mst, sten, lab in ((3.8, "Switch", "PERIPH_M", "Network Switch"),
                               (6.4, "Firewall", "PERIPH_M", "Firewall"),
                               (9.0, "PC", "COMPS_M", "GIS Desktop")):
        P.icon(sten, mst, cx, 10.4, 0.5)
        P.text(cx - 1.0, 10.68, 2.0, 0.26, lab, size=10, color="#000000", halign=1)
    P.conn(off, cf, dash=True, label="(รอยืนยัน)", layer="team")
    # question for the AWS team
    P.box(13.6, 8.0, 2.75, 1.0,
          "ขอให้ทีม AWS ยืนยัน\n• m7i.2xlarge ใช้ได้ใน ap-southeast-7\n• EBS gp3 500 GB ใช้ได้ใน ap-southeast-7",
          fill=WHITE, line=ORANGE, dash=True, weight=1.25, size=10, halign=0, valign=0,
          margin=0.1, layer="team")
    P.legend(10.5, 9.75)
    return P


_frames = {}


def P_cloud_frame(P):
    # the cloud frame is the first large shape tagged on the cloud layer
    for sh in P.p.Shapes:
        try:
            if abs(sh.CellsU("Width").Result(0) - 11.7) < 0.01 and \
               abs(sh.CellsU("Height").Result(0) - 7.95) < 0.01:
                return sh
        except Exception:
            pass
    raise RuntimeError("cloud frame not found")


# =============================================================================
# Page 4 - comparison
# =============================================================================
def page4():
    P = Pg(new_page(False), "4 เปรียบเทียบ")
    P.furniture("เปรียบเทียบตัวเลือก A และ B เพื่อเลือก cloud 1 ราย",
                "ข้อมูลเปรียบเทียบเชิงเทคนิค ไม่มีข้อสรุปว่าตัวเลือกใดดีกว่า", 4)
    rows = [
        ("หัวข้อ", "ตัวเลือก A: AWS", "ตัวเลือก B: NIPA"),
        ("ที่ตั้งข้อมูล", "Thailand ap-southeast-7", "(รอยืนยัน)"),
        ("เครื่องขนาด 8 vCPU / 32 GB RAM x 3", "EC2 m7i.2xlarge (8 vCPU / 32 GiB) (ใน ap-southeast-7 (รอยืนยัน))",
         "Compute apa.2xlarge.v2 (สเปก vCPU/RAM ของ flavor (รอยืนยัน))"),
        ("ดิสก์ 500 GB ต่อเครื่อง", "EBS gp3 500 GB",
         "Block Storage (Standard-HDD / Standard-SSD / Premium-SSD (รอยืนยัน))"),
        ("การแยกโซน Secure / Application / DMZ", "VPC + subnet + Security Group / NACL (สมมติ)",
         "Network แยกโซน (สมมติ) + firewall / security group (รอยืนยัน)"),
        ("ที่เก็บภาพ 1 TB เริ่มต้น", "Amazon S3 (S3 API ต้นฉบับ)",
         "Object Storage (S3-compatible API)"),
        ("การเข้ารหัสข้อมูลที่เก็บ", "KMS", "(รอยืนยัน)"),
        ("การตรวจสอบการใช้งาน", "CloudWatch", "(รอยืนยัน)"),
    ]
    hts = [0.45, 0.5, 0.8, 0.7, 0.7, 0.5, 0.5, 0.5]
    cw = [3.9, 5.7, 5.7]
    top = 1.35
    for r, row in enumerate(rows):
        x = 0.6
        for c, txt in enumerate(row):
            hdr = r == 0
            P.box(x, top, cw[c], hts[r], txt,
                  fill=NAVY if hdr else ("#F2F2F2" if c == 0 else WHITE),
                  line="#BFBFBF", weight=0.75, text_color=WHITE if hdr else "#000000",
                  size=12 if hdr else 11, bold=hdr or c == 0, halign=0 if c else 0,
                  margin=0.1)
            x += cw[c]
        top += hts[r]
    # (b) info boxes
    t0 = top + 0.35
    bh = 2.0
    boxes = [
        ("สอบถาม AWS", ["m7i.2xlarge และ EBS gp3 มีใน ap-southeast-7 หรือไม่",
                         "รูปแบบ license ของ Windows Server 2025 บน EC2"]),
        ("สอบถาม NIPA", ["ที่ตั้งศูนย์ข้อมูล", "apa.2xlarge.v2 ได้ 8 vCPU / 32 GB หรือไม่",
                          "ชนิด Block Storage", "firewall / security group",
                          "การเข้ารหัสและการตรวจสอบการใช้งาน",
                          "รูปแบบ license ของ Windows Server 2025"]),
        ("ยืนยันภายใน อย./ทีม", ["ข้อกำหนดที่ตั้งข้อมูลของ อย.",
                                   "รูปแบบเชื่อมต่อสำนักงาน อย. เข้า cloud และบทบาทของ "
                                   "Firewall / Network Switch [Existing]",
                                   "license ของ SQL Server และ ArcGIS Enterprise บน cloud",
                                   "อัตราการเติบโตของข้อมูลภาพและขนาดที่ต้องรองรับในอนาคต",
                                   "OS ของ GIS Desktop: Home (อย. ระบุ) หรือ Pro/Enterprise "
                                   "(ข้อเสนอทีม)"]),
    ]
    for i, (hd, items) in enumerate(boxes):
        x = 0.6 + i * 5.2
        P.box(x, t0, 4.9, 0.4, hd, fill=ORANGE, text_color=WHITE, size=12, bold=True)
        P.box(x, t0 + 0.4, 4.9, bh, "\n".join("• " + s for s in items), fill=WHITE,
              line="#BFBFBF", weight=0.75, size=11, halign=0, valign=0, margin=0.12)
    # (c) decision
    t1 = t0 + 0.4 + bh + 0.3
    P.box(0.6, t1, 9.4, 1.05, "", fill=WHITE, line=NAVY, weight=1.5)
    P.text(0.75, t1 + 0.08, 1.6, 0.4, "ผลการเลือก:", size=13, color=NAVY, bold=True, valign=1)
    for j, lab in enumerate(("ตัวเลือก A: AWS", "ตัวเลือก B: NIPA")):
        bx = 2.5 + j * 3.7
        P.box(bx, t1 + 0.12, 0.3, 0.3, "", fill=WHITE, line="#000000", weight=1.25)
        P.text(bx + 0.4, t1 + 0.08, 2.9, 0.4, lab, size=12, color="#000000", valign=1)
    P.text(0.75, t1 + 0.62, 8.5, 0.3, "ค่าใช้จ่ายประเมินแยกต่างหาก", size=10, color=GRAY,
           valign=1)
    P.legend(10.3, t1 - 0.05)
    return P


build_aws_stencil()
page1()
aws_page()
cloud_page(False, "3 ตัวเลือก B NIPA", "ตัวเลือก B: NIPA Cloud", "nipa")
page4()

doc.SaveAs(VSDX)
doc.ExportAsFixedFormat(1, os.path.join(OUT, "FDA_SystemDesign_07092026_cloud.pdf"), 1, 0)
doc.Close()
for s in stencils.values():
    try:
        s.Close()
    except Exception:
        pass
app.Quit()
print("saved", VSDX)
