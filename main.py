from email.message import EmailMessage
from pathlib import Path
import os
import smtplib
import sys
import threading
import time
import urllib.error
import urllib.request


def load_env(env_path: Path = Path(__file__).with_name(".env")) -> None:
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        os.environ.setdefault(key, val)


load_env()

SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
ALERT_TO_EMAIL = os.getenv("ALERT_TO_EMAIL", SENDER_EMAIL)
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))

R = "\033[31;1m"
G = "\033[32;1m"
Y = "\033[33;1m"
C = "\033[36;1m"
W = "\033[37;1m"
RESET = "\033[0m"

INFO = f"{C}[{W}i{C}]{RESET}"
SELECT = f"{G}[{W}?{G}]{RESET}"
SUCCESS = f"{G}[{W}✓{G}]{RESET}"
WARN = f"{Y}[{W}!{Y}]{RESET}"
ERR = f"{R}[{W}x{R}]{RESET}"

BANNER = rf"""{R}
/\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\
████████╗██████╗ ██╗██████╗ ██╗    ██╗██╗██████╗ ███████╗
╚══██╔══╝██╔══██╗██║██╔══██╗██║    ██║██║██╔══██╗██╔════╝
   ██║   ██████╔╝██║██████╔╝██║ █╗ ██║██║██████╔╝█████╗
   ██║   ██╔══██╗██║██╔═══╝ ██║███╗██║██║██╔══██╗██╔══╝
   ██║   ██║  ██║██║██║     ╚███╔███╔╝██║██║  ██║███████╗
   ╚═╝   ╚═╝  ╚═╝╚═╝╚═╝      ╚══╝╚══╝ ╚═╝╚═╝  ╚═╝╚══════╝{RESET}
{C}------------------------ [ TRIPWIRE ] -------------------------{RESET}"""


def normalize_url(url: str) -> str:
    url = url.strip()
    if url and not url.startswith(("http://", "https://")):
        return f"https://{url}"
    return url


def send_tripwire_email(
    message: str,
    url: str,
    output_file: str,
    log_line: str,
) -> None:
    if not (SENDER_EMAIL and APP_PASSWORD and ALERT_TO_EMAIL):
        print(f"\n{WARN} Missing SENDER_EMAIL, APP_PASSWORD, or ALERT_TO_EMAIL in .env")
        return

    body = (
        f"Tripwire triggered!\n"
        f"File: {output_file}\n"
        f"Target URL: {url}\n"
        f"Message: {message or '(none)'}\n"
        f"Access Log Entry: {log_line}\n"
    )

    msg = EmailMessage()
    msg["Subject"] = f"[TRIPWIRE HIT] {Path(output_file).name}"
    msg["From"] = SENDER_EMAIL
    msg["To"] = ALERT_TO_EMAIL
    msg.set_content(body)

    try:
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=10) as server:
                server.login(SENDER_EMAIL, APP_PASSWORD)
                server.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(SENDER_EMAIL, APP_PASSWORD)
                server.send_message(msg)
        print(f"\n{SUCCESS} Alert email sent to {W}{ALERT_TO_EMAIL}{RESET}")
    except Exception as exc:
        print(f"\n{ERR} Failed to send email to {W}{ALERT_TO_EMAIL}{RESET}: {exc}")


def pptx(url: str, out_path: str = "sample_presentation.pptx") -> str:
    from pptx import Presentation
    from pptx.opc.constants import RELATIONSHIP_TYPE
    from pptx.oxml import parse_xml

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    r_id = slide.part.relate_to(url, RELATIONSHIP_TYPE.IMAGE, is_external=True)

    pic_xml = f"""
    <p:pic
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvPicPr>
        <p:cNvPr id="2" name="Beacon"/>
        <p:cNvPicPr/>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip r:link="{r_id}"/>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
    """
    slide.shapes._spTree.append(parse_xml(pic_xml))

    prs.save(out_path)
    return out_path


def docx(url: str, out_path: str = "sample_document.docx") -> str:
    from docx import Document
    from docx.opc.constants import RELATIONSHIP_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()
    r_id = doc.part.relate_to(url, RELATIONSHIP_TYPE.IMAGE, is_external=True)

    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:line"), "1")
    spacing.set(qn("w:lineRule"), "exact")
    p_pr.append(spacing)

    run = p.add_run()
    drawing = OxmlElement("w:drawing")
    inline = OxmlElement("wp:inline")
    extent = OxmlElement("wp:extent")
    extent.set("cx", "0")
    extent.set("cy", "0")
    doc_pr = OxmlElement("wp:docPr")
    doc_pr.set("id", "1")
    doc_pr.set("name", "Beacon")

    graphic = OxmlElement("a:graphic")
    graphic.set(qn("xmlns:a"), "http://schemas.openxmlformats.org/drawingml/2006/main")
    graphic_data = OxmlElement("a:graphicData")
    graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")

    pic = OxmlElement("pic:pic")
    pic.set(qn("xmlns:pic"), "http://schemas.openxmlformats.org/drawingml/2006/picture")
    nv_pic_pr = OxmlElement("pic:nvPicPr")
    c_nv_pr = OxmlElement("pic:cNvPr")
    c_nv_pr.set("id", "0")
    c_nv_pr.set("name", "beacon.png")
    c_nv_pic_pr = OxmlElement("pic:cNvPicPr")
    nv_pic_pr.append(c_nv_pr)
    nv_pic_pr.append(c_nv_pic_pr)

    blip_fill = OxmlElement("pic:blipFill")
    blip = OxmlElement("a:blip")
    blip.set(qn("r:link"), r_id)
    blip_fill.append(blip)

    sp_pr = OxmlElement("pic:spPr")
    xfrm = OxmlElement("a:xfrm")
    off = OxmlElement("a:off")
    off.set("x", "0")
    off.set("y", "0")
    ext = OxmlElement("a:ext")
    ext.set("cx", "0")
    ext.set("cy", "0")
    xfrm.append(off)
    xfrm.append(ext)
    prst_geom = OxmlElement("a:prstGeom")
    prst_geom.set("prst", "rect")
    sp_pr.append(xfrm)
    sp_pr.append(prst_geom)

    pic.append(nv_pic_pr)
    pic.append(blip_fill)
    pic.append(sp_pr)
    graphic_data.append(pic)
    graphic.append(graphic_data)

    inline.append(extent)
    inline.append(doc_pr)
    inline.append(graphic)
    drawing.append(inline)
    run._r.append(drawing)

    doc.save(out_path)
    return out_path


def img(url: str, out_path: str = "sample_image.png") -> str:
    from PIL import Image
    from PIL.PngImagePlugin import PngInfo

    canvas = Image.new("RGBA", (1, 1), color=(0, 0, 0, 0))
    meta = PngInfo()
    meta.add_text("Tripwire-URL", url)
    canvas.save(out_path, pnginfo=meta)
    return out_path


def watch_http_log(
    server_ip: str,
    port: int,
    message: str,
    url: str,
    output_file: str,
    stop_event: threading.Event,
    poll_sec: float = 1.0,
) -> None:
    log_url = f"http://{server_ip}:{port}/access.log"
    seen_lines = 0
    last_err = ""

    try:
        with urllib.request.urlopen(log_url, timeout=5) as resp:
            seen_lines = len(resp.read().decode("utf-8", errors="replace").splitlines())
        print(f"{INFO} Connected to {W}{log_url}{RESET} (baseline lines: {W}{seen_lines}{RESET})")
    except Exception as exc:
        print(f"{ERR} Initial fetch from {W}{log_url}{RESET} failed: {exc}")

    while not stop_event.is_set():
        time.sleep(poll_sec)
        try:
            with urllib.request.urlopen(log_url, timeout=5) as resp:
                lines = resp.read().decode("utf-8", errors="replace").splitlines()
            last_err = ""

            if len(lines) < seen_lines:
                seen_lines = 0

            for line in lines[seen_lines:]:
                clean_line = line.strip()
                if clean_line:
                    print(f"\n{SUCCESS} {W}ACCESS LOGGED{RESET} [{C}{server_ip}:{port}{RESET}]: {clean_line}")
                    send_tripwire_email(message, url, output_file, clean_line)

            seen_lines = len(lines)
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            err_msg = str(exc)
            if err_msg != last_err:
                print(f"\n{WARN} Poll error ({log_url}): {err_msg}")
                last_err = err_msg
            continue


def main() -> None:
    print(BANNER)
    print(f"\n  {Y}[1]{RESET} pptx")
    print(f"  {Y}[2]{RESET} docx")
    print(f"  {Y}[3]{RESET} image\n")

    choice = input(f"{SELECT} Select format {G}>>{RESET} ").strip().lower()
    format_map = {
        "1": "pptx",
        "pptx": "pptx",
        "2": "docx",
        "docx": "docx",
        "3": "image",
        "img": "image",
        "image": "image",
    }
    fmt = format_map.get(choice)
    if not fmt:
        print(f"{ERR} Invalid format choice.")
        sys.exit(1)

    url = normalize_url(input(f"{SELECT} Tunnel link {G}>>{RESET} "))
    if not url:
        print(f"{ERR} Tunnel link is required.")
        sys.exit(1)

    server_ip = input(
        f"{SELECT} Server IP {W}{RESET} {G}>>{RESET} "
    ).strip() or "10.22.39.204"

    port_raw = input(f"{SELECT} Log HTTP port {W}{RESET} {G}>>{RESET} ").strip()
    port = int(port_raw) if port_raw else 8080

    message = input(f"{SELECT} Alert message {G}>>{RESET} ").strip()

    if fmt == "pptx":
        saved_file = pptx(url)
    elif fmt == "docx":
        saved_file = docx(url)
    else:
        saved_file = img(url)

    print(f"\n{SUCCESS} Saved {W}{saved_file}{RESET}")
    print(f"{INFO} Polling {W}http://{server_ip}:{port}/access.log{RESET}... Press {Y}Ctrl+C{RESET} to stop.\n")

    stop_event = threading.Event()
    watcher = threading.Thread(
        target=watch_http_log,
        args=(server_ip, port, message, url, saved_file, stop_event),
        daemon=True,
    )
    watcher.start()

    try:
        while watcher.is_alive():
            time.sleep(0.5)
    except KeyboardInterrupt:
        stop_event.set()
        print(f"\n{WARN} Stopped.")


if __name__ == "__main__":
    main()