from email.message import EmailMessage
import os
from pathlib import Path
import smtplib
import sys
import time

# MaxPhisher-style ANSI palette
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


def send_tripwire_email(to_email: str, message: str, url: str, output_file: str) -> None:
    """Sends alert email if SMTP env vars are configured; otherwise prints a dry-run notice."""
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASS")
    from_email = os.getenv("SMTP_FROM", smtp_user or "")

    body = (
        f"Tripwire generated:\n"
        f"File: {output_file}\n"
        f"Target URL: {url}\n"
        f"Message: {message or '(none)'}\n"
    )

    if not (smtp_host and smtp_user and smtp_pass and from_email):
        print(f"{WARN} SMTP env vars not set (SMTP_HOST, SMTP_USER, SMTP_PASS); skipping live send.")
        print(f"{INFO} Queued alert for {W}{to_email}{RESET}: {message or '(no message)'}")
        return

    msg = EmailMessage()
    msg["Subject"] = f"[TRIPWIRE] Generated {Path(output_file).name}"
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(body)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=10) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

    print(f"{SUCCESS} Alert email sent to {W}{to_email}{RESET}")


def pptx(url: str, out_path: str = "sample_presentation.pptx") -> str:
    from pptx import Presentation

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = "Tripwire Resource"
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.clear()

    p = tf.paragraphs[0]
    p.text = "Open link to access resource: "
    run = p.add_run()
    run.text = url
    run.hyperlink.address = url

    prs.save(out_path)
    return out_path


def docx(url: str, out_path: str = "sample_document.docx") -> str:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE

    doc = Document()
    doc.add_heading("Tripwire Resource", level=1)

    p = doc.add_paragraph("Access link: ")
    r_id = doc.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    new_run.append(r_pr)

    text_elem = OxmlElement("w:t")
    text_elem.text = url
    new_run.append(text_elem)
    hyperlink.append(new_run)
    p._p.append(hyperlink)

    doc.save(out_path)
    return out_path


def img(url: str, out_path: str = "sample_image.png") -> str:
    from PIL import Image, ImageDraw, ImageFont
    from PIL.PngImagePlugin import PngInfo

    canvas = Image.new("RGB", (640, 140), color=(25, 28, 36))
    draw = ImageDraw.Draw(canvas)
    font = ImageFont.load_default()

    draw.text((20, 25), "TRIPWIRE ENDPOINT", fill=(255, 80, 80), font=font)
    draw.text((20, 60), f"URL: {url}", fill=(240, 240, 240), font=font)

    meta = PngInfo()
    meta.add_text("Tripwire-URL", url)

    canvas.save(out_path, pnginfo=meta)
    return out_path


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

    raw_url = input(f"{SELECT} Tunnel link {G}>>{RESET} ").strip()
    url = normalize_url(raw_url)
    if not url:
        print(f"{ERR} Tunnel link is required.")
        sys.exit(1)

    email = input(f"{SELECT} Alert email (optional) {G}>>{RESET} ").strip()
    message = input(f"{SELECT} Alert message (optional) {G}>>{RESET} ").strip()

    print(f"\n{INFO} Embedding {W}{url}{RESET} into {W}{fmt}{RESET}...")
    time.sleep(0.2)

    if fmt == "pptx":
        saved_file = pptx(url)
    elif fmt == "docx":
        saved_file = docx(url)
    else:
        saved_file = img(url)

    print(f"{SUCCESS} Saved {W}{saved_file}{RESET}")

    if email:
        send_tripwire_email(email, message, url, saved_file)
    elif message:
        print(f"{INFO} Message noted (no email provided): {W}{message}{RESET}")


if __name__ == "__main__":
    main()