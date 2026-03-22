"""
template.py — Resume PDF template
Contains all drawing primitives, layout logic, and zone registration.
Called by fire.py which acts as a thin wrapper.
"""

import json
import base64
import tempfile
import os
from pathlib import Path
from io import BytesIO

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors


def _hex(h):
    return colors.HexColor(h)


def _wrap(cv, text, fn, fs, max_w):
    words = str(text).split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if cv.stringWidth(trial, fn, fs) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def build_resume(data, main_color="#1a56db", secondary_color="#6b7280", font="Helvetica"):
    """
    Render resume data into a PDF.
    Returns (pdf_bytes, zones, page_width_pts, page_height_pts, total_pages)

    zones = list of {id, label, value, path, x, y, w, h, fs, page}
    All coordinates in PDF points, origin bottom-left.
    """
    if isinstance(data, (str, Path)):
        with open(data, "r", encoding="utf-8") as f:
            data = json.load(f)

    F  = font
    FB = font + "-Bold"
    FI = font + ("-Italic" if "Times" in font else "-Oblique")

    PW, PH = A4
    ML = 52       # left margin
    MR = 52       # right margin
    TM = 50       # top margin
    BM = 36       # bottom margin
    CW = PW - ML - MR  # content width

    C_ACC   = _hex(main_color)
    C_META  = _hex(secondary_color)
    C_BLACK = _hex("#111111")
    C_BODY  = _hex("#222222")
    C_RULE  = _hex("#e5e7eb")
    C_WHITE = colors.white

    buf = BytesIO()
    cv  = canvas.Canvas(buf, pagesize=A4)
    cv.setTitle(f"{data.get('name', 'Resume')} — Resume")
    cv.setFillColor(C_WHITE)
    cv.rect(0, 0, PW, PH, fill=1, stroke=0)

    y            = PH - TM
    zones        = []
    current_page = [1]
    _images      = data.get("others", {}).get("images", [])

    # ─── Images ──────────────────────────────────────────────────────────────

    def draw_images_for_page(page_num):
        for img in _images:
            try:
                if int(img.get("page", 1)) != page_num:
                    continue
                b64 = img.get("src_b64", "")
                if not b64:
                    continue
                raw    = base64.b64decode(b64)
                suffix = ".png" if len(raw) > 3 and raw[:4] == b"\x89PNG" else ".jpg"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tf:
                    tf.write(raw)
                    tmp = tf.name
                cv.drawImage(
                    tmp,
                    float(img.get("x", 52)),
                    float(img.get("y", 400)),
                    width=float(img.get("w", 80)),
                    height=float(img.get("h", 80)),
                    preserveAspectRatio=False,
                    mask="auto",
                )
                os.unlink(tmp)
            except Exception:
                pass

    # ─── Zone registration ───────────────────────────────────────────────────

    def zone(zid, label, value, path, x, yy, w, h, font_size=10):
        # yy   = bottom baseline of text block
        # h    = span between first and last baseline (0 for single line)
        # Expand by Helvetica ascent/descent so highlight covers full glyph
        asc = font_size * 0.718
        dsc = font_size * 0.207
        zones.append({
            "id":    zid,
            "label": label,
            "value": str(value) if value is not None else "",
            "path":  path,
            "x":     x,
            "y":     yy - dsc,
            "w":     w,
            "h":     h + asc + dsc,
            "fs":    font_size,
            "page":  current_page[0],
        })

    # ─── Page management ─────────────────────────────────────────────────────

    def new_page():
        nonlocal y
        draw_images_for_page(current_page[0])
        cv.showPage()
        current_page[0] += 1
        cv.setFillColor(C_WHITE)
        cv.rect(0, 0, PW, PH, fill=1, stroke=0)
        y = PH - TM

    def check_page(needed=30):
        if y - needed < BM + 8:
            new_page()

    # ─── Drawing helpers ─────────────────────────────────────────────────────

    def section(label):
        nonlocal y
        check_page(25)
        cv.setFont(FB, 8)
        cv.setFillColor(C_ACC)
        cv.drawString(ML, y, label.upper())
        y -= 3
        cv.setStrokeColor(C_RULE); cv.setLineWidth(0.5)
        cv.line(ML, y, ML+CW, y)
        lw = cv.stringWidth(label.upper(), FB, 8) + 4
        cv.setStrokeColor(C_ACC); cv.setLineWidth(0.8)
        cv.line(ML, y, ML+lw, y)
        y -= 10

    def bullet_item(text, zid, path, indent=12, lh=11.5):
        nonlocal y
        lines = _wrap(cv, text, F, 8.5, CW-indent-6)
        check_page(len(lines) * lh + 4)
        cv.setFillColor(C_ACC)
        cv.circle(ML+indent-5, y+2.8, 1.4, fill=1, stroke=0)
        cv.setFont(F, 8.5); cv.setFillColor(C_BODY)
        y_start = y
        for line in lines:
            cv.drawString(ML+indent, y, line)
            y -= lh
        last_bl = y_start - (len(lines)-1)*lh
        zone(zid, "Bullet", text, path, ML+indent,
             last_bl, CW-indent, (len(lines)-1)*lh, 8.5)

    def two_col(left, right, lf=None, ls=10, rf=None, rs=8, lc=None, rc=None):
        lf = lf or FB; rf = rf or F
        lc = lc or C_BLACK; rc = rc or C_META
        cv.setFont(lf, ls); cv.setFillColor(lc)
        cv.drawString(ML, y, left)
        rw = cv.stringWidth(right, rf, rs)
        cv.setFont(rf, rs); cv.setFillColor(rc)
        cv.drawString(ML+CW-rw, y, right)

    def subline(text, lh=11):
        nonlocal y
        cv.setFont(FI, 8.5); cv.setFillColor(C_META)
        cv.drawString(ML, y, text)
        y -= lh

    def spacer(h=7):
        nonlocal y
        y -= h

    # ═════════════════════════════════════════════════════════════════════════
    # HEADER
    # ═════════════════════════════════════════════════════════════════════════
    name  = data.get("name",  "")
    title = data.get("title", "")
    c_    = data.get("contact", {})

    # Name — large bold
    cv.setFont(FB, 26); cv.setFillColor(C_BLACK)
    cv.drawString(ML, y, name)
    zone("name", "Full Name", name, "name",
         ML, y, min(cv.stringWidth(name, FB, 26)+4, CW), 0, 26)
    y -= 16

    # Job title — no seniority badge
    if title:
        cv.setFont(F, 10); cv.setFillColor(C_ACC)
        cv.drawString(ML, y, title)
        zone("title", "Job Title", title, "title",
             ML, y, cv.stringWidth(title, F, 10)+4, 0, 10)
        y -= 13

    # Contact row
    contact_keys = [
        ("email",    "Email"),
        ("phone",    "Phone"),
        ("location", "Location"),
        ("github",   "GitHub"),
        ("website",  "Website"),
        ("linkedin", "LinkedIn"),
    ]
    parts = [(k, lbl, c_.get(k, "")) for k, lbl in contact_keys if c_.get(k, "")]
    sep   = "   ·   "
    sep_w = cv.stringWidth(sep, F, 7.8)
    cv.setFont(F, 7.8); cv.setFillColor(C_META)
    cx = ML
    for i, (key, lbl, val) in enumerate(parts):
        vw = cv.stringWidth(val, F, 7.8)
        cv.drawString(cx, y, val)
        zone(f"contact.{key}", lbl, val, f"contact.{key}", cx, y, vw, 0, 7.8)
        cx += vw
        if i < len(parts) - 1:
            cv.drawString(cx, y, sep)
            cx += sep_w
    y -= 14

    # Accent rule under header
    cv.setStrokeColor(C_ACC); cv.setLineWidth(1.8)
    cv.line(ML, y, ML+CW, y)
    y -= 14

    # ═════════════════════════════════════════════════════════════════════════
    # SUMMARY
    # ═════════════════════════════════════════════════════════════════════════
    summary = data.get("summary", "").strip()
    if summary:
        section("Summary")
        lines   = _wrap(cv, summary, FI, 9, CW)
        lh      = 12
        cv.setFont(FI, 9); cv.setFillColor(C_BODY)
        y_start = y
        for line in lines:
            cv.drawString(ML, y, line)
            y -= lh
        last_bl = y_start - (len(lines)-1)*lh
        zone("summary", "Summary", summary, "summary",
             ML, last_bl, CW, (len(lines)-1)*lh, 9)
        spacer(4)

    # ═════════════════════════════════════════════════════════════════════════
    # EXPERIENCE  — duration intentionally omitted
    # ═════════════════════════════════════════════════════════════════════════
    experience = data.get("experience", [])
    if experience:
        section("Experience")
        for ei, exp in enumerate(experience):
            role    = exp.get("role",        "")
            company = exp.get("company",     "")
            loc     = exp.get("location",    "")
            start   = exp.get("start_date",  "")
            end     = exp.get("end_date",    "Present")
            descs   = exp.get("description", [])
            # Show "Mar 2025 – Jun 2025", omit "(3 months)"
            date_str = f"{start} – {end}" if start else end
            comp_str = company + (f"  ·  {loc}" if loc else "")
            check_page(38)
            two_col(role, date_str)
            zone(f"exp.{ei}.role", "Role", role,
                 f"experience.{ei}.role",
                 ML, y, cv.stringWidth(role, FB, 10)+4, 0, 10)
            zone(f"exp.{ei}.dates", "Dates (start – end)", date_str,
                 f"experience.{ei}.__dates",
                 ML+CW-cv.stringWidth(date_str, F, 8)-2, y,
                 cv.stringWidth(date_str, F, 8)+4, 0, 8)
            y -= 13
            subline(comp_str)
            zone(f"exp.{ei}.company", "Company · Location", comp_str,
                 f"experience.{ei}.__comp", ML, y, CW//2, 0, 8.5)
            for di, d in enumerate(descs):
                bullet_item(d, f"exp.{ei}.desc.{di}",
                            f"experience.{ei}.description.{di}")
            spacer(7)

    # ═════════════════════════════════════════════════════════════════════════
    # PROJECTS
    # ═════════════════════════════════════════════════════════════════════════
    projects = data.get("projects", [])
    if projects:
        section("Projects")
        for pi, proj in enumerate(projects):
            pname = proj.get("name",  "")
            links = proj.get("links", "")
            descs = proj.get("description", [])
            check_page(25)
            cv.setFont(FB, 9.5); cv.setFillColor(C_BLACK)
            cv.drawString(ML, y, pname)
            zone(f"proj.{pi}.name", "Project Name", pname,
                 f"projects.{pi}.name",
                 ML, y, cv.stringWidth(pname, FB, 9.5)+4, 0, 9.5)
            if links:
                lw2 = cv.stringWidth(links, F, 8)
                cv.setFont(F, 8); cv.setFillColor(C_META)
                cv.drawString(ML+CW-lw2, y, links)
                zone(f"proj.{pi}.links", "Link", links,
                     f"projects.{pi}.links",
                     ML+CW-lw2-2, y, lw2+4, 0, 8)
            y -= 13
            for di, d in enumerate(descs):
                bullet_item(d, f"proj.{pi}.desc.{di}",
                            f"projects.{pi}.description.{di}")
            spacer(6)

    # ═════════════════════════════════════════════════════════════════════════
    # SKILLS
    # ═════════════════════════════════════════════════════════════════════════
    hard = data.get("hardskills", {})
    soft = data.get("softskills", [])
    if hard or soft:
        section("Skills")
        items_dict = (hard if isinstance(hard, dict)
                      else ({"Technical": hard} if hard else {}))
        for cat, vals in items_dict.items():
            val_str   = ", ".join(vals)
            cat_label = cat + ":"
            cat_w     = cv.stringWidth(cat_label+"  ", FB, 8.5)
            cv.setFont(FB, 8.5); cv.setFillColor(C_BLACK)
            cv.drawString(ML, y, cat_label)
            val_lines = _wrap(cv, val_str, F, 8.5, CW-cat_w)
            cv.setFont(F, 8.5); cv.setFillColor(C_BODY)
            y_sk = y
            for vline in val_lines:
                cv.drawString(ML+cat_w, y, vline)
                y -= 12
            last_bl = y_sk - (len(val_lines)-1)*12
            zone(f"skill.{cat}", f"Skills: {cat}", val_str,
                 f"hardskills.{cat}.__csv",
                 ML+cat_w, last_bl, CW-cat_w, (len(val_lines)-1)*12, 8.5)
        if soft:
            sft_label = "Soft:"
            sft_w     = cv.stringWidth(sft_label+"  ", FB, 8.5)
            cv.setFont(FB, 8.5); cv.setFillColor(C_BLACK)
            cv.drawString(ML, y, sft_label)
            soft_str = ", ".join(soft)
            cv.setFont(F, 8.5); cv.setFillColor(C_BODY)
            cv.drawString(ML+sft_w, y, soft_str)
            zone("softskills", "Soft Skills (comma separated)", soft_str,
                 "softskills.__csv", ML+sft_w, y, CW-sft_w, 0, 8.5)
            y -= 12
        spacer(4)

    # ═════════════════════════════════════════════════════════════════════════
    # EDUCATION
    # ═════════════════════════════════════════════════════════════════════════
    education = data.get("education", [])
    if education:
        section("Education")
        for ei, edu in enumerate(education):
            deg      = edu.get("degree",      "")
            field    = edu.get("field",       "")
            inst     = edu.get("institution", "")
            start    = edu.get("start_date",  "")
            end      = edu.get("end_date",    "")
            state    = edu.get("state",       "")
            gpa      = edu.get("gpa",         "")
            deg_str  = deg + (f" — {field}" if field else "")
            date_str = (f"{start} – " if start else "") + (end or "")
            if state:
                date_str += f"  [{state}]"
            check_page(33)
            two_col(deg_str, date_str)
            zone(f"edu.{ei}.deg", "Degree & Field", deg_str,
                 f"education.{ei}.__deg",
                 ML, y, cv.stringWidth(deg_str, FB, 10)+4, 0, 10)
            zone(f"edu.{ei}.date", "Dates", date_str,
                 f"education.{ei}.__date",
                 ML+CW-cv.stringWidth(date_str, F, 8)-2, y,
                 cv.stringWidth(date_str, F, 8)+4, 0, 8)
            y -= 13
            meta = "  ·  ".join(
                p for p in [inst, f"GPA: {gpa}" if gpa else ""] if p)
            subline(meta)
            zone(f"edu.{ei}.inst", "Institution", meta,
                 f"education.{ei}.__inst", ML, y, CW//2, 0, 8.5)
            spacer(5)

    # ═════════════════════════════════════════════════════════════════════════
    # CERTIFICATIONS
    # ═════════════════════════════════════════════════════════════════════════
    certs = data.get("certifications", [])
    if certs:
        section("Certifications")
        for ci, cert in enumerate(certs):
            bullet_item(str(cert), f"cert.{ci}", f"certifications.{ci}")
        spacer(4)

    # ═════════════════════════════════════════════════════════════════════════
    # LANGUAGES
    # ═════════════════════════════════════════════════════════════════════════
    langs = data.get("languages", [])
    if langs:
        section("Languages")
        lang_str = "  ·  ".join(langs)
        cv.setFont(F, 8.5); cv.setFillColor(C_BODY)
        cv.drawString(ML, y, lang_str)
        zone("languages", "Languages (comma separated)",
             ", ".join(langs), "languages.__csv", ML, y, CW, 0, 8.5)
        y -= 12
        spacer(3)

    # ═════════════════════════════════════════════════════════════════════════
    # AWARDS / PUBLICATIONS
    # ═════════════════════════════════════════════════════════════════════════
    for key, label in [("awards", "Awards"), ("publications", "Publications")]:
        items = data.get(key, [])
        if items:
            section(label)
            for ii, item in enumerate(items):
                bullet_item(str(item), f"{key}.{ii}", f"{key}.{ii}")
            spacer(4)

    # ═════════════════════════════════════════════════════════════════════════
    # FINALISE — draw images on last page; no footer text
    # ═════════════════════════════════════════════════════════════════════════
    draw_images_for_page(current_page[0])

    cv.save()
    buf.seek(0)
    return buf.read(), zones, PW, PH, current_page[0]