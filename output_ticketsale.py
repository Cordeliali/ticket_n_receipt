#!/usr/bin/env python3
import csv
import glob
import math
import re


TARGET_TICKET_ITEM = "[Ticket] 3/8 KP! 3rd Oneman Live"
OUTPUT_PDF = "ticketsale_summary.pdf"

PAGE_WIDTH = 612.0   # US Letter width (8.5in * 72)
PAGE_HEIGHT = 792.0  # US Letter height (11in * 72)
MARGIN = 36.0

TABLE_LEFT = MARGIN
TABLE_RIGHT = PAGE_WIDTH - MARGIN
TABLE_TOP = PAGE_HEIGHT - MARGIN
TABLE_BOTTOM = MARGIN
TABLE_WIDTH = TABLE_RIGHT - TABLE_LEFT
TABLE_HEIGHT = TABLE_TOP - TABLE_BOTTOM

HEADER_HEIGHT = 24.0
ROW_HEIGHT = 22.0

# Preferred Name, Email, Ticket Type, Check
COL_WIDTHS = [0.27, 0.43, 0.22, 0.08]


def pdf_escape(text):
    # Keep content-stream strings stable in basic Latin for Helvetica.
    safe = []
    for ch in text:
        if ord(ch) < 32 or ord(ch) > 126:
            safe.append("?")
        else:
            safe.append(ch)
    return "".join(safe).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def approx_text_width(text, font_size):
    return len(text) * font_size * 0.52


def truncate_text(text, max_width, font_size):
    t = (text or "").strip()
    if not t:
        return ""
    if approx_text_width(t, font_size) <= max_width:
        return t
    ellipsis = "..."
    while t and approx_text_width(t + ellipsis, font_size) > max_width:
        t = t[:-1]
    return (t + ellipsis) if t else ellipsis


def draw_text(commands, x, y, text, font="F1", size=10):
    commands.append(f"BT /{font} {size} Tf 1 0 0 1 {x:.2f} {y:.2f} Tm ({pdf_escape(text)}) Tj ET")


def draw_line(commands, x1, y1, x2, y2):
    commands.append(f"{x1:.2f} {y1:.2f} m {x2:.2f} {y2:.2f} l S")


class SimplePDF:
    def __init__(self):
        self.pages = []

    def add_page(self, page_commands):
        stream = "\n".join(page_commands).encode("utf-8")
        self.pages.append(stream)

    def _build(self):
        objects = [None]

        def add_obj(data):
            objects.append(data)
            return len(objects) - 1

        catalog_id = add_obj("<< /Type /Catalog /Pages 2 0 R >>")
        add_obj(None)  # Pages placeholder
        font_regular_id = add_obj("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        font_bold_id = add_obj("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")

        page_ids = []
        for content in self.pages:
            content_id = add_obj(
                f"<< /Length {len(content)} >>\nstream\n".encode("utf-8")
                + content
                + b"\nendstream"
            )
            page_obj = (
                "<< /Type /Page /Parent 2 0 R "
                f"/MediaBox [0 0 {PAGE_WIDTH:.0f} {PAGE_HEIGHT:.0f}] "
                f"/Resources << /Font << /F1 {font_regular_id} 0 R /F2 {font_bold_id} 0 R >> >> "
                f"/Contents {content_id} 0 R >>"
            )
            page_ids.append(add_obj(page_obj))

        kids = " ".join(f"{pid} 0 R" for pid in page_ids)
        objects[2] = f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>"

        out = bytearray()
        out.extend(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0] * len(objects)

        for obj_id in range(1, len(objects)):
            offsets[obj_id] = len(out)
            out.extend(f"{obj_id} 0 obj\n".encode("utf-8"))
            data = objects[obj_id]
            if isinstance(data, str):
                out.extend(data.encode("utf-8"))
            else:
                out.extend(data)
            out.extend(b"\nendobj\n")

        xref_start = len(out)
        out.extend(f"xref\n0 {len(objects)}\n".encode("utf-8"))
        out.extend(b"0000000000 65535 f \n")
        for obj_id in range(1, len(objects)):
            out.extend(f"{offsets[obj_id]:010d} 00000 n \n".encode("utf-8"))
        out.extend(
            (
                "trailer\n"
                f"<< /Size {len(objects)} /Root {catalog_id} 0 R >>\n"
                "startxref\n"
                f"{xref_start}\n"
                "%%EOF\n"
            ).encode("utf-8")
        )
        return bytes(out)

    def save(self, path):
        with open(path, "wb") as f:
            f.write(self._build())


def extract_preferred_name(item_modifiers):
    if not item_modifiers:
        return ""
    match = re.search(r"Preferred Name:\s*([^,]+)", item_modifiers, flags=re.IGNORECASE)
    return match.group(1).strip() if match else ""


def safe_int(value, default=1):
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError):
        return default


def normalized_ticket_type(ticket_type):
    return (ticket_type or "").strip().lower()


def ticket_type_rank(ticket_type):
    t = normalized_ticket_type(ticket_type)
    if t == "vip tickets":
        return 0
    if t == "general admission":
        return 1
    if t == "student tickets":
        return 2
    return 3


def load_ticket_rows():
    csv_files = sorted(glob.glob("*.csv"))
    if not csv_files:
        raise RuntimeError("No CSV files found in current folder.")

    rows = []
    for csv_path in csv_files:
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for record in reader:
                if (record.get("Item Name") or "").strip() != TARGET_TICKET_ITEM:
                    continue

                preferred_name = extract_preferred_name(record.get("Item Modifiers") or "")
                if not preferred_name:
                    preferred_name = (
                        (record.get("Order Name") or "").strip()
                        or (record.get("Recipient Name") or "").strip()
                    )
                email = (record.get("Recipient Email") or "").strip()
                ticket_type = (record.get("Item Variation") or "").strip()
                quantity = safe_int((record.get("Item Quantity") or "").strip() or "1", default=1)

                for _ in range(quantity):
                    rows.append(
                        {
                            "preferred_name": preferred_name,
                            "email": email,
                            "ticket_type": ticket_type,
                        }
                    )

    rows.sort(
        key=lambda r: (
            ticket_type_rank(r["ticket_type"]),
            normalized_ticket_type(r["ticket_type"]),
            (r["preferred_name"] or "").lower(),
            (r["email"] or "").lower(),
        )
    )
    return rows


def render_table_page(commands, page_rows):
    col_edges = [TABLE_LEFT]
    running_x = TABLE_LEFT
    for frac in COL_WIDTHS:
        running_x += TABLE_WIDTH * frac
        col_edges.append(running_x)

    rows_on_page = len(page_rows)
    top_y = TABLE_TOP
    bottom_y = TABLE_TOP - HEADER_HEIGHT - (rows_on_page * ROW_HEIGHT)

    draw_line(commands, TABLE_LEFT, top_y, TABLE_RIGHT, top_y)
    draw_line(commands, TABLE_LEFT, bottom_y, TABLE_RIGHT, bottom_y)
    draw_line(commands, TABLE_LEFT, top_y, TABLE_LEFT, bottom_y)
    draw_line(commands, TABLE_RIGHT, top_y, TABLE_RIGHT, bottom_y)

    for x in col_edges[1:-1]:
        draw_line(commands, x, top_y, x, bottom_y)

    header_bottom = top_y - HEADER_HEIGHT
    draw_line(commands, TABLE_LEFT, header_bottom, TABLE_RIGHT, header_bottom)
    for i in range(rows_on_page):
        y = header_bottom - ((i + 1) * ROW_HEIGHT)
        draw_line(commands, TABLE_LEFT, y, TABLE_RIGHT, y)

    headers = ["Preferred Name", "Email", "Ticket Type", "Check"]
    for i, label in enumerate(headers):
        cell_left = col_edges[i]
        cell_right = col_edges[i + 1]
        text_x = cell_left + 6
        text_y = top_y - HEADER_HEIGHT + 7
        if label == "Check":
            w = approx_text_width(label, 10)
            text_x = cell_left + ((cell_right - cell_left - w) / 2)
        draw_text(commands, text_x, text_y, label, font="F2", size=10)

    for row_idx, row in enumerate(page_rows):
        row_top = header_bottom - (row_idx * ROW_HEIGHT)
        baseline_y = row_top - 15

        values = [row["preferred_name"], row["email"], row["ticket_type"]]
        for col_idx, value in enumerate(values):
            cell_left = col_edges[col_idx]
            cell_right = col_edges[col_idx + 1]
            max_w = (cell_right - cell_left) - 12
            text = truncate_text(value, max_w, 9)
            draw_text(commands, cell_left + 6, baseline_y, text, font="F1", size=9)

        # Empty checkbox square in the last column.
        c_left = col_edges[3]
        c_right = col_edges[4]
        size = 10
        x = c_left + ((c_right - c_left - size) / 2)
        y = row_top - ((ROW_HEIGHT + size) / 2)
        commands.append(f"{x:.2f} {y:.2f} {size:.2f} {size:.2f} re S")


def render_pdf(ticket_rows, output_path):
    if not ticket_rows:
        raise RuntimeError(f'No matching items found for "{TARGET_TICKET_ITEM}".')

    max_rows_per_page = int(math.floor((TABLE_HEIGHT - HEADER_HEIGHT) / ROW_HEIGHT))
    if max_rows_per_page <= 0:
        raise RuntimeError("Table layout cannot fit rows on the page with current settings.")

    pdf = SimplePDF()
    for i in range(0, len(ticket_rows), max_rows_per_page):
        page_rows = ticket_rows[i:i + max_rows_per_page]
        commands = []
        render_table_page(commands, page_rows)
        pdf.add_page(commands)

    pdf.save(output_path)


def main():
    rows = load_ticket_rows()
    render_pdf(rows, OUTPUT_PDF)
    print(f"Generated {OUTPUT_PDF} with {len(rows)} ticket rows.")


if __name__ == "__main__":
    main()
