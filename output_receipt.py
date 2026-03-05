#!/usr/bin/env python3
import csv
import glob
import re
import math
from collections import OrderedDict

TICKET_ITEMS_TO_SKIP = {
    "[Ticket] 3/8 KP! 3rd Oneman Live",
    "[Ticket] 2/7 Valentine Pop-Up Live",
}

PAGE_WIDTH = 612.0   # US Letter width in points
PAGE_HEIGHT = 792.0  # US Letter height in points
COLS = 3
ROWS = 5
CELL_WIDTH = PAGE_WIDTH / COLS
CELL_HEIGHT = PAGE_HEIGHT / ROWS

def normalize_item_name(name):
    n = (name or "").strip()
    if n == "Cheki Tickets [KP! 3rd Live 3/8]":
        return "Cheki Tickets 3rd"
    if "Pixel Figures Acrylic Keychains" in n:
        return "Pixel Keychains"
    if "Hoodie" in n:
        return "Hoodie"
    if "Tote Bag" in n:
        return "Tote Bag"
    if "Kuri Graduation Merch Set" in n:
        return "Kuri Grad Set"
    if "Buzz Graduation Merch Set" in n:
        return "Buzz Grad Set"
    if "Buzzly Graduation" in n:
        return "Buzz Grad Set"
    return n


def pdf_escape(text):
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def approx_text_width(text, font_size):
    return len(text) * font_size * 0.52


def wrap_text(text, font_size, max_width):
    words = text.split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        candidate = current + " " + word
        if approx_text_width(candidate, font_size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def draw_text(commands, x, y, text, font="F1", size=10):
    safe = pdf_escape(text)
    commands.append(f"BT /{font} {size} Tf 1 0 0 1 {x:.2f} {y:.2f} Tm ({safe}) Tj ET")


class SimplePDF:
    def __init__(self):
        self.pages = []

    def add_page(self, page_commands):
        stream = "\n".join(page_commands).encode("utf-8")
        self.pages.append(stream)

    def _build(self):
        objects = [None]  # object numbers are 1-based

        def add_obj(data):
            objects.append(data)
            return len(objects) - 1

        catalog_id = add_obj("<< /Type /Catalog /Pages 2 0 R >>")
        add_obj(None)  # placeholder for /Pages
        font1_id = add_obj("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        font2_id = add_obj("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")

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
                f"/Resources << /Font << /F1 {font1_id} 0 R /F2 {font2_id} 0 R >> >> "
                f"/Contents {content_id} 0 R >>"
            )
            page_ids.append(add_obj(page_obj))

        kids = " ".join(f"{pid} 0 R" for pid in page_ids)
        objects[2] = f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>"

        output = bytearray()
        output.extend(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0] * len(objects)

        for obj_id in range(1, len(objects)):
            offsets[obj_id] = len(output)
            output.extend(f"{obj_id} 0 obj\n".encode("utf-8"))
            data = objects[obj_id]
            if isinstance(data, str):
                output.extend(data.encode("utf-8"))
            else:
                output.extend(data)
            output.extend(b"\nendobj\n")

        xref_start = len(output)
        output.extend(f"xref\n0 {len(objects)}\n".encode("utf-8"))
        output.extend(b"0000000000 65535 f \n")
        for obj_id in range(1, len(objects)):
            output.extend(f"{offsets[obj_id]:010d} 00000 n \n".encode("utf-8"))
        output.extend(
            (
                "trailer\n"
                f"<< /Size {len(objects)} /Root {catalog_id} 0 R >>\n"
                "startxref\n"
                f"{xref_start}\n"
                "%%EOF\n"
            ).encode("utf-8")
        )
        return bytes(output)

    def save(self, path):
        with open(path, "wb") as f:
            f.write(self._build())


def extract_preferred_name(item_modifiers):
    if not item_modifiers:
        return ""
    m = re.search(r"Preferred Name:\s*([^,]+)", item_modifiers, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


def load_receipts():
    grouped = OrderedDict()
    csv_files = sorted(glob.glob("*.csv"))
    if not csv_files:
        raise RuntimeError("No CSV files found in current folder.")

    for csv_path in csv_files:
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                item_name = (row.get("Item Name") or "").strip()
                if item_name in TICKET_ITEMS_TO_SKIP:
                    continue
                if "valentine" in item_name.lower():
                    continue

                email = (row.get("Recipient Email") or "").strip().lower()
                order_id = (row.get("Order") or "").strip()
                qty = (row.get("Item Quantity") or "").strip() or "1"
                variation = (row.get("Item Variation") or "").strip()
                preferred_name = extract_preferred_name(row.get("Item Modifiers") or "")
                fallback_name = (
                    (row.get("Order Name") or "").strip()
                    or (row.get("Recipient Name") or "").strip()
                    or "Unknown"
                )
                key = email or f"__no_email__::{order_id}"

                if key not in grouped:
                    grouped[key] = {
                        "email": email,
                        "name_candidates": [],
                        "fallback_name": fallback_name,
                        "items": [],
                    }

                if preferred_name:
                    grouped[key]["name_candidates"].append(preferred_name)
                grouped[key]["items"].append(
                    {
                        "order": order_id,
                        "qty": qty,
                        "item_name": item_name,
                        "variation": variation,
                    }
                )

    receipts = []
    for person in grouped.values():
        if not person["items"]:
            continue
        preferred_name = (
            person["name_candidates"][0] if person["name_candidates"] else person["fallback_name"]
        )
        receipts.append({"preferred_name": preferred_name, "items": person["items"]})
    return receipts


def receipt_lines(receipt):
    lines = []
    for item in receipt["items"]:
        item_name = normalize_item_name(item["item_name"])
        variation = item["variation"].strip()
        if variation.lower() == "regular":
            variation = ""

        if variation:
            lines.append(f"- {item['qty']} x {item_name} ({variation})")
        else:
            lines.append(f"- {item['qty']} x {item_name}")
    return lines



def render_pdf(receipts, output_path):
    if not receipts:
        raise RuntimeError("No non-ticket items found in CSV files.")

    pdf = SimplePDF()
    per_page = COLS * ROWS

    pad = 12
    name_size = 20
    detail_size = 10
    line_gap = 3
    max_text_width = CELL_WIDTH - (pad * 2)

    # Expand each person into one or more adjacent panels if details overflow.
    panels = []
    for receipt in receipts:
        base_lines = []
        for base_line in receipt_lines(receipt):
            base_lines.extend(wrap_text(base_line, detail_size, max_text_width))

        name_lines = min(2, len(wrap_text(receipt["preferred_name"], name_size, max_text_width)))
        available_height = CELL_HEIGHT - (
            pad  # top
            + (name_lines * (name_size + 3))
            + 4   # name/content gap
            + 14  # separator line height step
            + pad # bottom
        )
        lines_per_panel = max(1, int(math.floor(available_height / (detail_size + line_gap))))

        total_panels = max(1, int(math.ceil(len(base_lines) / lines_per_panel)))
        for i in range(total_panels):
            start = i * lines_per_panel
            end = start + lines_per_panel
            chunk_lines = base_lines[start:end]
            label = f" ({i + 1}/{total_panels})" if total_panels > 1 else ""
            panels.append(
                {
                    "preferred_name": receipt["preferred_name"] + label,
                    "lines": chunk_lines,
                }
            )

    for page_start in range(0, len(panels), per_page):
        page_receipts = panels[page_start: page_start + per_page]
        commands = []

        for idx, receipt in enumerate(page_receipts):
            row = idx // COLS
            col = idx % COLS
            x = col * CELL_WIDTH
            y = PAGE_HEIGHT - (row + 1) * CELL_HEIGHT

            commands.append(f"{x:.2f} {y:.2f} {CELL_WIDTH:.2f} {CELL_HEIGHT:.2f} re S")

            content_left = x + pad
            cursor_y = y + CELL_HEIGHT - pad - name_size

            name_lines = wrap_text(receipt["preferred_name"], name_size, max_text_width)
            for ln in name_lines[:2]:
                draw_text(commands, content_left, cursor_y, ln, font="F2", size=name_size)
                cursor_y -= name_size + 3

            cursor_y -= 4
            draw_text(commands, content_left, cursor_y, "-" * 34, font="F1", size=9)
            cursor_y -= 14

            max_bottom = y + pad
            for ln in receipt["lines"]:
                if cursor_y - detail_size < max_bottom:
                    break
                draw_text(commands, content_left, cursor_y, ln, font="F1", size=detail_size)
                cursor_y -= detail_size + line_gap

        pdf.add_page(commands)

    pdf.save(output_path)


def main():
    output_path = "receipts.pdf"
    receipts = load_receipts()
    render_pdf(receipts, output_path)
    print(f"Generated {output_path} with {len(receipts)} receipts.")


if __name__ == "__main__":
    main()
