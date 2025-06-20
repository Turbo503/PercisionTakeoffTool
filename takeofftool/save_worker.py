# save_worker.py -------------------------------------------------------------
import fitz, json, sys, tempfile, shutil, os, traceback, gc
LOG_PATH = tempfile.mktemp(suffix=".log")


def safe_color(rgb):
    return [max(0.0, min(1.0, float(c))) for c in rgb[:3]]

def main():
    try:
        dest, bundle_json = sys.argv[1], sys.argv[2]
        bundle = json.loads(open(bundle_json, "r", encoding="utf-8").read())
        pdf_bytes = bytes.fromhex(bundle["pdf_hex"])
        highlights = bundle["hl"]

        tmp = tempfile.mktemp(suffix=".pdf")
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            for hl in highlights:
                page = doc.load_page(hl["page"])
                if hl["kind"] == "rect":
                    x0, y0, x1, y1 = hl["rect"]
                    r = fitz.Rect(x0, y0, x1, y1) & page.rect  # clip
                    if r.is_empty or r.width == 0 or r.height == 0:
                        continue
                    annot = page.add_rect_annot(r)
                    annot.set_colors(stroke=None, fill=safe_color(hl["color"]))
                    annot.update()
                else:  # line
                    p1, p2 = hl["p1"], hl["p2"]
                    if p1 == p2:
                        continue
                    annot = page.add_line_annot(
                        fitz.Point(*p1),
                        fitz.Point(*p2)
                    )
                    annot.set_colors(stroke=safe_color(hl["color"]))
                    annot.set_border(width=max(0.1, float(hl["width"])))
                    annot.update()

            doc.save(tmp, incremental=False, garbage=4, deflate=True)
        finally:
            doc.close()
            gc.collect()
        shutil.move(tmp, dest)
        sys.exit(0)

    except Exception:
        # log the traceback so the parent can show it
        with open(LOG_PATH, "w", encoding="utf-8") as fp:
            fp.write(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()
