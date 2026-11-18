from core.paths import WORK_DIR, DATA_XLSX, PDF_DIR, MERGED_DIR, LOGS_DIR, EXPORTERS_XLSX
import os
import re
from PyPDF2 import PdfMerger

# --- Path Configuration ---
OUTPUT_DIR = MERGED_DIR

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


def merge_list(files, out_name):
    merger = PdfMerger()
    for f in files:
        merger.append(os.path.join(PDF_DIR, f))
    merger.write(os.path.join(OUTPUT_DIR, out_name))
    merger.close()


def merge_pdfs():
    all_files = [f for f in os.listdir(PDF_DIR) if f.lower().endswith('.pdf')]

    if not all_files:
        print("Folder එකේ පයිල් මුකුත් නැහැ මචං!")
        return

    boat_notes = []
    amend_docs = []
    e_files = []
    p_files = {}
    used_files = set()

    print("පයිල් ටික වර්ග කරනවා...")

    for f in all_files:
        fname = os.path.splitext(f)[0].upper().strip()

        # 0. Amend Docs (SS පයිල්) - අනිත් ඒවට කලින්ම check කරනවා
        if fname.startswith('SS'):
            amend_docs.append(f)

        # 1. Boat Notes (B පයිල්)
        elif fname.startswith('B'):
            boat_notes.append(f)

        # 2. CusDecs - P පයිල්
        elif fname.startswith('P'):
            num_p = re.sub(r'\D', '', fname)
            if num_p:
                p_files[num_p] = f

        # 3. CusDecs - E පයිල්
        elif fname.startswith('E'):
            e_files.append(f)

    # --- 1. BOAT NOTES MERGE ---
    if boat_notes:
        merge_list(sorted(boat_notes), "BOAT NOTES.pdf")
        used_files.update(boat_notes)
        print("✅ BOAT NOTES.pdf හැදුණා.")

    # --- 2. CUSDECS MERGE (E + P) ---
    cusdec_merger = PdfMerger()
    has_cusdec = False

    for p_num, p_file in p_files.items():
        matched_e = None
        for e_file in e_files:
            if p_num in e_file:
                matched_e = e_file
                break

        if matched_e:
            cusdec_merger.append(os.path.join(PDF_DIR, matched_e))
            cusdec_merger.append(os.path.join(PDF_DIR, p_file))
            used_files.add(matched_e)
            used_files.add(p_file)
            has_cusdec = True
            print(f"Matched: {matched_e} + {p_file}")

    if has_cusdec:
        cusdec_merger.write(os.path.join(OUTPUT_DIR, "CUSDECS.pdf"))
        print("✅ CUSDECS.pdf හැදුණා.")
    cusdec_merger.close()

    # --- 3. AMEND DOCS MERGE (SS පයිල්) ---
    if amend_docs:
        merge_list(sorted(amend_docs), "AMEND DOCS.pdf")
        used_files.update(amend_docs)
        print(f"✅ AMEND DOCS.pdf හැදුණා (SS පයිල් {len(amend_docs)}ක්).")

    # --- 4. CDNS MERGE (ඉතිරි ඔක්කොම) ---
    cdn_files = [f for f in all_files if f not in used_files]
    if cdn_files:
        merge_list(sorted(cdn_files), "CDNS.pdf")
        used_files.update(cdn_files)
        print(f"✅ ඉතිරි පයිල් {len(cdn_files)}ක් CDNS.pdf එකට එකතු කළා.")

    # --- 5. AUTO DELETE ORIGINAL FILES ---
    print("\nපරණ පයිල් ටික අයින් කරනවා (Cleaning up)...")
    for f in used_files:
        try:
            os.remove(os.path.join(PDF_DIR, f))
            print(f"Deleted: {f}")
        except Exception as e:
            print(f"Error deleting {f}: {e}")

    print("\nවැඩේ සාර්ථකව අවසන්! Merge වුණ පයිල් 'merged' folder එකේ ඇති.")


if __name__ == "__main__":
    merge_pdfs()