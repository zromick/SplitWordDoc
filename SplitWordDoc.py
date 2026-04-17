"""
Word Document Splitter
Splits large Word documents into smaller chunks (PDF or DOCX format)
"""

import os
import sys
import win32com.client

# Hardcoded MS Word constants to completely bypass COM cache errors
WD_STATISTIC_PAGES = 2
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_FROM_TO = 3
WD_EXPORT_DOCUMENT_WITH_MARKUP = 7
WD_EXPORT_CREATE_NO_BOOKMARKS = 0
WD_GO_TO_PAGE = 1
WD_GO_TO_ABSOLUTE = 1
WD_FORMAT_ORIGINAL_FORMATTING = 16
WD_FORMAT_XML_DOCUMENT = 12


def init_word():
    """Initializes Word using dynamic dispatch to bypass cache issues."""
    try:
        return win32com.client.Dispatch("Word.Application")
    except Exception as e:
        print(f"❌ Failed to initialize Word: {e}")
        sys.exit(1)


def select_folder(prompt="Select a folder"):
    """Cross-platform folder selection dialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()

        if sys.platform == "darwin":
            root.call("wm", "attributes", ".", "-topmost", True)
        else:
            root.attributes("-topmost", True)

        folder = filedialog.askdirectory(title=prompt)
        root.destroy()

        return folder if folder else None

    except Exception as e:
        print(f"\n⚠️  GUI not available: {e}")
        path = input(f"{prompt} (enter full path): ").strip()
        return path if os.path.isdir(path) else None


def select_file(prompt="Select a file", filetypes=None):
    """Cross-platform file selection dialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()

        if sys.platform == "darwin":
            root.call("wm", "attributes", ".", "-topmost", True)
        else:
            root.attributes("-topmost", True)

        if filetypes is None:
            filetypes = [("Word Documents", "*.docx"), ("All Files", "*.*")]

        file = filedialog.askopenfilename(title=prompt, filetypes=filetypes)
        root.destroy()

        return file if file else None

    except Exception as e:
        print(f"\n⚠️  GUI not available: {e}")
        path = input(f"{prompt} (enter full path): ").strip()
        return path if os.path.isfile(path) else None


def split_to_pdf_chunks(doc_path, out_dir, chunk_size):
    """Split document into PDF chunks."""
    doc_path = os.path.abspath(doc_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    print("\n🔄 Initializing Microsoft Word...")
    word = init_word()
    word.Visible = False

    try:
        print(f"📖 Opening document: {os.path.basename(doc_path)}")
        doc = word.Documents.Open(doc_path, ReadOnly=True)
        total_pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        print(f"📄 Total pages: {total_pages}")

        base = os.path.splitext(os.path.basename(doc_path))[0]
        start = 1
        chunk_num = 1

        while start <= total_pages:
            end = min(start + chunk_size - 1, total_pages)
            out_pdf = os.path.join(
                out_dir, f"{base}_chunk{chunk_num}_pages{start}-{end}.pdf"
            )

            print(f"📤 Exporting chunk {chunk_num}: pages {start}–{end}...")

            doc.ExportAsFixedFormat(
                OutputFileName=out_pdf,
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                Range=WD_EXPORT_FROM_TO,
                From=start,
                To=end,
                Item=WD_EXPORT_DOCUMENT_WITH_MARKUP,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=WD_EXPORT_CREATE_NO_BOOKMARKS,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )

            start = end + 1
            chunk_num += 1

        print(f"\n✅ Successfully created {chunk_num - 1} PDF chunks!")
        print(f"📂 Output location: {out_dir}")

    finally:
        print("🔒 Closing Word...")
        word.Quit()


def split_to_docx_chunks(doc_path, out_dir, chunk_size):
    """Split document into DOCX chunks."""
    doc_path = os.path.abspath(doc_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    print("\n🔄 Initializing Microsoft Word...")
    word = init_word()
    word.Visible = False

    try:
        print(f"📖 Opening document: {os.path.basename(doc_path)}")
        doc = word.Documents.Open(doc_path, ReadOnly=True)
        total_pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        print(f"📄 Total pages: {total_pages}")
        print("⚠️  Note: DOCX pagination may shift slightly in chunks")

        base = os.path.splitext(os.path.basename(doc_path))[0]
        start_page = 1
        chunk_num = 1

        while start_page <= total_pages:
            end_page = min(start_page + chunk_size - 1, total_pages)
            print(
                f"📤 Creating DOCX chunk {chunk_num}: pages {start_page}–{end_page}..."
            )

            startRange = doc.GoTo(
                What=WD_GO_TO_PAGE,
                Which=WD_GO_TO_ABSOLUTE,
                Count=start_page,
            ).Start
            if end_page < total_pages:
                endRange = doc.GoTo(
                    What=WD_GO_TO_PAGE,
                    Which=WD_GO_TO_ABSOLUTE,
                    Count=end_page + 1,
                ).Start
            else:
                endRange = doc.Content.End

            rng = doc.Range(Start=startRange, End=endRange)

            newdoc = word.Documents.Add()
            rng.Copy()
            newdoc.Range(0, 0).PasteAndFormat(WD_FORMAT_ORIGINAL_FORMATTING)

            out_docx = os.path.join(
                out_dir, f"{base}_chunk{chunk_num}_pages{start_page}-{end_page}.docx"
            )
            newdoc.SaveAs2(out_docx, FileFormat=WD_FORMAT_XML_DOCUMENT)
            newdoc.Close(SaveChanges=False)

            start_page = end_page + 1
            chunk_num += 1

        print(f"\n✅ Successfully created {chunk_num - 1} DOCX chunks!")
        print(f"📂 Output location: {out_dir}")

    finally:
        print("🔒 Closing Word...")
        word.Quit()


def main():
    """Interactive walkthrough for splitting Word documents."""

    print("=" * 70)
    print("📄 WORD DOCUMENT SPLITTER")
    print("=" * 70)
    print("\nThis tool splits large Word documents into smaller chunks.")
    print("Useful for documents that are too large to process or share easily.\n")

    print("STEP 1: Select the Word document to split")
    print("-" * 70)
    input_file = select_file(
        prompt="Select Word document (.docx)",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
    )

    if not input_file:
        print("❌ No file selected. Exiting...")
        return

    if not input_file.lower().endswith(".docx"):
        print("❌ Error: File must be a .docx document")
        return

    print(f"✅ Selected: {os.path.basename(input_file)}\n")

    print("STEP 2: Select output folder for chunks")
    print("-" * 70)
    print("(Where should the split files be saved?)")
    output_folder = select_folder(prompt="Select output folder")

    if not output_folder:
        print("❌ No folder selected. Exiting...")
        return

    print(f"✅ Output folder: {output_folder}\n")

    print("STEP 3: Set chunk size")
    print("-" * 70)
    print("How many pages per chunk?")
    print("  - Small files: 100-200 pages")
    print("  - Medium files: 300-500 pages")
    print("  - Large files: 500+ pages")

    while True:
        chunk_input = input("\nEnter pages per chunk (default: 500): ").strip()
        if not chunk_input:
            chunk_size = 500
            break
        if chunk_input.isdigit() and int(chunk_input) > 0:
            chunk_size = int(chunk_input)
            break
        print("❌ Please enter a valid positive number")

    print(f"✅ Chunk size: {chunk_size} pages\n")

    print("STEP 4: Choose output format")
    print("-" * 70)
    print("  1. PDF (recommended - maintains exact layout)")
    print("  2. DOCX (editable, but pagination may shift)")

    while True:
        format_choice = input("\nEnter 1 for DOCX or 2 for PDF (default: 1): ").strip()
        if not format_choice or format_choice == "1":
            output_format = "docx"
            break
        elif format_choice == "2":
            output_format = "pdf"
            break
        print("❌ Please enter 1 or 2")

    print(f"✅ Output format: {output_format.upper()}\n")

    print("STEP 5: Review and process")
    print("=" * 70)
    print(f"  Input file:     {os.path.basename(input_file)}")
    print(f"  Output folder:  {output_folder}")
    print(f"  Chunk size:     {chunk_size} pages")
    print(f"  Format:         {output_format.upper()}")
    print("=" * 70)

    confirm = input("\nProceed with splitting? (Y/n): ").strip().lower()
    if confirm and confirm != "y":
        print("❌ Cancelled by user.")
        return

    try:
        if output_format == "pdf":
            split_to_pdf_chunks(input_file, output_folder, chunk_size)
        else:
            split_to_docx_chunks(input_file, output_folder, chunk_size)
    except Exception as e:
        print(f"\n❌ Error during processing: {e}")
        return

    print("\n" + "=" * 70)
    print("🎉 SPLITTING COMPLETE!")
    print("=" * 70)


if __name__ == "__main__":
    main()
