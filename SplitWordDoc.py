import os
import sys
from win32com.client import Dispatch, constants, gencache

def split_to_pdf_chunks(doc_path, out_dir, chunk_size=500):
    doc_path = os.path.abspath(doc_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    gencache.EnsureDispatch('Word.Application')
    word = Dispatch('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(doc_path, ReadOnly=True)
        total_pages = doc.ComputeStatistics(constants.wdStatisticPages)
        print(f"Total pages: {total_pages}")

        base = os.path.splitext(os.path.basename(doc_path))[0]
        start = 1
        while start <= total_pages:
            end = min(start + chunk_size - 1, total_pages)
            out_pdf = os.path.join(out_dir, f"{base}_pages_{start}_to_{end}.pdf")
            print(f"Exporting pages {start}–{end} -> {out_pdf}")

            doc.ExportAsFixedFormat(
                OutputFileName=out_pdf,
                ExportFormat=constants.wdExportFormatPDF,
                OpenAfterExport=False,
                OptimizeFor=constants.wdExportOptimizeForPrint,
                Range=constants.wdExportFromTo,
                From=start,
                To=end,
                Item=constants.wdExportDocumentWithMarkup,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=constants.wdExportCreateNoBookmarks,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )

            start = end + 1

        print("Done.")
    finally:
        word.Quit()

def split_to_docx_chunks(doc_path, out_dir, chunk_size=500):
    # NOTE: DOCX is reflowable; pagination may shift a little in chunks.
    doc_path = os.path.abspath(doc_path)
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    gencache.EnsureDispatch('Word.Application')
    word = Dispatch('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(doc_path, ReadOnly=True)
        total_pages = doc.ComputeStatistics(constants.wdStatisticPages)
        print(f"Total pages: {total_pages}")

        base = os.path.splitext(os.path.basename(doc_path))[0]
        start_page = 1

        while start_page <= total_pages:
            end_page = min(start_page + chunk_size - 1, total_pages)
            print(f"Creating DOCX chunk for pages {start_page}–{end_page}...")

            startRange = doc.GoTo(What=constants.wdGoToPage,
                                  Which=constants.wdGoToAbsolute,
                                  Count=start_page).Start
            if end_page < total_pages:
                endRange = doc.GoTo(What=constants.wdGoToPage,
                                    Which=constants.wdGoToAbsolute,
                                    Count=end_page + 1).Start
            else:
                endRange = doc.Content.End

            rng = doc.Range(Start=startRange, End=endRange)

            newdoc = word.Documents.Add()
            rng.Copy()
            newdoc.Range(0, 0).PasteAndFormat(constants.wdFormatOriginalFormatting)

            out_docx = os.path.join(out_dir, f"{base}_pages_{start_page}_to_{end_page}.docx")
            newdoc.SaveAs2(out_docx, FileFormat=constants.wdFormatXMLDocument)
            newdoc.Close(SaveChanges=False)

            start_page = end_page + 1

        print("Done.")
    finally:
        word.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage:")
        print("  python SplitWordDoc.py <input.docx> <output_folder> [chunk_size] [docx|pdf]")
        print("Defaults: chunk_size=500, format=docx")
        sys.exit(1)

    input_docx = sys.argv[1]
    output_folder = sys.argv[2]
    chunk = int(sys.argv[3]) if len(sys.argv) >= 4 and sys.argv[3].isdigit() else 500
    # Default format is DOCX now
    fmt = (sys.argv[4].lower() if len(sys.argv) >= 5 else "docx")

    if fmt == "pdf":
        split_to_pdf_chunks(input_docx, output_folder, chunk)
    else:
        split_to_docx_chunks(input_docx, output_folder, chunk)