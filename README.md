# Split Word Doc (PDF & DOCX)

A Python utility to split large Microsoft Word documents into smaller chunks of a specific page length. It supports exporting to both **DOCX** (maintaining original formatting) and **PDF**.

## 🚀 Features

  * **Automated Pagination:** Automatically detects total page counts and calculates splits.
  * **Dual Format Support:**
      * **DOCX:** Extracts raw content and formatting into new Word documents.
      * **PDF:** Exports high-quality, print-optimized PDF chunks.
  * **Large File Handling:** Designed to handle massive documents by automating the Word Background Application.
  * **Format Preservation:** Uses `PasteAndFormat` to ensure the original styling, fonts, and layouts remain intact.

## 🛠️ Requirements

  * **Operating System:** Windows (Required for `pywin32` and Microsoft Word COM).
  * **Software:** Microsoft Word must be installed.
  * **Python Libraries:**
    ```bash
    pip install pywin32
    ```

## 💻 Usage

Run the script from the command line by providing the input file and the target output directory.

```bash
python SplitWordDoc.py <input.docx> <output_folder> [chunk_size] [format]
```

### Parameters:

| Parameter | Description | Default |
| :--- | :--- | :--- |
| `input.docx` | The path to the Word file you want to split. | *Required* |
| `output_folder`| Where the split files will be saved. | *Required* |
| `chunk_size` | Number of pages per split file. | `500` |
| `format` | Output type: `docx` or `pdf`. | `docx` |

### Examples:

**Split into 100-page Word documents:**

```bash
python SplitWordDoc.py manual.docx ./output_chunks 100 docx
```

**Split into 50-page PDF segments:**

```bash
python SplitWordDoc.py thesis.docx ./pdf_parts 50 pdf
```

## 🔍 How it Works

The script utilizes the `win32com` library to interface directly with the `Word.Application` engine.

1.  **PDF Export:** Uses Word's internal `ExportAsFixedFormat` engine for perfect PDF reproduction.
2.  **DOCX Export:** Navigates the document's `Range` objects via `GoTo` (Page level) to copy and paste specific sections into new document instances.

## ⚠️ Important Notes

  * **Reflowable Text:** DOCX is a "reflowable" format. While the script selects content based on Word's current pagination, small variations in layout might occur in the new DOCX chunks.
  * **Background Process:** The script runs Word in `Visible = False` mode. If the script crashes, you may occasionally see a `Word` process left open in your Task Manager.
