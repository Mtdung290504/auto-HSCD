import os
import sys
import pythoncom
import win32com.client
from win32com.client import constants


def compare_word_docs(orig_path, filled_path, output_path):
    pythoncom.CoInitialize()

    word = None
    doc_orig = None
    doc_filled = None
    doc_diff = None

    try:
        # ------------------------------------------------------------------
        # Resolve absolute paths
        # ------------------------------------------------------------------
        orig_abs = os.path.abspath(orig_path)
        filled_abs = os.path.abspath(filled_path)
        out_abs = os.path.abspath(output_path)

        for path in (orig_abs, filled_abs):
            if not os.path.isfile(path):
                print(f"[ERROR] File not found: {path}")
                return 1

        out_dir = os.path.dirname(out_abs)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        print("[INFO] Starting Microsoft Word...")

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = constants.wdAlertsNone
        except Exception:
            word.DisplayAlerts = 0

        print("[INFO] Opening documents...")

        doc_orig = word.Documents.Open(
            FileName=orig_abs,
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        doc_filled = word.Documents.Open(
            FileName=filled_abs,
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        print("[INFO] Comparing documents...")

        wdCompareDestinationNew = 2
        wdGranularityWordLevel = 1
        try:
            wdCompareDestinationNew = constants.wdCompareDestinationNew
        except Exception:
            pass
        try:
            wdGranularityWordLevel = constants.wdGranularityWordLevel
        except Exception:
            pass

        doc_diff = word.CompareDocuments(
            OriginalDocument=doc_orig,
            RevisedDocument=doc_filled,
            Destination=wdCompareDestinationNew,
            Granularity=wdGranularityWordLevel,
            CompareFormatting=True,
            CompareCaseChanges=True,
            CompareWhitespace=True,
            CompareTables=True,
            CompareHeaders=True,
            CompareFootnotes=True,
            CompareTextboxes=True,
            CompareFields=True,
            CompareComments=True,
            CompareMoves=True,
            RevisedAuthor="Diff Checker",
            IgnoreAllComparisonWarnings=True,
        )

        print("[INFO] Saving result...")

        wdFormatXMLDocument = 12
        try:
            wdFormatXMLDocument = constants.wdFormatXMLDocument
        except Exception:
            pass

        doc_diff.SaveAs2(
            FileName=out_abs,
            FileFormat=wdFormatXMLDocument,
            AddToRecentFiles=False,
        )

        print(f"[SUCCESS] Output saved: {out_abs}")

        return 0

    except Exception as e:
        print(f"[ERROR] {e}")
        return 1

    finally:

        for doc in (doc_diff, doc_filled, doc_orig):
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass

        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()


def main():
    if len(sys.argv) != 4:
        print(
            "Usage:\n" "python doc-diff.py original.docx edited.docx output.check.docx"
        )
        sys.exit(1)

    exit_code = compare_word_docs(
        sys.argv[1],
        sys.argv[2],
        sys.argv[3],
    )

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
