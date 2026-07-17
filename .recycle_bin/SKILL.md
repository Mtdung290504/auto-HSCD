# Compare Word Documents

## Purpose

Generate a Microsoft Word Track Changes document by comparing an original document and an edited document.

This skill must be used whenever the workflow requires Track Changes.

Do NOT implement document diff yourself.

Always delegate comparison to Microsoft Word Compare.

---

## Inputs

- Original document (.docx)
- Edited document (.docx)
- Output path (.check.docx)

---

## Execution

Run:

python doc_diff.py "<original>" "<edited>" "<output>"

---

## Output

One .check.docx document containing native Microsoft Word Track Changes.

---

## Failure

If the script returns a non-zero exit code:

- stop the workflow
- report the error
- do not claim success
