---
name: doc_diff
description: Generate a native Microsoft Word Track Changes document by comparing an original document with an edited document. Use this skill whenever a reviewable Word document with Track Changes is required.
---

# Word Document Comparison

This skill generates a native Microsoft Word Track Changes document by comparing an original document with an edited document using Microsoft Word's built-in Compare feature.

## When to use

Use this skill whenever the workflow requires:

- generating Track Changes;
- producing a reviewable Word document;
- comparing an original document with its edited version.

Do **not** implement your own document comparison algorithm.

Always delegate document comparison to Microsoft Word through this skill.

## Inputs

The script requires:

- Original document (`.docx`)
- Edited document (`.docx`)
- Output document path (`.check.docx`)

## Execution

Run the comparison script:

```bash
python "c:/z-Information Security Level Profile/.agents/skills/doc_diff/scripts/doc-diff.py" "<original_docx>" "<edited_docx>" "<output_check_docx>"
```

## Output

The script produces:

- one `.check.docx` document;
- native Microsoft Word Track Changes;
- all detected insertions, deletions, and formatting changes handled by Microsoft Word.

The generated `.check.docx` file should be treated as the final review document.

## Failure Handling

If the script exits with a non-zero exit code or reports an error:

- stop the current workflow;
- report the error to the user;
- do not claim the document was successfully generated;
- do not replace this workflow with another comparison method.
