# Foresight V5 — Model Builder Only Patch

This patch preserves the current V5 Home, cycle workspace, Explore, Ask AI, scenario workflow, profile menu, reports, and data dictionary.

It changes only the **Create a New Model** workflow by embedding the complete working V24 Model Builder.

Upload these items to the GitHub repository root:

- `index.html` (replace current file)
- `model-builder.html` (new file)
- `assets/js/v24-fixes.js`
- `assets/vendor/jszip.min.js`

Do not delete the existing `examples/` or other repository files.
