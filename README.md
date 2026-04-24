# Material matching

Static GitHub Pages website for multi-project spool material matching.

Current rule set:
- Each spool is uniquely identified by `Project Code + Drawing No. + Spool No.`
- The BOM format is:
  `Project Code, Drawing No., Spool No., Item Code, Total Q'ty, Priority, Total Wt., Total Inch Dia, Material`
- Each spool includes one master row where `Item Code` is blank
- The master row supplies spool-level `Priority`, `Total Wt.`, `Total Inch Dia`, and optional spool `Material`
- Component rows for the same spool supply `Item Code` and `Total Q'ty` for matching
- The stock format is:
  `Project Code, Item Code, Location, Total Q'ty`
- Stock from one `Project Code` is never used for another `Project Code`
- Stock is checked from `Store` first, then from `QC`

Available analysis modes:
- `By Priority`
  Filled priorities are checked from lowest to highest and blank priorities are treated as last
- `By Wt.`
  Spools are checked by `Total Wt.` from highest to lowest
- `By Inch Dia`
  Spools are checked by `Total Inch Dia` from highest to lowest
- `Best Analysis`
  Spools stay in uploaded order, but if a spool cannot be fully covered then its tentative allocation is released so later spools can use that stock

Final statuses:
- `100% Material Available`
- `100% Material Available but some items are at QC location`
- `Partial Material Available`
- `No Material Available`

Exports:
- `Fabrication readiness by spool` can be downloaded as Excel
- `Component detail` can be downloaded as Excel

GitHub Pages files included:
- `index.html`
- `material-checker.css`
- `material-checker.js`
- `.nojekyll`
- `.github/workflows/deploy-pages.yml`

Publishing note:
- Push this folder to the `main` branch of the GitHub repository
- In repository settings, set Pages source to `GitHub Actions`
- The workflow will publish the site automatically
