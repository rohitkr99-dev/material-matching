# Material matching

GitHub Pages-ready static website for BOM versus stock readiness checking.

Current rule set:
- The spool is identified by the combination of `Project Code + Drawing No. + Spool No.`
- Stock is allocated from `Store` first
- Remaining demand is then checked against `QC`
- Final statuses are:
  - `100% Material Available`
  - `100% Material Available but some items are at QC location`
  - `Partial Material Available`
  - `No Material Available`

GitHub Pages files included:
- `index.html`
- `material-checker.css`
- `material-checker.js`
- `.nojekyll`
- `.github/workflows/deploy-pages.yml`

Publishing note:
- Push to the `main` branch.
- In GitHub repository settings, set Pages source to `GitHub Actions`.
- The workflow will publish the site automatically.
