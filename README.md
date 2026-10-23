# Report Plotter

Upload Excel/CSV files (or a folder) and generate one report per file into `output/`.

## Run
Needs Python 3.9+.
- Mac/Linux: `chmod +x run.command && ./run.command`
- Windows: double-click `run.bat`

First run creates a local `.venv` and installs the packages in `requirements.txt` (nothing is installed system-wide).
Then open http://localhost:5055 (served on this computer only).

## Report types
- **DR8 Data Distribution** - PowerPoint deck (uses `dr8_template.pptx`, keep it next to `app.py`). Export to PDF from PowerPoint.
- **Chart report** - HTML with box plots or histograms.

Reports are written to `output/<file name>/`. Uploaded copies go to `uploads/`.
