# Generate 2024 monthly .xlsx files from PDFs

This adds a script scripts/generate_2024_from_pdfs.py which uses the templates in examples_2023_split/ to
create .xlsx files for each month found in data_2024/*.pdf and saves them in output_2024/.

How it works
- For each PDF named like GENNAIO_2024.pdf it uses the corresponding template "GENNAIO 2023.xlsx" from
  examples_2023_split/ if present; otherwise falls back to examples_2023_split/2023.xlsx.
- It copies the template to output_2024/<MONTH> 2024.xlsx and writes the extracted table rows into the first sheet
  starting at row 4 (configurable in the script).

Run locally
1. Ensure Python 3.8+ is installed.
2. Create a virtualenv and install dependencies:

   pip install pdfplumber pandas openpyxl xlrd

3. From repository root run:

   python3 scripts/generate_2024_from_pdfs.py

Notes
- The script makes heuristics about table extraction; review the generated xlsx for formatting/data alignment.
- If you want me to run the extraction and commit the generated .xlsx files directly into output_2024, I can do that next — but I cannot execute code in this environment. Run the script locally or in CI and then I can help commit the outputs.
