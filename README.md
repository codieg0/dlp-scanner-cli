# DLP Email & Attachment Scanner CLI

A Python command-line tool for scanning emails and attachments for sensitive data, including dictionary terms, smart identifiers (SSN, Credit Card, US Driver License), and custom terms. Supports scanning inside .eml, .msg, and .zip files.

## Features

- Scans files for sensitive information using a customizable dictionary (`dlp_terms.json`).
- Detects US Social Security Numbers (SSN), Credit Card numbers (with Luhn check), and US Driver License numbers.
- Supports scanning:
  - Plain text files (.txt)
  - Microsoft Office files (.docx, .xlsx, .xls)
  - PDF files (.pdf)
  - CSV files (.csv)
  - HTML/RTF files (.htm, .html, .rtf)
  - Email files (.eml, .msg)
  - Zip archives (.zip) recursively
- Colorful, emoji-enhanced CLI output for easy review.
- Logs errors and warnings to `Logs/dlp_errors.log`.
- Privacy warning at startup.

## Requirements

- Python 3.8+
- Install dependencies:

  ```powershell
  py.exe -m pip install -r requirements.txt
  ```

  (See `requirements.txt` for: colorama, pandas, PyPDF2, beautifulsoup4, extract-msg, python-docx)

- Create the folders in your Desktop folder

    ```powershell
    New-Item -Path "C:\Users\$env:USERNAME\Desktop\DLP\attachments" -ItemType Directory
    New-Item -Path "C:\Users\$env:USERNAME\Desktop\DLP\Logs" -ItemType Directory
    ```

## Usage

1. Place your attachments in the `attachments` folder (in the same directory as the script).
2. Ensure `dlp_terms.json` is present in the same directory as the script.
3. Run the script in PowerShell:

   ```powershell
   py.exe .\dlp_email_scanner.py
   ```

4. Follow the prompts:
   - Enter attachment filenames (comma-separated).
   - Choose dictionary scan options and/or smart identifier scan options.
   - Review the color-coded results.

## Output Example

```text
Attachment name - email.eml
    Body/subject:
        📗 • SSN   [SSN-Term]
        🆔 • 489-36-8350   [SSN]
    SSN.txt
        📗 • SSN   [SSN-Term]
        🆔 • 489-36-8350   [SSN]
```

- 📗 Green: Dictionary match
- 🆔 Red: SSN
- 💳 Green: Credit Card
- 🪪 Blue: US Driver License

## Error Logging

- Errors and warnings are logged to `Logs/dlp_errors.log` with timestamps.

## Privacy Warning

⚠️ This tool scans sensitive data. Please keep the output private and secure.
⚠️ Esta herramienta analiza datos sensibles. Por favor, mantenga el contenido en privado.

## License

MIT License

---

**Note:** This tool is for demonstration and internal use. Always handle sensitive data responsibly.
