# DLP scanner

### Requirements

1. **Install Python 3**
   - From the official website: [python.org/downloads](https://www.python.org/downloads/)
   - **Important:** Add Python to PATH during installation.

2. **Prepare your workspace**
  - This one-liner should create the folders in Desktop
      ```bash
      New-Item -Path "C:\Users\$env:USERNAME\Desktop\DLP\attachments" -ItemType Directory
      ```
  - Save the script in the newly created folder (`dlp_email_scanner.py`).
  - Download the dictionary from **Wiki**.
 

3. **Install dependencies**
  - Download requirements.txt
   - Open PowerShell in your project folder and run:
     ```powershell
     python.exe -m pip install requirements.txt
     ```

### Usage

```powershell
python dlp_email_scanner.py <dictionary.json> <file_or_directory> --scan <options>
```

This would be the output:
```bash
PS > py.exe .\dlp_email_scanner.py .\attachments\cc.eml --scan ssn cc

📧 cc.eml
  BODY:
    • 48X-XX-XXXX            [SSN]
    • 48XXXXXXX              [SSN]
    • 622XXXXXXXXXXXXX       [CreditCard]

  📎 Attachment: Testing.docx
    • 48X-XX-XXXX            [SSN]
    • 48XXXXXXX              [SSN]
    • 622XXXXXXXXXXXXX       [CreditCard]

  📎 Attachment: Testing.pdf
    • 48XXXXXXX              [SSN]
    • 622XXXXXXXXXXXXX       [CreditCard]
```

**Note**: It is optional to use the dictionary. If you want to use it, you will need to use: `--scan dict`.
This option will show you for the dictionaries you want to check against.

## Notes

- **DO NOT** provide the output to the customer. The output is for **internal** use only.

- To convert `.msg` files to `.eml`, use:

  ```bash
  msgconvert --mbox email.eml email.msg
  ```
  - You may need to install `msgconvert` in WSL:
    ```bash
    sudo apt install libemail-outlook-message-perl
    ```

- Make sure to use the correct Python version (3.x) to run the script. If you have multiple versions installed, use `py.exe` or specify the full path to the Python 3 executable.
- The script scans for sensitive information based on the terms defined in `dlp_terms.json`. It will output matches found in the email body or attachments. So please verify the rule and the terms that are triggering in the email.
- The script can handle various file types, including `.eml`, `.pdf`, `.docx`, `.txt`, `.zip`. 🚧 Soon: `.rar`, `.msg`, `.html`, `.csv`. Ensure that the attachments are in the `attachments` folder. 
- **Validate SSN:** [ssnregistry.org/validate](https://www.ssnregistry.org/validate/)
- **Validate Credit Cards:** [validcreditcardnumber.com](https://www.validcreditcardnumber.com/)
- **Validate NDC:** [dps.fda.gov/ndc](https://dps.fda.gov/ndc/)