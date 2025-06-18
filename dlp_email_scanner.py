import os
import sys
import re
import tempfile
import pandas as pd
import csv
from email import policy
from email.parser import BytesParser
from docx import Document
import PyPDF2
from bs4 import BeautifulSoup
import json
from colorama import init, Fore, Style
from datetime import datetime
import zipfile
import shutil
import extract_msg

# Force colorama to always convert ANSI codes
init(autoreset=True, strip=False, convert=True)

# # Test color output
# print(Fore.YELLOW + Style.BRIGHT + "[Colorama Test] If you see this in yellow, colorama is working!")

def load_dlp_dict(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data  # Keep as {category: [terms]}

def flatten_terms(dlp_dict):
    flat = {}
    for category, terms in dlp_dict.items():
        for term in terms:
            flat[term.strip()] = category
    return flat

def find_dlp_terms(text, dlp_dict, selected_terms=None):
    found = []
    for term, category in dlp_dict.items():
        if selected_terms and term not in selected_terms:
            continue
        pattern = re.compile(re.escape(term), re.IGNORECASE)
        if pattern.search(text):
            found.append((term, category))
    return found

def find_ssn(text):
    formatted = re.findall(r'\b(?!666|000|9\d{2})\d{3}-(?!00)\d{2}-(?!0000)\d{4}\b', text)
    unformatted = re.findall(r'\b(?!666|000|9\d{2})\d{9}\b', text)
    excluded_range = set(str(i) for i in range(87654320, 87654330))
    filtered_unformatted = [ssn for ssn in unformatted if ssn not in excluded_range]
    return list(set(formatted + filtered_unformatted))

def luhn_check(card_number):
    digits = [int(d) for d in card_number if d.isdigit()]
    checksum = 0
    reverse_digits = digits[::-1]
    for i, d in enumerate(reverse_digits):
        if i % 2 == 1:
            doubled = d * 2
            checksum += doubled - 9 if doubled > 9 else doubled
        else:
            checksum += d
    return checksum % 10 == 0

def find_credit_cards(text):
    cc_patterns = [
        r'\b3[47]\d{13}\b',                     # Amex
        r'\b3(0[0-5]|[68]\d)\d{11}\b',          # Diners Club
        r'\b6011\d{12}\b',                      # Discover
        r'\b5[1-5]\d{14}\b',                    # MasterCard
        r'\b62\d{14}\b',                        # Union Pay
        r'\b4\d{12}(\d{3})?\b'                  # Visa
    ]
    found = set()
    for pattern in cc_patterns:
        for match in re.findall(pattern, text):
            match_str = ''.join(match) if isinstance(match, tuple) else match
            if luhn_check(match_str):
                found.add(match_str)
    return list(found)

def find_us_driver_license(text):
    patterns = [
        r'\b\d{5,13}\b',
        r'\b[A-Z]{1,2}\d{5,13}\b'
    ]
    found = set()
    for pattern in patterns:
        found.update(re.findall(pattern, text))
    return list(found)

def extract_docx_text(path):
    doc = Document(path)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_pdf_text(path):
    text = ""
    with open(path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def process_attachment(filename, payload):
    ext = filename.lower().split('.')[-1]
    with tempfile.NamedTemporaryFile(delete=False, suffix='.'+ext, mode='wb') as tmp:
        tmp.write(payload)
        tmp_path = tmp.name

    try:
        if ext == 'docx':
            return extract_docx_text(tmp_path)
        elif ext == 'pdf':
            return extract_pdf_text(tmp_path)
        elif ext in ['xlsx', 'xls']:
            df = pd.read_excel(tmp_path, dtype=str)
            return df.fillna("").to_string()
        elif ext == 'csv':
            df = pd.read_csv(tmp_path, dtype=str)
            return df.fillna("").to_string()
        elif ext in ['htm', 'html']:
            with open(tmp_path, 'r', encoding='utf-8', errors='ignore') as f:
                soup = BeautifulSoup(f, 'html.parser')
                return soup.get_text(separator='\n')
        elif ext == 'rtf':
            with open(tmp_path, 'r', encoding='utf-8', errors='ignore') as f:
                raw = f.read()
                return re.sub(r'{\\[^{}]+}|\\[a-z]+\d* ?|[{}]', '', raw)
        elif ext == 'txt':
            with open(tmp_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        else:
            return ""
    except Exception:
        return ""

def process_eml(eml_path, attachment_names):
    with open(eml_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)
    subject = msg['subject'] or ""
    body = ""
    attachments = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body += part.get_content()
            elif part.get_filename():
                filename = part.get_filename()
                if filename in attachment_names:
                    payload = part.get_payload(decode=True)
                    attachments.append((filename, process_attachment(filename, payload)))
    else:
        body = msg.get_content()
    return subject, body, attachments

def process_msg(msg_path, attachment_names=None):
    msg = extract_msg.Message(msg_path)
    subject = msg.subject or ""
    body = msg.body or ""
    attachments = []
    for att in msg.attachments:
        att_name = att.longFilename or att.shortFilename
        if not att_name:
            continue
        payload = att.data
        attachments.append((att_name, payload))
    return subject, body, attachments

def scan_text(text, dlp_terms, smart_opts):
    results = []
    # Dictionary terms
    if dlp_terms:
        for term, category in dlp_terms:
            results.append((f"• {term}   [{category}]", "dict"))
    # Smart identifiers
    if "ssn" in smart_opts:
        for ssn in find_ssn(text):
            results.append((f"• {ssn}   [SSN]", "ssn"))
    if "cc" in smart_opts:
        for cc in find_credit_cards(text):
            results.append((f"• {cc}   [Credit Card]", "cc"))
    if "usdl" in smart_opts:
        for lic in find_us_driver_license(text):
            results.append((f"• {lic}   [US Driver License]", "usdl"))
    return results

def log_error(message):
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "dlp_errors.log")
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

def main():
    print("======")
    print(Fore.RED + Style.BRIGHT + "Privacy Warning / Advertencia de Privacidad")
    print(Fore.YELLOW +
        "⚠️ This tool scans sensitive data.\nPlease keep the output private and secure.\n"
        "⚠️ Esta herramienta analiza datos sensibles. Por favor, mantenga el contenido en privado.")
    print("======")
    att_names_input = input(f"Please copy the name of the {Fore.GREEN}attachments (comma-separated){Style.RESET_ALL}:\n  e.g. [email.eml, email.pdf, email.docx, email.txt, email.zip]\n> ").strip()
    att_names = [x.strip() for x in att_names_input.split(",") if x.strip()]
    if not att_names:
        print("No attachments provided. Exiting.")
        sys.exit(1)

    # Prepend current working directory + attachments folder to each filename
    attachments_folder = os.path.join(os.getcwd(), "attachments")
    att_names_full = [os.path.join(attachments_folder, x) for x in att_names]

    # 2. Ask about dictionary usage
    dict_terms = []
    dict_categories = []
    dlp_dict = {}
    selected_terms = None
    flat_dict = {}
    run_smart = False
    smart_opts = []  # Ensure smart_opts is always defined
    dict_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dlp_terms.json")
    if not os.path.isfile(dict_path):
        print("Dictionary file not found.")
        sys.exit(1)
    dlp_dict = load_dlp_dict(dict_path)
    use_dict = input(f"Run against the entire {Fore.CYAN}dictionary{Style.RESET_ALL}? (y/N): ").strip().lower()
    if use_dict in ["y", "yes"]:
        flat_dict = flatten_terms(dlp_dict)
        selected_terms = None
        run_smart = True
    else:
        # Ask about specific dictionary term/category
        use_specific = input(f"Run against a specific {Fore.CYAN}dictionary category{Style.RESET_ALL}? (y/N): ").strip().lower()
        if use_specific in ["y", "yes"]:
            categories = list(dlp_dict.keys())
            print(f"Please choose the specific {Fore.RED}term/category{Style.RESET_ALL}:")
            # Print categories in 3 columns, nicely aligned
            col_width = 28
            for i in range(0, len(categories), 3):
                row = "".join(f"[{i+j+1}] {categories[i+j]:<{col_width}}" if i+j < len(categories) else "" for j in range(3))
                print(row.rstrip())
            cat_choice = input(f"Enter the number(s) of the {Fore.CYAN}category{Style.RESET_ALL} to use (comma-separated): ").strip()
            cat_indices = [int(i)-1 for i in cat_choice.split(",") if i.strip().isdigit() and 0 < int(i) <= len(categories)]
            if not cat_indices:
                print("No valid category selected. Exiting.")
                sys.exit(1)
            selected_terms = []
            for idx in cat_indices:
                selected_terms.extend(dlp_dict[categories[idx]])
            flat_dict = flatten_terms(dlp_dict)
            run_smart = True
        else:
            flat_dict = {}
            selected_terms = None
            run_smart = False

    # Always prompt for smart identifiers if any dictionary is selected or if user chose to skip dictionary
    smart_opts = []
    ask_smart = False
    if (use_dict in ["y", "yes"]) or (use_specific in ["y", "yes"]) or (not flat_dict):
        use_smart = input(f"Run it against any of these 3 {Fore.MAGENTA}Smart Identifiers{Style.RESET_ALL} (SSN, CC, USDL)? (y/N): ").strip().lower()
        if use_smart in ["y", "yes"]:
            ask_smart = True
            print(f"Which {Fore.MAGENTA}Smart Identifier(s){Style.RESET_ALL}?")
            print(f"[1] {Fore.RED}SSN{Style.RESET_ALL} (formatted and unformatted)")
            print(f"[2] {Fore.GREEN}CC{Style.RESET_ALL}")
            print(f"[3] {Fore.BLUE}USDL{Style.RESET_ALL}")
            print(f"[4] {Fore.MAGENTA}The 3 above{Style.RESET_ALL}")
            smart_choice = input(f"Enter number(s) (comma-separated): ").strip()
            if "4" in smart_choice:
                smart_opts = ["ssn", "cc", "usdl"]
            else:
                if "1" in smart_choice:
                    smart_opts.append("ssn")
                if "2" in smart_choice:
                    smart_opts.append("cc")
                if "3" in smart_choice:
                    smart_opts.append("usdl")

    # 5. Scan each file
    for fname, fname_full in zip(att_names, att_names_full):
        print(f"\nAttachment name - {fname}")
        if not os.path.isfile(fname_full):
            print(Fore.RED + f"\t[File not found]")
            log_error(f"File not found: {fname_full}")
            continue
        ext = fname.lower().split('.')[-1]
        if ext == "eml":
            subject, body, attachments = process_eml(fname_full, att_names_full)
            # Scan body/subject
            dlp_terms = find_dlp_terms(subject + "\n" + body, flat_dict, selected_terms) if flat_dict else []
            results = []
            if flat_dict:
                results += scan_text(subject + "\n" + body, dlp_terms, smart_opts if run_smart else [])
            elif run_smart:
                results += scan_text(subject + "\n" + body, [], smart_opts)
            print("\tBody/subject:")
            if results:
                for res, typ in results:
                    if typ == "dict":
                        print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                    elif typ == "ssn":
                        print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                    elif typ == "cc":
                        print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                    elif typ == "usdl":
                        print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                    else:
                        print("\t\t" + res)
            # Scan attachments inside eml
            if attachments:
                for att_name, att_text in attachments:
                    dlp_terms = find_dlp_terms(att_text, flat_dict, selected_terms) if flat_dict else []
                    att_results = []
                    if flat_dict:
                        att_results += scan_text(att_text, dlp_terms, smart_opts if run_smart else [])
                    elif run_smart:
                        att_results += scan_text(att_text, [], smart_opts)
                    if att_results:
                        print(f"\t{att_name}")
                        for res, typ in att_results:
                            if typ == "dict":
                                print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                            elif typ == "ssn":
                                print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                            elif typ == "cc":
                                print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                            elif typ == "usdl":
                                print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                            else:
                                print("\t\t" + res)
        elif ext == "msg":
            subject, body, attachments = process_msg(fname_full)
            dlp_terms = find_dlp_terms(subject + "\n" + body, flat_dict, selected_terms) if flat_dict else []
            results = []
            if flat_dict:
                results += scan_text(subject + "\n" + body, dlp_terms, smart_opts if run_smart else [])
            elif run_smart:
                results += scan_text(subject + "\n" + body, [], smart_opts)
            print("\tBody/subject:")
            if results:
                for res, typ in results:
                    if typ == "dict":
                        print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                    elif typ == "ssn":
                        print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                    elif typ == "cc":
                        print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                    elif typ == "usdl":
                        print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                    else:
                        print("\t\t" + res)
            if attachments:
                for att_name, payload in attachments:
                    # Save attachment to temp file and scan if supported
                    att_ext = att_name.lower().split('.')[-1]
                    supported_exts = ["docx", "pdf", "xlsx", "xls", "csv", "htm", "html", "rtf", "txt", "eml", "msg", "zip"]
                    if att_ext in supported_exts:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.'+att_ext, mode='wb') as tmp:
                            tmp.write(payload)
                            tmp_path = tmp.name
                        try:
                            if att_ext == "docx":
                                text = extract_docx_text(tmp_path)
                            elif att_ext == "pdf":
                                text = extract_pdf_text(tmp_path)
                            elif att_ext in ["xlsx", "xls"]:
                                df = pd.read_excel(tmp_path, dtype=str)
                                text = df.fillna("").to_string()
                            elif att_ext == "csv":
                                df = pd.read_csv(tmp_path, dtype=str)
                                text = df.fillna("").to_string()
                            elif att_ext in ["htm", "html"]:
                                with open(tmp_path, 'r', encoding='utf-8', errors='ignore') as f:
                                    soup = BeautifulSoup(f, 'html.parser')
                                    text = soup.get_text(separator='\n')
                            elif att_ext == "rtf":
                                with open(tmp_path, 'r', encoding='utf-8', errors='ignore') as f:
                                    raw = f.read()
                                    text = re.sub(r'{\\[^{}]+}|\\[a-z]+\d* ?|[{}]', '', raw)
                            elif att_ext == "txt":
                                with open(tmp_path, 'r', encoding='utf-8', errors='ignore') as f:
                                    text = f.read()
                            else:
                                text = ""
                            dlp_terms = find_dlp_terms(text, flat_dict, selected_terms) if flat_dict else []
                            att_results = []
                            if flat_dict:
                                att_results += scan_text(text, dlp_terms, smart_opts if run_smart else [])
                            elif run_smart:
                                att_results += scan_text(text, [], smart_opts)
                            if att_results:
                                print(f"\t{att_name}")
                                for res, typ in att_results:
                                    if typ == "dict":
                                        print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                                    elif typ == "ssn":
                                        print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                                    elif typ == "cc":
                                        print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                                    elif typ == "usdl":
                                        print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                                    else:
                                        print("\t\t" + res)
                        finally:
                            os.remove(tmp_path)
        elif ext == "zip":
            # Ask user if they want to scan inside zip
            scan_zip = input(f"{Fore.CYAN}Scan inside zip file {fname}? (y/N): {Style.RESET_ALL}").strip().lower()
            if scan_zip not in ["y", "yes"]:
                print(Fore.YELLOW + f"\tSkipping zip file {fname}.")
                continue
            with tempfile.TemporaryDirectory() as tmpdir:
                try:
                    with zipfile.ZipFile(fname_full, 'r') as zip_ref:
                        zip_ref.extractall(tmpdir)
                    # Recursively scan extracted files
                    for root, dirs, files in os.walk(tmpdir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            rel_name = os.path.relpath(file_path, tmpdir)
                            ext2 = file.lower().split('.')[-1]
                            supported_exts = ["docx", "pdf", "xlsx", "xls", "csv", "htm", "html", "rtf", "txt", "eml", "msg"]
                            if ext2 in supported_exts:
                                print(f"\t[From zip] {rel_name}")
                                # Recurse: call main scanning logic for this file type
                                # For simplicity, treat as standalone file
                                try:
                                    if ext2 == "eml":
                                        subject, body, attachments = process_eml(file_path, [])
                                        dlp_terms = find_dlp_terms(subject + "\n" + body, flat_dict, selected_terms) if flat_dict else []
                                        results = []
                                        if flat_dict:
                                            results += scan_text(subject + "\n" + body, dlp_terms, smart_opts if run_smart else [])
                                        elif run_smart:
                                            results += scan_text(subject + "\n" + body, [], smart_opts)
                                        print("\tBody/subject:")
                                        if results:
                                            for res, typ in results:
                                                if typ == "dict":
                                                    print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                                                elif typ == "ssn":
                                                    print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                                                elif typ == "cc":
                                                    print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                                                elif typ == "usdl":
                                                    print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                                                else:
                                                    print("\t\t" + res)
                                    elif ext2 == "msg":
                                        subject, body, attachments = process_msg(file_path)
                                        # Explicitly delete the msg object and force garbage collection
                                        import gc
                                        gc.collect()
                                        dlp_terms = find_dlp_terms(subject + "\n" + body, flat_dict, selected_terms) if flat_dict else []
                                        results = []
                                        if flat_dict:
                                            results += scan_text(subject + "\n" + body, dlp_terms, smart_opts if run_smart else [])
                                        elif run_smart:
                                            results += scan_text(subject + "\n" + body, [], smart_opts)
                                        print("\tBody/subject:")
                                        if results:
                                            for res, typ in results:
                                                if typ == "dict":
                                                    print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                                                elif typ == "ssn":
                                                    print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                                                elif typ == "cc":
                                                    print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                                                elif typ == "usdl":
                                                    print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                                                else:
                                                    print("\t\t" + res)
                                    else:
                                        # Standalone file logic
                                        if ext2 == "docx":
                                            text = extract_docx_text(file_path)
                                        elif ext2 == "pdf":
                                            text = extract_pdf_text(file_path)
                                        elif ext2 in ["xlsx", "xls"]:
                                            df = pd.read_excel(file_path, dtype=str)
                                            text = df.fillna("").to_string()
                                        elif ext2 == "csv":
                                            df = pd.read_csv(file_path, dtype=str)
                                            text = df.fillna("").to_string()
                                        elif ext2 in ["htm", "html"]:
                                            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                                                soup = BeautifulSoup(f, 'html.parser')
                                                text = soup.get_text(separator='\n')
                                        elif ext2 == "rtf":
                                            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                                                raw = f.read()
                                                text = re.sub(r'{\\[^{}]+}|\\[a-z]+\d* ?|[{}]', '', raw)
                                        elif ext2 == "txt":
                                            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                                                text = f.read()
                                        else:
                                            text = ""
                                        dlp_terms = find_dlp_terms(text, flat_dict, selected_terms) if flat_dict else []
                                        results = []
                                        if flat_dict:
                                            results += scan_text(text, dlp_terms, smart_opts if run_smart else [])
                                        elif run_smart:
                                            results += scan_text(text, [], smart_opts)
                                        if results:
                                            print(f"\t{rel_name}")
                                            for res, typ in results:
                                                if typ == "dict":
                                                    print(Fore.GREEN + Style.BRIGHT + "\t\t📗 " + res + Style.RESET_ALL)
                                                elif typ == "ssn":
                                                    print(Fore.RED + Style.BRIGHT + "\t\t🆔 " + res + Style.RESET_ALL)
                                                elif typ == "cc":
                                                    print(Fore.GREEN + Style.BRIGHT + "\t\t💳 " + res + Style.RESET_ALL)
                                                elif typ == "usdl":
                                                    print(Fore.BLUE + Style.BRIGHT + "\t\t🪪 " + res + Style.RESET_ALL)
                                                else:
                                                    print("\t\t" + res)
                                except Exception as e:
                                    print(Fore.RED + f"\t[Error scanning {rel_name} in zip: {e}]")
                                    log_error(f"Error scanning {rel_name} in zip: {e}")
                except Exception as e:
                    print(Fore.RED + f"\t[Error extracting zip file: {e}]")
                    log_error(f"Error extracting zip file {fname_full}: {e}")
            continue

if __name__ == "__main__":
    main()