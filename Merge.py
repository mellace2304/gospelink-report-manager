import pandas as pd
import os
import fitz
from docx import Document
from pathlib import Path
import time
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from copy import deepcopy
import pythoncom
import win32com.client
from collections import defaultdict

# Add at top of Merge.py
import logging

def _validate_pdf(path: str) -> tuple[bool, str]:
    """Returns (is_valid, reason). Cheap check before fitz.open."""
    if not path:
        return False, "empty path"
    if not os.path.exists(path):
        return False, "file does not exist"
    try:
        size = os.path.getsize(path)
    except OSError as e:
        return False, f"stat failed: {e}"
    if size == 0:
        return False, "zero bytes (likely failed conversion)"
    if size < 100:
        return False, f"suspiciously small ({size} bytes)"
    return True, ""

class Preacher:
    def __init__(self,pNum,pName):
        self.pNum = pNum
        self.pName = pName
        self.note = ""
        self.report = ""
        self.noteNecessary = False
    def enforce(self):
        self.noteNecessary = True
    def ready(self)->bool:
        return (not self.noteNecessary or self.note!="") and (self.report!="")   
    def getFiles(self):
        files = [self.report]
        if self.noteNecessary:
            files.append(self.note)
        return [file for file in files if file!='']         
    def display(self):
        print()
        print("***Preacher Printing***")
        print("Name:",self.pName)
        print("Preacher Number:",self.pNum)
        print("Report:",self.report)
        print("TY note necessary:",self.noteNecessary)
        print("TY note:", self.note)
    def reasons(self):
        reasons = []
        if (self.noteNecessary and self.note==""):
            reasons.append(f"Missing TY note")
        if (self.report == ""):
            reasons.append(f"Missing report")
        return reasons
        
class Donor:
    
    def __init__(self, pNum = "", pName = "", eNum = "", eName = "", street = "", city = "", state = "", zip = "", email = "",QReport="", aNum = ""):
        self.eNum = eNum
        if "\n" in eName:
            self.eName = eName.split("\n")[0]
            self.additional = eName.split("\n")[1]
        else:
            self.eName = eName
            self.additional = ""
        self.street = street
        self.city = city
        self.state = state
        self.zip = zip
        self.email = email
        self.send_quarterly = (True if QReport == "E-Mail" else False)
        self.aNum = aNum
        self.preachers = {pNum: Preacher(pNum,pName)}
        self.coverLetter = ""
        self.extra_notes = []
    
    def __eq__(self, other):
        if not isinstance(other,Donor):
            return NotImplemented
        return self.aNum == other.aNum
    
    def __hash__(self):
        return hash((self.aNum))
    
    def addPreacher(self,pNum,pName):
        if pNum not in self.preachers:
            self.preachers[pNum] = Preacher(pNum, pName)

    def ready(self):
        for preacher in self.preachers.values():
            if not preacher.ready():
                return False
        if self.coverLetter == "":
            return False
        else:
            return True
        
    def getFiles(self) -> list[str]:
        files = []
        files.append(self.coverLetter)
        for preacher in self.preachers.values():
            files.extend(preacher.getFiles())
        
        for note in self.extra_notes:
            files.append(note)
        
        return [file for file in files if file!='']
    
    def display(self):
        print()
        print("========Donor Printing========")
        print("Envelope Number:",self.eNum)
        print("Envelope Name:",self.eName)
        print("Street:",self.street)
        print("City:",self.city)
        print("State:",self.state)
        print("ZIP Code:",self.zip)
        print("Email Address:",self.email)
        print("Account Number:",self.aNum)
        print("Coverletter:", self.coverLetter)
        print("Extra Notes/Reports:", ", ".join(self.extra_notes))
        print("Additional:",self.additional)
        for preacher in self.preachers.values():
            preacher.display()
    
    def reasons(self):
        reasons = {}
        if self.coverLetter=="":
            reasons[f"D#{self.eNum} (A#{self.aNum})"] = "Missing coversheet"
        for preacher in self.preachers.values():
            if not preacher.ready():
                reasons[preacher.pNum] = preacher.reasons()
        return reasons

def getDonors(file:str=".\\Data\\Spreadsheets\\2026-01-22 Cover Sheet Data.xlsx") -> dict[Donor]:
    coverSheetDF = pd.read_excel(file)
    dropColumns = ['Type', 'Status', 'Note', 'Sponsorship Amount', 'Pledge','Donor/Preacher Combo']
    coverSheetDF = coverSheetDF.dropna(subset=['Account Number'])
    coverSheetDF["Account Number"] = coverSheetDF["Account Number"].astype(int).astype(str) 
    coverSheetDF = (coverSheetDF.dropna(subset=['Account Number'])
                    .astype(str)
                    .drop(columns=dropColumns)
                    .fillna("")
                    .rename(columns={
                        "Preacher Number": "pNum",
                        "Preacher Name": "pName",
                        "Envelope Number": "eNum",
                        "Envelope Name": "eName",
                        "Primary Street": "street",
                        "Primary City": "city",
                        "Primary State": "state",
                        "Primary ZIP Code": "zip",
                        "Primary Email Address": "email",
                        "Account Number": "aNum",
                        'Send Quarterly Report Via': "QReport",
                    }).astype(str)
                    
                    )
    coverSheetDF["street"] = coverSheetDF["street"].str.replace("\n",' ')

    donors = {}
    for _, row in coverSheetDF.iterrows():
        row = row.fillna("").to_dict()
        aNum = row["aNum"]

        if aNum not in donors:
            donors[aNum] = Donor(**row)
        else:
            donors[aNum].addPreacher(row["pNum"], row["pName"])
    
    return donors

def findDonor(*, eNum=None, eName=None, aNum=None, donors=None) -> Donor | None:
    if donors is None:
        return None
    for donor in donors.values():
        if (
            (eNum  and donor.eNum  == eNum) or
            (eName and donor.eName == eName) or
            (aNum  and donor.aNum  == aNum)
        ):
            return donor

    return None


def enforceTY(donors:dict[Donor],extraGiftFile:str = "") -> None:
    # Get data from both sheets, rename columns, and concatenate into one DataFrame
    tempsheet1 = pd.read_excel(extraGiftFile,sheet_name="Spons Xtra")
    tempsheet1 = tempsheet1.rename(columns={
        "Actual Name": "eName",
        "Gross Amt": "Amo",
        "Prchr ID": "pNum",
        "Preacher Name": "pName",
        "Explanation": "eNum"
    })[["eNum","eName","Amo","pNum"]].dropna()
    tempsheet1["eNum"] = tempsheet1["eNum"].astype(int)
    tempsheet1["pNum"] = tempsheet1["pNum"].astype(int)

    tempsheet2 = pd.read_excel(extraGiftFile,sheet_name="DGR")
    tempsheet2 = tempsheet2.rename(columns={
        "Actual Name": "eName",
        "Amt to Send": "Amo",
        "FUNDS_NAME": "pNum",
        "First 2": "eNum"
    })[["eNum","eName","Amo","pNum"]].dropna()
    tempsheet2["eNum"] = tempsheet2["eNum"].astype(int)
    tempsheet2["pNum"] = tempsheet2["pNum"].astype(int)


    giftDF = pd.concat([tempsheet1,tempsheet2],ignore_index=True)
    giftDF.columns =["eNum","eName","Amo","pNum"]
    giftDF = giftDF.astype(str)
    
    donors_by_eNum = {d.eNum: d for d in donors.values() if d.eNum}
    sponsorBlacklist = ["21956","20325","1940","1287","1565","1787","3548","3628","3990","5705","7608","604110"]
    
    for _, row in giftDF.iterrows():
        if row["eNum"] in sponsorBlacklist:
             continue
        donor = donors_by_eNum.get(row["eNum"])
        if donor == None:
                print("what the.")
                print(str(row))
                continue #fix this
        preacher = donor.preachers.get(row["pNum"])
        if preacher is not None:
            preacher.enforce()
        
def addReports(donors:dict[Donor],directory:str=".\\Data\\Reports") -> None:
    reports = os.listdir(directory)
    blacklist=["pcf","widow","rdf",".docx","rd"]

    donors_by_eNum = {d.eNum: d for d in donors.values() if d.eNum}
    donors_by_pNum = defaultdict(list)
    for d in donors.values():
        for pNum in d.preachers:
            donors_by_pNum[pNum].append(d)

    for report in reports:

        if any(kw in report for kw in blacklist):
            continue
        
        path = os.path.join(directory, report)

        if "ty" in report:
            pNum, eNum = report[:-4].replace(' ','').split('ty')
            donor = donors_by_eNum.get(eNum)
            if donor is None:
                print("yuck donor,",report) #donor referenced was not found at all                
            else:
                tempPreacher = donor.preachers.get(pNum)
                if tempPreacher is not None:
                    tempPreacher.note = path
                else:
                    donor.extra_notes.append(path)
                    #preacher is not listed on donors donations, still added though
        else:

            pNum = report[:-4]
            for donor in donors_by_pNum.get(pNum, []):
                if donor.preachers[pNum] is not None:
                    donor.preachers[pNum].report = path

def merge(files, output_path: str):
    if not files:
        raise ValueError("merge() called with no files")

    # Validate every input before we start
    bad = []
    for f in files:
        ok, reason = _validate_pdf(f)
        if not ok:
            bad.append(f"  {f}: {reason}")
    if bad:
        raise FileNotFoundError(
            "Cannot merge — invalid input files:\n" + "\n".join(bad)
        )

    # Make sure output dir exists
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    merged_pdf = fitz.open()
    try:
        for file in files:
            pdf = fitz.open(file)
            merged_pdf.insert_pdf(pdf)
            pdf.close()
        merged_pdf.save(output_path)
    finally:
        merged_pdf.close()

    # Verify it actually wrote
    ok, reason = _validate_pdf(output_path)
    if not ok:
        raise IOError(f"Merge appeared to succeed but output is invalid: {reason}")
    
def mergeDonors(donors:list[Donor],output_path):

    for donor in donors:
        files = list(dict.fromkeys(donor.getFiles()))
        path = os.path.join(output_path, f"{'d'+donor.eNum if donor.eNum else 'a'+donor.aNum}{'e' if donor.send_quarterly else ''}.pdf")
        merge(files, path)


def fill_template(input_path, output_path, replacements):
    """
    Replace specified phrases/keywords in a DOCX file with new text.
    Handles cases where placeholders are split across multiple runs.
    Also handles text in text boxes and shapes.
    
    Args:
        input_path (str): Path to the input DOCX file
        output_path (str): Path where the modified DOCX will be saved
        replacements (dict): Dictionary mapping old text to new text
                           Example: {'«Envelope_Name»': 'John Doe'}
    
    Returns:
        str: Path to the output file
    """
    doc = Document(input_path)
    
    def get_paragraph_text(para_element):
        """Get full text from a paragraph element."""
        runs = para_element.findall(qn('w:r'))
        text = ''
        for run in runs:
            t_elements = run.findall(qn('w:t'))
            for t in t_elements:
                if t.text:
                    text += t.text
        return text
    
    def replace_in_paragraph_element(para_element):
        """Replace text in a paragraph element (works with both regular and textbox paragraphs)."""
        full_text = get_paragraph_text(para_element)
        
        # Check if any replacement is needed
        needs_replacement = False
        for old_text in replacements.keys():
            if old_text in full_text:
                needs_replacement = True
                break
        
        if not needs_replacement:
            return
        
        # Perform all replacements
        new_text = full_text
        for old_text, replacement in replacements.items():
            new_text = new_text.replace(old_text, replacement)
        
        if new_text == full_text:
            return
        
        # Get the first run's formatting
        runs = para_element.findall(qn('w:r'))
        first_run_props = None
        if runs:
            first_run_props = runs[0].find(qn('w:rPr'))
        
        # Remove all existing runs
        for run in runs:
            para_element.remove(run)
        
        # Split text by newlines and create runs/breaks
        text_parts = new_text.split('\n')
        for i, part in enumerate(text_parts):
            # Create a new run with replaced text
            r = OxmlElement('w:r')
            
            # Copy formatting from first run if it existed
            if first_run_props is not None:
                rPr = deepcopy(first_run_props)
                r.append(rPr)
            
            # Add the text
            if part:  # Only add text if part is not empty
                t = OxmlElement('w:t')
                t.set(qn('xml:space'), 'preserve')
                t.text = part
                r.append(t)
            
            # Add run to paragraph
            para_element.append(r)
            
            # Add line break if not the last part
            if i < len(text_parts) - 1:
                br_run = OxmlElement('w:r')
                if first_run_props is not None:
                    rPr = deepcopy(first_run_props)
                    br_run.append(rPr)
                br = OxmlElement('w:br')
                br_run.append(br)
                para_element.append(br_run)
    
    def replace_in_paragraph(paragraph):
        """Replace text in a paragraph object."""
        replace_in_paragraph_element(paragraph._element)
    
    # Process all paragraphs in the document body
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph)
    
    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph)
    
    # Process headers and footers
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_in_paragraph(paragraph)
        for paragraph in section.footer.paragraphs:
            replace_in_paragraph(paragraph)
    
    # Process text boxes
    body_element = doc.element.body
    textboxes = body_element.findall('.//' + qn('w:txbxContent'))
    for txbx in textboxes:
        paras = txbx.findall('.//' + qn('w:p'))
        for para_elem in paras:
            replace_in_paragraph_element(para_elem)
    
    # Save the modified document
    doc.save(output_path)


def createCoverLetter(donor:Donor,output_path=".\\Data\\Output\\CoverLetters",template_path=".\\Data\\Spreadsheets\\Cover Sheet Q3-2025.docx"):
    pdf_target  = os.path.join(output_path, f"{donor.aNum}.pdf")
    docx_target = os.path.join(output_path, f"{donor.aNum}.docx")

    # Only skip if a real PDF already exists
    ok, _ = _validate_pdf(pdf_target)
    if ok:
        print(f"{donor.aNum} CL Skipped (PDF exists)")
        return

    # Stale orphan docx from a prior failed run? Nuke it.
    if os.path.exists(docx_target):
        try:
            os.remove(docx_target)
        except OSError:
            pass
    if f"{donor.aNum}.docx" in os.listdir(output_path):
        print(f"{donor.aNum} CL Skipped!")
        return 
    
    data = {
        "«Envelope_Name»": donor.eName,
        "«Additional»": f"{'\n' if donor.additional!= '' else ''}{donor.additional}",
        "«Primary_Street»": donor.street,
        "«Primary_City»": donor.city,
        "«Primary_State»": donor.state,
        "«Primary_ZIP_Code»": donor.zip,
        "«Envelope_Number»": donor.eNum,
    }
    for i in range(1,14):
        if i < len(donor.preachers)+1:
            preacher = list(donor.preachers.values())[i-1]
            data[f"«Preacher_{i}»"] = preacher.pNum+' - '+preacher.pName
        else:
            data[f"«Preacher_{i}»"] = ""
    path = os.path.join(output_path, f"{donor.aNum}.docx")
    
    fill_template(template_path,path,data)
    print(f"{donor.aNum} created!")

def docx_to_pdf(input_dir, output_dir, restart_every=50, pause=0.5, files=None):
    input_dir = Path(input_dir).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if files is None:
        files = sorted(input_dir.glob("*.docx"))
    else:
        files = [Path(f) for f in files]

    # Skip files whose PDFs already exist before touching Word
    pending = [f for f in files if not (output_dir / f"{f.stem}.pdf").exists()]
    if not pending:
        return
    
    wdFormatPDF = 17
    
    pythoncom.CoInitialize()
    word = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        count = 0

        for idx, docx_path in enumerate(files, 1):
            pdf_path = output_dir / f"{docx_path.stem}.pdf"

            if pdf_path.exists():
                continue

            doc = None
            try:
                print(f"[{idx}/{len(files)}] {docx_path.name}")
                doc = word.Documents.Open(str(docx_path))
                doc.SaveAs(str(pdf_path), FileFormat=wdFormatPDF)
                count += 1
            except Exception as e:
                print(f"Failed: {docx_path.name}: {e}")
            finally:
                if doc is not None:
                    doc.Close(False)

            if count >= restart_every:
                word.Quit()
                time.sleep(pause)
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                count = 0

    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
def assignCoverLetters(donors, directory=".\\Data\\Output\\CoverLetters"):
    if not os.path.isdir(directory):
        print(f"[assignCoverLetters] directory missing: {directory}")
        return
    assigned = skipped = 0
    for fname in os.listdir(directory):
        if not fname.lower().endswith(".pdf"):
            continue
        full = os.path.join(directory, fname)
        ok, reason = _validate_pdf(full)
        if not ok:
            print(f"[assignCoverLetters] skipping {fname}: {reason}")
            skipped += 1
            continue
        aNum = fname[:-4]
        if aNum in donors:
            donors[aNum].coverLetter = full
            assigned += 1
    print(f"[assignCoverLetters] assigned={assigned} skipped={skipped}")
            
if __name__ == "__main__":
    output_dir = ".\\Data\\Output" 
    cover_letter_dir = output_dir+"\\CoverLetters"
    coverSheetFile = ".\\Data\\Spreadsheets\\2026-01-22 Cover Sheet Data.xlsx"
    extraGiftFile = ".\\Data\\Spreadsheets\\2026-01-22 Qtr 3 2025 extra gift reports.xlsx"
    donors = getDonors(coverSheetFile)
     
    # for donor in donors.values():
    #     createCoverLetter(donor,cover_letter_dir)
    # docx_to_pdf(".\\Data\\Output\\CoverLetters",".\\Data\\Output\\CoverLetters")
    assignCoverLetters(donors,cover_letter_dir)

    enforceTY(donors,extraGiftFile)


    addReports(donors)

    readyDonors = [donor for donor in donors.values() if donor.ready()]
    
    mergeDonors(readyDonors,".\\Data\\Output\\Reports")

