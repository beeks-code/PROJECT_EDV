""" To extract details from edv Confirmation Page And Save them On a Excell Sheet"""
from pathlib import Path
from pdfminer.high_level import extract_text
import re
import pandas as pd

def solve_name(name):
    first_name, last_name = name.split(", ")
    return f"{last_name} {first_name}"

insert_name=[]
conf_numb=[]
year=[]
digital_sig=[]
# Specify the path to your folder
directory_path = Path("path")

# Loop through each file in the directory
for file_path in directory_path.iterdir():
    if file_path.is_file() and file_path.suffix.lower() == ".pdf":  
        # Extract text from the PDF
        text = extract_text(file_path)
        print(f"Processing PDF: {file_path.name}")
        
        # Extract Entrant Name
        name_pattern = r"(?<=Entrant Name:\s)(.*?)(?=\s*Confirmation Number:)"
        name_match = re.search(name_pattern, text)
        if name_match:
            entrant_name = name_match.group(1)
            name=solve_name(entrant_name)
            insert_name.append(name)
            
        
        # for confrim nub
        confirmation_number_pattern = r"(?<=Confirmation Number:\s)(.*?)(?=\s*Year of Birth:)"
        confirmation_number_match = re.search(confirmation_number_pattern, text)
        if confirmation_number_match:
            confirmation_number = confirmation_number_match.group(1)
            conf_numb.append(confirmation_number)
            
        
        # to extarct year
        year_of_birth_pattern = r"(?<=Year of Birth:\s)(\d{4})"
        year_of_birth_match = re.search(year_of_birth_pattern, text)
        if year_of_birth_match:
            year_of_birth = year_of_birth_match.group(1)
            year.append(year_of_birth)
        
        # Extract Digital Signature
        digital_signature_pattern = r"(?<=Digital Signature:\s)(\S+)"
        digital_signature_match = re.search(digital_signature_pattern, text)
        if digital_signature_match:
            digital_signature = digital_signature_match.group(1)
            digital_sig.append(digital_signature)
detail={
"Name":insert_name,
"Confirmation Number":conf_numb,
"Year":year,
"Digital Signature":digital_sig

    
}            
df=pd.DataFrame(detail)
df.to_excel("Detail.xlsx",index=False)
