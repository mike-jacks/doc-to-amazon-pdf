# DOC TO AMAZON PDF

## Instructions

This python script converts a .docx or .md file to PDF per amazon publishing specs.

### Installation

Download doc_to_amazon_pdf.py and requirements.txt to your computer into a dedicated folder.  
navigate to that folder within terminal.

#### [Optional Create Virtual environment]

Create a virtual environment by running this command in the terminal:  
`python -m venv .venv`

Activate the python virtual environment:  
`source .venv/bin/activate`  
<br>
<br>

### Dependency installation

run in the terminal:  
`pip install -r requirements.txt`

<br><br>

### Usage

`python doc_to_amazon_pdf.py <input_file> <output_file.pdf>`
<br><br>

### Alternative Usage

after installing dependencies globally (not in a virtual environment) run this command:  
`chmod 744 doc_to_amazon_pdf.py`

add doc_to_amazon_pdf.py to a folder in your $PATH (ie ~/.local/bin)  
remove `.py` from the script filename and run `doc_to_amazon_pdf` from anywhere.
