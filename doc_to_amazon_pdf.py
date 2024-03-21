#!/usr/local/bin/python3

from docx import Document
from markdown2 import markdown
from weasyprint import HTML
import sys

def docx_to_html(docx_file) -> str:
    """Converts a .docx file to HTML

    Args:
        docx_file (str): Path to the .docx file

    Returns:
        str: HTML content
    """
    doc = Document(docx_file)
    html = ""
    in_list = False
    list_type = None
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('List Paragraph'):
            if not in_list:
                html, in_list, list_type = start_new_list(html, paragraph, doc)
            else:
                html += f'<li>{paragraph.text}</li>'
        else:
            if in_list:
                html += f'</{list_type}>'
                in_list = False
                list_type = None

            match paragraph.style.name:
                case "Heading 1":
                    html += f'<h1>{paragraph.text}</h1>'
                case "Heading 2":
                    html += f'<h2>{paragraph.text}</h2>'
                case "Heading 3":
                    html += f'<h3>{paragraph.text}</h3>'
                case "Heading 4":
                    html += f'<h4>{paragraph.text}</h4>'
                case "Heading 5":
                    html += f'<h5>{paragraph.text}</h5>'
                case "Heading 6" | "Heading 7" | "Heading 8" | "Heading 9":
                    html += f'<h6>{paragraph.text}</h6>'
                case "Normal":
                    html += f'<p>{paragraph.text}</p>'
                case "List Paragraph":
                    html += f'<li>{paragraph.text}</li>'
    if in_list:
        html += f'</{list_type}>'

    return html

def start_new_list(html, paragraph, doc):
    """Starts a new list

    Args:
        html (str): HTML content
        paragraph (docx.Paragraph): The paragraph
        doc (Document): The docx Document

    Returns:
        html (str): HTML content
        in_list (bool): Whether the list is open
        list_type (str): The type of list (ul or ol)
    """
    in_list = True
    list_type = 'ul'
    html += f'<{list_type}>'
    html += f'<li>{paragraph.text}</li>'
    return html, in_list, list_type
    

def md_to_html(md_file):
    """Converts a .md file to HTML

    Args:
        md_file (str): Path to the .md file

    Returns:
        _type_: _description_
    """
    with open(md_file, "r") as file:
        markdown_text = file.read()
    return markdown(markdown_text)

def create_pdf(html_content, output_name):
    """Creates a PDF from HTML content

    Args:
        html_content (str): HTML content
        output_name (str): Name of the output PDF file
    """
    # Define any sepecific CSS here
    trim_size_width = "6" if (user_input := input("Trim Size Width in inches (default: 6): ")) == "" else user_input
    trim_size_height = "9" if (user_input := input("Trim Size Height in inches (default: 9): ")) == "" else user_input
    bleed = False if (user_input:= input("Enable Bleed (default: False)? (True/False): ")).lower() == "false" else True if user_input.lower() == "true" else False
    inside_gutter = "0.375" if (user_input := input("Insider gutter size (default: 0.375): ")) == "" else user_input
    outside_margin = "0.25" if (user_input := input("Outside margin size (default: 0.25 if bleed is false, 0.375 if bleed is true): ")) and bleed == False else "0.375" if user_input == "" and bleed == True else user_input
    font_size = "11" if (user_input := input("Font Size (Default: 11): ")) == "" else user_input


    css = ('<style>'
            f'@page {{ size: {trim_size_width}in {trim_size_height}in; margin: {outside_margin}in; }}'
            f'@page :right {{ margin-left: {inside_gutter}in; }}'
            f'@page :left {{ margin-right: {inside_gutter}in; }}'
            f'body {{ font-family: Times New Roman; font-size: {font_size}pt; }}'
            'h1, h2 { text-align: center; }'
            'h2 { page-break-before: always }'
            'p { text-indent: 20px; text-align: justify; }'
            '@page { counter-increment: page; }'
            '@page :left { @bottom-left { content: "Page " counter(page); } }'
            '@page :right { @bottom-right { content: "Page " counter(page); } }'
            '</style>')
    html_with_css = css + html_content
    pdf = HTML(string=html_with_css).write_pdf()
    with open(output_name, 'wb') as output_file:
        output_file.write(pdf)

def main():
    """Main function to convert a .docx or .md file to a PDF
    """
    if len(sys.argv) != 3:
        print("Usage: python script.py input_file output_file.pdf")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if input_file.endswith('.docx'):
        html_content = docx_to_html(input_file)
    elif input_file.endswith('.md'):
        html_content = md_to_html(input_file)
    else:
        print("Unsupported file format. Please provide a .docx or a .md file.")
        sys.exit(1)
        
    create_pdf(html_content, output_file)
    print(f"PDF successfully created: {output_file}")


if __name__ == "__main__":
    main()
