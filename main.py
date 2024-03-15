import os
import re
import docx
import pypandoc
import win32com.client as win32
from pdf2docx import Converter
from win32com.client import constants
from utils import CFL, create_quiz, extract_format_text, is_option, is_question, process_options, split_options

# Function to convert .doc to .docx
def format_file(file_path: str, del_list: list, selected_options: list) -> tuple:
    """
    Format a document file to DOCX, extract formatted text, and return relevant information.
    The extracted highlights are based on specific formatting rules within the document.
    """
    def convert_to_docx(file_path_convert: str, name: str, del_list: list) -> list:
        global temp_path_docx
        temp_name = f'wteTemp{name}'
        
        # Load the DOCX document using pypandoc
        pypandoc.convert_file(file_path_convert, 'plain', extra_args=['--wrap=none'], outputfile=f'{temp_name}.txt')
        document = docx.Document()
        
        # Read the text from the file and replace soft returns with paragraph marks
        with open(f'{temp_name}.txt', 'r', encoding='utf-8') as file:
            text = file.readlines()
        
        for line in text:
            document.add_paragraph(line)
        
        document.save(f'{temp_name}.docx')
        os.remove(f'{temp_name}.txt')
        temp_path_docx = os.path.abspath(f'{temp_name}.docx')
        del_list.append(temp_path_docx)
        return del_list

    # Function to extract formatted text
    def extract_original_format(file_path: str, selected_options: list) -> list:
        highlights = []
        document = docx.Document(file_path)
        
        #Append the highlighted text
        for paragraph in document.paragraphs:
            highlighted_text = extract_format_text(paragraph, selected_options)
            
            match = re.match(r'^([a-dA-D])\.', highlighted_text)
            if is_option(highlighted_text):
                if "A,B,C,D" in selected_options:
                    highlights.append(CFL(match.group(1)))
                    print(CFL(match.group(1)))
                elif re.match(r'^[a-dA-D]\.(?=\s|$)(?=.+)',highlighted_text):
                    highlights.append(CFL(re.sub(r'^[a-dA-D]\.', '', highlighted_text).strip()))
        return highlights

    # Split the file path into name and extension
    name, ext = os.path.splitext(os.path.basename(file_path))
    abs_file_path = os.path.abspath(file_path)
    
    if ext == ".doc":
        # Convert .doc to .docx
        temp_name = f"wteDocTemp{name}"
        temp_path = os.path.abspath(f"wteDocTemp{name}.docx")
        
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(abs_file_path)
        doc.Activate()
        
        word.ActiveDocument.SaveAs(temp_name, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(False)
        word.Quit()
        
        del_list = convert_to_docx(temp_path, name, del_list)
        del_list.append(temp_path)
        highlights = extract_original_format(temp_path, selected_options)
        
        return temp_path_docx, highlights, del_list

    elif ext == ".docx":
        del_list = convert_to_docx(abs_file_path, name, del_list)
        highlights = extract_original_format(file_path, selected_options)
        return temp_path_docx, highlights, del_list   
    
    elif ext == ".pdf":
        # Convert .pdf to .docx
        new_abs_file_path = os.path.splitext(abs_file_path)[0] + '.docx'
        cv = Converter(abs_file_path)
        cv.convert(new_abs_file_path, start=0, end=None)
        cv.close()
        
        del_list = convert_to_docx(new_abs_file_path, name, del_list)
        highlights = extract_original_format(new_abs_file_path)
        return new_abs_file_path, highlights, del_list
    
    return False, None, None

# Function to process questions and options
def question_create(doc, current_question: str, current_options: list, highlights: list, data: list, platform: str, selected_options: list, question_numbers: int) -> int:
    """
    Process a document to create quiz questions and options based on specific formatting.
    The document structure and formatting rules must align with the processing logic for accurate results.
    """
    def last_question(current_question: str, current_options: list, highlights: list, data: list, platform: str, selected_options: list, question_numbers: int) -> int:
        if current_question and len(current_options) > 0:
            question_numbers += 1
            current_question, current_options = process_options(current_question, current_options, selected_options, question_numbers)
            create_quiz(data, current_question, current_options, highlights, platform, selected_options)
        return question_numbers
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if is_question(text):
            if current_question and len(current_options) > 0:
                current_question, current_options = process_options(current_question, current_options, selected_options, question_numbers)
                question_numbers += 1
                create_quiz(data, current_question, current_options, highlights, platform, selected_options)
            current_question = text
            current_options.clear()  # Clear the options list for the new questions
        elif is_option(text):
            for option in split_options(text):
                current_options.append(option)
    # Process the last question
    question_numbers = last_question(current_question, current_options, highlights, data, platform, selected_options, question_numbers)
    return question_numbers

