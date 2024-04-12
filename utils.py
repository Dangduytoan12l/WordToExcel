import os
import re
import openpyxl
import subprocess
import win32com.client
import pandas as pd
from tkinter.filedialog import askopenfilenames

# Helper function to open a window that specifies a file's path
def open_folder():
    """Opens a file dialog to select multiple files."""
    filepaths = askopenfilenames()
    return filepaths

# Capitalize first letter
def CFL(text: str) -> str:
    """Capitalize the first letter of the string."""
    if text:
        return text[0].upper() + text[1:]
    return text

# Helper function to check whether a text is a question
def is_question(text: str) -> bool:
    """Check if a text is a question."""
    regex_pattern = r"\b(?:Câu|câu|\d+)\b|\b(?:\d+)\."
    regex = re.compile(regex_pattern)
    return bool(regex.match(text))
# Helper function to check if a paragraph starts with an option (A, B, C, D)
def is_option(text: str) -> bool:
    """Check if a text is option"""
    if text.startswith(("A.", "B.", "C.", "D.", "a.", "b.", "c.", "d.","A ", "B ", "C ", "D ", "a ", "b ", "c ", "d ")):        
        return True

# Helper function to split options that are on the same line
def split_options(text: str) -> list:
    """Splits options that are on the same line into a list."""
    if is_option(text):
        return re.split(r'\s+(?=[a-dA-D]\.)', text, flags=re.IGNORECASE)


def extract_format_text(text: str, selected_options: list) -> str:
    """Extracts formatted text (highlighted, bold, underline, italic) if not return empty string."""
    format_text = ""
    
    formatting_conditions = {
        ("Bôi đen",): lambda run: run.font.bold,
        ("In nghiêng",): lambda run: run.font.italic,
        ("Gạch chân",): lambda run: run.font.underline,
        ("Bôi màu",): lambda run: run.font.highlight_color,
    }
    
    for run in text.runs:
        for options, condition in formatting_conditions.items():
            if all(option in selected_options for option in options) and condition(run):
                format_text += run.text
                break
    return format_text

#Get the correct answer index and remove that answer to optimze the performance
def get_correct_answer_index(options: list, highlights: list, contains_ABCD: bool ) -> int:
    """Find and return the index of the correct answer in a list of answer options based on highlighted text.    """

    # Gets the index of the correct answer from options based on highlighted text.
    for index, option_text in enumerate(options):
        try:
            if contains_ABCD:
                if option_text[0] == highlights[0]:
                    highlights.pop(0)
                    return index+1
            else:
                option_text = CFL(re.sub(r'^[a-dA-D]\.', '', option_text).strip())
                if option_text == highlights[0]:
                    highlights.pop(0)
                    return index+1
        except Exception:
            pass
    return None

def create_quiz(data: list, current_question: str, current_options: list, highlights: list, platform: str, selected_options: list) -> None:
    """Create a Quiz Question based on the specified platform."""
    # Creates a question based on the specified platform and adds it to the data list.
    contains_ABCD = False
    if "A,B,C,D" in selected_options:
        contains_ABCD = True
    def quizizz(data: list, current_question: str, current_options: list, highlights: list) -> list:
        # Creates a Quizizz-style question and adds it to the data list.
        data.append({
            'Question Text': current_question,
            'Question Type': "Multiple Choice",
            'Option 1': current_options[0] if len(current_options) > 0 else "",
            'Option 2': current_options[1] if len(current_options) > 1 else "",
            'Option 3': current_options[2] if len(current_options) > 2 else "",
            'Option 4': current_options[3] if len(current_options) > 3 else "",
            'Correct Answer': get_correct_answer_index(current_options, highlights, contains_ABCD),
            'Time in seconds': 30,
        })
        return data

    def kahoot(data: list, current_question: str, current_options: list, highlights: list) -> list:
        # Creates a Kahoot-style question and adds it to the data list.
        data.append({
            'Question': current_question,
            'Answer 1': current_options[0] if len(current_options) > 0 else "",
            'Answer 2': current_options[1] if len(current_options) > 1 else "",
            'Answer 3': current_options[2] if len(current_options) > 2 else "",
            'Answer 4': current_options[3] if len(current_options) > 3 else "",
            'Time limit': 30,
            'Correct Answer': get_correct_answer_index(current_options, highlights, contains_ABCD),
        })
        return data

    def blooket(data: list, current_question: str, current_options: list, highlights: list) -> list:
        # Creates a Blooket-style question and adds it to the data list.
        data.append({
            'Question Text': current_question,
            'Answer 1': current_options[0] if len(current_options) > 0 else "",
            'Answer 2': current_options[1] if len(current_options) > 1 else "",
            'Answer 3': current_options[2] if len(current_options) > 2 else "",
            'Answer 4': current_options[3] if len(current_options) > 3 else "",
            'Time limit': 30,
            'Correct Answer': get_correct_answer_index(current_options, highlights, contains_ABCD),
        })
        return data

    if platform == "Quizizz":
        quizizz(data, current_question, current_options, highlights)
    elif platform == "Kahoot":
        kahoot(data, current_question, current_options, highlights)
    elif platform == "Blooket":
        blooket(data, current_question, current_options, highlights)


def process_formats(current_question: str, current_options: list, selected_options:list, question_number: int) -> None:
    """Process and Format Questions and Answer Options.
    This function processes and formats question text and answer options based on selected formatting options 
    and the question number.
    """
    
    pattern = r'Câu (\d+)'
    match = re.search(pattern, current_question)
    r_match_1 = re.search(r'^Câu (\d+)\.', current_question)
    r_match_2 = re.search(r'^Câu (\d+)\:', current_question)
    current_question = current_question.replace('câu', 'Câu')

    if "Sửa lỗi định dạng" in selected_options:

        # Add a period after the number following "Câu" if it's missing
        if match and not r_match_1 and not r_match_2:
            # Add a period after the number
            current_question = re.sub(pattern, lambda m: f'Câu {m.group(1)}.', current_question, 1)

        # Capitalize the text after "Câu X."
        current_question = re.sub(r'Câu (\d+)\.\s*([a-zA-Z])', lambda match: f'Câu {match.group(1)}. {CFL(match.group(2))}', current_question)
        current_options = [re.sub(r'(^[a-dA-D])\.\s*(.*)', lambda match: f'{CFL(match.group(1))}. {CFL(match.group(2).strip())}', option) for option in current_options]

    if "Xóa chữ 'Câu'" in selected_options:
        current_question = CFL(re.sub(r'^Câu \d+\.', '', current_question).strip())
        current_question = CFL(re.sub(r'^Câu \d+\:', '', current_question).strip())
        current_question = CFL(re.sub(r'^\d+\.', '', current_question).strip())

    if "Xóa chữ 'A,B,C,D'" in selected_options and not "A,B,C,D" in selected_options:
        current_options = [CFL(re.sub(r'^[a-dA-D]\.\s*', '', option).strip()) for option in current_options]

    if "Thêm chữ 'Câu'" in selected_options and not "Câu" in current_question:
        current_question = re.sub(r"(\d+)", r'Câu \1', current_question, 1)
    
    if "Gộp nhiều tệp thành một" in selected_options:
        #Sync the question numbers
        if match:
            current_question = re.sub(pattern, f"Câu {question_number}", current_question)

    return current_question, current_options

def get_explorer_windows(target_path):
    """Find an Explorer window by a given path and bring it to the foreground."""
    shell_windows = win32com.client.Dispatch("Shell.Application").Windows()
    for window in shell_windows:
        # Only consider windows that are instances of File Explorer
        if window.Name == "File Explorer":
            try:
                window_path = window.Document.Folder.Self.Path
                if window_path.lower() == target_path.lower():
                    return True
            except Exception:
                pass
    return None

# Create excel file
def data_frame(data: list, file_path: str, selected_options: list, open_file: bool = True) -> None:
    """Convert a list of data into a DataFrame, optionally random rows, and save it as an Excel file."""
    output_directory = "Output"
    os.makedirs(output_directory, exist_ok=True)
    
    file_name = os.path.splitext(os.path.basename(file_path))[0] + ".xlsx"
    output_path = os.path.join(output_directory, file_name)
    
    df = pd.DataFrame(data)
    
    if "Xáo trộn câu hỏi" in selected_options:
        df = df.sample(frac=1)
    
    df.to_excel(output_path, index=False)
    if "A,B,C,D" in selected_options and "Xóa chữ 'A,B,C,D'" in selected_options:
        case = ("A.", "B.", "C.", "D.", "a.", "b.", "c.", "d.","A ", "B ", "C ", "D ", "a ", "b ", "c ", "d ")
        workbook = openpyxl.load_workbook(output_path)
        worksheet = workbook.active
        for row in worksheet.iter_rows():
            for ceil in row:
                if ceil.value and str(ceil.value).startswith(case):
                    new_value = ceil.value[2:]
                    ceil.value = new_value
        workbook.save(output_path)
    if open_file:
        os.startfile(output_path)
        
def close_excel():
    """Close all instances of excel."""
    # Closes an Excel application if it is open.
    subprocess.call("TASKKILL /F /IM EXCEL.EXE > nul 2>&1", shell=True)
