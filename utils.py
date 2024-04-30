import os
import re
import pandas as pd
from win32com.client import Dispatch
from tkinter.filedialog import askopenfilenames


# Helper function to open a window that specifies a file's path
def open_folder() -> list:
    """Opens a file dialog to select multiple files."""
    return askopenfilenames()

# Capitalize first letter
def CFL(text: str) -> str:
    """Capitalize the first letter of the string and dont make the rest lower case."""
    return text[0].upper() + text[1:] if text else ""

# Helper function to check if a string is a question
def is_question(text: str) -> bool:
    """Check if a text is a question."""
    return re.match(r'\b(?:Câu|câu|\d+)\b|\b(?:\d+)\.',text)

# Helper function to check if a string is an option.
def is_option(text: str) -> bool:
    """Check if a text is option"""
    return re.match(r'^[a-dA-D][\.:]', text)

# Helper function to split options that are on the same line
def split_options(text: str) -> list:
    """Splits options that are on the same line into a list and remove any redundant whitespace."""
    return re.split(r'\s+(?=[a-dA-D]\.\s+(?![a-dA-D]\.)|[a-dA-D]\.\s+(?![a-dA-D]\.))', re.sub(r" {2,}", " ", text))

def extract_format_text(text: str, selected_options: list) -> str:
    """Extracts formatted text (highlighted, bold, underline, italic) if not return None."""
    formatting_conditions = {
        ("Bôi đen",): lambda run: run.font.bold,
        ("In nghiêng",): lambda run: run.font.italic,
        ("Gạch chân",): lambda run: run.font.underline,
        ("Bôi màu",): lambda run: run.font.highlight_color,
    }

    for run in text.runs:
        for options, condition in formatting_conditions.items():
            if all(option in selected_options for option in options) and condition(run) and is_option(run.text):
                if "A,B,C,D" in selected_options:
                    return run.text[0]
                else:
                    return run.text.strip()

# Get the correct answer index
def get_correct_answer_index(options: list, highlights: list, contains_ABCD: bool ) -> int:
    """
    Find and return the index of the correct answer in a list of answer options based on highlighted text.

    Parameters:
    - options (list): A list of answer options.
    - highlights (list): A list of highlighted text.
    - contains_ABCD (bool): A flag indicating whether the answer options contain the letters A, B, C, D.

    Returns:
    - int: The index of the correct answer in the options list. If no correct answer is found, 0 is returned.
    """
    count = 0
    for index, option_text in enumerate(options):
        try:
            if contains_ABCD and option_text[0] == highlights[0]:
                highlights.pop(0)
                return index+1
            else:
                count+=1
                if option_text.lower() == highlights[0].strip().lower():
                    highlights.pop(0)
                    return index+1
                if count == 4:
                    highlights.pop(0)
                    return 0
        except:
            pass
        
def create_quiz(data: list, current_question: str, current_options: list, highlights: list, platform: str, selected_options: list) -> None:
    """Create a Quiz Question based on the specified platform."""
    contains_ABCD = "A,B,C,D" in selected_options
    def quizizz(data: list, current_question: str, current_options: list, highlights: list) -> list:
        data.append({
            'Question Text': current_question,
            'Question Type': "Multiple Choice",
            'Option 1': current_options[0] if len(current_options) >= 1 else "",
            'Option 2': current_options[1] if len(current_options) >= 2 else "",
            'Option 3': current_options[2] if len(current_options) >= 3 else "",
            'Option 4': current_options[3] if len(current_options) >= 4 else "",
            'Correct Answer': get_correct_answer_index(current_options, highlights, contains_ABCD),
            'Time in seconds': 30
        })
        return data

    def kahoot(data: list, current_question: str, current_options: list, highlights: list) -> list:
        data.append({
            'Question': current_question,
            'Answer 1': current_options[0] if len(current_options) >= 1 else "",
            'Answer 2': current_options[1] if len(current_options) >= 2 else "",
            'Answer 3': current_options[2] if len(current_options) >= 3 else "",
            'Answer 4': current_options[3] if len(current_options) >= 4 else "",
            'Time limit': 30,
            'Correct Answer': get_correct_answer_index(current_options, highlights, contains_ABCD)
        })
        return data

    def blooket(data: list, current_question: str, current_options: list, highlights: list) -> list:
        data.append({
            'Question Text': current_question,
            'Answer 1': current_options[0] if len(current_options) >= 1 else "",
            'Answer 2': current_options[1] if len(current_options) >= 2 else "",
            'Answer 3': current_options[2] if len(current_options) >= 3 else "",
            'Answer 4': current_options[3] if len(current_options) >= 4 else "",
            'Time limit': 30,
            'Correct Answer': get_correct_answer_index(current_options, highlights, contains_ABCD)
        })
        return data

    if platform == "Quizizz":
        quizizz(data, current_question, current_options, highlights)
    elif platform == "Kahoot":
        kahoot(data, current_question, current_options, highlights)
    elif platform == "Blooket":
        blooket(data, current_question, current_options, highlights)

def process_formats(current_question: str, current_options: list, selected_options:list, question_number: int) -> None:
    """
    Process and format questions and answer options based on selected formatting options and the question number.

    Args:
        current_question (str): The current question text.
        current_options (list): The list of current answer options.
        selected_options (list): The list of selected formatting options.
        question_number (int): The question number.

    Returns:
        None: This function does not return anything. It modifies the input parameters in-place.

    Description:
        This function processes and formats the question text and answer options based on the selected formatting options.
        The function performs the following operations:
        1. Adds a period after the number following "Câu" if it is missing.
        2. Capitalizes the text after "Câu X.".
        3. Removes lines containing '[]' from the question text.
        4. Formats answer options by capitalizing the first letter and removing leading whitespace.
        5. Removes the text 'Câu' from the question text if it exists.
        6. Removes the text 'A,B,C,D' from the answer options if it exists.
        7. Adds the text 'Câu' to the question text if it does not exist.
        8. Synchronizes the question numbers if the 'Gộp nhiều file thành một' option is selected.

    Note:
        This function modifies the input parameters in-place.
    """
    pattern = r'^Câu (\d+)'
    match = re.search(pattern, current_question)
    match_with_dot = re.search(r'^Câu (\d+)[\.:]', current_question)

    if "Sửa lỗi định dạng" in selected_options:

        # Add a period after the number following "Câu" if it is  missing
        if match and not match_with_dot:
            # Add a period after the number
            current_question = re.sub(pattern, lambda m: f'Câu {m.group(1)}.', current_question, 1)

        # Capitalize the text after "Câu X."
        current_question = re.sub(r'Câu (\d+)\.\s*([a-zA-Z])', lambda match: f'Câu {match.group(1)}. {CFL(match.group(2))}', current_question)
        current_question = '\n'.join(filter(lambda line: '[]' not in line, current_question.split('\n')))
        current_options = [re.sub(r'(^[a-dA-D])\.\s*(.*)', lambda match: f'{CFL(match.group(1))}. {CFL(match.group(2).strip())}', option) for option in current_options]

    if "Xóa chữ 'Câu'" in selected_options:
        current_question = CFL(re.sub(r'^(Câu \d+\.|Câu \d+\:|\d+\.)', '', current_question).strip())

    if "Xóa chữ 'A,B,C,D'" in selected_options and not "A,B,C,D" in selected_options:
        current_options = [CFL(re.sub(r'^[a-dA-D]\.\s*', '', option).strip()) for option in current_options]

    if "Thêm chữ 'Câu'" in selected_options and "Câu" not in current_question:
        current_question = re.sub(r"(\d+)", r'Câu \1', current_question, 1)
    
    if "Gộp nhiều file thành một" in selected_options:
        #Sync the question numbers
        if match:
            current_question = re.sub(pattern, f"Câu {question_number}", current_question)

    return current_question, current_options

def get_explorer_windows(target_path):
    """
    Find an Explorer window by a given name and bring it to the foreground.

    Parameters:
        target_path (str): The path of the Explorer window to find.

    Returns:
        bool: True if the Explorer window is found and brought to the foreground, False otherwise.
    """
    shell_windows = Dispatch("Shell.Application").Windows()
    for window in shell_windows:
        # Only consider windows that are instances of File Explorer
        if window.Name == "File Explorer":
            try:
                if window.Document.Folder.Self.Path.lower() == target_path.lower():
                    return True
            except Exception:
                pass
    return None

def get_unique_file_path(output_path):
    """
    Generates a unique file path by appending a count to the given output path if it already exists.

    Parameters:
        output_path (str): The original output path.

    Returns:
        str: The unique file path.

    Example:
        >>> get_unique_file_path("output.txt")
        'output (1).txt'
        >>> get_unique_file_path("output (1).txt")
        'output (2).txt'
    """
    count = 1
    name, ext = os.path.splitext(output_path)
    while os.path.exists(output_path):
        output_path = f"{name} ({count}){ext}"
        count += 1
    return output_path

def data_frame(data: list, file_path: str, selected_options: list, open_file: bool = True) -> None:
    """
    Creates a DataFrame from the given data list and saves it as an Excel file.

    Args:
        data (list): A list of dictionaries representing the data to be converted into a DataFrame.
        file_path (str): The path to the input file.
        selected_options (list): A list of selected options.
        open_file (bool, optional): Whether to open the output file after saving. Defaults to True.

    Returns:
        None

    This function creates a DataFrame from the given data list and saves it as an Excel file. 
    It takes the following steps:
    1. Creates the output directory if it doesn't exist.
    2. Generates a unique file name based on the input file path.
    3. Saves the DataFrame as an Excel file in the output directory.
    4. If the "Xáo trộn câu hỏi" option is selected, shuffles the DataFrame.
    5. If the "A,B,C,D" and "Xóa chữ 'A,B,C,D'" options are both selected, removes the leading characters 'A', 'B', 'C', or 'D' followed by a colon or period from each string in the DataFrame.
    6. Saves the DataFrame as an Excel file.
    7. If the `open_file` parameter is True, opens the output file using the default program associated with the file type.
    """
    output_directory = "Output"
    os.makedirs(output_directory, exist_ok=True)
    
    file_name = f"{os.path.splitext(os.path.basename(file_path))[0]}.xlsx"
    output_path = get_unique_file_path(os.path.join(output_directory, file_name))
    
    df = pd.DataFrame(data)

    
    if "Xáo trộn câu hỏi" in selected_options:
        df = df.sample(frac=1)
    
    if "A,B,C,D" in selected_options and "Xóa chữ 'A,B,C,D'" in selected_options:
        df = df.replace(r'^[a-dA-D][\.:]', '', regex=True)
    
    df.to_excel(output_path, index=False)
    
    if open_file:
        os.startfile(output_path)
