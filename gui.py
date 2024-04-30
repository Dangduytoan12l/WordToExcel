import os
import docx
import tkinter as tk
from subprocess import Popen
from contextlib import suppress
from utils import open_folder, get_explorer_windows, data_frame
from main import question_create, format_file


def answer_format() -> None:
    """Create a new tkinter window to select the answer format"""

    window_platform = tk.Toplevel(window)  # Use Toplevel for a new window
    window_platform.title("Đáp án")
    window_platform.geometry("300x180")

    answer_formats = ["Bôi màu","Bôi đen", "In nghiêng", "Gạch chân","A,B,C,D"]

    answer_var = tk.StringVar()
    answer_var.set(answer_formats[0])  # Setting default option

    platform_label = tk.Label(window_platform, text="Đáp án được định dạng:", font=("Helvetica", 12))
    platform_label.grid(row=0, column=0, columnspan=3, pady=10)
    ans_checkboxes = {}  # Dictionary to hold checkbox variables
    
    # Create checkboxes for each answer format
    for i, answer_format in enumerate(answer_formats):
        var = tk.BooleanVar()
        ans_checkboxes[answer_format] = var
        checkbox = tk.Checkbutton(window_platform, text=answer_format, variable=var, anchor="w")
        checkbox.grid(row=1 + (i // 3), column=i % 3, pady=10, padx=10, sticky="w")
    ans_checkboxes["Bôi đen"].set(True)

    # Function to handle the selection and close the window
    def on_select_button_click() -> None:
        selected_formats = [format for format, var in ans_checkboxes.items() if var.get()]
        if selected_formats:
            answer_var.set(selected_formats[0])  # Set the selected format
            window_platform.destroy()  # Close the window
            status_label.config(text="Đang xử lý...",fg="black")
            run(ans_checkboxes)  # Proceed with the main processing logic

    select_button = tk.Button(window_platform, text="Select", command=on_select_button_click)
    select_button.grid(row=3, column=1, pady=10)

curr, next = False, False
def update_checkboxes()-> None:
    """Function to make sure that only one checkbox can be selected at the same time."""
    global curr, next
    
    if curr and checkboxes["Xóa chữ 'Câu'"].get():
        checkboxes["Thêm chữ 'Câu'"].set(False)
    if not curr and next and checkboxes["Thêm chữ 'Câu'"].get():
        checkboxes["Xóa chữ 'Câu'"].set(False)
    curr = checkboxes["Thêm chữ 'Câu'"].get()
    next = checkboxes["Xóa chữ 'Câu'"].get()


def run(ans_checkboxes) -> None:
    """Execute the main processing logic for converting Word documents into quiz data."""
    # Step 1: Get selected file paths
    file_paths = open_folder()

    if not file_paths:
        status_label.config(text="Vui lòng chọn ít nhất một file Word", fg="red")
        return

    # Step 2: Get platform and selected options
    status_label.config(text="Đang xử lý...")
    platform = platform_selection.get()
    selected_options = [option for option, var in checkboxes.items() if var.get()]
    selected_options.extend([ans for ans, var in ans_checkboxes.items() if var.get()])
    # Step 3: Initialize data collection
    all_data = []
    del_list = []
    question_numbers = 1

    # Step 4: Process each selected file
    
    for file_path in file_paths:
        data = []
        current_question = ""
        current_options = []

        # Convert .doc to .docx if needed and get the new .docx file path
        path, highlights, del_list = format_file(file_path, del_list, selected_options)
        if path is None:
            status_label.config(text="Lỗi định dạng file, vui lòng chọn file Word!", fg="red")
            break
        
        doc = docx.Document(path)
        question_numbers = question_create(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers)

        if "Gộp nhiều file thành một" not in selected_options:
            question_numbers = 1
            data_frame(data, file_path, selected_options, open_file=True)
        else:
            all_data.extend(data)

    # Step 5: Merge multiple files if selected option is enabled
    if "Gộp nhiều file thành một" in selected_options:
        data_frame(all_data, "Merged_File.xlsx", selected_options, open_file=True)

    # Step 6: Delete temporary files
    for temp_file in del_list:
        os.remove(temp_file)
    output_path = os.path.abspath("Output")
    
    # Step 7: Open output directory
    if not get_explorer_windows(output_path):
        Popen(['explorer', "Output"], stdout=-1, stderr=-1)
    status_label.config(text="Chuyển đổi thành công!", fg="green")

# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter v2.3")
window.geometry("480x300")

# Main frame for organizing widgets
main_frame = tk.Frame(window)
main_frame.pack(pady=20, padx=10)

# Load the logo image
try:
    logo = tk.PhotoImage(file='logo.png')
except tk.TclError:
    logo = tk.PhotoImage(file='Images\logo.png')
window.iconphoto(True, logo)

# Header label
header_label = tk.Label(main_frame, text="Convert Word to Excel", font=("Helvetica", 16))
header_label.grid(row=0, column=0, columnspan=3, pady=10)

# File selection button
file_button = tk.Button(main_frame, text="Select Word Document", command=answer_format)
file_button.grid(row=1, column=0, columnspan=3, pady=10)

# Create a frame for the version label
version_label = tk.Label(main_frame, text="Author: caphefalumi", fg="blue", font=("Open sans", 8))
version_label.grid(row=5, column=2, sticky="e", padx=5, pady=10)

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row=5, column=0, columnspan=3, pady=10, padx=10)  # Center the label using "sticky"

# Platform radio buttons
platform_options = ["Quizizz", "Kahoot", "Blooket"]
platform_selection = tk.StringVar(window)
platform_selection.set(platform_options[0])

# Create radio buttons
platform_quizizz = tk.Radiobutton(main_frame, text="Quizizz", variable=platform_selection, value="Quizizz")
platform_kahoot = tk.Radiobutton(main_frame, text="Kahoot", variable=platform_selection, value="Kahoot")
platform_blooket = tk.Radiobutton(main_frame, text="Blooket", variable=platform_selection, value="Blooket")

# Place the radio buttons side by side
platform_quizizz.grid(row=2, column=0, pady=10, padx=10, sticky="w")
platform_kahoot.grid(row=2, column=1, pady=10, padx=10, sticky="w")
platform_blooket.grid(row=2, column=2, pady=10, padx=10, sticky="w")

# Choice checkboxes
checkbox_options = ["Xóa chữ 'Câu'", "Thêm chữ 'Câu'", "Sửa lỗi định dạng", "Xóa chữ 'A,B,C,D'", "Xáo trộn câu hỏi", "Gộp nhiều file thành một"]
checkboxes = {}


for i, option_text in enumerate(checkbox_options):
    var = tk.BooleanVar()
    checkboxes[option_text] = var
    checkbox = tk.Checkbutton(main_frame, text=option_text, variable=var, anchor="w",command=update_checkboxes)
    checkbox.grid(row=3 + (i // 3), column=i % 3, pady=10, padx=10, sticky="w")

# Set "Sửa lỗi định dạng" checkbox to be always checked
checkboxes["Sửa lỗi định dạng"].set(True)

# Start the GUI application
if __name__ == "__main__":
    try:
        with suppress(KeyboardInterrupt): window.mainloop()
    except:
        [os.remove(os.path.join(os.getcwd(), file)) for file in os.listdir() if file.startswith("wteTemp")]
