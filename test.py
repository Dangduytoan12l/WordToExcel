import win32com.client
import win32gui

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
            except Exception as e:
                print(f"Error accessing window's path: {e}")
    return None

# Example usage
if __name__ == "__main__":
    path_to_find = r"C:\Users\Toan\WordToExcel\Output"
    if get_explorer_windows(path_to_find):
        print(f"Found and foregrounded an Explorer window for '{path_to_find}'")
    else:
        print(f"No Explorer window found for '{path_to_find}'.")
