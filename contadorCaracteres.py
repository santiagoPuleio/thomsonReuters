import os
from tkinter import Tk, Label, Button, filedialog, StringVar
from docx import Document
import win32com.client as win32

def rename_files_in_folder(folder_path):
    try:
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                new_filename = filename.replace(" ", "_")
                new_file_path = os.path.join(folder_path, new_filename)
                if new_file_path != file_path:
                    os.rename(file_path, new_file_path)
                    print(f'Renamed "{filename}" to "{new_filename}"')
    except Exception as e:
        print(f"Error al renombrar archivos en {folder_path}: {e}")

def count_characters_in_doc(file_path):
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close(False)
        word.Quit()
        return len(text)
    except Exception as e:
        return f"Error: {e}"

def count_characters_in_docx(file_path):
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return len(''.join(full_text))
    except Exception as e:
        return f"Error: {e}"

def convert_and_count_characters_in_folder(folder_path):
    rename_files_in_folder(folder_path)
    total_characters = 0
    file_details = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # Truncate filename if too long
        display_name = filename if len(filename) <= 30 else filename[:27] + '...'
        status.set(f"Analizando: {display_name}")
        root.update_idletasks()

        if filename.endswith('.doc'):
            count = count_characters_in_doc(file_path)
        elif filename.endswith('.docx'):
            count = count_characters_in_docx(file_path)
        else:
            count = "No es un archivo de Word"

        if isinstance(count, int):
            total_characters += count
            file_details.append(f"{filename}: {count} caracteres")
        else:
            file_details.append(f"{filename}: Fallo")

    details = "\n".join(file_details)
    result.set(f'Total de caracteres (con espacios incluidos): {total_characters}\nDetalles:\n{details}')
    status.set("")  # Clear the status message
    root.update_idletasks()

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        convert_and_count_characters_in_folder(folder_selected)

root = Tk()
root.title("Contador de Caracteres")

result = StringVar()
status = StringVar()

Label(root, text="Selecciona una carpeta para contar caracteres:").pack(pady=10)
Button(root, text="Buscar carpeta", command=browse_folder).pack(pady=10)
Label(root, textvariable=status, width=50, anchor='w').pack(pady=5)  # Status label for analysis progress
Label(root, textvariable=result, justify='left').pack(pady=10)

root.mainloop()