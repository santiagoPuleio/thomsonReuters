import os
from tkinter import Tk, Label, Button, filedialog, StringVar
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

def count_characters_and_pages_in_doc(file_path):
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(file_path)
        
        text = doc.Content.Text
        page_count = doc.ComputeStatistics(2) 
        
        for section in doc.Sections:
            for header in section.Headers:
                text += header.Range.Text
            for footer in section.Footers:
                text += footer.Range.Text

        for shape in doc.Shapes:
            if shape.TextFrame.HasText:
                text += shape.TextFrame.TextRange.Text

        for footnote in doc.Footnotes:
            text += footnote.Range.Text
        
        for endnote in doc.Endnotes:
            text += endnote.Range.Text

        doc.Close(False)
        word.Quit()
        return len(text), page_count
    except Exception as e:
        return f"Error: {e}", 0

def count_characters_and_pages_in_docx(file_path):
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(file_path)
        
        text = doc.Content.Text
        page_count = doc.ComputeStatistics(2)
        
        doc.Close(False)
        word.Quit()
        
        return len(text), page_count
    except Exception as e:
        return f"Error: {e}", 0

def convert_and_count_characters_in_folder(folder_path):
    rename_files_in_folder(folder_path)
    total_characters = 0
    total_pages = 0
    file_details = []
    
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        display_name = filename if len(filename) <= 30 else filename[:27] + '...'
        status.set(f"Analizando: {display_name}")
        root.update_idletasks()

        if filename.endswith('.doc') or filename.endswith('.docx'):
            if filename.endswith('.doc'):
                count, pages = count_characters_and_pages_in_doc(file_path)
            else:
                count, pages = count_characters_and_pages_in_docx(file_path)
        else:
            count, pages = "No es un archivo de Word", 0

        if isinstance(count, int):
            total_characters += count
            total_pages += pages
            file_details.append(f"{filename}: {count} caracteres, {pages} páginas")
        else:
            file_details.append(f"{filename}: Fallo")

    if total_pages > 0:
        average_characters_per_page = total_characters / total_pages
        approximate_total_pages = total_characters / average_characters_per_page
    else:
        average_characters_per_page = 0
        approximate_total_pages = 0

    details = "\n".join(file_details)
    result.set(f'Total de caracteres (con espacios incluidos): {total_characters}\n'
               f'Promedio de caracteres por página: {average_characters_per_page:.2f}\n'
               f'Aproximado de total de páginas: {approximate_total_pages:.2f}\n'
               f'Detalles:\n{details}')
    status.set("")
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
Label(root, textvariable=status, width=50, anchor='w').pack(pady=5)
Label(root, textvariable=result, justify='left').pack(pady=10)

root.mainloop()
