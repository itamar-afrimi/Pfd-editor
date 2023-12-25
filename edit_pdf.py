import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from functools import partial
import logging
from time import perf_counter
import fitz  # PyMuPDF
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph

# logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s")
def _color_output(msg): return f'\033[1;36m{msg}\033[0m'


def delete_pages(file_path: str, page_to_delete : int = None, range : tuple = None, indexes: [] = None):
    """
    This function Deletes page.s from pdf file according to client input.
    :return:
    """
    t0 = perf_counter()
    logging.info('Start to delete %s', file_path)
    file_handler = fitz.open(file_path)
    output_file_path = file_path.removesuffix(".pdf") + "modified" + ".pdf"
    if page_to_delete != None:
        file_handler.delete_page(page_to_delete)
        file_handler.save(output_file_path)
    elif range != None:
        file_handler.delete_pages(range[0], range[1])
        file_handler.save(output_file_path)
    else:
        file_handler.select(indexes)
        file_handler.save(output_file_path)
    logging.info('Terminated in %.2fs.', perf_counter() - t0)


def convert_pdf_to_doc(input_path : str):
    """

    :param input_path:
    :return:
    """
    from pdf2docx import Converter

    pdf_file = input_path
    docx_file = pdf_file.removesuffix('.pdf') + '.docx'
    # docx_file = 'output.docx'

    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)

    cv.close()

def convert_docx_to_pdf(input_path : str):
    """

    :param input_path:
    :return:
    """
    from docx import Document

    pdf_file = input_path.removesuffix('.docx') + '.pdf'
    doc = Document(input_path)
    pdf = SimpleDocTemplate(pdf_file, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    default_font = styles['Normal'].fontName

    for paragraph in doc.paragraphs:
        ptext = paragraph.text
        if ptext:
            story.append(Paragraph(ptext, styles["Normal"], fontName=default_font))

    pdf.build(story)


def choose_file(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, file_path)
def execute_tool(tool_option, file_path_entry, root):
    def delete_pages_tool():
        def delete():
            filepath = file_path_entry.get()
            try:
                if page_delete_option.get() == "Page Number":
                    delete_pages(filepath, int(page_input.get()))
                elif page_delete_option.get() == "Page Range":
                    range_values = tuple(map(int, page_input.get().split(',')))
                    delete_pages(filepath, page_range=range_values)
                elif page_delete_option.get() == "Indexes":
                    index_list = list(map(int, page_input.get().split(',')))
                    delete_pages(filepath, indexes=index_list)
                messagebox.showinfo("Success", "Page(s) deleted successfully!")
                delete_pages_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

        delete_pages_window = tk.Toplevel(root)
        delete_pages_window.title("Delete Pages")

        page_delete_option = tk.StringVar()
        page_delete_option.set("Page Number")

        option_label = tk.Label(delete_pages_window, text="Select Delete Option:")
        option_label.pack()

        option_menu = tk.OptionMenu(delete_pages_window, page_delete_option, "Page Number", "Page Range", "Indexes")
        option_menu.pack()

        page_label = tk.Label(delete_pages_window, text="Enter Page(s):")
        page_label.pack()

        page_input = tk.Entry(delete_pages_window)
        page_input.pack()

        delete_button = tk.Button(delete_pages_window, text="Delete", command=delete)
        delete_button.pack()

    if tool_option == "Delete Pages":
        delete_pages_tool()
    else:
        filepath = file_path_entry.get()
        try:
            if tool_option == "PDF to DOCX":
                convert_pdf_to_doc(filepath)
                messagebox.showinfo("Success", "Conversion successful!")
            elif tool_option == "DOCX to PDF":
                convert_docx_to_pdf(filepath)
                messagebox.showinfo("Success", "Conversion successful!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main_gui():
    root = tk.Tk()
    root.title("PDF/DOCX Tool Options")

    label = tk.Label(root, text="Select Tool Option:")
    label.pack()

    tool_options = ["Delete Pages", "PDF to DOCX", "DOCX to PDF"]
    for option in tool_options:
        tool_button = tk.Button(root, text=option, command=lambda o=option: execute_tool(o, file_path_entry,root))
        tool_button.pack()

    file_path_label = tk.Label(root, text="File Path:")
    file_path_label.pack()

    file_path_entry = tk.Entry(root)
    file_path_entry.pack()

    choose_file_button = tk.Button(root, text="Choose File", command=lambda: choose_file(file_path_entry))
    choose_file_button.pack()

    root.mainloop()

main_gui()
