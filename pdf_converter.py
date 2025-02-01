import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
from PIL import Image

def convert_image_to_pdf():
    image_file = filedialog.askopenfilename(title="Selecciona una imagen", filetypes=[("Archivos de imagen", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")])
    
    if not image_file:
        return
    
    pdf_file = os.path.splitext(image_file)[0] + ".pdf"
    
    try:
        image = Image.open(image_file)
        image.save(pdf_file, "PDF", resolution=100.0)
        messagebox.showinfo("Éxito", f"Imagen convertida a PDF:\n{pdf_file}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir la imagen:\n{str(e)}")

def convert_word_to_pdf():
    word_file = filedialog.askopenfilename(title="Selecciona un archivo de Word", filetypes=[("Archivos de Word", "*.docx;*.doc")])
    
    if not word_file:
        return
    
    pdf_file = os.path.splitext(word_file)[0] + ".pdf"
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(word_file)
        doc.SaveAs(pdf_file, FileFormat=17)
        doc.Close()
        word.Quit()
        messagebox.showinfo("Éxito", f"Archivo de Word convertido a PDF:\n{pdf_file}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir el archivo de Word:\n{str(e)}")

def convert_powerpoint_to_pdf():
    ppt_file = filedialog.askopenfilename(title="Selecciona un archivo de PowerPoint", filetypes=[("Archivos de PowerPoint", "*.pptx;*.ppt")])
    
    if not ppt_file:
        return
    
    pdf_file = os.path.splitext(ppt_file)[0] + ".pdf"
    
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = False
        presentation = powerpoint.Presentations.Open(ppt_file)
        presentation.SaveAs(pdf_file, FileFormat=32)  # 32 es el formato para PDF
        presentation.Close()
        powerpoint.Quit()
        messagebox.showinfo("Éxito", f"Archivo de PowerPoint convertido a PDF:\n{pdf_file}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir el archivo de PowerPoint:\n{str(e)}")

def convert_excel_to_pdf():
    excel_file = filedialog.askopenfilename(title="Selecciona un archivo de Excel", filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    
    if not excel_file:
        return
    
    pdf_file = os.path.splitext(excel_file)[0] + ".pdf"
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_file)
        workbook.ExportAsFixedFormat(0, pdf_file)  # 0 es el formato para PDF
        workbook.Close()
        excel.Quit()
        messagebox.showinfo("Éxito", f"Archivo de Excel convertido a PDF:\n{pdf_file}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir el archivo de Excel:\n{str(e)}")

# Crear la ventana principal
root = tk.Tk()
root.title("Convertidor a PDF")
root.geometry("400x300")

# Crear botones para cada opción de conversión
btn_image = tk.Button(root, text="Convertir Imagen a PDF", command=convert_image_to_pdf)
btn_image.pack(pady=10)

btn_word = tk.Button(root, text="Convertir Word a PDF", command=convert_word_to_pdf)
btn_word.pack(pady=10)

btn_powerpoint = tk.Button(root, text="Convertir PowerPoint a PDF", command=convert_powerpoint_to_pdf)
btn_powerpoint.pack(pady=10)

btn_excel = tk.Button(root, text="Convertir Excel a PDF", command=convert_excel_to_pdf)
btn_excel.pack(pady=10)

# Iniciar el bucle de la interfaz gráfica
root.mainloop()