# Import tkinter for GUI creation
import tkinter as tk
# Import filedialog and messagebox for file selection and popup
from tkinter import filedialog, messagebox
# Import fitz (PyMuPDF) for reading PDF text
import fitz
# Import Document from python-docx to create DOCX files
from docx import Document
# Import Thread so the conversion runs in the background
from threading import Thread

# Define my class for our PDF to DOCX converter
class MyPDFConverter:
    # method to set up variables and GUI
    def __init__(self):
        # Store selected PDF file's path
        self.pdf_path = ""
        # Store output DOCX file's path
        self.docx_path = ""
        # Store total page count
        self.total_pages = 0
        # Store total word count
        self.total_words = 0

        # Create main tkinter window
        self.root = tk.Tk()
        # Set the window title
        self.root.title("Chehab's PDF to DOCX Converter")

        # Add the instruction label
        tk.Label(self.root, text="Select PDF File:").pack(pady=5)

        # Add a label to display the selected PDF path
        self.pdf_label = tk.Label(self.root, text="No file selected")
        self.pdf_label.pack(pady=5)

        # Add a button to browse PDF
        tk.Button(self.root, text="Browse PDF File", command=self.browse_pdf).pack(pady=5)

        # Add a button to start conversion
        tk.Button(self.root, text="Convert to DOCX Now", command=self.start_conversion).pack(pady=10)

        # Label to display progress messages
        self.progress_label = tk.Label(self.root, text="")
        self.progress_label.pack(pady=5)

        # Start tkinter main loop
        self.root.mainloop()

    # Function to open the file dialog and select PDF
    def browse_pdf(self):
        # Open the file selection dialog
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        # If a file is chosen, update the label
        if self.pdf_path:
            self.pdf_label.config(text=self.pdf_path)

    # Function to start conversion in the background thread
    def start_conversion(self):
        # If no file is selected, show an error message
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first!")
            return
        # Run conversion function in background to keep GUI responsive
        Thread(target=self.convert_pdf_to_docx).start()

    # Main conversion function
    def convert_pdf_to_docx(self):
        # Create a new Word document
        doc = Document()
        # Open the selected PDF with fitz
        pdf = fitz.open(self.pdf_path)

        # Initialize page counter manually
        page_counter = 0
        # Initialize word counter manually
        total_words = 0

        # Loop through each page of the PDF
        for i, page in enumerate(pdf):
            # Increase page counter
            page_counter += 1

            # Extract text from this page
            text = page.get_text()

            # Word counting logic
            words = text.split()       # split by whitespace
            word_count = len(words)    # count words
            total_words += word_count  # add to the total

            # Add text of the page to the Word file
            doc.add_paragraph(text)

            # Calculate percentage of progress
            percent = ((i + 1) / len(pdf)) * 100
            # Update progress label in GUI
            self.progress_label.config(text=f"Converting... {percent:.2f}%")
            # Refresh GUI to show updated progress
            self.root.update_idletasks()

        # Save final total pages and words
        self.total_pages = page_counter
        self.total_words = total_words

        # Create an output path by replacing .pdf with .docx
        self.docx_path = self.pdf_path.replace(".pdf", ".docx")
        # Save the Word file
        doc.save(self.docx_path)

        # Update GUI to show the completion message
        self.progress_label.config(
            text=f"Conversion is Complete!\nTotal Pages: {self.total_pages}\nTotal Words: {self.total_words}"
        )
        # Show popup message for success
        messagebox.showinfo("Success!", f"DOCX saved at {self.docx_path}")


# Run the program
if __name__ == "__main__":
    MyPDFConverter()
