PDF to DOCX Converter

A lightweight desktop application built with Python that converts PDF files into editable DOCX documents. Designed with GUI using Tkinter, the tool extracts text from each page, counts total words and pages, and provides real-time progress updates—all while maintaining a responsive interface through background threading.

A. Features

- Select and preview PDF file path
- Convert PDF to DOCX with one click
- Real-time progress updates during conversion
- Displays total word count and page count
- Responsive GUI using multithreading
- Error handling and success notifications

B. Technologies Used

| Library         | Purpose                                      |
|----------------|----------------------------------------------|
| `tkinter`       | GUI creation and user interaction            |
| `filedialog`    | File selection dialog                        |
| `messagebox`    | Popup alerts and error messages              |
| `fitz` (PyMuPDF)| PDF parsing and text extraction              |
| `python-docx`   | DOCX file creation and formatting            |
| `threading`     | Background processing for smooth UI          |

C. Installation

1. **Clone the repository**
   git clone https://github.com/Chippo90/Software-Design-and-Modeling.git
2. Install dependencies
   pip install pymupdf python-docx
3. Run the application
   main.py
   
D. Usage
1. Launch the app.
2. Click Browse PDF to select a file.
3. Click Convert to DOCX to begin conversion.
4. View progress and final word/page count.
5. Find the converted DOCX saved in the same directory.

E. Project Structure


Software-Design-and-Modeling/

├── main.py         
├── README.md            
├── GH1034223 - Chehab Hany's Final Assessment.pdf    
└── UML Diagrams/              
