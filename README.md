📁 Universal File Converter (Tkinter GUI)

A simple and modern multi-format file converter built with Python and Tkinter. This desktop application allows users to convert files between popular formats with an easy-to-use graphical interface.

✨ Features
    Convert between multiple file formats:
    DOCX → PDF
    PPT → PDF
    JPG → PNG
    PNG → JPG
    PDF → DOCX
    PDF → PPT
Clean and responsive GUI (Tkinter + ttk styling)
Automatic file picker based on conversion type
Multithreaded conversion (no UI freezing)
Progress bar during conversion
Error handling with user-friendly messages

⚙️ Requirements

    Install the required Python libraries:

    pip install -r requirements.txt

    Or manually install:

    pip install pillow docx2pdf pdf2docx comtypes pdf2pptx
🚀 How to Run
Clone the repository:\n
    git clone https://github.com/your-username/universal-file-converter.git\n
    cd universal-file-converter
    Run the application:
       |-python
       |-main.py
📌 Notes
    Windows Recommended (for PPT → PDF conversion using PowerPoint COM interface)
    Make sure:
    Microsoft Word is installed (for DOCX → PDF)
    Microsoft PowerPoint is installed (for PPT → PDF)
🧠 How It Works
Uses different libraries depending on file type:
    docx2pdf → DOCX conversion
    pdf2docx → PDF to Word and vice-versa
    pdf2pptx → PDF to PowerPoint and vice-versa
    jpg2png -> JPG to PNG conversion and vice-versa
    Pillow (PIL) → Image conversions
    comtypes → PowerPoint automation
Runs conversions in a separate thread to keep UI responsive
📂 Project Structure
    ├── app.py              # Main application file
    ├── README.md           # Project documentation
⚠️ Known Issues
    PPT conversion only works on Windows with PowerPoint installed
    Large PDF files may take longer to process
    Some formatting may be lost in PDF → DOCX/PPT conversions
💡 Future Improvements
    Drag & drop file support
    Batch file conversion
    Progress bar
    Dark/Light mode toggle
    More file formats (e.g., Excel, TXT)
🤝 Contributing

    Contributions are welcome!

    Fork the repo
    Create a new branch
    Make your changes
    Submit a Pull Request
📜 License
    This project is licensed under the MIT License.
