import sys
import os
import openai
import dotenv
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QTextEdit
from PyQt6.QtCore import Qt


try:
    import fitz  
except ImportError:
    raise ImportError("Please install PyMuPDF using 'pip install PyMuPDF'")
try:
    import docx 
except ImportError:
    raise ImportError("Please install python-docx using 'pip install python-docx'")
try:
    from fpdf import FPDF 
except ImportError:
    raise ImportError("Please install fpdf using 'pip install fpdf'")

#API key
dotenv.load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')  # Replace with your key if not using environment variable

def extract_text_from_pdf(file_path):
    """Extracts text from a PDF file using PyMuPDF."""
    text = ""
    try:
        doc = fitz.open(file_path)
        for page in doc:
            text += page.get_text()
    except Exception as e:
        text = f"Error extracting PDF text: {e}"
    return text

def extract_text_from_docx(file_path):
    """Extracts text from a DOCX file using python-docx."""
    text = ""
    try:
        doc = docx.Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        text = f"Error extracting DOCX text: {e}"
    return text

def get_resume_summary(resume_text):
    prompt = (
    "You are an expert HR assistant specializing in resume analysis. Your task is to analyze the following resume "
    "and generate a **professional HR-style summary** that highlights the candidates strengths, experience, and suitability for potential roles.\n\n"
    "**Extract and summarize key details, including:**\n"
    "- **Name and Contact Details**\n"
    "- **Skills, Education, and Certifications**\n"
    "- **Most importantly: Summarize Work Experience in a human-like manner, emphasizing relevant achievements, leadership, problem-solving, and impact.**\n\n"
    "**Instead of just listing responsibilities, provide a concise analysis of the candidates career journey, growth, and key contributions.**\n\n"
    "Now, analyze and summarize the following resume as if you were an HR professional presenting it to a hiring manager:\n\n"
    f"Resume:\n{resume_text}"
)

    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a resume analyzer."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=500,
        )
        summary = response.choices[0].message.content.strip()
    except Exception as e:
        summary = f"Error communicating with OpenAI API: {e}"
    return summary

def save_text_as_pdf(text, file_path):
    """Saves the given text into a PDF file using fpdf."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for line in text.splitlines():
        pdf.multi_cell(0, 10, line)
    pdf.output(file_path)

class ResumeScannerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.extracted_text = ""
        self.summary_text = ""
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Resume Scanner")
        self.setGeometry(100, 100, 600, 500)
        
        layout = QVBoxLayout()
        
        self.label = QLabel("Select a Resume (.pdf or .docx)")
        layout.addWidget(self.label)
        
        self.uploadButton = QPushButton("Browse File")
        self.uploadButton.clicked.connect(self.browseFile)
        layout.addWidget(self.uploadButton)
        
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        layout.addWidget(self.resultText)
        
        self.saveButton = QPushButton("Save Summary")
        self.saveButton.clicked.connect(self.saveSummary)
        self.saveButton.setEnabled(False)
        layout.addWidget(self.saveButton)
        
        self.setLayout(layout)
    
    def browseFile(self):
        filePath, _ = QFileDialog.getOpenFileName(
            self, "Open Resume File", "", "PDF Files (*.pdf);;Word Files (*.docx)"
        )
        if filePath:
            self.label.setText(f"Selected: {filePath}")
            self.processResume(filePath)
    
    def processResume(self, filePath):
        self.resultText.setText("Extracting text from resume...\n")
        ext = os.path.splitext(filePath)[1].lower()
        if ext == ".pdf":
            self.extracted_text = extract_text_from_pdf(filePath)
        elif ext == ".docx":
            self.extracted_text = extract_text_from_docx(filePath)
        else:
            self.resultText.append("Unsupported file format.")
            return
        
        if self.extracted_text.startswith("Error"):
            self.resultText.append(self.extracted_text)
            return
        
        self.resultText.append("Text extraction complete.\n\nSending to OpenAI for analysis...\n")
        QApplication.processEvents()  # Update UI
        
        self.summary_text = get_resume_summary(self.extracted_text)
        self.resultText.append("Summary Received:\n\n" + self.summary_text)
        self.saveButton.setEnabled(True)
    
    def saveSummary(self):
        options = QFileDialog.Options()
        filePath, selectedFilter = QFileDialog.getSaveFileName(
            self, "Save Summary", "", "PDF Files (*.pdf);;Text Files (*.txt)", options=options
        )
        if filePath:
            if selectedFilter == "PDF Files (*.pdf)" or filePath.lower().endswith(".pdf"):
                try:
                    save_text_as_pdf(self.summary_text, filePath)
                except Exception as e:
                    self.resultText.append(f"Error saving PDF: {e}")
            else:
                try:
                    with open(filePath, 'w', encoding='utf-8') as f:
                        f.write(self.summary_text)
                except Exception as e:
                    self.resultText.append(f"Error saving text file: {e}")
            self.resultText.append(f"\nSummary saved to: {filePath}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ResumeScannerApp()
    window.show()
    sys.exit(app.exec())
