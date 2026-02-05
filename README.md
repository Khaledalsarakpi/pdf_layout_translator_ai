# PDF Layout Translator AI

## 📄 Overview

**PDF Layout Translator AI** is a powerful Python tool designed to convert PDF documents into editable Word (`.docx`) files while preserving the original layout, images, and equations. It offers two primary modes of operation:
1.  **English Conversion:** Converts the PDF to an English Word document, maintaining exact positioning of text and elements.
2.  **Arabic AI Translation:** Translates the text content to Arabic using advanced AI models (MarianMT), while preserving the original layout (RTL support) and keeping mathematical equations and figures intact as images.

## ✨ Features

*   **Original Layout Preservation:** Uses floating text boxes and images to mimic the exact PDF structure in Word.
*   **AI-Powered Translation:** Utilizes the Helsinki-NLP/opus-mt-en-ar model for high-quality English-to-Arabic translation.
*   **Smart Math Detection:** Automatically detects mathematical equations and symbols to exclude them from translation, preserving scientific accuracy.
*   **Image Extraction:** Extracts figures and equations as high-quality images and places them precisely in the document.
*   **RTL Support:** Full support for Right-to-Left (Arabic) text direction and justification.
*   **Text Cleaning & Merging:** Advanced algorithms to fix OCR errors, merge broken lines into paragraphs, and separate merged words.

## 🛠️ Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd PDF_Layout_Translator_AI
    ```

2.  **Install dependencies:**
    Ensure you have Python 3.8+ installed. Install the required packages using pip:
    ```bash
    pip install -r requirements.txt
    ```

    **Key Dependencies:**
    *   `pymupdf` (PDF processing)
    *   `rapidocr_onnxruntime` (OCR)
    *   `python-docx` (Word generation)
    *   `transformers`, `torch` (AI Translation)
    *   `opencv-python` (Image processing)

## 🚀 Usage

Run the main script to start the interactive interface:

```bash
python main.py
```

Follow the on-screen prompts:
1.  Enter the path to your PDF file (or press Enter to use the default).
2.  (Optional) Enter the number of pages to process.
3.  Select the operation mode:
    *   Type `1` for **English Conversion (Original Layout)**.
    *   Type `2` for **Arabic AI Translation**.

The output `.docx` file will be saved in the same directory as the script.

## 📂 Project Structure

*   `main.py`: The entry point script that handles user input and calls the appropriate modules.
*   `gen_docx_english.py`: Module for converting PDF to English Word doc with layout preservation.
*   `gen_docx_arabic_ai.py`: Module for converting and translating PDF to Arabic Word doc with AI.
*   `requirements.txt`: List of Python dependencies.

## ⚠️ Notes

*   **Performance:** The AI translation process may take some time depending on your hardware (CPU/GPU) and the document length.
*   **Equations:** Equations are treated as images to ensure they display correctly in Word, as OCR often mishandles complex math symbols.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
