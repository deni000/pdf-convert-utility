# PDF Utility

This Python script provides several functions to manipulate PDF files. You can convert PDF files to Word documents, PowerPoint presentations, merge multiple PDF files, compress PDF files, and split PDF files into individual pages.

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/deni000/pdf-convert-utility.git
    cd pdf-utility
    ```
2. Create Python vertual environment for this tool:

    ```bash
    python3 -m venv pdf
    source pdf/bin/activate

    // Deactivate the environment
    deactivate 
    ```
3. Install the required Python libraries:

    ```bash
    pip install -r requirements.txt

    or

    pip install pdf2docx pdfplumber PyPDF2 python-pptx
    ```

## Usage

1. Run the script:

    ```bash
    python3 pdfconv.py
    ```

2. Choose the desired function from the list of available options.

3. Follow the prompts to provide input parameters and paths as required.

## Available Functions

1. **PDF to Word**: Convert a PDF file to a Word document.
2. **PDF to PowerPoint**: Convert a PDF file to a PowerPoint presentation.
3. **Merge PDF**: Merge multiple PDF files into a single PDF.
4. **Compress PDF**: Compress the size of a PDF file.
5. **Split PDF**: Split a PDF file into individual pages.

## Dependencies

- `os`
- `PyPDF2`
- `pdfplumber`
- `pdf2docx`
- `pptx`

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Feel free to modify and extend the script according to your requirements. If you encounter any issues or have suggestions for improvement, please open an issue or submit a pull request. Thank you!
