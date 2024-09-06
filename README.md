README - PDF to Excel Conversion with OCR
This Python script converts a PDF document into an Excel file using Adobe PDF Services API. It supports OCR (Optical Character Recognition) for scanned PDFs to extract text.


Features:
Converts PDFs to Excel format (.xlsx).
Supports OCR for extracting text from scanned PDFs.
Logs all operations to a log file with the current date.


Requirements:
adobe-pdfservices-sdk
python-dotenv
Valid Adobe PDF Services credentials.


Usage:
Set up a .env file with your Adobe API credentials.


Run the script using:
python script.py --file <path_to_pdf> --output <output_directory>

If --output is not specified, the Excel file is saved in the input file’s directory.


Environment Variables:
Set the following in your .env file:


PDF_SERVICES_CLIENT_ID=<your_client_id>
PDF_SERVICES_CLIENT_SECRET=<your_client_secret>
This script logs all actions and errors to a log file located in the specified directory.