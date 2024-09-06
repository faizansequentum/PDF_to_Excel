import logging
import os
from datetime import datetime
import argparse
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.io.cloud_asset import CloudAsset
from adobe.pdfservices.operation.io.stream_asset import StreamAsset
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_ocr_locale import ExportOCRLocale
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

# Get the current date
current_date = datetime.now().strftime("%Y-%m-%d")

# Get the directory where the current Python script is located
current_directory = os.path.dirname(os.path.abspath(__file__))

# Define the log folder path
log_folder = os.path.join(current_directory, 'log')

# Create the 'log' folder if it doesn't exist
os.makedirs(log_folder, exist_ok=True)

# Set the log file path
log_file = os.path.join(log_folder, "log_" + current_date + ".txt")

# Initialize the logger
logging.basicConfig(filename=log_file,
                    filemode='a',
                    format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S',
                    level=logging.INFO)

# read script parameters
parser = argparse.ArgumentParser(description="Processs PDF for Excel Convert")
parser.add_argument("--file", type=str, required=True, help="The path to the pdf file")
parser.add_argument("--output", type=str, required=False, help="The path for the output file (optional)")

args = parser.parse_args()

if args.file is None:
    print("Usage: python process_articles.py <json file name>")
    exit()

FILE_NAME = args.file
if args.output:
    OUTPUT_DIR = args.output
else:
    # Get the absolute path of the file if only the file name is provided
    if not os.path.dirname(FILE_NAME):
        # If no directory is specified in FILE_NAME, use the directory of the Python script
        OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
        FILE_NAME = os.path.join(OUTPUT_DIR, FILE_NAME)  # Combine the script directory with the file name
    else:
        # If a directory is provided, set OUTPUT_DIR to the directory of the input file
        OUTPUT_DIR = os.path.dirname(FILE_NAME)


class ExportPDFToExcelWithOCROption:
    def __init__(self, pdf_path: str, OUTPUT_DIR: str):
        try:
            logging.info("Starting PDF to Excel conversion with OCR")

            # Ensure the provided PDF path is valid
            if not os.path.isfile(pdf_path):
                raise FileNotFoundError(f"The specified PDF file does not exist: {pdf_path}")

            logging.info(f"Processing PDF file: {pdf_path}")

            # Read the PDF file
            with open(pdf_path, 'rb') as file:
                input_stream = file.read()
            logging.info("PDF file read successfully")

            client_id = os.environ.get('PDF_SERVICES_CLIENT_ID')
            client_secret = os.environ.get('PDF_SERVICES_CLIENT_SECRET')

            if not client_id or not client_secret:
                raise ValueError("Environment variables PDF_SERVICES_CLIENT_ID and PDF_SERVICES_CLIENT_SECRET must be set")

            logging.info("Creating ServicePrincipalCredentials instance")
            credentials = ServicePrincipalCredentials(
                client_id=client_id,
                client_secret=client_secret
            )
            logging.info("Credentials created successfully")

            # Create PDF Services instance
            logging.info("Creating PDFServices instance")
            pdf_services = PDFServices(credentials=credentials)
            logging.info("PDFServices instance created successfully")

            # Upload the PDF
            logging.info("Uploading the PDF file to Adobe PDF Services")
            input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.PDF)
            logging.info("PDF file uploaded successfully")

            # Create parameters for the job
            logging.info("Setting up parameters for the export job")
            export_pdf_params = ExportPDFParams(
                target_format=ExportPDFTargetFormat.XLSX,
                ocr_lang=ExportOCRLocale.EN_US
            )
            logging.info("Parameters set successfully")

            # Create and submit the export job
            logging.info("Creating and submitting the export job")
            export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)
            location = pdf_services.submit(export_pdf_job)
            logging.info(f"Job submitted successfully. Location: {location}")

            # Get the job result
            logging.info("Fetching the job result")
            pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
            logging.info("Job result fetched successfully")

            # Get content from the resulting asset
            logging.info("Getting content from the resulting asset")
            result_asset: CloudAsset = pdf_services_response.get_result().get_asset()
            stream_asset: StreamAsset = pdf_services.get_content(result_asset)
            logging.info("Content retrieved successfully")

            # Save the resulting Excel file with the same name as the PDF file
            logging.info("Saving the resulting Excel file")
            output_file_path = self.create_output_file_path(pdf_path, OUTPUT_DIR)
            with open(output_file_path, "wb") as file:
                file.write(stream_asset.get_input_stream())
            logging.info(f"Excel file saved at: {output_file_path}")

        except ValueError as e:
            logging.error(f"Configuration error: {e}")
        except (ServiceApiException, ServiceUsageException, SdkException) as e:
            logging.error(f"Adobe PDF Services API error: {e}")
        except FileNotFoundError as e:
            logging.error(f"File error: {e}")
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}")

    # Generates a string containing a directory structure and file name for the output file
    @staticmethod
    def create_output_file_path(pdf_path: str, output_dir: str) -> str:
        pdf_file_name = os.path.basename(pdf_path)
        pdf_file_name_without_ext = "processed_Excel"
        now = datetime.now()
        time_stamp = now.strftime("%Y-%m-%dT%H-%M-%S")
        output_folder = output_dir
        os.makedirs(output_folder, exist_ok=True)
        return os.path.join(output_folder, f"{pdf_file_name_without_ext}.xlsx")


if __name__ == "__main__":
    # Example usage with a specific PDF file
    specific_pdf_path = FILE_NAME  # pdf file path
    ExportPDFToExcelWithOCROption(specific_pdf_path, OUTPUT_DIR)

