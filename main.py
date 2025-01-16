import streamlit as st
import json
from PyPDF2 import PdfReader
from openai import AzureOpenAI
import pandas as pd
import requests
from io import BytesIO
import io
import re
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
from docx import Document  # Import for DOCX processing
from googleapiclient.errors import HttpError
import textract
import pytesseract
from PIL import Image  # Add this import at the top
import os

# Add Tesseract to PATH for Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Adjust path if needed

api_key = st.secrets["azure_openai"]["api_key"]
azure_endpoint = st.secrets["azure_openai"]["azure_endpoint"]

client = AzureOpenAI(
    azure_endpoint=azure_endpoint,
    api_key=api_key,
    api_version="2024-02-01"
)

# Authenticate and initialize the Google API clients

def authenticate_google_api():
    credentials_info = json.loads(st.secrets["google_api"]["credentials"])
    credentials = service_account.Credentials.from_service_account_info(
        credentials_info, scopes=[
            "https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/documents.readonly"]
    )
    drive_service = build("drive", "v3", credentials=credentials)
    docs_service = build("docs", "v1", credentials=credentials)
    return drive_service, docs_service

# Function to extract text from uploaded PDF

def extract_text_from_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    text = "\n".join(page.extract_text() for page in reader.pages)
    return text

# Extract text from Google Drive files

def extract_text_from_google_drive(drive_service, file_id):
    request = drive_service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    file_stream.seek(0)

    # Debugging: Check the size of the downloaded file
    print("Downloaded file size:", file_stream.getbuffer().nbytes)

    # Handle PDF and DOCX formats
    return extract_text_from_pdf(file_stream)

# Extract text from Google Docs

def extract_text_from_docx(docx_file):
    try:
        document = Document(docx_file)
        text = "\n".join([para.text for para in document.paragraphs])
        return text
    except Exception as e:
        raise ValueError(f"Failed to extract text from DOCX: {e}")

def extract_text_from_doc(doc_file):
    try:
        text = textract.process(doc_file, extension='doc').decode('utf-8')
        return text
    except Exception as e:
        raise ValueError(f"Failed to extract text from DOC file: {e}")

def extract_text_from_image(file_stream):
    try:
        image = Image.open(file_stream)
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        raise ValueError(f"Failed to perform OCR on the image: {e}")

def extract_file_id(url):
    """
    Extracts the file ID from a Google Docs URL using regex.
    """
    pattern = r"/d/([a-zA-Z0-9_-]+)"
    match = re.search(pattern, url)
    if match:
        return match.group(1)
    else:
        raise ValueError("Could not extract file ID from URL.")

def process_url(url, drive_service=None, docs_service=None):
    # Handle Google Docs URLs
    if "docs.google.com/document/d/" in url:
        # Extract the file ID from the URL using regex
        try:
            file_id = extract_file_id(url)
        except ValueError as ve:
            raise ValueError(f"Invalid Google Docs URL format: {ve}")
        
        # Fetch the file metadata
        try:
            file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType").execute()
            mime_type = file_metadata.get("mimeType")
            print(f"File ID: {file_id}, MIME Type: {mime_type}")  # Debugging line
        except Exception as e:
            raise ValueError(f"Failed to retrieve file metadata: {e}")
        
        if mime_type == "application/vnd.google-apps.document":
            # Native Google Docs file; export as DOCX
            try:
                request = drive_service.files().export_media(
                    fileId=file_id,
                    mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")  # Debugging line
                file_stream.seek(0)
                print(f"Downloaded DOCX size: {len(file_stream.getvalue())} bytes")  # Debugging line
                return extract_text_from_docx(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to export Google Doc as DOCX: {e}")
        
        elif mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # Already a DOCX file; download directly
            try:
                request = drive_service.files().get_media(fileId=file_id)
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")  # Debugging line
                file_stream.seek(0)
                print(f"Downloaded DOCX size: {len(file_stream.getvalue())} bytes")  # Debugging line
                return extract_text_from_docx(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to download DOCX from Google Drive: {e}")
        
        elif mime_type == "application/msword":
            # DOC file; download directly and extract text using textract
            try:
                request = drive_service.files().get_media(fileId=file_id)
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")  # Debugging line
                file_stream.seek(0)
                print(f"Downloaded DOC size: {len(file_stream.getvalue())} bytes")  # Debugging line
                return extract_text_from_doc(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to download DOC file from Google Drive: {e}")
        
        else:
            raise ValueError(f"Unsupported MIME type: {mime_type}. Only Google Docs files, DOCX, and DOC files can be processed.")
    
    # Handle other Google Drive URLs
    elif "drive.google.com" in url:
        if "open?id=" in url:
            file_id = url.split("id=")[-1].split("&")[0]
        elif "/file/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0]
        else:
            raise ValueError("Invalid Google Drive URL format.")
        
        # Fetch the file metadata
        try:
            file_metadata = drive_service.files().get(fileId=file_id, fields="mimeType").execute()
            mime_type = file_metadata.get("mimeType")
            print(f"File ID: {file_id}, MIME Type: {mime_type}")  # Debugging line
        except Exception as e:
            raise ValueError(f"Failed to retrieve file metadata: {e}")
        
        file_stream = io.BytesIO()
        # Add image handling condition
        if mime_type.startswith('image/'):
            try:
                request = drive_service.files().get_media(fileId=file_id)
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")
                file_stream.seek(0)
                print(f"Downloaded image size: {len(file_stream.getvalue())} bytes")
                return extract_text_from_image(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to download image from Google Drive: {e}")
        
        if mime_type == "application/pdf":
            try:
                request = drive_service.files().get_media(fileId=file_id)
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")  # Debugging line
                file_stream.seek(0)
                print(f"Downloaded PDF size: {len(file_stream.getvalue())} bytes")  # Debugging line
                return extract_text_from_pdf(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to download PDF from Google Drive: {e}")
        elif mime_type == "application/vnd.google-apps.document":
            # Native Google Docs file; export as DOCX
            try:
                request = drive_service.files().export_media(
                    fileId=file_id,
                    mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")  # Debugging line
                file_stream.seek(0)
                print(f"Downloaded DOCX size: {len(file_stream.getvalue())} bytes")  # Debugging line
                return extract_text_from_docx(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to export Google Doc as DOCX: {e}")
        elif mime_type in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword"]:
            # Handle DOCX and DOC files
            try:
                request = drive_service.files().get_media(fileId=file_id)
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download progress: {int(status.progress() * 100)}%")  # Debugging line
                file_stream.seek(0)
                file_size = len(file_stream.getvalue())
                print(f"Downloaded {'DOCX' if mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' else 'DOC'} size: {file_size} bytes")  # Debugging line

                if mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    return extract_text_from_docx(file_stream)
                elif mime_type == "application/msword":
                    return extract_text_from_doc(file_stream)
            except HttpError as e:
                raise ValueError(f"Failed to download DOC/DOCX file from Google Drive: {e}")
            except Exception as e:
                raise ValueError(f"Unexpected error processing DOC/DOCX file: {e}")

        else:
            raise ValueError(f"Unsupported MIME type: {mime_type}. Only PDFs, Google Docs, DOCX, DOC files, and images can be processed.")
    
    # Handle direct PDF URLs
    elif url.endswith(".pdf"):
        try:
            response = requests.get(url)
            if response.status_code != 200:
                raise ValueError(f"Failed to download PDF from URL: {url}")
            pdf_file = BytesIO(response.content)
            print(f"Downloaded PDF size: {len(pdf_file.getvalue())} bytes")  # Debugging line
            return extract_text_from_pdf(pdf_file)
        except Exception as e:
            raise ValueError(f"Failed to download PDF from URL: {e}")
    
        # Handle direct Image URLs
    
    elif url.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif')):
        try:
            response = requests.get(url)
            if response.status_code != 200:
                raise ValueError(f"Failed to download image from URL: {url}")
            image_file = BytesIO(response.content)
            print(f"Downloaded Image size: {len(image_file.getvalue())} bytes")  # Debugging line
            return extract_text_from_image(image_file)
        except Exception as e:
            raise ValueError(f"Failed to process image from URL: {e}")


    else:
        raise ValueError("Unsupported URL format or file type.")

def json_to_table(json_result, resume_link, column_order):
    # Initialize row with placeholders for all columns
    table_data = {"Resume Link/Text": [resume_link]}
    # Extract values directly based on position
    values = list(json_result.values())

    total_conditions = len(column_order)  # Total number of conditions
    met_conditions = 0  # Counter for met conditions

    for idx, key in enumerate(column_order):
        if idx < len(values):  # Ensure the position exists in the JSON
            value_data = values[idx]  # Access data by position
            if value_data.get("value") == 1:  # Check if the condition is met
                met_conditions += 1  # Increment counter for met conditions
        else:
            # Handle missing positions
            value_data = {"value": None, "remarks": "Data missing"}

        # Add the extracted value and remarks to the table
        table_data[f"{key} Value"] = [value_data.get("value")]
        table_data[f"{key} Remarks"] = [value_data.get("remarks")]

    # Add the summary column for met conditions
    summary = f"{met_conditions}/{total_conditions}"  # Format as met/total
    table_data["Total Score"] = [summary]  # Add summary to the table data

    return pd.DataFrame(table_data)

# AI-based evaluation function

def evaluate_with_ai(resume_text, job_description):
    print(job_description)  # Debugging line
    
    prompt = f"""
    You are tasked with evaluating how compatible a job candidate is with a specific job description by comparing their resume to the outlined criteria. 

    ### STRICTLY FOLLOW THESE INSTRUCTIONS:
    1) Evaluate based on the details explicitly mentioned in the resume. Do not assume information.
    2) Handle the following conditions strictly:
      1. **Must Have:** If explicitly mentioned as `NA`, skip evaluating this condition.
      2. **Broader Context:** Use this as reference if available to infer details from the resume.
    3) If the job description asks for `Age` evaluation:
        1. Check if age is mentioned in the resume.
        2. If DOB is not mentioned, infer age from graduation year.
        3. If no age-related details are found, return as `0`.
    4) If the job description asks for `Native Language` evaluation:
        1. Check if the language is explicitly mentioned in the resume.
        2. If not mentioned, infer based on candidate's work or native location.
        3. If details are missing, return as `0`.
    5) If the candidate worked under multiple roles in the same company, combine the duration for all roles.
    6) If the job description asks for Designation:
        **Designation-Specific Instructions**:
        1. You must find an **exact** or **explicit** match for the required designation. For example, if the job requires “Manager,” the resume must explicitly mention “Manager,” “Managerial,” or a clearly equivalent title.
        2. Do NOT assume that sales-related or counseling roles (e.g., Sales Executive, Academic Counselor, Business Development Associate) are automatically managerial unless the resume explicitly states managerial responsibilities or title.
        3. If the role was an internship and the job description specifically requires a full-time position, do NOT consider an internship as fulfilling that requirement.
        4. If a required designation is not stated in the resume, return `"value": 0` with remarks explaining why (e.g., “No exact managerial title mentioned. and mention what other designations the candidate have”).

    - While calculating the experience, Parse dates in **any standard date format**, including but not limited to:
        - DD/MM/YYYY
        - MM/DD/YYYY
        - YYYY-MM-DD

    7) If the job description asks for Work Experience:
        **Work Experience Calculation** (if relevant):
        1. Parse each **full-time** role’s start and end dates (ignore internships, training programs, or WILP unless the job description explicitly allows them).
        2. If an end date is “Present” or “Till date,” calculate experience through today (or note an approximate ongoing duration).
        3. Sum these durations across all **full-time roles only**.
        4. Compare the candidate's total full-time experience with the requirement from the job description:
        5. Do NOT double-count overlapping dates and do NOT consider internship periods, training programs, or WILP as full-time experience (unless the job description explicitly allows it).
        6. If the candidate’s experience includes internships, WILP, or training, explicitly state in the `"remarks"` why these are excluded from full-time experience.
            - Parse dates in **any valid date format**.
            - Handle overlapping dates carefully:
            - If roles overlap, **only count the non-overlapping period**.
            - For partial dates (e.g., MM/YYYY), assume the **1st day of the month**.
            - For single years (e.g., YYYY), assume **January 1st**.
            - Include **all valid full-time durations** explicitly in the `"remarks"` with their parsed start and end dates.
            - If a role is excluded (e.g., internship, training), clearly mention the **reason for exclusion** in `"remarks"`.
            - Ensure each date range is **added sequentially** and non-overlapping durations are **not ignored**.
            - If experience appears short, explicitly list **parsed durations per role** and how the total was calculated.

    8) If the job description specifies a particular industry and a specific designation (e.g., "1+ year of BDM or Managerial experience in Edtech"):
    - Confirm the candidate has a **full-time** role in that industry (Edtech) for at least the stated duration (1+ year).
    - Ensure the exact designation (e.g., BDM, Manager, Team Lead) is clearly stated in the resume. 
        - "Senior BDA" alone does not count as managerial unless explicitly stated (e.g., “led a team”).
    - If the candidate meets these conditions, set "value": 1 with a brief reason. Otherwise, set "value": 0 and note the shortfall (e.g., “Not managerial,” “Less than 1 year,” “Not Edtech,” etc.).

    - Evaluate each criterion (`Must Have`, `Broader Context`) with:
        - `"value"`: `1` if condition is met, otherwise `0`.
        - `"remarks"`: A brief explanation of why the condition was or was not met. If the condition is not met, provide relevant keyword information about the parameter the candidate possesses.
    ### JSON Output Format Example:
    {{
        "age": {{
            "value": 1,
            "remarks": "Candidate meets the age criteria mentioned in the job description."
        }},
        "native_language": {{
            "value": 0,
            "remarks": "Language not explicitly mentioned in resume."
        }},
        "experience": {{
            "value": 1,
            "remarks": "Candidate has sufficient experience in relevant domain."
        }}
    }}

    - If a condition is marked as `Must Have: NA`:
        1. Check if a `Broader Context for Prompt Criteria` exists.
        2. If Broader Context exists, evaluate the candidate based on that.
            - If the condition is met, set `"value": 1` and provide appropriate remarks.
            - If not met, set `"value": 0` and explain why. Also, provide relevant keyword information about the parameter the candidate possesses.

    Return the results strictly in the above JSON format without any additional text or explanations.
    
    Here are the details Inputs:

    **Job Description:**  
    {job_description}

    **Candidate Resume:**  
    {resume_text}


    """
    
    messages = [{"role": "system", "content": prompt}]
    
    # API Call
    response = client.chat.completions.create(
        model="gpt-4o",  # Replace with the correct model
        messages=messages
    )
    
    # Debugging: Print the raw response content
    raw_content = response.choices[0].message.content.strip()
    print("Raw response content:", raw_content)  # Debugging line
    
    # Sanitize the response
    if raw_content.startswith("```json") and raw_content.endswith("```"):
        raw_content = raw_content[8:-3].strip()
    elif raw_content.startswith("```") and raw_content.endswith("```"):
        raw_content = raw_content[3:-3].strip()
    
    # Parse JSON Response
    try:
        result = json.loads(raw_content)
    except json.JSONDecodeError:
        raise ValueError(
            f"The response is not valid JSON. Raw content: {raw_content}"
        )
    
    # Debugging: Print parsed JSON result
    print("Parsed Response:", result)
    
    return result

def convert_jd_to_json(jd_text):
    prompt = f"""
    You are tasked with converting a job description into a structured JSON format. Each parameter in the job description should be represented with:
    - `criteria`: A concise summary of the requirement.
    - `must_have`: The mandatory conditions or requirements (if it's not mandatory, return `NA`).
    - `broader_context`: Detailed steps or logic to evaluate the criteria.

    **Important Notes:**
    - If `Must Have` is mentioned as `NA`, include `"must_have": "NA"` in the JSON output.
    - If `Broader Context` is not specified, return `"broader_context": "NA"`.

    Example Output:
    {{
        "age": {{
            "criteria": "Age should be less than 30 (consider 31 if other parameters match).",
            "must_have": "Age <30. If rest of the parameters match, we can consider 31.",
            "broader_context": "1. Candidate will mention in Resume\n2. If DOB is not mentioned, calculate from the year of graduation\n3. If those two are not mentioned, calculate from the starting year of career\n4. Else give as 0"
        }},
        "native_language": {{
            "criteria": "Native Language or Known Language: Marathi",
            "must_have": "Must explicitly mention Marathi in resume.",
            "broader_context": "1. Candidate will mention in Resume\n2. If not mentioned, infer based on candidate work locations or native location\n3. Else give as missing"
        }}
    }}

    Return the results strictly in the above JSON format without any additional text or explanations.

    Job Description:
    {jd_text}
    """
    
    messages = [{"role": "system", "content": prompt}]
    
    response = client.chat.completions.create(
        model="gpt-4o",  # Replace with the deployment name in Azure
        messages=messages
    )
    
    # Debugging: Print the raw response content
    raw_content = response.choices[0].message.content.strip()
    
    # Sanitize the response
    if raw_content.startswith("```json") and raw_content.endswith("```"):
        raw_content = raw_content[8:-3].strip()
    elif raw_content.startswith("```") and raw_content.endswith("```"):
        raw_content = raw_content[3:-3].strip()
    
    # Parse the response to ensure it's valid JSON
    try:
        result = json.loads(raw_content)
    except json.JSONDecodeError:
        raise ValueError(
            f"The response is not valid JSON. Raw content: {raw_content}"
        )
    
    return result

jd_json = None

# Streamlit app
st.title("AI-Powered Resume Evaluator for Multiple Document Types")
st.write("Upload a CSV file containing URLs of resumes (Drive, Docs, PDFs, PNGs) and provide the job description to evaluate multiple candidates. The column name should be **Resume Links.** ")

# Job description input
st.header("Step 1: Provide Job Description")
job_description = st.text_area(
    "Paste the job description in plain text:", height=300)

if st.button("Convert to JSON"):
    if job_description:
        try:
            # Debugging: Print the job description
            print("Job Description:", job_description)  # Debugging line

            jd_json = convert_jd_to_json(job_description)
            st.subheader("Job Description JSON")
            st.json(jd_json)
            # Store in session state
            st.session_state.jd_json = jd_json
        except Exception as e:
            st.error(f"An error occurred during conversion: {e}")
    else:
        st.error("Please provide a job description.")

# Upload CSV file
st.header("Step 2: Upload CSV File")
csv_file = st.file_uploader(
    "Upload a CSV file containing URLs of resumes:", type=["csv"])

# Initialize Google API clients
drive_service, docs_service = authenticate_google_api()

column_order = None

if st.button("Process Resumes"):
    jd_json = st.session_state.get('jd_json', None)

    if csv_file and job_description and jd_json:
        try:
            df = pd.read_csv(csv_file)
            if "Resume Links" not in df.columns:
                st.error("The uploaded CSV file must have a column named 'Resume Links'.")
            else:
                results = []
                unsupported_resumes = []
                total_resumes = len(df)
                progress_bar = st.progress(0)
                progress_text = st.empty()

                # Get column_order from jd_json
                column_order = list(jd_json.keys())

                for index, row in df.iterrows():
                    resume_url = row["Resume Links"]
                    try:
                        resume_text = process_url(resume_url, drive_service, docs_service)
                        result = evaluate_with_ai(resume_text, jd_json)
                        result_table = json_to_table(result, resume_url, column_order)
                        results.append(result_table)
                    except Exception as resume_e:
                        # Create a DataFrame with NA values for failed processing
                        error_data = {"Resume Link/Text": resume_url}
                        
                        # Add NA values for each column in column_order
                        for key in column_order:
                            error_data[f"{key} Value"] = "NA"
                            error_data[f"{key} Remarks"] = str(resume_e)  # Add error message as remarks
                        
                        # Add Total Score as NA
                        error_data["Total Score"] = "NA"
                        
                        # Create DataFrame with error data and append to results
                        error_df = pd.DataFrame([error_data])
                        results.append(error_df)
                        
                        # Also add to unsupported_resumes for separate tracking
                        unsupported_resumes.append({
                            "Resume Link/Text": resume_url,
                            "Reason": str(resume_e)
                        })
                    
                    # Update progress
                    progress = (index + 1) / total_resumes
                    progress_bar.progress(progress)
                    progress_text.text(f"Processing: {index + 1}/{total_resumes}...")

                # Combine all results into one DataFrame
                if results:
                    final_results = pd.concat(results, ignore_index=True)
                    st.header("Evaluation Results (Including Failed Processes)")
                    st.dataframe(final_results)
                
                # Display unsupported resumes separately
                if unsupported_resumes:
                    st.header("Failed Processing Details")
                    unsupported_df = pd.DataFrame(unsupported_resumes)
                    st.dataframe(unsupported_df)

                # Summary statistics
                processed_count = len(results) - len(unsupported_resumes)
                unsupported_count = len(unsupported_resumes)
                st.subheader(f"Successfully Processed Resumes: {processed_count}")
                st.subheader(f"Failed Processing: {unsupported_count}")

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.error("Please upload a CSV file, provide a valid job description, and convert it to JSON.")
