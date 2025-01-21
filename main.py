import streamlit as st
import json
from PyPDF2 import PdfReader
from openai import AzureOpenAI
import pandas as pd
import requests
from io import BytesIO
import io
import re
from datetime import datetime  # Add this import

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

today_date = datetime.now().strftime('%Y-%m-%d')

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
    # CV Evaluation Expert

    You are an expert CV Evaluation Officer with extensive experience in talent assessment and recruitment. You are provided with the **Job Description** and **Candidate Resume**, your task is to perform a precise, criteria-based evaluation of candidate resumes against job requirements provided.

    # Inputs
    ```
    - You are provided with the **Job Description** and **Candidate Resume**
      - **Job Description** contains the requirements of a particular job.
      - **Candidate Resume** contains the details of a candidate
    ```

    ## Understanding the Job parameters:
    ```
    - GO through the **Job Description** details provided and understand the parameters of the job role.
    ```

    ## Data Extraction Instructions:
    ```
    Throughly review the candidate's resume and get the required details mentioned in the 'Job parameters'.
    1. Extract all dates (like age, work experience, etc) and put them in this format (YYYY-MM-DD)
    2. Identify the job roles and their durations (consider even if duration is still present)
    3. Extract the candidate's educational qualifications with completion dates (consider even if duration is still present)
    4. List all the skills mentioned in the resume.
    5. Note if any of the infomation (like: age, language, location) is not explicitly mentioned in the candidate resume then evaluate as below:
        - If candidate hasn't mentioned his date explicitily but provided the `date of birth` then calculate the years, else keep it as '0'
        - Extract the speaking language based on the location and work details provided, else return as '0'
        - If the end data of the work experience is still as present then calculate the experience from start date to till `today:{today_date}` and compare the total experience again the parameter.
    ```

    ## Job parameters Evaluations:
    ```
    - Based on the parameters provided in the **Job Description**, evaluate each parameters against the Candidate's resume to check the eligibility for the job role.
    ```
    ## Output Rules
    1. Each parameter evaluation must include:
    - Binary value (1/0), 1 indicating that candidate criteria requirement is met while 0 indicates that candidate has not met the requirement.
        - Example: if the student exerience is 7 months and the required experience is 3 then the binary value should be 1.
    - Evidence-based remarks for the above value
    - If the criteria is not met then mention the current qualification
    2. Handle NA requirements:
    - Skip Must Have evaluation if marked NA
    - Use Broader Context if available
    3. Document all exclusions with reasoning

    ## Output Format Instructions
    ```
    - Provide the output response in the standard JSON object format.
    ```

    ## Restrictions
    ```
    - Do not make any assumption with the information provide and only evaluate if particular mentioned in candidate's resume
    ```

    ## Common Missing Points:
    ```
    - Mentioning as less experience even though the experience is more than the required criteria
    -  Providing irrelevant value in a particular values or remarks
    ```

        Evaluation Process:
        ```
    - Evaluate each parameter clearly, how you are calculating the value and adding the remarks.
        ```
    
    ## Example Response Format:
    ```        
     - Here is an example of the JSON object Format Response for your referrence.
        ```
        {{
            "[parameter_1]": {{
                "value": [value],
                "remarks": "[remarks]"
            }},
            "[parameter_2]":": {{
                "value": [value],
                "remarks": "[remarks]"
            }},
            "[parameter_3]":": {{
                "value": [value],
                "remarks": "[remarks]"
            }}
        }}
        ```

    Note: Do not include any extra content in the output.
    ```

    #MandatoryResponseVerification:
    ```
    Take your time and cross verify below things correcly to avoid inaccurate responses.
    - Verify that you have correctly understand the job parameters mentioned in the **Job Description** provided
    - Verify you have extracted the required data mentioned in the **Job Description** from candidate's resume.
    - Strictly Verify that you have evaluated the the valid value/data for each and every parameter, if not recheck the candidate's resume again and assign the correct values.
    - Strictly recheck the you have correctly mentioned the binary value in the 'value' section especially age, work experience to maintain the accuracy clearly if applicable.
    ```
   
    Here are the details:

    **Job Description:**  
    {job_description}

    **Candidate Resume:**  
    {resume_text}
   
    """
    
    messages = [{"role": "system", "content": prompt}]
    
    # API Call
    response = client.chat.completions.create(
        model="gpt-4o",  # Replace with the correct model
        messages=messages,
        temperature= 0.00001
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

        Job Description:
        {jd_text}

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
                    progress_text.text(f"Processed: {index + 1}/{total_resumes}...")

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
