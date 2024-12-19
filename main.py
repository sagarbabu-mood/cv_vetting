import streamlit as st
import json
from PyPDF2 import PdfReader
from openai import AzureOpenAI
import pandas as pd
import requests
from io import BytesIO
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

api_key = st.secrets["azure_openai"]["api_key"]
azure_endpoint = st.secrets["azure_openai"]["azure_endpoint"]

client = AzureOpenAI(
    azure_endpoint=azure_endpoint,
    api_key=api_key,
    api_version="2024-02-01"
)

# Authenticate and initialize the Google API clients

def authenticate_google_api():
    credentials_path = "precise-passkey-441905-t2-5e8a96b58501.json"  # Update with your actual file path

    try:
        # Validate the file path
        with open(credentials_path, "r") as f:
            credentials_data = json.load(f)

        # Authenticate using the service account file
        credentials = service_account.Credentials.from_service_account_file(
            credentials_data,
            scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/documents"]
        )

        # Initialize Google API clients
        drive_service = build("drive", "v3", credentials=credentials)
        docs_service = build("docs", "v1", credentials=credentials)

        return drive_service, docs_service
    except FileNotFoundError:
        st.error(f"Credentials file not found at {credentials_path}. Ensure the file exists.")
        raise
    except json.JSONDecodeError as e:
        st.error(f"Error decoding JSON in the credentials file: {e}")
        raise
    except Exception as e:
        st.error(f"Failed to authenticate Google API. Error: {e}")
        raise

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

    # Handle PDF and DOCX formats
    return extract_text_from_pdf(file_stream)

# Extract text from Google Docs


def extract_text_from_google_docs(docs_service, document_id):
    doc = docs_service.documents().get(documentId=document_id).execute()
    content = []
    for element in doc.get("body").get("content", []):
        if "paragraph" in element:
            for text_run in element["paragraph"]["elements"]:
                content.append(text_run.get("textRun", {}).get("content", ""))
    return "".join(content)

# Convert URLs to downloadable format or detect file type


def process_url(url, drive_service, docs_service):
    if "drive.google.com" in url:
        if "open?id=" in url:
            file_id = url.split("id=")[-1].split("&")[0]
        elif "/file/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0]
        else:
            raise ValueError("Invalid Google Drive URL format.")
        return extract_text_from_google_drive(drive_service, file_id)
    elif "docs.google.com" in url:
        document_id = url.split("/d/")[1].split("/")[0]
        return extract_text_from_google_docs(docs_service, document_id)
    elif url.endswith(".pdf"):
        response = requests.get(url)
        pdf_file = BytesIO(response.content)
        return extract_text_from_pdf(pdf_file)
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
    Here are the details:

    Job Description:
    {job_description}

    Candidate Resume:
    {resume_text}

    STRICTLY FOLLOW THE INSTRUCTIONS BELOW:
    Instructions:
    - Evaluate based on the information provided in the resume.
    - If the job description asks for 'Age' evaluation:
        1. Check if the age is directly mentioned in the resume.
        2. If the Date of Birth (DOB) is not mentioned, calculate it from the year of graduation.
        3. If those two are not mentioned, calculate it from the starting year of the career.
        4. If none of these are available, consider as 0.
    - If the job description does not ask for age evaluation, ignore this criterion.

    - If the job description asks for 'Native Language' evaluation:
      1. Check if the language is explicitly mentioned in the resume.
      2. If it is not mentioned, infer the language based on candidate work locations or native location.
      3. If none of these are available, consider as 0.
    - If the job description does not ask for 'Native Language', skip this evaluation.

    - Do not assume the existence of documentation or certificates unless explicitly mentioned in the resume.
    - If the candidate has worked in the same company under multiple roles (e.g., a promotion), consider the combined duration for all roles at that company.
    - Identify if the candidate meets the required total tenure by summing the time spent across all roles within the same organization.
    - For each condition in the job description, evaluate the `condition(s)` provided for each criterion and provide the following:
      - "value": 1 if the condition is met, otherwise 0.
      - "remarks": A brief explanation of why the condition is either met or not met.
    - Do not provide separate responses for individual conditions.
    - 

    Example Output:
    {{
        "condition1": {{
            "value": 1 or 0,
            "remarks": "reason for why shortlisting or reason for why not shortlisting"
        }},
        "condition2": {{
            "value": 1 or 0,
            "remarks": "reason for why shortlisting or reason for why not shortlisting"
        }},
    }}

    Example Output with values:
    {{
        "age": {{
            "value": 1,
            "remarks": "Candidate is under 30, calculated based on year of birth from resume."
        }},
        "native_language": {{
            "value": 1,
            "remarks": "Marathi is mentioned as a known language in the resume."
        }},
    }}

    condition1 & condition2 can be replaced with the actual condition names.

    Return the results strictly in the above JSON format without any additional text or explanations.
    """
    messages = [{"role": "system", "content": prompt}]

    response = client.chat.completions.create(
        model="gpt-4o",  # Replace with the deployment name in Azure
        messages=messages
    )

    # Debugging: Print the raw response content
    # Strip whitespace
    raw_content = response.choices[0].message.content.strip()

    # Debugging: Print the raw content for inspection
    print("Raw response content:", raw_content)  # Debugging line

    # Sanitize the response
    if raw_content.startswith("```json") and raw_content.endswith("```"):
        # Remove the ```json and ``` markers
        raw_content = raw_content[8:-3].strip()
    elif raw_content.startswith("```") and raw_content.endswith("```"):
        # Remove the ``` markers if they are present
        raw_content = raw_content[3:-3].strip()

    # Parse the response to ensure it's valid JSON
    try:
        result = json.loads(raw_content)
    except json.JSONDecodeError:
        raise ValueError(
            f"The response is not valid JSON. Raw content: {raw_content}")

    return result


def convert_jd_to_json(jd_text):
    prompt = f"""
    You are tasked with converting a job description into a structured JSON format. Each parameter in the job description should be represented with:
    - `criteria`: A concise summary of the requirement.
    - `condition(s)`: Detailed steps or logic to evaluate the criteria.

    Job Description:
    {jd_text}

    Example Output:
    {{
        "age": {{
            "criteria": "Age should be less than 30 (consider 31 if other parameters match).",
            "condition(s)": "1. Candidate will mention in Resume\n2. If DOB is not mentioned, calculate from year of graduation\n3. If those two are not mentioned, calculate from starting year of career\n4. Else give as missing"
        }},
        "native_language": {{
            "criteria": "Native Language or Known Language: Marathi",
            "condition(s)": "1. Candidate will mention in Resume\n2. If not mentioned, infer based on candidate work locations or native location\n3. Else give as missing"
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
    raw_content = response.choices[0].message.content
    # Parse the response to ensure it's valid JSON
    try:
        result = json.loads(raw_content)
    except json.JSONDecodeError:
        raise ValueError(
            f"The response is not valid JSON. Raw content: {raw_content}")

    return result


jd_json = None

# Streamlit app
st.title("AI-Powered Resume Evaluator for Multiple Document Types")
st.write("Upload a CSV file containing URLs of resumes (Drive, Docs, PDFs) and provide the job description to evaluate multiple candidates.")

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
    # Retrieve jd_json from session state
    jd_json = st.session_state.get('jd_json', None)

    if csv_file and job_description and jd_json:  # Check if jd_json is defined
        try:
            # Read CSV file
            df = pd.read_csv(csv_file)
            if "Resume Links" not in df.columns:
                st.error(
                    "The uploaded CSV file must have a column named 'Resume Links'.")
            else:
                results = []
                total_resumes = len(df)  # Total number of resumes to process
                progress_bar = st.progress(0)  # Initialize progress bar
                progress_text = st.empty()  # Placeholder for progress text

                for index, row in df.iterrows():
                    resume_url = row["Resume Links"]
                    resume_text = process_url(
                        resume_url, drive_service, docs_service)
                    result = evaluate_with_ai(resume_text, jd_json)

                    # Extract keys from the first response as column_order
                    if column_order is None:
                        column_order = list(result.keys())

                    # Standardize the result to match column_order
                    result_table = json_to_table(
                        result, resume_url, column_order)
                    results.append(result_table)

                    # Update progress bar and text
                    # Calculate progress
                    progress = (index + 1) / total_resumes
                    progress_bar.progress(progress)  # Update progress bar
                    # Update progress text
                    progress_text.text(
                        f"Processing: {index + 1}/{total_resumes}...")

                # Combine all results into one DataFrame
                final_results = pd.concat(results, ignore_index=True)

                # Display the result in table format
                st.header("Evaluation Results")
                st.dataframe(final_results)
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.error(
            "Please upload a CSV file, provide a valid job description, and convert it to JSON.")
