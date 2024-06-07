import streamlit as st
from docx import Document
import openai
import pandas as pd

# Set your OpenAI API key here
openai.api_key = st.secrets['my_key']
# Function to extract text from the document
def extract_text_with_paragraphs(docx_path):
    doc = Document(docx_path)
    paragraphs = []
    for para in doc.paragraphs:
        paragraphs.append(para.text)
    return paragraphs

# Function to chunk text into smaller parts based on token count
def chunk_text(paragraphs, max_length=5000):
    chunks = []
    current_chunk = []
    current_length = 0
    
    for paragraph in paragraphs:
        if current_length + len(paragraph.split()) > max_length:
            chunks.append("\n".join(current_chunk))
            current_chunk = []
            current_length = 0
        current_chunk.append(paragraph)
        current_length += len(paragraph.split())
    
    if current_chunk:
        chunks.append("\n".join(current_chunk))
    
    return chunks

# Function to analyze text using GPT-3 and generate results
def analyze_text(chunks):
    responses = []
    for chunk in chunks:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": "Analyze this text to extract responsibilities & actions to be taken (all nesessary),  and responsible department or person, deadline (if exists) in Korean. If more than one responsible department, seperate them by ' && ' sign. the format should be Number|Activity|responsible|deadline. Put 'no deadline' if there is no deadline, don't use those seperating symbols in the context. Each line then should be seperated by semicolon. And avoid shortening the activities."},
                {"role": "user", "content": chunk}
            ],
            max_tokens=4000,
            temperature=0.7,
            n=1
        )
        responses.append(response.choices[0].message.content)
    return responses

# Main function to process the document
def process_document(docx_path):
    # Extract text from the document
    paragraphs = extract_text_with_paragraphs(docx_path)
    
    # Chunk text into smaller parts based on token count
    chunks = chunk_text(paragraphs)
    
    
    
    # Analyze each chunk using GPT-3 and generate results
    results = analyze_text(chunks)
    
    df = pd.DataFrame({
        "results": results,
    }) 
    df = df.stack().str.split(';', expand=True).stack().unstack(-2).reset_index(drop=True)
    df = df.results.str.split('|', expand = True)
    df.columns = ['Number', '책임', '담당자', 'deadline']
    df = df.set_index('Number')
    return df

# Main Streamlit code
st.title("Document Analysis")

# File uploader widget
uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import streamlit as st

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data


if uploaded_file is not None:
    # Process the uploaded document and generate the table
    table_df = process_document(uploaded_file)
    
    df_xlsx = to_excel(table_df)
    st.write(table_df)
    
    st.download_button(label='📥 Download Current Result',
                                data=df_xlsx ,
                                file_name= 'df_test.xlsx')
