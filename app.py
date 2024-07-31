import os
from dotenv import load_dotenv
import streamlit as st
from PyPDF2 import PdfReader
from langchain.text_splitter import CharacterTextSplitter
from langchain.embeddings.huggingface import HuggingFaceEmbeddings
from langchain.vectorstores import FAISS  # facebook AI similarity search
from langchain.chains.question_answering import load_qa_chain
from langchain import HuggingFaceHub

from pptxtopdf import convert
import comtypes.client
import docx
import json

# Sidebar contents
with st.sidebar:
    st.title('📄🖥️ Automatic form filling App powered by LLM')
    st.markdown('''
    ## About
    This app is an Automatic form filling app built using:
    - [Streamlit](https://streamlit.io/)
    - [Hugging Face Transformers](https://huggingface.co/google/flan-t5-large) for question answering

    Made by Developers of team J.A.R.V.I.S.,
    - Shubh Ranpara,
    - Priyanshu Savla,
    - Ronak Siddhpura,
    - Prem Jobanputra,
    - Pranav Tank.
    ''')

def convert_pptx_to_pdf(base_path, file, file_name):

    input_dir = os.path.join(base_path, file_name)
    output_dir = os.path.join(base_path)

    print(base_path, file_name)

    convert(input_dir, output_dir)

    output_file_name = base_path + "\\" + file_name[:-4] + "pdf"

    file = open(output_file_name, 'rb')
    return file

def convert_docx_to_pdf(base_path, file, file_name):

    word_path = base_path + "\\" + file_name
    pdf_path = base_path + "\\" + file_name[:-4] + "pdf"

    # Load the Word document
    doc = docx.Document(word_path)

    # Create a Word application object
    word = comtypes.client.CreateObject("Word.Application")

    # Get absolute paths for Word document and PDF file
    docx_path = os.path.abspath(word_path)
    pdf_path = os.path.abspath(pdf_path)

    # PDF format code
    pdf_format = 17

    # Make Word application invisible
    word.Visible = False

    try:
        # Open the Word document
        in_file = word.Documents.Open(docx_path)

        # Save the document as PDF
        in_file.SaveAs(pdf_path, FileFormat=pdf_format)

        print("Conversion successful. PDF saved at:", pdf_path)

    except Exception as e:
        print("Error:", e)

    finally:
        # Close the Word document and quit Word application
        if 'in_file' in locals():
            in_file.Close()
        word.Quit()

def main():
    load_dotenv()

    st.header("Upload Your PDF, DOCX, PPTX")

    file = st.file_uploader("Upload your file")

    file_name = ""

    if file:
        print(file.name)
        file_name = str(file.name)

    pdf = file

    # the path of directory where our pdf, docx or pptx is kept.
    base_path = "Enter the path of directory where your pdf or other files are kept."

    if (file_name.__contains__(".pptx")):
        pdf = convert_pptx_to_pdf(base_path, file, file_name)

    elif (file_name.__contains__(".docx")):
        convert_docx_to_pdf(base_path, file, file_name)
        output_file_name = base_path + "\\" + file_name[:-4] + "pdf"
        pdf = open(output_file_name, 'rb')

    if pdf is not None:
        pdf_reader = PdfReader(pdf)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()

        # split ito chuncks
        text_splitter = CharacterTextSplitter(
            separator="\n",
            chunk_size=512,
            chunk_overlap=0,
            length_function=len
        )
        chunks = text_splitter.split_text(text)

        # create embedding
        embeddings = HuggingFaceEmbeddings()

        knowledge_base = FAISS.from_texts(chunks, embeddings)

        responses = []
        issue = 0

        # Tried out different models

        # llm = HuggingFaceHub(repo_id="google/gemma-7b-it", model_kwargs={"temperature":0.1, "max_length":64})
        # llm = HuggingFaceHub(repo_id="google/gemma-2b-it", model_kwargs={"temperature":0.1, 'max_new_tokens':100})
        llm = HuggingFaceHub(repo_id="google/flan-t5-large", model_kwargs={"temperature": 5, "max_length": 64})  # best for now
        # llm = HuggingFaceHub(repo_id="google/flan-t5-large", model_kwargs={"temperature":6, "max_length":1024})
        # llm = HuggingFaceHub(repo_id="meta-llama/Llama-2-7b", model_kwargs={"temperature":0.1, "max_length":250})

        chain = load_qa_chain(llm, chain_type="stuff")

        # General
        # . Please give accurate and short answers.
        Questions = ['What is patient Name?', 'What is patient Age?', "What is patient's Date of Birth?", 'What is Gender of patient?', 'What is blood group of patient?', 'What is height of patient?',
                     'What is weight of patient?', 'What is Diagnosis of patient?', 'What is Follow-up Date?', 'What is Operation Date?', "What is patient's medical history?", 'What are additional comments?']

        # General
        Questions += ['What is hospital name?', 'What is Hospital address?',
                      'What is Doctor name?', 'What is Doctor Specialization?']

        for user_question in Questions:
            docs = knowledge_base.similarity_search(user_question)
            response = chain.run(input_documents=docs, question=user_question)
            responses.append(response)
            st.write(user_question, " : ", response)

        if ('Breast' or 'breast') in responses[7]:
            issue = 1
            st.write("The Patient has Breast Cancer.")

            # Breast Cancer
            Questions += ['Is there any family history of breast cancer?',
                          'Is there any history of cancer?']

            # Breast Cancer
            Questions += ["What is first Medication name, dosage and frequency?",
                          'What is second Medication name, dosage and frequency?', 'What is third Medication name, dosage and frequency?']

            # Breast Cancer
            Questions += ['What is Blood Pressure?', 'What is Heart Rate?',
                          'What is Respiratory Rate?', 'What is temperature?', 'What is Oxygen Saturation?']
            Questions += ['What is Tumor Marker (CA 15-3)?', 'What is Estrogen Receptor (ER) Status?',
                          'What is Progesterone Receptor (PR) Status?', 'What is HER2 Status?']

        elif ('Lung' or 'lung') in responses[7]:
            issue = 2
            st.write("The Patient has Lung Cancer.")

            # Lung Cancer
            Questions += ['Is there any family history of lung cancer?',
                          'Is there any previous history of cancer?']

            # Lung Cancer
            Questions += ["What is first Medication name, dosage and frequency?",
                          'What is second Medication name, dosage and frequency?']

            # Lung Cancer
            Questions += ['What is Blood Pressure?', 'What is Heart Rate?',
                          'What is Respiratory Rate?', 'What is temperature?', 'What is Oxygen Saturation?']
            Questions += ['What is Tumor Marker (CEA)?', 'What is EGFR Mutation?',
                          'What is ALK Reaarangement?', 'What is KRAS Mutation?']

        else:
            issue = 3
            st.write("The Patient has Orthopedic issue.")

            # Orthopedic / Fracture
            Questions += ['Is there any history of fractures?',
                          'Is there any family history?']
            
            Questions += ["What is Medication plan's name of first medicine, give only name.",
                          "What is Medication plan's dosage of first medicine, give only dosage.",
                          "What is Medication plan's frequency of first medicine, give only frequency."]
            Questions += ["What is first medicine start date?(mention NA if not there in document)",
                          "What is first medicine end date?(mention NA if not there in document)"]
            Questions += ["What is Medication plan's name of second medicine, give only name?",
                          "What is Medication plan's dosage of second medicine, give only dosage.",
                          "What is Medication plan's frequency of second medicine, give only frequency."]
            Questions += ["What is second medicine start date?(mention NA if not there in document)",
                          "What is second medicine end date?(mention NA if not there in document)"]

            # Orthopedic / Fracture
            Questions += ["What is Session Frequency of Physical Therapy Plan?",
                          "What is Start date of Physical Therapy Plan?", "What is end date of Physical Therapy Plan?"]
            Questions += ["What is Imaging technique?"]
            Questions += ["What is Fracture Type?",
                          "What is Fracture Location?", "What is Stability?"]
            Questions += ["What is Mobility?",
                          "What is Activities of Daily Living (ADLs)?"]

        for user_question in Questions[16:]:
            docs = knowledge_base.similarity_search(user_question)
            response = chain.run(input_documents=docs, question=user_question)

            st.write(user_question, " : ", response)
            responses.append(response)

        print(Questions)
        print(responses)

        # Creating dictionary to store key-value / question-answer pairs.

        response_dict = {}
        keys = []
        i = 0

        # Defining keys for every question

        if(issue==1):
            keys = ['name', 'age', 'dob', 'gender', 'BG', 'height', 'weight', 'issue', 'fp_date', 'op_date', 'medical_history', 'comments', 'hosp_name', 'hosp_add', 'doct_name', 'doct_sp', 'family_history', 'cancer_history', 'med1', 'med2', 'med3', 'BP', 'HR', 'RR', 'temp', 'OS', 'TM', 'ER', 'PR', 'HER2']
        elif(issue==2):
            keys = ['name', 'age', 'dob', 'gender', 'BG', 'height', 'weight', 'issue', 'fp_date', 'op_date', 'medical_history', 'comments', 'hosp_name', 'hosp_add', 'doct_name', 'doct_sp', 'family_history', 'cancer_history', 'med1', 'med2', 'BP', 'HR', 'RR', 'temp', 'OS', 'TM', 'EGFR', 'ALK', 'KRAS']
        else:
            keys = ['name', 'age', 'dob', 'gender', 'BG', 'height', 'weight', 'issue', 'fp_date', 'op_date', 'medical_history', 'comments', 'hosp_name', 'hosp_add', 'doct_name', 'doct_sp', 'fracture_history', 'family_history', 'med1','med1_dosage','med1_freq', 'med1_sd', 'med1_ed', 'med2','med2_dosage','med2_freq', 'med2_sd', 'med2_ed', 'pt_freq', 'pt_sd', 'pt_ed', 'IT', 'FT', 'FL', 'stability', 'mobility', 'ADLs']

        for response in responses:
            response_dict[keys[i]] = response
            i = i+1
            
        # Here we only have implemented for orthopedic patient form
        if(issue==3):
            st.write("The form for orthopedic patient is........................")
            # Display the patient's name
            patient_name = response_dict.get('name', 'Unknown')
            st.text_input("Patient Name", patient_name)
                
            # Age
            age = response_dict.get('age', 'Unknown')
            st.text_input("Age", age)

            # Date of Birth
            dob = response_dict.get('dob', 'Unknown')
            st.text_input("Date of Birth", dob)
                
            # Gender
            gender = response_dict.get('gender', 'Unknown')
            gender_options = ["Male", "Female", "Other"]
            selected_gender = gender if gender in gender_options else "Other"
            st.radio("Gender", gender_options, index=gender_options.index(selected_gender))
            
            # Blood Group
            blood_group = response_dict.get('BG', 'Unknown')
            st.text_input("Blood Group", blood_group)

            # Height
            height = response_dict.get('height', 'Unknown')
            st.text_input("Height (cm)", height)

            # Weight
            weight = response_dict.get('weight', 'Unknown')
            st.text_input("Weight (kg)", weight)

            # Diagnosis
            diagnosis = response_dict.get('issue', 'Unknown')
            st.text_input("Diagnosis", diagnosis)

            # Follow-up Date
            follow_up_date = response_dict.get('fp_date', 'Unknown')
            st.text_input("Follow-up Date", follow_up_date)

            # Operation Date
            operation_date = response_dict.get('op_date', 'Unknown')
            st.text_input("Operation Date", operation_date)

            # Your list of medical history options
            medical_history_options = [
                'Previous history of fractures',
                'Family history of bone disorders',
                'Other'
            ]

            # Get the selected medical history from response_dict, default to 'Other' if not found
            selected_history = response_dict.get('medical_history', 'Other')

            # Set the index of the default option
            default_index = 2  # 'Other' is at index 2

            # If selected_history is found in medical_history_options, update default_index
            if selected_history in medical_history_options:
                default_index = medical_history_options.index(selected_history)

            # Display the selected medical history with radio button
            st.write("Medical History")
            selected_history = st.radio("Medical History", medical_history_options, index=default_index)

            st.text_input("Other")
            
            
            # Get additional comments from response_dict, default to an empty string if not found
            comments = response_dict.get('comments', 'Unknown')
            st.text_input("Additional Comments", comments)
            
            st.write("...........................Hospital Information............................")
            
            # Get the hospital name from response_dict, default to 'Unknown' if not found
            hospital_name = response_dict.get('hosp_name', 'Unknown')
            st.text_input("Hospital Name", hospital_name)

            # Get the hospital address from response_dict, default to 'Unknown' if not found
            hospital_address = response_dict.get('hosp_add', 'Unknown')
            st.text_input("Hospital Address", hospital_address)

            # Get the doctor's name from response_dict, default to 'Unknown' if not found
            doctor_name = response_dict.get('doct_name', 'Unknown')
            st.text_input("Doctor Name", doctor_name)

            # Get the doctor's specialization from response_dict, default to 'Unknown' if not found
            doctor_specialization = response_dict.get('doct_sp', 'Unknown')
            st.text_input("Doctor Specialization", doctor_specialization)
            
            st.write("...........................Medication Plan............................")
            
            # Get Medicine 1 Name, Start Date, and End Date from response_dict, default to 'Unknown' if not found
            medicine1_name = response_dict.get('med1', 'Unknown')
            st.text_input("Medicine 1 Name", medicine1_name)
            
            medicine1_dosage = response_dict.get('med1_dosage', 'Unknown')
            st.text_input("Medicine 1 Dosage", medicine1_dosage)
            
            medicine1_freq = response_dict.get('med1_freq', 'Unknown')
            st.text_input("Medicine 1 Frequency", medicine1_freq)
            

            medicine1_start_date = response_dict.get('med1_sd', 'Unknown')
            st.text_input("Medicine 1 Start Date", medicine1_start_date)

            medicine1_end_date = response_dict.get('med1_ed', 'Unknown')
            st.text_input("Medicine 1 End Date", medicine1_end_date)

            # Get Medicine 2 Name, Start Date, and End Date from response_dict, default to 'Unknown' if not found
            medicine2_name = response_dict.get('med2', 'Unknown')
            st.text_input("Medicine 2 Name", medicine2_name)
            
            medicine2_dosage = response_dict.get('med2_dosage', 'Unknown')
            st.text_input("Medicine 2 Dosage", medicine2_dosage)
            
            medicine2_freq = response_dict.get('med2_freq', 'Unknown')
            st.text_input("Medicine 2 Frequency", medicine2_freq)
            

            medicine2_start_date = response_dict.get('med2_sd', 'Unknown')
            st.text_input("Medicine 2 Start Date", medicine2_start_date)

            medicine2_end_date = response_dict.get('med2_ed', 'Unknown')
            st.text_input("Medicine 2 End Date", medicine2_end_date)
            
            st.write("...........................Physical Therapy Plan............................")
            
                        # Get Session Frequency, Start Date, and End Date from response_dict, default to 'Unknown' if not found
            session_frequency = response_dict.get('pt_freq', 'Unknown')
            st.text_input("Session Frequency", session_frequency)

            start_date = response_dict.get('pt_sd', 'Unknown')
            st.text_input("Start Date", start_date)

            end_date = response_dict.get('pt_ed', 'Unknown')
            st.text_input("End Date", end_date)
                        
            st.write("...........................Imaging............................")
            x_ray1 = response_dict.get('IT', 'Unknown')
            st.text_input("X-ray:", x_ray1)
            
            st.write("...........................Orthopeadic Assessment............................")
                        # Get Fracture Type, Fracture Location, and Stability from response_dict, default to 'Unknown' if not found
            fracture_type = response_dict.get('FT', 'Unknown')
            st.text_input("Fracture Type", fracture_type)

            fracture_location = response_dict.get('FL', 'Unknown')
            st.text_input("Fracture Location", fracture_location)

            stability = response_dict.get('stability', 'Unknown')
            st.text_input("Stability", stability)
            
            st.write("...........................Functional Assessment............................")
                        # Get Mobility and Activities of Daily Living (ADLs) from response_dict, default to 'Unknown' if not found
            mobility = response_dict.get('mobility', 'Unknown')
            st.text_input("Mobility", mobility)

            ADLs = response_dict.get('ADLs', 'Unknown')
            st.text_input("Activities of Daily Living (ADLs)", ADLs)
            
            st.balloons()


if __name__ == '__main__':
    main()