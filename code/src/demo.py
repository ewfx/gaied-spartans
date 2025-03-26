from llama_cpp import Llama
from gpt4all import GPT4All 
from docx import Document
from collections import Counter
import os
import json


def read_docx_no_paragraphs(file_path):
    doc=Document(file_path)
    return " ".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])

def split_text(text,max_length=400):
    return [text[i:i+max_length]for i in range(0,len(text),max_length)]


model_file="Meta-Llama-3-8B-Instruct-Q4_0.gguf"
model_folder = "PropFile"
model_path = os.path.join(model_folder,model_file)
file_name="Invoice.docx"
file_path = os.path.join(model_folder,file_name)
llm=Llama(model_path)
document_text = read_docx_no_paragraphs(file_path)
chunks = split_text(document_text,max_length=400)

few_shot_prompt = (
    "You are an expert at classifying documents by their request type."
    "Based on the content, assign one of the following labels:Finacial, Techincal, or Legal.\n\n"
    "Return only Answer with no explanation.\n\n"
    "Example 1:\n"
    "Document:Invoice for payment due next month.\n"
    "Request Type:Financial\n\n"
    "Example 2:\n"
    "Document:API integration guide explaining authentication procedures.\n"
    "Request Type:Technical\n\n"
    "Now classify the following document.\n"
    "Document:{document_chunk}\n"
    "Request Type:"
)
results = []
for chunk in chunks:
    prompt = few_shot_prompt.format(document_chunk=chunk)
    response=llm(prompt)
    answer = response["choices"][0]["text"].strip()
    results.append(answer)

final_classification = Counter(results).most_common(1)[0][0]
output = {"request_type":final_classification}
print(json.dumps(output))