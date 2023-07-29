from flask import Flask, render_template, request, jsonify, send_file, make_response
from docx import Document
import os


 

app = Flask(__name__)

 

 

# Making docs file

def generate_document(data):

    template_file_path = 'Document.docx'

    output_file_path = data["name"]

    print(output_file_path)

 

    variables = {

        "${EMPLOYEE_NAME}": data["name"],

        "${EMPLOYEE_DESIGNATION}": data['designation'],

        "${EMP_LOCATION}": data['location'],

        "${DOJ}": "17/05/2022",

        "${ANNUAL_BASIC_CTC}": data['basic_percentage'],

        "${MON_BASIC_CTC}": data['monthly_basic'],

        "${ANNUAL_HRA}": data['hra_percentage'],

        "${MONTHLY_HRA}": data['monthly_hra'],

        "${SPL_ALLOWANCE}": data['special_allowance_percentage'],

        "${MON_ALLOWANCE}": data['monthly_special_allowance'],

        "${A_CONVEYANCE}": data['conveyance_percentage'],

        "${MON_CONVEYANCE}": data['monthly_conveyance'],

        "${ANNUAL_TOTAL}": data['annual_ctc'],

        "${MONTHLY_TOTAL}": data['monthly_ctc'],

    }

 

    template_document = load_template(template_file_path)

 

    for placeholder, replacement_value in variables.items():

        replace_text_in_paragraphs(

            template_document.paragraphs, placeholder, replacement_value)

        replace_text_in_tables(template_document.tables,

                               placeholder, replacement_value)

 

    save_document(template_document, output_file_path)

    return output_file_path

 

 

def load_template(file_path):

    return Document(file_path)

 

 

def replace_text_in_paragraphs(paragraphs, placeholder, replacement_value):

    for paragraph in paragraphs:

        if placeholder in paragraph.text:

            paragraph.text = paragraph.text.replace(

                placeholder, replacement_value)

 

 

def replace_text_in_tables(tables, placeholder, replacement_value):

    for table in tables:

        for row in table.rows:

            for cell in row.cells:

                replace_text_in_paragraphs(

                    cell.paragraphs, placeholder, replacement_value)

 

 

def save_document(document, file_path):

    document.save(file_path)

 

 

@app.route('/')

def index():

    return render_template('form.html')

 

 

@app.route('/calculate', methods=['POST'])

def calculate():

    data = request.get_json()

    file_path = generate_document(data)

    print(data['name'])

 

    # Get the filename from the path

    filename = data["name"]

 

    # Create a response with the file

    response = make_response(send_file(file_path))

 

    # Set the 'Content-Disposition' header for direct download

    response.headers["Content-Disposition"] = f"attachment; filename={filename}"

    print(response)

    # Return the response

    return response

 

if __name__ == '__main__':

    app.run(debug=True)