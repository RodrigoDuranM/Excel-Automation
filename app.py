import os
import re
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from actions import *  # File containing all the actions
from custom_parserAI import parse_command
from time import sleep
app = Flask(__name__)

# Directory to store uploaded and processed files
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Provides a route (link) to access the html
@app.route('/')
def index():
    return render_template('index.html')

# This is for the chat site (handles AJAX request)
@app.route('/chat', methods=['GET', 'POST'])
def chat():
    if request.method == 'POST':
        # Get the user message and the uploaded file
        user_message = request.form.get('message')
        uploaded_file = request.files.get('file')
        
        # Use the custom parser to extract the function name and its parameters
        command = parse_command(user_message)
        if not command:
            return "Command not recognized", 400
        
        func_name = command['function']  # e.g., "number_sorting"
        parameters = command['parameters']  # e.g., {"header": "Score", "order": "asc"}
        
        # Ensure a file is provided (if required by the command)
        if not uploaded_file:
            return "No file provided!", 400
        
        # Map function names to their corresponding function objects 
        FUNCTION_MAP = {
            'number_sorting': number_sorting,
            'word_sorting': word_sorting,
            'file_filter': file_filter,
            'dup_remover': dup_remover,
            'num_highlight': num_highlight,
            'word_highlight': word_highlight,
            'styling': styling
        }
        sleep(1)
        # Get function by name and call it dynamically
        if func_name in FUNCTION_MAP:
            return FUNCTION_MAP[func_name](uploaded_file, **parameters, app=app)
        else:
            return jsonify(response="Unknown function requested."), 400
    
    # For GET requests, simply render the chat interface.
    return render_template('chat.html')


# This is for the sorting site
@app.route('/sorting')
def sorting():
    return render_template('sorting.html')

@app.route('/sort_numbers', methods=['GET', 'POST'])
def sort_numbers():
    if request.method == 'POST':
        file = request.files['fileInput']
        header = request.form['headerInput']
        order = request.form['sortOrder']

        return number_sorting(file, header, order, app)

    return render_template('sort_numbers.html')

@app.route('/sort_words', methods=['GET', 'POST'])
def sort_words():
    if request.method == 'POST':
        file = request.files['fileInput']
        header = request.form['headerInput']
        order = request.form['sortOrder']

        return word_sorting(file, header, order, app)

    return render_template('sort_words.html')

# This is for the filtering site
@app.route('/filtering')
def filtering():
    return render_template('filtering.html')

@app.route('/filter', methods=['GET', 'POST'])
def filter():
    if request.method == 'POST':
        file = request.files.get('fileInput')
        header = request.form.get('headerInput')
        value = request.form.get('valueInput')
        sort_order = request.form.get('sortOrder')  # 'desc' (higher) or 'asc' (lower)

        return file_filter(file, header, value, sort_order, app)

    return render_template('filter.html')

@app.route('/remove', methods=['GET', 'POST'])
def remove():
    if request.method == 'POST':
        file = request.files.get('fileInput')

        return dup_remover(file, app)

    return render_template('remove.html')

@app.route('/highlightNum', methods=['GET', 'POST'])
def highlightNum():
    if request.method == 'POST':
        file = request.files['fileInput']
        header = request.form['headerInput']
        condition = request.form['sortOrder']
        key_value = request.form['valueInput']

        return num_highlight(file, header, condition, key_value, app)

    return render_template('highlightNum.html')

@app.route('/highlightWord', methods=['GET', 'POST'])
def highlightWord():
    if request.method == 'POST':
        file = request.files['fileInput']
        header = request.form['headerInput']
        key_value = request.form['valueInput']

        return word_highlight(file, header, key_value, app)

    return render_template('highlightWord.html')

# This is for the styling site
@app.route('/styling', methods=['GET', 'POST'])
def styling():
    if request.method == 'POST':
        file = request.files['fileInput']
        font_color = request.form['fontColor']
        header_background = request.form['headerBackground']
        bold_headers = request.form['boldHeaders']
        border_thickness = request.form['borders']
        header_alignment = request.form['headerAlignment']
        data_alignment = request.form['dataAlignment']

        return styling(file, font_color, header_background, bold_headers, border_thickness, header_alignment, data_alignment, app)

    return render_template('styling.html')

# This is for the 'more services' site
@app.route('/more')
def more():
    return render_template('more.html')

if __name__ == '__main__':
    app.run(debug=True)  # To run locally on this computer
    #app.run(host='0.0.0.0', port=5000)  # To run on other devices in the same network
