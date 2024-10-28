from flask import Flask, request, render_template, redirect
import openpyxl
from openpyxl import Workbook
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
EXCEL_FILE = 'seed_data.xlsx'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

if not os.path.exists(EXCEL_FILE):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Seed Data"
    sheet.append(["Name", "Email", "Seed Type", "Notes", "Photo Path"])  # Updated to "Notes"
    workbook.save(EXCEL_FILE)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        name = request.form['name']
        email = request.form['email']
        seed_type = request.form['seed_type']
        notes = request.form['notes']  # Updated to use 'notes'
        photo = request.files['seed_photo']
        
        if photo:
            photo_path = os.path.join(UPLOAD_FOLDER, photo.filename)
            photo.save(photo_path)
        else:
            photo_path = "No photo provided"

        workbook = openpyxl.load_workbook(EXCEL_FILE)
        sheet = workbook.active
        sheet.append([name, email, seed_type, notes, photo_path])  # Updated to save 'notes'
        workbook.save(EXCEL_FILE)

        return redirect('/')
    except Exception as e:
        print("An error occurred:", e)
        return "There was an error processing your request."

if __name__ == '__main__':
    app.run(debug=True)
