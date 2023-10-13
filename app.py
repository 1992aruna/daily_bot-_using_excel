from flask import Flask, request, jsonify
from flask_pymongo import PyMongo
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import os
import requests
import pandas as pd
from pymongo import MongoClient
from dotenv import load_dotenv
from messages import *
from utils import retrieve_user_answers, send_excel_file
from google.oauth2 import service_account
import gspread
import logging
import schedule
import datetime
import time


# Specify the path to your service account JSON key file
keyfile_path = 'google_cloud.json'

# Authenticate using the service account JSON key file
credentials = service_account.Credentials.from_service_account_file(
    keyfile_path, 
    # scopes=['https://www.googleapis.com/auth/spreadsheets']
    scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
)

# Authorize with gspread using the credentials
client = gspread.authorize(credentials)

# Open the Google Sheets document by its title
# spreadsheet = client.open('Daily_Questions')

# # Select the worksheet where you want to export data (if it exists)
# question_worksheet = spreadsheet.worksheet('Sheet1')  # Replace 'Sheet1' with your sheet name


# Load environment variables
load_dotenv()

# Replace with your MongoDB URI
# MONGO_DB = 'sbi'  # Replace with your database name
# STAFF_COLLECTION = 'staff'
# ANSWERS_COLLECTION = 'answers'

MONGO_URI = os.getenv("MONGO_URI")
API_URL = os.getenv("API_URL")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")

app = Flask(__name__)

app.config["MONGO_URI"] = MONGO_URI
mongo = PyMongo(app)
db = mongo.db.staff
answers_db = mongo.db.answers
# fs = gridfs.GridFS(mongo.db, collection="files")

# Initialize Wati API endpoint
WATI_API_ENDPOINT = f"{API_URL}/api/v1/sendSessionMessage"

scheduler = BackgroundScheduler()



# Function to send image message
def send_image_message(phone_number,image, caption):
    url = f"{API_URL}/api/v1/sendSessionFile/{phone_number}?caption={caption}"

    payload = {}
    files=[
    ('file',('file',open(image,'rb'),'image/jpeg'))
    ]
    headers = {
    'Authorization': ACCESS_TOKEN
    }

    response = requests.post(url, headers=headers, json=payload, files=files)
    print(response)
    print(response.json())

def get_questions_from_spreadsheet(worksheet):
    try:
        questions = worksheet.col_values(1)
        return questions
        # return questions[:10]
    except Exception as e:
        print(f"Error fetching questions from the Google Spreadsheet: {str(e)}")
        return []

def send_questions_to_contact(contact_number, questions):
    for question in questions:
        send_message(contact_number, question)  # Implement the send_message function
  
    
def send_branch_images():
    try:
        spreadsheet = client.open('Daily_Questions')
        worksheet = spreadsheet.worksheet('Sheet1')

        questions = get_questions_from_spreadsheet(worksheet)
        
        for staff in db.find({"status": ""}):
            # Check if 'branch' and 'phone_number' fields exist in the document
            if 'branch' in staff and 'phone_number' in staff:
                branch = staff['branch']
                phone_number = staff['phone_number']

                # Check if an image exists for this branch with either .png or .jpg extension
                image_extensions = ['.png', '.jpg']
                image_found = False

                for ext in image_extensions:
                    image_path = f'D:\\New Project\\Python\\New_Bot\\Bot\\daily_bot _using_excel\\branch_images\\{branch}{ext}'

                    if os.path.isfile(image_path):
                        image_found = True
                        print("Image exists. Sending to", phone_number)
                        # Provide a caption for the image message
                        caption = f'Here is your image for branch {branch}'
                        send_image_message(phone_number, image_path, caption)
                        print(f"Image sent for branch {branch} with extension {ext}")
                        send_questions_to_contact(phone_number, questions)
                        print(f"Questions sent for branch {branch} phone number {phone_number}")
                        db.update_one({"_id": staff["_id"]}, {"$set": {"status": "sent"}})


                if not image_found:
                    print(f"No image found for branch {branch}")
            else:
                print("Missing 'branch' or 'phone_number' field in the document.")

        # Close the MongoDB connection
        # client.close()
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def process_message(phone_number, message):
    spreadsheet = client.open('Daily_Questions')
    worksheet = spreadsheet.worksheet('Sheet1')

    questions = get_questions_from_spreadsheet(worksheet)

    print(f"Received message: {message} from phone_number: {phone_number}")
    
    question_number = extract_question_number(message)
    print(f"Extracted question number: {question_number}")
    
    # Get the corresponding question
    question = questions[question_number - 1]  # Subtract 1 because list indices start at 0

    # Extract only the response text from the message
    response_text = message.split('.', 1)[1].strip() if '.' in message else message

    # Check if a document for this phone number already exists
    answers_received = mongo.db.answers_received.find_one({'phone_number': phone_number})

    if answers_received:
        # If a document exists, update it with the new response
        mongo.db.answers_received.update_one(
            {'phone_number': phone_number},
            {'$set': {f'question_{question_number}': question, f'answer_{question_number}': response_text}}
        )
        print(f"Updated responses in database for phone number: {phone_number}")
    else:
        # If no document exists, create a new one
        answers_received = {
            'phone_number': phone_number,
            f'question_{question_number}': question,
            f'answer_{question_number}': response_text,
        }
        result = mongo.db.answers_received.insert_one(answers_received)
        print(f"Inserted responses into database, received ID: {result.inserted_id}")

def extract_question_number(message):
    # Split the message into words
    words = message.split()
    
    for word in words:
        # Remove any trailing period
        if word.endswith('.'):
            word = word[:-1]
        
        # Check if the word is a digit
        if word.isdigit():
            return int(word)
    
    return None  # Return None if no question number was found

def create_excel_report(user_answers):
    # Create a DataFrame from the user answers
    df = pd.DataFrame(user_answers)

    # Get the current date to create a filename with just the date
    current_date = datetime.date.today()
    formatted_date = current_date.strftime("%Y-%m-%d")
    output_folder_path = "D:/New Project/Python/New_Bot/Bot/daily_bot _using_excel/Output"
    file_name = f"answer_{formatted_date}.xlsx"
    file_path = os.path.join(output_folder_path, file_name)

    # Save the Excel file using the full file path
    df.to_excel(file_path, index=False)



def generate_report():
    print("Report generation task executed at", time.ctime())
    print("generate_report function started")  # Print when the function starts
    user_answers = retrieve_user_answers()
    print(f"user_answers: {user_answers}")  # Print the user answers
    create_excel_report(user_answers)
    print("Excel report created")  # Print after the report is created

    phone_number = "917892409211"
    print(f"Sending file to: {phone_number}")
    # mongo.db.answers_received.delete_many({})
    send_file(phone_number)
    

# schedule.every().day.at("17:51").do(generate_report)
# Schedule the generate_report function to run daily at 17:51
scheduler.add_job(generate_report, trigger=CronTrigger(hour=19, minute=20))

# Start the scheduler
scheduler.start()

def send_file(phone_number):
    dir = 'D:/New Project/Python/New_Bot/Bot/daily_bot _using_excel/Output'
    # phone_number = "917892409211"
    # Get the current date to create the file name
    current_date = datetime.date.today()
    formatted_date = current_date.strftime("%Y-%m-%d")
    
    file_name = f'answer_{formatted_date}.xlsx'
    file_path = f'{dir}/{file_name}'
    caption = 'Your daily report'
    send_excel_file(phone_number, file_path, caption)

allowed_extensions=["png", "jpg", "jpeg"]

@app.route('/')
def home():
  return "Ink Pen Bot Live 1.0"

@app.route("/webhook", methods=['GET'])
def connetwebhook():
    return "running whatsapp webhook"


@app.route('/webhook', methods=['POST'])
def webhook():
    
    try:
        # Extract message details from request
        # generate_report()
        print(f"Received POST request with JSON: {request.json}")
        message = request.json.get('text')
        phone_number = request.json.get('waId')

        print(f"Received POST request with message: {message} and phone number: {phone_number}")

        # Process the message and save the response
        process_message(phone_number, message)

        # generate_report()
        
        return jsonify({'message': 'Webhook executed successfully'}), 200
        
    except Exception as e:
        logging.exception("An error occurred: %s", e)
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    send_branch_images()
    app.run(debug=True)

