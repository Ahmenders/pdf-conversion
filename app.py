from flask import Flask, request
from utils import *
import logging

app = Flask(__name__)

# Route to get all data
@app.route('/convert', methods=['POST'])
def get_data():
    requestData = request.get_json()
    
    try:
        path = requestData["fileURL"]
        outputPath = requestData["outputPath"]
        current = requestData["from"]
        to = requestData["to"]
        logging.info(requestData)
    except Exception as exp:
        logging.exception("#Error in fetching the parameters")
        return {"status": "failed", "message": "Error in fetching the parameters", "data": None}
    
    response = None
    if current == "pdf" and to == "docx":
        response = pdf2doc(path, outputPath)
    elif current == "pdf" and to =="ppt":
        response = pdf2ppt(path, outputPath)
    elif current == "pdf" and to == "csv":
        response = pdf2csv(path, outputPath)
    else:
        logging.exception("#Error No support for the given conversion")
        return {"status": "failed", "message": "No support for the given conversion", "data": None}
    
    if response:
        return {"status": "success", "message": "File has been converted successfully", "data": response}
    else:
        return {"status": "failed", "message": "Error in converting the file", "data": None}

if __name__ == '__main__':
    app.run(debug=True)
