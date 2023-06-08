from flask import Flask, jsonify, request
from daily_email_audit import *
import pandas as pd

app = Flask(__name__)

@app.route('/create_audit_id', methods=['POST'])
def create_audit_id():
    record = request.get_json()
    result = crud.create_audit_id(record['audit_id'])
    return jsonify(result)

@app.route('/get_inbox_ids', methods=['POST'])
def get_inbox_ids():
    record = request.get_json()
    result = crud.get_inbox_ids(record['mailbox'], record['ews_config'])
    return jsonify(result)

@app.route('/get_journaled_email', methods=['POST'])
def get_journaled_email():
    record = request.get_json()
    result = crud.get_journaled_email(record['mailbox'], record['ews_config'], record['item_id'])
    return jsonify(result)

@app.route('/calculate_file_sha256', methods=['POST'])
def calculate_file_sha256():
    record = request.get_json()
    result = crud.calculate(record['file_path'])
    return jsonify(result)

@app.route('/write_email_audit', methods=['POST'])
def write_email_audit():
    record = request.get_json()
    result = crud.write_email_audit(
        record['audit_id'],
        record['audit_datetime'],
        record['message_id'],
        record['message_from'],
        record['message_to'],
        record['message_cc'],
        record['message_bcc'],
        record['message_subject'],
        record['message_body'],
        record['message_attachments'],
        record['message_other'],
        record['message_audit_result'],
        record['mailbox']
    )
    return jsonify(result)

@app.route('/open_eml_as_attachment', methods=['POST'])
def open_eml_as_attachment():
    record = request.get_json()
    result = crud.open_eml_as_attachment(record['eml_file_path'])
    return jsonify(result)

@app.route('/compare_email_files', methods=['POST'])
def compare_email_files():
    record = request.get_json()
    result = crud.compare_email_files(record['file_path1'], record['file_path2'])
    return jsonify(result)

@app.route('/get_attachments', methods=['POST'])
def get_attachments():
    record = request.get_json()
    result = crud.get_attachments(record['eml'])
    return jsonify(result)

if __name__ == '__main__':
    app.run()
