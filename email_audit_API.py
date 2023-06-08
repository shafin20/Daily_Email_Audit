from flask import Flask, jsonify, request
from daily_email_audit.py import * 

app = Flask(__name__)

@app.route('/call_functions', methods=['POST'])
def call_functions():
    # Call all imported functions at once
    function_results = {}
    for function_name in dir():
        if callable(eval(function_name)) and function_name != 'call_functions':
            function_results[function_name] = eval(function_name)()

    # Prepare the response
    response = jsonify(function_results)

    return response

if __name__ == '__main__':
    app.run()
