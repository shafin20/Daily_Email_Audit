from flask import app, jsonify
import pandas as pd

from api import get_database_connection


@app.route('/api', methods=['GET'])
def get_data():
    conn = get_database_connection()
    query = 'SELECT * FROM your_table'
    data = pd.read_sql_query(query, conn)
    conn.close()

    return jsonify(data.to_dict(orient='records'))
    
if __name__ == '__main__':
    app.run()
