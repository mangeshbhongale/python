from flask import Flask, request, jsonify
from datetime import datetime
import os
from aqi_downloader import AQIBulletinDownloader
from extract_aqi import AQIDataExtractor

app = Flask(__name__)

@app.route('/api/aqi', methods=['POST'])
def get_aqi_data():
    """
    API to fetch AQI data for a city within a date range.
    Input:
        - start_date: Start date in YYYY-MM-DD format
        - end_date: End date in YYYY-MM-DD format
        - city: Name of the city
    Output:
        - JSON response with dates and AQI data
    """
    try:
        # Parse request JSON
        request_data = request.get_json()
        start_date = request_data.get('start_date')
        end_date = request_data.get('end_date')
        city = request_data.get('city')

        # Validate input
        if not start_date or not end_date or not city:
            return jsonify({'error': 'start_date, end_date, and city are required'}), 400
        
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
            datetime.strptime(end_date, "%Y-%m-%d")
        except ValueError:
            return jsonify({'error': 'Invalid date format. Use YYYY-MM-DD'}), 400

        # Directories
        pdf_dir = "aqi_bulletins"
        
        # Step 1: Download AQI PDFs
        downloader = AQIBulletinDownloader(output_dir=pdf_dir)
        downloader.download_range(start_date, end_date)

        # Step 2: Extract AQI data
        extractor = AQIDataExtractor(pdf_dir)
        data = extractor.process_date_range(city, start_date, end_date)

        # Check if data is found
        if not data:
            return jsonify({'error': f'No data found for {city} in the specified date range'}), 404
        
        # Return extracted data as JSON
        return jsonify({'city': city, 'start_date': start_date, 'end_date': end_date, 'aqi_data': data}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
