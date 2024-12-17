# Excel Transcript Generator

## Overview
This Flask application allows users to convert Excel spreadsheets containing student grade data into individual PDF transcripts. The application supports:
- Uploading Excel files (.xlsx, .xls)
- Extracting student information and grades
- Generating individual PDF transcripts
- Creating a ZIP file with all generated transcripts

## Features
- Automatic grade calculation based on percentage
- Support for multiple courses and grade components
- Letter grade and grade point assignment
- SGPA (Semester Grade Point Average) calculation
- PDF transcript generation using HTML template
- ZIP file download of all generated transcripts

## Prerequisites
- Python 3.8+
- Flask
- openpyxl
- Jinja2
- WeasyPrint
- Werkzeug

## Installation
1. Clone the repository
2. Ensure you have the `requirements.txt` file in the project directory
3. Create a virtual environment (recommended):
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```
4. Install dependencies:
   ```bash
   pip3 install -r requirements.txt
   ```

## Project Structure
```
project_root/
│
├── app.py                # Main Flask application
├── requirements.txt      # Python dependencies
├── uploads/              # Directory for uploaded Excel files
├── templates/
│   └── transcript_template.html  # HTML template for transcripts
└── README.md             # This documentation file
```

## Usage
1. Run the application:
   ```bash
   python app.py
   ```
2. Navigate to `http://localhost:5000` in your web browser
3. Upload an Excel file with the following structure:
   - Column A: Serial Number
   - Column B: Student Name
   - Columns C onwards: Course details, components, and marks

## Excel File Requirements
- First row: Course names
- Second row: Component names
- Third row: Maximum marks for each component
- Subsequent rows: Student data

## Configuration
- Change `app.secret_key` for production
- Modify `UPLOAD_FOLDER` path if needed
- Adjust allowed file extensions in `ALLOWED_EXTENSIONS`

## Grading Scheme
The application uses a predefined grading scheme:
- 90% and above: A+ (4.00)
- 85-89%: A (3.67)
- 80-84%: B+ (3.33)
- 75-79%: B (3.00)
- 70-74%: C+ (2.67)
- 60-69%: C (2.33)
- Below 60%: F (0.00)

## Limitations
- Assumes 3 credits per course (can be modified)
- Simple SGPA calculation
- Requires a specific Excel file format

## Contributing
1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License
Specify your project's license here (e.g., MIT, Apache 2.0)

## Contact
Add your contact information or project maintainer details