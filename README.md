<<<<<<< HEAD

# PDF to Excel Automation

This script automates the process of downloading PDF attachments from Gmail, extracting tables, and saving them to an Excel file.

## Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/mrmlb94/pdf-to-excel-automation.git
   cd pdf-to-excel-automation
   ```

2. **Create a virtual environment and activate it**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. **Install the required packages**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Create a `.env` file in the project directory and add your environment variables**:
   ```dotenv
   EMAIL=your_email@gmail.com
   PASSWORD=your_app_password
   IMAP_SERVER=imap.gmail.com
   IMAP_PORT=993
   DOWNLOAD_FOLDER=/path/to/download/folder
   ```

5. **Download your `credentials.json` file from Google Cloud Console and place it in the project directory**.

6. **Run the script**:
   ```bash
   python main.py
   ```

## Usage

The script will:
1. Authenticate with Gmail.
2. Download the latest PDF attachment from a specified sender.
3. Extract tables from the PDF.
4. Save the tables to an Excel file with the same name as the PDF.

## Requirements

Create a `requirements.txt` file with the following content:
```
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
pdfplumber
pandas
python-dotenv
openpyxl
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
=======
# pdf-to-excel-automation
This script automates the process of downloading PDF attachments from Gmail, extracting tables, and saving them to an Excel file.
>>>>>>> 2d2311357a7ac0246bb44ea36a6c09c584b6330a
