# Employee Reminder Email Script

This Python script reads employee data from an Excel file and sends reminder emails to employees using Gmail's SMTP server.

## Prerequisites

- Python 3.x
- `openpyxl` library
- `python-dotenv` library
- Gmail account

## Installation

1. **Clone the repository or download the script.**
    ```bash
    git clone https://github.com/your_username/your_repository.git
    ```
2. **Navigate to the project directory:**
    ```bash
    cd your_repository
    ```
3. **Install required Python libraries:**
    ```bash
    pip install openpyxl python-dotenv
    ```

4. **Create a `.env` file in the same directory as the script with your Gmail credentials:**
    ```env
    SENDER_EMAIL=your_email@gmail.com
    SENDER_PASSWORD=your_password
    ```

## Usage

1. **Run the script:**
    ```bash
    python script_name.py
    ```

2. **The script will read employee data from `Employee.xlsx` (make sure to replace this with your Excel file's name and path).**
3. **It will send reminder emails to employees about upcoming events based on the data in the Excel file.**

## Excel File Format

The script expects the Excel file to have the following columns:

1. `Name`: Employee's name
2. `Email`: Employee's email address
3. `Reminder Date`: Date of the upcoming event

**Make sure to adjust the `read_employee_data()` function if your Excel file has different column names or structure.**

## Notes

- **Make sure to enable "Less secure app access" or generate an App Password for your Gmail account if you have two-factor authentication enabled.**
- **Replace `script_name.py` with the actual name of your script.**
- **Replace `Employee.xlsx` with the actual name and path of your Excel file.**
