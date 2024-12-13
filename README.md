# DocuFly

DocuFly is a Python application that automates the process of generating personalized documents from a Word template, converting them to PDFs, and sending them via email. This tool is especially useful for handling large volumes of customized documents and emails efficiently.

## Requirements

To run DocuFly, you need the following:

- **Python 3.6+**
- Required Python libraries:
  - `pandas`
  - `python-docx`
  - `openpyxl`
  - `docx2pdf`
  - `smtplib`
- A Gmail account for sending emails

## Installation

1. **Clone the repository** (or download the project files):

    ```sh
    git clone https://github.com/yourusername/docufly.git
    cd docufly
    ```

2. **Create and activate a virtual environment**:

    ```sh
    python3 -m venv myenv
    source myenv/bin/activate
    ```

3. **Install the required libraries**:

    ```sh
    pip install pandas==1.1.5 python-docx openpyxl docx2pdf
    ```

## Configuration

Create a `config.json` file in the root directory with the following structure:

```json
{
  "excel_path": "path/to/your/excel/file.xlsx",
  "word_template_path": "path/to/your/word/template.docx",
  "output_dir": "path/to/save/your/documents",
  "email_credentials": {
    "username": "your_gmail_address@gmail.com",
    "password": "your_app_specific_password"
  },
  "email_subject": "Your Document: ${name}",
  "file_name_pattern": "${name}.pdf",
  "email_body": "Dear ${name},\n\nWe are pleased to inform you that your application has been processed successfully.\nPlease find attached the details regarding your submission.\n\nBest regards,\nYour Company"
}
```

## Usage

1. **Run the application** by specifying the path to your `config.json` file:

    ```sh
    python script.py config.json
    ```

2. **Verify the generated PDFs and sent emails**.

## Example

### Excel File Structure

Ensure your Excel file has the following columns:

| Name  | Parameter1 | Parameter2 | Email                |
|-------|------------|------------|----------------------|
| John  | Value1     | Value2     | john@example.com     |
| Jane  | Value1     | Value2     | jane@example.com     |

### Word Template Placeholders

In your Word template, use placeholders in the format `${placeholder}`. For example:

```
Dear ${name},

We are pleased to inform you that your application has been ${Parameter1}. Please find attached the details regarding ${Parameter2}.

Best regards,
Your Company
```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or feature requests.

