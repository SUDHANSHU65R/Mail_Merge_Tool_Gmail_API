# Mail Merge Tool

## Description

Developed a mail merge tool using Google API, Google Sheets, and Python. The tool streamlines the mail merging process by automating the retrieval of template data, email addresses, HTML template creation, and email sending, enhancing efficiency and productivity.

## Key Features

- Utilized Google Sheets as a data source for mail merge operations.
- Extracted mail template data from Google Sheets and dynamically created HTML templates using Python.
- Integrated Gmail API to fetch email addresses from Google Sheets and send personalized emails based on the mail template.
- Implemented a logging mechanism to create a detailed report of sent emails, providing insights into the delivery status and ensuring accountability.

## Technologies Used

- Python
- Google API
- Google Sheets
- Gmail API
- GCP IAM Account for Clint Secret key.

## How It Works

1. The tool retrieves mail template data from Google Sheets.
2. Using Python, it dynamically creates HTML templates based on the retrieved data.
3. The Gmail API is utilized to fetch email addresses from Google Sheets and send personalized emails using the HTML templates.
4. A logging mechanism is implemented to create a detailed report of sent emails, providing insights into the delivery status and ensuring accountability.

## Benefits

- Automates the mail merging process, saving time and effort.
- Enhances efficiency and productivity by streamlining the workflow.
- Provides detailed insights into the delivery status of sent emails, ensuring accountability.

## Usage

To use the mail merge tool:
1. Clone the repository.
2. Install dependencies using `pip install -r requirements.txt`.
3. Configure Google API credentials and access to Google Sheets.
4. Run the tool and enjoy streamlined mail merging.

## Contributions

Contributions are welcome! Feel free to fork the repository, make improvements, and submit pull requests.

## License

This project is licensed under the [MIT License](LICENSE).

## Contact

For any inquiries or support, please contact [Sudhanshu Kumar](mailto:Sudhansu65r@gmail.com).
