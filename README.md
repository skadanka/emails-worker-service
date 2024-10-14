![מבנה מערכת](https://github.com/user-attachments/assets/807915ca-d0f4-4503-b895-5146c00dd2fd)![Worker drawio](https://github.com/user-attachments/assets/6c562743-9535-4511-bbe1-7f1f43b1deca)# HR Automation Service

## Overview

This project is an HR automation service designed to streamline the process of reading resumes from emails and creating Salesforce objects for candidates. The system integrates with Microsoft Outlook to retrieve emails, processes attached files, and submits the processed data to Salesforce. The project was developed iteratively, following client requirements, with a focus on modularity, maintainability, and performance.

## Features

- **Email Processing**: Connects to Microsoft Outlook, retrieves email messages, and extracts relevant data.
- **File Parsing**: Processes resumes in various formats, including `.pdf` and `.docx` using OCR where necessary.
- **Salesforce Integration**: Submits parsed candidate data to Salesforce via an HTTP API.
- **Error Handling & Logging**: Provides comprehensive error handling and logging for failed email processing or file parsing attempts.
- **Folder Management**: Automatically moves processed emails to appropriate folders (`Processed`, `Not Completed`).

## System Pipeline

The pipeline involves several components:

1. **Email Retrieval**: The system connects to the Outlook inbox, retrieves new emails, and filters them for relevant content (e.g., resumes or job applications).
2. **File Processing**:
    - Extracts text from attached files using the `DocumentReaderComponent`. This includes reading PDFs and DOCX files and applying OCR for image-based text extraction.
3. **Form Creation**: The `FormCreatorOutlook` processes the email and the extracted data, generating candidate profiles from structured content.
4. **Data Submission**: The parsed form models are submitted to Salesforce using the `FormSubmitSalesForce` service.
5. **Folder Organization**: Successfully processed emails are moved to the "Processed" folder, while emails that failed are moved to the "Not Completed" folder for further review.

## Components

### Outlook Service
- Connects to the Microsoft Outlook application to retrieve emails from specified folders.
- Marks emails as processed and moves them to the correct folders after successful processing.

### Document Reader Component
- A utility that handles reading resumes attached to emails.
- Supports PDF and DOCX file formats and uses OCR for image-based text extraction from PDFs.

### Form Creator
- Extracts relevant details from emails and resumes.
- Generates structured form models containing candidate details like name, email, phone, and resume content.

### Salesforce Integration
- Submits candidate form models to Salesforce using a secure API.
- Handles validation of required fields and logs errors in case of submission failures.

## System Design

_**(Placeholder for system design diagrams)**_

## Running the Windows Service

1. Publish the application as a self-contained executable:
   ```bash
   dotnet publish -c Release -r win10-x64
   ```

2. Install the service using Task Scheduler or register it as a Windows Service.
   _**(Instructions for Windows Service setup here)**_

## Usage

Once the service is running, it will:

1. Retrieve emails from the Outlook inbox.
2. Process any resumes attached to those emails.
3. Submit the processed data to Salesforce.
4. Log processing results and move emails to the appropriate folders.

## Error Handling

- Emails that fail to process correctly will be moved to the "Not Completed" folder, and an error log will be generated.
- The service logs both successful and failed form submissions.

## Contribution

1. Fork the repository.
2. Create a feature branch.
3. Submit a pull request.

## License

_**(Specify License if applicable)**_# HR Automation Service

## Overview

This project is an HR automation service designed to streamline the process of reading resumes from emails and creating Salesforce objects for candidates. The system integrates with Microsoft Outlook to retrieve emails, processes attached files, and submits the processed data to Salesforce. The project was developed iteratively, following client requirements, with a focus on modularity, maintainability, and performance.

## Features

- **Email Processing**: Connects to Microsoft Outlook, retrieves email messages, and extracts relevant data.
- **File Parsing**: Processes resumes in various formats, including `.pdf` and `.docx` using OCR where necessary.
- **Salesforce Integration**: Submits parsed candidate data to Salesforce via an HTTP API.
- **Error Handling & Logging**: Provides comprehensive error handling and logging for failed email processing or file parsing attempts.
- **Folder Management**: Automatically moves processed emails to appropriate folders (`Processed`, `Not Completed`).

## System Pipeline

The pipeline involves several components:

1. **Email Retrieval**: The system connects to the Outlook inbox, retrieves new emails, and filters them for relevant content (e.g., resumes or job applications).
2. **File Processing**:
    - Extracts text from attached files using the `DocumentReaderComponent`. This includes reading PDFs and DOCX files and applying OCR for image-based text extraction.
3. **Form Creation**: The `FormCreatorOutlook` processes the email and the extracted data, generating candidate profiles from structured content.
4. **Data Submission**: The parsed form models are submitted to Salesforce using the `FormSubmitSalesForce` service.
5. **Folder Organization**: Successfully processed emails are moved to the "Processed" folder, while emails that failed are moved to the "Not Completed" folder for further review.

## Components

### Outlook Service
- Connects to the Microsoft Outlook application to retrieve emails from specified folders.
- Marks emails as processed and moves them to the correct folders after successful processing.

### Document Reader Component
- A utility that handles reading resumes attached to emails.
- Supports PDF and DOCX file formats and uses OCR for image-based text extraction from PDFs.

### Form Creator
- Extracts relevant details from emails and resumes.
- Generates structured form models containing candidate details like name, email, phone, and resume content.

### Salesforce Integration
- Submits candidate form models to Salesforce using a secure API.
- Handles validation of required fields and logs errors in case of submission failures.

## System Design
![מבנה מערכת drawio](https://github.com/user-attachments/assets/f1798662-7098-4f74-ac8b-42c9e3a2f2a1)

![Worker](https://github.com/user-attachments/assets/b68c3d09-5697-4d49-8ef1-8484ab12601c)

## Running the Windows Service

1. Publish the application as a self-contained executable:
   ```bash
   dotnet publish -c Release -r win10-x64
   ```

2. Install the service using Task Scheduler or register it as a Windows Service.
   _**(Instructions for Windows Service setup here)**_

## Usage

Once the service is running, it will:

1. Retrieve emails from the Outlook inbox.
2. Process any resumes attached to those emails.
3. Submit the processed data to Salesforce.
4. Log processing results and move emails to the appropriate folders.

## Error Handling

- Emails that fail to process correctly will be moved to the "Not Completed" folder, and an error log will be generated.
- The service logs both successful and failed form submissions.

## Contribution

1. Fork the repository.
2. Create a feature branch.
3. Submit a pull request.

## License

_**(Specify License if applicable)**_
