# Certificate Generation and Email Bot

This Automation Anywhere bot takes an Excel file containing recipient data, generates certificates from a template, and sends them to the recipients via email.

## Prerequisites

- Automation Anywhere Enterprise Client
- Microsoft Excel
- Microsoft Word
- Email account configured in Automation Anywhere

## Bot Overview

The bot performs the following steps:
1. Creates a folder for storing generated certificates.
2. Opens the Excel file containing recipient data.
3. Loops through each row in the Excel file.
4. Copies the certificate template for each recipient.
5. Replaces placeholder text in the template with recipient-specific data.
6. Converts the Word document to PDF.
7. Sends the PDF certificate to the recipient via email.
8. Cleans up temporary files.

## Setup Instructions

1. **Create Folder**: 
   - Command: `Folder -> createFolder`
   - Parameters: `folderPath`, `isOverwrite = true`

2. **Open Excel Spreadsheet**: 
   - Command: `Excel_MS -> OpenSpreadsheet`
   - Parameters: `excelSourceOption = desktopfilepath`, `filePath = dir/CertificateData.xlsx`, `containsHeader = true`, `isSpecificSheet = false`, `fileAccessMode = EDIT`, `isSecure = false`, `loadAddIns = false`, `excludeHiddenSheets = false`, `containsChart = false`

3. **Loop Through Excel Rows**: 
   - Command: `Loop -> loop.commands.start`

4. **Copy Certificate Template**: 
   - Command: `File -> copyFiles`
   - Parameters: `sourceFilePath = dir/TemplateCertificateofAppreciation.docx`, `destinationPath`, `isOverwrite = true`, `isParallel = false`, `isSize = false`, `isDate = false`

5. **Replace Text in Template**: 
   - Command: `MSWordPackage -> ReplaceText`
   - Parameters: `filePath`, `replaceText`, `newText`
   - Repeat for each placeholder (e.g., Title, FirstName, LastName, DD, MM, YYYY, CompanyName)

6. **Convert Word Document to PDF**: 
   - Command: `FileConversion -> DOCXtoPDF`
   - Parameters: `inputFile`, `outputPath`

7. **Send Email with PDF Attachment**: 
   - Command: `Email -> SendMailV2`
   - Parameters: `toAddress`, `cc`, `bcc`, `subject = You are awesome!`, `fileList`

8. **Clean Up Temporary Files**: 
   - Command: `File -> deleteFiles`
   - Parameters: `filePath`, `isSize = false`, `isDate = false`

## Detailed Steps

1. **Create Folder**: 
   - Create a folder to store the generated certificates.

2. **Open Excel Spreadsheet**: 
   - Open the Excel file containing recipient data.

3. **Loop Through Excel Rows**: 
   - Loop through each row in the Excel file to process recipient data.

4. **Copy Certificate Template**: 
   - Copy the certificate template for each recipient.

5. **Replace Text in Template**: 
   - Replace placeholder text in the template with recipient-specific data (e.g., Title, FirstName, LastName, Date, CompanyName).

6. **Convert Word Document to PDF**: 
   - Convert the Word document to a PDF file.

7. **Send Email with PDF Attachment**: 
   - Send the PDF certificate to the recipient via email.

8. **Clean Up Temporary Files**: 
   - Delete temporary files created during the process.

---

Feel free to customize this README further based on your specific requirements!
