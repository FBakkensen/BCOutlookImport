# BC Outlook Email Import Extension

This extension adds functionality to Business Central to import emails directly from Microsoft Outlook via drag and drop.

## Features

- Drag and drop Outlook emails (.eml or .msg files) into a factbox on Customer Card and Customer List pages
- Emails are stored with metadata (subject, sender, date) and linked to the specific customer
- Email attachments are extracted and stored separately
- View email details and download the original email file or attachments

## Requirements

- Business Central 2022 Wave 2 (BC21) or later
- Modern client (web browser)
- Microsoft Outlook

## Usage

1. Open a Customer Card or Customer List
2. Find the "Customer Emails" factbox 
3. Drag an email from Outlook and drop it onto the drop zone in the factbox
4. The email will be processed and appear in the list
5. Click "View Email" to see details and attachments

## Technical Implementation

The extension is built using these components:

- Tables for storing emails and their attachments
- Control Add-in for drag and drop functionality using JavaScript
- Browser-side email parsing to extract metadata and attachments
- Factbox on Customer pages
- Clean separation between UI (parsing) and business logic (storage)

### Architecture

The solution follows SOLID principles:

1. **Single Responsibility**: Each component has a single responsibility
   - JavaScript control handles file dropping and parsing
   - EmailImportManagement codeunit manages the import process
   - EmailService codeunit handles data retrieval

2. **Open/Closed**: The system is open for extension but closed for modification
   - New email formats can be added by extending the JavaScript parsing

3. **Interface Segregation**: Components expose only what's needed
   - The UI only needs to know about the import process, not the storage details

4. **Dependency Inversion**: High-level modules don't depend on low-level modules
   - Business logic doesn't depend on how emails are parsed

## Object IDs

The extension uses object IDs in the range 50200-50249:

- Tables: 50200-50201
- Pages: 50200-50202
- Page Extensions: 50200-50201
- Codeunits: 50200, 50202