# Changelog

All notable changes to the "BC Outlook Email Import" extension will be documented in this file.

## [1.3.3] - 2025-03-06

### Fixed
- Fixed issue with email body displayed as Base64-encoded text instead of readable content
- Added detection and decoding of Base64-encoded content in email bodies
- Enhanced content block processing to handle Base64 encoding within multipart messages
- Improved email format detection for Microsoft Exchange emails with encoded content
- Added more detailed logging for Base64 content detection and decoding

## [1.3.2] - 2025-03-06

### Fixed
- Fixed issue with "boundary=" text appearing in email body for Microsoft Exchange/Outlook emails
- Added special handling for Microsoft-specific boundary formats
- Implemented more robust boundary detection and cleanup algorithms
- Added extensive logging for easier diagnostics of email parsing issues
- Enhanced content extraction with Microsoft Exchange specific content block handling

## [1.3.1] - 2025-03-06

### Fixed
- Fixed email body extraction for Microsoft Exchange/Outlook format emails
- Enhanced boundary detection for special formatted boundary markers
- Improved MIME multipart parsing for complex email formats
- Added more robust content extraction algorithms for difficult email formats
- Added additional cleanup of boundary markers in the final text output

## [1.3.0] - 2025-03-06

### Added
- Added functionality to delete emails and their related attachments
- Added delete action to the Email View page with confirmation dialog
- Added delete action to the Customer Email Factbox for quick deletion
- Implemented cascading deletion to ensure all related data is removed

## [1.2.1] - 2025-03-06

### Fixed
- Fixed email body extraction to properly handle multipart MIME emails
- Improved boundary detection for complex email formats
- Added support for quoted-printable encoding in email content
- Enhanced content extraction to handle both plain text and HTML parts
- Improved whitespace handling and formatting of extracted email bodies

## [1.2.0] - 2025-03-06

### Added
- Added email body storage and display capability
- Implemented real email content extraction from uploaded .eml and .msg files
- Added comprehensive parsing to extract subject, sender, date, and body from email headers
- Added HTML cleanup for better readability of email body content
- Added email body display in the Email View page
- Improved email content handling for better user experience

## [1.1.9] - 2025-03-06

### Fixed
- Fixed "The record in table Outlook Email Attachment already exists" error when importing files with attachments
- Improved temporary record handling to properly clean up between email imports
- Enhanced attachment processing to respect AutoIncrement behavior of Entry No. field
- Added thorough cleanup procedures to ensure clean state between operations

## [1.1.8] - 2025-03-06

### Fixed
- Fixed persistent "Please select a customer before processing email" error when dropping emails
- Added proper customer selection synchronization between Customer List and Email Drop factbox
- Implemented safe initialization checks to prevent runtime errors
- Added IsReady function to the factbox to safely verify control state before operations

## [1.1.7] - 2025-03-06

### Fixed
- Fixed initialization error "The control add-in on control EmailDrop on page Email Drop has not been instantiated"
- Added control ready state tracking to prevent errors when the page is first loading
- Improved error handling in Customer List page to handle control initialization gracefully
- Enhanced factbox visibility to ensure proper initialization before invoking control methods

## [1.1.6] - 2025-03-06

### Fixed
- Fixed issue where importing emails from Customer List view would show "Please select a customer" error even when a customer was selected
- Improved customer selection logic in the factbox to ensure proper synchronization with the currently selected customer
- Added visibility control for factboxes to hide them when no customer is selected

## [1.1.5] - 2025-03-06

### Added
- Implemented event handlers for `EmailParsed`, `AttachmentParsed`, and `EmailParsingComplete` events
- Added detailed logging in JavaScript for better troubleshooting of email processing
- Added performance optimization for drag-over events to improve responsiveness

### Fixed
- Fixed critical issue where dropped email files were not being processed due to missing event handlers
- Resolved event mapping between JavaScript and AL to ensure proper data flow
- Improved error handling and user feedback during the email import process

## [1.1.4] - 2025-03-06

### Added
- Enhanced visual indicator for the email drop zone to make it more intuitive
- Added SVG icon for improved visibility of the drop target
- Implemented subtle animation to draw attention to the drop area

### Changed
- Adjusted the email drop control dimensions to better fit in the factbox
- Improved visual styling with more distinct colors and borders
- Enhanced user experience with clearer drop instructions

### Fixed
- Fixed issue where the email drop control was not visible in the factbox
- Improved control initialization to ensure consistent rendering

## [1.1.3] - 2025-03-06

### Changed
- Separated user control and data viewing into different factboxes
- Created a new factbox for the user control
- Updated the existing factbox to only display email data
- Updated Customer Card and Customer List page extensions to include both factboxes

## [1.1.2] - 2025-03-06

### Changed
- Updated object IDs to match the new range specified in app.json (50200-50249)

## [1.1.1] - 2025-03-06

### Fixed
- Fixed linting errors in control add-in method definitions
- Updated JavaScript implementation to use events for returning values
- Corrected UpdatePropagation property in page extensions to use 'SubPart' value
- Removed promotion properties from ListPart actions
- Simplified user control event handling

## [1.1.0] - 2025-03-06

### Changed
- Moved email parsing logic from AL to JavaScript for better performance
- Updated control add-in to handle email metadata extraction and attachment parsing
- Renamed events in the control add-in to better reflect their purpose
- Implemented SOLID principles for better maintainability
- Created EmailImportManagement codeunit to handle the import process
- Simplified the EmailService codeunit to focus on data retrieval

### Removed
- Removed EmailParser codeunit as parsing is now done in JavaScript

## [1.0.0] - 2025-03-06

### Added
- Initial release of the BC Outlook Email Import extension
- Created tables for storing emails and attachments
- Implemented drag-and-drop functionality for emails using a custom control add-in
- Added factbox to Customer Card and Customer List pages
- Added email viewer with attachment support
- Implemented email metadata parsing and storage
- Added ability to download original emails and attachments