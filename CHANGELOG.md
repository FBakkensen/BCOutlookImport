# Changelog

All notable changes to the "BC Outlook Email Import" extension will be documented in this file.

## [1.1.3] - 2023-10-15

### Changed
- Separated user control and data viewing into different factboxes
- Created a new factbox for the user control
- Updated the existing factbox to only display email data
- Updated Customer Card and Customer List page extensions to include both factboxes

## [1.1.2] - 2023-10-15

### Changed
- Updated object IDs to match the new range specified in app.json (50200-50249)

## [1.1.1] - 2023-10-15

### Fixed
- Fixed linting errors in control add-in method definitions
- Updated JavaScript implementation to use events for returning values
- Corrected UpdatePropagation property in page extensions to use 'SubPart' value
- Removed promotion properties from ListPart actions
- Simplified user control event handling

## [1.1.0] - 2023-10-12

### Changed
- Moved email parsing logic from AL to JavaScript for better performance
- Updated control add-in to handle email metadata extraction and attachment parsing
- Renamed events in the control add-in to better reflect their purpose
- Implemented SOLID principles for better maintainability
- Created EmailImportManagement codeunit to handle the import process
- Simplified the EmailService codeunit to focus on data retrieval

### Removed
- Removed EmailParser codeunit as parsing is now done in JavaScript

## [1.0.0] - 2023-10-05

### Added
- Initial release of the BC Outlook Email Import extension
- Created tables for storing emails and attachments
- Implemented drag-and-drop functionality for emails using a custom control add-in
- Added factbox to Customer Card and Customer List pages
- Added email viewer with attachment support
- Implemented email metadata parsing and storage
- Added ability to download original emails and attachments