# Changelog

All notable changes to the "BC Outlook Email Import" extension will be documented in this file.

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