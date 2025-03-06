controladdin "Outlook Email Drop"
{
    Scripts = 'src/EmailManagement/Controls/Scripts/OutlookEmailDrop.js';
    StartupScript = 'src/EmailManagement/Controls/Scripts/StartupScript.js';
    StyleSheets = 'src/EmailManagement/Controls/Styles/OutlookEmailDrop.css';

    // Adjusted dimensions for better factbox fit
    MinimumWidth = 250;
    MinimumHeight = 120;
    MaximumHeight = 200;

    // Ensure stretching behavior
    VerticalStretch = true;
    VerticalShrink = false;  // Changed to false to prevent shrinking too small
    HorizontalStretch = true;
    HorizontalShrink = true;

    // Events triggered from JavaScript to AL
    event FileDropped(fileName: Text; fileExtension: Text; fileContent: Text; fileSize: Integer);
    event ControlReady();
    event DropError(errorMessage: Text);
    event EmailParsed(fileName: Text; fileExtension: Text; fileContent: Text;
                    subject: Text; senderEmail: Text; senderName: Text;
                    receivedDate: DateTime; hasAttachments: Boolean);
    event AttachmentParsed(emailFileName: Text; attachmentFileName: Text;
                        fileExtension: Text; mimeType: Text;
                        fileContent: Text; fileSize: Integer);
    event EmailParsingComplete();
    event ControlHeightReturned(height: Integer);

    // Methods called from AL to JavaScript
    procedure SetCustomerId(customerId: Text);
    procedure GetControlHeight();
    procedure SetPlaceholderText(text: Text);
}