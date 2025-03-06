page 50203 "Customer Email Drop Factbox"
{
    PageType = ListPart;
    SourceTable = "Outlook Email";
    Caption = 'Email Drop';
    InsertAllowed = false;
    ModifyAllowed = false;
    Editable = false;

    layout
    {
        area(Content)
        {
            usercontrol(EmailDrop; "Outlook Email Drop")
            {
                ApplicationArea = All;

                trigger ControlReady()
                begin
                    ControlIsReady := true;
                    CurrPage.EmailDrop.SetCustomerId(CustomerNo);
                    CurrPage.EmailDrop.SetPlaceholderText('Drop Outlook email here (.eml or .msg)');
                end;

                trigger FileDropped(FileName: Text; FileExtension: Text; FileContent: Text; FileSize: Integer)
                var
                    EmailImportMgt: Codeunit "Email Import Management";
                begin
                    // Check if we have a valid customer
                    if CustomerNo = '' then
                        Error('Please select a customer before dropping emails.');

                    // This is just an initial handling - the real processing happens in the other events
                    EmailImportMgt.InitializeSimpleEmailImport(CustomerNo, FileName, FileExtension, FileContent);
                    CurrPage.Update(false);
                    Message('Email "%1" imported successfully.', FileName);
                end;

                trigger EmailParsed(FileName: Text; FileExtension: Text; FileContent: Text;
                              Subject: Text; SenderEmail: Text; SenderName: Text;
                              ReceivedDate: DateTime; HasAttachments: Boolean; Body: Text)
                var
                    EmailImportMgt: Codeunit "Email Import Management";
                begin

                    // Check if we have a valid customer
                    if CustomerNo = '' then
                        Error('Please select a customer before processing email.');

                    // Initialize the email import with parsed data
                    EmailImportMgt.InitializeEmailImport(
                        CustomerNo,
                        FileName,
                        FileExtension,
                        FileContent,
                        Subject,
                        SenderEmail,
                        SenderName,
                        ReceivedDate,
                        HasAttachments,
                        Body
                    );

                    // Display a message about the parsed email
                    Message('Email from %1 (%2) parsed successfully.', SenderName, SenderEmail);
                end;

                trigger EmailParsingComplete()
                var
                    EmailImportMgt: Codeunit "Email Import Management";
                begin
                    // Finalize the email import process
                    EmailImportMgt.FinalizeEmailImport();
                    CurrPage.Update(false);
                end;

                trigger AttachmentParsed(EmailFileName: Text; AttachmentFileName: Text;
                                       FileExtension: Text; MimeType: Text;
                                       FileContent: Text; FileSize: Integer)
                var
                    EmailImportMgt: Codeunit "Email Import Management";
                begin
                    // Add attachment to the current email import
                    EmailImportMgt.AddAttachmentToImport(
                        AttachmentFileName,
                        FileExtension,
                        MimeType,
                        FileContent,
                        FileSize
                    );
                end;

                trigger DropError(ErrorMessage: Text)
                begin
                    Error(ErrorMessage);
                end;
            }
        }
    }

    var
        ControlIsReady: Boolean;
        CustomerNo: Text;

    // Add a procedure to set the current customer number
    procedure SetCustomerNo(_CustomerNo: Text)
    begin
        CustomerNo := _CustomerNo;

        // Only update the control if it's ready
        if ControlIsReady then
            CurrPage.EmailDrop.SetCustomerId(_CustomerNo);
    end;

    // Add a procedure to check if the control is ready
    procedure IsReady(): Boolean
    begin
        exit(ControlIsReady);
    end;
}