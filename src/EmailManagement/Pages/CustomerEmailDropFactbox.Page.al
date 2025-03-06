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
                    CurrPage.EmailDrop.SetCustomerId(Rec."Customer No.");
                    CurrPage.EmailDrop.SetPlaceholderText('Drop Outlook email here (.eml or .msg)');
                end;

                trigger FileDropped(FileName: Text; FileExtension: Text; FileContent: Text; FileSize: Integer)
                var
                    EmailImportMgt: Codeunit "Email Import Management";
                begin
                    if Rec."Customer No." = '' then
                        Error('Please select a customer before dropping emails.');

                    // This is just an initial handling - the real processing happens in the other events
                    EmailImportMgt.InitializeSimpleEmailImport(Rec."Customer No.", FileName, FileExtension, FileContent);
                    CurrPage.Update(false);
                    Message('Email "%1" imported successfully.', FileName);
                end;

                trigger EmailParsed(FileName: Text; FileExtension: Text; FileContent: Text;
                              Subject: Text; SenderEmail: Text; SenderName: Text;
                              ReceivedDate: DateTime; HasAttachments: Boolean)
                var
                    EmailImportMgt: Codeunit "Email Import Management";
                begin
                    if Rec."Customer No." = '' then
                        Error('Please select a customer before processing email.');

                    // Initialize the email import with parsed data
                    EmailImportMgt.InitializeEmailImport(
                        Rec."Customer No.",
                        FileName,
                        FileExtension,
                        FileContent,
                        Subject,
                        SenderEmail,
                        SenderName,
                        ReceivedDate,
                        HasAttachments
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
}