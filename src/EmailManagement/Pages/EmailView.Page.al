page 50201 "Email View"
{
    PageType = Card;
    SourceTable = "Outlook Email";
    Caption = 'Email View';
    InsertAllowed = false;
    ModifyAllowed = false;
    DeleteAllowed = false;
    Editable = false;

    layout
    {
        area(Content)
        {
            group(General)
            {
                Caption = 'General';

                field("Email Subject"; Rec."Email Subject")
                {
                    ApplicationArea = All;
                    Style = Strong;
                    StyleExpr = true;
                }
                field("Sender Name"; Rec."Sender Name")
                {
                    ApplicationArea = All;
                }
                field("Sender Email"; Rec."Sender Email")
                {
                    ApplicationArea = All;
                }
                field("Received Date"; Rec."Received Date")
                {
                    ApplicationArea = All;
                }
                field("Customer No."; Rec."Customer No.")
                {
                    ApplicationArea = All;
                    DrillDown = true;

                    trigger OnDrillDown()
                    var
                        CustomerCard: Page "Customer Card";
                        Customer: Record Customer;
                    begin
                        if Customer.Get(Rec."Customer No.") then begin
                            CustomerCard.SetRecord(Customer);
                            CustomerCard.Run();
                        end;
                    end;
                }
                field("File Name"; Rec."File Name")
                {
                    ApplicationArea = All;
                }
                field("File Extension"; Rec."File Extension")
                {
                    ApplicationArea = All;
                }
            }

            group(EmailContent)
            {
                Caption = 'Email Body';

                field(EmailBodyText; GetEmailBody())
                {
                    ApplicationArea = All;
                    Caption = 'Body';
                    MultiLine = true;
                    ShowCaption = false;
                }
            }

            part(Attachments; "Email Attachments Subpage")
            {
                ApplicationArea = All;
                Caption = 'Attachments';
                SubPageLink = "Email Entry No." = field("Entry No.");
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(DownloadEmailFile)
            {
                ApplicationArea = All;
                Caption = 'Download Email File';
                Image = Download;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;
                ToolTip = 'Download the original email file.';

                trigger OnAction()
                var
                    EmailService: Codeunit "Email Service";
                    FileManagement: Codeunit "File Management";
                    TempBlob: Codeunit "Temp Blob";
                    FileName: Text;
                    InStream: InStream;
                    OutStream: OutStream;
                begin
                    if not EmailService.DownloadEmailFile(Rec."Entry No.", InStream, FileName) then
                        Error('The email content is not available for download.');

                    TempBlob.CreateOutStream(OutStream);
                    CopyStream(OutStream, InStream);
                    TempBlob.CreateInStream(InStream);

                    FileManagement.BLOBExport(TempBlob, FileName, true);
                end;
            }

            action(DeleteEmail)
            {
                ApplicationArea = All;
                Caption = 'Delete Email';
                Image = Delete;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;
                ToolTip = 'Delete this email and all its attachments.';

                trigger OnAction()
                var
                    EmailService: Codeunit "Email Service";
                    ConfirmMsg: Label 'Are you sure you want to delete this email and all its attachments? This action cannot be undone.';
                    SuccessMsg: Label 'Email and attachments have been deleted successfully.';
                    ErrorMsg: Label 'An error occurred while trying to delete the email.';
                begin
                    if not Confirm(ConfirmMsg, false) then
                        exit;

                    if EmailService.DeleteEmail(Rec."Entry No.") then begin
                        Message(SuccessMsg);
                        CurrPage.Close();
                    end else
                        Error(ErrorMsg);
                end;
            }
        }
    }

    local procedure GetEmailBody(): Text
    var
        InStream: InStream;
        EmailBody: Text;
        CR: Char;
        LF: Char;
    begin
        CR := 13;
        LF := 10;

        Rec.CalcFields("Email Body");
        if not Rec."Email Body".HasValue then
            exit('No email body content available.');

        Rec."Email Body".CreateInStream(InStream);
        InStream.ReadText(EmailBody);

        // Ensure proper line breaks
        EmailBody := EmailBody.Replace(Format(CR) + Format(LF), '<br>');
        EmailBody := EmailBody.Replace(Format(LF), '<br>');

        exit(EmailBody);
    end;
}