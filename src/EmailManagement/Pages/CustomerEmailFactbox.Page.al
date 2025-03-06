page 50200 "Customer Email Factbox"
{
    PageType = ListPart;
    SourceTable = "Outlook Email";
    Caption = 'Customer Emails';
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

                trigger DropError(ErrorMessage: Text)
                begin
                    Error(ErrorMessage);
                end;
            }

            repeater(Emails)
            {
                Caption = 'Emails';
                FreezeColumn = "Email Subject";

                field("Entry No."; Rec."Entry No.")
                {
                    ApplicationArea = All;
                    Visible = false;
                }
                field("Email Subject"; Rec."Email Subject")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the subject of the email.';
                    Style = Strong;
                }
                field("Sender Email"; Rec."Sender Email")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the email address of the sender.';
                }
                field("Sender Name"; Rec."Sender Name")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the sender.';
                    Visible = false;
                }
                field("Received Date"; Rec."Received Date")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies when the email was received.';
                }
                field("Has Attachments"; Rec."Has Attachments")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies whether the email has attachments.';
                }
                field("Created At"; Rec."Created At")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies when the email was imported.';
                    Visible = false;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(ViewEmail)
            {
                ApplicationArea = All;
                Caption = 'View Email';
                Image = ViewDocumentLine;
                ToolTip = 'View the selected email.';

                trigger OnAction()
                var
                    EmailViewPage: Page "Email View";
                begin
                    EmailViewPage.SetRecord(Rec);
                    EmailViewPage.RunModal();
                end;
            }
        }
    }

    trigger OnOpenPage()
    var
        EmailService: Codeunit "Email Service";
    begin
        if Rec."Customer No." <> '' then
            EmailService.GetCustomerEmails(Rec."Customer No.", Rec);
    end;

    trigger OnAfterGetRecord()
    begin
        CurrPage.EmailDrop.SetCustomerId(Rec."Customer No.");
    end;
}