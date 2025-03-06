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

            action(DeleteEmail)
            {
                ApplicationArea = All;
                Caption = 'Delete Email';
                Image = Delete;
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
                        CurrPage.Update(false);
                    end else
                        Error(ErrorMsg);
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
}