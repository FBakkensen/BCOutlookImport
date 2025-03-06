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
        }
    }
}