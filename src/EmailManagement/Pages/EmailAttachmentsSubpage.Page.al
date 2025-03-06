page 50202 "Email Attachments Subpage"
{
    PageType = ListPart;
    SourceTable = "Outlook Email Attachment";
    Caption = 'Email Attachments';
    InsertAllowed = false;
    ModifyAllowed = false;
    DeleteAllowed = false;
    Editable = false;

    layout
    {
        area(Content)
        {
            repeater(Attachments)
            {
                field("File Name"; Rec."File Name")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the name of the attachment file.';
                }
                field("File Extension"; Rec."File Extension")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the file extension of the attachment.';
                }
                field("File Size (KB)"; Rec."File Size (KB)")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the size of the attachment in KB.';
                }
                field("MIME Type"; Rec."MIME Type")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the MIME type of the attachment.';
                    Visible = false;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(DownloadAttachment)
            {
                ApplicationArea = All;
                Caption = 'Download Attachment';
                Image = Download;
                ToolTip = 'Download the selected attachment.';

                trigger OnAction()
                var
                    EmailService: Codeunit "Email Service";
                    FileManagement: Codeunit "File Management";
                    TempBlob: Codeunit "Temp Blob";
                    FileName: Text;
                    InStream: InStream;
                    OutStream: OutStream;
                begin
                    if not EmailService.DownloadAttachmentFile(Rec."Entry No.", InStream, FileName) then
                        Error('The attachment content is not available for download.');

                    TempBlob.CreateOutStream(OutStream);
                    CopyStream(OutStream, InStream);
                    TempBlob.CreateInStream(InStream);

                    FileManagement.BLOBExport(TempBlob, FileName, true);
                end;
            }
        }
    }
}