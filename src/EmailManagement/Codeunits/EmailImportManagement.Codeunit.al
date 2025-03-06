codeunit 50202 "Email Import Management"
{
    Access = Public;
    SingleInstance = true;

    var
        TempEmailImport: Record "Outlook Email" temporary;
        TempAttachmentImports: Record "Outlook Email Attachment" temporary;
        EmailService: Codeunit "Email Service";
        IsInitialized: Boolean;

    /// <summary>
    /// Initializes a simple email import process with basic file data
    /// </summary>
    procedure InitializeSimpleEmailImport(CustomerNo: Code[20]; FileName: Text; FileExtension: Text; FileContent: Text)
    var
        OutlookEmail: Record "Outlook Email";
        Base64Convert: Codeunit "Base64 Convert";
        OutStream: OutStream;
        Subject: Text;
    begin
        // Extract subject from filename if possible
        Subject := GetSubjectFromFileName(FileName);

        // Create the email record directly
        OutlookEmail.Init();
        OutlookEmail."Customer No." := CustomerNo;
        OutlookEmail."File Name" := CopyStr(FileName, 1, MaxStrLen(OutlookEmail."File Name"));
        OutlookEmail."File Extension" := CopyStr(FileExtension, 1, MaxStrLen(OutlookEmail."File Extension"));
        OutlookEmail."Email Subject" := CopyStr(Subject, 1, MaxStrLen(OutlookEmail."Email Subject"));
        OutlookEmail."Received Date" := CurrentDateTime;

        // Store the email content
        OutlookEmail."Email Content".CreateOutStream(OutStream);
        Base64Convert.FromBase64(FileContent, OutStream);

        // Insert the email record
        OutlookEmail.Insert(true);
    end;

    local procedure GetSubjectFromFileName(FileName: Text): Text
    begin
        // Remove file extension if present
        if StrPos(FileName, '.') > 0 then
            exit(CopyStr(FileName, 1, StrPos(FileName, '.') - 1));

        exit(FileName);
    end;

    /// <summary>
    /// Initializes a new email import process with the parsed email data
    /// </summary>
    procedure InitializeEmailImport(CustomerNo: Code[20]; FileName: Text; FileExtension: Text; FileContent: Text;
                                  Subject: Text; SenderEmail: Text; SenderName: Text;
                                  ReceivedDate: DateTime; HasAttachments: Boolean; BodyContent: Text)
    var
        Base64Convert: Codeunit "Base64 Convert";
        OutStream: OutStream;
        BodyOutStream: OutStream;
    begin
        // Log initialization
        Message('Initializing email import for: %1', FileName);

        // Ensure clean state - reset all temporary data
        Clear(TempEmailImport);
        TempEmailImport.DeleteAll();
        TempAttachmentImports.Reset();
        TempAttachmentImports.DeleteAll();

        // Prepare the email record with parsed data
        TempEmailImport.Init();
        TempEmailImport."Customer No." := CustomerNo;
        TempEmailImport."File Name" := CopyStr(FileName, 1, MaxStrLen(TempEmailImport."File Name"));
        TempEmailImport."File Extension" := CopyStr(FileExtension, 1, MaxStrLen(TempEmailImport."File Extension"));
        TempEmailImport."Email Subject" := CopyStr(Subject, 1, MaxStrLen(TempEmailImport."Email Subject"));
        TempEmailImport."Sender Email" := CopyStr(SenderEmail, 1, MaxStrLen(TempEmailImport."Sender Email"));
        TempEmailImport."Sender Name" := CopyStr(SenderName, 1, MaxStrLen(TempEmailImport."Sender Name"));
        TempEmailImport."Received Date" := ReceivedDate;
        TempEmailImport."Has Attachments" := HasAttachments;

        // Store the email content
        TempEmailImport."Email Content".CreateOutStream(OutStream);
        Base64Convert.FromBase64(FileContent, OutStream);

        // Store the email body
        if BodyContent <> '' then begin
            TempEmailImport."Email Body".CreateOutStream(BodyOutStream);
            BodyOutStream.WriteText(BodyContent);
        end;

        TempEmailImport.Insert();
        IsInitialized := true;
    end;

    /// <summary>
    /// Adds an attachment to the current email import process
    /// </summary>
    procedure AddAttachmentToImport(AttachmentFileName: Text; FileExtension: Text; MimeType: Text;
                                  FileContent: Text; FileSize: Integer)
    var
        Base64Convert: Codeunit "Base64 Convert";
        OutStream: OutStream;
    begin
        if not IsInitialized then
            Error('Email import not initialized. Call InitializeEmailImport first.');

        TempAttachmentImports.Init();
        TempAttachmentImports."Entry No." := GetNextAttachmentEntryNo();
        TempAttachmentImports."File Name" := CopyStr(AttachmentFileName, 1, MaxStrLen(TempAttachmentImports."File Name"));
        TempAttachmentImports."File Extension" := CopyStr(FileExtension, 1, MaxStrLen(TempAttachmentImports."File Extension"));
        TempAttachmentImports."MIME Type" := CopyStr(MimeType, 1, MaxStrLen(TempAttachmentImports."MIME Type"));
        TempAttachmentImports."File Size (KB)" := Round(FileSize / 1024, 0.01);

        // Store the attachment content
        TempAttachmentImports."Attachment Content".CreateOutStream(OutStream);
        Base64Convert.FromBase64(FileContent, OutStream);

        TempAttachmentImports.Insert();
    end;

    /// <summary>
    /// Finalizes the email import process and saves all data to the database
    /// </summary>
    /// <returns>The entry number of the created email record</returns>
    procedure FinalizeEmailImport(): Integer
    var
        OutlookEmail: Record "Outlook Email";
        OutlookEmailAttachment: Record "Outlook Email Attachment";
        InStream: InStream;
        OutStream: OutStream;
        EmailEntryNo: Integer;
    begin
        if not IsInitialized then
            Error('Email import not initialized. Call InitializeEmailImport first.');

        // Log finalization
        Message('Finalizing email import for: %1', TempEmailImport."File Name");

        // Create the email record
        OutlookEmail.Init();
        OutlookEmail."Customer No." := TempEmailImport."Customer No.";
        OutlookEmail."Email Subject" := TempEmailImport."Email Subject";
        OutlookEmail."Sender Email" := TempEmailImport."Sender Email";
        OutlookEmail."Sender Name" := TempEmailImport."Sender Name";
        OutlookEmail."Received Date" := TempEmailImport."Received Date";
        OutlookEmail."File Name" := TempEmailImport."File Name";
        OutlookEmail."File Extension" := TempEmailImport."File Extension";
        OutlookEmail."Has Attachments" := TempEmailImport."Has Attachments";

        // Copy email content
        TempEmailImport.CalcFields("Email Content");
        if TempEmailImport."Email Content".HasValue then begin
            TempEmailImport."Email Content".CreateInStream(InStream);
            OutlookEmail."Email Content".CreateOutStream(OutStream);
            CopyStream(OutStream, InStream);
        end;

        // Copy email body if exists
        TempEmailImport.CalcFields("Email Body");
        if TempEmailImport."Email Body".HasValue then begin
            TempEmailImport."Email Body".CreateInStream(InStream);
            OutlookEmail."Email Body".CreateOutStream(OutStream);
            CopyStream(OutStream, InStream);
        end;

        // Insert the email record and get the entry no
        OutlookEmail.Insert(true);
        EmailEntryNo := OutlookEmail."Entry No.";

        // Process attachments if any
        if TempAttachmentImports.FindSet() then
            repeat
                // Create a new record for each attachment
                Clear(OutlookEmailAttachment);
                OutlookEmailAttachment.Init();
                // Don't set Entry No, let AutoIncrement handle it
                OutlookEmailAttachment."Email Entry No." := EmailEntryNo;
                OutlookEmailAttachment."File Name" := TempAttachmentImports."File Name";
                OutlookEmailAttachment."File Extension" := TempAttachmentImports."File Extension";
                OutlookEmailAttachment."MIME Type" := TempAttachmentImports."MIME Type";
                OutlookEmailAttachment."File Size (KB)" := TempAttachmentImports."File Size (KB)";

                // Copy attachment content
                TempAttachmentImports.CalcFields("Attachment Content");
                if TempAttachmentImports."Attachment Content".HasValue then begin
                    TempAttachmentImports."Attachment Content".CreateInStream(InStream);
                    OutlookEmailAttachment."Attachment Content".CreateOutStream(OutStream);
                    CopyStream(OutStream, InStream);
                end;

                // Insert the attachment record
                OutlookEmailAttachment.Insert(true);
            until TempAttachmentImports.Next() = 0;

        // Clean up and reset state
        Clear(TempEmailImport);
        TempEmailImport.DeleteAll();
        TempAttachmentImports.Reset();
        TempAttachmentImports.DeleteAll();
        IsInitialized := false;

        exit(EmailEntryNo);
    end;

    /// <summary>
    /// Gets the next entry number for temporary attachment records
    /// </summary>
    local procedure GetNextAttachmentEntryNo(): Integer
    var
        MaxEntryNo: Integer;
    begin
        TempAttachmentImports.Reset();
        if TempAttachmentImports.FindLast() then
            MaxEntryNo := TempAttachmentImports."Entry No.";

        exit(MaxEntryNo + 1);
    end;
}