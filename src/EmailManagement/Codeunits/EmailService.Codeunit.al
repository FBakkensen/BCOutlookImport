codeunit 50200 "Email Service"
{
    Access = Public;

    /// <summary>
    /// Gets all emails for a specific customer
    /// </summary>
    /// <param name="CustomerNo">The customer number</param>
    /// <param name="OutlookEmails">Output parameter that will contain the emails</param>
    procedure GetCustomerEmails(CustomerNo: Code[20]; var OutlookEmails: Record "Outlook Email")
    begin
        OutlookEmails.Reset();
        OutlookEmails.SetRange("Customer No.", CustomerNo);
        OutlookEmails.SetCurrentKey("Received Date");
        OutlookEmails.Ascending(false);
    end;

    /// <summary>
    /// Gets all attachments for a specific email
    /// </summary>
    /// <param name="EmailEntryNo">The entry number of the email</param>
    /// <param name="OutlookEmailAttachments">Output parameter that will contain the attachments</param>
    procedure GetEmailAttachments(EmailEntryNo: Integer; var OutlookEmailAttachments: Record "Outlook Email Attachment")
    begin
        OutlookEmailAttachments.Reset();
        OutlookEmailAttachments.SetRange("Email Entry No.", EmailEntryNo);
    end;

    /// <summary>
    /// Retrieves an email file as a stream for download
    /// </summary>
    /// <param name="EmailEntryNo">The entry number of the email</param>
    /// <param name="InStream">Output parameter that will contain the email file stream</param>
    /// <param name="FileName">Output parameter that will contain the filename</param>
    /// <returns>True if the email content was successfully retrieved</returns>
    procedure DownloadEmailFile(EmailEntryNo: Integer; var InStream: InStream; var FileName: Text): Boolean
    var
        OutlookEmail: Record "Outlook Email";
    begin
        if not OutlookEmail.Get(EmailEntryNo) then
            exit(false);

        OutlookEmail.CalcFields("Email Content");
        if not OutlookEmail."Email Content".HasValue then
            exit(false);

        OutlookEmail."Email Content".CreateInStream(InStream);
        FileName := OutlookEmail."File Name";
        exit(true);
    end;

    /// <summary>
    /// Retrieves an email attachment file as a stream for download
    /// </summary>
    /// <param name="AttachmentEntryNo">The entry number of the attachment</param>
    /// <param name="InStream">Output parameter that will contain the attachment file stream</param>
    /// <param name="FileName">Output parameter that will contain the filename</param>
    /// <returns>True if the attachment content was successfully retrieved</returns>
    procedure DownloadAttachmentFile(AttachmentEntryNo: Integer; var InStream: InStream; var FileName: Text): Boolean
    var
        OutlookEmailAttachment: Record "Outlook Email Attachment";
    begin
        if not OutlookEmailAttachment.Get(AttachmentEntryNo) then
            exit(false);

        OutlookEmailAttachment.CalcFields("Attachment Content");
        if not OutlookEmailAttachment."Attachment Content".HasValue then
            exit(false);

        OutlookEmailAttachment."Attachment Content".CreateInStream(InStream);
        FileName := OutlookEmailAttachment."File Name";
        exit(true);
    end;

    /// <summary>
    /// Deletes an email and all its related attachments
    /// </summary>
    /// <param name="EmailEntryNo">The entry number of the email to delete</param>
    /// <returns>True if the email was successfully deleted, false otherwise</returns>
    procedure DeleteEmail(EmailEntryNo: Integer): Boolean
    var
        OutlookEmail: Record "Outlook Email";
        OutlookEmailAttachments: Record "Outlook Email Attachment";
    begin
        // First, check if the email exists
        if not OutlookEmail.Get(EmailEntryNo) then
            exit(false);

        // Delete all related attachments first
        OutlookEmailAttachments.Reset();
        OutlookEmailAttachments.SetRange("Email Entry No.", EmailEntryNo);
        if OutlookEmailAttachments.FindSet() then
            OutlookEmailAttachments.DeleteAll(true);

        // Then delete the email itself
        exit(OutlookEmail.Delete(true));
    end;
}