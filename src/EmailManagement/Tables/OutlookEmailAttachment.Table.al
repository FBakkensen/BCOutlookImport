table 50201 "Outlook Email Attachment"
{
    Caption = 'Outlook Email Attachment';
    DataClassification = CustomerContent;

    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
            DataClassification = SystemMetadata;
            Editable = false;
            AutoIncrement = true;
        }
        field(2; "Email Entry No."; Integer)
        {
            Caption = 'Email Entry No.';
            DataClassification = SystemMetadata;
            TableRelation = "Outlook Email"."Entry No.";
        }
        field(3; "File Name"; Text[250])
        {
            Caption = 'File Name';
            DataClassification = CustomerContent;
        }
        field(4; "File Extension"; Text[10])
        {
            Caption = 'File Extension';
            DataClassification = CustomerContent;
        }
        field(5; "File Size (KB)"; Decimal)
        {
            Caption = 'File Size (KB)';
            DataClassification = CustomerContent;
            DecimalPlaces = 0 : 2;
        }
        field(6; "Attachment Content"; Blob)
        {
            Caption = 'Attachment Content';
            DataClassification = CustomerContent;
        }
        field(7; "MIME Type"; Text[100])
        {
            Caption = 'MIME Type';
            DataClassification = SystemMetadata;
        }
        field(8; "Created At"; DateTime)
        {
            Caption = 'Created At';
            DataClassification = SystemMetadata;
            Editable = false;
        }
    }

    keys
    {
        key(PK; "Entry No.")
        {
            Clustered = true;
        }
        key(EmailIndex; "Email Entry No.")
        {
        }
    }

    trigger OnInsert()
    begin
        "Created At" := CurrentDateTime;
    end;
}