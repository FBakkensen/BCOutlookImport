table 50200 "Outlook Email"
{
    Caption = 'Outlook Email';
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
        field(2; "Customer No."; Code[20])
        {
            Caption = 'Customer No.';
            DataClassification = CustomerContent;
            TableRelation = Customer."No.";
        }
        field(3; "Email Subject"; Text[250])
        {
            Caption = 'Email Subject';
            DataClassification = CustomerContent;
        }
        field(4; "Sender Email"; Text[250])
        {
            Caption = 'Sender Email';
            DataClassification = CustomerContent;
        }
        field(5; "Sender Name"; Text[100])
        {
            Caption = 'Sender Name';
            DataClassification = CustomerContent;
        }
        field(6; "Received Date"; DateTime)
        {
            Caption = 'Received Date';
            DataClassification = CustomerContent;
        }
        field(7; "Email Content"; Blob)
        {
            Caption = 'Email Content';
            DataClassification = CustomerContent;
        }
        field(8; "File Name"; Text[250])
        {
            Caption = 'File Name';
            DataClassification = CustomerContent;
        }
        field(9; "File Extension"; Text[10])
        {
            Caption = 'File Extension';
            DataClassification = CustomerContent;
        }
        field(10; "Has Attachments"; Boolean)
        {
            Caption = 'Has Attachments';
            DataClassification = CustomerContent;
        }
        field(11; "Created By"; Code[50])
        {
            Caption = 'Created By';
            DataClassification = EndUserIdentifiableInformation;
            Editable = false;
        }
        field(12; "Created At"; DateTime)
        {
            Caption = 'Created At';
            DataClassification = SystemMetadata;
            Editable = false;
        }
        field(13; "Email Body"; Blob)
        {
            Caption = 'Email Body';
            DataClassification = CustomerContent;
        }
    }

    keys
    {
        key(PK; "Entry No.")
        {
            Clustered = true;
        }
        key(CustomerIndex; "Customer No.")
        {
        }
    }

    trigger OnInsert()
    begin
        "Created By" := UserId;
        "Created At" := CurrentDateTime;
    end;
}