pageextension 50201 "Customer List Email Ext" extends "Customer List"
{
    layout
    {
        addfirst(factboxes)
        {
            part(CustomerEmailDropFactbox; "Customer Email Drop Factbox")
            {
                ApplicationArea = All;
                Caption = 'Email Drop';
                SubPageLink = "Customer No." = field("No.");
                UpdatePropagation = SubPart;
            }
            part(CustomerEmailFactbox; "Customer Email Factbox")
            {
                ApplicationArea = All;
                Caption = 'Customer Emails';
                SubPageLink = "Customer No." = field("No.");
                UpdatePropagation = SubPart;
            }
        }
    }
}