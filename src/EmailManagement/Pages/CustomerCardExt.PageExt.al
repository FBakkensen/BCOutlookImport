pageextension 50200 "Customer Card Email Ext" extends "Customer Card"
{
    layout
    {
        addfirst(factboxes)
        {
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