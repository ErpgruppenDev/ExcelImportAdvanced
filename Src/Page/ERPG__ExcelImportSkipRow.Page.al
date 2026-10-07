page 68762 "ERPG ExcelImport_SkipRow"
{
    ApplicationArea = All;
    Caption = 'ExcelImport_SkipRow';
    PageType = ListPart;
    SourceTable = "ERPG ExcelImport_SkipRow";

    layout
    {
        area(Content)
        {
            repeater(General)
            {
                field("Skip Row No"; Rec."Skip Row No")
                {
                    ToolTip = 'Specifies the value of the Skip Row No field.', Comment = '%';
                }
            }
        }
    }
}
