page 65004 ERPG_ExcelImport_SkipRow
{
    ApplicationArea = All;
    Caption = 'ExcelImport_SkipRow';
    PageType = ListPart;
    SourceTable = ERPG_ExcelImport_SkipRow;

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
