Page 68764 "ERPG Field list"
{

    SourceTable = Field;
    PageType = List;
    Caption = 'Field list', Comment = 'DAN="Felt liste"';

    layout
    {
        area(content)
        {
            repeater(Group)
            {
                Field("No."; Rec."No.")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the No. of the Field', Comment = 'DAN="Angiver felt nr."';
                }

                Field("Field Caption"; Rec."Field Caption")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the Caption of the Field name', Comment = 'DAN="Angiver felt navn"';
                }
                Field(TableNo; Rec.TableNo)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the No. of the table', Comment = 'DAN="Angiver tabel nr."';
                }
                Field(TableName; Rec.Tablename)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the Table name', Comment = 'DAN="Angiver tabel navn"';
                }
            }
        }
    }


}