page 68763 "ERPG ExcelImp. List, Setup"
{
    Caption = 'Excel Import List, Setup', Comment = 'DAN="Excel import Liste, Opsætning"';
    CardPageID = "ERPG Excel import Setup";
    Editable = false;
    PageType = List;
    SourceTable = "ERPG Excel Import Setup";
    UsageCategory = Tasks;
    ApplicationArea = All;

    layout
    {
        area(content)
        {
            repeater(Group)
            {
                Field(Code; Rec.Code)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify code', Comment = 'DAN="Angiver koden"';

                }
                Field(Description; Rec.Description)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify description', Comment = 'DAN="Angiver beskrivelsen"';
                }
                Field("No. of Columns"; Rec."No. of Columns")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify Quantity of colunms in Excel file', Comment = 'DAN="Angiver hvor mange kolonner der er i Excel fil"';
                }
                field(ImportToTable; Rec.ImportToTable) { ApplicationArea = All; Visible = true; }
                Field(TabelNo; Rec.TableNo)
                {
                    Visible = true;
                    ApplicationArea = All;
                    ToolTip = 'Specify tableno.', Comment = 'DAN="Angiver tabelnr."';
                }
                Field(Tablename; Rec.Tablename)
                {
                    Visible = true;
                    ApplicationArea = All;
                    ToolTip = 'Specify tablename', Comment = 'DAN="Angiver tabelnavn"';
                }
                field(JournalTemplate; Rec.JournalTemplate) { ApplicationArea = All; Visible = true; }
                field(JournalBatch; Rec.JournalBatch) { ApplicationArea = All; Visible = true; }
                // Field("Gen. Journal Template"; Rec."Gen. Journal Template")
                // {
                //     ApplicationArea = All;
                //     ToolTip = 'Specify Journal Template name', Comment = 'DAN="Angiver kladdetype"';

                // }
                // Field("Gen. Journal Batch"; Rec."Gen. Journal Batch")
                // {
                //     ApplicationArea = All;
                //     ToolTip = 'Specify Batch Name', Comment = 'DAN="Angiver Kladdenavn"';

                // }
                Field("Use Postingdate"; Rec."Use Postingdate")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify postingdate to use in import', Comment = 'DAN="Angiver bogføringsdato der angives i kladden, ved import"';

                }
                Field("Last modified date"; Rec."Last modified date")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the date of latest modfication', Comment = 'DAN="Angiver hvilken dato, data sidst er rettet"';
                }
                Field("Last modified by"; Rec."Last modified by")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the User who lastest has modified data', Comment = 'DAN="Angiver hvem der senest har rettet data"';
                }
            }
        }
    }

    actions
    {
    }
}

