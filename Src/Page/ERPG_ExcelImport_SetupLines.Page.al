page 65001 "ERPG Excelimp.lines Setupline"
{
    PageType = ListPart;
    SourceTable = "ERPG ExcelImport_SetupLine";
    Caption = 'Excel Import lines', Comment = 'DAN="Excel import linjer"';
    ApplicationArea = All;
    UsageCategory = Tasks;


    layout
    {
        area(content)
        {
            repeater(Group)
            {
                Field(Code; Rec.Code)
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify code', Comment = 'DAN="Angiver koden"';
                }
                Field(LineNo; Rec.LineNo)
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify lineno.', Comment = 'DAN="Angiver linjenr."';
                }
                Field(Column; Rec.Column)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify column No. in the Excel file (A column is no. 1, B is no. 2 etc)', Comment = 'DAN="Angiver kolonnenr. i Excel fil (A kolonne vil være nr. 1, B nr. 2 osv.)"';
                }
                Field(Description; Rec.Description)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify description', Comment = 'DAN="Angiver beskrivelsen"';
                }
                Field(TabelNo; Rec.TabelNo)
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify tableno.', Comment = 'DAN="Angiver tabelnr."';
                }
                Field(Tablename; Rec.Tablename)
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify tablename', Comment = 'DAN="Angiver tabelnavn"';
                }
                Field(FieldNo; Rec.FieldNo)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify fieldno.', Comment = 'DAN="Angiver feltnr."';
                }
                Field(Fieldname; Rec.Fieldname)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify fieldname', Comment = 'DAN="Angiver feltnavn"';
                }
                Field("Default Value"; Rec."Default Value")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the default value for the Field', Comment = 'DAN="Specificere standardværdi for feltet"';
                }
                Field(ReverseSign; Rec.ReverseSign)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the if value for the Field need to reverse sign', Comment = 'DAN="Specificere værdi i Excel skal have vendt fortegn"';
                }
                Field(MandatoryValue; Rec.Value_Mandatory)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify that the column must have a value.', Comment = 'DAN="Specificere kolonne skal have en værdi"';
                }
                // field(ValueNotInUse; Rec.ValueNotInUse)
                // {
                //     ApplicationArea = All;
                //     ToolTip = 'Specify that the column is not in use.', Comment = 'DAN="Specificere kolonne er ikke i brug"';
                // }
                Field("Created date"; Rec."Created date")
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify the date of creation', Comment = 'DAN="Agiver dato for oprettelsen"';
                }
                Field("Created by"; Rec."Created by")
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify the User who created data', Comment = 'DAN="Angiver hvem der har oprettet data"';
                }
                Field("Last modified date"; Rec."Last modified date")
                {
                    Visible = false;
                    ApplicationArea = All;
                    ToolTip = 'Specify the date of latest modfication', Comment = 'DAN="Angiver hvilken dato, data sidst er rettet"';
                }
                Field("Last modified by"; Rec."Last modified by")
                {
                    Visible = false;
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

