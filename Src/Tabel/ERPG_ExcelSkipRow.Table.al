table 65002 ERPG_ExcelImport_SkipRow
{
    Caption = 'Excel_SkipRow';
    DataClassification = ToBeClassified;

    fields
    {
        field(1; "Code"; Code[10])
        {
            Caption = 'Code', Comment = 'DAN="Kode"';
            DataClassification = ToBeClassified;
            TableRelation = "ERPG Excel Import Setup";
        }
        field(5; "Skip Row No"; Integer)
        {
            Caption = 'Skip Row No';
        }
    }
    keys
    {
        key(PK; "Code", "Skip Row No")
        {
            Clustered = true;
        }
    }
}
