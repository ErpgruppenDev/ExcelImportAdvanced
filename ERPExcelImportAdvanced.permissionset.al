/// <summary>
/// Unknown ERP_ExcelImportAdvanced (ID 60000).
/// </summary>
permissionset 65000 ERP_ExcelImpAdvanced
{
    Assignable = true;
    Permissions = tabledata "ERP Excel Import Setup"=RIMD,
        tabledata "ERP ExcelImport_SetupLine"=RIMD,
        table "ERP Excel Import Setup"=X,
        table "ERP ExcelImport_SetupLine"=X,
        report "ERP Import ExcelFile"=X,
        page "ERP Excel import Setup"=X,
        page "ERP ExcelImp. List, Setup"=X,
        page "ERP Excelimp.lines Setupline"=X,
        page "ERP Field list"=X;
}