/// <summary>
/// Unknown ERP_ExcelImportAdvanced (ID 60000).
/// </summary>
permissionset 65000 ERP_ExcelImpAdvanced
{
    Assignable = true;
    Permissions = tabledata "ERPG Excel Import Setup"=RIMD,
        tabledata "ERPG ExcelImport_SetupLine"=RIMD,
        table "ERPG Excel Import Setup"=X,
        table "ERPG ExcelImport_SetupLine"=X,
        report "ERPG Import ExcelFile"=X,
        page "ERPG Excel import Setup"=X,
        page "ERPG ExcelImp. List, Setup"=X,
        page "ERPG Excelimp.lines Setupline"=X,
        page "ERPG Field list"=X,
        tabledata ERPG_ExcelImport_SkipRow=RIMD,
        table ERPG_ExcelImport_SkipRow=X,
        page ERPG_ExcelImport_SkipRow=X;
}