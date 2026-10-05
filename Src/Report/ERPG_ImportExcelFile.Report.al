report 65000 "ERPG Import ExcelFile"
{
    ProcessingOnly = true;

    dataset
    {
    }

    requestpage
    {
        // SaveValues = true;

        layout
        {
            area(content)
            {
                group(Indstillinger)
                {
                    Caption = 'Options';
                    Field(ExcelFile; ExcelFile_Req)
                    {
                        Caption = 'ExcelFile', Comment = 'DAN="Excel filnavn"';
                        ApplicationArea = All;
                        ToolTip = 'Specify the Excel file name', Comment = 'DAN="Angiver Excel fil navn"';

                        trigger OnLookup(var Text: Text): Boolean
                        var
                        //Sheetname: Text[250];
                        begin

                            //FileStream: InStream; var TempNameValueBufferOut: Record "Name/Value Buffer" temporary) SheetsFound: Boolean
                            UploadIntoStream(ERP001Lbl, '', 'All Files (*.*)|*.xlsx;*.xls', ExcelFile_Req, InstreamVar);
                            if TempExcelBuffer.GetSheetsNameListFromStream(InstreamVar, TempNameValueBufferOut) then
                                if TempNameValueBufferOut.FindFirst() then
                                    SheetName_Req := TempNameValueBufferOut.Value;
                        end;

                    }
                    Field(SheetName; SheetName_Req)
                    {
                        Caption = 'SheetName', Comment = 'DAN="Ark navn"';
                        ApplicationArea = All;
                        ToolTip = 'Specify the Excel Sheet name', Comment = 'DAN="Angiver Excel ark navn"';


                        trigger OnLookup(var Text: Text): Boolean
                        begin
                            //SheetName := ExcelBuf.SelectSheetsName(ExcelFile);
                            if TempNameValueBufferOut.FindFirst() then
                                if Page.RunModal(Page::"Name/Value Lookup", TempNameValueBufferOut) = action::LookupOK then
                                    SheetName_Req := TempNameValueBufferOut.Value;
                        end;
                    }
                    Field("Excel Import Setupcode"; ExcelImportCode)
                    {
                        Caption = 'Excel import code', Comment = 'DAN="Excel import kode"';
                        ApplicationArea = All;
                        ToolTip = 'Specify the Excel Setup kode', Comment = 'DAN="Angiver Excel skabelon kode"';
                    }
                }
            }
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPostReport()
    begin

        WindowDialog.Open('Wait... data import in progess');

        TempExcelBuffer.SetFriendlyFilename(ExcelFile_Req);
        TempExcelBuffer.OpenBookStream(InstreamVar, SheetName_Req);
        //TempExcelBuffer.OpenBookStream()
        //TempExcelBuffer.OpenExcel();
        TempExcelBuffer.ReadSheet();
        SkipRowUpdate();
        if tempExcelBuffer.FindFirst() then
            repeat
                TempExcelBuffer3 := TempExcelBuffer;
                TempExcelBuffer3.Insert()
            until TempExcelBuffer.Next() = 0;
        if TempExcelBuffer.FindLast() then
            if ERP_ExcelImportSetup."Columns header in file" then begin
                for i := 2 to TempExcelBuffer."Row No." do
                    if TempExcelBuffer.GET(i, 1) then
                        case ERP_ExcelImportSetup.TableNo of
                            81:
                                CreateGenJnl(i);
                            83:
                                CreateItemJnl(i);
                            210:
                                CreateJobJnl(i);
                        end;

            end else begin
                for i := 1 to TempExcelBuffer."Row No." do
                    if TempExcelBuffer.Get(i, 1) then
                        case ERP_ExcelImportSetup.TableNo of
                            81:
                                CreateGenJnl(i);
                            83:
                                CreateItemJnl(i);
                            210:
                                CreateJobJnl(i);
                        end;
            end;
        WindowDialog.Close();
        Message(ERP002Lbl, ImportedLines);
    end;

    trigger OnPreReport()
    var
        GenJournalLineLocal: Record "Gen. Journal Line";
        GenJournalBatch: Record "Gen. Journal Batch";
        JobJournalBatch: Record "Job Journal Batch";
        JobJournalLine: Record "Job Journal Line";
        ItemJournalBatch: Record "Item Journal Batch";
        ItemJournalLine: Record "Item Journal Line";
    //GUIALLOWED := false;
    begin
        LanguageID := GLOBALLANGUAGE;

        ERP_ExcelImportSetup.GET(ExcelImportCode);
        //ERP_ExcelImportGenJnlLine.TESTFIELD("Use Postingdate");
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer2.DeleteAll();
        TempExcelBuffer3.DeleteAll();

        Clear(GenJournalLineLocal);
        Clear(GenJournalBatch);
        clear(JobJournalBatch);
        clear(JobJournalLine);
        case ERP_ExcelImportSetup.TableNo of
            81:
                begin
                    GenJournalLineLocal.SetRange("Journal Template Name", ERP_ExcelImportSetup.JournalTemplate);
                    GenJournalLineLocal.SetRange("Journal Batch Name", ERP_ExcelImportSetup.JournalBatch);
                    if GenJournalLineLocal.FindLast() then
                        NextLine := GenJournalLineLocal."Line No." + 10000
                    else
                        NextLine := 100000;

                    GenJournalBatch.Get(ERP_ExcelImportSetup.JournalTemplate, ERP_ExcelImportSetup.JournalBatch);
                    if GenJournalBatch."No. Series" <> '' then
                        NextDocno := NoSeriesManagement.GetNextNo(GenJournalBatch."No. Series", Today, false);
                end;
            83:
                begin
                    ItemJournalLine.SetRange("Journal Template Name", ERP_ExcelImportSetup.JournalTemplate);
                    ItemJournalLine.SetRange("Journal Batch Name", ERP_ExcelImportSetup.JournalBatch);
                    if ItemJournalLine.FindLast() then
                        NextLine := ItemJournalLine."Line No." + 10000
                    else
                        NextLine := 10000;
                    ItemJournalBatch.Get(ERP_ExcelImportSetup.JournalTemplate, ERP_ExcelImportSetup.JournalBatch);
                    if JobJournalBatch."No. Series" <> '' then
                        NextDocno := NoSeriesManagement.GetNextNo(ItemJournalBatch."No. Series", Today, false);
                end;
            210:
                begin
                    JobJournalLine.SetRange("Journal Template Name", ERP_ExcelImportSetup.JournalTemplate);
                    JobJournalLine.SetRange("Journal Batch Name", ERP_ExcelImportSetup.JournalBatch);
                    if JobJournalLine.FindLast() then
                        NextLine := JobJournalLine."Line No." + 10000
                    else
                        NextLine := 10000;
                    JobJournalBatch.Get(ERP_ExcelImportSetup.JournalTemplate, ERP_ExcelImportSetup.JournalBatch);
                    if JobJournalBatch."No. Series" <> '' then
                        NextDocno := NoSeriesManagement.GetNextNo(JobJournalBatch."No. Series", Today, false);
                end;
        end;

    end;

    var
        ERP_ExcelImportSetup: Record "ERPG Excel Import Setup";
        ERP_ExcelImportLines: Record "ERPG ExcelImport_SetupLine";
        TempExcelBuffer: Record "Excel Buffer";  // temporary;
        TempExcelBuffer2: Record "Excel Buffer" temporary;
        TempExcelBuffer3: Record "Excel Buffer" temporary;
        TempNameValueBufferOut: Record "Name/Value Buffer" temporary;
        ConfigValidateManagement: Codeunit "Config. Validate Management";
        NoSeriesManagement: Codeunit "No. Series"; //"NoSeriesManagement";
        RecordRef_: RecordRef;
        FieldRef_: FieldRef;
        InstreamVar: InStream;
        WindowDialog: Dialog;
        ExcelImportCode: Code[10];
        ExcelFile_Req: Text;
        SheetName_Req: Text;
        i: Integer;
        ERP001Lbl: Label 'Import File', Comment = 'DAN="Importer fil"';
        ERP002Lbl: Label 'Import completed - Imported lines %1', Comment = 'DAN="Import gennemført - antal linjer %1"';
        LanguageID: Integer;
        ImportedLines: Integer;
        NextLine: Integer;
        NextDocno: Code[20];

    procedure CheckForMandatoryFields(iRow: Integer): Boolean
    var

        ReturnValue: Boolean;
    begin
        ReturnValue := true;
        TempExcelBuffer3.SetRange("Row No.", iRow);

        Clear(ERP_ExcelImportLines);
        ERP_ExcelImportLines.SetRange(ERP_ExcelImportLines.Code, ERP_ExcelImportSetup.Code);
        ERP_ExcelImportLines.SetRange(value_Mandatory, true);
        if ERP_ExcelImportLines.FindFirst() then
            repeat
                TempExcelBuffer3.SetRange("Column No.", ERP_ExcelImportLines.Column);
                if TempExcelBuffer3.FindFirst() then
                    repeat
                        if TempExcelBuffer3."Column No." = ERP_ExcelImportLines.Column then
                            if (TempExcelBuffer3."Cell Value as Text" = '') and ERP_ExcelImportLines.Value_Mandatory then
                                ReturnValue := false;
                    until TempExcelBuffer3.Next() = 0
                else
                    ReturnValue := false;
            until ERP_ExcelImportLines.Next() = 0;

        exit(ReturnValue);
    end;

    local procedure CreateGenJnl(iRow: Integer)
    var
        GenJournalLine: Record "Gen. Journal Line";
    begin

        TempExcelBuffer.SetRange(TempExcelBuffer."Row No.", iRow);
        if TempExcelBuffer.FindFirst() then begin
            //Has row mandatory values?         
            if not CheckForMandatoryFields(TempExcelBuffer."Row No.") then
                exit;
            Clear(GenJournalLine);
            GenJournalLine.Init();
            GenJournalLine.Validate("Journal Template Name", ERP_ExcelImportSetup.JournalTemplate);
            GenJournalLine.Validate("Journal Batch Name", ERP_ExcelImportSetup.JournalBatch);
            GenJournalLine.Validate("Line No.", NextLine);
            if ERP_ExcelImportSetup."Use Postingdate" <> 0D then
                GenJournalLine.Validate("Posting Date", ERP_ExcelImportSetup."Use Postingdate");
            GenJournalLine.Validate("Document No.", NextDocno);
            GenJournalLine.Insert(true);

            RecordRef_.GetTable(GenJournalLine);
            HandleFieldDefault();

            NextDocno := IncStr(NextDocno);
            NextLine += 10000;
            repeat
                HandleField(TempExcelBuffer."Column No.", TempExcelBuffer."Cell Value as Text");
            until TempExcelBuffer.Next() = 0;
            ImportedLines += 1;
        end;
    end;

    local procedure CreateJobJnl(iRow: Integer)
    var
        JobJournalLine: Record "Job Journal Line";
    begin

        TempExcelBuffer.SetRange(TempExcelBuffer."Row No.", iRow);
        if TempExcelBuffer.FindFirst() then begin
            //Has row mandatory values?
            if not CheckForMandatoryFields(TempExcelBuffer."Row No.") then
                exit;
            Clear(JobJournalLine);
            JobJournalLine.Init();
            JobJournalLine.Validate("Journal Template Name", ERP_ExcelImportSetup.JournalTemplate);
            JobJournalLine.Validate("Journal Batch Name", ERP_ExcelImportSetup.JournalBatch);
            JobJournalLine.Validate("Line No.", NextLine);
            if ERP_ExcelImportSetup."Use Postingdate" <> 0D then
                JobJournalLine.Validate("Posting Date", ERP_ExcelImportSetup."Use Postingdate");
            JobJournalLine.Validate("Document No.", NextDocno);
            JobJournalLine.Insert(true);

            RecordRef_.GetTable(JobJournalLine);
            HandleFieldDefault();

            NextDocno := IncStr(NextDocno);
            NextLine += 10000;
            repeat
                HandleField(TempExcelBuffer."Column No.", TempExcelBuffer."Cell Value as Text");
            until TempExcelBuffer.Next() = 0;
            ImportedLines += 1;
        end;
    end;

    local procedure CreateItemJnl(iRow: Integer)
    var
        ItemJournalLine: Record "Item Journal Line";
    begin

        TempExcelBuffer.SetRange(TempExcelBuffer."Row No.", iRow);
        if TempExcelBuffer.FindFirst() then begin
            //Has row mandatory values?
            if not CheckForMandatoryFields(TempExcelBuffer."Row No.") then
                exit;
            Clear(ItemJournalLine);
            ItemJournalLine.Init();
            ItemJournalLine.Validate("Journal Template Name", ERP_ExcelImportSetup.JournalTemplate);
            ItemJournalLine.Validate("Journal Batch Name", ERP_ExcelImportSetup.JournalBatch);
            ItemJournalLine.Validate("Line No.", NextLine);
            if ERP_ExcelImportSetup."Use Postingdate" <> 0D then
                ItemJournalLine.Validate("Posting Date", ERP_ExcelImportSetup."Use Postingdate");
            ItemJournalLine.Validate("Document No.", NextDocno);
            ItemJournalLine.Insert(true);

            RecordRef_.GetTable(ItemJournalLine);
            HandleFieldDefault();

            NextDocno := IncStr(NextDocno);
            NextLine += 10000;
            repeat
                HandleField(TempExcelBuffer."Column No.", TempExcelBuffer."Cell Value as Text");
            until TempExcelBuffer.Next() = 0;
            ImportedLines += 1;
        end;
    end;

    local procedure HandleField(parColumn: Integer; parValue: Text)
    var
        TEMPDecimal: Decimal;
    begin
        Clear(ERP_ExcelImportLines);
        ERP_ExcelImportLines.SetRange(ERP_ExcelImportLines.Code, ERP_ExcelImportSetup.Code);
        ERP_ExcelImportLines.SetRange(Column, parColumn);
        if ERP_ExcelImportLines.FindFirst() then begin
            if (parValue = '') and ERP_ExcelImportLines.Value_Mandatory then
                exit;
            FieldRef_ := RecordRef_.Field(ERP_ExcelImportLines.FieldNo);
            if ERP_ExcelImportLines.ReverseSign then begin
                Clear(TEMPDecimal);
                Evaluate(TEMPDecimal, parValue);
                TEMPDecimal := TEMPDecimal * -1;
                parValue := format(TEMPDecimal);
            end;
            if delchr(parValue, '=', ' ') <> '' then
                ConfigValidateManagement.ValidateFieldValue(RecordRef_, FieldRef_, parValue, false, LanguageID);
            RecordRef_.Modify();
        end;
    end;

    local procedure HandleFieldDefault()
    begin
        Clear(ERP_ExcelImportLines);
        ERP_ExcelImportLines.SetRange(ERP_ExcelImportLines.Code, ERP_ExcelImportSetup.Code);
        ERP_ExcelImportLines.SetRange(Column, 0);
        if ERP_ExcelImportLines.FindFirst() then
            repeat
                FieldRef_ := RecordRef_.Field(ERP_ExcelImportLines.FieldNo);
                if (ERP_ExcelImportLines."Default Value" <> '') then
                    ConfigValidateManagement.ValidateFieldValue(RecordRef_, FieldRef_, ERP_ExcelImportLines."Default Value", false, LanguageID);
                RecordRef_.Modify();
            until ERP_ExcelImportLines.Next() = 0;
    end;

    procedure SetExcelSetup(parExcelCode: Code[10])
    begin
        ExcelImportCode := parExcelCode;
    end;

    procedure SkipRowUpdate()
    var
        ExcelSkipRow: Record "ERPG_ExcelImport_SkipRow";
    begin
        ExcelSkipRow.SetRange(Code, ERP_ExcelImportSetup.Code);
        if ExcelSkipRow.FindSet() then
            repeat
                TempExcelBuffer.SetRange(TempExcelBuffer."Row No.", ExcelSkipRow."Skip Row No");
                if TempExcelBuffer.FindSet() then
                    TempExcelBuffer.DeleteAll();
            until ExcelSkipRow.Next() = 0;
        TempExcelBuffer.Reset();
    end;
}

