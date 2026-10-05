table 65000 "ERPG Excel Import Setup"
{
    // 001.  DTx.xx  10-02-2022  HRE  []
    //                                Obj. created
    LookupPageId = "ERPG ExcelImp. List, Setup";
    Caption = 'Excel Import, Setup', Comment = 'DAN="Excel Import, finanskladde"';

    fields
    {
        Field(1; "Code"; Code[10])
        {
            Caption = 'Code', Comment = 'DAN="Kode"';
            DataClassification = ToBeClassified;
        }
        Field(5; Description; Text[30])
        {
            Caption = 'Description', Comment = 'DAN="Beskrivelse"';
            DataClassification = ToBeClassified;
        }
        Field(20; "No. of Columns"; Integer)
        {
            Caption = 'No. of Columns', Comment = 'DAN="Antal kolonner"';
            DataClassification = ToBeClassified;
        }
        Field(23; "Columns header in file"; Boolean)
        {
            Caption = 'Columns header in file', Comment = 'DAN="Kolonneoverskifter i fil"';
            DataClassification = ToBeClassified;

        }
        Field(25; TableNo; Integer)
        {
            Caption = 'TabelNo', Comment = 'DAN="Tabel nr."';
            DataClassification = ToBeClassified;
            TableRelation = AllObj."Object ID" where("Object Type" = filter(Table));

            trigger OnValidate()
            begin
                CalcFields(Tablename);
            end;
        }
        // Field(30; "Gen. Journal Template"; Code[10])
        // {
        //     Caption = 'Gen. Journal Template', Comment = 'DAN="Kladdetype"';
        //     DataClassification = ToBeClassified;
        //     TableRelation = "Gen. Journal Template";
        // }
        // Field(35; "Gen. Journal Batch"; Code[10])
        // {
        //     Caption = 'Gen. Journal Batch', Comment = 'DAN="Kladdenavn"';
        //     DataClassification = ToBeClassified;
        //     TableRelation = "Gen. Journal Batch";

        //     trigger OnLookup()
        //     var
        //         LocalBatch: Record "Gen. Journal Batch";
        //     begin
        //         Clear(LocalBatch);
        //         LocalBatch.SetRange("Journal Template Name", "Gen. Journal Template");
        //         if LocalBatch.FindFirst() then
        //             if Page.RunModal(Page::"General Journal Batches", LocalBatch) = Action::LookupOK then
        //                 "Gen. Journal Batch" := LocalBatch.Name;
        //     end;
        // }
        // field(36; JobJournalTemplate; Code[10])
        // {
        //     Caption = 'Job Journal Template', Comment = 'DAN="Kladdenavn"';
        //     DataClassification = ToBeClassified;
        //     TableRelation = "Job Journal Template";
        // }
        // field(37; JobJournalBacth; Code[10])
        // {
        //     Caption = 'Job Journal Batch', Comment = 'DAN="Kladdenavn"';
        //     DataClassification = ToBeClassified;
        //     TableRelation = "Job Journal Batch";
        //     trigger OnLookup()
        //     var
        //         LocalBatch: Record "Job Journal Batch";
        //     begin
        //         Clear(LocalBatch);
        //         LocalBatch.SetRange("Journal Template Name", JobJournalTemplate);
        //         if LocalBatch.FindFirst() then
        //             if LocalBatch.FindFirst() then
        //                 if Page.RunModal(Page::"Job Journal Batches", LocalBatch) = Action::LookupOK then
        //                     JobJournalBacth := LocalBatch.Name;
        //     end;
        // }
        Field(40; "Use Postingdate"; Date)
        {
            Caption = 'Use Postingdate', Comment = 'DAN="Bogføringsdato i kladde"';
            DataClassification = ToBeClassified;
        }
        Field(50; Tablename; Text[249])
        {
            CalcFormula = Lookup(AllObjWithCaption."Object Caption" where("Object Type" = const(Table), "Object ID" = Field(TableNo)));
            Caption = 'Tablename', Comment = 'DAN="Tabelnavn"';
            FieldClass = FlowField;
        }
        field(55; ImportToTable; Enum ERPG_ImportToTable)
        {
            trigger OnValidate()
            begin
                case ImportToTable of
                    ImportToTable::Generel:
                        TableNo := 81;
                    ImportToTable::Item:
                        TableNo := 83;
                    ImportToTable::Job:
                        TableNo := 210;
                end;
            end;
        }
        field(60; JournalTemplate; Code[10])
        {
            Caption = 'Journal Template';
            TableRelation = if (ImportToTable = const(Generel)) "Gen. Journal Template"
            else if (ImportToTable = const(Job)) "Job Journal Template"
            else if (ImportToTable = const(Item)) "Item Journal Template";
        }
        field(62; JournalBatch; Code[10])
        {
            Caption = 'Journal Batch';
            TableRelation = if (ImportToTable = const(Generel)) "Gen. Journal Batch".Name where("Journal Template Name" = field(JournalTemplate))
            else if (ImportToTable = const(Job)) "Job Journal Batch".Name where("Journal Template Name" = field(JournalTemplate))
            else if (ImportToTable = const(Item)) "Item Journal Batch".Name where("Journal Template Name" = field(JournalTemplate));
        }
        Field(1000; "Created date"; Date)
        {
            Caption = 'Created date', Comment = 'DAN="Oprettet den"';
            DataClassification = ToBeClassified;
        }
        Field(1005; "Created by"; Text[80])
        {
            Caption = 'Created by', Comment = 'DAN="Oprettet af"';
            DataClassification = ToBeClassified;
        }
        Field(1010; "Last modified date"; Date)
        {
            Caption = 'Last modified date', Comment = 'DAN="Rettet den"';
            DataClassification = ToBeClassified;
        }
        Field(1015; "Last modified by"; Text[80])
        {
            Caption = 'Last modified by', Comment = 'DAN="Rettet af"';
            DataClassification = ToBeClassified;
        }
    }

    keys
    {
        key(Key1; "Code")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }

    trigger OnInsert()
    begin
        "Created date" := Today;
        "Created by" := UserId;

        //TableNo := 81;
    end;

    trigger OnModify()
    begin
        "Last modified date" := Today;
        "Last modified by" := UserId;
    end;

    trigger OnDelete()
    var
        Lines: Record "ERPG ExcelImport_SetupLine";
    begin
        Clear(Lines);
        Lines.Setrange(Code, Rec.Code);
        Lines.DeleteAll(true);
    end;


    procedure CreateLines_ColumnQty()
    var
        ERP_ExcelimpLinesGenJnlLine: Record "ERPG ExcelImport_SetupLine";
        i: Integer;
    begin
        Clear(ERP_ExcelimpLinesGenJnlLine);

        for i := 1 to "No. of Columns" do begin
            ERP_ExcelimpLinesGenJnlLine.Init();
            ERP_ExcelimpLinesGenJnlLine.Code := Code;
            ERP_ExcelimpLinesGenJnlLine.LineNo := i;
            ERP_ExcelimpLinesGenJnlLine.Column := i;
            ERP_ExcelimpLinesGenJnlLine.TabelNo := TableNo;
            ERP_ExcelimpLinesGenJnlLine.Insert(true);

        end;
    end;
}

