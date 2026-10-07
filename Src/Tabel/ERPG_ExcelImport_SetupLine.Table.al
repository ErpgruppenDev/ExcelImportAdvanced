table 68761 "ERPG ExcelImport_SetupLine"
{
    Caption = 'Excel Import SetupLine', Comment = 'DAN="Excel Import linjer"';

    fields
    {
        Field(1; "Code"; Code[10])
        {
            Caption = 'Code', Comment = 'DAN="Kode"';
            DataClassification = ToBeClassified;
            TableRelation = "ERPG Excel Import Setup";
        }
        Field(5; LineNo; Integer)
        {
            AutoIncrement = true;
            Caption = 'LineNo', Comment = 'DAN="Linjenr."';
            DataClassification = ToBeClassified;
        }
        Field(10; Description; Text[30])
        {
            Caption = 'Description', Comment = 'DAN="Beskrivele"';
            DataClassification = ToBeClassified;
        }
        Field(15; Column; Integer)
        {
            Caption = 'Column', Comment = 'DAN="Kolonne"';
            DataClassification = ToBeClassified;
        }
        Field(20; TabelNo; Integer)
        {
            Caption = 'TabelNo', Comment = 'DAN="Tabelnr."';
            DataClassification = ToBeClassified;
            TableRelation = AllObj."Object ID" where("Object Type" = filter(Table));

            trigger OnValidate()
            begin
                CalcFields(Tablename);
            end;
        }
        Field(25; FieldNo; Integer)
        {
            Caption = 'FieldNo', Comment = 'DAN="Feltnr."';
            DataClassification = ToBeClassified;
            //TableRelation = Field."No." where(TableNo = Field(TabelNo));
            trigger OnLookup()
            var
                LocalFieldRec: Record Field;

            begin
                Clear(LocalFieldRec);
                LocalFieldRec.SetRange(TableNo, Rec.TabelNo);
                if LocalFieldRec.FindFirst() then
                    if Page.RunModal(Page::"ERPG Field list", LocalFieldRec) = Action::LookupOK then
                        FieldNo := LocalFieldRec."No.";
            end;

            trigger OnValidate()
            begin
                CalcFields(Fieldname);
            end;
        }
        Field(30; "Default Value"; Text[250])
        {
            Caption = 'Default Value', Comment = 'DAN="Standardværdi"';
            DataClassification = ToBeClassified;

            trigger OnValidate()
            var
                ConfigValidateMgt: Codeunit "Config. Validate Management";
                RecRef: RecordRef;
                FieldRef: FieldRef;
                ValidationError: Text[250];
            begin
                if (FieldNo <> 0) and ("Default Value" <> '') then begin
                    RecRef.Open(TabelNo, true);
                    FieldRef := RecRef.Field(FieldNo);
                    //valuateValueWithValidate(var FieldRef: FieldRef; Value: Text; XMLValue: Boolean): Text
                    ValidationError := ConfigValidateMgt.EvaluateValueWithValidate(FieldRef, "Default Value", false);
                    if ValidationError <> '' then
                        error(ValidationError);

                    "Default Value" := Format(FieldRef.Value);
                    Clear(Column);
                end
            end;
        }
        //Field(35; ValidateValue; Boolean)
        Field(35; ReverseSign; Boolean)
        {

            Caption = 'Reverse sign', Comment = 'DAN="Vend fortegn"';
            DataClassification = ToBeClassified;
            //Virker ikke som ønsket, så feltet skal væk igen
            // var brugt til om felt skulle valideres eller ej
            //ObsoleteState = Removed;
            //ObsoleteReason = 'Does not work - skip again';
        }
        field(40; Value_Mandatory; Boolean)
        {
            Caption = 'Value is Mandatory', Comment = 'DAN="Værdi er Obligatorisk"';
            DataClassification = ToBeClassified;
        }
        // field(45; ValueNotInUse; Boolean)
        // {
        //     Caption = 'Value Not In Use', Comment = 'DAN="Værdi ikke i brug"';
        //     DataClassification = ToBeClassified;
        // }
        Field(50; Tablename; Text[249])
        {
            CalcFormula = Lookup(AllObjWithCaption."Object Caption" where("Object Type" = const(Table), "Object ID" = Field(TabelNo)));
            Caption = 'Tablename', Comment = 'DAN="Tabelnavn"';
            FieldClass = FlowField;
        }
        Field(55; Fieldname; Text[80])
        {
            CalcFormula = Lookup(Field."Field Caption" where(TableNo = Field(TabelNo), "No." = Field(FieldNo)));
            Caption = 'Fieldname', Comment = 'DAN="Feltnavn"';
            Editable = false;
            FieldClass = FlowField;
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
        key(Key1; "Code", LineNo)
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }

    trigger OnInsert()
    begin
        ERP_ExcelImportGenJnlLine.Get(Rec.Code);
        Validate(TabelNo, ERP_ExcelImportGenJnlLine.TableNo);

    end;

    var
        ERP_ExcelImportGenJnlLine: Record "ERPG Excel Import Setup";
}

