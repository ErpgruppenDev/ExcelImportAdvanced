page 68760 "ERPG Excel import Setup"
{
    PageType = Card;
    SourceTable = "ERPG Excel Import Setup";
    Caption = 'Excel import, Setup', Comment = 'DAN="Excel import, opsætning"';
    UsageCategory = Tasks;

    layout
    {
        area(content)
        {
            group(General)
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
                Field("Columns header in file"; Rec."Columns header in file")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify if the file contains colunms headers', Comment = 'DAN="Angiver om der er kolonne overskrifter i filen"';
                }
                field(ImportToTable; Rec.ImportToTable) { ApplicationArea = All; Visible = true; }
                Field(TableNo; Rec.TableNo)
                {
                    ApplicationArea = All;
                    Visible = false;
                    ToolTip = 'Specify tableno. (81 for Gen.Journal, 83 for Item Journal and 210 for ProjectJournal)', Comment = 'DAN="Angiver tabelnr."';

                    // trigger OnValidate()
                    // begin
                    //     if Rec.TableNo <> 0 then
                    //         if (Rec.TableNo <> 81) and (Rec.TableNo <> 210) then
                    //             Error('Der kan kun angives tabel 81 (finanskladde) eller 210 (projektkladde)');
                    //     case Rec.TableNo of
                    //         0:
                    //             begin
                    //                 EditAbleGL := false;
                    //                 EditAbleJob := false;
                    //             end;
                    //         81:
                    //             begin
                    //                 EditAbleGL := true;
                    //                 EditAbleJob := false;
                    //             end;
                    //         210:
                    //             begin
                    //                 EditAbleGL := false;
                    //                 EditAbleJob := true;
                    //             end;
                    //     end;
                    //     CurrPage.Update(true);
                    // end;
                }
                Field(Tablename; Rec.Tablename)
                {
                    Visible = true;
                    ApplicationArea = All;
                    ToolTip = 'Specify tablename', Comment = 'DAN="Angiver tabelnavn"';
                }
                field(JournalTemplate; Rec.JournalTemplate) { ApplicationArea = All; Visible = true; }
                field(JournalBatch; Rec.JournalBatch) { ApplicationArea = All; Visible = true; }
                Field("Use Postingdate"; Rec."Use Postingdate")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify which dato to use in the journal, in import', Comment = 'DAN="Angiver bogføringsdato der benyttes i kladde, ved import"';
                }
                field(SameDocNoAllLines; Rec.SameDocNoAllLines) { ApplicationArea = All; ToolTip = 'Specify if all lines need to have same Document No.';}
            }
            part("Excel import lines Gen.Jnlline"; "ERPG Excelimp.lines Setupline")
            {
                Caption = 'Lines / Column Specify', Comment = 'DAN="Linjer / kolonnespecifikation"';
                ApplicationArea = all;
                SubPageLink = Code = Field(Code);
            }
            part("Excel import Skip Row"; "ERPG ExcelImport_SkipRow")
            {
                Caption = 'Skip Row', Comment = 'DAN="Spring over rækker"';
                ApplicationArea = all;
                SubPageLink = Code = Field(Code);
            }
            group(Information)
            {
                Field("Created date"; Rec."Created date")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the date of creation', Comment = 'DAN="Agiver dato for oprettelsen"';
                }
                Field("Created by"; Rec."Created by")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specify the User who created data', Comment = 'DAN="Angiver hvem der har oprettet data"';
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
        area(navigation)
        {
            group("&Funktioner")
            {
                Caption = 'Functions', Comment = 'DAN="Funktioner"';
                Image = ReferenceData;
                action("CreateLines for Column")
                {
                    Caption = 'CreateLines for Column', Comment = 'DAN="Opret linjer (antal kolonner)"';
                    ApplicationArea = All;
                    ToolTip = 'Creates lines, according to the quantity, fill in No. of Columns', Comment = 'DAN="Opretter samme antal linjer, som der er angivet i kolonne antal"';
                    Image = Column;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedIsBig = true;
                    ShortCutKey = 'F7';

                    trigger OnAction()
                    begin
                        Rec.CreateLines_ColumnQty();
                    end;
                }
                action("Import Excelfile")
                {
                    Caption = 'Import Excelfile', Comment = 'DAN="Importer Excel fil"';
                    ApplicationArea = All;
                    ToolTip = 'Importer Excelfile', Comment = 'DAN="Importer Excel fil"';
                    Image = Excel;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedIsBig = true;

                    trigger OnAction()
                    var
                        RptImport: Report "ERPG Import ExcelFile";
                    begin
                        RptImport.SetExcelSetup(Rec.Code);
                        RptImport.UseRequestPage(true);
                        RptImport.RunModal();
                    end;
                }

                action("Show Journal")
                {
                    Caption = 'Show Journal', Comment = 'DAN="Vis kladde"';
                    ApplicationArea = All;
                    ToolTip = 'Show the Journal where data is imported', Comment = 'DAN="Viser kladde er importeret til"';
                    Image = Journal;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedIsBig = true;

                    trigger OnAction()
                    var
                        GenJournal: Record "Gen. Journal Template";
                        GenLine: Record "Gen. Journal Line";
                        JobJournal: Record "Job Journal Template";
                        JobLine: record "Job Journal Line";
                        ItemJournal: Record "Item Journal Template";
                        ItemJournalLine: Record "Item Journal Line";
                        NoJournalLbl: Label 'No Lines in journal', Comment = 'DAN="Der findes ingen kladdelinjer"';
                    begin
                        case Rec.TableNo of
                            81:
                                begin
                                    Clear(GenJournal);
                                    GenJournal.SetRange(Name, Rec.JournalTemplate);
                                    if GenJournal.FindFirst() then begin
                                        Clear(GenLine);
                                        GenLine.SetRange("Journal Template Name", Rec.JournalTemplate);
                                        GenLine.SetRange("Journal Batch Name", Rec.JournalBatch);
                                        if GenLine.FindFirst() then
                                            Page.RunModal(GenJournal."Page ID", GenLine)
                                        else
                                            Message(NoJournalLbl);
                                    end;
                                end;
                            83:
                                begin
                                    Clear(ItemJournal);
                                    ItemJournal.SetRange(Name, Rec.JournalTemplate);
                                    if ItemJournal.FindFirst() then begin
                                        Clear(ItemJournalLine);
                                        ItemJournalLine.SetRange("Journal Template Name", Rec.JournalTemplate);
                                        ItemJournalLine.SetRange("Journal Batch Name", Rec.JournalBatch);
                                        if ItemJournalLine.FindFirst() then
                                            Page.RunModal(ItemJournal."Page ID", ItemJournalLine)
                                        else
                                            Message(NoJournalLbl);
                                    end
                                end;
                            210:
                                begin
                                    Clear(JobJournal);
                                    JobJournal.SetRange(Name, Rec.JournalTemplate);
                                    if JobJournal.FindFirst() then begin
                                        Clear(JobLine);
                                        JobLine.SetRange("Journal Template Name", Rec.JournalTemplate);
                                        JobLine.SetRange("Journal Batch Name", Rec.JournalBatch);
                                        if JobLine.FindFirst() then
                                            Page.RunModal(JobJournal."Page ID")
                                        else
                                            Message(NoJournalLbl);
                                    end;
                                end;
                        end;
                    end;
                }
            }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.CalcFields(Tablename);
        // case Rec.TableNo of
        //     81:
        //         begin
        //             EditAbleGL := true;
        //             EditAbleJob := false;
        //         end;
        //     210:
        //         begin
        //             EditAbleGL := false;
        //             EditAbleJob := true;
        //         end;
        // end;
    end;

    trigger OnAfterGetCurrRecord()
    begin
        // case Rec.TableNo of
        //     81:
        //         begin
        //             EditAbleGL := true;
        //             EditAbleJob := false;
        //         end;
        //     210:
        //         begin
        //             EditAbleGL := false;
        //             EditAbleJob := true;
        //         end;
        // end;
    end;



    var
    // EditAbleGL: Boolean;
    // EditAbleJob: Boolean;
}

