pageextension 50101 ZYItemListExt extends "Item List"
{
    actions
    {
        addfirst(processing)
        {
            action(ImportLinksFMExcel)
            {
                Caption = 'Import Links from Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Import;
                trigger OnAction()
                begin
                    ReadExcelSheet();
                    ImportLinksFromExcel();
                end;
            }
            action(ImportNotesFMExcel)
            {
                Caption = 'Import Notes From Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Import;
                trigger OnAction()
                begin
                    ReadExcelSheet();
                    ImportNotesFromExcel();
                end;
            }
            action(ClearAllLinksAndNotes)
            {
                Caption = 'Clear All Links and Notes';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Import;
                trigger OnAction()
                begin
                    ClearAllLinksAndNotes(Rec);
                end;
            }
        }
    }
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        UploadExcelMsg: Label 'Please Choose the Excel file.';
        NoFileFoundMsg: Label 'No Excel file found!';
        ExcelImportSucess: Label 'Excel is successfully imported.';
        LastLinkID: Integer;

    local procedure ImportLinksFromExcel()
    var
        RowNo: Integer;
        ColNo: Integer;
        LineNo: Integer;
        MaxRowNo: Integer;
        RecordLink: Record "Record Link";
        Item: Record Item;
    begin
        RowNo := 0;
        ColNo := 0;
        MaxRowNo := 0;
        LineNo := 0;
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then
            MaxRowNo := TempExcelBuffer."Row No.";
        LastLinkID := 0;
        GetLastLinkID();
        for RowNo := 2 to MaxRowNo do begin
            LastLinkID += 1;
            RecordLink.Init();
            RecordLink."Link ID" := LastLinkID;
            RecordLink.Company := CompanyName;
            RecordLink.Type := RecordLink.Type::Link;
            RecordLink.Created := CurrentDateTime;
            RecordLink."User ID" := UserId;
            Item.Get(GetValueAtCell(RowNo, 1));
            RecordLink."Record ID" := Item.RecordId;
            Evaluate(RecordLink.URL1, GetValueAtCell(RowNo, 3));
            Evaluate(RecordLink.Description, GetValueAtCell(RowNo, 4));
            RecordLink.Insert();
        end;
        Message(ExcelImportSucess);
    end;

    local procedure ImportNotesFromExcel()
    var
        RowNo: Integer;
        ColNo: Integer;
        LineNo: Integer;
        MaxRowNo: Integer;
        RecordLink: Record "Record Link";
        Item: Record Item;
        RecordLinkMgt: Codeunit "Record Link Management";
    begin
        RowNo := 0;
        ColNo := 0;
        MaxRowNo := 0;
        LineNo := 0;
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then
            MaxRowNo := TempExcelBuffer."Row No.";
        LastLinkID := 0;
        GetLastLinkID();
        for RowNo := 2 to MaxRowNo do begin
            LastLinkID += 1;
            RecordLink.Init();
            RecordLink."Link ID" := LastLinkID;
            RecordLink.Insert();
            RecordLink.Company := CompanyName;
            RecordLink.Type := RecordLink.Type::Note;
            RecordLink.Created := CurrentDateTime;
            RecordLink."User ID" := UserId;
            Item.Get(GetValueAtCell(RowNo, 1));
            RecordLink."Record ID" := Item.RecordId;
            RecordLinkMgt.WriteNote(RecordLink, GetValueAtCell(RowNo, 3));
            RecordLink.Modify();

        end;
        Message(ExcelImportSucess);
    end;

    local procedure ClearAllLinksAndNotes(Item: Record Item)
    var
        RecordLink: Record "Record Link";
    begin
        Item.Reset();
        if Item.FindSet() then
            repeat
                RecordLink.Reset();
                RecordLink.SetRange("Record ID", Item.RecordId);
                if RecordLink.FindSet() then
                    RecordLink.DeleteAll();
            until Item.Next() = 0;
    end;

    local procedure GetLastLinkID()
    var
        RecordLink: Record "Record Link";
    begin
        RecordLink.Reset();
        if RecordLink.FindLast() then
            LastLinkID := RecordLink."Link ID"
        else
            LastLinkID := 0;
    end;

    local procedure ReadExcelSheet()
    var
        FileMgt: Codeunit "File Management";
        IStream: InStream;
        FromFile: Text[100];
    begin
        UploadIntoStream(UploadExcelMsg, '', '', FromFile, IStream);
        if FromFile = '' then
            Error(NoFileFoundMsg);
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(IStream, TempExcelBuffer.SelectSheetsNameStream(IStream));
        TempExcelBuffer.ReadSheet();
    end;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin

        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text")
        else
            exit('');
    end;
}
