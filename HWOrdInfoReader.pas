unit HWOrdInfoReader;

interface

uses
  Classes, ComObj, DateUtils, SysUtils, CommUtils, Variants;

type
  THWOrdInfoRecord = packed record
    sproj: string;
    sordno: string;
    snumber: string;
    sname: string;
    scolor: string;
    scap: string;
    dqty: Double;
    soutdate: string;
    Tag: TObject;
  end;
  PHWOrdInfoRecord = ^THWOrdInfoRecord;

  THWOODRecord = packed record
    sproj: string;
    sordno: string;
    snumber: string;
    sname: string;
    scolor: string;
    scap: string;
    dqty: Double;  
    Tag: TObject;
  end;
  PHWOODRecord = ^THWOODRecord;
  
  THWOrdInfoReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FOrdInfoList: TList;
    FOODList: TList;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetOrdInfoCount: Integer;
    function GetOrdInfoItems(i: Integer): PHWOrdInfoRecord;
    function GetOODCount: Integer;                                          
    function GetOODItems(i: Integer): PHWOODRecord;
  public 
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;  
    property OrdInfoCount: Integer read GetOrdInfoCount;
    property OrdInfoItems[i: Integer]: PHWOrdInfoRecord read GetOrdInfoItems;     
    property OODCount: Integer read GetOODCount;
    property OODItems[i: Integer]: PHWOODRecord read GetOODItems;
  end;

implementation
 
{ THWOrdInfoReader }

constructor THWOrdInfoReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FLogEvent := aLogEvent; 
  FFile := sfile;
  FOrdInfoList := TList.Create;
  FOODList := TList.Create;
  Open;
end;

destructor THWOrdInfoReader.Destroy;
begin
  Clear;
  FOrdInfoList.Free;
  FOODList.Free;
  inherited;
end;

procedure THWOrdInfoReader.Clear;
var
  i: Integer;
  p: PHWOrdInfoRecord;
  aHWOODRecordPtr: PHWOODRecord;
begin
  for i := 0 to FOrdInfoList.Count - 1 do
  begin
    p := PHWOrdInfoRecord(FOrdInfoList[i]);
    Dispose(p);
  end;
  FOrdInfoList.Clear;
  
  for i := 0 to FOrdInfoList.Count - 1 do
  begin
    aHWOODRecordPtr := PHWOODRecord(FOODList[i]);
    Dispose(aHWOODRecordPtr);
  end;
  FOrdInfoList.Clear;
end;
 
procedure THWOrdInfoReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function THWOrdInfoReader.GetOrdInfoCount: Integer;
begin
  Result := FOrdInfoList.Count;
end;

function THWOrdInfoReader.GetOrdInfoItems(i: Integer): PHWOrdInfoRecord;
begin
  Result := PHWOrdInfoRecord(FOrdInfoList[i]);
end;

function THWOrdInfoReader.GetOODCount: Integer;
begin
  Result := FOODList.Count;
end;

function THWOrdInfoReader.GetOODItems(i: Integer): PHWOODRecord;
begin
  Result := PHWOODRecord(FOODList[i]);
end;  

procedure THWOrdInfoReader.Open;
var
  irow: Integer;
  irow1: Integer;
  stitle1, stitle2, stitle3, stitle4,
  stitle5, stitle6: string;
  stitle: string;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  snumber: string;
  sproj: string;
  aHWOrdInfoRecordPtr: PHWOrdInfoRecord;
  aHWOODRecordPtr: PHWOODRecord;
  v: Variant;
  dt: TDateTime;
begin  
  Clear;
 
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;

        stitle := stitle1 + stitle2 + stitle3 + stitle4 +
          stitle5 + stitle6;
        if stitle <> '项目生产单单号料号描述颜色容量' then
        begin
          Log(sSheet + '格式不符');
          Continue;
        end;

        //   4-7订单完成状况 
        if ExcelApp.Cells[irow, 19].Value = '延迟天数' then
        begin
          irow := 2;
          irow1 := irow;
          snumber := ExcelApp.Cells[irow, 3].Value;     
          sproj := ExcelApp.Cells[irow, 1].Value;
          while snumber <> '' do
          begin
            aHWOrdInfoRecordPtr := New(PHWOrdInfoRecord);
            FOrdInfoList.Add(aHWOrdInfoRecordPtr);

            if not IsCellMerged(ExcelApp, irow1, 1, irow, 1) then
            begin
              sproj := ExcelApp.Cells[irow, 1].Value;  
              irow1 := irow;   
            end;
            aHWOrdInfoRecordPtr^.Tag := nil;
            aHWOrdInfoRecordPtr^.sproj := sproj;          
            aHWOrdInfoRecordPtr^.sordno := ExcelApp.Cells[irow, 2].Value;
            aHWOrdInfoRecordPtr^.snumber := ExcelApp.Cells[irow, 3].Value;
            aHWOrdInfoRecordPtr^.sname := ExcelApp.Cells[irow, 4].Value;
            aHWOrdInfoRecordPtr^.scolor := ExcelApp.Cells[irow, 5].Value;
            aHWOrdInfoRecordPtr^.scap := ExcelApp.Cells[irow, 6].Value;
            aHWOrdInfoRecordPtr^.dqty := ExcelApp.Cells[irow, 8].Value;
            v := ExcelApp.Cells[irow, 18].Value;
            if VarIsType(v, varDate) then
            begin
              dt := v;
              aHWOrdInfoRecordPtr^.soutdate := FormatDateTime('YYYY-MM-DD', dt);
            end
            else
            begin
              aHWOrdInfoRecordPtr^.soutdate := '';
            end;

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 3].Value;
          end;
        end
        else  //   Open 订单
        begin
          irow := 2;
          irow1 := irow;
          snumber := ExcelApp.Cells[irow, 3].Value;     
          sproj := ExcelApp.Cells[irow, 1].Value;
          while snumber <> '' do
          begin
            aHWOODRecordPtr := New(PHWOODRecord);
            FOODList.Add(aHWOODRecordPtr);

            if not IsCellMerged(ExcelApp, irow1, 1, irow, 1) then
            begin
              sproj := ExcelApp.Cells[irow, 1].Value;  
              irow1 := irow;   
            end;
            aHWOODRecordPtr^.Tag := nil;
            aHWOODRecordPtr^.sproj := sproj;
            aHWOODRecordPtr^.sordno := ExcelApp.Cells[irow, 2].Value;
            aHWOODRecordPtr^.snumber := ExcelApp.Cells[irow, 3].Value;
            aHWOODRecordPtr^.sname := ExcelApp.Cells[irow, 4].Value;
            aHWOODRecordPtr^.scolor := ExcelApp.Cells[irow, 5].Value;
            aHWOODRecordPtr^.scap := ExcelApp.Cells[irow, 6].Value;
            aHWOODRecordPtr^.dqty := ExcelApp.Cells[irow, 8].Value; 
            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 3].Value;
          end;
        end; 
      end;
        
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit;
  end;
end;

end.
