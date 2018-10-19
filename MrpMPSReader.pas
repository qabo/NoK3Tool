unit MrpMPSReader;

interface

uses
  Classes, SysUtils, ComObj;

type
  TMrpMPSLine = class
    sbillno: string; //单据编号
    needdate: TDateTime;//	预测开始日期
    date: TDateTime;//	日期
    qty_net: Double;//	净需求数量
    snumber: string; //	产品编码
    sname: string; //	产品名称
  end;
  
  TMrpMPSReader = class      
  private
    FFile: string;    
    FList: TStringList;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    function GetItems(i: Integer): TMrpMPSLine;
    function GetCount: Integer;
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    property Items[i: Integer]: TMrpMPSLine read GetItems;
    property Count: Integer read GetCount;
    function GetQty(const sNumber: string; dt1, dt2: TDateTime): Double;
  end;

implementation

{ TMrpMPSReader }
             
constructor TMrpMPSReader.Create(const sfile: string);
begin
  FList := TStringList.Create;
  FFile := sfile;
  Open;
end;

destructor TMrpMPSReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TMrpMPSReader.Log(const str: string);
begin

end;

function TMrpMPSReader.GetItems(i: Integer): TMrpMPSLine;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := TMrpMPSLine(FList[i]);
  end
  else Result := nil;
end;

function TMrpMPSReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

procedure TMrpMPSReader.Clear;
var
  i: Integer;
  aMrpMPSLine: TMrpMPSLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aMrpMPSLine := TMrpMPSLine(FList.Objects[i]);
    aMrpMPSLine.Free;
  end;
  FList.Clear;
end;

procedure TMrpMPSReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer; 
  snumber: string;     
  aMrpMPSLine: TMrpMPSLine;
begin
  Clear;
  if not FileExists(FFile) then Exit;

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
        Log(sSheet);

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;   
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;
        if stitle <> '单据编号预测开始日期日期净需求数量产品编码产品名称' then
        begin
          Log(sSheet +'  不是简易MRP MPS格式');
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 5].Value;
        while snumber <> '' do
        begin
          aMrpMPSLine := TMrpMPSLine.Create;
          FList.AddObject(snumber, aMrpMPSLine);

          aMrpMPSLine.sbillno := ExcelApp.Cells[irow, 1].Value;
          aMrpMPSLine.needdate := ExcelApp.Cells[irow, 2].Value;
          aMrpMPSLine.date := ExcelApp.Cells[irow, 3].Value;
          aMrpMPSLine.qty_net := ExcelApp.Cells[irow, 4].Value;
          aMrpMPSLine.snumber := snumber;
          aMrpMPSLine.sname := ExcelApp.Cells[irow, 6].Value;

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 5].Value;
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

function TMrpMPSReader.GetQty(const sNumber: string; dt1, dt2: TDateTime): Double;
var
  i: Integer;
  aMrpMPSLine: TMrpMPSLine;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aMrpMPSLine := TMrpMPSLine(FList.Objects[i]);
    if (aMrpMPSLine.snumber = sNumber)
      and (aMrpMPSLine.needdate >= dt1)
      and (aMrpMPSLine.needdate <= dt2) then
    begin
      Result := Result + aMrpMPSLine.qty_net;
    end;
  end;
end;

end.
