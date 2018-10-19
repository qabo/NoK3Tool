unit XHBomReader;

interface

uses
  Classes, SysUtils, ComObj, CommUtils, KeyICItemSupplyReader;

type 
  TXHBomRecord = packed record
    snumber: string;
    sname: string;
    smodel: string;
    serpclsid: string;
    sextra: string;
  end;
  PXHBomRecord = ^TXHBomRecord;
                 
  TXHBomChilcRecord = packed record
    snumber: string;
    sname: string;
    smodel: string;
    susage: string;
    slocation: string;
  end;
  PXHBomChilcRecord = ^TXHBomChilcRecord;

  TXHBomReader = class
  private          
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PXHBomChilcRecord;
  public
    FXHBomRecord: TXHBomRecord;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PXHBomChilcRecord read GetItems;
  end;

implementation
       
{ TXHBomReader }

constructor TXHBomReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open; 
end;

destructor TXHBomReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TXHBomReader.Clear;
var
  i: Integer;
  p: PXHBomRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PXHBomRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

procedure TXHBomReader.Log(const str: string);
begin

end;

function TXHBomReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TXHBomReader.GetItems(i: Integer): PXHBomChilcRecord;
begin
  Result := PXHBomChilcRecord(FList.Objects[i]);
end;

procedure TXHBomReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9, stitle10: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  ptrXHBomChilcRecord: PXHBomChilcRecord;
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
        Log(sSheet);

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
        if stitle <> '物料代码物料名称规格型号物料属性辅助属性' then
        begin
          Log(sSheet +'  不是BOM格式');
          Continue;
        end;

        irow := 2;
        FXHBomRecord.snumber := ExcelApp.Cells[irow, 1].Value;
        FXHBomRecord.sname := ExcelApp.Cells[irow, 2].Value;
        FXHBomRecord.smodel := ExcelApp.Cells[irow, 3].Value;
        FXHBomRecord.serpclsid := ExcelApp.Cells[irow, 4].Value;
        FXHBomRecord.sextra := ExcelApp.Cells[irow, 5].Value;

        irow := 4;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          ptrXHBomChilcRecord := New(PXHBomChilcRecord);

          ptrXHBomChilcRecord^.snumber := snumber;
          ptrXHBomChilcRecord^.sname := ExcelApp.Cells[irow, 2].Value;
          ptrXHBomChilcRecord^.smodel := ExcelApp.Cells[irow, 3].Value;
          ptrXHBomChilcRecord^.susage := ExcelApp.Cells[irow, 4].Value;
          ptrXHBomChilcRecord^.slocation := ExcelApp.Cells[irow, 5].Value;
          FList.AddObject(snumber, TObject(ptrXHBomChilcRecord));

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 1].Value;     
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
