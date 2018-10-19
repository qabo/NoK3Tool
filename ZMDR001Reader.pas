unit ZMDR001Reader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TZMDR001Record = packed record
    sNumber: string;
    sName: string;
    sCategory: string; //'物料类型描述';
    sGroupName: string; //'物料组描述';
    sRom: string; //'ROM';
    sRam: string; //'RAM';
    scap: string;
    sColor: string; //'颜色';
    sVer: string; //'制式';
    sBrand: string; //'品牌';
    sProj: string;
    isFG: Boolean;
  end;
  PZMDR001Record = ^TZMDR001Record;
   
  TSAPMaterialReader = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FProjs: TStringList;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PZMDR001Record;
    function GetProjCount: Integer;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear; 
    property Count: Integer read GetCount;
    property Items[i: Integer]: PZMDR001Record read GetItems;
    function GetSAPMaterialRecord(const snumber: string): PZMDR001Record;
    property ProjCount: Integer read GetProjCount;
    function ProjNo2Name(const sProjNo: string): string;
    class function IsHW(p: PZMDR001Record): Boolean;
  end;

implementation
       
{ TSAPMaterialReader }

constructor TSAPMaterialReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  FProjs := TStringList.Create;
  Open;
end;

destructor TSAPMaterialReader.Destroy;
begin
  Clear;
  FList.Free;
  FProjs.Free;
  inherited;
end;

procedure TSAPMaterialReader.Clear;
var
  i: Integer;
  p: PZMDR001Record;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PZMDR001Record(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
  FProjs.Clear;
end;
 
function TSAPMaterialReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPMaterialReader.GetItems(i: Integer): PZMDR001Record;
begin
  Result := PZMDR001Record(FList.Objects[i]);
end;

function TSAPMaterialReader.GetProjCount: Integer;
begin
  Result := FProjs.Count;
end;

function TSAPMaterialReader.ProjNo2Name(const sProjNo: string): string;
begin
  Result := FProjs.Values[sProjNo];
end;

class function TSAPMaterialReader.IsHW(p: PZMDR001Record): Boolean;
begin
  Result := p^.sVer = '海外';
end;  

function TSAPMaterialReader.GetSAPMaterialRecord(const snumber: string): PZMDR001Record;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PZMDR001Record(FList.Objects[idx]);
  end
  else Result := nil;
end;

procedure TSAPMaterialReader.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function IndexOfCol(ExcelApp: Variant; irow: Integer; const scol: string): Integer;
var
  i: Integer;
  s: string;
begin
  Result := -1;
  for i := 1 to 50 do
  begin
    s := ExcelApp.Cells[irow, i].Value;
    if s = scol then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TSAPMaterialReader.Open;
const
  CSNumber = '物料编码';
  CSName = '物料描述';
  CSCategory = '物料类型描述';
  CSGroupName = '物料组描述';
  CSRom = 'ROM';
  CSRam = 'RAM';
  CSColor = '颜色';
  CSVer = '制式';
  CSBrand = '品牌';
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPMaterialRecordPtr: PZMDR001Record;
  iColNumber: Integer;
  iColName: Integer;
  iCategory: Integer; //'物料类型描述';
  iGroupName: Integer; //'物料组描述';
  iRom: Integer; //'ROM';
  iRam: Integer; //'RAM';
  iColor: Integer; //'颜色';
  iVer: Integer; //'制式';
  iBrand: Integer; //'品牌';
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
        
        iColNumber := IndexOfCol(ExcelApp, irow, CSNumber);
        iColName := IndexOfCol(ExcelApp, irow, CSName);

        iCategory := IndexOfCol(ExcelApp, irow, CSCategory);
        iGroupName := IndexOfCol(ExcelApp, irow, CSGroupName);
        iRom := IndexOfCol(ExcelApp, irow, CSRom);
        iRam := IndexOfCol(ExcelApp, irow, CSRam);
        iColor := IndexOfCol(ExcelApp, irow, CSColor);
        iVer := IndexOfCol(ExcelApp, irow, CSVer);
        iBrand := IndexOfCol(ExcelApp, irow, CSBrand);
   
        if (iColNumber = -1) or (iColName = -1) or (iCategory = -1)
          or (iGroupName = -1) or (iRom = -1) or (iRam = -1)
          or (iColor = -1) or (iVer = -1) or (iBrand = -1)
          then
        begin
          Log(sSheet +'  不是  SAP导出物料  格式');
          Continue;
        end;
                
        irow := 2;
        snumber := ExcelApp.Cells[irow, iColNumber].Value;
        while snumber <> '' do
        begin                                
          aSAPMaterialRecordPtr := New(PZMDR001Record);
          FList.AddObject(snumber, TObject(aSAPMaterialRecordPtr));

          aSAPMaterialRecordPtr^.sNumber := snumber;
          aSAPMaterialRecordPtr^.sName := ExcelApp.Cells[irow, iColName].Value;
          aSAPMaterialRecordPtr^.sCategory := ExcelApp.Cells[irow, iCategory].Value;
          aSAPMaterialRecordPtr^.sGroupName := ExcelApp.Cells[irow, iGroupName].Value;
          aSAPMaterialRecordPtr^.sRom := ExcelApp.Cells[irow, iRom].Value;
          aSAPMaterialRecordPtr^.sRam := ExcelApp.Cells[irow, iRam].Value;
          aSAPMaterialRecordPtr^.sColor := ExcelApp.Cells[irow, iColor].Value;
          aSAPMaterialRecordPtr^.sVer := ExcelApp.Cells[irow, iVer].Value;
          aSAPMaterialRecordPtr^.sBrand := ExcelApp.Cells[irow, iBrand].Value;

          aSAPMaterialRecordPtr^.sCategory := StringReplace(aSAPMaterialRecordPtr^.sCategory, '（', '(', [rfReplaceAll]);
          aSAPMaterialRecordPtr^.sCategory := StringReplace(aSAPMaterialRecordPtr^.sCategory, '）', ')', [rfReplaceAll]);
          aSAPMaterialRecordPtr^.sProj := Copy(aSAPMaterialRecordPtr^.sGroupName, 1, Pos('(', aSAPMaterialRecordPtr^.sGroupName) - 1);

          aSAPMaterialRecordPtr^.isFG := Pos('量产整机', aSAPMaterialRecordPtr^.sCategory) > 0;

          aSAPMaterialRecordPtr^.scap := aSAPMaterialRecordPtr^.sRom;
          if aSAPMaterialRecordPtr^.sRam <> '' then
          begin
            aSAPMaterialRecordPtr^.scap := aSAPMaterialRecordPtr^.sRam + '+' + aSAPMaterialRecordPtr^.sRom;
          end;


          FProjs.Values[Copy(snumber, 1, 5)] := aSAPMaterialRecordPtr^.sProj;

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, iColNumber].Value;
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
