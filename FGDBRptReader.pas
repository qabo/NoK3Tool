unit FGDBRptReader;

interface

uses
  Classes, CommUtils, SysUtils, ComObj, MakeFGReportCommon;

type
  TFGDBRptRecord = packed record
    snumber: string; //�Ϻ�
    sname: string; //��Ʒ����
    sBatchNo: string; // ����
    sstock: string; 
    dqty: Double;
  end;
  PFGDBRptRecord = ^TFGDBRptRecord;
  
  TFGDBRptReader = class
  private
    FFile: string;
    FList: TStringList;               
    FLogEvent: TLogEvent;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);       
    function GetCount: Integer;
    function GetItems(i: Integer): PFGDBRptRecord;
  public
    constructor Create(const sfile: string; const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;                  
    procedure GetNumberSet(slNumber: TStringList);    
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;    
    property Count: Integer read GetCount;
    property Items[i: Integer]: PFGDBRptRecord read GetItems;

  end;

implementation

{ TFGDBRptReader}

constructor TFGDBRptReader.Create(const sfile: string; const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;   
  FLogEvent := LogEvent;
  Open;
end;

destructor TFGDBRptReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TFGDBRptReader.Clear;
var
  i: Integer;
  p: PFGDBRptRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p :=  PFGDBRptRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
          
procedure TFGDBRptReader.GetNumberSet(slNumber: TStringList);
var
  i: Integer;
  p: PFGDBRptRecord;     
  pn: PFGNumberRecord;
begin
  for i := 0 to self.Count - 1 do
  begin
    p := Items[i];
    if slNumber.IndexOf(p^.snumber + '=' + p^.sBatchNo ) < 0 then
    begin        
      pn := New(PFGNumberRecord);
      pn^.snumber := p^.snumber;
      pn^.sname := p^.sname;
      pn^.sBatchNo := p^.sBatchNo;
      slNumber.AddObject(p^.snumber + '=' + p^.sBatchNo, TObject(pn) );
    end;
  end;
end;

procedure TFGDBRptReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;  
       
function TFGDBRptReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TFGDBRptReader.GetItems(i: Integer): PFGDBRptRecord;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PFGDBRptRecord(FList.Objects[i]);
  end
  else Result := nil;
end;

procedure TFGDBRptReader.Open;
const
  CINumber = 1; //���ϴ���
  CIName = 2;   //��������
  CIStock = 3;  //�ֿ�����
  CIQty = 4;    //���õ�λ����

var
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;     
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  p: PFGDBRptRecord;
begin
     
  Clear;

  if not FileExists(FFile) then Exit;

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
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
        stitle := stitle1 + stitle2 + stitle3 + stitle4;
        if stitle <> '���ϴ����������Ʋֿ����Ƴ��õ�λ����' then
        begin
          Log(sSheet +'  ���� ��Ʒ���� δ���� ��ʽ');
          Continue;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        while snumber <> '' do
        begin
          p := New(PFGDBRptRecord);
          
          p^.snumber := ExcelApp.Cells[irow, CINumber].Value;
          p^.sname := ExcelApp.Cells[irow, CIName].Value;
          p^.sstock := ExcelApp.Cells[irow, CIStock].Value;
          p^.dqty := ExcelApp.Cells[irow, CIQty].Value;

          FList.AddObject(snumber, TObject(p));
          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, CINumber].Value;
        end;

      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit;
  end;  
end;

end.
