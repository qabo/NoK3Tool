unit FGRptWinReader;

interface

uses
  Classes, CommUtils, SysUtils, ComObj, MakeFGReportCommon;

type
  TFGRptWinRecord = packed record 
    dt: TDateTime; //�������
    sNumber: string; //���ϳ�����
    sName: string; //��������
    dQty: Double; //����
    sFac: string; //������
    sBatchNo: string; //����
    sNote: string; //��ע
    bSum: Boolean;
  end;
  PFGRptWinRecord = ^TFGRptWinRecord;
  
  TFGRptWinReader = class
  private
    FFile: string;
    FList: TStringList;               
    FLogEvent: TLogEvent;
    ExcelApp, WorkBook: Variant;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PFGRptWinRecord;
  public
    constructor Create(const sfile: string; const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
    procedure GetNumberSet(slNumber: TStringList);
    property Count: Integer read GetCount;
    property Items[i: Integer]: PFGRptWinRecord read GetItems;
    property ReadOk: Boolean read FReadOk;
  end;

implementation

{ TFGRptWinReader}

constructor TFGRptWinReader.Create(const sfile: string; const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;   
  FLogEvent := LogEvent;
  Open;
end;

destructor TFGRptWinReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TFGRptWinReader.Clear;
var
  i: Integer;
  p: PFGRptWinRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p :=  PFGRptWinRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

procedure TFGRptWinReader.GetNumberSet(slNumber: TStringList);
var
  i: Integer;
  p: PFGRptWinRecord;    
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

procedure TFGRptWinReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;

function TFGRptWinReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TFGRptWinReader.GetItems(i: Integer): PFGRptWinRecord;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PFGRptWinRecord(FList.Objects[i]);
  end
  else Result := nil;
end;

procedure TFGRptWinReader.Open;
const
  CIDate = 1; //�������
  CINumber = 2; //���ϳ�����
  CIName = 3; //��������
  CIQty = 4; //����
  CIFac = 5; //������
  CIBatchNo = 6; //����
  CINote = 7; //��ע
 
var
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;     
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7: string;
  stitle: string;
  irow: Integer;
  snumber, sname: string;
  p: PFGRptWinRecord;
begin
  log(FFile);
     
  FReadOk := False;
    
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
        stitle4 := ExcelApp.Cells[irow, 4]. Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 +
          stitle7;
        if stitle <> '��������ϺŲ�Ʒ�����������������α�ע' then
        begin

          Log(sSheet +'  ���� ��Ʒ���� ��� ��ʽ');
          Continue;
        end;
             
        FReadOk := True;
  
        irow := 2;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        sname := ExcelApp.Cells[irow, CIName].Value;
        while (snumber <> '') and (sname <> '') do
        begin
          p := New(PFGRptWinRecord);
          p^.bSum := False;
          
          p^.dt := ExcelApp.Cells[irow, CIDate].Value;
          p^.sNumber := snumber;
          p^.sName := sname;
          p^.dQty := ExcelApp.Cells[irow, CIQty].Value;
          p^.sFac := ExcelApp.Cells[irow, CIFac].Value;
          p^.sBatchNo := ExcelApp.Cells[irow, CIBatchNo].Value;
          p^.sNote := ExcelApp.Cells[irow, CINote].Value; 
          FList.AddObject(snumber, TObject(p));

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, CINumber].Value;
          sname := ExcelApp.Cells[irow, CIName].Value;
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
