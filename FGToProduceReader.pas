unit FGToProduceReader;

interface

uses
  Classes, CommUtils, SysUtils, ComObj, MakeFGReportCommon, FGRptWinReader,
  ZMDR001Reader;

type
  TCountrySet = (csC, csR, csU, csI, csO); // �й�������˹���ڿ�����ӡ�ᣬ����
  
  TFGToProduceRecord = packed record 
    sproj: string; //����
    sprojname: string; //��Ʒ����
    sitem: string; //��ɫ\��ʽ\����
//    dtotal_in: Double; //��Ŀ����         ����
//    dtotal_out: Double; //��Ŀ����        ����
//    dproduct_in: Double; //����������     ����
//    dproduct_out: Double; //����������    ����
    dtoproduce_in: Double; //��������		    ����        10
    dtoproduce_out: Double; //��������      С��        11
  end;
  PFGToProduceRecord = ^TFGToProduceRecord;
  
  TFGToProduceReader = class
  private
    FFile: string;
    FList: TStringList;               
    FLogEvent: TLogEvent;
    ExcelApp, WorkBook: Variant;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PFGToProduceRecord;
    procedure SubWinLine(aZMDR001RecordPtr: PZMDR001Record;
      aFGRptWinRecordPtr: PFGRptWinRecord);
  public
    constructor Create(const sfile: string; const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    procedure SubWin( aFGRptWinReader: TFGRptWinReader;
      aSAPMaterialReader: TSAPMaterialReader; dt1, dt2: TDateTime);
    function GetQty(const sProjName,
      sitem: string; slMap: TStrings; bHW: Boolean; acs: TCountrySet): Double;
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PFGToProduceRecord read GetItems;
    property ReadOk: Boolean read FReadOk;
  end;

implementation

{ TFGToProduceReader}

constructor TFGToProduceReader.Create(const sfile: string; const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;   
  FLogEvent := LogEvent;
  Open;
end;

destructor TFGToProduceReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TFGToProduceReader.Clear;
var
  i: Integer;
  p: PFGToProduceRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p :=  PFGToProduceRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
procedure TFGToProduceReader.SubWinLine(aZMDR001RecordPtr: PZMDR001Record;
  aFGRptWinRecordPtr: PFGRptWinRecord);
var
  i: Integer;
  p: PFGToProduceRecord;
begin
  for i := 0 to Count - 1 do
  begin
    p := Items[i];
    
    if (p^.sproj = aZMDR001RecordPtr^.sProj) and
      ((p^.sitem = aZMDR001RecordPtr^.sVer) or
       (p^.sitem = aZMDR001RecordPtr^.scap) or
       (p^.sitem = aZMDR001RecordPtr^.sColor)
       ) then
    begin
      if TSAPMaterialReader.IsHW(aZMDR001RecordPtr) then
      begin
        p^.dtoproduce_out := p^.dtoproduce_out - aFGRptWinRecordPtr^.dQty;
      end
      else
      begin
        p^.dtoproduce_in := p^.dtoproduce_in - aFGRptWinRecordPtr^.dQty;
      end;
    end; 
  end;
end;  

procedure TFGToProduceReader.SubWin( aFGRptWinReader: TFGRptWinReader;
  aSAPMaterialReader: TSAPMaterialReader; dt1, dt2: TDateTime);
var
  i: Integer;
  aFGRptWinRecordPtr: PFGRptWinRecord;
  aZMDR001RecordPtr: PZMDR001Record;
begin
  dt1 := myStrToDateTime( FormatDateTime('yyyy-MM-dd', dt1) );
  dt2 := myStrToDateTime( FormatDateTime('yyyy-MM-dd', dt2) );

  for i := 0 to aFGRptWinReader.Count - 1 do
  begin
    aFGRptWinRecordPtr := aFGRptWinReader.Items[i];
    if (aFGRptWinRecordPtr^.dt < dt1) or (aFGRptWinRecordPtr^.dt > dt2) then
    begin
      Continue;
    end;
    aZMDR001RecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aFGRptWinRecordPtr^.sNumber);
    if aZMDR001RecordPtr = nil then
    begin
      Log(aFGRptWinRecordPtr^.sNumber + ' �Ҳ������ϻ�������');
      Continue;
    end;

    SubWinLine(aZMDR001RecordPtr, aFGRptWinRecordPtr);
  end;  
end;  

function TFGToProduceReader.GetQty(
  const sProjName, sitem: string; slMap: TStrings; bHW: Boolean; acs: TCountrySet): Double;
var
  i: Integer;
  p: PFGToProduceRecord;
  sitem_p: string;
  idx: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
  begin
    p := Items[i];

    sitem_p := p^.sitem;
    if slMap <> nil then
    begin
      idx := slMap.IndexOfName(sitem_p);
      if idx >= 0 then
      begin
        sitem_p := slMap.ValueFromIndex[idx];
      end;
    end;
    
    if (p^.sproj = sProjName) and (sitem_p = sitem) then
    begin
      if bHW then
      begin

        case acs of
          csC: Continue;
          csR:
            if (Pos('����', p^.sitem) = 0)
              and (Pos('����˹', p^.sitem) = 0) then Continue;
          csU:
            if Pos('�ڿ���', p^.sitem) = 0 then Continue;
          csI:
            if Pos('ӡ��', p^.sitem) = 0 then Continue;
          csO:
            if (Pos('����', p^.sitem) > 0)
              or (Pos('����˹', p^.sitem) > 0)
              or (Pos('�ڿ���', p^.sitem) > 0)
              or (Pos('ӡ��', p^.sitem) > 0) then Continue;
        end;
        Result := Result + p^.dtoproduce_out;
      end
      else
      begin
        Result := Result + p^.dtoproduce_in;
      end;
    end; 
  end;
end;  

procedure TFGToProduceReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;

function TFGToProduceReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TFGToProduceReader.GetItems(i: Integer): PFGToProduceRecord;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PFGToProduceRecord(FList.Objects[i]);
  end
  else Result := nil;
end;

procedure TFGToProduceReader.Open;
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
  stitle1, stitle2, stitle3, stitle4, stitle7, stitle10: string;
  stitle: string;
  irow: Integer;
  sitem: string;
  p: PFGToProduceRecord;
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
        stitle7 := ExcelApp.Cells[irow, 7].Value;   
        stitle10 := ExcelApp.Cells[irow, 10].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle7 + stitle10;

        if stitle <> '���Ͳ�Ʒ������ɫ\��ʽ\������Ŀ����������������������' then
        begin

          Log(sSheet +'  ���� ��Ʒ���� ��� ��ʽ');
          Continue;
        end;
             
        FReadOk := True;
  
        irow := 3;
        sitem := ExcelApp.Cells[irow, 3].Value; 
        while sitem <> '' do
        begin
          if sitem <> 'TOTAL' then
          begin
            p := New(PFGToProduceRecord);
            if not IsCellMerged(ExcelApp, irow, 1, irow - 1, 1) then
            begin
              p^.sproj := ExcelApp.Cells[irow, 1].Value; //����
              p^.sprojname := ExcelApp.Cells[irow, 2].Value; //��Ʒ����
            end;
            p^.sitem := sitem; //��ɫ\��ʽ\����
            p^.dtoproduce_in := ExcelApp.Cells[irow, 10].Value; //��������		    ����        10
            p^.dtoproduce_out := ExcelApp.Cells[irow, 11].Value; //��������      С��        11

            FList.AddObject(sitem, TObject(p));
          end;

          irow := irow + 1;
          sitem := ExcelApp.Cells[irow, 3].Value; 
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
