unit SAPMaterialReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TSAPMaterialRecord = packed record
    sNumber: string;
    sName: string;
    dRecvTime: Double;
    sMRPType: string; // PD mrp, �⹺  M0 mps ���ư��Ʒ
    dMOQ: Double; // 
    dSPQ: Double; // = '����ֵ';
    dLT_PD: Double; // = '�ƻ�����ʱ��';
    dLT_M0: Double; // = '��������ʱ��';
    iLowestCode: Integer; // ��λ��  
    sMRPer: string;
    sMRPerDesc: string;
    sBuyer: string;
    sMRPGroup: string;
    sPType: string;
    sAbc: string;
    sGroupName: string;
    sPlanNumber: string;
    sMMTypeDesc: string;
  end;
  PSAPMaterialRecord = ^TSAPMaterialRecord;
   
  TSAPMaterialReader = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string); 
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPMaterialRecord;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear; 
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPMaterialRecord read GetItems;
    function GetSAPMaterialRecord(const snumber: string): PSAPMaterialRecord;
  end;

implementation
       
{ TSAPMaterialReader }

constructor TSAPMaterialReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPMaterialReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPMaterialReader.Clear;
var
  i: Integer;
  p: PSAPMaterialRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPMaterialRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPMaterialReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPMaterialReader.GetItems(i: Integer): PSAPMaterialRecord;
begin
  Result := PSAPMaterialRecord(FList.Objects[i]);
end;

function TSAPMaterialReader.GetSAPMaterialRecord(const snumber: string): PSAPMaterialRecord;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PSAPMaterialRecord(FList.Objects[idx]);
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
  CSNumber = '���ϱ���';
  CSName = '��������';
  CSRecvTime = '�ջ�����ʱ��';
  CSMRPType = 'MRP����';
  CSSPQ = '����ֵ';
  CSLT_PD = '�ƻ�����ʱ��';
  CSLT_M0 = '��������ʱ��';
  CSLT_MOQ = '��С������С';
  CSMRPGroup = 'MRP��';
  CSPType = '�ɹ�����';
  CSAbc = 'ABC��ʶ';
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  iColNumber: Integer;
  iColName: Integer;
  iColRecvTime: Integer;
  iColMRPType: Integer;
  iColSPQ: Integer;
  iColLT_PD: Integer;
  iColLT_M0: Integer;
  iColMOQ: Integer;
  iMRPGroup: Integer;
  iPType: Integer;
  iAbc: Integer;
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
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;
        if stitle <> '���������������ϱ�������������������������������' then
        begin
          Log(sSheet +'  ����  SAP��������  ��ʽ');
          Continue;
        end;

        
        iColNumber := IndexOfCol(ExcelApp, irow, CSNumber);
        iColName := IndexOfCol(ExcelApp, irow, CSName);
        iColRecvTime := IndexOfCol(ExcelApp, irow, CSRecvTime);
        iColMRPType := IndexOfCol(ExcelApp, irow, CSMRPType);
        iColSPQ := IndexOfCol(ExcelApp, irow, CSSPQ);
        iColLT_PD := IndexOfCol(ExcelApp, irow, CSLT_PD);
        iColLT_M0 := IndexOfCol(ExcelApp, irow, CSLT_M0);
        iColMOQ := IndexOfCol(ExcelApp, irow, CSLT_MOQ);
        iMRPGroup := IndexOfCol(ExcelApp, irow, CSMRPGroup);
        iPType := IndexOfCol(ExcelApp, irow, CSPType);
        iAbc := IndexOfCol(ExcelApp, irow, CSAbc);

        if (iColNumber = -1) or (iColName = -1) or (iColRecvTime = -1)
          or (iColMRPType = -1) or (iColSPQ = -1) or (iColLT_PD = -1)
          or (iColLT_M0 = -1) or (iColMOQ = -1) or (iMRPGroup = -1)
          or (iPType = -1) or (iAbc = -1)
          then
        begin
          Log(sSheet +'  ����  SAP��������  ��ʽ');
          Continue;
        end;
                
        irow := 2;
        snumber := ExcelApp.Cells[irow, iColNumber].Value;
        while snumber <> '' do
        begin                                
          aSAPMaterialRecordPtr := New(PSAPMaterialRecord);
          FList.AddObject(snumber, TObject(aSAPMaterialRecordPtr));

          aSAPMaterialRecordPtr^.sNumber := snumber;
          aSAPMaterialRecordPtr^.sName := ExcelApp.Cells[irow, iColName].Value;
          aSAPMaterialRecordPtr^.dRecvTime := ExcelApp.Cells[irow, iColRecvTime].Value;
          aSAPMaterialRecordPtr^.sMRPType := ExcelApp.Cells[irow, iColMRPType].Value;
          aSAPMaterialRecordPtr^.dSPQ := ExcelApp.Cells[irow, iColSPQ].Value;            
          aSAPMaterialRecordPtr^.dMOQ := ExcelApp.Cells[irow, iColMOQ].Value;
          aSAPMaterialRecordPtr^.dLT_PD := ExcelApp.Cells[irow, iColLT_PD].Value;
          aSAPMaterialRecordPtr^.dLT_M0 := ExcelApp.Cells[irow, iColLT_M0].Value;
          aSAPMaterialRecordPtr^.iLowestCode := 0; // ��λ�룬Ĭ��Ϊ1
          aSAPMaterialRecordPtr^.sMRPGroup := ExcelApp.Cells[irow, iMRPGroup].Value;
          aSAPMaterialRecordPtr^.sPType := ExcelApp.Cells[irow, iPType].Value;
          aSAPMaterialRecordPtr^.sAbc := ExcelApp.Cells[irow, iAbc].Value;
 
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, iColNumber].Value;
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
