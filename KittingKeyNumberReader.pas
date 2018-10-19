unit KittingKeyNumberReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TKittingKeyNumber = class
  private
    function GetName: string;
  public
    id     : string;     //  �ؼ���
    sproj  : string;  //  ��Ŀ
    sname  : string;  //  ����
    scat   : string;   //  ����
    sver   : string;   //  ����BOM�б�׼��ʽ
    scap   : string;   //	����BOM������
    scolor : string; //	����BOM����ɫ
    snumber: string;//	���ϱ���
    dlt    : Double;    //	 ������ǰ��
    dusage : Double; //	����

    property name: string read GetName;
  end;
  
  TKittingKeyNumberReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FList: TStringList;

    procedure Open;
    procedure Log(const str: string);

    function GetCount: Integer;
    function GetItems(i: Integer): TKittingKeyNumber;
  public 
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;

    property Count: Integer read GetCount;
    property Items[i: Integer]: TKittingKeyNumber read GetItems;
  end;

implementation

{ TKittingKeyNumber }

function TKittingKeyNumber.GetName: string;
begin
  Result := sproj + '_' + sname + '_' + scat;
end;

{ TKittingKeyNumberReader }

constructor TKittingKeyNumberReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TKittingKeyNumberReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TKittingKeyNumberReader.Clear;
var
  i: Integer;
  aKittingKeyNumber: TKittingKeyNumber;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aKittingKeyNumber := TKittingKeyNumber(FList.Objects[i]);
    aKittingKeyNumber.Free;
  end;
  FList.Clear;
end;

procedure TKittingKeyNumberReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function TKittingKeyNumberReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TKittingKeyNumberReader.GetItems(i: Integer): TKittingKeyNumber;
begin
  Result := TKittingKeyNumber(FList.Objects[i]);
end;  

procedure TKittingKeyNumberReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4,
  stitle5, stitle6, stitle7: string;
  stitle: string;
  irow: Integer; 

  snumber: string;

  id: string;      //  �ؼ���
  sproj: string;   //  ��Ŀ
  sname: string;   //  ����
  scat: string;    //  ����
  sver: string;    //  ����BOM�б�׼��ʽ
  scap: string;    //	����BOM������
  scolor: string;  //	����BOM����ɫ

  id0: string;     //  �ؼ���
  sproj0: string;  //  ��Ŀ
  sname0: string;  //  ����
  scat0: string;   //  ����
  sver0: string;   //  ����BOM�б�׼��ʽ
  scap0: string;   //	����BOM������
  scolor0: string; //	����BOM����ɫ  

  aKittingKeyNumber: TKittingKeyNumber;
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
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 +
          stitle5 + stitle6 + stitle7;
        if stitle <> '�ؼ�����Ŀ���Ϸ�������BOM�б�׼��ʽ����BOM����������BOM����ɫ' then
        begin
          Log(sSheet +'  ����   �ֹ��ؼ������嵥');
          Continue;
        end;
 
        irow := 2;
        snumber := ExcelApp.Cells[irow, 8].Value;
        while snumber <> '' do
        begin
          aKittingKeyNumber := TKittingKeyNumber.Create;
          FList.AddObject(snumber, aKittingKeyNumber);
                                             
          id     := ExcelApp.Cells[irow, 1].Value;
          sproj  := ExcelApp.Cells[irow, 2].Value;
          sname  := ExcelApp.Cells[irow, 3].Value;
          scat   := ExcelApp.Cells[irow, 4].Value;
          sver   := ExcelApp.Cells[irow, 5].Value;
          scap   := ExcelApp.Cells[irow, 6].Value;
          scolor := ExcelApp.Cells[irow, 7].Value;

          if id <> '' then id0 := id;
          if sproj <> '' then sproj0 := sproj;
          if sname <> '' then sname0 := sname;
          if scat <> '' then scat0 := scat;
          if sver <> '' then sver0 := sver;
          if scap <> '' then scap0 := scap;
          if scolor <> '' then scolor0 := scolor;

          aKittingKeyNumber.id := id0;
          aKittingKeyNumber.sproj := sproj0;
          aKittingKeyNumber.sname := sname0;
          aKittingKeyNumber.scat := scat0;
          aKittingKeyNumber.sver := sver0;
          aKittingKeyNumber.scat := scap0;
          aKittingKeyNumber.scolor := scolor0;
          aKittingKeyNumber.snumber := snumber;
          aKittingKeyNumber.dlt := ExcelApp.Cells[irow, 9].Value;
          aKittingKeyNumber.dusage := ExcelApp.Cells[irow, 10].Value;
 
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 8].Value;
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
