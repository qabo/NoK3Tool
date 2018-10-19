unit SimMPSReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type 
  TSimMPSReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
  public
    FList: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
  end;

implementation
  
{ TSimMPSReader }

constructor TSimMPSReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TSimMPSReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSimMPSReader.Clear; 
begin
  FList.Clear;
end;

procedure TSimMPSReader.Log(const str: string);
begin

end;

procedure TSimMPSReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  dQty: Integer;
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
        stitle := stitle1 + stitle2; 
        if stitle <> '�Ϻ�����' then
        begin
          Log(sSheet +'  ����  Ӫ��ģ��MPS  ��ʽ  �Ϻ�����');
          Continue;
        end;
 
        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          dQty := ExcelApp.Cells[irow, 2].Value;
          FList.AddObject(snumber, TObject(dQty));
          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 1].Value;
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
