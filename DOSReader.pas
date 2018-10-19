unit DOSReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TDOSRecord = packed record
    snumber: string;
    sproj: string;
    dos: Double;
    dos_dest: Double;
  end;
  PDOSRecord = ^TDOSRecord;

  TDOSReader = class
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

{ TDOSReader }

constructor TDOSReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TDOSReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TDOSReader.Clear;
var
  i: Integer;
  p: PDOSRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PDOSRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

procedure TDOSReader.Log(const str: string);
begin

end;

procedure TDOSReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  p: PDOSRecord;
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

        irow := 2;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle := stitle1 + stitle2;
//        if stitle <> 'ProjectSKU No' then
//        begin
//          Log(sSheet +'  �����¾�SKU��ʽ');
//          Continue;
//        end;

        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4;
        if stitle <> 'ProjectSKU NoDOSĿ��' then
        begin
          Log(sSheet +'  �����¾�SKU��ʽ');
          Continue;
        end;


        irow := 3;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin                                
          p := New(PDOSRecord);
          FList.AddObject(snumber, TObject(p));

          p^.snumber := snumber;
          p^.sproj := ExcelApp.Cells[irow, 1].Value;
          p^.dos := ExcelApp.Cells[irow, 3].Value;
//          p^.dos_dest := ExcelApp.Cells[irow, 4].Value;

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 2].Value;
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
