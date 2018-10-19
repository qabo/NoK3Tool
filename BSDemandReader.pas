unit BSDemandReader;

interface
      
uses
  Windows, Classes, SysUtils, ComObj, Variants, CommUtils;

type
  TBSDemand = class
  public
    FNumber99: string;
    FDate: string;
    FQty: Double;
  end;

  TBSDemandReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Log(const str: string);
    procedure Open;
  public 
    FList: TList;             
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
  end;

implementation

{ TBSDemandReader }

constructor TBSDemandReader.Create(const sfile: string);
begin
  FFile := sfile;

  FList := TList.Create;

  Open;

end;

destructor TBSDemandReader.Destroy; 
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TBSDemandReader.Clear;
var
  i: Integer;
  aBSDemand: TBSDemand;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aBSDemand := TBSDemand(FList[i]);
    aBSDemand.Free;
  end;
  FList.Clear;
end;
      
procedure TBSDemandReader.Log(const str: string);
begin

end;

procedure TBSDemandReader.Open;

const
  CINumber99 = 2; // 子物料编码    
  CINumber = 5; // 子物料编码
  
var
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  stitle: string;
  irow: Integer;
  icol: Integer;
  snumber: string;
  snumber99: string;
  bFound: Boolean;
  irow1: Integer;
  slDate: TStringList;
  v: Variant;
  dt: TDateTime;
  i: Integer;    
  aBSDemand: TBSDemand;
  dQty: Double;
begin
  Clear;

  slDate := TStringList.Create;

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

        bFound := False;
        for irow := 1 to 10 do
        begin
          stitle1 := ExcelApp.Cells[irow, 1].Value;
          stitle2 := ExcelApp.Cells[irow, 2].Value;
          stitle3 := ExcelApp.Cells[irow, 3].Value;
          stitle4 := ExcelApp.Cells[irow, 4].Value;
          stitle5 := ExcelApp.Cells[irow, 5].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
          if stitle = 'Item99料号闻泰物料编码主物料编码子物料编码' then
          begin
            bFound := True;
            Break;
          end;
        end;

        if not bFound then
        begin     
          Log(sSheet + '  格式不符合');
          Continue;
        end;

        icol := 12;
        v := ExcelApp.Cells[irow, icol].Value;
        while VarIsType(v, varDate) do
        begin
          dt := v;
          slDate.AddObject(FormatDateTime('yyyy-MM-dd', dt), TObject(icol));
          icol := icol + 1;      
          v := ExcelApp.Cells[irow, icol].Value;
        end;
        
        irow := irow + 1;   // 跳过标题栏，到数据开始行

        irow1 := irow;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        while snumber <> '' do
        begin

          while IsCellMerged(ExcelApp, irow1, CINumber99, irow, CINumber99) do
          begin
            irow := irow + 1;
            Continue;
          end;

          snumber99 := ExcelApp.Cells[irow1, CINumber99].Value;
          for i := 0 to slDate.Count - 1 do
          begin
            icol := Integer(slDate.Objects[i]);

            dQty := 0;
            v := ExcelApp.Cells[irow1, icol].Value;
            if VarIsNumeric(v) then
              dQty := v
            else
            begin
              Log('irow: ' + IntToStr(irow1) + '  icol: ' + IntToStr(icol) + ' is not a valid number ' + v);
            end;


            if dQty > 0 then
            begin
              aBSDemand := TBSDemand.Create;

              aBSDemand.FNumber99 := snumber99;
              aBSDemand.FDate := slDate[i];
              aBSDemand.FQty := dQty;

              FList.Add(aBSDemand);
            end;
            

          end;
          
          irow1 := irow;
          snumber := ExcelApp.Cells[irow, CINumber].Value;
        end;

        Break;

      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit;

    slDate.Free;
  end;            
end;

end.
