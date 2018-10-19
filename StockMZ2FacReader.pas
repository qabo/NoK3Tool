unit StockMZ2FacReader;

interface

uses
  Classes, SysUtils, ComObj, CommUtils, KeyICItemSupplyReader;

type
  TStockMZ2FacReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Log(const str: string);
  protected
    procedure Open; virtual; abstract;
  public
    FList: TStringList;   
    FList_no: TStringList;
    slToName: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function Fac2MZ(const sfac: string): string;     
    function Fac2MZ_no(const sfac: string): string;
    function ToName(const sstock: string): string;
  end;

  TStockMZ2FacReader_ml = class(TStockMZ2FacReader)
  protected
    procedure Open; override;
  end;    

  TStockMZ2FacReader_wt = class(TStockMZ2FacReader)
  protected
    procedure Open; override;
  end;      

  TStockMZ2FacReader_yd = class(TStockMZ2FacReader)
  protected
    procedure Open; override;
  end;

implementation
       
{ TStockMZ2FacReader }

constructor TStockMZ2FacReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  FList_no := TStringList.Create;
  slToName := TStringList.Create;
  Open; 
end;

destructor TStockMZ2FacReader.Destroy;
begin
  Clear;
  FList.Free;
  FList_no.Free;
  slToName.Free;
  inherited;
end;

procedure TStockMZ2FacReader.Clear;
begin
  FList.Clear;
  FList_no.Clear;
  slToName.Clear;
end;

procedure TStockMZ2FacReader.Log(const str: string);
begin

end;

function TStockMZ2FacReader.Fac2MZ(const sfac: string): string;
begin
  if FList.IndexOfName(sfac) < 0 then
  begin
    Result := sfac;
  end
  else
  begin
    Result := FList.Values[sfac];
  end;
end;       

function TStockMZ2FacReader.Fac2MZ_no(const sfac: string): string;
begin
  if FList_no.IndexOfName(sfac) < 0 then
  begin
    Result := sfac;
  end
  else
  begin
    Result := FList_no.Values[sfac];
  end;
end;

function TStockMZ2FacReader.ToName(const sstock: string): string;
begin
  Result := slToName.Values[sstock];
end;


{ TStockMZ2FacReader_ml }

procedure TStockMZ2FacReader_ml.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9, stitle10: string;
  stitle: string;
  irow: Integer; 
  sfac, smeizu: string; 
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
        stitle6 := ExcelApp.Cells[irow, 6].Value;                                    
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := Trim(stitle1) + Trim(stitle2) + Trim(stitle3) + Trim(stitle4) + Trim(stitle5) + Trim(stitle6) + Trim(stitle7);
        if stitle <> 'storuomstorDescriptionmzDescSAPmzDesc' then
        begin
          Log(sSheet +'  不是  魅族代工厂仓库对照表  格式');
          Continue;
        end;

        irow := 2;
        smeizu := ExcelApp.Cells[irow, 5].Value;
        while smeizu <> '' do
        begin
          sfac := ExcelApp.Cells[irow, 4].Value;

          FList.Add(sfac + '=' + smeizu);

          irow := irow + 1;
          smeizu := ExcelApp.Cells[irow, 5].Value;
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

{ TStockMZ2FacReader_wt }

procedure TStockMZ2FacReader_wt.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9, stitle10: string;
  stitle: string;
  irow: Integer; 
  sfac, smeizu: string;

  s2, s12, s13: string;
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
        stitle6 := ExcelApp.Cells[irow, 6].Value;                                    
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := Trim(stitle1) + Trim(stitle2) + Trim(stitle3) + Trim(stitle4) + Trim(stitle5) + Trim(stitle6) + Trim(stitle7);
        if stitle <> 'plant闻泰仓位DescriptionuommzDescmzStro' then
        begin
          Log(sSheet +'  不是  魅族代工厂仓库对照表（闻泰: plant闻泰仓位DescriptionuommzDescmzStro）  格式');
          Continue;
        end;

        irow := 2;
        smeizu := ExcelApp.Cells[irow, 13].Value;
        while smeizu <> '' do
        begin
          sfac := ExcelApp.Cells[irow, 4].Value;

          FList.Add(sfac + '=' + smeizu);

          s2 := ExcelApp.Cells[irow, 2].Value;
          s12 := ExcelApp.Cells[irow, 12].Value;
          s13 := ExcelApp.Cells[irow, 13].Value;
          slToName.Values[s12] := s13;

          FList_no.Values[s2] := s12;

          irow := irow + 1;
          smeizu := ExcelApp.Cells[irow, 13].Value;
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
            

{ TStockMZ2FacReader_yd }

procedure TStockMZ2FacReader_yd.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9, stitle10: string;
  stitle: string;
  irow: Integer; 
  sfac, smeizu: string;   
  sfac_no, smeizu_no: string;
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
        stitle6 := ExcelApp.Cells[irow, 6].Value;                                    
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := Trim(stitle1) + Trim(stitle2) + Trim(stitle3) + Trim(stitle4) + Trim(stitle5) + Trim(stitle6) + Trim(stitle7);
        if stitle <> '组织+代码组织SECONDARY_INVENTORY_NAMEDESCRIPTIONSTATUS_CODEATTRIBUTE1魅族仓库名称' then
        begin
          Log(sSheet +'  不是  魅族代工厂仓库对照表（与德: 组织+代码组织SECONDARY_INVENTORY_NAMEDESCRIPTIONSTATUS_CODEATTRIBUTE1魅族仓库名称）  格式');
          Continue;
        end;

        irow := 2;
        smeizu := ExcelApp.Cells[irow, 7].Value;
        while smeizu <> '' do
        begin
          sfac := ExcelApp.Cells[irow, 4].Value;
          sfac_no := ExcelApp.Cells[irow, 1].Value;
          smeizu_no := ExcelApp.Cells[irow, 8].Value;

          FList.Add(sfac + '=' + smeizu);
          FList_no.Add(sfac_no + '=' + smeizu_no);

          irow := irow + 1;
          smeizu := ExcelApp.Cells[irow, 7].Value;
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
