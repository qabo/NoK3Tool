unit LTPCMSConfirmReader;

interface

uses
  Classes, ComObj, ActiveX, SysUtils, Windows, CommUtils, DateUtils, Variants,
  SAPStockReader, SBomReader;

type 
  TLTPCMSConfirmRecord = packed record
    dtCreateDate: TDateTime; //单据创建日期
    sNumber: string; //物料编码
    sName: string; //物料名称
    sUnit: string; //单位
    dQtyNeed: Double; //需求数量
    dtDateNeed: TDateTime; //计划到料日期
    sBuyerNo: string; //采购员代码
    sBuyerName: string; //采购员名称
    dQtyConfirm: Double; //回复数量
    dtConfirm: TDateTime; //回复到料日期
  end;
  PLTPCMSConfirmRecord = ^TLTPCMSConfirmRecord;

  TTPCMSConfirmReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
  public
    FList: TList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    procedure GetNumberList(slNumbers: TStringList; var dt1, dt2: TDateTime);     
    function GetQtyDemand(sNumber: string; dt: TDateTime): Double;
    function GetQtyConfirm(sNumber: string; dt: TDateTime): Double;

    procedure Save(aSAPStockReader: TSAPStockReader; aSBomReader: TSBomReader;
      const sfile: string);
  end;

implementation

const
  CICreateDate = 1; //单据创建日期
  CINumber = 2; //物料编码
  CIName = 3; //物料名称
  CIUnit = 4; //单位
  CIQtyNeed = 5; //需求数量
  CIDateNeed = 6; //计划到料日期
  CIBuyerNo = 7; //采购员代码
  CIBuyerName = 8; //采购员名称
  CIQtyConfirm = 12; //回复数量
  CIDateConfirm = 13; //回复到料日期
 
{ TTPCMSConfirmReader }

constructor TTPCMSConfirmReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TList.Create;
  Open;
end;

destructor TTPCMSConfirmReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TTPCMSConfirmReader.Clear;
var
  i: Integer;
  p: PLTPCMSConfirmRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PLTPCMSConfirmRecord(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

procedure TTPCMSConfirmReader.GetNumberList(slNumbers: TStringList;
  var dt1, dt2: TDateTime);
var
  i: Integer;
  p: PLTPCMSConfirmRecord;
begin
  dt2 := myStrToDateTime('1900-01-01');
  dt1 := myStrToDateTime('2100-01-01');
  slNumbers.Clear;
  for i := 0 to FList.Count - 1 do
  begin
    p := PLTPCMSConfirmRecord(FList[i]);
    if p^.dQtyConfirm > 0 then 
    begin
      if dt1 > p^.dtConfirm then
      begin
        dt1 := p^.dtConfirm;
      end;
      if dt2 < p^.dtConfirm then
      begin
        dt2 := p^.dtConfirm;
      end;
    end;
    
    if slNumbers.IndexOf(p^.sNumber) < 0 then
    begin
      slNumbers.AddObject(p^.sNumber, TObject(p));
    end;
  end;
end;

function TTPCMSConfirmReader.GetQtyDemand(sNumber: string; dt: TDateTime): Double;
var
  i: Integer;
  p: PLTPCMSConfirmRecord;
begin
  Result := 0;
  
  for i := 0 to FList.Count - 1 do
  begin
    p := PLTPCMSConfirmRecord(FList[i]);
    if p^.dtConfirm <> dt then Continue;
    if p^.sNumber <> sNumber then Continue;
    Result := Result + p^.dQtyNeed;
  end;
end;        

function TTPCMSConfirmReader.GetQtyConfirm(sNumber: string; dt: TDateTime): Double;
var
  i: Integer;
  p: PLTPCMSConfirmRecord;
begin
  Result := 0;
  
  for i := 0 to FList.Count - 1 do
  begin
    p := PLTPCMSConfirmRecord(FList[i]);
    if p^.dtConfirm <> dt then Continue;
    if p^.sNumber <> sNumber then Continue;
    Result := Result + p^.dQtyConfirm;
  end;
end;

procedure TTPCMSConfirmReader.Log(const str: string);
begin

end;
 
procedure TTPCMSConfirmReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  snumber: string;    
  p: PLTPCMSConfirmRecord;
  sdate: string;
  v: Variant;
  s: string;
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
        stitle := stitle1 + stitle2 + stitle3 + stitle4;
                     
        if stitle <> '单据创建日期物料编码物料名称单位' then
        begin
          Log(sSheet +'  不是SAP导出BOM格式');
          Continue;
        end;


        irow := 2;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        while snumber <> '' do
        begin                   
          p := New(PLTPCMSConfirmRecord);
          FList.Add(p);

          sdate := ExcelApp.Cells[irow, CICreateDate].Value;
          sdate := StringReplace(sdate, '/', '-', [rfReplaceAll]);
          p^.dtCreateDate := myStrToDateTime(sdate);
          
          p^.sNumber := snumber;
          p^.sName := ExcelApp.Cells[irow, CIName].Value;
          p^.sUnit := ExcelApp.Cells[irow, CIUnit].Value;
          p^.dQtyNeed := ExcelApp.Cells[irow, CIQtyNeed].Value;

          sdate := ExcelApp.Cells[irow, CIDateNeed].Value;
          sdate := StringReplace(sdate, '/', '-', [rfReplaceAll]);
          p^.dtDateNeed :=  myStrToDateTime(sdate);
          
          p^.sBuyerNo := ExcelApp.Cells[irow, CIBuyerNo].Value;
          p^.sBuyerName := ExcelApp.Cells[irow, CIBuyerName].Value;
          v := ExcelApp.Cells[irow, CIQtyConfirm].Value;
          if VarIsNumeric(v) then
          begin
            p^.dQtyConfirm := v;
          end
          else if VarIsStr(v) then
          begin
            s := v;
            p^.dQtyConfirm := StrToFloatDef(s, 0);
          end;

          if p^.dQtyConfirm > 0 then
          begin
            sdate := ExcelApp.Cells[irow, CIDateConfirm].Value;
            sdate := StringReplace(sdate, '/', '-', [rfReplaceAll]);
            p^.dtConfirm := myStrToDateTime(sdate);
          end;


          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, CINumber].Value;
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

procedure TTPCMSConfirmReader.Save(aSAPStockReader: TSAPStockReader;
  aSBomReader: TSBomReader; const sfile: string);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  i: Integer;
  p: PLTPCMSConfirmRecord;
  slNumbers: TStringList;
  dt1, dt2: TDateTime;
  col1: Integer;
  col: Integer;
  idate: Integer;
  dc: Integer;
  dt: TDateTime;
begin


  // 开始保存 Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;

  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;

  slNumbers := TStringList.Create;
  GetNumberList(slNumbers, dt1, dt2);
  if dt1 > dt2 then
  begin
    dc := 0;
  end
  else
  begin
    dc := DaysBetween(dt2, dt1);
  end;

  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := '采购交期回复';
  try
    irow := 3;

    ExcelApp.Cells[irow, 1].Value := '物料类别';
    ExcelApp.Cells[irow, 2].Value := '物料编码';
    ExcelApp.Cells[irow, 3].Value := '物料名称';
    ExcelApp.Cells[irow, 4].Value := '子料名称';
    ExcelApp.Cells[irow, 5].Value := '所用项目';
    ExcelApp.Cells[irow, 6].Value := '半成品库存';
    ExcelApp.Cells[irow, 7].Value := '可用库存';
    ExcelApp.Cells[irow, 8].Value := '新外购入库';
    ExcelApp.Cells[irow, 9].Value := 'MRP当天外购入库';    
    ExcelApp.Cells[irow, 10].Value := 'Commitment';
                                            
    ExcelApp.Columns[2].ColumnWidth := 16;
    ExcelApp.Columns[3].ColumnWidth := 20;
    ExcelApp.Columns[10].ColumnWidth := 18;

    col1 := 11;
    for idate := 0 to dc - 1 do
    begin
      col := col1 + idate; 
      ExcelApp.Cells[irow - 1, col].Value := 'WK' + IntToStr(WeekOf(dt1 + idate));
      dt := dt1 + idate;
      ExcelApp.Cells[irow, col].Value := dt;
    end;
                                                          
    AddColor(ExcelApp, irow - 1, col1 - 1, irow - 1, col1 + dc - 1, $E8DEB7);
    AddColor(ExcelApp, irow, 1, irow, col1 + dc - 1, $E8DEB7);
          
    irow := 4;
    for i := 0 to slNumbers.Count - 1 do
    begin

      p := PLTPCMSConfirmRecord(slNumbers.Objects[i]);

      ExcelApp.Cells[irow, 1].Value := '';
      ExcelApp.Cells[irow, 2].Value := p^.sNumber;
      ExcelApp.Cells[irow, 3].Value := p^.sName;
      ExcelApp.Cells[irow, 4].Value := '';
      ExcelApp.Cells[irow, 5].Value := '';
      ExcelApp.Cells[irow, 6].Value := aSBomReader.GetAvailStockSemi(p^.sNumber);
      ExcelApp.Cells[irow, 7].Value := aSAPStockReader.GetAvailStock(p^.sNumber);
      ExcelApp.Cells[irow, 8].Value := '';                                    
      ExcelApp.Cells[irow, 9].Value := '';
      ExcelApp.Cells[irow, 10].Value := 'Demand By MPS';
      ExcelApp.Cells[irow + 1, 10].Value := 'Demand By 要货计划';
      ExcelApp.Cells[irow + 2, 10].Value := 'Delta MPS';
      ExcelApp.Cells[irow + 3, 10].Value := 'Confirm supply';

      
      MergeCells(ExcelApp, irow, 1, irow + 3, 1); 
      MergeCells(ExcelApp, irow, 2, irow + 3, 2);
      MergeCells(ExcelApp, irow, 3, irow + 3, 3);
      MergeCells(ExcelApp, irow, 4, irow + 3, 4);
      MergeCells(ExcelApp, irow, 5, irow + 3, 5);
      MergeCells(ExcelApp, irow, 6, irow + 3, 6);
      MergeCells(ExcelApp, irow, 7, irow + 3, 7);
      MergeCells(ExcelApp, irow, 8, irow + 3, 8);

      col1 := 11;
      for idate := 0 to dc - 1 do
      begin
        col := col1 + idate;
        //ExcelApp.Cells[irow, col].Value := 'Demand By MPS';
        ExcelApp.Cells[irow + 1, col].Value := GetQtyDemand(p^.sNumber, dt1 + idate);
        ExcelApp.Cells[irow + 2, col].Value := '=' + GetRef(col) + IntToStr(irow + 1) + '-' + GetRef(col) + IntToStr(irow);
        ExcelApp.Cells[irow + 3, col].Value := GetQtyConfirm(p^.sNumber, dt1 + idate);
      end;
               
      AddColor(ExcelApp, irow + 2, col1 - 1, irow + 2, col1 + dc - 1, $E4DCD6); 
      AddColor(ExcelApp, irow + 3, col1 - 1, irow + 3, col1 + dc - 1, $D6E4FC);

      irow := irow + 4;
    end;

    col1 := 10;
    AddBorder(ExcelApp, 1, 1, irow - 1, col1 + dc);

  
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;

    slNumbers.Free;
  end; 

end;    

end.
