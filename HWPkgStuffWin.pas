unit HWPkgStuffWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin, ComObj, Math, CommUtils;

type
  TMDemand = packed record
    Number: string;
    Name: string;
    DemandType: string;
    sUnit: string;
    SourceType: string;
    SourceBillNo: string;
    SourceEntryID: string;
    Note: string;
    Qty: Double;
    Date1: TDateTime;
    Date2: TDateTime;
  end;
  PMDemand = ^TMDemand;

  TProjDayDemand = packed record
    Date1: TDateTime;
    Date2: TDateTime;
    Qty: Double;
    DemandType: string;
    sUnit: string;
    SourceType: string;
    SourceBillNo: string;
    SourceEntryID: string;
    Note: string;
  end;
  PProjDayDemand = ^TProjDayDemand;

  TProjDemand = class
  private
    FList: TList;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    function IndexOf(dt: TDateTime): Integer;
  end;

  TPkgItem = class
  private
    FNumbers: TStringList;
    FUsage: Double;
    F99: string;
    FFGNumber: string;
  public
    constructor Create;
    destructor Destroy; override;   
    procedure Clear;
  end;

  TPkgStuff = class
  private
    FItems: TList;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
  end;

  TfrmHWPkgStuff = class(TForm)
    leMDemand: TLabeledEdit;
    btnMDemand: TButton;
    OpenDialog1: TOpenDialog;
    lePkgStuff: TLabeledEdit;
    btnPkgStuff: TButton;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ImageList1: TImageList;
    leProjNumber: TLabeledEdit;
    btnProjNumber: TButton;
    Memo1: TMemo;
    SaveDialog1: TSaveDialog;
    procedure btnMDemandClick(Sender: TObject);
    procedure btnPkgStuffClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnProjNumberClick(Sender: TObject);
  private
    { Private declarations }
    FMDemands: TList;
    FProjMDemand: TStringList;
    FProjNumber: TStringList;
    FPkgStuff: TStringList;
    procedure Clear;
//    procedure OpenMDemand(const sfile: string);       
    procedure OpenMDemand2(const sfile: string);
    procedure OpenProjNumber(const sfile: string);
    procedure OpenPkgStuff(const sfile: string);
    procedure SumMDemand;
    procedure MakePkgStuffDemand(const sfile: string);
    function GetFGDemand(const sdate: string; slFGNumber: TStringList; var aProjDayDemand: TProjDayDemand): Boolean;
  public
    { Public declarations }
    class procedure ShowForm;
  end;



implementation

{$R *.dfm}

var
  gFormatSettings: TFormatSettings;

class procedure TfrmHWPkgStuff.ShowForm;   
var
  frmHWPkgStuff: TfrmHWPkgStuff;
begin
  frmHWPkgStuff := TfrmHWPkgStuff.Create(nil);
  try
    frmHWPkgStuff.ShowModal;
  finally
    frmHWPkgStuff.Free;
  end;
end;

function CellMerged(ExcelApp: Variant; irow1, irow2: Integer; icol: Integer): Boolean;
var
  vma1, vma2: Variant;
begin
  vma1 := ExcelApp.Cells[irow1, icol].MergeArea;
  vma2 := ExcelApp.Cells[irow2, icol].MergeArea;
  Result := vma1.Address = vma2.Address;
end;
   
{ TProjDemand }

constructor TProjDemand.Create;
begin
  FList := TList.Create;
end;

destructor TProjDemand.Destroy;
begin
  Clear;
  FList.Free;
end;

procedure TProjDemand.Clear;
var
  i: Integer;
  aProjDayDemand: PProjDayDemand;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aProjDayDemand := PProjDayDemand(FList[i]);
    Dispose(aProjDayDemand);
  end;
  FList.Clear;
end;

function TProjDemand.IndexOf(dt: TDateTime): Integer;
var
  i: Integer;
  aProjDayDemand: PProjDayDemand;
begin
  Result := -1;
  for i := 0 to FList.Count - 1 do
  begin
    aProjDayDemand := PProjDayDemand(FList[i]);
    if aProjDayDemand.Date1 = dt then
    begin
      Result := i;
      Break;
    end;
  end;
end;  

{ TPkgItem }

constructor TPkgItem.Create;
begin
  FNumbers := TStringList.Create;
end;

destructor TPkgItem.Destroy;
begin
  Clear;
  FNumbers.Free;
end;

procedure TPkgItem.Clear;
begin
  FNumbers.Clear;
end;

{ TPkgStuff }

constructor TPkgStuff.Create;
begin
  FItems := TList.Create;
end;

destructor TPkgStuff.Destroy;
begin
  Clear;
  FItems.Free;
  inherited;
end;

procedure TPkgStuff.Clear;
var
  i: Integer;
  aPkgItem: TPkgItem;
begin
  for i := 0 to FItems.Count - 1 do
  begin
    aPkgItem := TPkgItem(FItems[i]);
    aPkgItem.Free;
  end;
  FItems.Clear;
end;




procedure TfrmHWPkgStuff.FormCreate(Sender: TObject);
begin
  GetLocaleFormatSettings(GetThreadLocale, gFormatSettings);
  gFormatSettings.DateSeparator := '-';

  Memo1.Clear;
  leProjNumber.Text := 'C:\Users\qiujinbo\Desktop\海外公用包材报表\项目编码前缀.xlsx';
  leMDemand.Text := 'C:\Users\qiujinbo\Desktop\海外公用包材报表\市场预测.xls';
  lePkgStuff.Text := 'C:\Users\qiujinbo\Desktop\海外公用包材报表\国内&海外共用包材.xlsx';
  FMDemands := TList.Create;
  FProjNumber := TStringList.Create;
  FPkgStuff := TStringList.Create;
  FProjMDemand := TStringList.Create;
end;
   
procedure TfrmHWPkgStuff.FormDestroy(Sender: TObject);
begin
  Clear;
  FMDemands.Free;
  FProjNumber.Free;
  FPkgStuff.Free;
  FProjMDemand.Free;
end;
  
procedure TfrmHWPkgStuff.btnProjNumberClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options - [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  leProjNumber.Text := OpenDialog1.FileName;
end;

procedure TfrmHWPkgStuff.btnMDemandClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options - [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  leMDemand.Text := OpenDialog1.FileName;
end;

procedure TfrmHWPkgStuff.btnPkgStuffClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options - [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  lePkgStuff.Text := OpenDialog1.FileName;
end;

function GetCol(irow: Integer; ExcelApp: Variant; const scol: string): Integer;
var
  icol: Integer;
  svalue: string;
begin
  Result := 0;
  for icol := 1 to 100 do
  begin
    svalue := ExcelApp.Cells[irow, icol].Value;
    if svalue = scol then
    begin
      Result := icol;
      Break;
    end;
  end;
  if Result = 0 then
  begin
    raise Exception.Create('列 "' + scol + '" 不存在');
  end;
end;

procedure TfrmHWPkgStuff.SumMDemand;
var
  aMDemandPtr: PMDemand;
  aProjDemand: TProjDemand;
  i: Integer;
  idx: Integer;
  sprojNumber: string;
  sprojName: string;
  aProjDayDemand: PProjDayDemand;
begin
  for i := 0 to FMDemands.Count - 1 do
  begin
    aMDemandPtr := FMDemands[i];
    sprojNumber := Copy(aMDemandPtr^.Number, 1, 5);

    idx := FProjNumber.IndexOfName(sprojNumber);
    if idx < 0 then
    begin
      Memo1.Lines.Add('项目编号名称不存在 ' + aMDemandPtr^.Number);
      Continue;
    end;
    sprojName := FProjNumber.ValueFromIndex[idx];


    idx := FProjMDemand.IndexOf(sprojName);
    if idx < 0 then
    begin
      aProjDemand := TProjDemand.Create;
      FProjMDemand.AddObject(sprojName, aProjDemand);
    end
    else
    begin
      aProjDemand := TProjDemand(FProjMDemand.Objects[idx]);
    end;

    idx := aProjDemand.IndexOf(aMDemandPtr^.Date1);
    if idx < 0 then
    begin
      aProjDayDemand := New(PProjDayDemand);
      aProjDayDemand^.Date1 := aMDemandPtr^.Date1;
      aProjDayDemand^.Date2 := aMDemandPtr^.Date2;
      aProjDayDemand^.Qty := aMDemandPtr^.Qty;
      aProjDayDemand^.DemandType := aMDemandPtr^.DemandType;
      aProjDayDemand^.sUnit := aMDemandPtr^.sUnit;

      aProjDayDemand^.SourceType := aMDemandPtr^.SourceType;
      aProjDayDemand^.SourceBillNo := aMDemandPtr^.SourceBillNo;
      aProjDayDemand^.SourceEntryID := aMDemandPtr^.SourceEntryID;
      aProjDayDemand^.Note := aMDemandPtr^.Note;

      aProjDemand.FList.Add(aProjDayDemand);
    end
    else
    begin
      aProjDayDemand := PProjDayDemand(aProjDemand.FList[idx]);
      aProjDayDemand^.Qty := aProjDayDemand^.Qty + aMDemandPtr^.Qty;
    end;
  end;
end;

function TfrmHWPkgStuff.GetFGDemand(const sdate: string; slFGNumber: TStringList;
  var aProjDayDemand: TProjDayDemand): Boolean;
var
  iDemand: Integer;
  aMDemandPtr: PMDemand;
  dt: TDateTime;
begin
  Result := False;                                          
  dt := myStrToDateTime(sdate);
  for iDemand := 0 to FMDemands.Count - 1 do
  begin
    aMDemandPtr := PMDemand(FMDemands[iDemand]);

    if aMDemandPtr^.Date1 <> dt then Continue;
    if slFGNumber.IndexOf(aMDemandPtr^.Number) < 0 then Continue;
    if not Result then
    begin
      aProjDayDemand.Date1 := aMDemandPtr^.Date1;
      aProjDayDemand.Date2 := aMDemandPtr^.Date2;
      aProjDayDemand.Qty := aMDemandPtr^.Qty;
      aProjDayDemand.DemandType := aMDemandPtr^.DemandType;
      aProjDayDemand.sUnit := aMDemandPtr^.sUnit;
      aProjDayDemand.SourceType := aMDemandPtr^.SourceType;
      aProjDayDemand.SourceBillNo := aMDemandPtr^.SourceBillNo;
      aProjDayDemand.SourceEntryID := aMDemandPtr^.SourceEntryID;
      aProjDayDemand.Note := aMDemandPtr^.Note;
 
      Result := True;
    end
    else
    begin
      aProjDayDemand.Qty := aProjDayDemand.Qty + aMDemandPtr^.Qty;
    end;
  end;
end;

procedure TfrmHWPkgStuff.MakePkgStuffDemand(const sfile: string);
const
  CSBillNo = '编号*';
  CSDemandType = '需求类型*';
  CSNumber = '物料长编码*';
  CSUnit = '单位*';
  CSQty = '数量*';
  CSDate1 = '预测开始日期*';
  CSDate2 = '预测截止日期*';
  CSAvgType = '均化周期类型*';
  CSSourceType = '源单类型*';
  CSSourceBillNo = '源单号*';
  CSSourceEntryID =	'源单行号*';
  CSNote = '备注*';

var
  iProj: Integer;
  iDate: Integer;
  iPkg: Integer;
  idx: Integer;
  sprojName: string;
  aProjDemand: TProjDemand;
  aPkgStuff: TPkgStuff;
  aProjDayDemandPtr: PProjDayDemand;
  irow: Integer;
  ExcelApp, WorkBook: Variant;
  aPkgItem: TPkgItem;
  slFGNumber: TStringList;
  iDemand: Integer;
  aMDemandPtr: PMDemand;
  sldate: TStringList;           
  sdate: string;
  aProjDayDemand: TProjDayDemand;
begin

  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  slFGNumber := TStringList.Create;

  WorkBook := ExcelApp.WorkBooks.Add;
  WorkBook.Sheets[1].Name := '产品预测单';
  
  try
    irow := 1;
    ExcelApp.Cells[irow, 1].Value := CSBillNo;
    ExcelApp.Cells[irow, 2].Value := CSDemandType;
    ExcelApp.Cells[irow, 3].Value := CSNumber;
    ExcelApp.Cells[irow, 4].Value := CSUnit;
    ExcelApp.Cells[irow, 5].Value := CSQty;
    ExcelApp.Cells[irow, 6].Value := CSDate1;;
    ExcelApp.Cells[irow, 7].Value := CSDate2;    
    ExcelApp.Cells[irow, 8].Value := CSAvgType;
    ExcelApp.Cells[irow, 9].Value := CSSourceType;
    ExcelApp.Cells[irow, 10].Value := CSSourceBillNo;
    ExcelApp.Cells[irow, 11].Value := CSSourceEntryID;
    ExcelApp.Cells[irow, 12].Value := CSNote;

    ExcelApp.Columns[1].ColumnWidth := 13;   
    ExcelApp.Columns[2].ColumnWidth := 9;   
    ExcelApp.Columns[3].ColumnWidth := 15;  
    ExcelApp.Columns[4].ColumnWidth := 6;
    ExcelApp.Columns[5].ColumnWidth := 7;
    ExcelApp.Columns[6].ColumnWidth := 13;
    ExcelApp.Columns[7].ColumnWidth := 13; 
    ExcelApp.Columns[8].ColumnWidth := 13;
    ExcelApp.Columns[9].ColumnWidth := 8;
    ExcelApp.Columns[10].ColumnWidth := 8;
    ExcelApp.Columns[11].ColumnWidth := 8;
    ExcelApp.Columns[12].ColumnWidth := 34;

    irow := irow + 1;
    for iProj := 0 to FProjMDemand.Count - 1 do
    begin
      sprojName := FProjMDemand[iProj];

      idx := FPkgStuff.IndexOf(sprojName);
      if idx < 0 then
      begin
        Memo1.Lines.Add('没有包材清单 ' + sprojName);
        Continue;
      end;

      aPkgStuff := TPkgStuff(FPkgStuff.Objects[idx]);
      aProjDemand := TProjDemand(FProjMDemand.Objects[iProj]);

      for iDate := 0 to aProjDemand.FList.Count - 1 do
      begin
        aProjDayDemandPtr := PProjDayDemand(aProjDemand.FList[iDate]);

        for iPkg := 0 to aPkgStuff.FItems.Count - 1 do
        begin
          aPkgItem := TPkgItem(aPkgStuff.FItems[iPkg]);
          if aPkgItem.FFGNumber <> '' then Continue;

          ExcelApp.Cells[irow, 1].Value := 'HWYC0000XX';
          ExcelApp.Cells[irow, 2].Value := aProjDayDemandPtr^.DemandType;
          ExcelApp.Cells[irow, 3].Value := aPkgItem.F99;
          ExcelApp.Cells[irow, 4].Value := aProjDayDemandPtr^.sUnit;
          ExcelApp.Cells[irow, 5].Value := Ceil(aProjDayDemandPtr^.Qty * aPkgItem.FUsage); //向上取整
          ExcelApp.Cells[irow, 6].Value := aProjDayDemandPtr^.Date1;
          ExcelApp.Cells[irow, 7].Value := aProjDayDemandPtr^.Date2;
          ExcelApp.Cells[irow, 8].Value := '不均化';
          ExcelApp.Cells[irow, 9].Value := aProjDayDemandPtr^.SourceType;
          ExcelApp.Cells[irow, 10].Value := aProjDayDemandPtr^.SourceBillNo;
          ExcelApp.Cells[irow, 11].Value := aProjDayDemandPtr^.SourceEntryID;
          ExcelApp.Cells[irow, 12].Value := aProjDayDemandPtr^.Note;
            
          irow := irow + 1;
        end;

      end;
    end;

    ///////// 分颜色的包材 //////////////////////////////
    sldate := TStringList.Create;
    try
      for iDemand := 0 to FMDemands.Count - 1 do
      begin
        aMDemandPtr := PMDemand(FMDemands[iDemand]);
        sdate := FormatDateTime('yyyy-MM-dd', aMDemandPtr^.Date1);
        if sldate.IndexOf(sdate) >= 0 then Continue;
        sldate.Add(sdate);
      end;

      for iProj := 0 to FPkgStuff.Count - 1 do
      begin
        aPkgStuff := TPkgStuff(FPkgStuff.Objects[iProj]);
        for iPkg := 0 to aPkgStuff.FItems.Count - 1 do
        begin
          aPkgItem := TPkgItem(aPkgStuff.FItems[iPkg]);
          if aPkgItem.FFGNumber = '' then Continue;

          slFGNumber.Text := StringReplace(aPkgItem.FFGNumber, ';', #13#10, [rfReplaceAll]);

          for iDate := 0 to sldate.Count - 1 do
          begin
            if not GetFGDemand(sldate[iDate], slFGNumber, aProjDayDemand) then Continue;


            ExcelApp.Cells[irow, 1].Value := 'HWYC0000XX';
            ExcelApp.Cells[irow, 2].Value := aProjDayDemand.DemandType;
            ExcelApp.Cells[irow, 3].Value := aPkgItem.F99;
            ExcelApp.Cells[irow, 4].Value := aProjDayDemand.sUnit;
            ExcelApp.Cells[irow, 5].Value := Ceil(aProjDayDemand.Qty * aPkgItem.FUsage); //向上取整
            ExcelApp.Cells[irow, 6].Value := aProjDayDemand.Date1;
            ExcelApp.Cells[irow, 7].Value := aProjDayDemand.Date2;
            ExcelApp.Cells[irow, 8].Value := '不均化';
            ExcelApp.Cells[irow, 9].Value := aProjDayDemand.SourceType;
            ExcelApp.Cells[irow, 10].Value := aProjDayDemand.SourceBillNo;
            ExcelApp.Cells[irow, 11].Value := aProjDayDemand.SourceEntryID;
            ExcelApp.Cells[irow, 12].Value := aProjDayDemand.Note;      
            irow := irow + 1;

          end; 
        end;
      end;
    finally
      sldate.Free;
    end;


    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 12]].Borders.LineStyle := 1; //加边框

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
 
       
  finally
    WorkBook.Close;
    ExcelApp.Quit;

    slFGNumber.Free;
  end;
  
end;
 

procedure TfrmHWPkgStuff.OpenMDemand2(const sfile: string);
const
  CSNumber = '物料长编码*';
  CSDemandType = '需求类型*';
  CSQty = '数量*';
  CSDate1 = '预测开始日期*';
  CSDate2 = '预测截止日期*';
  CSUnit = '单位*';
  
var
  irow: Integer;
  iNumber: Integer; 
  iDemandType: Integer;
  iQty: Integer;
  iDate1: Integer;
  iDate2: Integer;
  iUnit: Integer; 
  snumber: string;
  sName: string;
  sDemandType: string;
  sUnit: string;
  sQty: Integer; 
  dtDate1: TDateTime;
  dtDate2: TDateTime;
  ExcelApp, WorkBook: Variant;
  aMDemandPtr: PMDemand;
begin
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    try
      irow := 1;
      iNumber := GetCol(irow, ExcelApp, CSNumber);
      iDemandType := GetCol(irow, ExcelApp, CSDemandType);
      iQty := GetCol(irow, ExcelApp, CSQty);
      iDate1 := GetCol(irow, ExcelApp, CSDate1);
      iDate2 := GetCol(irow, ExcelApp, CSDate2);    
      iUnit := GetCol(irow, ExcelApp, CSUnit);
 
      irow := 2;
      snumber := ExcelApp.Cells[irow, iNumber].Value;
      while snumber <> '' do
      begin
        sname := '';
        sdemandtype := ExcelApp.Cells[irow, iDemandType].Value;
        sqty := ExcelApp.Cells[irow, iQty].Value;
        dtDate1 := ExcelApp.Cells[irow, iDate1].Value;
        dtDate2 := ExcelApp.Cells[irow, iDate2].Value;    
        sUnit := ExcelApp.Cells[irow, iUnit].Value;
    
        aMDemandPtr := New(PMDemand);
        aMDemandPtr.Number := snumber;
        aMDemandPtr.Name := sname;
        aMDemandPtr.DemandType := sdemandtype;
        aMDemandPtr.sUnit := sUnit;
        aMDemandPtr.Qty := sqty;
        aMDemandPtr.Date1 := dtdate1;
        aMDemandPtr.Date2 := dtdate2;
        FMDemands.Add(aMDemandPtr);

        irow := irow + 1;
        snumber := ExcelApp.Cells[irow, iNumber].Value;
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

procedure TfrmHWPkgStuff.OpenProjNumber(const sfile: string);
const
  CSNumber = '项目编码';
  CSName = '项目名称';
var
  irow: Integer;
  iNumber: Integer;
  iName: Integer;
  snumber: string; 
  sName: string; 
  ExcelApp, WorkBook: Variant;
  idx: Integer; 
begin
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    try
      irow := 1;
      iNumber := GetCol(irow, ExcelApp, CSNumber);  
      iName := GetCol(irow, ExcelApp, CSName); 

      irow := 2;
      snumber := ExcelApp.Cells[irow, iNumber].Value;
      while snumber <> '' do
      begin
        sname := ExcelApp.Cells[irow, iName].Value;
        if Copy(snumber, 1, 1) = '''' then
          snumber := Copy(snumber, 2, Length(snumber) - 1);
        idx := FProjNumber.IndexOfName(snumber);
        if idx < 0 then
          FProjNumber.Add(snumber + '=' + sName);

        irow := irow + 1;
        snumber := ExcelApp.Cells[irow, iNumber].Value;
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

function StringListIndexOfValue(sl: TStringList; const s: string): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to sl.Count - 1 do
  begin
    if sl.ValueFromIndex[i] = s then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TfrmHWPkgStuff.OpenPkgStuff(const sfile: string);
const
  CSNumber = '物料编码';
  CSName = '物料名称';
  CSUsage = '用量';
  CS99 = '99编码';  
  CSFGNumber = '产品编码';
var
  irow: Integer;
  irow0: Integer;
  iNumber: Integer;
  iName: Integer;
  iUsage: Integer;
  i99: Integer;
  iFGNumber: Integer;
  snumber: string;
  sName: string;
  sUsage: Double;
  s99: string;
  sFGNumber: string;
  ExcelApp, WorkBook: Variant;   
  iSheetCount, iSheet: Integer;
  aPkgStuff: TPkgStuff;
  aPkgItem: TPkgItem;
  idx: Integer;
  sproj: string;
begin 
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    try             
      iSheetCount := ExcelApp.WorkSheets.Count;
      for iSheet := 1 to iSheetCount do
      begin         
        ExcelApp.WorkSheets[iSheet].Activate;

        sproj := ExcelApp.WorkSheets[iSheet].Name;
        idx := StringListIndexOfValue(FProjNumber, sproj);
        if idx < 0 then
        begin
          memo1.Lines.Add(sproj + '不存在');
          Continue;
        end;

        idx := FPkgStuff.IndexOf(sproj);
        if idx <= 0 then
        begin
          aPkgStuff := TPkgStuff.Create;
          FPkgStuff.AddObject(sproj, aPkgStuff);
        end
        else
        begin
          aPkgStuff := TPkgStuff(FPkgStuff[idx]);
        end;

        irow := 1;
        iNumber := GetCol(irow, ExcelApp, CSNumber);  
        iName := GetCol(irow, ExcelApp, CSName);
        iUsage := GetCol(irow, ExcelApp, CSUsage);
        i99 := GetCol(irow, ExcelApp, CS99);
        iFGNumber := GetCol(irow, ExcelApp, CSFGNumber);
 
        irow0 := 0;
        aPkgItem := nil;
        
        irow := 2;
        snumber := ExcelApp.Cells[irow, iNumber].Value;
        while snumber <> '' do
        begin   
          sname := ExcelApp.Cells[irow, iName].Value;
          sUsage := ExcelApp.Cells[irow, iUsage].Value;
          s99 := ExcelApp.Cells[irow, i99].Value;
          sFGNumber := ExcelApp.Cells[irow, iFGNumber].Value;

          if irow0 = 0 then
          begin
            aPkgItem := TPkgItem.Create;
            aPkgItem.FNumbers.Add(snumber + '=' + sName);
            aPkgItem.FUsage := sUsage;
            aPkgItem.F99 := s99;
            aPkgItem.FFGNumber := sFGNumber;
            aPkgStuff.FItems.Add(aPkgItem);

            irow0 := irow;
          end
          else
          begin
            if CellMerged(ExcelApp, irow0, irow, iUsage) then
            begin
              aPkgItem.FNumbers.Add(snumber + '=' + sName);
            end
            else
            begin
              aPkgItem := TPkgItem.Create;
              aPkgItem.FNumbers.Add(snumber + '=' + sName);
              aPkgItem.FUsage := sUsage;     
              aPkgItem.F99 := s99;
              aPkgItem.FFGNumber := sFGNumber;
              aPkgStuff.FItems.Add(aPkgItem);

              irow0 := irow;
            end;
          end;

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, iNumber].Value;
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

procedure TfrmHWPkgStuff.ToolButton1Click(Sender: TObject);
var
  sfile: string;
begin
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := 'HWPKG' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;
  sfile := SaveDialog1.FileName;

  Clear;                    
  OpenProjNumber(leProjNumber.Text); 
  OpenMDemand2(leMDemand.Text);
  OpenPkgStuff(lePkgStuff.Text);

  SumMDemand;

  MakePkgStuffDemand(sfile);
  
  MessageBox(Handle, '完成', '金蝶提示', 0);
end;

procedure TfrmHWPkgStuff.Clear;
var
  aPkgStuff: TPkgStuff;
  aProjDemand: TProjDemand;
  p: PMDemand;
  i: Integer;
begin
  for i := 0 to FPkgStuff.Count - 1 do
  begin
    aPkgStuff := TPkgStuff(FPkgStuff.Objects[i]);
    aPkgStuff.Free;
  end;
  FPkgStuff.Clear;

  for i := 0 to FMDemands.Count - 1 do
  begin
    p := PMDemand(FMDemands[i]);
    Dispose(p);
  end;
  FMDemands.Clear;

  for i := 0 to FProjMDemand.Count - 1 do
  begin
    aProjDemand := TProjDemand(FProjMDemand.Objects[i]);
    aProjDemand.Free;
  end;
  FProjMDemand.Clear;

  FProjNumber.Clear;
end;

end.
