unit MakeFGReportWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, FGWinReader, FGStockRptReader, FGUnOutRptReader, IniFiles,
  ComCtrls, ToolWin, ImgList, ProjNameNoWin, ExtCtrls, CommUtils,
  ComObj, MakeFGReportCommon, FGRptWinReader;

type
  TfrmMakeFGReport = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ToolButton1: TToolButton;
    tbProjNameNo: TToolButton;
    mmoWins: TMemo;
    mmoStocks: TMemo;
    mmoUnOut: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnWins: TButton;
    btnStocks: TButton;
    btnUnOut: TButton;
    btnSF: TButton;
    btnKJ: TButton;
    btnORT: TButton;
    mmoDB: TMemo;
    Label4: TLabel;
    btnDB: TButton;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    mmoWinsBS: TMemo;
    mmoStocksBS: TMemo;
    mmoUnOutBS: TMemo;
    btnWinsBS: TButton;
    btnStocksBS: TButton;
    btnUnOutBS: TButton;
    mmoDBBS: TMemo;
    btnDBBS: TButton;
    leFGRptWin: TLabeledEdit;
    btnFGRptWin: TButton;
    Memo1: TMemo;
    Memo2: TMemo;
    Label9: TLabel;
    mmoSF: TMemo;
    Label10: TLabel;
    mmoKJ: TMemo;
    mmoORT: TMemo;
    Label11: TLabel;
    procedure tbProjNameNoClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnWinsClick(Sender: TObject);
    procedure btnStocksClick(Sender: TObject);
    procedure btnUnOutClick(Sender: TObject);
    procedure btnSFClick(Sender: TObject);
    procedure btnKJClick(Sender: TObject);
    procedure btnORTClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnDBClick(Sender: TObject);
    procedure btnWinsBSClick(Sender: TObject);
    procedure btnStocksBSClick(Sender: TObject);
    procedure btnUnOutBSClick(Sender: TObject);
    procedure btnDBBSClick(Sender: TObject);
    procedure btnFGRptWinClick(Sender: TObject);
  private
    { Private declarations }
    procedure Log(const s: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmMakeFGReport.ShowForm;
var
  frmMakeFGReport: TfrmMakeFGReport;
begin
  frmMakeFGReport := TfrmMakeFGReport.Create(nil);
  try
    frmMakeFGReport.ShowModal;
  finally
    frmMakeFGReport.Free;
  end;
end;
    
procedure TfrmMakeFGReport.tbProjNameNoClick(Sender: TObject);
begin
  TfrmProjNameNo.ShowForm;
end;

procedure TfrmMakeFGReport.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMakeFGReport.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, mmoWins.Name, '');
    mmoWins.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmoStocks.Name, '');
    mmoStocks.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmoUnOut.Name, '');
    mmoUnOut.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmoDB.Name, '');
    mmoDB.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
                                
    s := ini.ReadString(self.ClassName, mmoWinsBS.Name, '');
    mmoWinsBS.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmoStocksBS.Name, '');
    mmoStocksBS.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmoUnOutBS.Name, '');
    mmoUnOutBS.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmoDBBS.Name, '');
    mmoDBBS.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);

    s := ini.ReadString(self.ClassName, mmoSF.Name, '');
    mmoSF.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
                                                         
    s := ini.ReadString(self.ClassName, mmoKJ.Name, '');
    mmoKJ.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
                                                         
    s := ini.ReadString(self.ClassName, mmoORT.Name, '');
    mmoORT.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);

    leFGRptWin.Text := ini.ReadString(self.ClassName, leFGRptWin.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmMakeFGReport.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := StringReplace(mmoWins.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoWins.Name, s); 
    s := StringReplace(mmoStocks.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoStocks.Name, s);
    s := StringReplace(mmoUnOut.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoUnOut.Name, s);
    s := StringReplace(mmoDB.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoDB.Name, s);
                   
    s := StringReplace(mmoWinsBS.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoWinsBS.Name, s);
    s := StringReplace(mmoStocksBS.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoStocksBS.Name, s);
    s := StringReplace(mmoUnOutBS.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoUnOutBS.Name, s);
    s := StringReplace(mmoDBBS.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoDBBS.Name, s);

    s := StringReplace(mmoSF.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoSF.Name, s);
    
    s := StringReplace(mmoKJ.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoKJ.Name, s);

    s := StringReplace(mmoORT.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoORT.Name, s);

    ini.WriteString(self.ClassName, leFGRptWin.Name, leFGRptWin.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmMakeFGReport.btnWinsClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoWins.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnStocksClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoStocks.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnUnOutClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoUnOut.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;
     
procedure TfrmMakeFGReport.btnDBClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoDB.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnSFClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoSF.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnKJClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoKJ.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnORTClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoORT.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;
    
function NameNeedSum(const sName: string; slIgnoreNames: TStringList): Boolean;
var
  sl: TStringList;
  i: Integer;
  j: Integer;
  bFound: Boolean;
  s1: string;
  s2: string;
begin
  Result := True;
  sl := TStringList.Create;
  for i := 0 to slIgnoreNames.Count - 1 do
  begin
    s1 := Trim(slIgnoreNames[i]);
    if s1 = '' then Continue;

    sl.Text := StringReplace(s1, ' ', #13#10, [rfReplaceAll]);
    bFound := True;
    for j := 0 to sl.Count - 1 do
    begin
      s2 := Trim(sl[j]);
      if s2 = '' then Continue;
      if Pos(s2, sName) = 0 then
      begin
        bFound := False;
        Break;
      end;
    end;

    if bFound then
    begin
      Result := False;
      Break;
    end;
  end;
  sl.Free;
end;

procedure SaveWin(ExcelApp: Variant; aFGRptWinReader: TFGRptWinReader;
  slWin, slWinBS: TStringList);
var
  aFGWinReader: TFGWinReader;
  irow: Integer;
  i, j: Integer;    
  iino: Integer;
  bino: Boolean;  
  slIgnoreNos: TStringList;     
  slIgnoreNames: TStringList;
begin
  irow := 1;
                     
  slIgnoreNos := TfrmProjNameNo.GetIgnoreNos;  
  slIgnoreNames := TfrmProjNameNo.GetIgnoreName4Sum;

  ExcelApp.Cells[irow, 1].Value := '入库日期';
  ExcelApp.Cells[irow, 2].Value := '料号';
  ExcelApp.Cells[irow, 3].Value := '产品名称';
  ExcelApp.Cells[irow, 4].Value := '数量';
  ExcelApp.Cells[irow, 5].Value := '代工厂';
  ExcelApp.Cells[irow, 6].Value := '批次';     
  ExcelApp.Cells[irow, 7].Value := '备注';
  ExcelApp.Cells[irow, 8].Value := '汇总';
     
  ExcelApp.Columns[1].ColumnWidth := 12;
  ExcelApp.Columns[2].ColumnWidth := 17;
  ExcelApp.Columns[3].ColumnWidth := 34;
  ExcelApp.Columns[4].ColumnWidth := 9;
  ExcelApp.Columns[5].ColumnWidth := 13;
  ExcelApp.Columns[6].ColumnWidth := 13;
  ExcelApp.Columns[7].ColumnWidth := 17;

  irow := 2;

  // 先写之前的入库
  for j := 0 to aFGRptWinReader.Count - 1 do
  begin       
    aFGRptWinReader.Items[j].bSum := NameNeedSum(aFGRptWinReader.Items[j].sName, slIgnoreNames);

    ExcelApp.Cells[irow, 1].Value := aFGRptWinReader.Items[j].dt;
    ExcelApp.Cells[irow, 2].Value := aFGRptWinReader.Items[j].snumber;
    ExcelApp.Cells[irow, 3].Value := aFGRptWinReader.Items[j].sname;
    ExcelApp.Cells[irow, 4].Value := aFGRptWinReader.Items[j].dqty;
    ExcelApp.Cells[irow, 5].Value := aFGRptWinReader.Items[j].sfac;
    ExcelApp.Cells[irow, 6].Value := aFGRptWinReader.Items[j].sBatchNo;
    ExcelApp.Cells[irow, 7].Value := aFGRptWinReader.Items[j].snote;
    if aFGRptWinReader.Items[j].bSum then   
      ExcelApp.Cells[irow, 8].Value := 'Y';
    irow := irow + 1;
  end;


  for i := 0 to slWin.Count - 1 do
  begin
    aFGWinReader := TFGWinReader(slWin.Objects[i]);
    for j := 0 to aFGWinReader.Count - 1 do
    begin                                                 


      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGWinReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGWinReader.Items[j].bSum := NameNeedSum(aFGWinReader.Items[j].sName, slIgnoreNames);

      if aFGWinReader.Items[j].snote = '返工入库' then Continue;
      ExcelApp.Cells[irow, 1].Value := aFGWinReader.Items[j].dt;
      ExcelApp.Cells[irow, 2].Value := aFGWinReader.Items[j].snumber;
      ExcelApp.Cells[irow, 3].Value := aFGWinReader.Items[j].sname;
      ExcelApp.Cells[irow, 4].Value := aFGWinReader.Items[j].dqty;
      ExcelApp.Cells[irow, 5].Value := aFGWinReader.Items[j].sfac;
      ExcelApp.Cells[irow, 6].Value := aFGWinReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 7].Value := aFGWinReader.Items[j].snote;
      if aFGWinReader.Items[j].bSum then
        ExcelApp.Cells[irow, 8].Value := 'Y';
      irow := irow + 1;
    end;
  end;

  for i := 0 to slWinBS.Count - 1 do
  begin
    aFGWinReader := TFGWinReader(slWinBS.Objects[i]);
    for j := 0 to aFGWinReader.Count - 1 do
    begin                                                 


      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGWinReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGWinReader.Items[j].bSum := NameNeedSum(aFGWinReader.Items[j].sName, slIgnoreNames);

      if aFGWinReader.Items[j].snote = '返工入库' then Continue;
      ExcelApp.Cells[irow, 1].Value := aFGWinReader.Items[j].dt;
      ExcelApp.Cells[irow, 2].Value := aFGWinReader.Items[j].snumber;
      ExcelApp.Cells[irow, 3].Value := aFGWinReader.Items[j].sname;
      ExcelApp.Cells[irow, 4].Value := aFGWinReader.Items[j].dqty;
      ExcelApp.Cells[irow, 5].Value := aFGWinReader.Items[j].sfac;
      ExcelApp.Cells[irow, 6].Value := aFGWinReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 7].Value := aFGWinReader.Items[j].snote;
      if aFGWinReader.Items[j].bSum then
        ExcelApp.Cells[irow, 8].Value := 'Y';
      irow := irow + 1;
    end;
  end;
       
  slIgnoreNos.Free;   
  slIgnoreNames.Free;
 
  AddBorder(ExcelApp, 1, 1, irow - 1, 8);
  AddColor(ExcelApp, 1, 1, 1, 8, $DBDCF2);
  AddHorizontalAlignment(ExcelApp, 1, 1, 1, 8, xlCenter);

  ExcelApp.Range[ExcelApp.Cells[2, 4], ExcelApp.Cells[irow - 1, 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
      
end;
             
procedure SaveStock(ExcelApp: Variant; slStock, slStockBS, slDB, slDBBS: TStringList);
var
  aFGStockRptReader: TFGStockRptReader;
  irow: Integer;
  i, j: Integer;   
  iino: Integer;
  bino: Boolean;     
  slIgnoreNos: TStringList;     
  slIgnoreNames: TStringList;
begin
  irow := 1;
                                             
  slIgnoreNos := TfrmProjNameNo.GetIgnoreNos;    
  slIgnoreNames := TfrmProjNameNo.GetIgnoreName4Sum;

  ExcelApp.Cells[irow, 1].Value := '料号';
  ExcelApp.Cells[irow, 2].Value := '产品名称';
  ExcelApp.Cells[irow, 3].Value := '库存总数';
  ExcelApp.Cells[irow, 4].Value := '其中';
  ExcelApp.Cells[irow + 1, 4].Value := '返工';
  ExcelApp.Cells[irow + 1, 5].Value := '待检验'#13#10'（超过3个月仓龄）';
  MergeCells(ExcelApp, 1, 4, 1, 5);
  ExcelApp.Cells[irow, 6].Value := '存货地点';
  ExcelApp.Cells[irow, 7].Value := '批次';
  ExcelApp.Cells[irow, 8].Value := '备注';    
  ExcelApp.Cells[irow, 9].Value := '汇总';
       
  ExcelApp.Columns[1].ColumnWidth := 17;
  ExcelApp.Columns[2].ColumnWidth := 21;
  ExcelApp.Columns[3].ColumnWidth := 9;
  ExcelApp.Columns[4].ColumnWidth := 9;  
  ExcelApp.Columns[5].ColumnWidth := 12;
  ExcelApp.Columns[6].ColumnWidth := 13;
  ExcelApp.Columns[7].ColumnWidth := 13;
  ExcelApp.Columns[8].ColumnWidth := 8;
                           
  MergeCells(ExcelApp, 1, 1, 2, 1);
  MergeCells(ExcelApp, 1, 2, 2, 2);
  MergeCells(ExcelApp, 1, 3, 2, 3);
  MergeCells(ExcelApp, 1, 6, 2, 6);
  MergeCells(ExcelApp, 1, 7, 2, 7);
  MergeCells(ExcelApp, 1, 8, 2, 8);   
  MergeCells(ExcelApp, 1, 9, 2, 9);

  irow := 3;

  for i := 0 to slStock.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin

      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGStockRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGStockRptReader.Items[j].bSum := NameNeedSum(aFGStockRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGStockRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGStockRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGStockRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGStockRptReader.Items[j].drework;
      ExcelApp.Cells[irow, 5].Value := aFGStockRptReader.Items[j].duncheck;
      ExcelApp.Cells[irow, 6].Value := aFGStockRptReader.Items[j].saddr;
      if Trim(aFGStockRptReader.Items[j].sBatchNo) = '' then
        ExcelApp.Cells[irow, 7].Value := '   '
      else
        ExcelApp.Cells[irow, 7].Value := aFGStockRptReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 8].Value := aFGStockRptReader.Items[j].snote;
      if aFGStockRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 9].Value := 'Y';
      irow := irow + 1;
    end;
  end;
      
  for i := 0 to slDB.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin      

      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGStockRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGStockRptReader.Items[j].bSum := NameNeedSum(aFGStockRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGStockRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGStockRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGStockRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGStockRptReader.Items[j].drework;
      ExcelApp.Cells[irow, 5].Value := aFGStockRptReader.Items[j].duncheck;
      ExcelApp.Cells[irow, 6].Value := aFGStockRptReader.Items[j].saddr;
      if Trim(aFGStockRptReader.Items[j].sBatchNo) = '' then
        ExcelApp.Cells[irow, 7].Value := '   '
      else
        ExcelApp.Cells[irow, 7].Value := aFGStockRptReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 8].Value := aFGStockRptReader.Items[j].snote;
      if aFGStockRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 9].Value := 'Y';
      irow := irow + 1;
    end;
  end;
         
  for i := 0 to slStockBS.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStockBS.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin    

      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGStockRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGStockRptReader.Items[j].bSum := NameNeedSum(aFGStockRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGStockRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGStockRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGStockRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGStockRptReader.Items[j].drework;
      ExcelApp.Cells[irow, 5].Value := aFGStockRptReader.Items[j].duncheck;
      ExcelApp.Cells[irow, 6].Value := aFGStockRptReader.Items[j].saddr;
      if Trim(aFGStockRptReader.Items[j].sBatchNo) = '' then
        ExcelApp.Cells[irow, 7].Value := '   '
      else
        ExcelApp.Cells[irow, 7].Value := aFGStockRptReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 8].Value := aFGStockRptReader.Items[j].snote;
      if aFGStockRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 9].Value := 'Y';
      irow := irow + 1;
    end;
  end;
      
  for i := 0 to slDBBS.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDBBS.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin       

      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGStockRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGStockRptReader.Items[j].bSum := NameNeedSum(aFGStockRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGStockRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGStockRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGStockRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGStockRptReader.Items[j].drework;
      ExcelApp.Cells[irow, 5].Value := aFGStockRptReader.Items[j].duncheck;
      ExcelApp.Cells[irow, 6].Value := aFGStockRptReader.Items[j].saddr;
      if Trim(aFGStockRptReader.Items[j].sBatchNo) = '' then
        ExcelApp.Cells[irow, 7].Value := '   '
      else
        ExcelApp.Cells[irow, 7].Value := aFGStockRptReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 8].Value := aFGStockRptReader.Items[j].snote;
      if aFGStockRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 9].Value := 'Y';
      irow := irow + 1;
    end;
  end;
           
  slIgnoreNos.Free;           
  slIgnoreNames.Free;
 
  AddBorder(ExcelApp, 1, 1, irow - 1, 9);
  AddColor(ExcelApp, 1, 1, 2, 9, $DEF1EB);   
  AddHorizontalAlignment(ExcelApp, 1, 1, 2, 9, xlCenter);
  
  ExcelApp.Range[ExcelApp.Cells[3, 3], ExcelApp.Cells[irow - 1, 5]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
      
end;

procedure SaveUnOut(ExcelApp: Variant; slUnOut, slUnOutBS: TStringList);
var
  aFGUnOutRptReader: TFGUnOutRptReader;
  irow: Integer;
  i, j: Integer;    
  iino: Integer;
  bino: Boolean;     
  slIgnoreNos: TStringList;    
  slIgnoreNames: TStringList;
begin
  irow := 1;
                             
  slIgnoreNos := TfrmProjNameNo.GetIgnoreNos;  
  slIgnoreNames := TfrmProjNameNo.GetIgnoreName4Sum;

  ExcelApp.Cells[irow, 1].Value := '料号';
  ExcelApp.Cells[irow, 2].Value := '产品名称';
  ExcelApp.Cells[irow, 3].Value := '数量';
  ExcelApp.Cells[irow, 4].Value := '备注';    
  ExcelApp.Cells[irow, 5].Value := '汇总';

  ExcelApp.Columns[1].ColumnWidth := 20;
  ExcelApp.Columns[2].ColumnWidth := 40;
  ExcelApp.Columns[3].ColumnWidth := 13;
  ExcelApp.Columns[4].ColumnWidth := 14;

  irow := 2;

  for i := 0 to slUnOut.Count - 1 do
  begin
    aFGUnOutRptReader := TFGUnOutRptReader(slUnOut.Objects[i]);
    for j := 0 to aFGUnOutRptReader.Count - 1 do
    begin             

      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGUnOutRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGUnOutRptReader.Items[j].bSum := NameNeedSum(aFGUnOutRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGUnOutRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGUnOutRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGUnOutRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGUnOutRptReader.Items[j].snote;
      if aFGUnOutRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 5].Value := 'Y';
      irow := irow + 1;
    end;
  end;

  for i := 0 to slUnOutBS.Count - 1 do
  begin
    aFGUnOutRptReader := TFGUnOutRptReader(slUnOutBS.Objects[i]);
    for j := 0 to aFGUnOutRptReader.Count - 1 do
    begin                

      bino := False;
      for iino := 0 to slIgnoreNos.Count - 1 do
      begin
        if Copy(aFGUnOutRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
        begin
          bino := True;
          Break;
        end;
      end;

      if bino then
      begin
        Continue;
      end;

      aFGUnOutRptReader.Items[j].bSum := NameNeedSum(aFGUnOutRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGUnOutRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGUnOutRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGUnOutRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGUnOutRptReader.Items[j].snote;
      if aFGUnOutRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 5].Value := 'Y';
      irow := irow + 1;
    end;
  end;
       
  slIgnoreNos.Free;         
  slIgnoreNames.Free;
 
  AddBorder(ExcelApp, 1, 1, irow - 1, 5);
  AddColor(ExcelApp, 1, 1, 1, 5, $DBDCF2);
  AddHorizontalAlignment(ExcelApp, 1, 1, 1, 5, xlCenter);
  
  ExcelApp.Range[ExcelApp.Cells[2, 3], ExcelApp.Cells[irow - 1, 3]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    
end;
              
procedure Save_SF_KJ_ORT(ExcelApp: Variant; sl: TStringList);
var 
  irow: Integer;
  j: Integer;
  iino: Integer;
  bino: Boolean;     
  slIgnoreNos: TStringList;    
  slIgnoreNames: TStringList;
  i: Integer;
  aFGStockRptReader: TFGStockRptReader;
begin
  irow := 1;
 
  ExcelApp.Cells[irow, 1].Value := '料号';
  ExcelApp.Cells[irow, 2].Value := '产品名称';
  ExcelApp.Cells[irow, 3].Value := '库存总数';
  ExcelApp.Cells[irow, 4].Value := '其中';
  ExcelApp.Cells[irow + 1, 4].Value := '返工';
  ExcelApp.Cells[irow + 1, 5].Value := '待检验'#13#10'（超过3个月仓龄）';
  MergeCells(ExcelApp, 1, 4, 1, 5);
  ExcelApp.Cells[irow, 6].Value := '存货地点';
  ExcelApp.Cells[irow, 7].Value := '批次';
  ExcelApp.Cells[irow, 8].Value := '备注';   
  ExcelApp.Cells[irow, 9].Value := '汇总';

  MergeCells(ExcelApp, 1, 1, 2, 1);
  MergeCells(ExcelApp, 1, 2, 2, 2);
  MergeCells(ExcelApp, 1, 3, 2, 3);
  MergeCells(ExcelApp, 1, 6, 2, 6);
  MergeCells(ExcelApp, 1, 7, 2, 7);
  MergeCells(ExcelApp, 1, 8, 2, 8);  
  MergeCells(ExcelApp, 1, 9, 2, 9);

  ExcelApp.Columns[1].ColumnWidth := 17;
  ExcelApp.Columns[2].ColumnWidth := 21;
  ExcelApp.Columns[3].ColumnWidth := 9;
  ExcelApp.Columns[4].ColumnWidth := 9;  
  ExcelApp.Columns[5].ColumnWidth := 12;
  ExcelApp.Columns[6].ColumnWidth := 13;
  ExcelApp.Columns[7].ColumnWidth := 13;
  ExcelApp.Columns[8].ColumnWidth := 8;

  irow := 3;
                     
  slIgnoreNos := TfrmProjNameNo.GetIgnoreNos; 
  slIgnoreNames := TfrmProjNameNo.GetIgnoreName4Sum;

  for i := 0 to sl.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(sl.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin


        bino := False;
        for iino := 0 to slIgnoreNos.Count - 1 do
        begin
          if Copy(aFGStockRptReader.Items[j].snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
          begin
            bino := True;
            Break;
          end;
        end;

        if bino then
        begin
          Continue;
        end;

          
  //    if (Copy(aFGStockRptReader.Items[j].snumber, 1, 3) = '07.')
  //      or (Copy(aFGStockRptReader.Items[j].snumber, 1, 3) = '66.')
  //      or (Copy(aFGStockRptReader.Items[j].snumber, 1, 6) = '03.54.')
  //      or (Copy(aFGStockRptReader.Items[j].snumber, 1, 6) = '83.17.')
  //      or (Copy(aFGStockRptReader.Items[j].snumber, 1, 3) = '02.') then
  //    begin
  //      Continue;
  //    end;

      aFGStockRptReader.Items[j].bSum := NameNeedSum(aFGStockRptReader.Items[j].sName, slIgnoreNames);

      ExcelApp.Cells[irow, 1].Value := aFGStockRptReader.Items[j].snumber;
      ExcelApp.Cells[irow, 2].Value := aFGStockRptReader.Items[j].sname;
      ExcelApp.Cells[irow, 3].Value := aFGStockRptReader.Items[j].dqty;
      ExcelApp.Cells[irow, 4].Value := aFGStockRptReader.Items[j].drework;
      ExcelApp.Cells[irow, 5].Value := aFGStockRptReader.Items[j].duncheck;
      ExcelApp.Cells[irow, 6].Value := aFGStockRptReader.Items[j].saddr;
      ExcelApp.Cells[irow, 7].Value := aFGStockRptReader.Items[j].sBatchNo;
      ExcelApp.Cells[irow, 8].Value := aFGStockRptReader.Items[j].snote;
      if aFGStockRptReader.Items[j].bSum then
        ExcelApp.Cells[irow, 9].Value := 'Y';
      irow := irow + 1;
    end;
  end;

  slIgnoreNos.Free;     
  slIgnoreNames.Free;
 
  AddBorder(ExcelApp, 1, 1, irow - 1, 9);
  AddColor(ExcelApp, 1, 1, 2, 9, $DEF1EB);
  AddHorizontalAlignment(ExcelApp, 1, 1, 2, 9, xlCenter);
  
  ExcelApp.Range[ExcelApp.Cells[3, 3], ExcelApp.Cells[irow - 1, 5]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    
end;

function GetSumWin(const sProjNo: string; aFGRptWinReader: TFGRptWinReader;
  slWin: TStringList; bBS: Boolean): Double;
var
  aFGWinReader: TFGWinReader; 
  i, j: Integer;
  iPos: Integer;
begin
  Result := 0;
   
  for i := 0 to slWin.Count - 1 do
  begin
    aFGWinReader := TFGWinReader(slWin.Objects[i]);
    for j := 0 to aFGWinReader.Count - 1 do
    begin                                                           
      if Copy(aFGWinReader.Items[j]^.snumber, 1, 5) = sProjNo then
      begin
        if aFGWinReader.Items[j].snote = '返工入库' then Continue;   

        if not aFGWinReader.Items[j].bSum then Continue;

        Result := Result + aFGWinReader.Items[j].dqty;
      end;
    end;
  end;

  for j := 0 to aFGRptWinReader.Count - 1 do
  begin                                 
    if Copy(aFGRptWinReader.Items[j]^.snumber, 1, 5) = sProjNo then
    begin
      if aFGRptWinReader.Items[j].snote = '返工入库' then Continue;
      iPos := Pos('BS', aFGRptWinReader.Items[j]^.sFac);
      if (iPos > 0) and not bBS then Continue; // 如果包含 BS， 但是不是 BS的，跳过
      if (iPos = 0) and bBS then Continue;     // 如果不包含 BS， 但是 BS 的，跳过    
      
      if not aFGRptWinReader.Items[j].bSum then Continue;

      Result := Result + aFGRptWinReader.Items[j].dqty;
    end;
  end;
end;

function GetSumStockAlll(const sProjNo: string; slStock, slDB: TStringList): Double;
var
  aFGStockRptReader: TFGStockRptReader;
  i, j: Integer;
begin
  Result := 0;
  
  for i := 0 to slStock.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin            
      if not aFGStockRptReader.Items[j].bSum then Continue;
                           
      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].dqty;
    end;
  end;

  for i := 0 to slDB.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin                
      if not aFGStockRptReader.Items[j].bSum then Continue;
                                      
      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].dqty;
    end;
  end;               
end;

function GetSumStockNoUnOut(const sProjNo: string; slStock, slDB, slUnOut: TStringList): Double;
var
  aFGStockRptReader: TFGStockRptReader;
  aFGUnOutRptReader: TFGUnOutRptReader;
  i, j: Integer;
begin
  Result := 0;
  
  for i := 0 to slStock.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin     
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].dqty;
    end;
  end;

  for i := 0 to slDB.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin     
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].dqty;
    end;
  end;

  for i := 0 to slUnOut.Count - 1 do
  begin
    aFGUnOutRptReader := TFGUnOutRptReader(slUnOut.Objects[i]);
    for j := 0 to aFGUnOutRptReader.Count - 1 do
    begin   
      if not aFGUnOutRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGUnOutRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result - aFGUnOutRptReader.Items[j].dqty;
    end;
  end;
end;

function GetSumStockNoUnOut_zhouzuan(const sProjNo: string; slStock, slDB: TStringList): Double;
var
  aFGStockRptReader: TFGStockRptReader; 
  i, j: Integer;
begin
  Result := 0;
  
  for i := 0 to slStock.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin
      if Pos('周转', aFGStockRptReader.Items[j]^.saddr) <= 0 then Continue;
             
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].dqty;
    end;
  end;

  for i := 0 to slDB.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin                 
      if Pos('周转', aFGStockRptReader.Items[j]^.saddr) <= 0 then Continue;
           
      if not aFGStockRptReader.Items[j]^.bSum then Continue;
                         
      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].dqty;
    end;
  end;
end;

function GetSumStockRework(const sProjNo: string; slStock, slDB: TStringList): Double;
var
  aFGStockRptReader: TFGStockRptReader; 
  i, j: Integer;
begin
  Result := 0;
  
  for i := 0 to slStock.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin   
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].drework;
    end;
  end;

  for i := 0 to slDB.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin      
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].drework;
    end;
  end; 
end;

function GetSumStockUncheck(const sProjNo: string; slStock, slDB: TStringList): Double;
var
  aFGStockRptReader: TFGStockRptReader; 
  i, j: Integer;
begin
  Result := 0;
  
  for i := 0 to slStock.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin     
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].duncheck;
    end;
  end;

  for i := 0 to slDB.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin      
      if not aFGStockRptReader.Items[j]^.bSum then Continue;

      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j].duncheck;
    end;
  end; 
end;

function GetSumSF_KJ_ORT(const sProjNo: string; sl: TStringList): Double;
var
  j: Integer;
  aFGStockRptReader: TFGStockRptReader;
  i: Integer;
begin
  Result := 0;

  for i := 0 to sl.Count - 1 do
  begin
    aFGStockRptReader := TFGStockRptReader(sl.Objects[i]);
    for j := 0 to aFGStockRptReader.Count - 1 do
    begin                    
      if not aFGStockRptReader.Items[j]^.bSum then Continue;
      
      if Copy(aFGStockRptReader.Items[j]^.snumber, 1, 5) = sProjNo then
        Result := Result + aFGStockRptReader.Items[j]^.dqty;
    end;
  end;
end;
      
procedure TfrmMakeFGReport.btnSave2Click(Sender: TObject);
var
  slWin: TStringList;
  slStock: TStringList;
  slUnOut: TStringList;
  slDB: TStringList;
                      
  slWinBS: TStringList;
  slStockBS: TStringList;
  slUnOutBS: TStringList;
  slDBBS: TStringList;

  slSF: TStringList;    
  slKJ: TStringList;
  slORT: TStringList;

  slNumber: TStringList;

  i: Integer;
  aFGWinReader: TFGWinReader;
  aFGStockRptReader: TFGStockRptReader;
  aFGUnOutRptReader: TFGUnOutRptReader;
//  aFGDBRptReader: TFGDBRptReader;

  aFGStockRptReader_kjX: TFGStockRptReader;
  aFGStockRptReader_sfX: TFGStockRptReader;
  aFGStockRptReader_ortX: TFGStockRptReader;

  aFGRptWinReader: TFGRptWinReader;

  pn: PFGNumberRecord;

  sfile: string;
  ExcelApp, WorkBook: Variant;
  irow: Integer;

  slProj: TStringList;
  sProjNo0: string;
  irow1, irow2: Integer;

  irow_oem, irow_odm: Integer;

  slProjNo_OEM: TStringList;
  slProjNo_ODM: TStringList;

  sSumWin, sSumStockAll, dSumStockNoUnOut: Double;
  dRework, dUncheck: Double;
  slNumberWritten: TStringList;
  slIgnoreNos: TStringList;
  iino: Integer;
  bino: Boolean;
begin
  sfile := '成品报表' + FormatDateTime('yyyyMMdd', Now);
  if not ExcelSaveDialog(sfile) then Exit;

  slIgnoreNos := TfrmProjNameNo.GetIgnoreNos;

  slWin := TStringList.Create;
  slStock := TStringList.Create;
  slUnOut := TStringList.Create;
  slDB := TStringList.Create;
        
  slWinBS := TStringList.Create;
  slStockBS := TStringList.Create;
  slUnOutBS := TStringList.Create;
  slDBBS := TStringList.Create;
  
  slSF := TStringList.Create;
  slKJ := TStringList.Create;
  slORT := TStringList.Create;

  slNumber := TStringList.Create;
  slProj := TStringList.Create;
  slNumberWritten := TStringList.Create;

  slProjNo_OEM := TfrmProjNameNo.GetProjNos_OEM;
  slProjNo_ODM := TfrmProjNameNo.GetProjNos_ODM;

  aFGRptWinReader := TFGRptWinReader.Create(leFGRptWin.Text, Log);  
  if not aFGRptWinReader.ReadOk then
  begin
    Memo2.Lines.Add('读取入库失败 成品库存报表  ' + leFGRptWin.Text);
  end;
                         
  try
    aFGRptWinReader.GetNumberSet(slNumber);



    for i := 0 to mmoSF.Lines.Count - 1 do
    begin
      if Trim(mmoSF.Lines[i]) = '' then Continue;
 
      aFGStockRptReader_sfX := TFGStockRptReader.Create(mmoSF.Lines[i], Log);
      if not aFGStockRptReader_sfX.ReadOk then
      begin
        Memo2.Lines.Add('读取库存失败 物流仓 顺丰  ' + mmoSF.Lines[i]);
      end;
         
      slSF.AddObject(mmoSF.Lines[i], aFGStockRptReader_sfX);

      aFGStockRptReader_sfX.GetNumberSet(slNumber);
    end;
        

    for i := 0 to mmoKJ.Lines.Count - 1 do
    begin
      if Trim(mmoKJ.Lines[i]) = '' then Continue;

      aFGStockRptReader_kjX := TFGStockRptReader.Create(mmoKJ.Lines[i], Log);
      if not aFGStockRptReader_kjX.ReadOk then
      begin
        Memo2.Lines.Add('读取库存失败 物流仓 科捷  ' + mmoKJ.Lines[i]);
      end;
         
      slKJ.AddObject(mmoKJ.Lines[i], aFGStockRptReader_kjX);

      aFGStockRptReader_kjX.GetNumberSet(slNumber);
    end;
       

    for i := 0 to mmoORT.Lines.Count - 1 do
    begin
      if Trim(mmoORT.Lines[i]) = '' then Continue;
 
      aFGStockRptReader_ortX := TFGStockRptReader.Create(mmoORT.Lines[i], Log);
      if not aFGStockRptReader_ortX.ReadOk then
      begin
        Memo2.Lines.Add('读取库存失败 物流仓 欧瑞特  ' + mmoORT.Lines[i]);
      end;
         
      slORT.AddObject(mmoORT.Lines[i], aFGStockRptReader_ortX);

      aFGStockRptReader_ortX.GetNumberSet(slNumber);
    end;
       




    

    for i := 0 to mmoWins.Lines.Count - 1 do
    begin
      if Trim(mmoWins.Lines[i]) = '' then Continue;
      aFGWinReader := TFGWinReader.Create(mmoWins.Lines[i], Log);      
      if not aFGWinReader.ReadOk then
      begin
        Memo2.Lines.Add('读取入库失败 ' + mmoWins.Lines[i]);
      end;
      slWin.AddObject(mmoWins.Lines[i], aFGWinReader);

      aFGWinReader.GetNumberSet(slNumber);
    end;
        
    for i := 0 to mmoStocks.Lines.Count - 1 do
    begin                 
      if Trim(mmoStocks.Lines[i]) = '' then Continue;
      aFGStockRptReader := TFGStockRptReader.Create(mmoStocks.Lines[i], Log);
      if not aFGStockRptReader.ReadOk then
      begin
        Memo2.Lines.Add('读取库存失败 ' + mmoStocks.Lines[i]);
      end;
      slStock.AddObject(mmoStocks.Lines[i], aFGStockRptReader);
      
      aFGStockRptReader.GetNumberSet(slNumber);
    end;

    for i := 0 to mmoUnOut.Lines.Count - 1 do
    begin          
      if Trim(mmoUnOut.Lines[i]) = '' then Continue;
      aFGUnOutRptReader := TFGUnOutRptReader.Create(mmoUnOut.Lines[i], Log);
      if not aFGUnOutRptReader.ReadOk then
      begin
        Memo2.Lines.Add('读取未出库失败 ' + mmoUnOut.Lines[i]);
      end;
      slUnOut.AddObject(mmoUnOut.Lines[i], aFGUnOutRptReader);

      aFGUnOutRptReader.GetNumberSet(slNumber);
    end;
                             
    for i := 0 to mmoDB.Lines.Count - 1 do
    begin       
      if Trim(mmoDB.Lines[i]) = '' then Continue;
      aFGStockRptReader := TFGStockRptReader.Create(mmoDB.Lines[i], Log);  
      if not aFGStockRptReader.ReadOk then
      begin
        Memo2.Lines.Add('读取周转仓失败 ' + mmoDB.Lines[i]);
      end;
      slDB.AddObject(mmoDB.Lines[i], aFGStockRptReader);
      
      aFGStockRptReader.GetNumberSet(slNumber);
    end;

    // ------ BS  ------------------------------------------

    for i := 0 to mmoWinsBS.Lines.Count - 1 do
    begin     
      if Trim(mmoWinsBS.Lines[i]) = '' then Continue;
      aFGWinReader := TFGWinReader.Create(mmoWinsBS.Lines[i], Log);    
      if not aFGWinReader.ReadOk then
      begin
        Memo2.Lines.Add('读取入库失败 BS ' + mmoWinsBS.Lines[i]);
      end;
      slWinBS.AddObject(mmoWinsBS.Lines[i], aFGWinReader);

      aFGWinReader.GetNumberSet(slNumber);
    end;
        
    for i := 0 to mmoStocksBS.Lines.Count - 1 do
    begin      
      if Trim(mmoStocksBS.Lines[i]) = '' then Continue;
      aFGStockRptReader := TFGStockRptReader.Create(mmoStocksBS.Lines[i], Log);
      if not aFGStockRptReader.ReadOk then
      begin
        Memo2.Lines.Add('读取库存失败 BS ' + mmoStocksBS.Lines[i]);
      end;
      slStockBS.AddObject(mmoStocksBS.Lines[i], aFGStockRptReader);
      
      aFGStockRptReader.GetNumberSet(slNumber);
    end;

    for i := 0 to mmoUnOutBS.Lines.Count - 1 do
    begin       
      if Trim(mmoUnOutBS.Lines[i]) = '' then Continue;
      aFGUnOutRptReader := TFGUnOutRptReader.Create(mmoUnOutBS.Lines[i], Log);  
      if not aFGUnOutRptReader.ReadOk then
      begin
        Memo2.Lines.Add('读取未出库失败 BS ' + mmoUnOutBS.Lines[i]);
      end;
      slUnOutBS.AddObject(mmoUnOutBS.Lines[i], aFGUnOutRptReader);

      aFGUnOutRptReader.GetNumberSet(slNumber);
    end;
                             
    for i := 0 to mmoDBBS.Lines.Count - 1 do
    begin                   
      if Trim(mmoDBBS.Lines[i]) = '' then Continue;
      aFGStockRptReader := TFGStockRptReader.Create(mmoDBBS.Lines[i], Log);  
      if not aFGStockRptReader.ReadOk then
      begin
        Memo2.Lines.Add('读取周转仓失败 BS ' + mmoDBBS.Lines[i]);
      end;
      slDBBS.AddObject(mmoDBBS.Lines[i], aFGStockRptReader);
      
      aFGStockRptReader.GetNumberSet(slNumber);
    end;

    // -----------------------------------------------------    



    slNumber.Sort;


    // 开始保存 Excel
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

    WorkBook := ExcelApp.WorkBooks.Add;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;
                     
    ExcelApp.Sheets[1].Activate; 
    ExcelApp.Sheets[1].Name := '最新汇总';
                              
    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[1]);
    ExcelApp.Sheets[2].Activate;
    ExcelApp.Sheets[2].Name := '结存';

    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[2]);
    ExcelApp.Sheets[3].Activate;
    ExcelApp.Sheets[3].Name := '入库';

    SaveWin(ExcelApp, aFGRptWinReader, slWin, slWinBS);


    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[3]);
    ExcelApp.Sheets[4].Activate;
    ExcelApp.Sheets[4].Name := '库存';

    SaveStock(ExcelApp, slStock, slStockBS, slDB, slDBBS);


    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[4]);
    ExcelApp.Sheets[5].Activate;
    ExcelApp.Sheets[5].Name := '未出货';

    SaveUnOut(ExcelApp, slUnOut, slUnOutBS);


    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[5]);
    ExcelApp.Sheets[6].Activate;
    ExcelApp.Sheets[6].Name := '顺丰仓库存';

    Save_SF_KJ_ORT(ExcelApp, slSF);


    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[6]);
    ExcelApp.Sheets[7].Activate;
    ExcelApp.Sheets[7].Name := '科捷（整机）';

    Save_SF_KJ_ORT(ExcelApp, slKJ);


    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[7]);
    ExcelApp.Sheets[8].Activate;
    ExcelApp.Sheets[8].Name := '欧瑞特（整机）';
                                         
    Save_SF_KJ_ORT(ExcelApp, slORT);
    

    try              
      ExcelApp.Sheets[2].Activate;
      
      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '料号';
      ExcelApp.Cells[irow, 2].Value := '物料描述';
      ExcelApp.Cells[irow, 3].Value := '批号';
      ExcelApp.Cells[irow, 4].Value := '入库';
      ExcelApp.Cells[irow, 5].Value := '实物结存';
      ExcelApp.Cells[irow, 6].Value := '未出货数量';
      ExcelApp.Cells[irow, 7].Value := '结存（实物结存-已抛单）';
      ExcelApp.Cells[irow, 8].Value := '顺丰仓库存';
      ExcelApp.Cells[irow, 9].Value := '科捷（整机）';
      ExcelApp.Cells[irow, 10].Value := '欧瑞特（整机）';

      ExcelApp.Columns[1].ColumnWidth := 17;
      ExcelApp.Columns[2].ColumnWidth := 65;
      ExcelApp.Columns[3].ColumnWidth := 5;
      ExcelApp.Columns[4].ColumnWidth := 9;
      ExcelApp.Columns[5].ColumnWidth := 9;
      ExcelApp.Columns[6].ColumnWidth := 9;
      ExcelApp.Columns[7].ColumnWidth := 9;
      ExcelApp.Columns[8].ColumnWidth := 9;
      ExcelApp.Columns[9].ColumnWidth := 9;
      ExcelApp.Columns[10].ColumnWidth := 9;

      sProjNo0 := '';
      irow1 := 0; 
      irow := 2;
      for i := 0 to slNumber.Count - 1 do
      begin
        pn := PFGNumberRecord(slNumber.Objects[i]);

        bino := False;
        for iino := 0 to slIgnoreNos.Count - 1 do
        begin
          if Copy(pn^.snumber, 1, Length(slIgnoreNos[iino])) = slIgnoreNos[iino] then
          begin
            bino := True;
            Break;
          end;
        end;

        if bino then
        begin
          Continue;
        end;
        
//        if (Copy(pn^.snumber, 1, 3) = '07.')
//          or (Copy(pn^.snumber, 1, 3) = '66.')
//          or (Copy(pn^.snumber, 1, 6) = '03.54.')
//          or (Copy(pn^.snumber, 1, 6) = '83.17.')          
//          or (Copy(pn^.snumber, 1, 3) = '02.') then
//        begin
//          Continue;
//        end;

        if slNumberWritten.IndexOf(pn^.snumber) >= 0 then
        begin
          Continue;
        end;
        slNumberWritten.Add(pn^.snumber);

        ExcelApp.Cells[irow, 1].Value := pn^.snumber;
        ExcelApp.Cells[irow, 2].Value := pn^.sname;
        if Trim(pn^.sBatchNo) = '' then
          ExcelApp.Cells[irow, 3].Value := '   '
        else
          ExcelApp.Cells[irow, 3].Value := pn^.sBatchNo;

        ExcelApp.Cells[irow, 4].Value := '=SUMIF(入库!B:B,A:A,入库!D:D)';
        ExcelApp.Cells[irow, 5].Value := '=SUMIF(库存!A:A,A:A,库存!C:C)';
        ExcelApp.Cells[irow, 6].Value := '=SUMIF(未出货!A:A,A:A,未出货!C:C)';
        ExcelApp.Cells[irow, 7].Value := '=E' + IntToStr(irow) + '-F' + IntToStr(irow);
        ExcelApp.Cells[irow, 8].Value := '=SUMIF(顺丰仓库存!A:A,A:A,顺丰仓库存!C:C)';
        ExcelApp.Cells[irow, 9].Value := '=SUMIF(''科捷（整机）''!A:A,A:A,''科捷（整机）''!C:C)';
        ExcelApp.Cells[irow, 10].Value := '=SUMIF(''欧瑞特（整机）''!A:A,A:A,''欧瑞特（整机）''!C:C)';
 
        if sProjNo0 <> Copy(pn^.snumber, 1, 5) then  // 新一个项目
        begin
          if sProjNo0 = '' then  // 第一行
          begin
          end
          else
          begin
            irow2 := irow - 1;
            slProj.AddObject(sProjNo0 + '=' + IntToStr(irow1), TObject(irow2));
          end;      
          irow1 := irow;
        end;
        sProjNo0 := Copy(pn^.snumber, 1, 5);

        irow := irow + 1;
      end;
      
      irow2 := irow - 1;
      slProj.AddObject(sProjNo0 + '=' + IntToStr(irow1), TObject(irow2));

      irow1 := irow;
      
      for i := 0 to slProj.Count - 1 do
      begin
        irow1 := StrToInt(slProj.ValueFromIndex[i]);
        irow2 := Integer( slProj.Objects[i] );
        if slProjNo_OEM.Values[ slProj.Names[i] ] <> '' then
        begin           
          ExcelApp.Cells[irow, 1].Value := slProjNo_OEM.Values[ slProj.Names[i] ];
        end
        else
        begin        
          if slProjNo_ODM.Values[ slProj.Names[i] ] <> '' then
          begin        
            ExcelApp.Cells[irow, 1].Value := slProjNo_ODM.Values[ slProj.Names[i] ];
          end
          else
          begin
            ExcelApp.Cells[irow, 1].Value := slProj.Names[i];
          end;
        end;
        MergeCells(ExcelApp, irow, 1, irow, 3);
        ExcelApp.Cells[irow, 4].Value := '=SUM(D' + IntToStr(irow1) + ':D' + IntToStr(irow2) + ')';
        ExcelApp.Cells[irow, 5].Value := '=SUM(E' + IntToStr(irow1) + ':E' + IntToStr(irow2) + ')';
        ExcelApp.Cells[irow, 6].Value := '=SUM(F' + IntToStr(irow1) + ':F' + IntToStr(irow2) + ')';
        ExcelApp.Cells[irow, 7].Value := '=E' + IntToStr(irow) + '-F' + IntToStr(irow);
        ExcelApp.Cells[irow, 8].Value := '=SUM(H' + IntToStr(irow1) + ':H' + IntToStr(irow2) + ')';
        ExcelApp.Cells[irow, 9].Value := '=SUM(I' + IntToStr(irow1) + ':I' + IntToStr(irow2) + ')';
        ExcelApp.Cells[irow, 10].Value := '=SUM(J' + IntToStr(irow1) + ':J' + IntToStr(irow2) + ')';
        irow := irow + 1;
      end;
      irow2 := irow - 1;

      AddColor(ExcelApp, 1, 1, 1, 7, $50D092);     
      AddColor(ExcelApp, 1, 8, 1, 10, $B4D5FC);
      AddBorder(ExcelApp, 1, 1, irow - 1, 10); 
      AddHorizontalAlignment(ExcelApp, irow1, 1, irow2, 1, xlCenter);
                 
      ExcelApp.Range[ExcelApp.Cells[2, 4], ExcelApp.Cells[irow - 1, 10]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
      

      //////////////////////////////////////////////////////////////////////////////

        
      ExcelApp.Sheets[1].Activate;

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '机型';
      ExcelApp.Cells[irow, 2].Value := '入库';
      ExcelApp.Cells[irow, 3].Value := '含未出货结存';
      ExcelApp.Cells[irow, 4].Value := '结存'#13#10'（实物结存-已抛单） ';
      ExcelApp.Cells[irow, 5].Value := '顺丰仓库存';
      ExcelApp.Cells[irow, 6].Value := '科捷库存';
      ExcelApp.Cells[irow, 7].Value := '欧瑞特库存';
      ExcelApp.Cells[irow, 8].Value := '合计';
 
      ExcelApp.Columns[1].ColumnWidth := 13;
      ExcelApp.Columns[2].ColumnWidth := 12;
      ExcelApp.Columns[3].ColumnWidth := 12;
      ExcelApp.Columns[4].ColumnWidth := 22;
      ExcelApp.Columns[5].ColumnWidth := 12;
      ExcelApp.Columns[6].ColumnWidth := 12;
      ExcelApp.Columns[7].ColumnWidth := 12;
      ExcelApp.Columns[8].ColumnWidth := 12;

      ExcelApp.Rows[irow].RowHeight := 27;

      AddColor(ExcelApp, 1, 1, 1, 8, $DAC0CC);

      irow := 2;

      //////  OEM  //////////////////////////////

      irow1 := irow;
      for i := 0 to slProjNo_OEM.Count - 1 do
      begin
        ExcelApp.Cells[irow, 1].Value := slProjNo_OEM.ValueFromIndex[i];     
        ExcelApp.Cells[irow, 2].Value := GetSumWin(slProjNo_OEM.Names[i], aFGRptWinReader, slWin, False);
        ExcelApp.Cells[irow, 3].Value := GetSumStockAlll(slProjNo_OEM.Names[i], slStock, slDB);
        ExcelApp.Cells[irow, 4].Value := GetSumStockNoUnOut(slProjNo_OEM.Names[i], slStock, slDB, slUnOut);
        ExcelApp.Cells[irow, 5].Value := GetSumSF_KJ_ORT(slProjNo_OEM.Names[i], slSF);
        ExcelApp.Cells[irow, 6].Value := GetSumSF_KJ_ORT(slProjNo_OEM.Names[i], slKJ);
        ExcelApp.Cells[irow, 7].Value := GetSumSF_KJ_ORT(slProjNo_OEM.Names[i], slORT);
        ExcelApp.Cells[irow, 8].Value := '=C' + IntToStr(irow) + '+E' + IntToStr(irow) + '+F' + IntToStr(irow) + '+G' + IntToStr(irow);
        irow := irow + 1;
                   

        sSumWin := GetSumWin(slProjNo_OEM.Names[i], aFGRptWinReader, slWinBS, True);
        sSumStockAll := GetSumStockAlll(slProjNo_OEM.Names[i], slStockBS, slDBBS);
        dSumStockNoUnOut := GetSumStockNoUnOut(slProjNo_OEM.Names[i], slStockBS, slDBBS, slUnOutBS);
        if (sSumWin > 0) or (sSumStockAll > 0) or (dSumStockNoUnOut > 0) then
        begin
          ExcelApp.Cells[irow, 1].Value := slProjNo_OEM.ValueFromIndex[i] + 'BS';
          ExcelApp.Cells[irow, 2].Value := sSumWin;
          ExcelApp.Cells[irow, 3].Value := sSumStockAll;
          ExcelApp.Cells[irow, 4].Value := dSumStockNoUnOut;
          ExcelApp.Cells[irow, 5].Value := 0;
          ExcelApp.Cells[irow, 6].Value := 0;
          ExcelApp.Cells[irow, 7].Value := 0;
          ExcelApp.Cells[irow, 8].Value := '=C' + IntToStr(irow) + '+E' + IntToStr(irow) + '+F' + IntToStr(irow) + '+G' + IntToStr(irow);
          irow := irow + 1;
        end;
      end;
      irow2 := irow - 1;

      irow_oem := irow;

      ExcelApp.Cells[irow, 1].Value := '自研产品小计';    
      ExcelApp.Cells[irow, 2].Value := '=SUM(B' + IntToStr(irow1) + ':B' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 3].Value := '=SUM(C' + IntToStr(irow1) + ':C' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 4].Value := '=SUM(D' + IntToStr(irow1) + ':D' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 5].Value := '=SUM(E' + IntToStr(irow1) + ':E' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 6].Value := '=SUM(F' + IntToStr(irow1) + ':F' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 7].Value := '=SUM(G' + IntToStr(irow1) + ':G' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 8].Value := '=SUM(H' + IntToStr(irow1) + ':H' + IntToStr(irow2) + ')';
                
      AddColor(ExcelApp, irow, 1, irow, 8, $00FFFF);

      irow := irow + 1;

              
      //////  ODM  //////////////////////////////

      irow1 := irow;
      for i := 0 to slProjNo_ODM.Count - 1 do
      begin
        ExcelApp.Cells[irow, 1].Value := slProjNo_ODM.ValueFromIndex[i];
        ExcelApp.Cells[irow, 2].Value := GetSumWin(slProjNo_ODM.Names[i], aFGRptWinReader, slWin, False);
        ExcelApp.Cells[irow, 3].Value := GetSumStockAlll(slProjNo_ODM.Names[i], slStock, slDB);
        ExcelApp.Cells[irow, 4].Value := GetSumStockNoUnOut(slProjNo_ODM.Names[i], slStock, slDB, slUnOut);
        ExcelApp.Cells[irow, 5].Value := GetSumSF_KJ_ORT(slProjNo_ODM.Names[i], slSF);
        ExcelApp.Cells[irow, 6].Value := GetSumSF_KJ_ORT(slProjNo_ODM.Names[i], slKJ);
        ExcelApp.Cells[irow, 7].Value := GetSumSF_KJ_ORT(slProjNo_ODM.Names[i], slORT);
        ExcelApp.Cells[irow, 8].Value := '=C' + IntToStr(irow) + '+E' + IntToStr(irow) + '+F' + IntToStr(irow) + '+G' + IntToStr(irow);
        irow := irow + 1;    
                   

        sSumWin := GetSumWin(slProjNo_ODM.Names[i], aFGRptWinReader, slWinBS, True);
        sSumStockAll := GetSumStockAlll(slProjNo_ODM.Names[i], slStockBS, slDBBS);
        dSumStockNoUnOut := GetSumStockNoUnOut(slProjNo_ODM.Names[i], slStockBS, slDBBS, slUnOutBS);
        if (sSumWin > 0) or (sSumStockAll > 0) or (dSumStockNoUnOut > 0) then
        begin
          ExcelApp.Cells[irow, 1].Value := slProjNo_ODM.ValueFromIndex[i] + 'BS';
          ExcelApp.Cells[irow, 2].Value := sSumWin;
          ExcelApp.Cells[irow, 3].Value := sSumStockAll;
          ExcelApp.Cells[irow, 4].Value := dSumStockNoUnOut;
          ExcelApp.Cells[irow, 5].Value := 0;
          ExcelApp.Cells[irow, 6].Value := 0;
          ExcelApp.Cells[irow, 7].Value := 0;
          ExcelApp.Cells[irow, 8].Value := '=C' + IntToStr(irow) + '+E' + IntToStr(irow) + '+F' + IntToStr(irow) + '+G' + IntToStr(irow);
          irow := irow + 1;
        end;
      end;
      irow2 := irow - 1;

      irow_odm := irow;

      ExcelApp.Cells[irow, 1].Value := '外研公司小计';    
      ExcelApp.Cells[irow, 2].Value := '=SUM(B' + IntToStr(irow1) + ':B' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 3].Value := '=SUM(C' + IntToStr(irow1) + ':C' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 4].Value := '=SUM(D' + IntToStr(irow1) + ':D' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 5].Value := '=SUM(E' + IntToStr(irow1) + ':E' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 6].Value := '=SUM(F' + IntToStr(irow1) + ':F' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 7].Value := '=SUM(G' + IntToStr(irow1) + ':G' + IntToStr(irow2) + ')';
      ExcelApp.Cells[irow, 8].Value := '=SUM(H' + IntToStr(irow1) + ':H' + IntToStr(irow2) + ')';

      AddColor(ExcelApp, irow, 1, irow, 8, $00FFFF);

      ///////////  合计  ////////////////////////
      irow := irow + 1;
      ExcelApp.Cells[irow, 1].Value := '合计';
      ExcelApp.Cells[irow, 2].Value := '=B' + IntToStr(irow_oem) + '+B' + IntToStr(irow_odm);
      ExcelApp.Cells[irow, 3].Value := '=C' + IntToStr(irow_oem) + '+C' + IntToStr(irow_odm);
      ExcelApp.Cells[irow, 4].Value := '=D' + IntToStr(irow_oem) + '+D' + IntToStr(irow_odm);
      ExcelApp.Cells[irow, 5].Value := '=E' + IntToStr(irow_oem) + '+E' + IntToStr(irow_odm);
      ExcelApp.Cells[irow, 6].Value := '=F' + IntToStr(irow_oem) + '+F' + IntToStr(irow_odm);
      ExcelApp.Cells[irow, 7].Value := '=G' + IntToStr(irow_oem) + '+G' + IntToStr(irow_odm);
      ExcelApp.Cells[irow, 8].Value := '=H' + IntToStr(irow_oem) + '+H' + IntToStr(irow_odm);
                      
      AddColor(ExcelApp, irow, 1, irow, 8, $4696F7);
      AddHorizontalAlignment(ExcelApp, 1, 1, 1, 8, xlCenter);       
      ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 8]].Font.Bold  := True;

      ExcelApp.Range[ExcelApp.Cells[2, 2], ExcelApp.Cells[irow, 8]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
        
      ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, 8]].Font.Size := 10; 


      //////////////////////////////////////////////////////////////////////////
      //////////////////////////////////////////////////////////////////////////
      //////////////////////////////////////////////////////////////////////////
                  
      irow := irow + 2;
      irow1 := irow;

      ExcelApp.Cells[irow, 2].Value := '待返工与在途库存明细';
      MergeCells(ExcelApp, irow, 2, irow, 7);

      irow := irow + 1;

      ExcelApp.Cells[irow, 2].Value := '机型';
      ExcelApp.Cells[irow, 3].Value := '结存（实物结存-已抛单）';
      ExcelApp.Cells[irow, 4].Value := '待返工';
      ExcelApp.Cells[irow, 5].Value := '超期';
      ExcelApp.Cells[irow, 6].Value := '在途';
      ExcelApp.Cells[irow, 7].Value := '可发库存';
                                                   
      AddHorizontalAlignment(ExcelApp, irow1, 2, irow, 7, xlCenter);
          
      //////  OEM  //////////////////////////////

      irow := irow + 1;
      for i := 0 to slProjNo_OEM.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slProjNo_OEM.ValueFromIndex[i]; 
        ExcelApp.Cells[irow, 3].Value := GetSumStockNoUnOut(slProjNo_OEM.Names[i], slStock, slDB, slUnOut)
          + GetSumSF_KJ_ORT(slProjNo_OEM.Names[i], slKJ) + GetSumSF_KJ_ORT(slProjNo_OEM.Names[i], slORT);    
        ExcelApp.Cells[irow, 4].Value := GetSumStockRework(slProjNo_OEM.Names[i], slStock, slDB);        
        ExcelApp.Cells[irow, 5].Value := GetSumStockUncheck(slProjNo_OEM.Names[i], slStock, slDB);
        ExcelApp.Cells[irow, 6].Value := GetSumStockNoUnOut_zhouzuan(slProjNo_OEM.Names[i], slStock, slDB);    
        ExcelApp.Cells[irow, 7].Value := '=C' + IntToStr(irow) + '-D' + IntToStr(irow) + '-E' + IntToStr(irow) + '-F' + IntToStr(irow); 
        
        irow := irow + 1;
 
        dSumStockNoUnOut := GetSumStockNoUnOut(slProjNo_OEM.Names[i], slStockBS, slDBBS, slUnOutBS);
        dRework := GetSumStockRework(slProjNo_OEM.Names[i], slStockBS, slDBBS);
        dUncheck := GetSumStockUncheck(slProjNo_OEM.Names[i], slStockBS, slDBBS);
        if (dSumStockNoUnOut > 0) or (dRework > 0) or (dUncheck > 0) then
        begin
          ExcelApp.Cells[irow, 2].Value := slProjNo_OEM.ValueFromIndex[i] + 'BS';
          ExcelApp.Cells[irow, 3].Value := dSumStockNoUnOut;
          ExcelApp.Cells[irow, 4].Value := dRework;
          ExcelApp.Cells[irow, 5].Value := dUncheck;
          ExcelApp.Cells[irow, 6].Value := 0;
          ExcelApp.Cells[irow, 7].Value := '=C' + IntToStr(irow) + '-D' + IntToStr(irow) + '-E' + IntToStr(irow) + '-F' + IntToStr(irow); 
          irow := irow + 1;
        end;
      end;

      //////  ODM  //////////////////////////////
 
      for i := 0 to slProjNo_ODM.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slProjNo_ODM.ValueFromIndex[i];
        ExcelApp.Cells[irow, 3].Value := GetSumStockNoUnOut(slProjNo_ODM.Names[i], slStock, slDB, slUnOut)
          + GetSumSF_KJ_ORT(slProjNo_ODM.Names[i], slKJ) + GetSumSF_KJ_ORT(slProjNo_ODM.Names[i], slORT);
        ExcelApp.Cells[irow, 4].Value := GetSumStockRework(slProjNo_ODM.Names[i], slStock, slDB);
        ExcelApp.Cells[irow, 5].Value := GetSumStockUncheck(slProjNo_ODM.Names[i], slStock, slDB);
        ExcelApp.Cells[irow, 6].Value := GetSumStockNoUnOut_zhouzuan(slProjNo_ODM.Names[i], slStock, slDB);    
        ExcelApp.Cells[irow, 7].Value := '=C' + IntToStr(irow) + '-D' + IntToStr(irow) + '-E' + IntToStr(irow) + '-F' + IntToStr(irow);
        
        irow := irow + 1;    
                   
        dSumStockNoUnOut := GetSumStockNoUnOut(slProjNo_ODM.Names[i], slStockBS, slDBBS, slUnOutBS);
        dRework := GetSumStockRework(slProjNo_ODM.Names[i], slStockBS, slDBBS);
        dUncheck := GetSumStockUncheck(slProjNo_ODM.Names[i], slStockBS, slDBBS);
        if (dSumStockNoUnOut > 0) or (dRework > 0) or (dUncheck > 0) then
        begin
          ExcelApp.Cells[irow, 2].Value := slProjNo_ODM.ValueFromIndex[i] + 'BS';
          ExcelApp.Cells[irow, 3].Value := dSumStockNoUnOut;
          ExcelApp.Cells[irow, 4].Value := dRework;
          ExcelApp.Cells[irow, 5].Value := dUncheck;
          ExcelApp.Cells[irow, 6].Value := 0;
          ExcelApp.Cells[irow, 7].Value := '=C' + IntToStr(irow) + '-D' + IntToStr(irow) + '-E' + IntToStr(irow) + '-F' + IntToStr(irow);
          irow := irow + 1;
        end;
      end; 

      AddBorder(ExcelApp, irow1, 2, irow - 1, 7);
      
 
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
    
  finally
    for i := 0 to slWin.Count - 1 do
    begin
      aFGWinReader := TFGWinReader(slWin.Objects[i]);
      aFGWinReader.Free;
    end;
    slWin.Free;
                   
    for i := 0 to slStock.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slStock.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slStock.Free;
                  
    for i := 0 to slUnOut.Count - 1 do
    begin
      aFGUnOutRptReader := TFGUnOutRptReader(slUnOut.Objects[i]);
      aFGUnOutRptReader.Free;
    end;
    slUnOut.Free;
                   
    for i := 0 to slDB.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slDB.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slDB.Free;
 
    // ------  BS --------------------------------------------------------
          
    for i := 0 to slWinBS.Count - 1 do
    begin
      aFGWinReader := TFGWinReader(slWinBS.Objects[i]);
      aFGWinReader.Free;
    end;
    slWinBS.Free;
                   
    for i := 0 to slStockBS.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slStockBS.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slStockBS.Free;
                  
    for i := 0 to slUnOutBS.Count - 1 do
    begin
      aFGUnOutRptReader := TFGUnOutRptReader(slUnOutBS.Objects[i]);
      aFGUnOutRptReader.Free;
    end;
    slUnOutBS.Free;
                   
    for i := 0 to slDBBS.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slDBBS.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slDBBS.Free;
                   
    for i := 0 to slSf.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slSf.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slSf.Free;   
                   
    for i := 0 to slKJ.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slKJ.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slKJ.Free;   
                   
    for i := 0 to slORT.Count - 1 do
    begin
      aFGStockRptReader := TFGStockRptReader(slORT.Objects[i]);
      aFGStockRptReader.Free;
    end;
    slORT.Free;

    // -------------------------------------------------------------------
          
    for i := 0 to slNumber.Count - 1 do
    begin
      pn := PFGNumberRecord(slNumber.Objects[i]);
      Dispose(pn);
    end;

    slNumber.Free;
    slNumberWritten.Free;

    slProj.Free;
    slProjNo_OEM.Free;
    slProjNo_ODM.Free;

    aFGRptWinReader.Free;

    slIgnoreNos.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

(*
1、魅力周转仓格式
2、顺丰仓的格式
3、华贝 BS的， 入库，格式没改
4、华贝BS的未出货，格式
5、闻泰BS入库    
6、闻泰BS库存    
7、闻泰BS未出货

*)

procedure TfrmMakeFGReport.btnWinsBSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoWinsBS.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnStocksBSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoStocksBS.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnUnOutBSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoUnOutBS.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnDBBSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoDBBS.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmMakeFGReport.btnFGRptWinClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  leFGRptWin.Text := sfile;
end;

procedure TfrmMakeFGReport.Log(const s: string);
begin
  Memo1.Lines.Add(s);
end;

end.
