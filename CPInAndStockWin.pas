unit CPInAndStockWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin, CommUtils, IniFiles,
  SAPMB51Reader2, CPInAndStockReader, SAPStockReader2, SAPMrpAreaStockReader,
  ComObj;

type
  TfrmCPInAndStock = class(TForm)
    leMB51: TLabeledEdit;
    btnMB51: TButton;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    leCPInAndStock: TLabeledEdit;
    btnCPInAndStock: TButton;
    Memo1: TMemo;
    mmoProjNo2Name: TMemo;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    leSAPStock: TLabeledEdit;
    btnSAPStock: TButton;
    leAreaOfStock: TLabeledEdit;
    btnAreaOfStock: TButton;
    procedure btnMB51Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnCPInAndStockClick(Sender: TObject);
    procedure btnSAPStockClick(Sender: TObject);
    procedure btnAreaOfStockClick(Sender: TObject);
  private
    { Private declarations }
    procedure LogEvent(const str: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmCPInAndStock.ShowForm;
var
  frmCPInAndStock: TfrmCPInAndStock;
begin
  frmCPInAndStock := TfrmCPInAndStock.Create(nil);
  try
    frmCPInAndStock.ShowModal;
  finally
    frmCPInAndStock.Free;
  end;
end;

procedure TfrmCPInAndStock.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  sdate: string;
  s: string;
begin
   ini := TIniFile.Create(AppIni);
  try
    leMB51.Text := ini.ReadString(self.ClassName, leMB51.Name, '');
    leCPInAndStock.Text := ini.ReadString(self.ClassName, leCPInAndStock.Name, '');
    sdate := ini.ReadString(self.ClassName, DateTimePicker1.Name, FormatDateTime('yyyy-MM-dd', Now));
    DateTimePicker1.DateTime := myStrToDateTime(sdate);

//    s := ini.ReadString(self.ClassName, mmoArea.Name, '');
//    mmoArea.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);

    s := ini.ReadString(self.ClassName, mmoProjNo2Name.Name, '');
    mmoProjNo2Name.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);

    leAreaOfStock.Text := ini.ReadString(self.ClassName, leAreaOfStock.Name, '');
    leSAPStock.Text := ini.ReadString(self.ClassName, leSAPStock.Name, '');

  finally
    ini.Free;
  end;
end;

procedure TfrmCPInAndStock.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leMB51.Name, leMB51.Text);
    ini.WriteString(self.ClassName, leCPInAndStock.Name, leCPInAndStock.Text);
    ini.WriteString(self.ClassName, DateTimePicker1.Name, FormatDateTime('yyyy-MM-dd', DateTimePicker1.DateTime));
            
//    s := StringReplace(mmoArea.Text, #13#10, '||', [rfReplaceAll]);
//    ini.WriteString(self.ClassName, mmoArea.Name, s);

    s := StringReplace(mmoProjNo2Name.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmoProjNo2Name.Name, s);

    ini.WriteString(self.ClassName, leAreaOfStock.Name, leAreaOfStock.Text);
    ini.WriteString(self.ClassName, leSAPStock.Name, leSAPStock.Text);

  finally
    ini.Free;
  end;
end;

procedure TfrmCPInAndStock.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmCPInAndStock.btnMB51Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMB51.Text := sfile;
end;
        
procedure TfrmCPInAndStock.btnCPInAndStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leCPInAndStock.Text := sfile;
end;

procedure TfrmCPInAndStock.btnAreaOfStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leAreaOfStock.Text := sfile;
end;
      
procedure TfrmCPInAndStock.btnSAPStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;

procedure TfrmCPInAndStock.LogEvent(const str: string);
begin
  Memo1.Lines.Add(str);
end;  

procedure TfrmCPInAndStock.btnSave2Click(Sender: TObject);
var
  sfile: string;
  aSAPMB51Reader2: TSAPMB51Reader2;
  aCPInAndStockReader: TCPInAndStockReader;
        
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  icol: Integer;
  aCPInAndStorkColHeaderPtr: PCPInAndStorkColHeader;

  idate: Integer;
  iproj: Integer;
  iline: Integer;
  aCPInAndStorkProj: TCPInAndStorkProj;
  aCPInAndStorkLine: TCPInAndStorkLine;
  aCPInAndStorkColPtr: PCPInAndStorkCol;
  irow1: Integer;
  i: Integer;

  aSAPStockReader2: TSAPStockReader2;

  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader;

  aSAPMB51RecordPtr: PSAPMB51Record;
  sno: string;
  idx: Integer; 
begin
  if not ExcelSaveDialog(sfile) then Exit;

  aCPInAndStockReader := TCPInAndStockReader.Create(leCPInAndStock.Text, LogEvent);

  aSAPMrpAreaStockReader := TSAPMrpAreaStockReader.Create(leAreaOfStock.Text, LogEvent);
  aSAPStockReader2 := TSAPStockReader2.Create(leSAPStock.Text, LogEvent);

  aSAPMB51Reader2 := TSAPMB51Reader2.Create(leMB51.Text, LogEvent);

  for i := 0 to aSAPMB51Reader2.Count - 1 do
  begin
    aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i];
    if (aSAPMB51RecordPtr^.smovingtype <> '101') and
      (aSAPMB51RecordPtr^.smovingtype <> '102') then Continue;

    sno := aSAPMB51RecordPtr^.snumber;
    sno := Copy(sno, 1, 5);
    idx := mmoProjNo2Name.Lines.IndexOfName(sno);
    if idx < 0 then Continue;
    aCPInAndStockReader.AddCPIn(aSAPMB51RecordPtr, mmoProjNo2Name.Lines.ValueFromIndex[idx]);
  end;

  
  aCPInAndStockReader.AddTodayStock(DateTimePicker1.DateTime, aSAPStockReader2,
    aSAPMrpAreaStockReader);

  try
           
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := True;
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
 
    try               
      Memo1.Lines.Add('write complete ');

      ExcelApp.Sheets[1].Activate;
                     
      irow := 1;
      ExcelApp.Cells[irow, 1].Value:= '入库';
      MergeCells(ExcelApp, 1, 1, 1, 5);
                    
      irow := 2;
      ExcelApp.Cells[irow, 1].Value:= '机型';
      ExcelApp.Cells[irow, 2].Value:= '料号';
      ExcelApp.Cells[irow, 3].Value:= '描述';
      ExcelApp.Cells[irow, 4].Value:= '当月累计入库数';
      ExcelApp.Cells[irow, 5].Value:= FormatDateTime('MM/dd', DateTimePicker1.DateTime) + '入库';
          
      ExcelApp.Columns[1].ColumnWidth := 6.88;
      ExcelApp.Columns[2].ColumnWidth := 18.75;
      ExcelApp.Columns[3].ColumnWidth := 55.38;
      ExcelApp.Columns[4].ColumnWidth := 14.50;
      ExcelApp.Columns[5].ColumnWidth := 11;

      icol := 6;
      for idate := 0 to aCPInAndStockReader.lstDate.Count -1  do
      begin
        aCPInAndStorkColHeaderPtr := PCPInAndStorkColHeader(aCPInAndStockReader.lstDate[idate]);
        ExcelApp.Cells[irow, icol].Value := aCPInAndStorkColHeaderPtr^.sml;   
        ExcelApp.Cells[irow, icol + 1].Value := aCPInAndStorkColHeaderPtr^.sfox;
        ExcelApp.Cells[irow, icol + 2].Value := aCPInAndStorkColHeaderPtr^.sdate;
        ExcelApp.Cells[irow, icol + 2].NumberFormatLocal := 'm"月"d"日"';
        icol :=icol + 3;
      end;

      irow := 3;
      for iproj := 0 to aCPInAndStockReader.Count - 1 do
      begin
        aCPInAndStorkProj := aCPInAndStockReader.Items[iproj];
        ExcelApp.Cells[irow, 1].Value := aCPInAndStorkProj.sname;

        irow1 := irow;
        for iline := 0 to aCPInAndStorkProj.Count -1  do
        begin
          aCPInAndStorkLine := aCPInAndStorkProj.Items[iline];
          ExcelApp.Cells[irow, 2].Value := aCPInAndStorkLine.snumber;
          ExcelApp.Cells[irow, 3].Value := aCPInAndStorkLine.sname;
          ExcelApp.Cells[irow, 4].Value := aCPInAndStorkLine.dCPInAcct;
          ExcelApp.Cells[irow, 5].Value := aCPInAndStorkLine.dCPInYesterday;
          icol := 6;
          for idate := 0 to aCPInAndStorkLine.Count - 1 do
          begin
            aCPInAndStorkColPtr := aCPInAndStorkLine.Items[idate];
            ExcelApp.Cells[irow, icol].Value := aCPInAndStorkColPtr^.dQtyML;
            ExcelApp.Cells[irow, icol + 1].Value := aCPInAndStorkColPtr^.dQtyFox;
            ExcelApp.Cells[irow, icol + 2].Value := '=' + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol + 1) + IntToStr(irow);
            icol := icol + 3;
          end;

          irow := irow + 1;
        end;                                     
        ExcelApp.Cells[irow, 2].Value := aCPInAndStorkProj.sname + '小计';
        icol := 6;
        for idate := 0 to aCPInAndStockReader.lstDate.Count - 1 do
        begin
          ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow - 1) + ')';
          ExcelApp.Cells[irow, icol + 1].Value := '=SUM(' + GetRef(icol + 1) + IntToStr(irow1) + ':' + GetRef(icol + 1) + IntToStr(irow - 1) + ')';
          ExcelApp.Cells[irow, icol + 2].Value := '=' + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol + 1) + IntToStr(irow);
          icol := icol + 3;
        end;


        MergeCells(ExcelApp, irow, 2, irow, 3);     // 小计
        MergeCells(ExcelApp, irow1, 1, irow, 1);    // Proj Name
        
        irow := irow + 1;
      end;

      ExcelApp.Range[ ExcelApp.Cells[3, 6], ExcelApp.Cells[irow - 1, 6 + aCPInAndStockReader.lstDate.Count * 3] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
      AddBorder(ExcelApp, 2, 1, irow - 1, 5 + aCPInAndStockReader.lstDate.Count * 3);
                                           
      Memo1.Lines.Add('save ');
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      Memo1.Lines.Add('quit excel ');
      WorkBook.Close;
      ExcelApp.Quit; 
    end;

  finally

    aCPInAndStockReader.Free;
    aSAPMrpAreaStockReader.Free;
    aSAPStockReader2.Free;
    aSAPMB51Reader2.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
  
end;

end.
