unit MRPAreaStockCheckWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ComCtrls, ToolWin, StdCtrls, ExtCtrls, CommUtils, IniFiles,
  SAPMrpAreaStockReader, StockListReader, ComObj;

type
  TfrmMRPAreaStockCheck = class(TForm)
    ToolBar1: TToolBar;
    tbSave: TToolButton;
    ImageList1: TImageList;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    leMRPAreaStockList: TLabeledEdit;
    leStockList: TLabeledEdit;
    btnMRPAreaStockList: TButton;
    btnStockList: TButton;
    mmAreas: TMemo;
    Label1: TLabel;
    procedure btnMRPAreaStockListClick(Sender: TObject);
    procedure btnStockListClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmMRPAreaStockCheck.ShowForm;
var
  frmMRPAreaStockCheck: TfrmMRPAreaStockCheck;
begin
  frmMRPAreaStockCheck := TfrmMRPAreaStockCheck.Create(nil);
  try
    frmMRPAreaStockCheck.ShowModal;
  finally
    frmMRPAreaStockCheck.Free;
  end;
end;  

procedure TfrmMRPAreaStockCheck.btnMRPAreaStockListClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMRPAreaStockList.Text := sfile;
end;

procedure TfrmMRPAreaStockCheck.btnStockListClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leStockList.Text := sfile;
end;

procedure TfrmMRPAreaStockCheck.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    leMRPAreaStockList.Text := ini.ReadString(self.ClassName, leMRPAreaStockList.Name, '');
    leStockList.Text := ini.ReadString(self.ClassName, leStockList.Name, '');
    s := ini.ReadString(self.ClassName, mmAreas.Name, '');
    s := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    mmAreas.Text := s; 
  finally
    ini.Free;
  end;
end;

procedure TfrmMRPAreaStockCheck.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leMRPAreaStockList.Name, leMRPAreaStockList.Text);
    ini.WriteString(self.ClassName, leStockList.Name, leStockList.Text);
    s := StringReplace(mmAreas.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmAreas.Name, s);
  finally
    ini.Free;
  end;
end;

procedure TfrmMRPAreaStockCheck.ToolButton2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmMRPAreaStockCheck.tbSaveClick(Sender: TObject);
const
  sNoMrpArea: array[0..2] of string = ('NOMRP', 'MZRD1', 'MZCS1');
  function IsNoMrpArea(const s: string): Boolean;
  var
    i: Integer;
  begin
    Result := False;
    for i := 0 to Length(sNoMrpArea) - 1 do
    begin
      if sNoMrpArea[i] = s then
      begin
        Result := True;
        Break;
      end;
    end;
  end;
var
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader;
  aStockListReader: TStockListReader;
  iStock: Integer;
  aStockInfoRecordPtr: PStockInfoRecord;
  
  sfile: string;        
  ExcelApp, WorkBook: Variant;
  irow: Integer; 
begin
  sfile := 'MRP区域对应仓库列表' + FormatDateTime('YYYYMMDDHHmmSS', Now) + '.xlsx';
  if not ExcelSaveDialog(sfile) then Exit;
 
  aSAPMrpAreaStockReader := TSAPMrpAreaStockReader.Create(leMRPAreaStockList.Text);
  aStockListReader := TStockListReader.Create(leStockList.Text);
  try
    for iStock := 0 to aStockListReader.Count - 1 do
    begin
      aStockInfoRecordPtr := aStockListReader.Items[iStock];
      aStockInfoRecordPtr^.sMrpArea := aSAPMrpAreaStockReader.MrpAreaOfStockNo(aStockInfoRecordPtr^.snumber);
      if aStockInfoRecordPtr^.sMrpArea = '' then
      begin
        aStockInfoRecordPtr^.sIsMrp := 'N';
      end
      else
      begin
        if IsNoMrpArea(aStockInfoRecordPtr^.sMrpArea) then
        begin
          aStockInfoRecordPtr^.sIsMrp := 'N';
        end
        else
        begin
          aStockInfoRecordPtr^.sIsMrp := 'Y';
        end;
      end;
    end;


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


   try                      
      ExcelApp.Sheets[1].Activate;
      ExcelApp.Sheets[1].Name := 'MRP区域对应仓库列表';
      
      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '工厂';   
      ExcelApp.Cells[irow, 2].Value := 'MRP区域';
      ExcelApp.Cells[irow, 3].Value := 'MRP区域描述';
      ExcelApp.Cells[irow, 4].Value := '仓库';
      ExcelApp.Cells[irow, 5].Value := '仓储描述';
      ExcelApp.Cells[irow, 6].Value := '是否参与MRP计算';

      irow := 2;
      for iStock := 0 to aStockListReader.Count - 1 do
      begin
        aStockInfoRecordPtr := aStockListReader.Items[iStock];
        ExcelApp.Cells[irow, 1].Value := '''' + aStockInfoRecordPtr^.sfac;
        ExcelApp.Cells[irow, 2].Value := aStockInfoRecordPtr^.sMrpArea;
        ExcelApp.Cells[irow, 3].Value := mmAreas.Lines.Values[aStockInfoRecordPtr^.sMrpArea];

        if StrToIntDef(aStockInfoRecordPtr^.snumber, -9999) <> -9999 then
        begin
          ExcelApp.Cells[irow, 4].Value := '''' + aStockInfoRecordPtr^.snumber;
        end
        else
        begin
          ExcelApp.Cells[irow, 4].Value := aStockInfoRecordPtr^.snumber;
        end;
        
        ExcelApp.Cells[irow, 5].Value := aStockInfoRecordPtr^.sname;
        ExcelApp.Cells[irow, 6].Value := aStockInfoRecordPtr^.sIsMrp;
        irow := irow + 1;
      end;

//
//
//      ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[1]);   
//      ExcelApp.Sheets[2].Activate;
//      ExcelApp.Sheets[2].Name := 'MRP区域列表';
//
//      irow := 1;
//      ExcelApp.Cells[irow, 1].Value := aStockInfoRecordPtr^.snumber;
//
//      irow := 2;
//      for iArea := 0 to slArea.Count - 1 do
//      begin
//      
//      end;
//

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end; 

  finally
    aSAPMrpAreaStockReader.Free;
    aStockListReader.Free;
 
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
