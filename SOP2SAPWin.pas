unit SOP2SAPWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin, ComObj, DateUtils,
  Buttons, IniFiles, CommUtils;

type 
  TfrmSOP2SAP = class(TForm)
    leSOP: TLabeledEdit;
    btnSOP: TButton;
    OpenDialog1: TOpenDialog;
    ToolBar1: TToolBar;
    ImageList1: TImageList;
    SaveDialog1: TSaveDialog;
    tbSave: TToolButton;
    ToolButton1: TToolButton;
    mmoYearOfProj: TMemo;
    Label2: TLabel;
    dtpDate: TDateTimePicker;
    Label3: TLabel;
    tbQuit: TToolButton;
    GroupBox1: TGroupBox;
    Memo1: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure btnSOPClick(Sender: TObject);
    procedure tbQuitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

   
implementation

uses SOPReaderUnit;

{$R *.dfm}
            
class procedure TfrmSOP2SAP.ShowForm;
var
  frmSOP2SAP: TfrmSOP2SAP;
begin
  frmSOP2SAP := TfrmSOP2SAP.Create(nil);
  try
    frmSOP2SAP.ShowModal;
  finally
    frmSOP2SAP.Free;
  end;
end;
   
procedure TfrmSOP2SAP.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leSOP.Text := ini.ReadString(self.ClassName, leSOP.Name, '');
    dtpDate.DateTime := myStrToDateTime( ini.ReadString(self.ClassName, dtpDate.Name, '1900-01-01') );
    mmoYearOfProj.Text := StringReplace( ini.ReadString(self.ClassName, mmoYearOfProj.Name, ''), '|', #13#10, [rfReplaceAll] );
  finally
    ini.Free;
  end;
end;

procedure TfrmSOP2SAP.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leSOP.Name, leSOP.Text);
    ini.WriteString(self.ClassName, dtpDate.Name, FormatDateTime('yyyy-MM-dd', dtpDate.DateTime));
    ini.WriteString(self.ClassName, mmoYearOfProj.Name, StringReplace(mmoYearOfProj.Text, #13#10, '|', [rfReplaceAll]));
  finally
    ini.Free;
  end;
end; 
                
procedure TfrmSOP2SAP.btnSOPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOP.Text := sfile;
end;
  
procedure TfrmSOP2SAP.tbQuitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSOP2SAP.tbSaveClick(Sender: TObject);
var
  sfile: string;        
  ExcelApp, WorkBook: Variant;
  aSOPReader: TSOPReader;
  iProj: Integer;
  aSOPProj: TSOPProj;
  irow: Integer;
  iLine: Integer;
  aSOPLine: TSOPLine;
  iDate: Integer;
  aSOPCol: TSOPCol;
  slDate: TStringList;
  sdate: string;
  iQty: Integer; 
  idx: Integer;
  bBlank: Boolean;
  slNumberDetail: TStringList;
  iColDate1: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;


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
  //ExcelApp.Sheets[1].Name := '产品预测单';

  irow := 1;
  ExcelApp.Cells[irow, 1].Value := 'MATNR';
 
  ExcelApp.Columns[1].ColumnWidth := 16;

  slDate := TStringList.Create;
  slNumberDetail := TStringList.Create;
  try
    aSOPReader := TSOPReader.Create(TStringList( mmoYearOfProj.Lines ), leSOP.Text);

    iColDate1 := 2;
    if aSOPReader.HaveArea then
    begin
      ExcelApp.Cells[irow, 2].Value := 'BERID';
      iColDate1 := 3;
    end;

    try
      irow := 2;
      for iProj := 0 to aSOPReader.ProjCount - 1 do
      begin
        aSOPProj := aSOPReader.Projs[iProj];
        for iLine := 0 to aSOPProj.LineCount - 1 do
        begin
          bBlank := True;
          aSOPLine := aSOPProj.Lines[iLine];
          ExcelApp.Cells[irow, 1].Value := aSOPLine.sNumber;      
          if aSOPReader.HaveArea then
          begin
            ExcelApp.Cells[irow, 2].Value := aSOPLine.sMRPArea;
          end;

          for iDate := 0 to aSOPLine.DateCount - 1 do
          begin
            aSOPCol := aSOPLine.Dates[iDate];
            if aSOPCol.dt1 < dtpDate.DateTime then Continue;
            if aSOPCol.iQty > 0 then // 有数据
              bBlank := False;
            sdate := FormatDateTime('yyyy', aSOPCol.dt1) + Copy( IntToStr( 100 + WeekOf(aSOPCol.dt1)), 2, 2 );
            idx := slDate.IndexOf(sdate);
            if idx >= 0 then   // 周 已经 存在
            begin
              iQty := Integer(slDate.Objects[idx]);
              iQty := iQty + Round(aSOPCol.iQty);
              slDate.Objects[idx] := TObject(iQty);
            end
            else
            begin            
              iQty := Round(aSOPCol.iQty);
              slDate.AddObject(sdate, TObject(iQty));
            end; 
          end;

          if bBlank then
          begin
            ExcelApp.Cells[irow, 1].Value := '';
            if aSOPReader.HaveArea then
            begin
              ExcelApp.Cells[irow, 2].Value := '';
            end;
            Continue; // 无数据，跳过
          end;

          for iDate := 0 to slDate.Count - 1 do
          begin
            iQty := Integer(slDate.Objects[iDate]);
            if iQty > 0 then
            begin
              ExcelApp.Cells[irow, iDate + iColDate1].Value := iQty;
            end;

            slDate.Objects[iDate] := TObject(0);
          end;

          aSOPLine.sProj := aSOPProj.FName;
          slNumberDetail.AddObject(IntToStr(irow), aSOPLine);


          irow := irow + 1;
        end;
      end;

      irow := 1;
      for iDate := 0 to slDate.Count - 1 do
      begin
        ExcelApp.Cells[irow, iDate + iColDate1].Value := slDate[iDate];  
      end;

      irow := 1;
      ExcelApp.Cells[irow, slDate.Count + iColDate1 + 3].Value := '产品编码';
      ExcelApp.Cells[irow, slDate.Count + iColDate1 + 4].Value := '版本';
      ExcelApp.Cells[irow, slDate.Count + iColDate1 + 5].Value := '颜色';
      ExcelApp.Cells[irow, slDate.Count + iColDate1 + 6].Value := '容量';
      ExcelApp.Cells[irow, slDate.Count + iColDate1 + 7].Value := '项目';      
      for iLine := 0 to slNumberDetail.Count - 1 do
      begin
        aSOPLine := TSOPLine(slNumberDetail.Objects[iLine]);
        irow := StrToInt(slNumberDetail[iLine]);
        ExcelApp.Cells[irow, slDate.Count + iColDate1 + 3].Value := aSOPLine.sNumber;
        ExcelApp.Cells[irow, slDate.Count + iColDate1 + 4].Value := aSOPLine.sVer;
        ExcelApp.Cells[irow, slDate.Count + iColDate1 + 5].Value := aSOPLine.sColor;
        ExcelApp.Cells[irow, slDate.Count + iColDate1 + 6].Value := aSOPLine.sCap;
        ExcelApp.Cells[irow, slDate.Count + iColDate1 + 7].Value := aSOPLine.sProj;
      end;

    finally
      aSOPReader.Free;
    end;

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;

    slDate.Free;
    slNumberDetail.Free;
  end; 
  MessageBox(Handle, '完成', '提示', 0);
end;

end.


