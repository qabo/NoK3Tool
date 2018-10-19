unit WhereUseWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, CommUtils, StdCtrls, ComCtrls, ToolWin, ImgList,
  ExcelBomReader, NumberCheckReader, ComObj;

type
  TfrmWhereUse = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    GroupBox1: TGroupBox;
    mmoBoms: TMemo;
    btnBoms: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnBomsClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmWhereUse.ShowForm;
var
  frmWhereUse: TfrmWhereUse;
begin
  frmWhereUse := TfrmWhereUse.Create(nil);
  try
    frmWhereUse.ShowModal;
  finally
    frmWhereUse.Free;
  end;
end;

procedure TfrmWhereUse.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, mmoBoms.Name, '');
    mmoBoms.Text := StringReplace(Trim(s), '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;

procedure TfrmWhereUse.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := StringReplace(Trim(mmoBoms.Text), #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmoBoms.Name, s);
  finally
    ini.Free;
  end;
end;

procedure TfrmWhereUse.btnExitClick(Sender: TObject);
begin
  Close;
end;
   
procedure TfrmWhereUse.btnBomsClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  sfile := StringReplace(sfile, ';', #13#10, [rfReplaceAll]);
  mmoBoms.Lines.Add( sfile );
end;

procedure TfrmWhereUse.btnSave2Click(Sender: TObject);
type
  TWURecord = packed record
    snumber: string;
    sname: string;
    swhereuse: string;
  end;
  PWURecord = ^TWURecord;
  
var
  sfile: string;
  lstBom: TList;
  ibom: Integer;
  aExcelBomReader: TExcelBomReader;
  aBomLoc: TBomLoc;
  iBomLoc: Integer;
  slNumber: TStringList;
  aWURecordPtr: PWURecord;
  iBomLocLine: Integer;
  aBomLocLinePtr: PBomLocLine;
  inumber: Integer;
  idx: Integer;


  ExcelApp, WorkBook: Variant;
  irow: Integer;
  sfile_save: string;
begin
  if not ExcelSaveDialog(sfile_save) then Exit;

  lstBom := TList.Create;
  slNumber := TStringList.Create;

  try
    for ibom := 0 to mmoBoms.Lines.Count - 1 do
    begin
      sfile := Trim(mmoBoms.Lines[ibom]);
      if sfile = '' then Continue;

      aExcelBomReader := TExcelBomReader.Create(sfile);
      lstBom.Add(aExcelBomReader);
    end;




    for ibom := 0 to lstBom.Count - 1 do
    begin
      aExcelBomReader := TExcelBomReader(lstBom[ibom]);
      for iBomLoc := 0 to aExcelBomReader.FList.Count - 1 do
      begin
        aBomLoc := TBomLoc(aExcelBomReader.FList.Objects[iBomLoc]);
        for iBomLocLine := 0 to aBomLoc.FList.Count - 1 do
        begin
          aBomLocLinePtr := PBomLocLine(aBomLoc.FList.Objects[iBomLocLine]);
          idx := slNumber.IndexOf( aBomLocLinePtr^.snumber );
          if idx < 0 then
          begin
            aWURecordPtr := New(PWURecord);
            aWURecordPtr^.snumber := aBomLocLinePtr^.snumber;
            aWURecordPtr^.sname := aBomLocLinePtr^.sname;
            aWURecordPtr^.swhereuse := aExcelBomReader.sProj;
            slNumber.AddObject(aBomLocLinePtr^.snumber, TObject(aWURecordPtr));
          end
          else
          begin
            aWURecordPtr := PWURecord(slNumber.Objects[idx]);
            if Pos(aExcelBomReader.sProj, aWURecordPtr^.swhereuse) <= 0 then
            begin
              aWURecordPtr^.swhereuse := aWURecordPtr^.swhereuse + '/' + aExcelBomReader.sProj;
            end;
          end;
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

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '物料';
    ExcelApp.Cells[irow, 2].Value := '物料描述';    
    ExcelApp.Cells[irow, 3].Value := '使用项目';
    irow := irow + 1;
    for inumber := 0 to slNumber.Count - 1 do
    begin
      aWURecordPtr := PWURecord(slNumber.Objects[inumber]);
      ExcelApp.Cells[irow, 1].Value := aWURecordPtr^.snumber;
      ExcelApp.Cells[irow, 2].Value := aWURecordPtr^.sname;
      ExcelApp.Cells[irow, 3].Value := aWURecordPtr^.swhereuse;
      irow := irow + 1;
    end;  


    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存



    WorkBook.Close;
    ExcelApp.Quit;



    
    for ibom := 0 to lstBom.Count - 1 do
    begin
      aExcelBomReader := TExcelBomReader(lstBom[ibom]);
      aExcelBomReader.Free;
    end;

    for inumber := 0 to slNumber.Count - 1 do
    begin
      aWURecordPtr := PWURecord(slNumber.Objects[inumber]);
      Dispose(aWURecordPtr);
    end;
    slNumber.Clear;

  finally
    lstBom.Free;
    slNumber.Free;
  end;

  MessageBox(self.Handle, '完成', '提示', 0);
end;

end.
