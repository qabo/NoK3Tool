unit MergeBomWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExcelUnit, CommVars, ComCtrls, ADODB, CommUtils,
  Grids, ValEdit, Menus, IniFiles, DB;

type
  TfrmMergeBom = class(TForm)
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    vleBomsFOX: TValueListEditor;
    vleBomsML: TValueListEditor;
    Button2: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    ADOQuery1: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations } 
    FType: string; 
    function CheckInput: Boolean;    
  public
    { Public declarations }
    class procedure ShowForm(const str: string);
  end;

implementation

uses MainWin;

{$R *.dfm}

class procedure TfrmMergeBom.ShowForm(const str: string);
var
  frmMergeBom: TfrmMergeBom;
begin
  frmMergeBom := TfrmMergeBom.Create(nil);  
  frmMergeBom.FType := str;
  frmMergeBom.Width := 866;
  frmMergeBom.Height := 470;
  frmMergeBom.Show;
end;  

procedure TfrmMergeBom.FormCreate(Sender: TObject);   
var
  ini: TIniFile;   
  sfile: string;
  skeysFOX, skeysML: TStringList;
  svalue: string;
  i: Integer;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);
  skeysFOX := TStringList.Create;
  skeysML := TStringList.Create;
  try
    ini.ReadSection('bomsFOX', skeysFOX);
    ini.ReadSection('bomsML', skeysML);
    if skeysFOX.Count = skeysML.Count then
    begin
      for i := 0 to skeysFOX.Count - 1 do
      begin
        svalue := ini.ReadString('bomsFOX', skeysFOX[i], '');
        vleBomsFOX.InsertRow(skeysFOX[i], svalue, True);
      end;

      ini.ReadSection('bomsML', skeysML);
      for i := 0 to skeysML.Count - 1 do
      begin
        svalue := ini.ReadString('bomsML', skeysML[i], '');
        vleBomsML.InsertRow(skeysML[i], svalue, True);
      end;
    end
    else
    begin
      ini.EraseSection('bomsFOX');  
      ini.EraseSection('bomsML');
    end;
  finally
    skeysFOX.Free;    
    skeysML.Free;
    ini.Free;
  end;
end;

procedure TfrmMergeBom.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  sfile: string;
  i: Integer;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);
  try
    ini.EraseSection('bomsFOX');
    for i := 1 to vleBomsFOX.RowCount - 1 do
    begin
      if vleBomsFOX.Cells[0, i] = EmptyStr then Break;
      ini.WriteString('bomsFOX', vleBomsFOX.Cells[0, i], vleBomsFOX.Cells[1, i]);
    end;

    ini.EraseSection('bomsML');
    for i := 1 to vleBomsML.RowCount - 1 do
    begin
      if vleBomsML.Cells[0, i] = EmptyStr then Break;
      ini.WriteString('bomsML', vleBomsML.Cells[0, i], vleBomsML.Cells[1, i]);
    end;
  finally
    ini.Free;
  end;
end;

procedure TfrmMergeBom.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;
 
function TfrmMergeBom.CheckInput: Boolean;
var
  i: Integer;
  svalue: string;
  ivalue: Integer;
begin
  Result := False;
  if vleBomsFOX.RowCount <= 1 then
  begin
    MessageBox(0, '请选择Bom', 'Ok', 0);
    Exit;
  end;
  if vleBomsFOX.Cells[0, 1] = EmptyStr then
  begin
    MessageBox(0, '请选择Bom', 'Ok', 0);
    Exit;
  end;
  for i := 1 to vleBomsFOX.RowCount - 1 do
  begin
    svalue := vleBomsFOX.Cells[1, i];
    ivalue := StrToIntDef(svalue, -9999);
    if (ivalue = -9999) and (svalue <> EmptyStr) then
    begin
      MessageBox(0, '生产数量请输入整数', 'Ok', 0);
      Exit;
    end;
  end;    
  for i := 1 to vleBomsML.RowCount - 1 do
  begin
    svalue := vleBomsML.Cells[1, i];
    ivalue := StrToIntDef(svalue, -9999);
    if (ivalue = -9999) and (svalue <> EmptyStr) then
    begin
      MessageBox(0, '生产数量请输入整数', 'Ok', 0);
      Exit;
    end;
  end;
  Result := True;
end;

procedure TfrmMergeBom.Button1Click(Sender: TObject);
var
  ifile: Integer;
  irow: Integer;
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options + [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  for ifile := 0 to OpenDialog1.Files.Count - 1 do
  begin
    if vleBomsFOX.FindRow(OpenDialog1.Files[ifile], irow) then
      Continue;
    vleBomsFOX.InsertRow(OpenDialog1.Files[ifile], EmptyStr, True);
    vleBomsML.InsertRow(OpenDialog1.Files[ifile], EmptyStr, True);
  end;
end;

procedure TfrmMergeBom.Button2Click(Sender: TObject);   
var
  aExcelBom1, aExcelBom2: TExcelBom;
  ifile: Integer;
  sModel: string;
  dwTick: DWORD;
begin
  if not CheckInput then Exit;

  sModel := ChangeFileExt(ExtractFileName(vleBomsFOX.Cells[0, 1]), '');
  for ifile := 2 to vleBomsFOX.RowCount - 1 do
  begin
    sModel := sModel + '&' + ChangeFileExt(ExtractFileName(vleBomsFOX.Cells[0, ifile]), '');
  end;
  if Length(sModel) > 200 then
    sModel := Copy(sModel, 1, 200);
  sModel := sModel + '料况表' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := sModel;
  if not SaveDialog1.Execute then Exit;

  dwTick := GetTickCount;
  
  frmmain.StatusBar1.Panels[0].Text := '正在生成料况表...';
  Application.ProcessMessages;

  try
    sModel := ChangeFileExt(ExtractFileName(vleBomsFOX.Cells[0, 1]), '');
    aExcelBom1 := TExcelBom.Create(sModel, FType, StrToIntDef(vleBomsFOX.Cells[1, 1], 0), StrToIntDef(vleBomsML.Cells[1, 1], 0));
    aExcelBom1.Open(vleBomsFOX.Cells[0, 1]);
    aExcelBom1.Close;
    for ifile := 2 to vleBomsFOX.RowCount - 1 do
    begin
      sModel := ChangeFileExt(ExtractFileName(vleBomsFOX.Cells[0, ifile]), '');
      aExcelBom2 := TExcelBom.Create(sModel, FType, StrToIntDef(vleBomsFOX.Cells[1, ifile], 0), StrToIntDef(vleBomsML.Cells[1, ifile], 0));
      aExcelBom2.Open(vleBomsFOX.Cells[0, ifile]);
      aExcelBom2.Close;
      aExcelBom1.MergeBom(aExcelBom2);
      aExcelBom2.Free;
    end;

    aExcelBom1.SaveAs(SaveDialog1.FileName);
    aExcelBom1.Free;
  finally 
    dwTick := GetTickCount - dwTick;
    frmmain.StatusBar1.Panels[0].Text := '生成料况表完成，耗用时间：' + IntToStr(dwTick div 1000) + ' 秒';
  end;
end;

procedure TfrmMergeBom.N1Click(Sender: TObject);
begin
  if (vleBomsFOX.Row > 0) and (vleBomsFOX.Row < vleBomsFOX.RowCount) then
  begin
    vleBomsFOX.DeleteRow(vleBomsFOX.Row);    
    vleBomsML.DeleteRow(vleBomsFOX.Row);
  end;
end;

end.
