unit SAPBom2SBomWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, IniFiles,
  SAPBomReader, SAPStockReader;

type
  TfrmSAPBom2SBom = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leSAPBom: TLabeledEdit;
    btnSAPBom: TButton;
    leSAPStock: TLabeledEdit;
    btnSAPStock: TButton;
    procedure btnExitClick(Sender: TObject);
    procedure btnSAPBomClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnSAPStockClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

uses CommUtils;

{$R *.dfm}

class procedure TfrmSAPBom2SBom.ShowForm;
var
  frmSAPBom2SBom: TfrmSAPBom2SBom;
begin
  frmSAPBom2SBom := TfrmSAPBom2SBom.Create(nil);
  try
    frmSAPBom2SBom.ShowModal;
  finally
    frmSAPBom2SBom.Free;
  end;
end;

procedure TfrmSAPBom2SBom.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSAPBom2SBom.btnSAPBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;
         
procedure TfrmSAPBom2SBom.btnSAPStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;

procedure TfrmSAPBom2SBom.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leSAPBom.Text := ini.ReadString(self.ClassName, leSAPBom.Name, '');
    leSAPStock.Text := ini.ReadString(self.ClassName, leSAPStock.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSAPBom2SBom.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leSAPBom.Name, leSAPBom.Text);
    ini.WriteString(self.ClassName, leSAPStock.Name, leSAPStock.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmSAPBom2SBom.btnSave2Click(Sender: TObject);
var
  sfile: string;
  aSAPBomReader: TSAPBomReader;     
  aSAPStockReader: TSAPStockReader;
begin
  if not ExcelSaveDialog(sfile) then Exit;
                                         
  aSAPStockReader := TSAPStockReader.Create(leSAPStock.Text);
  aSAPBomReader := TSAPBomReader.Create(leSAPBom.Text);
  try
    aSAPBomReader.SaveSBom(sfile, aSAPStockReader);
  finally
    aSAPBomReader.Free;     
    aSAPStockReader.Free;
  end;
end;

end.
