unit LTP_CMS2MRPSimWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, CommUtils, StdCtrls, ExtCtrls, ComCtrls, ToolWin, ImgList, IniFiles,
  LTPCMSConfirmReader;

type
  TfrmLTP_CMS2MRPSim = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    btnCMSConfirm: TButton;
    leCMSConfirm: TLabeledEdit;
    procedure btnCMSConfirmClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmLTP_CMS2MRPSim.ShowForm;
var
  frmLTP_CMS2MRPSim: TfrmLTP_CMS2MRPSim;
begin
  frmLTP_CMS2MRPSim := TfrmLTP_CMS2MRPSim.Create(nil);
  try
    frmLTP_CMS2MRPSim.ShowModal;
  finally
    frmLTP_CMS2MRPSim.Free;
  end;
end;

procedure TfrmLTP_CMS2MRPSim.btnCMSConfirmClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leCMSConfirm.Text := sfile;
end;

procedure TfrmLTP_CMS2MRPSim.btnExitClick(Sender: TObject);
begin
  Close;
end;
   
procedure TfrmLTP_CMS2MRPSim.btnSave2Click(Sender: TObject);   
var
  aSAPBomReader: TTPCMSConfirmReader;
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  aSAPBomReader := TTPCMSConfirmReader.Create(leCMSConfirm.Text);
  try 
    aSAPBomReader.Save(nil, nil, sfile);
  finally
    aSAPBomReader.Free;
  end;
end;

procedure TfrmLTP_CMS2MRPSim.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leCMSConfirm.Text := ini.ReadString(self.ClassName, leCMSConfirm.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmLTP_CMS2MRPSim.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leCMSConfirm.Name, leCMSConfirm.Text);
  finally
    ini.Free;
  end;
end;

end.
