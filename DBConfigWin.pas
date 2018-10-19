unit DBConfigWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, CommUtils, IniFiles;

type
  TfrmDBConfig = class(TForm)
    leServer: TLabeledEdit;
    leUser: TLabeledEdit;
    lePwd: TLabeledEdit;
    btnOk: TButton;
    btnCancel: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
    class procedure Load;
  end;


implementation

{$R *.dfm}

class procedure TfrmDBConfig.ShowForm;
var
  frmDBConfig: TfrmDBConfig;
begin
  frmDBConfig := TfrmDBConfig.Create(nil);
  try
    with frmDBConfig do
    begin
      leServer.Text := gserver;
      leUser.Text := guser;
      lePwd.Text := gpwd;
    end;
    frmDBConfig.ShowModal;
  finally
    frmDBConfig.Free;
  end;
end;

class procedure TfrmDBConfig.Load;
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    gserver := ini.ReadString('options', 'server', '127.0.0.1');
    guser := ini.ReadString('options', 'user', 'sa');
    gpwd := ini.ReadString('options', 'pwd', 'Pmc010601');
  finally
    ini.Free;
  end;
end;

procedure TfrmDBConfig.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leServer.Text := ini.ReadString('options', 'server', '127.0.0.1');
    leUser.Text := ini.ReadString('options', 'user', 'sa');
    lePwd.Text := ini.ReadString('options', 'pwd', 'Pmc010601');
  finally
    ini.Free;
  end;
end;

procedure TfrmDBConfig.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);         
var
  ini: TIniFile;
begin
  if ModalResult <> mrOk then Exit;
  ini := TIniFile.Create(AppIni);
  try
    gserver := leServer.Text;
    guser := leUser.Text;
    gpwd := lePwd.Text;
    ini.WriteString('options', 'server', gserver);
    ini.WriteString('options', 'user', guser);
    ini.WriteString('options', 'pwd', gpwd);
  finally
    ini.Free;
  end;
end;

end.
