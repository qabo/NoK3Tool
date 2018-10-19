unit ProjNameNoWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, StdCtrls, CommUtils;

type
  TfrmProjNameNo = class(TForm)
    mmoOEM: TMemo;
    btnCancel: TButton;
    btnOk: TButton;
    mmoODM: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    mmoIgnoreNo: TMemo;
    Label3: TLabel;
    mmoIgnoreName4Sum: TMemo;
    Label4: TLabel;
    procedure btnOkClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
    class function GetProjNos_OEM: TStringList;
    class function GetProjNos_ODM: TStringList;   
    class function GetIgnoreNos: TStringList;
    class function GetIgnoreName4Sum: TStringList;
  end;


implementation

{$R *.dfm}

class procedure TfrmProjNameNo.ShowForm;
var
  frmProjYear: TfrmProjNameNo;
begin
  frmProjYear := TfrmProjNameNo.Create(nil);
  try
    frmProjYear.ShowModal;
  finally
    frmProjYear.Free;
  end;
end;

class function TfrmProjNameNo.GetProjNos_OEM: TStringList;
var
  ini: TIniFile;
begin
  Result := TStringList.Create;
  ini := TIniFile.Create(AppIni);
  try
    ini.ReadSectionValues(self.ClassName + '-OEM', Result);
  finally
    ini.Free;
  end;
end;

class function TfrmProjNameNo.GetProjNos_ODM: TStringList;
var
  ini: TIniFile;
begin
  Result := TStringList.Create;
  ini := TIniFile.Create(AppIni);
  try
    ini.ReadSectionValues(self.ClassName + '-ODM', Result);
  finally
    ini.Free;
  end;
end;

class function TfrmProjNameNo.GetIgnoreNos: TStringList;
var
  ini: TIniFile;
  s: string;
begin
  Result := TStringList.Create;
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, 'mmoIgnoreNo', '');
    Result.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;  

class function TfrmProjNameNo.GetIgnoreName4Sum: TStringList;
var
  ini: TIniFile;
  s: string;
begin
  Result := TStringList.Create;
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, 'mmoIgnoreName4Sum', '');
    Result.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;  

procedure TfrmProjNameNo.btnOkClick(Sender: TObject);
var
  ini: TIniFile;
  iline: Integer;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.EraseSection(self.ClassName + '-OEM');
    for iline := 0 to mmoOEM.Lines.Count - 1 do
    begin
      ini.WriteString(self.ClassName + '-OEM', mmoOEM.Lines.Names[iline], mmoOEM.Lines.ValueFromIndex[iline]);
    end;

    ini.EraseSection(self.ClassName + '-ODM');
    for iline := 0 to mmoODM.Lines.Count - 1 do
    begin
      ini.WriteString(self.ClassName + '-ODM', mmoODM.Lines.Names[iline], mmoODM.Lines.ValueFromIndex[iline]);
    end;

    ini.WriteString(self.ClassName, mmoIgnoreNo.Name,  StringReplace(mmoIgnoreNo.Text, #13#10, '||', [rfReplaceAll])  );
    ini.WriteString(self.ClassName, mmoIgnoreName4Sum.Name,  StringReplace(mmoIgnoreName4Sum.Text, #13#10, '||', [rfReplaceAll])  );

  finally
    ini.Free;
  end;
  Close;
end;

procedure TfrmProjNameNo.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmProjNameNo.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  mmoOEM.Clear;
  mmoODM.Clear;
  ini := TIniFile.Create(AppIni);
  try
    ini.ReadSectionValues(Self.ClassName + '-OEM', mmoOEM.Lines);
    ini.ReadSectionValues(Self.ClassName + '-ODM', mmoODM.Lines);
    mmoIgnoreNo.Text := StringReplace( ini.ReadString(self.ClassName, mmoIgnoreNo.Name, ''), '||', #13#10, [rfReplaceAll] );
    mmoIgnoreName4Sum.Text := StringReplace( ini.ReadString(self.ClassName, mmoIgnoreName4Sum.Name, ''), '||', #13#10, [rfReplaceAll] );
  finally
    ini.Free;
  end;


end;

end.
