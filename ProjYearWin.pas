unit ProjYearWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, StdCtrls, CommUtils;

type
  TfrmProjYear = class(TForm)
    Memo1: TMemo;
    GroupBox1: TGroupBox;
    btnOk: TButton;
    btnCancel: TButton;
    procedure btnOkClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
    class function GetProjYears: TStringList;
  end;


implementation

{$R *.dfm}

class procedure TfrmProjYear.ShowForm;
var
  frmProjYear: TfrmProjYear;
begin
  frmProjYear := TfrmProjYear.Create(nil);
  try
    frmProjYear.ShowModal;
  finally
    frmProjYear.Free;
  end;
end;

class function TfrmProjYear.GetProjYears: TStringList;
var
  ini: TIniFile;
begin
  Result := TStringList.Create;
  ini := TIniFile.Create(AppIni);
  try
    ini.ReadSectionValues('frmProjYear', Result);
  finally
    ini.Free;
  end; 
end;

procedure TfrmProjYear.btnOkClick(Sender: TObject);
var
  ini: TIniFile;
  iline: Integer;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.EraseSection(self.Name);
    for iline := 0 to Memo1.Lines.Count - 1 do
    begin
      ini.WriteString(self.Name, Memo1.Lines.Names[iline], Memo1.Lines.ValueFromIndex[iline]);
    end;
  finally
    ini.Free;
  end;
  Close;
end;

procedure TfrmProjYear.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmProjYear.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  Memo1.Clear;
  ini := TIniFile.Create(AppIni);
  try
    ini.ReadSectionValues(Self.Name, Memo1.Lines);
  finally
    ini.Free;
  end; 
end;

end.
