unit KittingWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, IniFiles,
  CommUtils, KittingPPReader, StdCtrls, ExtCtrls, KittingKeyNumberReader;

type
  TfrmKitting = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leKittingPP: TLabeledEdit;
    btnKittingPP: TButton;
    Memo1: TMemo;
    leMKeyNumbers: TLabeledEdit;
    btnMKeyNumbers: TButton;
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnKittingPPClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure OnLogEvent(const s: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

{ TfrmKitting }

class procedure TfrmKitting.ShowForm;
var
  frmKitting: TfrmKitting;
begin
  frmKitting := TfrmKitting.Create(nil);
  frmKitting.ShowModal;
  frmKitting.Close;
end;

procedure TfrmKitting.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmKitting.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leKittingPP.Text := ini.ReadString(self.ClassName, leKittingPP.Name, '');
    leMKeyNumbers.Text := ini.ReadString(self.ClassName, leMKeyNumbers.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmKitting.FormClose(Sender: TObject; var Action: TCloseAction);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leKittingPP.Name, leKittingPP.Text);
    ini.WriteString(self.ClassName, leMKeyNumbers.Name, leMKeyNumbers.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmKitting.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

procedure TfrmKitting.btnKittingPPClick(Sender: TObject);
begin
  ExcelSaveDialogBtnClick(self, sender);    
end;

procedure TfrmKitting.btnSave2Click(Sender: TObject);
var
  aKittingPPReader: TKittingPPReader;
  aKittingKeyNumberReader: TKittingKeyNumberReader;
  slKittingNumber: TStringList;
  inumber: Integer;
  ikeynumber: Integer;
  aKittingNumber: TKittingNumber;
  aKittingPPSheet: TKittingPPSheet;
  isheet: Integer;
  aKittingKeyNumber: TKittingKeyNumber;
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  slKittingNumber := TStringList.Create;
  aKittingPPReader := TKittingPPReader.Create(leKittingPP.Text, OnLogEvent);
  aKittingKeyNumberReader := TKittingKeyNumberReader.Create(leMKeyNumbers.Text, OnLogEvent);
  try
    for isheet := 0 to aKittingPPReader.Count - 1 do
    begin
      aKittingPPSheet := aKittingPPReader.Sheets[isheet];
      OnLogEvent(aKittingPPSheet.FName);

      aKittingPPSheet.GenCombos(slKittingNumber);
      
      for inumber := 0 to slKittingNumber.Count - 1 do
      begin
        aKittingNumber := TKittingNumber(slKittingNumber.Objects[inumber]);
        OnLogEvent('aKittingNumber: ' + aKittingNumber.Name);
      end;

      OnLogEvent('');
    end;

                      
    for inumber := 0 to slKittingNumber.Count - 1 do
    begin
      aKittingNumber := TKittingNumber(slKittingNumber.Objects[inumber]);
      
      for ikeynumber := 0 to aKittingKeyNumberReader.Count - 1 do
      begin
        aKittingKeyNumber := aKittingKeyNumberReader.Items[ikeynumber];

        aKittingNumber.AddChildIfOk(aKittingKeyNumber);
        //ikeynumber

        OnLogEvent( Format('%s, %s, %s, %s, %s      %s, %s, %s, %0.0f, %0.0f', [
          aKittingKeyNumber.id,
          aKittingKeyNumber.sproj,
          aKittingKeyNumber.sname,
          aKittingKeyNumber.scat,
          aKittingKeyNumber.sver,
          aKittingKeyNumber.scap,
          aKittingKeyNumber.scolor,
          aKittingKeyNumber.snumber,
          aKittingKeyNumber.dlt,
          aKittingKeyNumber.dusage
         ]) );
      end;

    end;

    OnLogEvent('');
    
  finally                
    for inumber := 0 to slKittingNumber.Count - 1 do
    begin
      aKittingNumber := TKittingNumber(slKittingNumber.Objects[inumber]);
      aKittingNumber.Free;
    end;

    slKittingNumber.Clear;

    slKittingNumber.Free;
    aKittingPPReader.Free;
    aKittingKeyNumberReader.Free;
  end;
  
  MessageBox(Handle, '完成', '提示', 0);
end;

end.
