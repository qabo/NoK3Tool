unit FGDemandConfigWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  IniFiles, Dialogs, StdCtrls, ExtCtrls;

type  
  TFGDemandConfig = class
  public
    UpperLimit: Double;
    LowerLimit: Double;
    UpperBrush: TColor;
    LowerBrush: TColor;
    UpperFont: TColor;
    LowerFont: TColor;
    PromptHistoryChanged: Boolean;
    procedure LoadConfig(const ssection: string; const sfile: string);
    procedure SaveConfig(const ssection: string; const sfile: string);
  end;

  TfrmFGDemandConfig = class(TForm)
    leUpperLimit: TLabeledEdit;
    leLowerLimit: TLabeledEdit;
    btnCancel: TButton;
    btnOk: TButton;
    pnlUpperBrush: TPanel;
    pnlLowerBrush: TPanel;
    pnlUpperFont: TPanel;
    pnlLowerFont: TPanel;
    ColorDialog1: TColorDialog;
    FontDialog1: TFontDialog;
    procedure pnlUpperBrushClick(Sender: TObject);
    procedure pnlUpperFontClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class function ShowForm(aFGDemandConfig: TFGDemandConfig): Boolean;
  end;


implementation

{$R *.dfm}

{ TFGDemandConfig }

procedure TFGDemandConfig.LoadConfig(const ssection: string; const sfile: string);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(sfile);
  try
    UpperLimit := ini.ReadFloat(ssection, 'UpperLimit', 1.2);
    LowerLimit := ini.ReadFloat(ssection, 'LowerLimit', 0.8);  
    UpperBrush := ini.ReadInteger(ssection, 'UpperBrush', clRed);
    LowerBrush := ini.ReadInteger(ssection, 'LowerBrush', clYellow);
    UpperFont := ini.ReadInteger(ssection, 'UpperFont', clBlue);
    LowerFont := ini.ReadInteger(ssection, 'LowerFont', clYellow);   
    PromptHistoryChanged := ini.ReadBool(ssection, 'PromptHistoryChanged', True);
  finally
    ini.Free;
  end;
end;

procedure TFGDemandConfig.SaveConfig(const ssection: string; const sfile: string);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(sfile);
  try
    ini.WriteFloat(ssection, 'UpperLimit', UpperLimit);
    ini.WriteFloat(ssection, 'LowerLimit', LowerLimit); 
    ini.WriteInteger(ssection, 'UpperBrush', UpperBrush);
    ini.WriteInteger(ssection, 'LowerBrush', LowerBrush);
    ini.WriteInteger(ssection, 'UpperFont', UpperFont);
    ini.WriteInteger(ssection, 'LowerFont', LowerFont);
  finally
    ini.Free;
  end;
end;

class function TfrmFGDemandConfig.ShowForm(aFGDemandConfig: TFGDemandConfig): Boolean;
var
  frmFGDemandConfig: TfrmFGDemandConfig;
  mr: TModalResult;
begin
  Result := False;
  frmFGDemandConfig := TfrmFGDemandConfig.Create(nil);
  try 
    frmFGDemandConfig.leUpperLimit.Text := FloatToStr(aFGDemandConfig.UpperLimit);
    frmFGDemandConfig.leLowerLimit.Text := FloatToStr(aFGDemandConfig.LowerLimit);
    frmFGDemandConfig.pnlUpperBrush.Color := aFGDemandConfig.UpperBrush;
    frmFGDemandConfig.pnlUpperFont.Font.Color := aFGDemandConfig.UpperFont;
    frmFGDemandConfig.pnlLowerBrush.Color := aFGDemandConfig.LowerBrush;
    frmFGDemandConfig.pnlLowerFont.Font.Color := aFGDemandConfig.LowerFont;
    frmFGDemandConfig.Invalidate;
    mr := frmFGDemandConfig.ShowModal;
    if mr <> mrOk then Exit;
    Result := True;
    aFGDemandConfig.UpperLimit := StrToFloatDef(frmFGDemandConfig.leUpperLimit.Text, 0.2);   
    aFGDemandConfig.LowerLimit := StrToFloatDef(frmFGDemandConfig.leLowerLimit.Text, -0.2);
    aFGDemandConfig.UpperBrush := frmFGDemandConfig.pnlUpperBrush.Color;
    aFGDemandConfig.UpperFont := frmFGDemandConfig.pnlUpperFont.Font.Color;
    aFGDemandConfig.LowerBrush := frmFGDemandConfig.pnlLowerBrush.Color;
    aFGDemandConfig.LowerFont := frmFGDemandConfig.pnlLowerFont.Font.Color;
  finally
    frmFGDemandConfig.Free;
  end;
end;

procedure TfrmFGDemandConfig.pnlUpperBrushClick(Sender: TObject);
begin
  ColorDialog1.Color := TPanel(Sender).Color;
  if not ColorDialog1.Execute then Exit;
  TPanel(Sender).Color := ColorDialog1.Color;
end;

procedure TfrmFGDemandConfig.pnlUpperFontClick(Sender: TObject);
begin
  ColorDialog1.Color := TPanel(Sender).Font.Color;
  if not ColorDialog1.Execute then Exit;
  TPanel(Sender).Font.Color := ColorDialog1.Color;
end;

end.
