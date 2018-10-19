unit MPACompareWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, CommUtils, IniFiles, ImgList, ComCtrls, ComObj,
  ToolWin, DataIntAnalysisReader;

type
  TfrmMPACompare = class(TForm)
    leWeek1: TLabeledEdit;
    btnWeek1: TButton;
    leWeek2: TLabeledEdit;
    btnWeek2: TButton;
    ToolBar1: TToolBar;
    btnCompare: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    mmoWeeks: TMemo;
    Label1: TLabel;
    Memo1: TMemo;
    procedure btnWeek1Click(Sender: TObject);
    procedure btnWeek2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCompareClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations } 
  public
    { Public declarations }
    class procedure ShowForm; 
  end;
 
implementation

{$R *.dfm}


class procedure TfrmMPACompare.ShowForm;
var
  frmMPACompare: TfrmMPACompare;
begin
  frmMPACompare := TfrmMPACompare.Create(nil);
  try
    frmMPACompare.ShowModal;
  finally
    frmMPACompare.Free;
  end;
end;

procedure TfrmMPACompare.btnWeek1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leWeek1.Text := sfile;
end;

procedure TfrmMPACompare.btnWeek2Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leWeek2.Text := sfile;
end;

procedure TfrmMPACompare.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.Name, leWeek1.Name, leWeek1.Text);
    ini.WriteString(self.Name, leWeek2.Name, leWeek2.Text);
    ini.WriteString(self.Name, mmoWeeks.Name, StringReplace(mmoWeeks.Text, #13#10, '||', [rfReplaceAll] ) );
  finally
    ini.Free;
  end;
end;

procedure TfrmMPACompare.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leWeek1.Text := ini.ReadString(self.Name, leWeek1.Name, '');   
    leWeek2.Text := ini.ReadString(self.Name, leWeek2.Name, '');
    mmoWeeks.Text :=  StringReplace( ini.ReadString(self.Name, mmoWeeks.Name, ''), '||', #13#10, [rfReplaceAll]) ;
  finally
    ini.Free;
  end;
end;

procedure TfrmMPACompare.btnCompareClick(Sender: TObject);
var
  reader1: TDataIntAnalysisReader;
  reader2: TDataIntAnalysisReader;
  iline1: Integer;
  aDataIntAnalysisLine1: TDataIntAnalysisLine;
  aDataIntAnalysisLine2: TDataIntAnalysisLine;
  iweek: Integer;
  aDataIntAnalysisColPtr1: PDataIntAnalysisCol;
  aDataIntAnalysisColPtr2: PDataIntAnalysisCol;
  iDiffCount: Integer;
begin

  Memo1.Lines.Add('  ��ʼ��ȡ .............................................  ');
  Memo1.Lines.Add('  ======================================================  ');
  Memo1.Lines.Add('  ======================================================  ');
  Memo1.Lines.Add('  ======================================================  ');
  reader1 := TDataIntAnalysisReader.Create(leWeek1.Text, mmoWeeks.Lines);
  try
    reader2 := TDataIntAnalysisReader.Create(leWeek2.Text, mmoWeeks.Lines);
    try                    
      Memo1.Lines.Add('  ��ʼ�Ա� .............................................  ');
      iDiffCount := 0;
      Memo1.Lines.Add('KPI����-S&OP��Ӧ�ƻ� VS ���ۼƻ�');
      for iline1 := 0 to reader1.FListSOPvsDemand.Count - 1 do
      begin
        aDataIntAnalysisLine1 := TDataIntAnalysisLine(reader1.FListSOPvsDemand[iline1]);
        aDataIntAnalysisLine1.fcalc := True;

        aDataIntAnalysisLine2 := reader2.FindLine(reader2.FListSOPvsDemand, aDataIntAnalysisLine1);
        if aDataIntAnalysisLine2 = nil then
        begin
          iDiffCount := iDiffCount + 1;
          Memo1.Lines.Add('�� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' ��week2������');
          Continue;
        end;
        aDataIntAnalysisLine2.fcalc := True;

        for iweek := 0 to aDataIntAnalysisLine1.slweeks.Count - 1 do
        begin
          aDataIntAnalysisColPtr1 := PDataIntAnalysisCol(aDataIntAnalysisLine1.slweeks.Objects[iweek]);
          aDataIntAnalysisColPtr2 := aDataIntAnalysisLine2.FindWeek(aDataIntAnalysisColPtr1^.sweek);
          if aDataIntAnalysisColPtr2 = nil then
          begin
            iDiffCount := iDiffCount + 1;
            Memo1.Lines.Add('�� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr1^.icol) + ' ��week2������');
            Continue;
          end;
          aDataIntAnalysisColPtr2^.fcalc := True;

          if (aDataIntAnalysisColPtr1^.qty1 <> aDataIntAnalysisColPtr2^.qty1)
            or (aDataIntAnalysisColPtr1^.qty2 <> aDataIntAnalysisColPtr1^.qty2) then
          begin
            iDiffCount := iDiffCount + 1;
            Memo1.Lines.Add('week1 �� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr1^.icol) + ' ��ֵ ' + FloatToStr(aDataIntAnalysisColPtr1^.qty1) + ', ' + FloatToStr(aDataIntAnalysisColPtr1^.qty2) + '  ��week2�� '
              + IntToStr(aDataIntAnalysisLine2.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr2^.icol) + ' ��ֵ ' + FloatToStr(aDataIntAnalysisColPtr2^.qty1) + ', ' + FloatToStr(aDataIntAnalysisColPtr2^.qty2) + '  ����');
          end;
        end;
      end;
      if iDiffCount = 0 then
      begin
        Memo1.Lines.Add('KPI����-S&OP��Ӧ�ƻ� VS ���ۼƻ� week1 �� week2 �޲��� ');
      end
      else
      begin
        Memo1.Lines.Add('KPI����-S&OP��Ӧ�ƻ� VS ���ۼƻ� week1 �� week2 �� ' + IntToStr(iDiffCount) + ' ������ ');
      end;


      iDiffCount := 0;
      Memo1.Lines.Add('KPI����-ʵ�ʲ��� VS S&OP��Ӧ�ƻ�');
      for iline1 := 0 to reader1.FListACTvsDemand.Count - 1 do
      begin
        aDataIntAnalysisLine1 := TDataIntAnalysisLine(reader1.FListACTvsDemand[iline1]);
        aDataIntAnalysisLine1.fcalc := True;

        aDataIntAnalysisLine2 := reader2.FindLine(reader2.FListACTvsDemand, aDataIntAnalysisLine1);
        if aDataIntAnalysisLine2 = nil then
        begin
          iDiffCount := iDiffCount + 1;
          Memo1.Lines.Add('�� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' ��week2������');
          Continue;
        end;
        aDataIntAnalysisLine2.fcalc := True;

        for iweek := 0 to aDataIntAnalysisLine1.slweeks.Count - 1 do
        begin
          aDataIntAnalysisColPtr1 := PDataIntAnalysisCol(aDataIntAnalysisLine1.slweeks.Objects[iweek]);
          aDataIntAnalysisColPtr2 := aDataIntAnalysisLine2.FindWeek(aDataIntAnalysisColPtr1^.sweek);
          if aDataIntAnalysisColPtr2 = nil then
          begin
            iDiffCount := iDiffCount + 1;
            Memo1.Lines.Add('�� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr1^.icol) + ' ��week2������');
            Continue;
          end;
          aDataIntAnalysisColPtr2^.fcalc := True;

          if (aDataIntAnalysisColPtr1^.qty1 <> aDataIntAnalysisColPtr2^.qty1)
            or (aDataIntAnalysisColPtr1^.qty2 <> aDataIntAnalysisColPtr1^.qty2) then
          begin
            iDiffCount := iDiffCount + 1;
            Memo1.Lines.Add('week1 �� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr1^.icol) + ' ��ֵ ' + FloatToStr(aDataIntAnalysisColPtr1^.qty1) + ', ' + FloatToStr(aDataIntAnalysisColPtr1^.qty2) + '  ��week2�� '
              + IntToStr(aDataIntAnalysisLine2.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr2^.icol) + ' ��ֵ ' + FloatToStr(aDataIntAnalysisColPtr2^.qty1) + ', ' + FloatToStr(aDataIntAnalysisColPtr2^.qty2) + '  ����');
          end;
        end;
      end;
      if iDiffCount = 0 then
      begin
        Memo1.Lines.Add('KPI����-ʵ�ʲ��� VS S&OP��Ӧ�ƻ� week1 �� week2 �޲��� ');
      end
      else
      begin
        Memo1.Lines.Add('KPI����-ʵ�ʲ��� VS S&OP��Ӧ�ƻ� week1 �� week2 �� ' + IntToStr(iDiffCount) + ' ������ ');
      end;       


      iDiffCount := 0;
      Memo1.Lines.Add('KPI����-ʵ�ʲ��� VS �Ų��ƻ�');
      for iline1 := 0 to reader1.FListACTvsSch.Count - 1 do
      begin
        aDataIntAnalysisLine1 := TDataIntAnalysisLine(reader1.FListACTvsSch[iline1]);
        aDataIntAnalysisLine1.fcalc := True;

        aDataIntAnalysisLine2 := reader2.FindLine(reader2.FListACTvsSch, aDataIntAnalysisLine1);
        if aDataIntAnalysisLine2 = nil then
        begin
          iDiffCount := iDiffCount + 1;
          Memo1.Lines.Add('�� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' ��week2������');
          Continue;
        end;
        aDataIntAnalysisLine2.fcalc := True;

        for iweek := 0 to aDataIntAnalysisLine1.slweeks.Count - 1 do
        begin
          aDataIntAnalysisColPtr1 := PDataIntAnalysisCol(aDataIntAnalysisLine1.slweeks.Objects[iweek]);
          aDataIntAnalysisColPtr2 := aDataIntAnalysisLine2.FindWeek(aDataIntAnalysisColPtr1^.sweek);
          if aDataIntAnalysisColPtr2 = nil then
          begin
            iDiffCount := iDiffCount + 1;
            Memo1.Lines.Add('�� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr1^.icol) + ' ��week2������');
            Continue;
          end;
          aDataIntAnalysisColPtr2^.fcalc := True;

          if (aDataIntAnalysisColPtr1^.qty1 <> aDataIntAnalysisColPtr2^.qty1)
            or (aDataIntAnalysisColPtr1^.qty2 <> aDataIntAnalysisColPtr1^.qty2) then
          begin
            iDiffCount := iDiffCount + 1;
            Memo1.Lines.Add('week1 �� ' + IntToStr(aDataIntAnalysisLine1.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr1^.icol) + ' ��ֵ ' + FloatToStr(aDataIntAnalysisColPtr1^.qty1) + ', ' + FloatToStr(aDataIntAnalysisColPtr1^.qty2) + '  ��week2�� '
              + IntToStr(aDataIntAnalysisLine2.irow) + ' �� ' + GetRef(aDataIntAnalysisColPtr2^.icol) + ' ��ֵ ' + FloatToStr(aDataIntAnalysisColPtr2^.qty1) + ', ' + FloatToStr(aDataIntAnalysisColPtr2^.qty2) + '  ����');
          end;
        end;
      end;
      if iDiffCount = 0 then
      begin
        Memo1.Lines.Add('KPI����-ʵ�ʲ��� VS �Ų��ƻ� week1 �� week2 �޲��� ');
      end
      else
      begin
        Memo1.Lines.Add('KPI����-ʵ�ʲ��� VS �Ų��ƻ� week1 �� week2 �� ' + IntToStr(iDiffCount) + ' ������ ');
      end;  
    finally
      reader2.Free;
    end;
  finally
    reader1.Free;
  end;
end;
 
procedure TfrmMPACompare.btnExitClick(Sender: TObject);
begin
  Close;
end;

end.
