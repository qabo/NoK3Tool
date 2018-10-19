unit SOPVSActWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, CommUtils, ImgList, ComCtrls, ToolWin, IniFiles,
  ComObj, SOPReaderUnit, SOPVSActReaderUnit;

type
  TfrmSOPVSAct = class(TForm)
    leWeek: TLabeledEdit;
    Label1: TLabel;
    leSOP: TLabeledEdit;
    leMPS: TLabeledEdit;
    btnSOP: TButton;
    btnMPS: TButton;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    Memo1: TMemo;
    leSOPvsAct: TLabeledEdit;
    btnSOPvsAct: TButton;
    procedure btnSOPClick(Sender: TObject);
    procedure btnMPSClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSOPvsActClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmSOPVSAct.ShowForm;
var
  frmPlanVSAct: TfrmSOPVSAct;
begin
  frmPlanVSAct := TfrmSOPVSAct.Create(nil);
  try
    frmPlanVSAct.ShowModal;
  finally
    frmPlanVSAct.Free;
  end;
end;

procedure TfrmSOPVSAct.btnSOPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOP.Text := sfile;
end;

procedure TfrmSOPVSAct.btnMPSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMPS.Text := sfile;
end;
   
procedure TfrmSOPVSAct.btnSOPvsActClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOPvsAct.Text := sfile;
end;

procedure TfrmSOPVSAct.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSOPVSAct.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  leWeek.Text := ini.ReadString(Self.ClassName, leWeek.Name, '');     
  leSOP.Text := ini.ReadString(Self.ClassName, leSOP.Name, '');
  leMPS.Text := ini.ReadString(Self.ClassName, leMPS.Name, '');   
  leSOPvsAct.Text := ini.ReadString(Self.ClassName, leSOPvsAct.Name, '');
  ini.Free;
end;

procedure TfrmSOPVSAct.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  ini.WriteString(Self.ClassName, leWeek.Name, leWeek.Text);          
  ini.WriteString(Self.ClassName, leSOP.Name, leSOP.Text);
  ini.WriteString(Self.ClassName, leMPS.Name, leMPS.Text);
  ini.WriteString(Self.ClassName, leSOPvsAct.Name, leSOPvsAct.Text);
  ini.Free;
end;
  
procedure TfrmSOPVSAct.btnSave2Click(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;
//  iSheetCount, iSheet: Integer;
//  sSheet: string;

  aSOPReader_sop: TSOPReader;
  aSOPReader_mps: TSOPReader;

  irow: Integer;
  icol: Integer;
  icolPast: Integer;
  iProj: Integer;
  sfile: string;

  aSOPProj: TSOPProj;    
  aMPSProj: TSOPProj;
  aSOPLine: TSOPLine;
  aMPSLine: TSOPLine;
  aSOPCol: TSOPCol;
  aMPSCol: TSOPCol;
  iLine: Integer;
//  iMonth: Integer;
  iWeek: Integer;

//  slWeeks: TStringList;
  iWeekCount: Integer;
  irow1: Integer;

  aSOPVSActReader: TSOPVSActReader;
  aSOPVSActProj: TSOPVSActProj;
  bPass: Boolean;
  sl: TStringList;
begin
  if not ExcelSaveDialog(sfile) then Exit;

//  sfile := ExtractFilePath(Application.ExeName) + 'aa.xlsx';

  sl := TStringList.Create;

  aSOPReader_sop := TSOPReader.Create(sl, leSOP.Text);
  aSOPReader_mps := TSOPReader.Create(sl, leMPS.Text);

  sl.Free;

  aSOPVSActReader := TSOPVSActReader.Create(leSOPvsAct.Text);


  try


    // ��ʼ���� Excel
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(Handle, PChar(e.Message), '�����ʾ', 0);
        Exit;
      end;
    end;

    WorkBook := ExcelApp.WorkBooks.Add;
    while ExcelApp.Sheets.Count < aSOPReader_sop.FProjs.Count do
    begin
      ExcelApp.Sheets.Add;
    end;
        
    try
      try
        for iProj := 0 to aSOPReader_sop.FProjs.Count - 1 do
        begin
          aSOPProj := TSOPProj(aSOPReader_sop.FProjs.Objects[iProj]);

          aMPSProj := aSOPReader_mps.GetProj(aSOPProj.FName);
          if aMPSProj = nil then
          begin
            Memo1.Lines.Add('��Ŀ ' + aSOPProj.FName + ' �Ҳ��� MPS ');
            Continue;
          end;

          ExcelApp.Sheets[iProj + 1].Activate;
          ExcelApp.Sheets[iProj + 1].Name := aSOPProj.FName;

          icolPast := 9999;
                 
          irow := 1; 

          ExcelApp.Cells[irow, 1].Value := 'Week';
          MergeCells(ExcelApp, irow, 1, irow + 1, 1);
          ExcelApp.Cells[irow, 2].Value := '��ʽ';
          MergeCells(ExcelApp, irow, 2, irow + 1, 2);
          ExcelApp.Cells[irow, 3].Value := '���ϱ���';
          MergeCells(ExcelApp, irow, 3, irow + 1, 3);
          ExcelApp.Cells[irow, 4].Value := '��ɫ';
          MergeCells(ExcelApp, irow, 4, irow + 1, 4);
          ExcelApp.Cells[irow, 5].Value := '����';
          MergeCells(ExcelApp, irow, 5, irow + 1, 5);
          ExcelApp.Cells[irow, 6].Value := '����';
          MergeCells(ExcelApp, irow, 6, irow + 1, 6);

          iWeekCount := 0;
          if aSOPProj.FList.Count > 0 then
          begin
            aSOPLine := TSOPLine(aSOPProj.FList[0]);
            iWeekCount := aSOPLine.FList.Count;
            for iWeek := 0 to aSOPLine.FList.Count - 1 do
            begin
              aSOPCol := TSOPCol(aSOPLine.FList[iWeek]);

              ExcelApp.Cells[irow, iWeek + 7].Value := aSOPCol.sWeek;
              ExcelApp.Cells[irow + 1, iWeek + 7].Value := aSOPCol.sDate;

              if aSOPCol.sDate = leWeek.Text then
              begin
                icolPast := iWeek;
              end;
            end;
          end;
                   
          irow := 3;
          
          // д��ʷ ///////////////////////////////
          aSOPVSActProj := aSOPVSActReader.GetProj(aSOPProj.FName);
          if aSOPVSActProj <> nil then
          begin
            for iLine := 0 to aSOPVSActProj.FList.Count - 1 do
            begin
              aSOPLine := TSOPLine(aSOPVSActProj.FList[iLine]);


              ExcelApp.Cells[irow, 1].Value := aSOPLine.sDate;
              MergeCells(ExcelApp, irow, 1, irow + 2, 1);
              ExcelApp.Cells[irow, 2].Value := aSOPLine.sVer;
              MergeCells(ExcelApp, irow, 2, irow + 2, 2);
              ExcelApp.Cells[irow, 3].Value := aSOPLine.sNumber;
              MergeCells(ExcelApp, irow, 3, irow + 2, 3);
              ExcelApp.Cells[irow, 4].Value := aSOPLine.sColor;
              MergeCells(ExcelApp, irow, 4, irow + 2, 4);
              ExcelApp.Cells[irow, 5].Value := aSOPLine.sCap;
              MergeCells(ExcelApp, irow, 5, irow + 2, 5);    
              ExcelApp.Cells[irow, 6].Value := 'S&OP';
              ExcelApp.Cells[irow + 1, 6].Value := 'MPS';   
              ExcelApp.Cells[irow + 2, 6].Value := 'Delta';

                        
              bPass := True;

              icol := 7;
              for iWeek := 0 to aSOPLine.FList.Count - 1 do
              begin
                aSOPCol := TSOPCol(aSOPLine.FList[iWeek]);
                ExcelApp.Cells[irow, icol + iWeek].Value := aSOPCol.iQty_sop;   // ��һ�У� SOP�ƻ�
                ExcelApp.Cells[irow + 1, icol + iWeek].Value := aSOPCol.iQty_mps;   // ��һ�У� SOP�ƻ�   
                ExcelApp.Cells[irow + 2, icol + iWeek].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1);   // �����У� Delta

                if bPass then
                begin                                         
                  AddColor(ExcelApp, irow, 7 + iWeek, irow + 2, 7 + iWeek, $EAEAEA);
                end;
                
                if aSOPCol.sDate = aSOPLine.sDate then
                begin
                  bPass := False;
                end;
              end;

              irow := irow + 3;
            end;
          end;


          // д��ǰSOP//////////////////////////////////
                         
          irow1 := irow;
          for iLine := 0 to aSOPProj.FList.Count -1 do
          begin
            aSOPLine := TSOPLine(aSOPProj.FList[iLine]);

            aMPSLine := aMPSProj.GetLine(aSOPLine.sVer, aSOPLine.sNumber,
              aSOPLine.sColor, aSOPLine.sCap);
                                                           
            ExcelApp.Cells[irow, 1].Value := leWeek.Text;
            MergeCells(ExcelApp, irow, 1, irow + 2, 1);
            ExcelApp.Cells[irow, 2].Value := aSOPLine.sVer;
            MergeCells(ExcelApp, irow, 2, irow + 2, 2);
            ExcelApp.Cells[irow, 3].Value := aSOPLine.sNumber;
            MergeCells(ExcelApp, irow, 3, irow + 2, 3);
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sColor;
            MergeCells(ExcelApp, irow, 4, irow + 2, 4);
            ExcelApp.Cells[irow, 5].Value := aSOPLine.sCap;
            MergeCells(ExcelApp, irow, 5, irow + 2, 5);
            ExcelApp.Cells[irow, 6].Value := 'S&OP';
            ExcelApp.Cells[irow + 1, 6].Value := 'MPS';  
            ExcelApp.Cells[irow + 2, 6].Value := 'Delta';

            for iWeek := 0 to aSOPLine.FList.Count - 1 do
            begin
              aSOPCol := TSOPCol(aSOPLine.FList[iWeek]);

              icol := iWeek + 7;

              ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;   // ��һ�У� SOP�ƻ�
            
              if iWeek <= icolPast then  // ���У� ��ȥ���ܣ���ʵ�����Ա�
              begin
                if aMPSLine <> nil then
                begin
                  aMPSCol := aMPSLine.GetCol(aSOPCol.sDate);
                  if aMPSCol <> nil then
                  begin
                    ExcelApp.Cells[irow + 1, icol].Value := aMPSCol.iQty;   // ��һ�У� SOP�ƻ�
                  end;
                end;
              end
              else                      // ���У� δ�����ܣ� ��SOP�ƻ����Ա�
              begin
                ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.iQty;   // ��һ�У� SOP�ƻ�
              end;    
              ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1);   // �����У� Delta
            end;
          
            irow := irow + 3;
          end;

          if icolPast + 7 <= iWeekCount + 6 then
          begin
            AddColor(ExcelApp, irow1, 7, irow - 1, icolPast + 7, $EAEAEA);
          end
          else
          begin
            AddColor(ExcelApp, irow1, 7, irow - 1, iWeekCount + 6, $EAEAEA);
          end;

          AddBorder(ExcelApp, 1, 1, irow - 1, iWeekCount + 6);
        end;
      except
        on e: Exception do
        begin
          raise e;
        end;
      end;
                         
      ExcelApp.Sheets[1].Activate;

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end; 

  finally
    aSOPReader_sop.Free;
    aSOPReader_mps.Free;
    aSOPVSActReader.Free;
  end;

  MessageBox(Handle, '���', '��ʾ', 0);

end;

end.
