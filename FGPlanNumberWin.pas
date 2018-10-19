unit FGPlanNumberWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommVars, CommUtils,
  IniFiles, SAPMaterialReader, SAPMaterialReader2, SAPBomReader3, SAPBomReader,
  ComObj, ExcelConsts;

type
  TfrmFGPlanNumber = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    lemmlist: TLabeledEdit;
    btnmmlist: TButton;
    leBom: TLabeledEdit;
    btnBom: TButton;
    mmProjArea: TMemo;
    StatusBar1: TStatusBar;
    ProgressBar1: TProgressBar;
    procedure btnmmlistClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnBomClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmFGPlanNumber.ShowForm;
var
  frmFGPlanNumber: TfrmFGPlanNumber;
begin
  frmFGPlanNumber := TfrmFGPlanNumber.Create(nil);
  try
    frmFGPlanNumber.ShowModal;
  finally
    frmFGPlanNumber.Free;
  end;
end;  
  
procedure TfrmFGPlanNumber.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    lemmlist.Text := ini.ReadString(self.ClassName, lemmlist.Name, '');
    leBom.Text := ini.ReadString(self.ClassName, leBom.Name, '');
    s := ini.ReadString(self.ClassName, mmProjArea.Name, '');
    mmProjArea.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;

procedure TfrmFGPlanNumber.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, lemmlist.Name, lemmlist.Text);
    ini.WriteString(self.ClassName, leBom.Name, leBom.Text);
    s := StringReplace(mmProjArea.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmProjArea.Name, s);
  finally
    ini.Free;
  end;
end;

procedure TfrmFGPlanNumber.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmFGPlanNumber.btnmmlistClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  lemmlist.Text := sfile;
end;
       
procedure TfrmFGPlanNumber.btnBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leBom.Text := sfile;
end;

procedure TfrmFGPlanNumber.btnSave2Click(Sender: TObject);
type
  TPlanNumber = packed record
    snumber: string;
    sname: string;
    sMRPGroup: string;
    sMRPType: string;
    sMRPer: string;
    dLT_M0: Double;
    splannumber: string;
    sMMTypeDesc: string;
  end;
  PPlanNumber = ^TPlanNumber;
var
  aSAPBomReader3: TSAPBomReader3;
  aSAPMaterialReader2: TSAPMaterialReader2;
  iCount: Integer;
  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  aSapBom: TSapBom;
  aPlanNumberPtr: PPlanNumber;
  slPlanNumber: TStringList;
  splannumber: string;
  sfile: string;      
  ExcelApp, WorkBook: Variant;
  irow: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  sfile := ChangeFileExt(sfile, '.xls');

  StatusBar1.Panels[0].Text := '��ʼ��ȡ����...';
  
  aSAPBomReader3 := TSAPBomReader3.Create(leBom.Text);
  aSAPMaterialReader2 := TSAPMaterialReader2.Create(lemmlist.Text);

  StatusBar1.Panels[0].Text := 'Step 1...';

  slPlanNumber := TStringList.Create;
  try
    ProgressBar1.Max := aSAPMaterialReader2.Count;
    ProgressBar1.Position := 0;
    for iCount := 0 to aSAPMaterialReader2.Count - 1 do
    begin
      ProgressBar1.Position := iCount;

      aSAPMaterialRecordPtr := aSAPMaterialReader2.Items[iCount];

      if Pos('��������', aSAPMaterialRecordPtr.sGroupName) = 0 then Continue;
       
      if aSAPMaterialRecordPtr^.sPlanNumber <> '' then Continue;

      aSapBom := aSAPBomReader3.GetSapBom(aSAPMaterialRecordPtr^.sNumber, '');
      if aSapBom = nil then Continue;

      splannumber := aSapBom.GetPlanNumber;
      if splannumber = '' then Continue;

      if not IsNameHW(aSAPMaterialRecordPtr^.sNumber, aSAPMaterialRecordPtr^.sName) then Continue;

      aPlanNumberPtr := New(PPlanNumber);
      aPlanNumberPtr^.snumber := aSAPMaterialRecordPtr^.sNumber;
      aPlanNumberPtr^.sname := aSAPMaterialRecordPtr^.sName;
      aPlanNumberPtr^.sMRPGroup := aSAPMaterialRecordPtr^.sMRPGroup;
      aPlanNumberPtr^.sMRPType := aSAPMaterialRecordPtr^.sMRPType;
      aPlanNumberPtr^.sMRPer := aSAPMaterialRecordPtr^.sMRPer;
      aPlanNumberPtr^.dLT_M0 := aSAPMaterialRecordPtr^.dLT_M0;
      aPlanNumberPtr^.sMMTypeDesc := aSAPMaterialRecordPtr^.sMMTypeDesc;
      aPlanNumberPtr^.splannumber := splannumber;

      slPlanNumber.AddObject(aSAPMaterialRecordPtr^.sNumber, TObject(aPlanNumberPtr));
    end;
             
    StatusBar1.Panels[0].Text := 'Step 2...';

    ProgressBar1.Max := slPlanNumber.Count;
    
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

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;
    ExcelApp.Sheets[1].Name := '�����Ʒ�ƻ�����';

    try
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '���ϱ���';
      ExcelApp.Cells[irow, 2].Value := '����';
      ExcelApp.Cells[irow, 3].Value := 'MRP����';
      ExcelApp.Cells[irow, 4].Value := 'MRP��';
      ExcelApp.Cells[irow, 5].Value := 'MRP����';
      ExcelApp.Cells[irow, 6].Value := 'MRP������';
      ExcelApp.Cells[irow, 7].Value := '������С';
      ExcelApp.Cells[irow, 8].Value := '��С������С';
      ExcelApp.Cells[irow, 9].Value := '����ɹ���-����';
      ExcelApp.Cells[irow, 10].Value := '��������ʱ��';
      ExcelApp.Cells[irow, 11].Value := '���������С';
      ExcelApp.Cells[irow, 12].Value := '����ֵ';
      ExcelApp.Cells[irow, 13].Value := '����ɹ���-MRP����';
      ExcelApp.Cells[irow, 14].Value := 'ABC��ʶ';
      ExcelApp.Cells[irow, 15].Value := '�ƻ�����';
      ExcelApp.Cells[irow, 16].Value := '';
      ExcelApp.Cells[irow, 17].Value := '��Ʒ����';
      ExcelApp.Cells[irow, 18].Value := '������������';
      ExcelApp.Cells[irow, 19].Value := '����������';

      irow := irow + 1;
      
      for iCount := 0 to slPlanNumber.Count - 1 do
      begin
        ProgressBar1.Position := iCount;
        
        aPlanNumberPtr := PPlanNumber(slPlanNumber.Objects[iCount]);
        ExcelApp.Cells[irow, 1].Value := aPlanNumberPtr^.snumber;
        ExcelApp.Cells[irow, 2].Value := '1001';
        ExcelApp.Cells[irow, 3].Value := mmProjArea.Lines.Values[ Copy(aPlanNumberPtr^.snumber, 1, 5) ];
        ExcelApp.Cells[irow, 4].Value := aPlanNumberPtr^.sMRPGroup;
        ExcelApp.Cells[irow, 5].Value := aPlanNumberPtr^.sMRPType;
        ExcelApp.Cells[irow, 6].Value := aPlanNumberPtr^.sMRPer;
        ExcelApp.Cells[irow, 7].Value := 'EX'; //'������С';
        ExcelApp.Cells[irow, 8].Value := ''; //'��С������С';
        ExcelApp.Cells[irow, 9].Value := ''; //'����ɹ���-����';
        ExcelApp.Cells[irow, 10].Value := aPlanNumberPtr^.dLT_M0;
        ExcelApp.Cells[irow, 11].Value := ''; //'���������С';
        ExcelApp.Cells[irow, 12].Value := ''; //'����ֵ';
        ExcelApp.Cells[irow, 13].Value := ''; //'����ɹ���-MRP����';
        ExcelApp.Cells[irow, 14].Value := ''; //'ABC��ʶ';
        ExcelApp.Cells[irow, 15].Value := aPlanNumberPtr^.splannumber; // '�ƻ�����';
        ExcelApp.Cells[irow, 16].Value := '';
        ExcelApp.Cells[irow, 17].Value := aPlanNumberPtr^.sname;  // '��Ʒ����';
        ExcelApp.Cells[irow, 18].Value := aPlanNumberPtr^.sMMTypeDesc;  //'������������';
        ExcelApp.Cells[irow, 19].Value := '����������';
      end;

      ProgressBar1.Position := ProgressBar1.Max;
      
      WorkBook.SaveAs(sfile, xlExcel8);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end; 
        
  finally
    aSAPBomReader3.Free;
    aSAPMaterialReader2.Free;

    for iCount := 0 to slPlanNumber.Count - 1 do
    begin
      aPlanNumberPtr := PPlanNumber(slPlanNumber.Objects[iCount]);
      Dispose(aPlanNumberPtr);
    end;
    slPlanNumber.Free;
  end;          
  StatusBar1.Panels[0].Text := '���';

  MessageBox(Handle, '���', '��ʾ', 0);
end;

end.
