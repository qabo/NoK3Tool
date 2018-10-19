unit PCNumberWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommVars, CommUtils,
  IniFiles, SAPMaterialReader, SAPMaterialReader2, SAPBomReader3, SAPBomReader,
  ComObj, ExcelConsts, ZMDR001Reader;

type
  TfrmPCNumber = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    mmProjArea: TMemo;
    StatusBar1: TStatusBar;
    ProgressBar1: TProgressBar;
    leZMDR001: TLabeledEdit;
    btnZMDR001: TButton;
    Label1: TLabel;
    mmProjNo: TMemo;
    Label2: TLabel;
    mmProjMrper: TMemo;
    Label3: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnZMDR001Click(Sender: TObject);
  private
    { Private declarations }
    function GetProjArea(aZMDR001RecordPtr: PZMDR001Record): string;
    function GetMrpGroup(aZMDR001RecordPtr: PZMDR001Record): string;
    function GetProjMrper(aZMDR001RecordPtr: PZMDR001Record): string;
    function GetProjLT(aZMDR001RecordPtr: PZMDR001Record): string;
    function GetABC(const sMRPGroup: string; aZMDR001RecordPtr: PZMDR001Record): string;
    function ProjName2No(const sprojname: string): string;
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmPCNumber.ShowForm;
var
  frmFGPlanNumber: TfrmPCNumber;
begin
  frmFGPlanNumber := TfrmPCNumber.Create(nil);
  try
    frmFGPlanNumber.ShowModal;
  finally
    frmFGPlanNumber.Free;
  end;
end;  
  
procedure TfrmPCNumber.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, mmProjArea.Name, '');
    mmProjArea.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    leZMDR001.Text := ini.ReadString(self.ClassName, leZMDR001.Name, '');
    s := ini.ReadString(self.ClassName, mmProjNo.Name, '');
    mmProjNo.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    s := ini.ReadString(self.ClassName, mmProjMrper.Name, '');
    mmProjMrper.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;

procedure TfrmPCNumber.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := StringReplace(mmProjArea.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmProjArea.Name, s);
    s := StringReplace(mmProjNo.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmProjNo.Name, s);
    s := StringReplace(mmProjMrper.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmProjMrper.Name, s);
    ini.WriteString(self.ClassName, leZMDR001.Name, leZMDR001.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmPCNumber.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmPCNumber.btnZMDR001Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leZMDR001.Text := sfile;
end;
        
function TfrmPCNumber.ProjName2No(const sprojname: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to mmProjNo.Lines.Count - 1 do
  begin
    if mmProjNo.Lines.ValueFromIndex[i] = sprojname then
    begin
      Result := mmProjNo.Lines.Names[i];
      Break;
    end;
  end;
end;

function TfrmPCNumber.GetProjArea(aZMDR001RecordPtr: PZMDR001Record): string;
var
  sprojno: string;
  sprojname: string;
  idx: Integer;
begin
  Result := '';
  
  if aZMDR001RecordPtr^.sCategory = '�ƻ�����' then
  begin
    sprojno := '03.' + Copy(aZMDR001RecordPtr^.sNumber, 4, 2);
    if mmProjArea.Lines.IndexOfName(sprojno) >= 0 then
    begin
      Result := mmProjArea.Lines.Values[sprojno];
      Exit;
    end;
    
    sprojno := '83.' + Copy(aZMDR001RecordPtr^.sNumber, 4, 2);
    if mmProjArea.Lines.IndexOfName(sprojno) >= 0 then
    begin
      Result := mmProjArea.Lines.Values[sprojno];
      Exit;
    end;

    Exit;
  end;
            
  if (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') or (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') then
  begin
    sprojno := Copy(aZMDR001RecordPtr^.sNumber, 1, 5);
    if mmProjArea.Lines.IndexOfName(sprojno) >= 0 then
    begin
      Result := mmProjArea.Lines.Values[sprojno];
      Exit;
    end;

    Exit;
  end;

  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    idx := Pos('(', aZMDR001RecordPtr^.sGroupName);
    if idx > 0 then
    begin
      sprojname := Copy(aZMDR001RecordPtr^.sGroupName, 1, idx - 1);
      sprojno := ProjName2No(sprojname);
      if sprojno <> '' then
      begin
        Result := mmProjArea.Lines.Values[sprojno];
        Exit;
      end;
    end;

    sprojname := Copy(aZMDR001RecordPtr^.sName, 1, 5);
    sprojno := ProjName2No(sprojname);
    idx := mmProjArea.Lines.IndexOfName(sprojno);
    if idx >= 0 then
    begin
      Result := mmProjArea.Lines.ValueFromIndex[idx];
      Exit;
    end;

    Exit;
  end;
 
  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    idx := Pos('(', aZMDR001RecordPtr^.sGroupName);
    if idx > 0 then
    begin
      sprojname := Copy(aZMDR001RecordPtr^.sGroupName, 1, idx - 1);
      sprojno := ProjName2No(sprojname);
      if sprojno <> '' then
      begin
        Result := mmProjArea.Lines.Values[sprojno];
        Exit;
      end;
    end;   

    sprojname := Copy(aZMDR001RecordPtr^.sName, 1, 5);
    sprojno := ProjName2No(sprojname);
    idx := mmProjArea.Lines.IndexOfName(sprojno);
    if idx >= 0 then
    begin
      Result := mmProjArea.Lines.ValueFromIndex[idx];
      Exit;
    end;

    Exit;
  end;   
end;  

function TfrmPCNumber.GetMrpGroup(aZMDR001RecordPtr: PZMDR001Record): string; 
begin
  Result := '';
  
  if aZMDR001RecordPtr^.sCategory = '�ƻ�����' then
  begin
    Result := 'Z001';
    Exit;
  end;
  
  if aZMDR001RecordPtr^.sGroupName = 'ί����Ʒ�ػ�' then
  begin
    Result := 'Z001';
    Exit;
  end;
            
  if (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') or (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') then
  begin          
    if Pos('�ۺ����', aZMDR001RecordPtr^.sGroupName) > 0 then
    begin
      Result := 'Z001';
    end
    else if Pos('�������', aZMDR001RecordPtr^.sGroupName) > 0 then
    begin
      Result := 'Z002';
    end
    else
    begin
      Result := 'Z001';
    end;

    Exit;
  end;

  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    if (Pos('�ۺ���Ʒ', aZMDR001RecordPtr^.sGroupName) > 0) and
      (Pos('(����)', aZMDR001RecordPtr^.sGroupName) > 0) then
    begin
      Result := 'Z001';
    end
    else
    begin
      Result := 'Z002';
    end;
    Exit;
  end;

  // ���а��Ʒ
  Result := 'Z002';
end;

function TfrmPCNumber.GetProjMrper(aZMDR001RecordPtr: PZMDR001Record): string; 
var
  sprojno: string;
  sprojname: string;
  idx: Integer;
begin
  Result := '';
  
  if aZMDR001RecordPtr^.sCategory = '�ƻ�����' then
  begin
    sprojno := '03.' + Copy(aZMDR001RecordPtr^.sNumber, 4, 2);
    if mmProjNo.Lines.IndexOfName(sprojno) >= 0 then
    begin
      sprojname := mmProjNo.Lines.Values[sprojno];
      Result := mmProjMrper.Lines.Values[sprojname];
      Exit;
    end;
    
    sprojno := '83.' + Copy(aZMDR001RecordPtr^.sNumber, 4, 2);
    if mmProjNo.Lines.IndexOfName(sprojno) >= 0 then
    begin
      sprojname := mmProjNo.Lines.Values[sprojno];    
      Result := mmProjMrper.Lines.Values[sprojname];
      Exit;
    end;

    Exit;
  end;
            
  if (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') or (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') then
  begin
    sprojno := Copy(aZMDR001RecordPtr^.sNumber, 1, 5);
    if mmProjNo.Lines.IndexOfName(sprojno) >= 0 then
    begin
      sprojname := mmProjNo.Lines.Values[sprojno];    
      Result := mmProjMrper.Lines.Values[sprojname];
      Exit;
    end;

    Exit;
  end;


  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    idx := Pos('(', aZMDR001RecordPtr^.sGroupName);
    if idx > 0 then
    begin
      sprojname := Copy(aZMDR001RecordPtr^.sGroupName, 1, idx - 1);
      Result := mmProjMrper.Lines.Values[sprojname];
      if Result <> '' then Exit;
    end;

    sprojname := Copy(aZMDR001RecordPtr^.sName, 1, 5);
    idx := mmProjMrper.Lines.IndexOfName(sprojname);
    if idx >= 0 then
    begin
      Result := mmProjMrper.Lines.ValueFromIndex[idx];
      Exit;
    end;
    
    Exit;
  end;


  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    idx := Pos('(', aZMDR001RecordPtr^.sGroupName);
    if idx > 0 then
    begin
      sprojname := Copy(aZMDR001RecordPtr^.sGroupName, 1, idx - 1);
      Result := mmProjMrper.Lines.Values[sprojname];
      if Result <> '' then Exit;
    end;

    sprojname := Copy(aZMDR001RecordPtr^.sName, 1, 5);
    idx := mmProjMrper.Lines.IndexOfName(sprojname);
    if idx >= 0 then
    begin
      Result := mmProjMrper.Lines.ValueFromIndex[idx];
      Exit;
    end;
    
    Exit;
  end;  


end;

function TfrmPCNumber.GetProjLT(aZMDR001RecordPtr: PZMDR001Record): string; 
begin
  Result := '';
  
  if aZMDR001RecordPtr^.sCategory = '�ƻ�����' then
  begin
    Result := '2';
    Exit;
  end;
            
  if (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') then
  begin
    if Pos('�������', aZMDR001RecordPtr^.sGroupName) > 0 then
    begin
      Result := '3';
    end
    else
    begin
      Result := '2';
    end;
  end;

  if (aZMDR001RecordPtr^.sCategory = '���г�Ʒ') then
  begin
    if Pos('�������', aZMDR001RecordPtr^.sGroupName) > 0 then
    begin
      Result := '4';
    end   
    else if Pos('�Բ����', aZMDR001RecordPtr^.sGroupName) > 0 then
    begin
      Result := '4';
    end
    else
    begin
      Result := '2';
    end;

    Exit;
  end;


  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    if Pos('(PCBA)', aZMDR001RecordPtr^.sGroupName) > 0 then
    begin
      Result := '3';
      Exit;
    end;

    if Pos('���', aZMDR001RecordPtr^.sName) > 0 then
    begin
      Result := '1';
      Exit;
    end;

    if Pos('���', aZMDR001RecordPtr^.sName) > 0 then
    begin
      Result := '1';
      Exit;
    end;

    if Pos('����/', aZMDR001RecordPtr^.sName) > 0 then
    begin
      Result := '3';
      Exit;
    end;

    Exit;
  end;


  if aZMDR001RecordPtr^.sCategory = '���а��Ʒ' then
  begin
    if Pos('(PCBA)', aZMDR001RecordPtr^.sGroupName) > 0  then
    begin
      Result := '5';
    end;
    
    if Pos('���', aZMDR001RecordPtr^.sName) > 0 then
    begin
      Result := '1';
      Exit;
    end;
           
    if Pos('���', aZMDR001RecordPtr^.sName) > 0 then
    begin
      Result := '1';
      Exit;
    end;

    if Pos('����/', aZMDR001RecordPtr^.sName) > 0 then
    begin
      Result := '5';
      Exit;
    end;

    Exit;
  end;  


end;

function TfrmPCNumber.GetABC(const sMRPGroup: string;
  aZMDR001RecordPtr: PZMDR001Record): string;
begin
  if (sMRPGroup = 'Z002') and (Pos('�ۺ�', aZMDR001RecordPtr^.sGroupName) = 0) then
  begin
    Result := 'Y';
  end
  else
  begin
    Result := '';
  end;
end;

procedure TfrmPCNumber.btnSave2Click(Sender: TObject);
type
  TPlanNumber = packed record
    snumber: string;
    sname: string;
    sarea: string;
    sMRPGroup: string;
    sMRPType: string;
    sMRPer: string;
    dLT_M0: string;
    splannumber: string;
    sabc: string;
    sMMTypeDesc: string;
    sGroupName: string;
  end;
  PPlanNumber = ^TPlanNumber;
var
  iCount: Integer;
  aZMDR001RecordPtr: PZMDR001Record; 
  aPlanNumberPtr: PPlanNumber;
  slPlanNumber: TStringList; 
  sfile: string;      
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  aSAPMaterialReader: TSAPMaterialReader; 
begin
  sfile := '���Ʒ��Ʒģ��' + FormatDateTime('MMDD', Now);
  if not ExcelSaveDialog(sfile) then Exit;
  sfile := ChangeFileExt(sfile, '.xls');

  StatusBar1.Panels[0].Text := '��ʼ��ȡ����...';

  aSAPMaterialReader := TSAPMaterialReader.Create(leZMDR001.Text);

  StatusBar1.Panels[0].Text := 'Step 1...';

  slPlanNumber := TStringList.Create;
  try
    ProgressBar1.Max := aSAPMaterialReader.Count;
    ProgressBar1.Position := 0;
    for iCount := 0 to aSAPMaterialReader.Count - 1 do
    begin
      ProgressBar1.Position := iCount;

      aZMDR001RecordPtr := aSAPMaterialReader.Items[iCount];


      aPlanNumberPtr := New(PPlanNumber);
      aPlanNumberPtr^.snumber := aZMDR001RecordPtr^.sNumber;
      aPlanNumberPtr^.sname := aZMDR001RecordPtr^.sName;
      aPlanNumberPtr^.sarea := GetProjArea(aZMDR001RecordPtr);
      aPlanNumberPtr^.sMRPGroup := GetMrpGroup(aZMDR001RecordPtr);
      aPlanNumberPtr^.sMRPType := 'M0';
      aPlanNumberPtr^.sMRPer := GetProjMrper(aZMDR001RecordPtr); 
      aPlanNumberPtr^.dLT_M0 := GetProjLT(aZMDR001RecordPtr);
      aPlanNumberPtr^.sMMTypeDesc := aZMDR001RecordPtr^.sCategory;
      aPlanNumberPtr^.splannumber := '';
      aPlanNumberPtr^.sGroupName := aZMDR001RecordPtr.sGroupName;
      aPlanNumberPtr^.sabc := GetABC(aPlanNumberPtr^.sMRPGroup, aZMDR001RecordPtr);

      slPlanNumber.AddObject(aZMDR001RecordPtr^.sNumber, TObject(aPlanNumberPtr));
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
        ExcelApp.Cells[irow, 3].Value := aPlanNumberPtr^.sarea;
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
        ExcelApp.Cells[irow, 14].Value := aPlanNumberPtr^.sabc; //'ABC��ʶ';
        ExcelApp.Cells[irow, 15].Value := aPlanNumberPtr^.splannumber; // '�ƻ�����';
        ExcelApp.Cells[irow, 16].Value := '';
        ExcelApp.Cells[irow, 17].Value := aPlanNumberPtr^.sname;  // '��Ʒ����';
        ExcelApp.Cells[irow, 18].Value := aPlanNumberPtr^.sMMTypeDesc;  //'������������';
        ExcelApp.Cells[irow, 19].Value := aPlanNumberPtr^.sGroupName;

        irow := irow  + 1;
      end;

      ProgressBar1.Position := ProgressBar1.Max;
      
      WorkBook.SaveAs(sfile, xlExcel8);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end; 
        
  finally
    aSAPMaterialReader.Free;

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
