unit MrpSimDemandWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ComCtrls, ToolWin, CommUtils, IniFiles, StdCtrls, ComObj,
  ExtCtrls, SAPS618Reader, ManMrpReader, SimMPSReader;

type
  TfrmMrpSimDemand = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    leSimMPS: TLabeledEdit;
    leManMrp: TLabeledEdit;
    leFCST: TLabeledEdit;
    btnFCST: TButton;
    btnManMrp: TButton;
    btnSimMPS: TButton;
    Memo1: TMemo;
    mmoError: TMemo;
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnFCSTClick(Sender: TObject);
    procedure btnManMrpClick(Sender: TObject);
    procedure btnSimMPSClick(Sender: TObject);
  private
    { Private declarations }
    procedure OnLogEvent(const s: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmMrpSimDemand.ShowForm;
var
  frmMrpSimDemand: TfrmMrpSimDemand;
begin
  frmMrpSimDemand := TfrmMrpSimDemand.Create(nil);
  try
    frmMrpSimDemand.ShowModal;
  finally
    frmMrpSimDemand.Free;
  end;
end;

procedure TfrmMrpSimDemand.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leFCST.Text := ini.ReadString(self.ClassName, leFCST.Name, '');
    leManMrp.Text := ini.ReadString(self.ClassName, leManMrp.Name, '');
    leSimMPS.Text := ini.ReadString(self.ClassName, leSimMPS.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmMrpSimDemand.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leFCST.Name, leFCST.Text);
    ini.WriteString(self.ClassName, leManMrp.Name, leManMrp.Text);
    ini.WriteString(self.ClassName, leSimMPS.Name, leSimMPS.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmMrpSimDemand.btnExitClick(Sender: TObject);
begin
  Close;
end;
 
procedure TfrmMrpSimDemand.btnFCSTClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leFCST.Text := sfile;
end;

procedure TfrmMrpSimDemand.btnManMrpClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leManMrp.Text := sfile;
end;

procedure TfrmMrpSimDemand.btnSimMPSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSimMPS.Text := sfile;
end;

procedure TfrmMrpSimDemand.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

procedure TfrmMrpSimDemand.btnSave2Click(Sender: TObject);
var
  sfile: string;
  aSAPS618Reader: TSAPPIRReader;
  aManMrpReader: TManMrpReader;
  aSimMPSReader: TSimMPSReader;
  slDelta: TStringList;
  iMPS: Integer;
  snumber: string;
  iQty: Integer;
  idxManMrp: Integer;
  slSubs: TStringList;
  iSub: Integer;

  ExcelApp, WorkBook: Variant;
  irow: Integer;
  iweek: Integer;
  icol: Integer;
  i: Integer;
  aSAPS618: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
begin
  if not ExcelSaveDialog(sfile) then Exit;

         
  Memo1.Lines.Add('开始读取 营销模拟MPS  ' + leFCST.Text);
  aSimMPSReader := TSimMPSReader.Create(leSimMPS.Text);

  Memo1.Lines.Add('开始读取 PIR  ' + leFCST.Text);
  aSAPS618Reader := TSAPPIRReader.Create(leFCST.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 手工MRP  ' + leFCST.Text);
  aManMrpReader := TManMrpReader.Create(leManMrp.Text);

  slDelta := TStringList.Create;
  try
    Memo1.Lines.Add('开始计算总量差额');

    for iMPS := 0 to aSimMPSReader.FList.Count - 1 do
    begin
      snumber := aSimMPSReader.FList[iMPS];
      iQty := Integer(aSimMPSReader.FList.Objects[iMPS]);

      idxManMrp := aManMrpReader.FList.IndexOf(snumber);
      if idxManMrp >= 0 then
      begin
        iQty := TManMrpLine(aManMrpReader.FList.Objects[idxManMrp]).dSum - iQty;  // 负数代表增加需求
        slDelta.AddObject(snumber, TObject(iQty));
      end
      else
      begin
        mmoError.Lines.Add('模拟MPS ' + snumber + '(' + IntToStr(iQty) + ') 在手工物料需求找不到');
      end; 
    end;

    slSubs := TStringList.Create;
                                              
    Memo1.Lines.Add('开始合并总差额到独立需求');
    
    for iMPS := 0 to slDelta.Count - 1 do
    begin
      snumber := slDelta[iMPS];
      iQty := Integer(slDelta.Objects[iMPS]);

      slSubs.Clear;
      aManMrpReader.GetSubs(snumber, slSubs);

      for iSub := 0 to slSubs.Count - 1 do
      begin
        if iQty = 0 then Break;

        aSAPS618Reader.SubNumber(slSubs[iSub], iQty);
      end;
                                            
      if iQty <> 0 then
      begin
        Memo1.Lines.Add(snumber + ' : ' + IntToStr(iQty));
      end;
    end;
 
    slSubs.Free;




    // 开始保存 Excel
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
        Exit;
      end;
    end;

    WorkBook := ExcelApp.WorkBooks.Add;

    try
      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '物料';
      ExcelApp.Cells[irow, 2].Value := '物料描述';
      ExcelApp.Cells[irow, 3].Value := '物料组';
      ExcelApp.Cells[irow, 4].Value := '物料组描述';
      ExcelApp.Cells[irow, 5].Value := '工厂';
      ExcelApp.Cells[irow, 6].Value := '需求类型';
      ExcelApp.Cells[irow, 7].Value := '版本';
      ExcelApp.Cells[irow, 8].Value := 'Act';
      ExcelApp.Cells[irow, 9].Value := '需求计划号';
      ExcelApp.Cells[irow, 10].Value := 'MRP 范围';
      ExcelApp.Cells[irow, 11].Value := 'MRP控制者';
      ExcelApp.Cells[irow, 12].Value := '基本计量单位';

      ExcelApp.Columns[1].ColumnWidth := 16;
      ExcelApp.Columns[2].ColumnWidth := 44;

      for iweek := 0 to aSAPS618Reader.slWeek.Count - 1 do
      begin
        icol := iweek + 13;
        ExcelApp.Cells[irow, icol].Value := aSAPS618Reader.slWeek.Names[iweek];
        ExcelApp.Columns[icol].ColumnWidth := 9.5;
      end;

      AddColor(ExcelApp, irow, 1, irow, aSAPS618Reader.slWeek.Count + 12, $C0C0C0);

      irow := 2;
      for i := 0 to aSAPS618Reader.Count - 1 do
      begin
        aSAPS618 := aSAPS618Reader.Items[i];
        ExcelApp.Cells[irow, 1].Value := aSAPS618.FNumber; // '物料';
        ExcelApp.Cells[irow, 2].Value := aSAPS618.sname; //'物料描述';
        ExcelApp.Cells[irow, 3].Value := aSAPS618.sgroup; //'物料组';
        ExcelApp.Cells[irow, 4].Value := aSAPS618.sgroupname; //'物料组描述';
        ExcelApp.Cells[irow, 5].Value := aSAPS618.sfac; //'工厂';
        ExcelApp.Cells[irow, 6].Value := aSAPS618.sDemandType; //'需求类型';
        ExcelApp.Cells[irow, 7].Value := '''' + aSAPS618.sDemandVer; //'版本';
        ExcelApp.Cells[irow, 8].Value := aSAPS618.sAct; //'Act';
        ExcelApp.Cells[irow, 9].Value := aSAPS618.sPlanNo; //'需求计划号';
        ExcelApp.Cells[irow, 10].Value :=aSAPS618.FMrpArea; // 'MRP 范围';
        ExcelApp.Cells[irow, 11].Value :=aSAPS618.sMrper; // 'MRP控制者';
        ExcelApp.Cells[irow, 12].Value :=aSAPS618.sUnit; // '基本计量单位';

        for iweek := 0 to aSAPS618.Count - 1 do
        begin               
          icol := iweek + 13;
          aSAPS618ColPtr := aSAPS618.Items[iweek];
          ExcelApp.Cells[irow, icol].Value := aSAPS618ColPtr^.dqty;
        end;

        irow := irow + 1;
      end;  
       
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;


  finally
    aSAPS618Reader.Free;
    aManMrpReader.Free;
    aSimMPSReader.Free;

    slDelta.Free;
  end;
  MessageBox(Handle, '完成', '提示', 0);
end;

end.
