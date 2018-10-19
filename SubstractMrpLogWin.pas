unit SubstractMrpLogWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, CommUtils, StdCtrls, ExtCtrls, IniFiles, ComCtrls, ToolWin,
  ImgList, MrpLogReader2, ComObj;

type
  TfrmSubstractMrpLog = class(TForm)
    leMrpLog: TLabeledEdit;
    btnMrpLog: TButton;
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    ToolButton2: TToolButton;
    tbSave: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    leNumber: TLabeledEdit;
    procedure btnMrpLogClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
  private
    { Private declarations }
    procedure SaveLog(const sfile: string; lst: TList);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmSubstractMrpLog.ShowForm;
var
  frmSubstractMrpLog: TfrmSubstractMrpLog;
begin
  frmSubstractMrpLog := TfrmSubstractMrpLog.Create(nil);
  frmSubstractMrpLog.ShowModal;
  frmSubstractMrpLog.Free;
end;

procedure TfrmSubstractMrpLog.btnMrpLogClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMrpLog.Text := sfile;
end;

procedure TfrmSubstractMrpLog.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leMrpLog.Text := ini.ReadString(Self.ClassName, leMrpLog.Name, '');
    leNumber.Text := ini.ReadString(self.ClassName, leNumber.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSubstractMrpLog.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leMrpLog.Name, leMrpLog.Text);
    ini.WriteString(Self.ClassName, leNumber.Name, leNumber.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmSubstractMrpLog.ToolButton4Click(Sender: TObject);
begin
  Close;
end;

function  ListSortCompare(Item1, Item2: Pointer): Integer;
var
  p1, p2: PMrpLogRecord;
begin
  p1 := Item1;
  p2 := Item2;
  if p1^.id > p2^.id then
    Result := 1
  else if p1^.id < p2^.id then
    Result := -1
  else Result := 0;
end;

procedure TfrmSubstractMrpLog.SaveLog(const sfile: string; lst: TList);
var
  ExcelApp, WorkBook: Variant;  
  aMrpLogRecordPtr: PMrpLogRecord;
  irow: Integer;
  i: Integer;
begin


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

  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[1].Delete;
  end;
  ExcelApp.Sheets[1].Name := 'Mrp Log(' + leNumber.Text + ')';

  try
    irow := 1;
    ExcelApp.Cells[irow, 1].Value := 'ID';
    ExcelApp.Cells[irow, 2].Value := '父ID';
    ExcelApp.Cells[irow, 3].Value := '物料';
    ExcelApp.Cells[irow, 4].Value := '物料名称';
    ExcelApp.Cells[irow, 5].Value := '需求日期';
    ExcelApp.Cells[irow, 6].Value := '建议下单日期';
    ExcelApp.Cells[irow, 7].Value := '需求数量';
    ExcelApp.Cells[irow, 8].Value := '可用库存';
    ExcelApp.Cells[irow, 9].Value := 'OPO';
    ExcelApp.Cells[irow, 10].Value := '净需求';
    ExcelApp.Cells[irow, 11].Value := '替代组';
    ExcelApp.Cells[irow, 12].Value := 'MRP控制者';
    ExcelApp.Cells[irow, 13].Value := '采购员';
    ExcelApp.Cells[irow, 14].Value := 'MRP区域';
    ExcelApp.Cells[irow, 15].Value := '上层料号';
    ExcelApp.Cells[irow, 16].Value := '根料号';
    ExcelApp.Cells[irow, 17].Value := 'L/T';
                  
    irow := irow + 1;
    
    for i := 0 to lst.Count - 1 do
    begin
      aMrpLogRecordPtr := lst[i];

      ExcelApp.Cells[irow, 1].Value := aMrpLogRecordPtr^.id;// 'ID';
      ExcelApp.Cells[irow, 2].Value := aMrpLogRecordPtr^.pid;//'父ID';
      ExcelApp.Cells[irow, 3].Value := aMrpLogRecordPtr^.snumber;//'物料';
      ExcelApp.Cells[irow, 4].Value := aMrpLogRecordPtr^.sname;//'物料名称';
      ExcelApp.Cells[irow, 5].Value := aMrpLogRecordPtr^.dt;//'需求日期';
      ExcelApp.Cells[irow, 6].Value := aMrpLogRecordPtr^.dtReq;
      ExcelApp.Cells[irow, 7].Value := aMrpLogRecordPtr^.dqty;//'需求数量';
      ExcelApp.Cells[irow, 8].Value := aMrpLogRecordPtr^.dqtyStock;//'可用库存';
      ExcelApp.Cells[irow, 9].Value := aMrpLogRecordPtr^.dqtyOPO;//'OPO';
      ExcelApp.Cells[irow, 10].Value := aMrpLogRecordPtr^.dqtyNet;//'净需求';
      ExcelApp.Cells[irow, 11].Value := aMrpLogRecordPtr^.sGroup;//'替代组';
      ExcelApp.Cells[irow, 12].Value := aMrpLogRecordPtr^.sMrp;//'MRP控制者';
      ExcelApp.Cells[irow, 13].Value := aMrpLogRecordPtr^.sBuyer;//'采购员';
      ExcelApp.Cells[irow, 14].Value := aMrpLogRecordPtr^.sArea;//'MRP区域';
      ExcelApp.Cells[irow, 15].Value := aMrpLogRecordPtr^.spnumber;//'上层料号';
      ExcelApp.Cells[irow, 16].Value := aMrpLogRecordPtr^.srnumber;//'根料号';
      ExcelApp.Cells[irow, 17].Value := aMrpLogRecordPtr^.slt;//'L/T';

      irow := irow + 1;
    end;
    

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;
    
end;

{*
查找指定的值是否在当前数组中(数组已经是有序的)
*}
function SearchData(sl: TStringList; id: longint ): Integer;
var
  idMid: integer;
  idLow, idHigh: integer;
begin
  idLow := 0;
  idHigh := sl.Count - 1;


  while ( idLow <= idHigh ) do
  begin
    if idLow = idHigh then
    begin
      if PMrpLogRecord(sl.Objects[ idLow ]).id = id then
      begin
        Result := idLow;
      end
      else
      begin
        Result := -1;
      end;
      Exit;
    end;
    idMid := ( idLow + idHigh ) div 2;
    if PMrpLogRecord(sl.Objects[ idMid ]).id = id then
    begin
      Result := idMid;
      Exit;
    end;

    if PMrpLogRecord(sl.Objects[ idMid ]).id > id then idHigh := idMid - 1;
    if PMrpLogRecord(sl.Objects[ idMid ]).id < id then idLow := idMid + 1;
  end;

  Result := -1;
end;

procedure TfrmSubstractMrpLog.tbSaveClick(Sender: TObject);
var
  aMrpLogReader2: TMrpLogReader2;
  aMrpLogRecordPtr: PMrpLogRecord;
  lstID: TList;
  i: Integer;
  ig: Integer; 
  sfile: string;
  idx: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  
  lstID := TList.Create;
  aMrpLogReader2 := TMrpLogReader2.Create(leMrpLog.Text, nil);
  try
    for i := 0 to aMrpLogReader2.Count - 1 do
    begin
      aMrpLogRecordPtr := aMrpLogReader2.Items[i];
      if aMrpLogRecordPtr^.snumber = leNumber.Text then
      begin
        lstID.Add(aMrpLogRecordPtr);
        aMrpLogRecordPtr^.bCalc := True;
        for ig := 0 to aMrpLogReader2.Count - 1 do
        begin
          if aMrpLogReader2.Items[ig]^.bCalc then Continue;
          if aMrpLogReader2.Items[ig]^.sGroup = aMrpLogRecordPtr^.sGroup then
          begin
            lstID.Add(aMrpLogReader2.Items[ig]);
            aMrpLogReader2.Items[ig]^.bCalc := True;
          end;
        end;

        while aMrpLogRecordPtr^.pid <> 0 do
        begin
          idx := SearchData(aMrpLogReader2.FList, aMrpLogRecordPtr^.pid); // 根据ID查找
          if idx >= 0 then
          begin    
            aMrpLogRecordPtr := aMrpLogReader2.Items[idx];
              
            lstID.Add(aMrpLogRecordPtr);
            aMrpLogRecordPtr^.bCalc := True;

            if aMrpLogRecordPtr^.sGroup <> '0' then
            begin
              for ig := 0 to aMrpLogReader2.Count - 1 do   // 查找同替代组
              begin
                if aMrpLogReader2.Items[ig]^.bCalc then Continue;
                if aMrpLogReader2.Items[ig]^.sGroup = aMrpLogRecordPtr^.sGroup then
                begin
                  lstID.Add(aMrpLogReader2.Items[ig]);
                  aMrpLogReader2.Items[ig]^.bCalc := True;
                end;
              end;
            end;
          end
          else
            raise Exception.Create('id not found ' + IntToStr(aMrpLogRecordPtr^.pid));
        end;
      end;
    end;

    lstID.Sort(ListSortCompare);

    SaveLog(sfile, lstID);

    MessageBox(Handle, '完成', '提示', 0);
  finally
    aMrpLogReader2.Free;
    lstID.Free;
  end;
end;

end.
