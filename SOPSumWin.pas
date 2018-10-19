unit SOPSumWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ComCtrls, ToolWin, CommUtils, IniFiles, StdCtrls, ComObj,
  ExtCtrls, ZMDR001Reader, FGStockRptReader, FGRptWinReader, DateUtils,
  FGToProduceReader, SAPS618Reader, HWOrdInfoReader;

type
  TfrmSOPSum = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ImageList1: TImageList;
    leMMList: TLabeledEdit;
    btnMMList: TButton;
    Memo1: TMemo;
    leFGReport1: TLabeledEdit;
    btnFGReport1: TButton;
    dtpWin1: TDateTimePicker;
    dtpWin2: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    btnToProduce: TButton;
    leToProduce: TLabeledEdit;
    leFGReport2: TLabeledEdit;
    btnFGReport2: TButton;
    mmoMap: TMemo;
    Label3: TLabel;
    lePIR_HW: TLabeledEdit;
    btnPIR_HW: TButton;
    leHWODInfo: TLabeledEdit;
    btnHWODInfo: TButton;
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnMMListClick(Sender: TObject);
    procedure btnToProduceClick(Sender: TObject);
    procedure btnFGReport1Click(Sender: TObject);
    procedure btnFGReport2Click(Sender: TObject);
    procedure btnPIR_HWClick(Sender: TObject);
    procedure btnHWODInfoClick(Sender: TObject);
  private
    { Private declarations }
      procedure OnLogEvent(const s: string);
      function GetQtyPIR(aSAPPIRReader: TSAPPIRReader; const sProjName,
        skey: string; slMap: TStrings; acs: TCountrySet): Double;
      function GetHWOrdInfoQty(aHWOrdInfoReader: THWOrdInfoReader; const sProjName,
        skey: string; slMap: TStrings; acs: TCountrySet): Double;
      function GetHWOODQty(aHWOrdInfoReader: THWOrdInfoReader; const sProjName,
        skey: string; slMap: TStrings; acs: TCountrySet): Double;
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

type
  TFGColType = (ctSF, ctGN, ctHW);   


const
  CSColType: array[TFGColType] of string = ('顺丰', '国内', '海外');

class procedure TfrmSOPSum.ShowForm;
var
  frmSOPSum: TfrmSOPSum;
begin
  frmSOPSum := TfrmSOPSum.Create(nil);
  try
    frmSOPSum.ShowModal;
  finally
    frmSOPSum.Free;
  end;
end;  
   
procedure TfrmSOPSum.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    leMMList.Text := ini.ReadString(self.ClassName, leMMList.Name, '');
    leFGReport1.Text := ini.ReadString(self.ClassName, leFGReport1.Name, '');
    leFGReport2.Text := ini.ReadString(self.ClassName, leFGReport2.Name, '');    
    leToProduce.Text := ini.ReadString(self.ClassName, leToProduce.Name, '');
    s := ini.ReadString(self.ClassName, mmoMap.Name, '');
    mmoMap.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
    lePIR_HW.Text := ini.ReadString(self.ClassName, lePIR_HW.Name, '');
    leHWODInfo.Text := ini.ReadString(self.ClassName, leHWODInfo.Name, '');
  finally
    ini.Free;
  end;

  dtpWin1.DateTime := StartOfTheWeek(Now);
  dtpWin2.DateTime := dtpWin1.DateTime + 2;
end;

procedure TfrmSOPSum.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leMMList.Name, leMMList.Text);
    ini.WriteString(self.ClassName, leFGReport1.Name, leFGReport1.Text);
    ini.WriteString(self.ClassName, leFGReport2.Name, leFGReport2.Text);
    ini.WriteString(self.ClassName, leToProduce.Name, leToProduce.Text);
    s := StringReplace(mmoMap.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmoMap.Name, s);
    ini.WriteString(self.ClassName, lePIR_HW.Name, lePIR_HW.Text);
    ini.WriteString(self.ClassName, leHWODInfo.Name, leHWODInfo.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmSOPSum.btnExitClick(Sender: TObject);
begin
  Close;
end;
     
procedure TfrmSOPSum.btnMMListClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMMList.Text := sfile;
end;

procedure TfrmSOPSum.btnToProduceClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leToProduce.Text := sfile;
end;

procedure TfrmSOPSum.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

function GetQty4Color(const sProjNo, skey: string; aFGStockRptReader: TFGStockRptReader;
  ct: TFGColType; acs: TCountrySet): Double;
var
  i: Integer;
  aFGStockRptRecordPtr: PFGStockRptRecord;      
  aZMDR001RecordPtr: PZMDR001Record;
begin
  Result := 0;
  case ct of
    ctSF:
    begin
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) = 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sColor <> skey then Continue;
        Result := Result + aFGStockRptRecordPtr^.dqty;
      end;
    end;

    ctGN:
    begin   
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];   
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) > 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);    
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sColor <> skey then Continue;
        if TSAPMaterialReader.IsHW(aZMDR001RecordPtr) then Continue;
        Result := Result + aFGStockRptRecordPtr^.dqty;      
      end;
    end;

    ctHW:
    begin
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];   
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) > 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sColor <> skey then Continue;
        if not TSAPMaterialReader.IsHW( aZMDR001RecordPtr ) then Continue;

        case acs of
          csC:;         // 海外，不分国家
            //Continue;
          csR:
            if (Pos('俄文', aZMDR001RecordPtr^.sName) = 0)
              and (Pos('俄罗斯', aZMDR001RecordPtr^.sName) = 0) then Continue;
          csU:
            if Pos('乌克兰', aZMDR001RecordPtr^.sName) = 0 then Continue;
          csI:
            if Pos('印尼', aZMDR001RecordPtr^.sName) = 0 then Continue;
          csO:
            if (Pos('俄文', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('俄罗斯', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('乌克兰', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('印尼', aZMDR001RecordPtr^.sName) > 0) then Continue;
        end;

        Result := Result + aFGStockRptRecordPtr^.dqty;      
      end;
    end;
  end;
end;

function TfrmSOPSum.GetQtyPIR(aSAPPIRReader: TSAPPIRReader; const sProjName,
  skey: string; slMap: TStrings; acs: TCountrySet): Double;
var
  i: Integer;
  aZMDR001RecordPtr: PZMDR001Record;
  aSAPS618: TSAPS618;
  skey_p: string;
  idx: Integer;
begin
  Result := 0;

  skey_p := skey;
  if slMap <> nil then
  begin
    idx := slMap.IndexOfName(skey_p);
    if idx >= 0 then
    begin
      skey_p := slMap.ValueFromIndex[idx];
    end;
  end;


  for i := 0 to aSAPPIRReader.Count - 1 do
  begin
    aSAPS618 := aSAPPIRReader.Items[i];
    if aSAPS618.Tag = nil then Continue;
    aZMDR001RecordPtr := PZMDR001Record(aSAPS618.Tag);
    if sProjName <> aZMDR001RecordPtr^.sProj then Continue;
    if (skey_p = aZMDR001RecordPtr^.sColor) or (skey_p = aZMDR001RecordPtr^.sVer)
      or (skey_p = aZMDR001RecordPtr^.scap) then
    begin

      case acs of
        csC:;         // 海外，不分国家
          //Continue;
        csR:
          if (Pos('俄文', aZMDR001RecordPtr^.sName) = 0)
            and (Pos('俄罗斯', aZMDR001RecordPtr^.sName) = 0) then Continue;
        csU:
          if Pos('乌克兰', aZMDR001RecordPtr^.sName) = 0 then Continue;
        csI:
          if Pos('印尼', aZMDR001RecordPtr^.sName) = 0 then Continue;
        csO:
          if (Pos('俄文', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('俄罗斯', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('乌克兰', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('印尼', aZMDR001RecordPtr^.sName) > 0) then Continue;
      end;
    
      Result := Result + aSAPS618.GetSum;
    end;  
  end;
end;

function TfrmSOPSum.GetHWOrdInfoQty(aHWOrdInfoReader: THWOrdInfoReader; const sProjName,
  skey: string; slMap: TStrings; acs: TCountrySet): Double;
var
  i: Integer;
  aZMDR001RecordPtr: PZMDR001Record;
  aHWOrdInfoRecordPtr: PHWOrdInfoRecord;
  skey_p: string;
  idx: Integer;
begin
  Result := 0;

  skey_p := skey;
  if slMap <> nil then
  begin
    idx := slMap.IndexOfName(skey_p);
    if idx >= 0 then
    begin
      skey_p := slMap.ValueFromIndex[idx];
    end;
  end;


  for i := 0 to aHWOrdInfoReader.OrdInfoCount - 1 do
  begin
    aHWOrdInfoRecordPtr := aHWOrdInfoReader.OrdInfoItems[i];
    if aHWOrdInfoRecordPtr^.Tag = nil then Continue;
    aZMDR001RecordPtr := PZMDR001Record(aHWOrdInfoRecordPtr^.Tag);
    if sProjName <> aZMDR001RecordPtr^.sProj then Continue;
    if (skey_p = aZMDR001RecordPtr^.sColor) or (skey_p = aZMDR001RecordPtr^.sVer)
      or (skey_p = aZMDR001RecordPtr^.scap) then
    begin

      case acs of
        csC:;         // 海外，不分国家
          //Continue;
        csR:
          if (Pos('俄文', aZMDR001RecordPtr^.sName) = 0)
            and (Pos('俄罗斯', aZMDR001RecordPtr^.sName) = 0) then Continue;
        csU:
          if Pos('乌克兰', aZMDR001RecordPtr^.sName) = 0 then Continue;
        csI:
          if Pos('印尼', aZMDR001RecordPtr^.sName) = 0 then Continue;
        csO:
          if (Pos('俄文', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('俄罗斯', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('乌克兰', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('印尼', aZMDR001RecordPtr^.sName) > 0) then Continue;
      end;
    
      Result := Result + aHWOrdInfoRecordPtr^.dqty;
    end;  
  end;
end;

function TfrmSOPSum.GetHWOODQty(aHWOrdInfoReader: THWOrdInfoReader; const sProjName,
  skey: string; slMap: TStrings; acs: TCountrySet): Double;
var
  i: Integer;
  aZMDR001RecordPtr: PZMDR001Record;
  aHWOODRecordPtr: PHWOODRecord;
  skey_p: string;
  idx: Integer;
begin
  Result := 0;

  skey_p := skey;
  if slMap <> nil then
  begin
    idx := slMap.IndexOfName(skey_p);
    if idx >= 0 then
    begin
      skey_p := slMap.ValueFromIndex[idx];
    end;
  end;
 
  for i := 0 to aHWOrdInfoReader.OODCount - 1 do
  begin
    aHWOODRecordPtr := aHWOrdInfoReader.OODItems[i];
    if aHWOODRecordPtr^.Tag = nil then Continue;
    aZMDR001RecordPtr := PZMDR001Record(aHWOODRecordPtr^.Tag);
    if sProjName <> aZMDR001RecordPtr^.sProj then Continue;
    if (skey_p = aZMDR001RecordPtr^.sColor) or (skey_p = aZMDR001RecordPtr^.sVer)
      or (skey_p = aZMDR001RecordPtr^.scap) then
    begin

      case acs of
        csC:;         // 海外，不分国家
          //Continue;
        csR:
          if (Pos('俄文', aZMDR001RecordPtr^.sName) = 0)
            and (Pos('俄罗斯', aZMDR001RecordPtr^.sName) = 0) then Continue;
        csU:
          if Pos('乌克兰', aZMDR001RecordPtr^.sName) = 0 then Continue;
        csI:
          if Pos('印尼', aZMDR001RecordPtr^.sName) = 0 then Continue;
        csO:
          if (Pos('俄文', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('俄罗斯', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('乌克兰', aZMDR001RecordPtr^.sName) > 0)
            or (Pos('印尼', aZMDR001RecordPtr^.sName) > 0) then Continue;
      end;
    
      Result := Result + aHWOODRecordPtr^.dqty;
    end;  
  end;
end;
      
function GetQty4Cap(const sprojno, skey: string; aFGStockRptReader: TFGStockRptReader;
  ct: TFGColType; acs: TCountrySet): Double;
var
  i: Integer;
  aFGStockRptRecordPtr: PFGStockRptRecord;      
  aZMDR001RecordPtr: PZMDR001Record;
begin
  Result := 0;
  case ct of
    ctSF:
    begin
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sprojno then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) = 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);    
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sCap <> skey then Continue;
        Result := Result + aFGStockRptRecordPtr^.dqty;
      end;
    end;

    ctGN:
    begin   
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];    
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sprojno then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) > 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);  
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sCap <> skey then Continue;
        if aZMDR001RecordPtr^.sVer = '海外' then Continue;
        Result := Result + aFGStockRptRecordPtr^.dqty;      
      end;
    end;

    ctHW:
    begin
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];   
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sprojno then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) > 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr); 
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sCap <> skey then Continue;
        if aZMDR001RecordPtr^.sVer <> '海外' then Continue;     

        case acs of
          csC:;         // 海外，不分国家
            //Continue;
          csR:
            if (Pos('俄文', aZMDR001RecordPtr^.sName) = 0)
              and (Pos('俄罗斯', aZMDR001RecordPtr^.sName) = 0) then Continue;
          csU:
            if Pos('乌克兰', aZMDR001RecordPtr^.sName) = 0 then Continue;
          csI:
            if Pos('印尼', aZMDR001RecordPtr^.sName) = 0 then Continue;
          csO:
            if (Pos('俄文', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('俄罗斯', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('乌克兰', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('印尼', aZMDR001RecordPtr^.sName) > 0) then Continue;
        end;

        Result := Result + aFGStockRptRecordPtr^.dqty;      
      end;
    end;
  end;
end;   
          
function GetQty4Ver(const sProjNo, skey: string; aFGStockRptReader: TFGStockRptReader;
  ct: TFGColType; acs: TCountrySet): Double;
var
  i: Integer;
  aFGStockRptRecordPtr: PFGStockRptRecord;      
  aZMDR001RecordPtr: PZMDR001Record;
begin
  Result := 0;
  case ct of
    ctSF:
    begin
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) = 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);   
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sVer <> skey then Continue;
        Result := Result + aFGStockRptRecordPtr^.dqty;
      end;
    end;

    ctGN:
    begin   
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];  
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) > 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);  
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sVer <> skey then Continue;
        if aZMDR001RecordPtr^.sVer = '海外' then Continue;
        Result := Result + aFGStockRptRecordPtr^.dqty;      
      end;
    end;

    ctHW:
    begin
      for i := 0 to aFGStockRptReader.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader.Items[i];     
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then Continue;
        if Pos('顺丰', aFGStockRptRecordPtr^.ssheet) > 0 then Continue;
        aZMDR001RecordPtr := PZMDR001Record(aFGStockRptRecordPtr^.ptr);  
        if aZMDR001RecordPtr = nil then Continue;
        if aZMDR001RecordPtr^.sVer <> skey then Continue;
        if aZMDR001RecordPtr^.sVer <> '海外' then Continue;   

        case acs of
          csC: ;  // 海外，不分国家
            //Continue;
          csR:
            if (Pos('俄文', aZMDR001RecordPtr^.sName) = 0)
              and (Pos('俄罗斯', aZMDR001RecordPtr^.sName) = 0) then Continue;
          csU:
            if Pos('乌克兰', aZMDR001RecordPtr^.sName) = 0 then Continue;
          csI:
            if Pos('印尼', aZMDR001RecordPtr^.sName) = 0 then Continue;
          csO:
            if (Pos('俄文', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('俄罗斯', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('乌克兰', aZMDR001RecordPtr^.sName) > 0)
              or (Pos('印尼', aZMDR001RecordPtr^.sName) > 0) then Continue;
        end;

        Result := Result + aFGStockRptRecordPtr^.dqty;      
      end;
    end;
  end;
end;

procedure TfrmSOPSum.btnSave2Click(Sender: TObject);
var
  sfile: string;
  aSAPMaterialReader: TSAPMaterialReader;
  aFGStockRptReader1: TFGStockRptReader;
  aFGStockRptReader2: TFGStockRptReader;
  iProj: Integer;
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  irow_proj1: Integer;
  icol: Integer;
  slCol: TStringList;
  slCap: TStringList;
  slVer: TStringList;
  sProjNo: string;
  sProjName: string;
  inumber: Integer;
  aFGStockRptRecordPtr: PFGStockRptRecord;
  aZMDR001RecordPtr: PZMDR001Record; 
  ct: TFGColType;

  aFGRptWinReader1: TFGRptWinReader;
  aFGRptWinReader2: TFGRptWinReader;
  aFGToProduceReader: TFGToProduceReader;
  aSAPPIRReader: TSAPPIRReader;
  aHWOrdInfoReader: THWOrdInfoReader;
  aSAPS618: TSAPS618;
  aHWOrdInfoRecordPtr: PHWOrdInfoRecord;
  aHWOODRecordPtr: PHWOODRecord;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  Memo1.Lines.Add('开始......');

  Memo1.Lines.Add('打开物料清单...');
  aSAPMaterialReader := TSAPMaterialReader.Create(leMMList.Text, OnLogEvent);

  Memo1.Lines.Add('打开成品报表  库存...');
  aFGStockRptReader1 := TFGStockRptReader.Create(leFGReport1.Text, OnLogEvent);
  aFGStockRptReader2 := TFGStockRptReader.Create(leFGReport2.Text, OnLogEvent);

  Memo1.Lines.Add('打开成品报表  产品入库...');
  aFGRptWinReader1 := TFGRptWinReader.Create(leFGReport1.Text, OnLogEvent);
  aFGRptWinReader2 := TFGRptWinReader.Create(leFGReport2.Text, OnLogEvent);

  Memo1.Lines.Add('打开待产数据...');
  aFGToProduceReader := TFGToProduceReader.Create(leToProduce.Text, OnLogEvent);
  aFGToProduceReader.SubWin(aFGRptWinReader1, aSAPMaterialReader, dtpWin1.DateTime, dtpWin2.DateTime);    
  aFGToProduceReader.SubWin(aFGRptWinReader2, aSAPMaterialReader, dtpWin1.DateTime, dtpWin2.DateTime);

                       
  Memo1.Lines.Add('打开海外独立需求...');
  aSAPPIRReader := TSAPPIRReader.Create(lePIR_HW.Text);
                       
  Memo1.Lines.Add('打开海外订单完成情况...');
  aHWOrdInfoReader := THWOrdInfoReader.Create(leHWODInfo.Text);

  for inumber := 0 to aSAPPIRReader.Count - 1 do
  begin
    aSAPS618 := aSAPPIRReader.Items[inumber];
    aSAPS618.Tag := TObject(aSAPMaterialReader.GetSAPMaterialRecord(aSAPS618.FNumber));
    if aSAPS618.Tag = nil then
    begin
      Memo1.Lines.Add(aSAPS618.FNumber + ' 找不到物料信息');
    end;
  end;  
       
  for inumber := 0 to aHWOrdInfoReader.OrdInfoCount - 1 do
  begin
    aHWOrdInfoRecordPtr := aHWOrdInfoReader.OrdInfoItems[inumber];
    aHWOrdInfoRecordPtr^.Tag := TObject(aSAPMaterialReader.GetSAPMaterialRecord(aHWOrdInfoRecordPtr^.snumber));
    if aHWOrdInfoRecordPtr^.Tag = nil then
    begin
      Memo1.Lines.Add(aHWOrdInfoRecordPtr^.sNumber + ' 海外订单 找不到物料信息');
    end;
  end;  
           
  for inumber := 0 to aHWOrdInfoReader.OODCount - 1 do
  begin
    aHWOODRecordPtr := aHWOrdInfoReader.OODItems[inumber];
    aHWOODRecordPtr^.Tag := TObject(aSAPMaterialReader.GetSAPMaterialRecord(aHWOODRecordPtr^.snumber));
    if aHWOODRecordPtr^.Tag = nil then
    begin
      Memo1.Lines.Add(aHWOODRecordPtr^.sNumber + ' 海外未结订单 找不到物料信息');
    end;
  end;  

  try

    Memo1.Lines.Add('开始统计成品库存...');
 
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
      ExcelApp.Sheets[2].Delete;
    end;

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '机型';
    ExcelApp.Cells[irow, 2].Value := '项目';
    icol := 3;
    ExcelApp.Cells[irow, icol].Value := '库存';
    MergeCells(ExcelApp, irow, icol, irow, icol + Ord(High(TFGColType)));
    for ct := Low(TFGColType) to High(TFGColType) do
    begin
      ExcelApp.Cells[irow + 1, icol + Ord(ct)].Value := CSColType[ct];
    end;
    icol := icol + Ord(High(TFGColType)) + 1;
    ExcelApp.Cells[irow, icol].Value := '待产';                             
    MergeCells(ExcelApp, irow, icol, irow, icol + 1);
    ExcelApp.Cells[irow + 1, icol].Value := '国内';
    ExcelApp.Cells[irow + 1, icol + 1].Value := '海外';

    MergeCells(ExcelApp, irow, 1, irow + 1, 1); 
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);

    CenterCells(ExcelApp, 1, 1, 2, 7);

    irow := irow + 2;
    irow_proj1 := irow;

    for iProj := 0 to aFGStockRptReader1.ProjCount - 1 do
    begin
      sProjNo := aFGStockRptReader1.Projs[iProj];

      slCol := TStringList.Create;
      slCap := TStringList.Create;
      slVer := TStringList.Create;

      sProjName := aSAPMaterialReader.ProjNo2Name(sProjNo);
      ExcelApp.Cells[irow, 1].Value := sProjName;

      for inumber := 0 to aFGStockRptReader1.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader1.Items[inumber];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then
          Continue;

        aZMDR001RecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aFGStockRptRecordPtr^.snumber);

        if aZMDR001RecordPtr = nil then
        begin
          Memo1.Lines.Add(aFGStockRptRecordPtr^.snumber + '  找不到物料主数据' );
          Continue;
        end;

        aFGStockRptRecordPtr^.ptr := aZMDR001RecordPtr;

        if slCol.IndexOf(aZMDR001RecordPtr^.sColor) < 0 then
        begin
          slCol.Add(aZMDR001RecordPtr^.sColor);
        end;


        if slCap.IndexOf(aZMDR001RecordPtr^.scap) < 0 then
        begin
          slCap.Add(aZMDR001RecordPtr^.scap);
        end;

        if slVer.IndexOf(aZMDR001RecordPtr^.sVer) < 0 then
        begin
          slVer.Add(aZMDR001RecordPtr^.sVer);
        end;
      end;

      for inumber := 0 to slCol.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCol[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctGN, csC);    //国内
        ExcelApp.Cells[irow, 5].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctHW, csC);    //海外
        ExcelApp.Cells[irow, 6].Value := aFGToProduceReader.GetQty(sProjName, slCol[inumber], nil, False, csC);    //国内
        ExcelApp.Cells[irow, 7].Value := aFGToProduceReader.GetQty(sProjName, slCol[inumber], nil, True, csC);     //海外   
        irow := irow + 1;
      end;

      for inumber := 0 to slCap.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCap[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctGN, csC);    //国内
        ExcelApp.Cells[irow, 5].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctHW, csC);    //海外
        ExcelApp.Cells[irow, 6].Value := aFGToProduceReader.GetQty(sProjName, slCap[inumber], nil, False, csC);    //国内
        ExcelApp.Cells[irow, 7].Value := aFGToProduceReader.GetQty(sProjName, slCap[inumber], nil, True, csC);     //海外
        irow := irow + 1;
      end;
                                  
      for inumber := 0 to slVer.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slVer[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctGN, csC);    //国内
        ExcelApp.Cells[irow, 5].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctHW, csC);    //海外
        ExcelApp.Cells[irow, 6].Value := aFGToProduceReader.GetQty(sProjName, slVer[inumber], mmoMap.Lines, False, csC);    //国内
        ExcelApp.Cells[irow, 7].Value := aFGToProduceReader.GetQty(sProjName, slVer[inumber], mmoMap.Lines, True, csC);     //海外
        irow := irow + 1;
      end;
                        
      slCol.Free;
      slCap.Free;
      slVer.Free;

      MergeCells(ExcelApp, irow_proj1, 1, irow - 1, 1);
      irow_proj1 := irow;
    end;




    for iProj := 0 to aFGStockRptReader2.ProjCount - 1 do
    begin
      sProjNo := aFGStockRptReader2.Projs[iProj];

      slCol := TStringList.Create;
      slCap := TStringList.Create;
      slVer := TStringList.Create;

      sProjName := aSAPMaterialReader.ProjNo2Name(sProjNo);
      ExcelApp.Cells[irow, 1].Value := sProjName;

      for inumber := 0 to aFGStockRptReader2.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader2.Items[inumber];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then
          Continue;

        aZMDR001RecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aFGStockRptRecordPtr^.snumber);

        if aZMDR001RecordPtr = nil then
        begin
          Memo1.Lines.Add(aFGStockRptRecordPtr^.snumber + '  找不到物料主数据' );
          Continue;
        end;

        aFGStockRptRecordPtr^.ptr := aZMDR001RecordPtr;

        if slCol.IndexOf(aZMDR001RecordPtr^.sColor) < 0 then
        begin
          slCol.Add(aZMDR001RecordPtr^.sColor);
        end;


        if slCap.IndexOf(aZMDR001RecordPtr^.scap) < 0 then
        begin
          slCap.Add(aZMDR001RecordPtr^.scap);
        end;

        if slVer.IndexOf(aZMDR001RecordPtr^.sVer) < 0 then
        begin
          slVer.Add(aZMDR001RecordPtr^.sVer);
        end;
      end;

      for inumber := 0 to slCol.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCol[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctGN, csC);    //国内
        ExcelApp.Cells[irow, 5].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctHW, csC);    //海外      
        ExcelApp.Cells[irow, 6].Value := aFGToProduceReader.GetQty(sProjName, slCol[inumber], nil, False, csC);    //国内
        ExcelApp.Cells[irow, 7].Value := aFGToProduceReader.GetQty(sProjName, slCol[inumber], nil, True, csC);     //海外
        irow := irow + 1;
      end;

      for inumber := 0 to slCap.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCap[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctGN, csC);    //国内
        ExcelApp.Cells[irow, 5].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctHW, csC);    //海外
        ExcelApp.Cells[irow, 6].Value := aFGToProduceReader.GetQty(sProjName, slCap[inumber], nil, False, csC);    //国内
        ExcelApp.Cells[irow, 7].Value := aFGToProduceReader.GetQty(sProjName, slCap[inumber], nil, True, csC);     //海外
        irow := irow + 1;
      end;
                                  
      for inumber := 0 to slVer.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slVer[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctGN, csC);    //国内
        ExcelApp.Cells[irow, 5].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctHW, csC);    //海外    
        ExcelApp.Cells[irow, 6].Value := aFGToProduceReader.GetQty(sProjName, slVer[inumber], mmoMap.Lines, False, csC);    //国内
        ExcelApp.Cells[irow, 7].Value := aFGToProduceReader.GetQty(sProjName, slVer[inumber], mmoMap.Lines, True, csC);     //海外
        irow := irow + 1;
      end;
                        
      slCol.Free;
      slCap.Free;
      slVer.Free;

      MergeCells(ExcelApp, irow_proj1, 1, irow - 1, 1);
      irow_proj1 := irow;
    end;


    AddBorder(ExcelApp, 1, 1, irow - 1, 7);

    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////

    ExcelApp.Sheets.Add(After := ExcelApp.Sheets[1]);
    ExcelApp.Sheets[2].Activate;
    ExcelApp.Sheets[2].Name := '分国家';
                               
    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '机型';
    ExcelApp.Cells[irow, 2].Value := '项目';     
    MergeCells(ExcelApp, irow, 1, irow + 2, 1);
    MergeCells(ExcelApp, irow, 2, irow + 2, 2);
    
    icol := 3;
    ExcelApp.Cells[irow, icol].Value := '库存';
    MergeCells(ExcelApp, irow, icol, irow, icol + 5);
    
    ExcelApp.Cells[irow + 1, icol].Value := '顺丰';   
    ExcelApp.Cells[irow + 1, icol + 1].Value := '国内';
    ExcelApp.Cells[irow + 1, icol + 2].Value := '海外';
    MergeCells(ExcelApp, irow + 1, icol, irow + 2, icol);
    MergeCells(ExcelApp, irow + 1, icol + 1, irow + 2, icol + 1);
    MergeCells(ExcelApp, irow + 1, icol + 2, irow + 1, icol + 5);

    ExcelApp.Cells[irow + 2, icol + 2].Value := '俄罗斯';
    ExcelApp.Cells[irow + 2, icol + 3].Value := '乌克兰';
    ExcelApp.Cells[irow + 2, icol + 4].Value := '印尼';
    ExcelApp.Cells[irow + 2, icol + 5].Value := '其它';

    icol := icol + 6;
    ExcelApp.Cells[irow, icol].Value := '海外独立需求';
    MergeCells(ExcelApp, irow, icol, irow + 1, icol + 3); 
    ExcelApp.Cells[irow + 2, icol].Value := '俄罗斯';
    ExcelApp.Cells[irow + 2, icol + 1].Value := '乌克兰';
    ExcelApp.Cells[irow + 2, icol + 2].Value := '印尼';
    ExcelApp.Cells[irow + 2, icol + 3].Value := '其它';
                       
    icol := icol + 4;
    ExcelApp.Cells[irow, icol].Value := '订单';
    MergeCells(ExcelApp, irow, icol, irow + 1, icol + 3);
    ExcelApp.Cells[irow + 2, icol].Value := '俄罗斯';
    ExcelApp.Cells[irow + 2, icol + 1].Value := '乌克兰';
    ExcelApp.Cells[irow + 2, icol + 2].Value := '印尼';
    ExcelApp.Cells[irow + 2, icol + 3].Value := '其它';

    icol := icol + 4;
    ExcelApp.Cells[irow, icol].Value := '应提';
    MergeCells(ExcelApp, irow, icol, irow + 1, icol + 3); 
    ExcelApp.Cells[irow + 2, icol].Value := '俄罗斯';
    ExcelApp.Cells[irow + 2, icol + 1].Value := '乌克兰';
    ExcelApp.Cells[irow + 2, icol + 2].Value := '印尼';
    ExcelApp.Cells[irow + 2, icol + 3].Value := '其它';
 

    CenterCells(ExcelApp, 1, 1, 3, 20);

    irow := irow + 3;
    irow_proj1 := irow;

    for iProj := 0 to aFGStockRptReader1.ProjCount - 1 do
    begin
      sProjNo := aFGStockRptReader1.Projs[iProj];

      slCol := TStringList.Create;
      slCap := TStringList.Create;
      slVer := TStringList.Create;

      sProjName := aSAPMaterialReader.ProjNo2Name(sProjNo);
      ExcelApp.Cells[irow, 1].Value := sProjName;

      for inumber := 0 to aFGStockRptReader1.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader1.Items[inumber];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then
          Continue;

        aZMDR001RecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aFGStockRptRecordPtr^.snumber);

        if aZMDR001RecordPtr = nil then
        begin
          Memo1.Lines.Add(aFGStockRptRecordPtr^.snumber + '  找不到物料主数据' );
          Continue;
        end;

        aFGStockRptRecordPtr^.ptr := aZMDR001RecordPtr;

        if slCol.IndexOf(aZMDR001RecordPtr^.sColor) < 0 then
        begin
          slCol.Add(aZMDR001RecordPtr^.sColor);
        end;


        if slCap.IndexOf(aZMDR001RecordPtr^.scap) < 0 then
        begin
          slCap.Add(aZMDR001RecordPtr^.scap);
        end;

        if slVer.IndexOf(aZMDR001RecordPtr^.sVer) < 0 then
        begin
          slVer.Add(aZMDR001RecordPtr^.sVer);
        end;
      end;

      for inumber := 0 to slCol.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCol[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctGN, csC);    //国内

        ExcelApp.Cells[irow, 5].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctHW, csR);    //俄罗斯
        ExcelApp.Cells[irow, 6].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctHW, csU);    //乌克兰
        ExcelApp.Cells[irow, 7].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctHW, csI);    //印尼
        ExcelApp.Cells[irow, 8].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader1, ctHW, csO);    //其他

        ExcelApp.Cells[irow, 9].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 10].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 11].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 12].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csO);     //其他
                                 
        ExcelApp.Cells[irow, 13].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csR)
           + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 14].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csU)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 15].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csI)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 16].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csO)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csO);     //其他

        ExcelApp.Cells[irow, 17].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 18].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 19].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 20].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csO);     //其他

        irow := irow + 1;
      end;

      for inumber := 0 to slCap.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCap[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctGN, csC);    //国内

        ExcelApp.Cells[irow, 5].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctHW, csR);    //俄罗斯
        ExcelApp.Cells[irow, 6].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctHW, csU);    //乌克兰
        ExcelApp.Cells[irow, 7].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctHW, csI);    //印尼
        ExcelApp.Cells[irow, 8].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader1, ctHW, csO);    //其他
                                                                      
        ExcelApp.Cells[irow, 9].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 10].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 11].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 12].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csO);     //其他
                    
        ExcelApp.Cells[irow, 13].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csR)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 14].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csU)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 15].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csI)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 16].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csO)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csO);     //其他
        
        ExcelApp.Cells[irow, 17].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 18].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 19].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 20].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csO);     //其他

        irow := irow + 1;
      end;
                                  
      for inumber := 0 to slVer.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slVer[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctGN, csC);    //国内
        
        ExcelApp.Cells[irow, 5].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctHW, csR);    //俄罗斯
        ExcelApp.Cells[irow, 6].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctHW, csU);    //乌克兰
        ExcelApp.Cells[irow, 7].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctHW, csI);    //印尼
        ExcelApp.Cells[irow, 8].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader1, ctHW, csO);    //其他

        ExcelApp.Cells[irow, 9].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 10].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 11].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 12].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csO);     //其他
                          
        ExcelApp.Cells[irow, 13].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csR)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 14].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csU)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 15].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csI)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 16].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csO)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csO);     //其他
            
        ExcelApp.Cells[irow, 17].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 18].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 19].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 20].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csO);     //其他

        irow := irow + 1;
      end;
                        
      slCol.Free;
      slCap.Free;
      slVer.Free;

      MergeCells(ExcelApp, irow_proj1, 1, irow - 1, 1);
      irow_proj1 := irow;
    end;




    for iProj := 0 to aFGStockRptReader2.ProjCount - 1 do
    begin
      sProjNo := aFGStockRptReader2.Projs[iProj];

      slCol := TStringList.Create;
      slCap := TStringList.Create;
      slVer := TStringList.Create;

      sProjName := aSAPMaterialReader.ProjNo2Name(sProjNo);
      ExcelApp.Cells[irow, 1].Value := sProjName;

      for inumber := 0 to aFGStockRptReader2.Count - 1 do
      begin
        aFGStockRptRecordPtr := aFGStockRptReader2.Items[inumber];
        if Copy(aFGStockRptRecordPtr^.snumber, 1, 5) <> sProjNo then
          Continue;

        aZMDR001RecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aFGStockRptRecordPtr^.snumber);

        if aZMDR001RecordPtr = nil then
        begin
          Memo1.Lines.Add(aFGStockRptRecordPtr^.snumber + '  找不到物料主数据' );
          Continue;
        end;

        aFGStockRptRecordPtr^.ptr := aZMDR001RecordPtr;

        if slCol.IndexOf(aZMDR001RecordPtr^.sColor) < 0 then
        begin
          slCol.Add(aZMDR001RecordPtr^.sColor);
        end;


        if slCap.IndexOf(aZMDR001RecordPtr^.scap) < 0 then
        begin
          slCap.Add(aZMDR001RecordPtr^.scap);
        end;

        if slVer.IndexOf(aZMDR001RecordPtr^.sVer) < 0 then
        begin
          slVer.Add(aZMDR001RecordPtr^.sVer);
        end;
      end;

      for inumber := 0 to slCol.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCol[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctGN, csC);    //国内
        
        ExcelApp.Cells[irow, 5].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctHW, csR);    //俄罗斯
        ExcelApp.Cells[irow, 6].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctHW, csU);    //乌克兰
        ExcelApp.Cells[irow, 7].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctHW, csI);    //印尼
        ExcelApp.Cells[irow, 8].Value := GetQty4Color(sProjNo, slCol[inumber], aFGStockRptReader2, ctHW, csO);    //其他

        ExcelApp.Cells[irow, 9].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 10].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 11].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 12].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCol[inumber], mmoMap.Lines, csO);     //其他    
                  
        ExcelApp.Cells[irow, 13].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csR)
           + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 14].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csU)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 15].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csI)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 16].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csO)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csO);     //其他
        
        ExcelApp.Cells[irow, 17].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 18].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 19].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 20].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCol[inumber], mmoMap.Lines, csO);     //其他

        irow := irow + 1;
      end;

      for inumber := 0 to slCap.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCap[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctGN, csC);    //国内   

        ExcelApp.Cells[irow, 5].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctHW, csR);    //俄罗斯
        ExcelApp.Cells[irow, 6].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctHW, csU);    //乌克兰
        ExcelApp.Cells[irow, 7].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctHW, csI);    //印尼
        ExcelApp.Cells[irow, 8].Value := GetQty4Cap(sProjNo, slCap[inumber], aFGStockRptReader2, ctHW, csO);    //其他

        ExcelApp.Cells[irow, 9].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 10].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 11].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 12].Value := GetQtyPIR(aSAPPIRReader, sProjName, slCap[inumber], mmoMap.Lines, csO);     //其他
                    
        ExcelApp.Cells[irow, 13].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csR)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 14].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csU)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 15].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csI)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 16].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csO)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csO);     //其他
        
        ExcelApp.Cells[irow, 17].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 18].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 19].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 20].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slCap[inumber], mmoMap.Lines, csO);     //其他

        irow := irow + 1;
      end;
                                  
      for inumber := 0 to slVer.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slVer[inumber];
        ExcelApp.Cells[irow, 3].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctSF, csC);    //顺丰
        ExcelApp.Cells[irow, 4].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctGN, csC);    //国内

        
        ExcelApp.Cells[irow, 5].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctHW, csR);    //俄罗斯
        ExcelApp.Cells[irow, 6].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctHW, csU);    //乌克兰
        ExcelApp.Cells[irow, 7].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctHW, csI);    //印尼
        ExcelApp.Cells[irow, 8].Value := GetQty4Ver(sProjNo, slVer[inumber], aFGStockRptReader2, ctHW, csO);    //其他

        ExcelApp.Cells[irow, 9].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 10].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 11].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 12].Value := GetQtyPIR(aSAPPIRReader, sProjName, slVer[inumber], mmoMap.Lines, csO);     //其他   
                 
        ExcelApp.Cells[irow, 13].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csR)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 14].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csU)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 15].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csI)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 16].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csO)
          + GetHWOODQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csO);     //其他
        
        ExcelApp.Cells[irow, 17].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csR);     //俄罗斯
        ExcelApp.Cells[irow, 18].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csU);     //乌克兰
        ExcelApp.Cells[irow, 19].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csI);     //印尼
        ExcelApp.Cells[irow, 20].Value := GetHWOrdInfoQty(aHWOrdInfoReader, sProjName, slVer[inumber], mmoMap.Lines, csO);     //其他

        irow := irow + 1;
      end;
                        
      slCol.Free;
      slCap.Free;
      slVer.Free;

      MergeCells(ExcelApp, irow_proj1, 1, irow - 1, 1);
      irow_proj1 := irow;
    end;

    AddBorder(ExcelApp, 1, 1, irow - 1, 20);

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    WorkBook.Close;
    ExcelApp.Quit;
  finally
    aFGStockRptReader1.Free;
    aFGStockRptReader2.Free;
    aSAPMaterialReader.Free;
    aFGRptWinReader1.Free;
    aFGRptWinReader2.Free;
    aFGToProduceReader.Free;
    aSAPPIRReader.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

procedure TfrmSOPSum.btnFGReport1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leFGReport1.Text := sfile;
end;

procedure TfrmSOPSum.btnFGReport2Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leFGReport2.Text := sfile;
end;

procedure TfrmSOPSum.btnPIR_HWClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  lePIR_HW.Text := sfile;
end;

procedure TfrmSOPSum.btnHWODInfoClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leHWODInfo.Text := sfile;
end;

end.
