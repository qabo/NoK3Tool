unit MRPSimulationWin;

(*
1�� ������������������ϣ�SAP ZPPR020 ��������ȡ��Load���Ϲ�Ӧ��ʱ����Ҫ��Ҫ��ȡ���


*)

(*
ϵͳ����ģ��MRP������ģ�ⲽ�裺
  1�������������Ҫ���ƻ�  ZPPR028
  2������BOM ZPPR021�� �ϺŴӵ������Ҫ���ƻ�ȡ
  3��������� MB52
  4��BOMתSBOM  ��������
  5���������Ҫ���ƻ�תS&OP��ʽ��ע��SKU��Ҫ�ظ�
  6��SKU���ȼ����㣬 SKU���ȼ��ļ���ĲŻ�������׷���
  7��ģ��MRP
  8�����׷��� *******8
*)

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, StdCtrls, ExtCtrls, IniFiles, CommUtils,
  DB, ADODB, Provider, DBClient, SOPReaderUnit, SAPBom2SBomWin, SAPS618Reader, 
  DateUtils, ComObj, ExcelConsts, MrpMPSReader, NewSKUReader, LTP_CMS2MRPSimWin,
  KeyICItemSupplyReader, SBomReader, SOPSimReader, FGPriorityReader, Clipbrd,
  ProjYearWin, DOSReader, MRPWinReader, jpeg, ImgList, LTPCMSConfirmReader,
  SAPStockReader, SAPBomReader, SAPBomReader2, SAPMaterialReader, SAPMaterialReader2;

type
  TDeltaRecord = packed record
    sproj: string;
    snumber: string;
    sdate: string;
    iqty: Double;
    iqty_org: Double;
  end;
  PDeltaRecord = ^TDeltaRecord;
                                                             
  TfrmMRPSimulation = class(TForm)
    ToolBar1: TToolBar;
    tbClose: TToolButton;
    ToolButton5: TToolButton;
    Memo1: TMemo;
    StatusBar1: TStatusBar;
    PageControl1: TPageControl;
    TabSheet5: TTabSheet;
    leLastMPS: TLabeledEdit;
    leLastSOPAnalysis: TLabeledEdit;
    leSupplyEval: TLabeledEdit;
    leSBOM: TLabeledEdit;
    leDemand5: TLabeledEdit;
    btnLastMPS: TButton;
    btnLastSOPAnalysis: TButton;
    btnSupplyEval: TButton;
    btnSBOM: TButton;
    btnDemand: TButton;
    ProgressBar1: TProgressBar;
    ToolButton1: TToolButton;
    btnSim: TButton;
    TabSheet2: TTabSheet;
    leDemand_p4: TLabeledEdit;
    btnDemand2: TButton;
    leNewSKU_p4: TLabeledEdit;
    btnNewSKU: TButton;
    leDOS_p4: TLabeledEdit;
    btnDOS: TButton;
    btnPriority: TButton;
    lePriority: TLabeledEdit;
    btnCalcPriority: TButton;
    Button7: TButton;
    Button8: TButton;
    Button5: TButton;
    Button6: TButton;
    tbProjYear: TToolButton;
    Label5: TLabel;
    dtpDemandBeginDate1: TDateTimePicker;
    Label6: TLabel;
    Label7: TLabel;
    dtpDemandBeginDate4: TDateTimePicker;
    Label8: TLabel;
    ImageList1: TImageList;
    TabSheet1: TTabSheet;
    Button1: TButton;
    btnCMS2MRPSim: TButton;
    btnCMSConfirm: TButton;
    leCMSConfirm: TLabeledEdit;
    leSAPBom: TLabeledEdit;
    btnSAPBom: TButton;
    leSAPStock: TLabeledEdit;
    btnSAPStock: TButton;
    leSBOM1: TLabeledEdit;
    btnSBOM1: TButton;
    leStock: TLabeledEdit;
    btnStock: TButton;
    TabSheet3: TTabSheet;
    leSAPStock3: TLabeledEdit;
    btnSAPStock3: TButton;
    leSAPBom3: TLabeledEdit;
    btnSAPBom3: TButton;
    leDemand3: TLabeledEdit;
    btnDemand3: TButton;
    btnMRP: TButton;
    Memo3: TMemo;
    Label1: TLabel;
    dtpDemandBeginDate3: TDateTimePicker;
    Label2: TLabel;
    Memo2: TMemo;
    leICItem: TLabeledEdit;
    btnICItem: TButton;
    btnSim2: TButton;
    tbSIM2SAP: TTabSheet;
    leSIM2SAP5: TLabeledEdit;
    btnSIM2SAP5: TButton;
    btnSIM2SAP5_save: TButton;
    procedure FormCreate(Sender: TObject); 
    procedure FormDestroy(Sender: TObject);
    procedure tbCloseClick(Sender: TObject);
    procedure btnPrev1Click(Sender: TObject);
    procedure btnNext1Click(Sender: TObject);
    procedure btnLastMPSClick(Sender: TObject);
    procedure btnLastSOPAnalysisClick(Sender: TObject);
    procedure btnSupplyEvalClick(Sender: TObject);
    procedure btnSBOMClick(Sender: TObject);
    procedure btnSimClick(Sender: TObject);
    procedure btnPriorityClick(Sender: TObject);
    procedure btnDemand2Click(Sender: TObject);
    procedure btnNewSKUClick(Sender: TObject);
    procedure btnDOSClick(Sender: TObject);
    procedure btnCalcPriorityClick(Sender: TObject);
    procedure tbProjYearClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnCMS2MRPSimClick(Sender: TObject);
    procedure btnCMSConfirmClick(Sender: TObject);
    procedure btnSAPBomClick(Sender: TObject);
    procedure btnSAPStockClick(Sender: TObject);
    procedure btnSBOM1Click(Sender: TObject);
    procedure btnStockClick(Sender: TObject);
    procedure btnMRPClick(Sender: TObject);
    procedure btnSAPStock3Click(Sender: TObject);
    procedure btnSAPBom3Click(Sender: TObject);
    procedure btnDemand3Click(Sender: TObject);
    procedure btnICItemClick(Sender: TObject);
    procedure btnSim2Click(Sender: TObject);
    procedure btnSIM2SAP5Click(Sender: TObject);
    procedure btnSIM2SAP5_saveClick(Sender: TObject);
  private
    { Private declarations } 
    procedure SaveMPS(const sfile_save: string; aSOPReader: TSOPReader;
      aSOPSimReader: TSOPSimReader; aMrpMPSReader: TMrpMPSReader;
      aKeyICItemSupplyReader: TKeyICItemSupplyReader);  
    procedure SaveSAP(const sfile_save: string; aSAPS618Reader: TSAPS618Reader);
    procedure OnLogEvent(const s: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation
 
{$R *.dfm}

type
  TNumberInfo = packed record
    scategory: string; //�������
    sg: string; //ͨ����
    snumber: string; //���ϱ���(����)
    sname: string; //��������
    slt: string; //LT
    smc: string; //MC
    sproj: string; //��Ŀ
  end;
  PNumberInfo = ^TNumberInfo;
    
  TOrderByConditions = packed record
    isnew: Boolean;
    dos: Double;
    demand: Double;
  end;
  POrderByConditions = ^TOrderByConditions;

  function StringListSortCompare(List: TStringList; Index1, Index2: Integer): Integer;
  var
    p1, p2: POrderByConditions;
  begin
    Result := 0;
    p1 := POrderByConditions(List.Objects[Index1]);
    p2 := POrderByConditions(List.Objects[Index2]);

    if p1^.isnew <> p2^.isnew then
    begin
      if p1^.isnew then
        Result := -1
      else Result := 1;
    end
    else
    begin
      if p1^.dos <> p2.dos then
      begin
        if p1^.dos > p2^.dos then
          Result := 1
        else Result := -1;
      end
      else
      begin
        if p1^.demand <> p2^.demand then
        begin
          if p1^.demand > p2^.demand then
            Result := -1
          else Result := 1;
        end;
      end;
    end; 
  end;

class procedure TfrmMRPSimulation.ShowForm;
var
  frmMRPSimulation: TfrmMRPSimulation;
begin
  frmMRPSimulation := TfrmMRPSimulation.Create(nil);
  try
    frmMRPSimulation.ShowModal;
  finally
    frmMRPSimulation.Free;
  end;
end;
   
procedure TfrmMRPSimulation.FormCreate(Sender: TObject);
var
  sfile: string;
  ini: TIniFile;
  sdt: string;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);

  sdt := ini.ReadString(self.Name, dtpDemandBeginDate4.Name, FormatDateTime('yyyy-MM-dd', Now));
  dtpDemandBeginDate4.DateTime := myStrToDateTime(sdt);
  leLastSOPAnalysis.Text := ini.ReadString(self.Name, leLastSOPAnalysis.Name, '');
  leDemand5.Text := ini.ReadString(self.Name, leDemand5.Name, '');
  lePriority.Text := ini.ReadString(self.Name, lePriority.Name, '');
  leSupplyEval.Text := ini.ReadString(self.Name, leSupplyEval.Name, '');
  leLastMPS.Text := ini.ReadString(self.Name, leLastMPS.Name, '');
  leSBOM.Text := ini.ReadString(self.Name, leSBOM.Name, '');


  leDemand_p4.Text := ini.ReadString(self.Name, leDemand_p4.Name, '');
  leNewSKU_p4.Text := ini.ReadString(self.Name, leNewSKU_p4.Name, '');
  leDOS_p4.Text := ini.ReadString(self.Name, leDOS_p4.Name, '');
  dtpDemandBeginDate1.DateTime := myStrToDateTime( StringReplace( ini.ReadString( self.Name, dtpDemandBeginDate1.Name, FormatDateTime('yyyy-MM-dd', Now) ), '/', '-', [rfReplaceAll] ));

  leSAPStock.Text := ini.ReadString(self.Name, leSAPStock.Name, '');
  leSAPBom.Text := ini.ReadString(self.Name, leSAPBom.Name, '');           
  leSBOM1.Text := ini.ReadString(self.Name, leSBOM1.Name, '');
  leCMSConfirm.Text := ini.ReadString(self.Name, leCMSConfirm.Name, '');
  leStock.Text := ini.ReadString(self.Name, leStock.Name, '');

  leSAPStock3.Text := ini.ReadString(self.Name, leSAPStock3.Name, '');
  leSAPBom3.Text := ini.ReadString(self.Name, leSAPBom3.Name, '');
  leDemand3.Text := ini.ReadString(self.Name, leDemand3.Name, '');
  dtpDemandBeginDate3.DateTime := myStrToDateTime( StringReplace( ini.ReadString( self.Name, dtpDemandBeginDate3.Name, FormatDateTime('yyyy-MM-dd', Now) ), '/', '-', [rfReplaceAll] ));

  leICItem.Text := ini.ReadString(self.Name, leICItem.Name, '');

  leSIM2SAP5.Text := ini.ReadString(self.Name, leSIM2SAP5.Name, '');

  ini.Free;




//  PageControl1.ActivePageIndex := 1;
end;

procedure TfrmMRPSimulation.FormDestroy(Sender: TObject);
var
  sfile: string;
  ini: TIniFile;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);

  ini.WriteDateTime(self.Name, dtpDemandBeginDate4.Name, dtpDemandBeginDate4.DateTime);
  ini.WriteString(self.Name, leLastSOPAnalysis.Name, leLastSOPAnalysis.Text);
  ini.WriteString(self.Name, leDemand5.Name, leDemand5.Text);
  ini.WriteString(self.Name, lePriority.Name, lePriority.Text);
  ini.WriteString(self.Name, leSupplyEval.Name, leSupplyEval.Text);
  ini.WriteString(self.Name, leLastMPS.Name, leLastMPS.Text);
  ini.WriteString(self.Name, leSBOM.Name, leSBOM.Text);

  ini.WriteString(self.Name, leDemand_p4.Name, leDemand_p4.Text);
  ini.WriteString(self.Name, leNewSKU_p4.Name, leNewSKU_p4.Text);
  ini.WriteString(self.Name, leDOS_p4.Name, leDOS_p4.Text);
  ini.WriteString(self.Name, dtpDemandBeginDate1.Name, FormatDateTime('yyyy-MM-dd', dtpDemandBeginDate1.DateTime));

  ini.WriteString(self.Name, leSAPStock.Name, leSAPStock.Text);
  ini.WriteString(self.Name, leSAPBom.Name, leSAPBom.Text);
  ini.WriteString(self.Name, leSBOM1.Name, leSBOM1.Text);
  ini.WriteString(self.Name, leCMSConfirm.Name, leCMSConfirm.Text);
  ini.WriteString(self.Name, leStock.Name, leStock.Text);

  ini.WriteString(self.Name, leSAPStock3.Name, leSAPStock3.Text);
  ini.WriteString(self.Name, leSAPBom3.Name, leSAPBom3.Text);
  ini.WriteString(self.Name, leDemand3.Name, leDemand3.Text);
  ini.WriteString(self.Name, dtpDemandBeginDate3.Name, FormatDateTime('yyyy-MM-dd', dtpDemandBeginDate3.DateTime));

  ini.WriteString(self.Name, leICItem.Name, leICItem.Text);

  ini.WriteString(self.Name, leSIM2SAP5.Name, leSIM2SAP5.Text);

  ini.Free;
end;
 
function ExtractSOPDate(const sdate: string): TDateTime;
var
  s: string;
begin
  try
    s := sdate;
    if Pos('/', s) > 0 then
    begin
      if Pos('-', s) > 0 then
      begin
        s := Copy(s, 1, Pos('-', s) - 1);
      end;
      s := '2017-' + StringReplace(s, '/', '-', [rfReplaceAll]);
    end;
    Result := myStrToDateTime(s);
  except
    Result := 0;
  end;
end;  

function GetNearestMPS(aMPSs: TList; aDeltaPtr: PDeltaRecord): PDeltaRecord;
var
  iDist: Double;
  i: Integer;
  p: PDeltaRecord;
  dt1, dt2: TDateTime;
begin
  Result := nil;
  iDist := MaxInt;
  for i := 0 to aMPSs.Count - 1 do
  begin
    p := PDeltaRecord(aMPSs[i]);
    if p^.iqty = 0 then Continue;
    if p^.snumber = aDeltaPtr^.snumber then
    begin
      dt1 := ExtractSOPDate(p^.sdate);
      dt2 := ExtractSOPDate(aDeltaPtr^.sdate);
      if iDist > Abs(dt1 - dt2) then
      begin
        iDist := Abs(dt1 - dt2);
        Result := p;
      end;
    end;
  end;  
end;
 
procedure TfrmMRPSimulation.tbCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMRPSimulation.btnPrev1Click(Sender: TObject);
begin
  if PageControl1.ActivePageIndex > 0 then
    PageControl1.ActivePageIndex := PageControl1.ActivePageIndex - 1;
end;

procedure TfrmMRPSimulation.btnNext1Click(Sender: TObject);
begin
  if PageControl1.ActivePageIndex < PageControl1.PageCount - 1 then
    PageControl1.ActivePageIndex := PageControl1.ActivePageIndex + 1;
end;

procedure TfrmMRPSimulation.btnLastMPSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leLastMPS.Text := sfile;
end;

procedure TfrmMRPSimulation.btnLastSOPAnalysisClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leLastSOPAnalysis.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSupplyEvalClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSupplyEval.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSBOMClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSBOM.Text := sfile;
end;

procedure TfrmMRPSimulation.SaveMPS(const sfile_save: string; aSOPReader: TSOPReader;
  aSOPSimReader: TSOPSimReader; aMrpMPSReader: TMrpMPSReader;
  aKeyICItemSupplyReader: TKeyICItemSupplyReader);
var
  ExcelApp, WorkBook: Variant;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  irow: Integer;
  icol: Integer;
  dt0: TDateTime;
  icol1: Integer;  
  icol2: Integer;

//  aSOPSimProj: TSOPSimProj;
//  aSOPSimLine: TSOPSimLine;

  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
  aKeyICItemSupplyDate: TKeyICItemSupplyDate;
begin

  Memo1.Lines.Add('��ʼ����EXCEL ');
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '�����ʾ', 0);
      Exit;
    end;
  end;

  icol2 := 1;

   
  WorkBook := ExcelApp.WorkBooks.Add;
  ExcelApp.DisplayAlerts := False;


  try
    while ExcelApp.Sheets.Count < aSOPReader.FProjs.Count + 1 do
    begin
      ExcelApp.Sheets.Add;
    end;


    for iproj := 0 to aSOPReader.FProjs.Count - 1 do
    begin
      aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[iproj]);

      Memo1.Lines.Add(aSOPProj.FName + ' sheet');
 
      ExcelApp.Sheets[iproj + 1].Activate;
      ExcelApp.Sheets[iproj + 1].Name := aSOPProj.FName;    

      
      irow := 1;                             
      ExcelApp.Cells[irow, 1].Value := 'MRP����';
      ExcelApp.Cells[irow, 2].Value := '��Ŀ';
      ExcelApp.Cells[irow, 3].Value := '���ϱ���';
      ExcelApp.Cells[irow, 4].Value := '��׼��ʽ';
      ExcelApp.Cells[irow, 5].Value := '��ɫ';
      ExcelApp.Cells[irow, 6].Value := '����';
      ExcelApp.Cells[irow, 7].Value := '������';
            

      MergeCells(ExcelApp, irow, 1, irow + 1, 1);
      MergeCells(ExcelApp, irow, 2, irow + 1, 2);
      MergeCells(ExcelApp, irow, 3, irow + 1, 3);
      MergeCells(ExcelApp, irow, 4, irow + 1, 4);
      MergeCells(ExcelApp, irow, 5, irow + 1, 5);
      MergeCells(ExcelApp, irow, 6, irow + 1, 6);  
      MergeCells(ExcelApp, irow, 7, irow + 1, 7);
                                              
      ExcelApp.Columns[1].ColumnWidth := 8; 
      ExcelApp.Columns[2].ColumnWidth := 8;
      ExcelApp.Columns[3].ColumnWidth := 13;
      ExcelApp.Columns[4].ColumnWidth := 13;
      ExcelApp.Columns[5].ColumnWidth := 8;
      ExcelApp.Columns[6].ColumnWidth := 8;
      ExcelApp.Columns[7].ColumnWidth := 23;

      AddColor(ExcelApp, irow, 1, irow + 1, 7, $00CC99);
 
      irow := 3; 
      for iline := 0 to aSOPProj.FList.Count - 1 do
      begin
        aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);
                                                                
        ExcelApp.Cells[irow, 1].Value := aSOPLine.sMRPArea;
        ExcelApp.Cells[irow, 2].Value := aSOPLine.sProj;
        ExcelApp.Cells[irow, 3].Value := aSOPLine.sNumber;
        ExcelApp.Cells[irow, 4].Value := aSOPLine.sVer;
        ExcelApp.Cells[irow, 5].Value := aSOPLine.sColor;
        ExcelApp.Cells[irow, 6].Value := aSOPLine.sCap;

        ExcelApp.Cells[irow,7].Value := 'Ҫ���ƻ���';
        ExcelApp.Cells[irow + 1, 7].Value := '�ɹ�Ӧ��';
        ExcelApp.Cells[irow + 2, 7].Value := 'Ҫ���ƻ���ɹ�Ӧ������';
        ExcelApp.Cells[irow + 3, 7].Value := 'Ƿ��';

                            
        MergeCells(ExcelApp, irow, 1, irow + 3, 1);
        MergeCells(ExcelApp, irow, 2, irow + 3, 2);
        MergeCells(ExcelApp, irow, 3, irow + 3, 3);
        MergeCells(ExcelApp, irow, 4, irow + 3, 4);
        MergeCells(ExcelApp, irow, 5, irow + 3, 5);  
        MergeCells(ExcelApp, irow, 6, irow + 3, 6);

 
        dt0 := 0;     
        icol := 8;

        if iline = 0 then  // ��һ��SKU���룬д�б���
        begin


          // д������ /////////////////////////////////////////
          for idate := 0 to aSOPLine.FList.Count - 1 do
          begin
            aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
            if aSOPCol.dt1 < dtpDemandBeginDate4.DateTime then Continue;

            if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
            begin
              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(dt0)) + '��';
              MergeCells(ExcelApp, 1, icol, 2, icol);
              AddColor(ExcelApp, 1, icol, 2, icol, $CCFFFF);
              icol := icol + 1;
            end;
               
            ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
            ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;   
            icol := icol + 1;

            // ���һ������
            if idate = aSOPLine.FList.Count - 1 then
            begin
              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(aSOPCol.dt1)) + '��';
              MergeCells(ExcelApp, 1, icol, 2, icol);        
              AddColor(ExcelApp, 1, icol, 2, icol, $CCFFFF);

              icol2 := icol;
            end;

            dt0 := aSOPCol.dt1;
          end;
        end;
        
                         
        dt0 := 0;     
        icol := 8;
        icol1 := icol;

        for idate := 0 to aSOPLine.FList.Count - 1 do
        begin
          aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);    
          if aSOPCol.dt1 < dtpDemandBeginDate4.DateTime then Continue;

          if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow); 
            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1); 
            ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
//            ExcelApp.Cells[irow + 3, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 3) + ':' + GetRef(icol - 1) + IntToStr(irow + 3);
//            ExcelApp.Cells[irow + 4, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 1);
//            ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 5) + ':' + GetRef(icol - 1) + IntToStr(irow + 5);
//            ExcelApp.Cells[irow + 6, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 5) + '-' + GetRef(icol) + IntToStr(irow + 1);
            icol := icol + 1;
            icol1 := icol;
          end;
            
          ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;
          ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.iQty_ok;
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
//          if aSOPSimLine <> nil then
//          begin
//            ExcelApp.Cells[irow + 3, icol].Value := aSOPSimLine.GetQty(aSOPCol.dt1);
//          end;
//          ExcelApp.Cells[irow + 4, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 1);
//          if aMrpMPSReader <> nil then
//          begin
//            ExcelApp.Cells[irow + 5, icol].Value := aMrpMPSReader.GetQty(aSOPLine.sNumber, aSOPCol.dt1, aSOPCol.dt2);
//          end;
//          ExcelApp.Cells[irow + 6, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 5) + '-' + GetRef(icol) + IntToStr(irow + 1);
          ExcelApp.Cells[irow + 3, icol].Value := aSOPCol.sShortageICItem;

          icol := icol + 1;

          // ���һ������
          if idate = aSOPLine.FList.Count - 1 then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow);
            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1);
            ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
//            ExcelApp.Cells[irow + 3, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 3) + ':' + GetRef(icol - 1) + IntToStr(irow + 3);
//            ExcelApp.Cells[irow + 4, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 1);
//            ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 5) + ':' + GetRef(icol - 1) + IntToStr(irow + 5);
//            ExcelApp.Cells[irow + 6, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 5) + '-' + GetRef(icol) + IntToStr(irow + 1);
          end;
          
          dt0 := aSOPCol.dt1;

        end;

//        AddColor(ExcelApp, irow + 3, 6, irow + 4, icol2, $D6E4FC);     
//        AddColor(ExcelApp, irow + 5, 6, irow + 6, icol2, $99FFFF);

        irow := irow + 4;
      end;

      AddBorder(ExcelApp, 1, 1, irow - 1, icol2);
          
      ExcelApp.Range[ ExcelApp.Cells[3, 8], ExcelApp.Cells[3, 8] ].Select;
      ExcelApp.ActiveWindow.FreezePanes := True;

    end;

    Memo1.Lines.Add('���Ϲ�Ӧ����sheet');
    // �� ���Ϲ�Ӧ �ķ�������ʾ
    ExcelApp.Sheets[aSOPReader.FProjs.Count + 1].Activate;
    ExcelApp.Sheets[aSOPReader.FProjs.Count + 1].Name := '�ؼ����Ϲ�Ӧ';

    irow := 3;
    for iline := 0 to aKeyICItemSupplyReader.FList.Count - 1 do
    begin
      aKeyICItemSupplyLine := TKeyICItemSupplyLine(aKeyICItemSupplyReader.FList.Objects[iline]);

      if iline = 0 then
      begin
        ExcelApp.Cells[1, 1].Value := '���ϱ���';   
        ExcelApp.Cells[1, 2].Value := '��������';
        ExcelApp.Cells[1, 3].Value := '������';

        ExcelApp.Cells[1, 4].Value := '����';
        ExcelApp.Cells[1, 5].Value := '���Ʒ���';
        ExcelApp.Cells[1, 6].Value := '���ÿ��';
        ExcelApp.Cells[1, 7].Value := '���⹺���';
        ExcelApp.Cells[1, 8].Value := 'MRP�����⹺���';

        MergeCells(ExcelApp, 1, 1, 2, 1);
        MergeCells(ExcelApp, 1, 2, 2, 2); 
        MergeCells(ExcelApp, 1, 3, 2, 3);
        MergeCells(ExcelApp, 1, 4, 2, 4); 
        MergeCells(ExcelApp, 1, 5, 2, 5);
        MergeCells(ExcelApp, 1, 6, 2, 6);
        MergeCells(ExcelApp, 1, 7, 2, 7);
        MergeCells(ExcelApp, 1, 8, 2, 8);

        for idate := 0 to aKeyICItemSupplyLine.FList.Count - 1 do
        begin
          aKeyICItemSupplyDate := TKeyICItemSupplyDate(aKeyICItemSupplyLine.FList[idate]);
          ExcelApp.Cells[1, idate + 8 + 1].Value := aKeyICItemSupplyDate.sweek;
          ExcelApp.Cells[2, idate + 8 + 1].Value := aKeyICItemSupplyDate.sdate;
        end;

        icol2 := aKeyICItemSupplyLine.FList.Count + 8;
      end;     

      ExcelApp.Cells[irow, 1].Value := aKeyICItemSupplyLine.sNumber99;
      ExcelApp.Cells[irow, 2].Value := aKeyICItemSupplyLine.sName;
      ExcelApp.Cells[irow, 3].Value := aKeyICItemSupplyLine.sNumber;
      ExcelApp.Cells[irow, 4].Value := '������';
      ExcelApp.Cells[irow + 1, 4].Value := '������';
      ExcelApp.Cells[irow, 5].Value := aKeyICItemSupplyLine.sStock_semi;
      ExcelApp.Cells[irow, 6].Value := aKeyICItemSupplyLine.sStock;
      ExcelApp.Cells[irow, 7].Value := aKeyICItemSupplyLine.sWin_new;
      ExcelApp.Cells[irow, 8].Value := aKeyICItemSupplyLine.sWin_day_of_mrp;
        
      MergeCells(ExcelApp, 1, 1, 2, 1);
      MergeCells(ExcelApp, 1, 2, 2, 2);
      MergeCells(ExcelApp, 1, 3, 2, 3);
        
      MergeCells(ExcelApp, 1, 5, 2, 5);

      ExcelApp.Cells[irow + 1, 6].Value := aKeyICItemSupplyLine.dqty_alloc;

      MergeCells(ExcelApp, 1, 7, 2, 7);
      MergeCells(ExcelApp, 1, 8, 2, 8);

      for idate := 0 to aKeyICItemSupplyLine.FList.Count - 1 do
      begin
        aKeyICItemSupplyDate := TKeyICItemSupplyDate(aKeyICItemSupplyLine.FList[idate]);
        ExcelApp.Cells[irow, idate + 8 + 1].Value := aKeyICItemSupplyDate.ConfirmSupply;
        ExcelApp.Cells[irow + 1, idate + 8 + 1].Value := aKeyICItemSupplyDate.Alloc;
      end;

      irow := irow + 2;
    end;
    AddBorder(ExcelApp, 1, icol2, irow - 1, icol2);
 

    ExcelApp.Sheets[1].Activate;
      
    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end;
  Memo1.Lines.Add('����EXCEL ����');  
end;        

procedure TfrmMRPSimulation.SaveSAP(const sfile_save: string;
  aSAPS618Reader: TSAPS618Reader);
var
  ExcelApp, WorkBook: Variant;
  iline: Integer;
  aSAPS618: TSAPS618;
  idate: Integer;
  aSAPS618ColPtr: PSAPS618Col;
  irow: Integer;
  icol: Integer; 
begin

  Memo1.Lines.Add('��ʼ����EXCEL ');
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '�����ʾ', 0);
      Exit;
    end;
  end;

 
   
  WorkBook := ExcelApp.WorkBooks.Add;
  ExcelApp.DisplayAlerts := False;


  try
    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;
                               
    ExcelApp.Sheets[1].Activate;
    ExcelApp.Sheets[1].Name := '��Ӧ����';
  
    irow := 1; 
    ExcelApp.Cells[irow, 1].Value := 'MATNR';
    ExcelApp.Cells[irow, 2].Value := 'BERID';      
    ExcelApp.Cells[irow, 3].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 4].Value := '��Ŀ'; 

    irow := 2;
    for iline := 0 to aSAPS618Reader.Count - 1 do
    begin
      aSAPS618 := aSAPS618Reader.Items[iline];
        
      ExcelApp.Cells[irow, 1].Value := aSAPS618.FNumber;
      ExcelApp.Cells[irow, 2].Value := aSAPS618.FMrpArea;   
      ExcelApp.Cells[irow, 3].Value := aSAPS618.sname;

      icol := 5;
      if iline = 0 then  // ��һ��SKU���룬д�б���
      begin 
        // д������ /////////////////////////////////////////
        for idate := 0 to aSAPS618.Count - 1 do
        begin
          aSAPS618ColPtr := aSAPS618.Items[idate];
          if aSAPS618ColPtr^.dt1 < dtpDemandBeginDate4.DateTime then Continue;
 
          ExcelApp.Cells[1, icol].Value := FormatDateTime('YYYY', aSAPS618ColPtr^.dt1) + Copy( IntToStr(100 + WeekOfTheYear(aSAPS618ColPtr^.dt1)), 2, 2);   
          icol := icol + 1;

        end;
      end;
        
                       
      icol := 5;  
      for idate := 0 to aSAPS618.Count - 1 do
      begin
        aSAPS618ColPtr := aSAPS618.Items[idate];
        if aSAPS618ColPtr^.dt1 < dtpDemandBeginDate4.DateTime then Continue;
 
        ExcelApp.Cells[irow, icol].Value := aSAPS618ColPtr^.dQty;
        ExcelApp.Cells[irow + 1, icol].Value := aSAPS618ColPtr^.dQty_ok;   
        ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow); 
        ExcelApp.Cells[irow + 3, icol].Value := aSAPS618ColPtr^.sShortageICItem;


        icol := icol + 1;

      end;

      ExcelApp.Cells[irow, 4].Value := '�������Ҫ�� �ƻ�';
      ExcelApp.Cells[irow + 1, 4].Value := '��Ӧ����';
      ExcelApp.Cells[irow + 2, 4].Value := '����';
      ExcelApp.Cells[irow + 3, 4].Value := 'Ƿ��';

      MergeCells(ExcelApp, irow, 1, irow + 4, 1);
      MergeCells(ExcelApp, irow, 2, irow + 4, 2);
      MergeCells(ExcelApp, irow, 3, irow + 4, 3);
 
      irow := irow + 4;
    end;

  
    ExcelApp.Sheets[1].Activate;
      
    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end;
  Memo1.Lines.Add('����EXCEL ����');
end;

procedure TfrmMRPSimulation.btnSimClick(Sender: TObject);
  procedure AddShortageICItem(p: PSAPS618Col; const smsg: string);
  begin
    if p^.sShortageICItem <> '' then
      p^.sShortageICItem := p^.sShortageICItem + #13#10'';
    p^.sShortageICItem := p^.sShortageICItem + smsg;
  end;
var 
  aKeyICItemSupplyReader: TKeyICItemSupplyReader;

  aSBomReader: TSBomReader;
 
  aSAPS618Reader: TSAPS618Reader;
  slProjYear: TStringList;
  idate: Integer;
  sldate: TStringList;
  aSAPS618ColPtr: PSAPS618Col;

  inumber: Integer;
  iper: Integer;

  aSAPS618ColPtr_demand: PSAPS618Col;
  aSBom: TSBom;
  aSBomChild: TSBomChild;

  ibom: Integer;
  ichild: Integer;

  dqty: Double;
  dqty_child_min: Double;
  dqty_a: Double;

  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
 
  sfile_save: string;
 
  aFGPriorityReader: TFGPriorityReader;
  lstDemand: TList;
  iDemand: Integer;
  smsg: string;

  dtMemand: TDateTime;
  dDemandQty: Double;

  igroup: Integer;
  
  aSAPStockReader: TSAPStockReader;

 
begin
  sfile_save := 'SOP ' + FormatDateTime('yyyyMMdd-hhmmss', Now); // 20170705-093325
  if not ExcelSaveDialog(sfile_save) then Exit;

  dtpDemandBeginDate4.DateTime := EncodeDateTime(YearOf(dtpDemandBeginDate4.DateTime),
    MonthOf(dtpDemandBeginDate4.DateTime), DayOf(dtpDemandBeginDate4.DateTime), 0, 0, 0, 0);

  lstDemand := TList.Create;

  Memo1.Lines.Add('��ȡ��Ŀ��� ......');
  slProjYear := TfrmProjYear.GetProjYears;
  sldate := TStringList.Create;

  Memo1.Lines.Add('��ȡ�ؼ����Ϲ�Ӧ���� ......');
  aKeyICItemSupplyReader := TKeyICItemSupplyReader.Create(leSupplyEval.Text);

  Memo1.Lines.Add('��ȡ����BOM ......');
  aSBomReader := TSBomReader.Create(leSBOM.Text);
          
  Memo3.Lines.Add('��ʼ��ȡ Ҫ���ƻ�  �������Ҫ���ƻ�');
  aSAPS618Reader := TSAPS618Reader.Create(leDemand5.Text, '�������Ҫ���ƻ�', OnLogEvent);


  Memo1.Lines.Add('��ȡSKU���ȼ� ......');
  aFGPriorityReader := TFGPriorityReader.Create(lePriority.Text);
                                          
  Memo1.Lines.Add('��ȡSAP��� ......');
  aSAPStockReader := TSAPStockReader.Create(leStock.Text);

  aKeyICItemSupplyReader.SAPStock := aSAPStockReader; // ��Щitem�п�棬����û�ó��ɹ����� �ظ���������û������ϡ����Ƿ�����ʱ��û������ϣ��������Ϊ�ɹ�ӦΪ0.���Ի�ȡ���Ϲ�Ӧ��ʱ��û�н��ڵģ�ȡ���
         
  try
                                        
    Memo1.Lines.Add('�� Bom ���������ϣ� ����Ӧ������ϵ�� ......');
    // �� Bom ���������ϣ� ����Ӧ������ϵ��
    for ibom := 0 to aSBomReader.FList.Count - 1 do
    begin
      aSBom := TSBom(aSBomReader.FList.Objects[ibom]);
      for ichild := 0 to aSBom.FList.Count - 1 do
      begin
        aSBomChild := TSBomChild(aSBom.FList.Objects[ichild]);
        for igroup := 0 to aSBomChild.FList.Count - 1 do
        begin
          PBomGroupChild(aSBomChild.FList.Objects[igroup])^.supp := aKeyICItemSupplyReader.GetSupplyLine(aSBomChild.FList[igroup]);
        end;
      end;
    end;
                        
    Memo1.Lines.Add('ƥ���Ʒ�Ϻ�Bom ......');
    // ƥ���Ʒ�Ϻ�Bom
    for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
    begin
      aFGPriorityReader.FList.Objects[inumber] := aSBomReader.GetBom(aFGPriorityReader.FList[inumber]);
      if aFGPriorityReader.FList.Objects[inumber] = nil then
      begin
        Memo1.Lines.Add(aFGPriorityReader.FList[inumber] + '  û��BOM');
      end;
    end;
    
    Memo1.Lines.Add('ȡ�����б� ......');
    // ȡ�����б�
//    aSOPReader.GetDateList(sldate);
    aSAPS618Reader.GetDateList(sldate);

    Memo1.Lines.Add('ÿһ�����ڼ��� ......');
    // ÿһ�����ڼ���
    for idate := 0 to sldate.Count - 1 do
    begin     
      aSAPS618ColPtr := PSAPS618Col(sldate.Objects[idate]); // ����
      if aSAPS618ColPtr^.dt1 < dtpDemandBeginDate4.DateTime then Continue;
 
      dtMemand := aSAPS618ColPtr^.dt2 - 2;
      if dtMemand < aSAPS618ColPtr^.dt1 then
        dtMemand := aSAPS618ColPtr^.dt1;

      Memo1.Lines.Add( FormatDateTime('yyyy-MM-dd', dtMemand) + '    ' + aSAPS618ColPtr^.sweek + ' ......');
    
      // ÿһ������ѭ������10�Σ� ÿ������10%�������һ������ʣ���ȫ��
      for iper := 1 to 10 do
      begin
        // ÿ��SKU�Ϻ�ѭ���� slFGNumber ��Ϊ�Ѱ����ȼ�˳������
        for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
        begin 
          // ��ȡҪ���ƻ�� ��SKU�Ϻţ� ���ڵ�����
          aSAPS618Reader.GetDemands(aFGPriorityReader.FList[inumber], aSAPS618ColPtr^.dt1,
            dtpDemandBeginDate4.DateTime, lstDemand);
 
          for iDemand := 0 to lstDemand.Count - 1 do
          begin
            aSAPS618ColPtr_demand := PSAPS618Col(lstDemand[iDemand]);

            aSBom := TSBom(aFGPriorityReader.FList.Objects[inumber]);
            if aSBom = nil then Continue;  // ���û��Bom�� ����

            dDemandQty := aSAPS618ColPtr_demand^.dDemandQty;
            // ���һ�Σ� ������������
            if iper = 10 then
            begin
              //  ��ȥ���������� ���ϴ�δ��������������ǽ�ȥ
              dqty := dDemandQty - aSAPS618ColPtr_demand^.dQty_ok;
              aSAPS618ColPtr_demand^.dQty_calc := dDemandQty;
            end
            else  // �������� 10%
            begin
              dqty := Round(dDemandQty * 0.1);
              if aSAPS618ColPtr_demand^.dQty_calc + dqty > dDemandQty then
              begin
                //dqty := dDemandQty - aSOPCol_demand.iQty_calc;    
                aSAPS618ColPtr_demand^.dQty_calc := dDemandQty;
              end
              else
              begin
                aSAPS618ColPtr_demand^.dQty_calc := aSAPS618ColPtr_demand^.dQty_calc + dqty;
              end;
              //  ��ȥ���������� ���ϴ�δ��������������ǽ�ȥ
              dqty := aSAPS618ColPtr_demand^.dQty_calc - aSAPS618ColPtr_demand^.dQty_ok;
            end;

            // dqty ���μ��� ��������
            dqty_child_min := -9999;
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
                                                                          
              // ȡ�ɹ�Ӧ����
              dqty_a := aSBomChild.GetQtyAvail(dtMemand - aSBomChild.FLT);

              if DoubleE( dqty_a , 0 ) then
              begin
                dqty_child_min := 0;
                Break;  // ���������Ϊ0�� ���������
              end;                                               
            
              if DoubleE(dqty_child_min , -9999) or DoubleG(dqty_child_min , Trunc( dqty_a / aSBomChild.dUsage) ) then
              begin
                dqty_child_min := Trunc(dqty_a / aSBomChild.dUsage);
              end;
            end;

            if DoubleLE(dqty_child_min , 0) then
            begin 
              Continue; // �������������������Ϊ0������������һSKU
            end;
                                     
            if DoubleG( dqty , dqty_child_min) then // ����������������
            begin
              dqty := dqty_child_min;
            end;

            // ��������ٿ��������� ����
            aSAPS618ColPtr_demand^.dQty_ok := aSAPS618ColPtr_demand^.dQty_ok + dqty; // ����������
            // ���ӹ�Ӧ���ѷ�����
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
              aSBomChild.AllocQty(aFGPriorityReader.FList[inumber], dtMemand - aSBomChild.FLT, dqty * aSBomChild.dUsage); 
            end;
          end;
        end;
      end;


      // Ƿ�Ϸ��䣬���㲻����ƻ���ȱ���ϣ� slFGNumber ��Ϊ�Ѱ����ȼ�˳������
      for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
      begin
        dtMemand := aSAPS618ColPtr^.dt2 - 2;
        if dtMemand < aSAPS618ColPtr^.dt1 then
          dtMemand := aSAPS618ColPtr^.dt1;
                     
          if (aFGPriorityReader.FList[inumber] = '93.56.3564201') then
          begin
            Sleep(1);
          end;  

        aSAPS618Reader.GetDemands(aFGPriorityReader.FList[inumber], aSAPS618ColPtr^.dt1,
          dtpDemandBeginDate4.DateTime, lstDemand);
        for iDemand := 0 to lstDemand.Count - 1 do
        begin
          aSAPS618ColPtr_demand := PSAPS618Col(lstDemand[iDemand]);

          aSBom := TSBom(aFGPriorityReader.FList.Objects[inumber]);
          if aSBom = nil then
          begin
            AddShortageICItem(aSAPS618ColPtr_demand, aFGPriorityReader.FList[inumber] + ' ��BOM');
            Continue;  // ���û��Bom�� ����
          end;

          // һ�μ����꣬������    // ������ iQty�� ǰ�������Ϊ����ģ�����Ϊ����Ƿ��
          dqty := aSAPS618ColPtr_demand^.dQty - aSAPS618ColPtr_demand^.dQty_ok;
          if DoubleLE(dqty , 0) then Continue; // �����ˣ���Ƿ��

          // dqty ���μ��� ��������
          for ichild := 0 to aSBom.FList.Count - 1 do
          begin
            aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);

            dqty_a := aSBomChild.GetQtyAvail2(dtMemand - aSBomChild.FLT);

            if Trunc(dqty_a / aSBomChild.dUsage) < dqty then  // ��������¼��Ƿ����
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty_a);
              smsg := aSBomChild.FList[0] + '(' + aSBomChild.FGroup + '): ' + Format('%.0f', [dqty * aSBomChild.dUsage - dqty_a]);
              AddShortageICItem(aSAPS618ColPtr_demand, smsg);
            end
            else   //  ��������
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty * aSBomChild.dUsage);
            end; 
          end;
        end;
      end; 
    end;

    SaveSAP(sfile_save, aSAPS618Reader);

    
    MessageBox(Handle, '���', '��ʾ', 0);

  finally
    aFGPriorityReader.Free;
          
    aSAPS618Reader.Free;
    
    slProjYear.Free;

    for idate := 0 to sldate.Count - 1 do
    begin
      aSAPS618ColPtr := PSAPS618Col(sldate.Objects[idate]);
      Dispose(aSAPS618ColPtr);
    end;
    sldate.Free;

                      
    aKeyICItemSupplyReader.Free;
    aSBomReader.Free;

    lstDemand.Free;

    aSAPStockReader.Free;
  end;
end;

procedure TfrmMRPSimulation.btnPriorityClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  lePriority.Text := sfile;
end;

procedure TfrmMRPSimulation.btnDemand2Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leDemand_p4.Text := sfile;
  leDemand5.Text := sfile;

end;

procedure TfrmMRPSimulation.btnNewSKUClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leNewSKU_p4.Text := sfile;
end;

procedure TfrmMRPSimulation.btnDOSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leDOS_p4.Text := sfile;
end;

procedure TfrmMRPSimulation.btnCalcPriorityClick(Sender: TObject);
var 
  aNewSKUReader: TNewSKUReader;
  aDOSReader: TDOSReader;
  aSOPReader: TSOPReader;

  slProjYear: TStringList;

  slFGNumber: TStringList;
  inumber: Integer;
  p: POrderByConditions;
  idx: Integer;
  aDOSRecordPptr: PDOSRecord;

  sfile_save: string;     
  ExcelApp, WorkBook: Variant;
  irow: Integer;
begin
  sfile_save := 'Prio ' + FormatDateTime('yyyyMMdd-hhmmss', Now); // 20170705-093325
  if not ExcelSaveDialog(sfile_save) then Exit;
 
  dtpDemandBeginDate1.DateTime := EncodeDateTime(YearOf(dtpDemandBeginDate1.DateTime),
    MonthOf(dtpDemandBeginDate1.DateTime), DayOf(dtpDemandBeginDate1.DateTime), 0, 0, 0, 0);

  lePriority.Text := sfile_save;

  aNewSKUReader := TNewSKUReader.Create(leNewSKU_p4.Text);

  aDOSReader := TDOSReader.Create(leDOS_p4.Text);


  slProjYear := TfrmProjYear.GetProjYears;

  aSOPReader := TSOPReader.Create(slProjYear, leDemand_p4.Text);


  slFGNumber := TStringList.Create;

  aSOPReader.GetNumberList(slFGNumber);


  for inumber := 0 to slFGNumber.Count - 1 do
  begin
    p := New(POrderByConditions);
    slFGNumber.Objects[inumber] := TObject(p);

    p^.isnew := aNewSKUReader.FList.IndexOf(slFGNumber[inumber]) >= 0;
    idx := aDOSReader.FList.IndexOf(slFGNumber[inumber]);
    if idx >= 0 then
    begin
      aDOSRecordPptr := PDOSRecord(aDOSReader.FList.Objects[idx]);
      p^.dos := aDOSRecordPptr^.dos;
    end
    else
    begin
      p^.dos := 0;
    end;

    p^.demand := aSOPReader.GetDemandSum(dtpDemandBeginDate1.DateTime, slFGNumber[inumber]);
  end;


  slFGNumber.CustomSort(StringListSortCompare);



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
  ExcelApp.DisplayAlerts := False;
//  while ExcelApp.Sheets.Count > 1 do
//  begin
//    ExcelApp.Sheets.Delete(1);
//  end;

  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := 'SKU����˳��';

  irow := 1;
  ExcelApp.Cells[irow, 1].Value := '��Ʒ����';
  ExcelApp.Cells[irow, 2].Value := '���ȼ�';

  irow := 2;
  for inumber := 0 to slFGNumber.Count - 1 do
  begin
    ExcelApp.Cells[irow, 1].Value := slFGNumber[inumber];
    ExcelApp.Cells[irow, 2].Value := inumber + 1;
    irow := irow + 1;
  end;


  try

    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;



  for inumber := 0 to slFGNumber.Count - 1 do
  begin
    p := POrderByConditions(slFGNumber.Objects[inumber]);
    Dispose(p);
  end;
  slFGNumber.Free;



  aSOPReader.Free;
  slProjYear.Free;

  aDOSReader.Free;

  aNewSKUReader.Free;

  MessageBox(Handle, '���', '��ʾ', 0);

end;

procedure TfrmMRPSimulation.tbProjYearClick(Sender: TObject);
begin
  TfrmProjYear.ShowForm;
end;

procedure TfrmMRPSimulation.Button1Click(Sender: TObject);
var
  sfile: string;
  aSAPBomReader: TSAPBomReader2;
  aSAPStockReader: TSAPStockReader;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  aSAPStockReader := TSAPStockReader.Create(leSAPStock.Text);
  aSAPBomReader := TSAPBomReader2.Create(leSAPBom.Text);
  try
    aSAPBomReader.SaveSBom(sfile, aSAPStockReader);
  finally
    aSAPBomReader.Free;
    aSAPStockReader.Free;
  end;
  
  MessageBox(Self.Handle, '���', '��ʾ', 0);
end;

procedure TfrmMRPSimulation.btnCMS2MRPSimClick(Sender: TObject);
var
  aTPCMSConfirmReader: TTPCMSConfirmReader;
  aSAPStockReader: TSAPStockReader;
  aSBomReader: TSBomReader;
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  aTPCMSConfirmReader := TTPCMSConfirmReader.Create(leCMSConfirm.Text);
  aSAPStockReader := TSAPStockReader.Create(leSAPStock.Text);
  aSBomReader := TSBomReader.Create(leSBOM1.Text);

  try 
    aTPCMSConfirmReader.Save(aSAPStockReader, aSBomReader, sfile);
  finally                
    aTPCMSConfirmReader.Free;
    aSAPStockReader.Free;
    aSBomReader.Free;
  end;      
  MessageBox(Handle, '���', '��ʾ', 0);
end;

procedure TfrmMRPSimulation.btnCMSConfirmClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leCMSConfirm.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSAPBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSAPStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSBOM1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSBOM1.Text := sfile;
end;

procedure TfrmMRPSimulation.btnStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leStock.Text := sfile;
end;

  // �򻯵�MRP���㣬�����ǵ�λ��
  type
    TMRPUnit = packed record
      snumber: string;
      sname: string;
      dt: TDateTime;
      dQty: Double;
      dQtyStock: Double;
      dQtyStockParent: Double;
      bExpend: Boolean;
      aBom: TSapBom;
      aParentBom: TSapBom;
    end;
    PMRPUnit = ^TMRPUnit;
    function StringListSortCompare_DateTime (List: TStringList; Index1, Index2: Integer): Integer;
    var
      p1, p2: PMRPUnit;
    begin
      p1 := PMRPUnit(List.Objects[Index1]);
      p2 := PMRPUnit(List.Objects[Index2]);
      if DoubleG(p1^.dt, p2^.dt) then
        Result := 1
      else if DoubleL(p1^.dt, p2^.dt) then
        Result := -1
      else Result := 0;
    end;

function GetDemand(lstDemand: TStringList; const snumber: string; dt1, dt2: TDateTime): Double;
var
  i: Integer;
  aMRPUnitPtr: PMRPUnit;
begin
  Result := 0;
  for i := 0 to lstDemand.Count - 1 do
  begin
    if lstDemand[i] <> snumber then Continue;
    aMRPUnitPtr := PMRPUnit(lstDemand.Objects[i]);
    if (aMRPUnitPtr^.dt >= dt1) and (aMRPUnitPtr^.dt < dt2) then
    begin
      Result := Result + aMRPUnitPtr^.dQty;
    end;
  end;
end;

procedure TfrmMRPSimulation.btnMRPClick(Sender: TObject);
var                                                                                                       
  sfile: string;
  aSAPBomReader: TSAPBomReader2;
  aSAPStockReader: TSAPStockReader;
  aSAPMaterialReader2: TSAPMaterialReader2;
  aSAPStockSum: TSAPStockSum;

  aSAPS618Reader: TSAPS618Reader;
  lstDemand: TStringList;  
  lstDemand_tmp: TStringList;
  
  aSAPS618: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
  iLine: Integer;
  iDate: Integer; 
  aMRPUnitPtr: PMRPUnit;
  aMRPUnitPtr_Dep: PMRPUnit;
  iMrpUnit: Integer;

  iChild: Integer;
  iChildItem: Integer;
  aSapItemGroup: TSapItemGroup;

  aSapBomChild: TSapBom;
  
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  bLoop: BOOL;
  dQty: Double;
  dQty_Stock_a: Double;
  lstNumber: TStringList;
  iNumber: Integer;
  dtMin, dtMax: TDateTime;
  dt: TDateTime;
  icol: Integer;
  icolMax: Integer;
  idx: Integer;
  dtToday: TDateTime;
  dtNextMonday: TDateTime;
  iSheet: Integer;

  ExcelApp2, WorkBook2: Variant;    
  iSheetCount2, iSheet2: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7, stitle8, stitle9: string;
  stitle: string;

  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  sl: TStringList;
  sline: string;
  sname_p, snumber_p: string;
  sabc: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;

   

  Memo3.Lines.Add('��ʼ��ȡ ���');
  aSAPStockReader := TSAPStockReader.Create(leSAPStock3.Text, OnLogEvent);

  if not aSAPStockReader.ReadOk then
  begin
    aSAPStockReader.Free;
    MessageBox(Handle, '��ȡ���ʧ��', '����', 0);
    Exit;
  end;

  Memo3.Lines.Add('��ʼ��ȡ BOM');
  aSAPBomReader := TSAPBomReader2.Create(leSAPBom3.Text, OnLogEvent);
  
  Memo3.Lines.Add('��ʼ��ȡ Ҫ���ƻ�');
  aSAPS618Reader := TSAPS618Reader.Create(leDemand3.Text, '�������Ҫ���ƻ�', OnLogEvent);
                      
  Memo3.Lines.Add('��ʼ��ȡ ������Ϣ');
  aSAPMaterialReader2 := TSAPMaterialReader2.Create(leICItem.Text, OnLogEvent);
  


  aSAPStockSum := TSAPStockSum.Create;
  aSAPStockReader.SumTo(aSAPStockSum);   
  aSAPStockReader.Free;
  
  lstDemand := TStringList.Create;
  lstNumber := TStringList.Create;
                          
  Memo3.Lines.Add('����Ҫ���ƻ�����');
  for iLine := 0 to aSAPS618Reader.Count - 1 do
  begin
    aSAPS618 := aSAPS618Reader.Items[iLine];
    for iDate := 0 to aSAPS618.Count - 1 do
    begin
      aSAPS618ColPtr := aSAPS618.Items[iDate];
      
      // ����Ϊ0�Ĳ������
      if DoubleE( aSAPS618ColPtr^.dqty, 0 ) then Continue;
      //�������ڿ�ʼ���ڵĲ����㣬����
      if aSAPS618ColPtr^.dt1 < dtpDemandBeginDate3.DateTime then Continue;

      aMRPUnitPtr := New(PMRPUnit);   
      aMRPUnitPtr^.snumber := aSAPS618ColPtr^.snumber;
      aMRPUnitPtr^.sname := aSAPS618ColPtr^.sname;
      aMRPUnitPtr^.dt := aSAPS618ColPtr^.dt1;
      aMRPUnitPtr^.dQty := aSAPS618ColPtr^.dqty;
      aMRPUnitPtr^.dQtyStock := 0;
      aMRPUnitPtr^.dQtyStockParent := 0;
      aMRPUnitPtr^.bExpend := False;
      aMRPUnitPtr^.aBom := nil;     
      aMRPUnitPtr^.aParentBom := nil;
      lstDemand.AddObject(aSAPS618ColPtr^.snumber, TObject(aMRPUnitPtr));
    
    end;
  end;

  aSAPS618Reader.Free;
                            
  Memo3.Lines.Add('��ʼģ��MRP����');
  try
    bLoop := True;
    while bLoop do
    begin
      bLoop := False;
      
      //��������
      lstDemand_tmp := TStringList.Create;
      for iMrpUnit := 0 to lstDemand.Count - 1 do
      begin
        lstDemand_tmp.AddObject(lstDemand[iMrpUnit], lstDemand.Objects[iMrpUnit]);
      end;
      lstDemand.Clear;

      //���򣬰�����
      lstDemand_tmp.CustomSort(StringListSortCompare_DateTime);
 
      for iMrpUnit := 0 to lstDemand_tmp.Count - 1 do
      begin
        aMRPUnitPtr := PMRPUnit(lstDemand_tmp.Objects[iMrpUnit]);
        if not aMRPUnitPtr^.bExpend then
        begin
          // ���ڵ�  �����ǿ�� ////////////////////////////////////////////////////
          if aMRPUnitPtr^.aParentBom = nil then
          begin   
            aMRPUnitPtr^.bExpend := True;      
            aMRPUnitPtr^.aBom := aSAPBomReader.GetSapBom(lstDemand_tmp[iMrpUnit], '');
            if aMRPUnitPtr^.aBom = nil then  // û��BOM���쳣����¼��־
            begin
              Memo3.Lines.Add(lstDemand_tmp[iMrpUnit] + ' û��BOM');
              Continue;
            end;
            for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
            begin
              aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
              // BOM �ϲ����������  �� ���ȥ�ϲ��ѷ�����
//              dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapItemGroup.Items[0].dusage;     
              dQty := aMRPUnitPtr^.dQty * aSapItemGroup.Items[0].dusage;      // ÿһ������󣬶��ǿۼ����ѷ������
              // ������ÿ��
              if dQty > 0 then
              begin
                for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                begin
                  aSapBomChild := aSapItemGroup.Items[iChildItem];
                  dQty_Stock_a := aSAPStockSum.GetAvailStock(aSapBomChild.FNumber);

                  if DoubleE(dQty_Stock_a, 0) then Continue;
                  if dQty <= dQty_Stock_a then
                  begin
                    aSapBomChild.FStock := dQty;
                    aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty);  
                    dQty := 0;
                    Break;
                  end
                  else
                  begin
                    aSapBomChild.FStock := dQty_Stock_a;           
                    aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty_Stock_a);
                    dQty := dQty - dQty_Stock_a;
                  end;
                end;
              end;
              
              //չ�������²�
              for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
              begin                        
                aSapBomChild := aSapItemGroup.Items[iChildItem];
                aMRPUnitPtr_Dep := New(PMRPUnit);
                aMRPUnitPtr_Dep^.snumber := aSapBomChild.FNumber;
                aMRPUnitPtr_Dep^.sname := aSapBomChild.FName;
                aMRPUnitPtr_Dep^.dt := aMRPUnitPtr^.dt - aMRPUnitPtr^.aBom.lt;
                if aSapBomChild.sgroup = '' then
                begin
                  aMRPUnitPtr_Dep^.dQty := dQty;   
                end
                else
                begin
                  aMRPUnitPtr_Dep^.dQty := dQty * aSapBomChild.dPer / 100;  // ���Ʒ����ϰ���ȷ�
                end;
                
                aMRPUnitPtr_Dep^.dQtyStock := aSapBomChild.FStock;   
                aSapBomChild.FStock := 0;  // ��ֵ����0��������������ֵ��Ӱ����һ�ּ���
                aMRPUnitPtr_Dep^.dQtyStockParent := aMRPUnitPtr^.dQtyStock + aMRPUnitPtr^.dQtyStockParent;  // ������
                aMRPUnitPtr_Dep^.bExpend := False;
                aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                lstDemand.AddObject(aSapBomChild.FNumber, TObject(aMRPUnitPtr_Dep));
              end;
              bLoop := True;
            end;
          end
          else
          // �Ǹ��ڵ�  �����ǿ�� //////////////////////////////////////////////////      
          begin
            // Ҷ�ӽڵ㣬 ֱ����ӵ�����////////////////////////////////////////////
            if aMRPUnitPtr^.aBom.ChildCount = 0 then
            begin
              aMRPUnitPtr_Dep := New(PMRPUnit);
              aMRPUnitPtr_Dep^ := aMRPUnitPtr^;
              lstDemand.AddObject(lstDemand_tmp[iMrpUnit], TObject(aMRPUnitPtr_Dep));
            end        
            // ��Ҷ�ӽڵ㣬 �Ǹ��ڵ㣬���ǰ��Ʒ��չ������//////////////////////////
            else    // �����迼�ǰ��Ʒ��� ////////////////////////////////////////
            begin
              aMRPUnitPtr^.bExpend := True;
 
              for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
              begin
                aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];

                // BOM �ϲ����������  �� ���ȥ�ϲ��ѷ�����
                //dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapItemGroup.Items[0].dusage;
                dQty := aMRPUnitPtr^.dQty* aSapItemGroup.Items[0].dusage;  // ÿһ������󣬶��ǿۼ����ѷ������

                // ������ÿ��
                if dQty > 0 then
                begin
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem]; 
 
                    dQty_Stock_a := aSAPStockSum.GetAvailStock(aSapBomChild.FNumber);
                    if DoubleE(dQty_Stock_a, 0) then Continue;
                  
                    if dQty <= dQty_Stock_a then
                    begin
                      aSapBomChild.FStock := dQty;
                      aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty);
                      dQty := 0;
                      Break;
                    end
                    else
                    begin
                      aSapBomChild.FStock := dQty_Stock_a;           
                      aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty_Stock_a);
                      dQty := dQty - dQty_Stock_a;
                    end;
                  end;
                end;

                for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                begin
                  aSapBomChild := aSapItemGroup.Items[iChildItem];
                  aMRPUnitPtr_Dep := New(PMRPUnit);
                  aMRPUnitPtr_Dep^.snumber := aSapBomChild.FNumber;
                  aMRPUnitPtr_Dep^.sname := aSapBomChild.FName;
                  aMRPUnitPtr_Dep^.dt := aMRPUnitPtr^.dt - aMRPUnitPtr^.aBom.lt;
                  if aSapBomChild.sgroup = '' then
                  begin
                    aMRPUnitPtr_Dep^.dQty := dQty;  // ���Ʒ����ϰ���ȷ�
                  end
                  else
                  begin
                    aMRPUnitPtr_Dep^.dQty := dQty * aSapBomChild.dPer / 100;  // ���Ʒ����ϰ���ȷ�
                  end;
                  aMRPUnitPtr_Dep^.dQtyStock := aSapBomChild.FStock;
                  aSapBomChild.FStock := 0;  // ��ֵ����0��������������ֵ��Ӱ����һ�ּ���

                  aMRPUnitPtr_Dep^.dQtyStockParent := aMRPUnitPtr^.dQtyStock + aMRPUnitPtr^.dQtyStockParent;
                  aMRPUnitPtr_Dep^.bExpend := False;
                  aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                  aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                  lstDemand.AddObject(aSapBomChild.FNumber, TObject(aMRPUnitPtr_Dep));
                end;
                bLoop := True;
              end;          
            end;
          end;
        end;
//        Memo2.Lines.Add(aMRPUnitPtr^.snumber + #9 + aMRPUnitPtr^.sname  + #9 + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt) + #9 + FloatToStr(aMRPUnitPtr^.dQty) + #9 + FloatToStr(aMRPUnitPtr^.dQtyStock) + #9 + FloatToStr(aMRPUnitPtr^.dQtyStockParent));
        Dispose(aMRPUnitPtr);
      end;
      lstDemand_tmp.Free;
    end;
    
   
               
    //���򣬰�����
    lstDemand.CustomSort(StringListSortCompare_DateTime);

    dtMin := myStrToDateTime('2100-12-31');
    dtMax := myStrToDateTime('1900-01-01');
    if lstDemand.Count > 0 then
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand.Objects[0]);
      dtMin := aMRPUnitPtr^.dt;
      aMRPUnitPtr := PMRPUnit(lstDemand.Objects[lstDemand.Count - 1]);
      dtMax := aMRPUnitPtr^.dt;
    end;

    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand.Objects[iMrpUnit]);
//      Memo2.Lines.Add(aMRPUnitPtr^.snumber + #9 + aMRPUnitPtr^.sname  + #9 + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt) + #9 + FloatToStr(aMRPUnitPtr^.dQty) + #9 + FloatToStr(aMRPUnitPtr^.dQtyStock) + #9 + FloatToStr(aMRPUnitPtr^.dQtyStockParent));
      idx := lstNumber.IndexOf(lstDemand[iMrpUnit]);
      if idx < 0 then
      begin
        aMRPUnitPtr_Dep := New(PMRPUnit);
        aMRPUnitPtr_Dep^ := aMRPUnitPtr^;
        lstNumber.AddObject(lstDemand[iMrpUnit], TObject(aMRPUnitPtr_Dep));
      end
      else
      begin
        aMRPUnitPtr_Dep := PMRPUnit(lstNumber.Objects[idx]);
        aMRPUnitPtr_Dep^.dQtyStockParent := aMRPUnitPtr_Dep^.dQtyStockParent + aMRPUnitPtr^.dQtyStockParent;
      end;  
    end;


    // ���� //////////////////////////////////////////////////////////////////////

    Memo3.Lines.Add('��ʼ����ģ��MRP��������ֻ����ؼ�����');
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
    ExcelApp.DisplayAlerts := False;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'ģ��MRP';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 2].Value := '����';
    ExcelApp.Cells[irow, 3].Value := '����';    
    ExcelApp.Cells[irow, 4].Value := '���';

    ExcelApp.Columns[1].ColumnWidth := 17;
    ExcelApp.Columns[2].ColumnWidth := 13;
    ExcelApp.Columns[3].ColumnWidth := 9;
    ExcelApp.Columns[4].ColumnWidth := 9;

    ProgressBar1.Max := lstDemand.Count;
    ProgressBar1.Position := 0;

    sl := TStringList.Create;
    sline := '���ϱ���'#9'��������'#9'����'#9'��������'#9'���'#9'������'#9'ABC��־'#9'��������'#9'��������';
    sl.Add(sline);

    irow := 2;
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand.Objects[iMrpUnit]);

      if aMRPUnitPtr^.aParentBom <> nil then
      begin
        snumber_p := aMRPUnitPtr^.aParentBom.FNumber;
        sname_p := aMRPUnitPtr^.aParentBom.FName;
      end
      else
      begin
        snumber_p := '';
        sname_p := '';
      end;
      if aMRPUnitPtr^.aBom <> nil then
      begin
        sabc := aMRPUnitPtr^.aBom.abc;
      end
      else
      begin
        sabc := '';
      end;  
      sline := aMRPUnitPtr^.snumber + #9 +
        aMRPUnitPtr^.sname + #9 +
        FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt) + #9 +
        Format('%0.0f'#9'%0.0f'#9'%0.0f', [aMRPUnitPtr^.dQty, aMRPUnitPtr^.dQtyStock, aMRPUnitPtr^.dQtyStockParent]) + #9 +  // 3��
        sabc + #9 +
        snumber_p + #9 +
        sname_p;
      sl.Add(sline);

      // ֻд�ؼ�����
      if aMRPUnitPtr^.aBom.abc = '' then Continue;

      ExcelApp.Cells[irow, 1].Value := lstDemand[iMrpUnit];
      ExcelApp.Cells[irow, 2].Value := aMRPUnitPtr^.dt;
      ExcelApp.Cells[irow, 3].Value := aMRPUnitPtr^.dQty;
      ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr^.dQtyStock;
      irow := irow + 1;
      ProgressBar1.Position := ProgressBar1.Position + 1;
    end;

    ////////////////////////////////////////////////////////////////////////////

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'ģ��MRP���ڻظ�';
                  
    ExcelApp.Cells[1, 1].Value := 'MRPģ��--Ҫ���ƻ�������������ƻ�';
    ExcelApp.Cells[2, 1].Value := 'ģ������:';                   
    ExcelApp.Cells[2, 2].Value := FormatDateTime('yyyy-MM-dd HH:mm', Now);
    ExcelApp.Cells[3, 1].Value := '�������:';

    
    irow := 4;
    ExcelApp.Cells[irow, 1].Value := '���ϱ���';
    ExcelApp.Cells[irow, 2].Value := '��������';
    ExcelApp.Cells[irow, 3].Value := '���������';
    ExcelApp.Cells[irow, 4].Value := '������';
    ExcelApp.Cells[irow, 5].Value := 'Commitment';

    ExcelApp.Columns[1].ColumnWidth := 14;
    ExcelApp.Columns[2].ColumnWidth := 20;
    ExcelApp.Columns[3].ColumnWidth := 8;
    ExcelApp.Columns[4].ColumnWidth := 8;
    ExcelApp.Columns[5].ColumnWidth := 13;

    icolMax := 1;

    dtToday := myStrToDateTime(FormatDateTime('yyyy-MM-dd', Now));
    
    ProgressBar1.Max := lstNumber.Count;
    ProgressBar1.Position := 0;
    irow := 5;
    for iNumber := 0 to lstNumber.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstNumber.Objects[iNumber]);
      // ֻд�ؼ�����
      if aMRPUnitPtr^.aBom.abc = '' then
      begin
        ProgressBar1.Position := ProgressBar1.Position + 1;
        Continue;
      end;
            
      aSAPMaterialRecordPtr := aSAPMaterialReader2.GetSAPMaterialRecord(aMRPUnitPtr^.snumber);
      
      ExcelApp.Cells[irow, 1].Value := lstNumber[iNumber];
      ExcelApp.Cells[irow, 2].Value := aMRPUnitPtr^.sname;
      ExcelApp.Cells[irow, 3].Value := aMRPUnitPtr^.dQtyStockParent;  // ������
      ExcelApp.Cells[irow, 4].Value := aSAPStockSum.GetStock(lstNumber[iNumber]);
      ExcelApp.Cells[irow, 5].Value := 'Ҫ���ƻ�������';
      ExcelApp.Cells[irow + 1, 5].Value := '�ɹ�Ӧ��';
      ExcelApp.Cells[irow + 2, 5].Value := '��Ӧ����';



      icol := 6;
      dt := dtMin;
      while dt < dtMax do
      begin
        if irow = 5 then
        begin                   
          if icol = 6 then
          begin
            ExcelApp.Cells[3, icol].Value := 'WK' + IntToStr( WeekOfTheYear(dtToday) );
            ExcelApp.Cells[4, icol].Value := FormatDateTime('yyyy-MM-dd', dtToday);
          end
          else
          begin
            ExcelApp.Cells[3, icol].Value := 'WK' + IntToStr( WeekOfTheYear(dt) );
            ExcelApp.Cells[4, icol].Value := FormatDateTime('yyyy-MM-dd', dt);
          end;
          if icolMax < icol then
          begin
            icolMax := icol;
          end;               
          ExcelApp.Columns[icol].ColumnWidth := 10;
        end;
        if icol = 6 then
        begin
          dtNextMonday := dtToday - DayOfWeek(dtToday) + 2 + 7;
          ExcelApp.Cells[irow, icol].Value := GetDemand(lstDemand, lstNumber[iNumber], dt, dtNextMonday);  // ����һ֮ǰ
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);

          dt := dtNextMonday;   // ���ڶ�λ������1
        end
        else
        begin
          ExcelApp.Cells[irow, icol].Value := GetDemand(lstDemand, lstNumber[iNumber], dt, dt + 7);
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol - 1) + IntToStr(irow + 2) + '+' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
               
          dt := dt + 7;
        end;
        icol := icol + 1;
      end;
      
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;
      
//      AddColor(ExcelApp, irow + 2, 10, irow + 2, icolMax, $E4DCD6);
//      AddColor(ExcelApp, irow + 3, 10, irow + 3, icolMax, $D6E4FC);

      MergeCells(ExcelApp, irow, 1, irow + 2, 1);
      MergeCells(ExcelApp, irow, 2, irow + 2, 2);
      MergeCells(ExcelApp, irow, 3, irow + 2, 3);
      MergeCells(ExcelApp, irow, 4, irow + 2, 4);

      if aSAPMaterialRecordPtr <> nil then
      begin
        ExcelApp.Cells[irow, icolMax + 1].Value := aSAPMaterialRecordPtr^.sMRPer;
        ExcelApp.Cells[irow, icolMax + 2].Value := aSAPMaterialRecordPtr^.sBuyer;
        
        ExcelApp.Cells[irow + 1, icolMax + 1].Value := aSAPMaterialRecordPtr^.sMRPer;
        ExcelApp.Cells[irow + 1, icolMax + 2].Value := aSAPMaterialRecordPtr^.sBuyer;

        ExcelApp.Cells[irow + 2, icolMax + 1].Value := aSAPMaterialRecordPtr^.sMRPer;
        ExcelApp.Cells[irow + 2, icolMax + 2].Value := aSAPMaterialRecordPtr^.sBuyer;
      end;

      irow := irow + 3;
      ProgressBar1.Position := ProgressBar1.Position + 1;
    end;
                                                
    ExcelApp.Cells[4, icolMax + 1].Value := 'MC';
    ExcelApp.Cells[4, icolMax + 2].Value := '�ɹ�';
    
    AddBorder(ExcelApp, 1, 1, irow - 1, icolMax + 2);
    AddColor(ExcelApp, 3, 6, 3, icolMax + 2, $E8DEB7);
    AddColor(ExcelApp, 4, 1, 4, icolMax + 2, $E8DEB7);

    ExcelApp.Cells[irow + 2, 1].Value := '˵����';
    ExcelApp.Cells[irow + 3, 1].Value := 'Ҫ���ƻ���������ָ���������������Ʒ����100����Ʒ��治���ǣ����ṩ���Ǿ�Ҫ���ƻ��Ѿ����ǹ��ˣ����м���Ʒ���20��ԭ�ϱ�����10����Ҫ���ƻ�������=100-20-10=70��';
    ExcelApp.Cells[irow + 4, 1].Value := '�ɹ�Ӧ�����ɹ��ظ�����������Ժ�Ӧ�̿��Խ���֮�������磺�������������=100��Ҫ���ƻ�������Ϊ100����ʱ���ϲ�����ʵΪ200������Ӧ��ȷ�Ͽɽ���80����ɹ�Ӧ��=80��';
    ExcelApp.Cells[irow + 5, 1].Value := '��Ӧ���죺���������=ǰ��ʣ����+����ɹ�Ӧ��-����Ҫ���ƻ�������';

                  
    aSAPMaterialReader2.Free;
                                 
    aSAPStockSum.Free;

    aSAPBomReader.Free;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '�������Ҫ���ƻ�';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';
    try

      WorkBook2 := ExcelApp2.WorkBooks.Open(leDemand3.Text);


      try
        iSheetCount2 := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount2 do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet := ExcelApp2.Sheets[iSheet2].Name;
          Memo1.Lines.Add(sSheet);

          irow := 1;
          stitle1 := ExcelApp2.Cells[irow, 1].Value;
          stitle2 := ExcelApp2.Cells[irow, 2].Value;
          stitle3 := ExcelApp2.Cells[irow, 3].Value;
          stitle4 := ExcelApp2.Cells[irow, 4].Value;
          stitle5 := ExcelApp2.Cells[irow, 5].Value;
          stitle6 := ExcelApp2.Cells[irow, 6].Value;
          stitle7 := ExcelApp2.Cells[irow, 7].Value;
          stitle8 := ExcelApp2.Cells[irow, 8].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8;

          if stitle <> '���Ϲ���MRP ��Χ���������ƻ������ַ�����������λ�ϼ�' then
          begin        
            Memo1.Lines.Add(sSheet +'  ����  �������Ҫ���ƻ�  ��ʽ');
            Continue;
          end;

          try
            ExcelApp2.ActiveSheet.Cells.Copy;

            ExcelApp.Sheets[iSheet].Paste;
          except
            on e: Exception do
            begin
              Memo1.Lines.Add(e.Message);
            end;

          end;
//          ExcelApp.ActiveSheet.Range('A1').Select;

          ExcelApp2.CutCopyMode := False;

          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;
         
          
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '���ÿ��';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';
    try

      WorkBook2 := ExcelApp2.WorkBooks.Open(leSAPStock3.Text);


      try
        iSheetCount2 := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount2 do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet := ExcelApp2.Sheets[iSheet2].Name;
          Memo1.Lines.Add(sSheet);

          irow := 1;
          stitle1 := ExcelApp2.Cells[irow, 1].Value;
          stitle2 := ExcelApp2.Cells[irow, 2].Value;
          stitle3 := ExcelApp2.Cells[irow, 3].Value;
          stitle4 := ExcelApp2.Cells[irow, 4].Value;
          stitle5 := ExcelApp2.Cells[irow, 5].Value;
          stitle6 := ExcelApp2.Cells[irow, 6].Value;                                   
          stitle7 := ExcelApp2.Cells[irow, 7].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7;

          if stitle <> '�������ص�ִ��ص������������������������ʹ�õĿ����ת��ת��' then
          begin        
            Memo1.Lines.Add(sSheet +'  ����  SAP�������  ��ʽ');
            Continue;
          end;

          try
            ExcelApp2.ActiveSheet.Cells.Copy;

            ExcelApp.Sheets[iSheet].Paste;
          except
            on e: Exception do
            begin
              Memo1.Lines.Add(e.Message);
            end;
          end;

//          ExcelApp.ActiveSheet.Range('A1').Select;

          ExcelApp2.CutCopyMode := False;

          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;
                 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'BOM';
          

    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';
    try

      WorkBook2 := ExcelApp2.WorkBooks.Open(leSAPBom3.Text);


      try
        iSheetCount2 := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount2 do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet := ExcelApp2.Sheets[iSheet2].Name;
          Memo1.Lines.Add(sSheet);

          irow := 1;
          stitle1 := ExcelApp2.Cells[irow, 1].Value;
          stitle2 := ExcelApp2.Cells[irow, 2].Value;
          stitle3 := ExcelApp2.Cells[irow, 3].Value;
          stitle4 := ExcelApp2.Cells[irow, 4].Value;                 
          stitle5 := ExcelApp2.Cells[irow, 5].Value;
          stitle6 := ExcelApp2.Cells[irow, 6].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;
                        
          if stitle <> 'ĸ�����ϱ���ĸ����������������;������ĸ��L/T' then
          begin
            Memo1.Lines.Add(sSheet +'  ����SAP����BOM��ʽ');
            Continue;
          end;

          try
            ExcelApp2.ActiveSheet.Cells.Copy;

            ExcelApp.Sheets[iSheet].Paste;
          except
            on e: Exception do
            begin
              Memo1.Lines.Add(e.Message);
            end;
          end;
//          ExcelApp.ActiveSheet.Range('A1').Select;

          ExcelApp2.CutCopyMode := False;


          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;
         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '���������־';


    try
      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;
    except
      on e: Exception do
      begin
        Memo1.Lines.Add(e.Message);
        sl.SaveToFile('c:\a.txt');
      end;
    end;

    sl.Free;

    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

  

    
  finally

    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand.Objects[iMrpUnit]);
      Dispose(aMRPUnitPtr);
    end;
    lstDemand.Free;

    for iMrpUnit := 0 to lstNumber.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstNumber.Objects[iMrpUnit]);
      Dispose(aMRPUnitPtr);
    end;
    lstNumber.Free;
  end;

  MessageBox(Handle, '���', '��ʾ', 0);
end;

procedure TfrmMRPSimulation.btnSAPStock3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock3.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSAPBom3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom3.Text := sfile;
end;

procedure TfrmMRPSimulation.btnDemand3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leDemand3.Text := sfile;
end;

procedure TfrmMRPSimulation.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

procedure TfrmMRPSimulation.btnICItemClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leICItem.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSim2Click(Sender: TObject);

var 
  aKeyICItemSupplyReader: TKeyICItemSupplyReader;

  aSBomReader: TSBomReader;

  aSOPReader: TSOPReader;
 
  slProjYear: TStringList;
  idate: Integer;
  sldate: TStringList;
  aSOPCol: TSOPCol;

  inumber: Integer;
  iper: Integer;

  aSOPCol_demand: TSOPCol;
  aSBom: TSBom;
  aSBomChild: TSBomChild;

  ibom: Integer;
  ichild: Integer;

  dqty: Double;
  dqty_child_min: Double;
  dqty_a: Double;

  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
 
  sfile_save: string;

  aSOPSimReader: TSOPSimReader;
  aMrpMPSReader: TMrpMPSReader;

  aFGPriorityReader: TFGPriorityReader;
  lstDemand: TList;
  iDemand: Integer;
  smsg: string;

  dtMemand: TDateTime;
  dDemandQty: Double;

  igroup: Integer;
  
  aSAPStockReader: TSAPStockReader;

 
begin
  sfile_save := 'SOP ' + FormatDateTime('yyyyMMdd-hhmmss', Now); // 20170705-093325
  if not ExcelSaveDialog(sfile_save) then Exit;

  dtpDemandBeginDate4.DateTime := EncodeDateTime(YearOf(dtpDemandBeginDate4.DateTime),
    MonthOf(dtpDemandBeginDate4.DateTime), DayOf(dtpDemandBeginDate4.DateTime), 0, 0, 0, 0);

  lstDemand := TList.Create;

  Memo1.Lines.Add('��ȡ��Ŀ��� ......');
  slProjYear := TfrmProjYear.GetProjYears;
  sldate := TStringList.Create;

  aMrpMPSReader := TMrpMPSReader.Create(leLastMPS.Text);
  Memo1.Lines.Add('��ȡ����MRP¼��MPS ......');

  Memo1.Lines.Add('��ȡ����S&OPģ�� ......');
  aSOPSimReader := TSOPSimReader.Create(leLastSOPAnalysis.Text, slProjYear);

  Memo1.Lines.Add('��ȡ�ؼ����Ϲ�Ӧ���� ......');
  aKeyICItemSupplyReader := TKeyICItemSupplyReader.Create(leSupplyEval.Text);

  Memo1.Lines.Add('��ȡ����BOM ......');
  aSBomReader := TSBomReader.Create(leSBOM.Text);
             
  Memo1.Lines.Add('��ȡҪ���ƻ� ......');
  aSOPReader := TSOPReader.Create(slProjYear, leDemand5.Text);
 

  Memo1.Lines.Add('��ȡSKU���ȼ� ......');
  aFGPriorityReader := TFGPriorityReader.Create(lePriority.Text);
                                          
  Memo1.Lines.Add('��ȡSAP��� ......');
  aSAPStockReader := TSAPStockReader.Create(leStock.Text);

  aKeyICItemSupplyReader.SAPStock := aSAPStockReader; // ��Щitem�п�棬����û�ó��ɹ����� �ظ���������û������ϡ����Ƿ�����ʱ��û������ϣ��������Ϊ�ɹ�ӦΪ0.���Ի�ȡ���Ϲ�Ӧ��ʱ��û�н��ڵģ�ȡ���
         
  try
                                        
    Memo1.Lines.Add('�� Bom ���������ϣ� ����Ӧ������ϵ�� ......');
    // �� Bom ���������ϣ� ����Ӧ������ϵ��
    for ibom := 0 to aSBomReader.FList.Count - 1 do
    begin
      aSBom := TSBom(aSBomReader.FList.Objects[ibom]);
      for ichild := 0 to aSBom.FList.Count - 1 do
      begin
        aSBomChild := TSBomChild(aSBom.FList.Objects[ichild]);
        for igroup := 0 to aSBomChild.FList.Count - 1 do
        begin
          PBomGroupChild(aSBomChild.FList.Objects[igroup])^.supp := aKeyICItemSupplyReader.GetSupplyLine(aSBomChild.FList[igroup]);
        end;
      end;
    end;
                        
    Memo1.Lines.Add('ƥ���Ʒ�Ϻ�Bom ......');
    // ƥ���Ʒ�Ϻ�Bom
    for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
    begin
      aFGPriorityReader.FList.Objects[inumber] := aSBomReader.GetBom(aFGPriorityReader.FList[inumber]);
      if aFGPriorityReader.FList.Objects[inumber] = nil then
      begin
        Memo1.Lines.Add(aFGPriorityReader.FList[inumber] + '  û��BOM');
      end;
    end;
    
    Memo1.Lines.Add('ȡ�����б� ......');
    // ȡ�����б�
    aSOPReader.GetDateList(sldate);
//    aSAPS618Reader.GetDateList(sldate);

    Memo1.Lines.Add('ÿһ�����ڼ��� ......');
    // ÿһ�����ڼ���
    for idate := 0 to sldate.Count - 1 do
    begin     
      aSOPCol := TSOPCol(sldate.Objects[idate]); // ����
      if aSOPCol.dt1 < dtpDemandBeginDate4.DateTime then Continue;

//      if aSOPCol.sDate =  '12/25-12/31' then
//      begin
//        aKeyICItemSupplyLine := aKeyICItemSupplyReader.GetSupplyLine('01.01.1012959');
//        aKeyICItemSupplyLine.GetQtyAvailx(  myStrToDateTime('2017-12-03'), dqty_a  );
//        Memo1.Lines.Add(Format('dqty_a: %0.0f', [dqty_a]));
//      end;

      dtMemand := aSOPCol.dt2 - 2;
      if dtMemand < aSOPCol.dt1 then
        dtMemand := aSOPCol.dt1;

      Memo1.Lines.Add( FormatDateTime('yyyy-MM-dd', dtMemand) + '    ' + aSOPCol.sDate + ' ......');
    
      // ÿһ������ѭ������10�Σ� ÿ������10%�������һ������ʣ���ȫ��
      for iper := 1 to 10 do
      begin
        // ÿ��SKU�Ϻ�ѭ���� slFGNumber ��Ϊ�Ѱ����ȼ�˳������
        for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
        begin
 

          // ��ȡҪ���ƻ�� ��SKU�Ϻţ� ���ڵ�����
          aSOPReader.GetDemands(aFGPriorityReader.FList[inumber], aSOPCol.dt1,
            dtpDemandBeginDate4.DateTime, lstDemand);
 
          for iDemand := 0 to lstDemand.Count - 1 do
          begin
            aSOPCol_demand := TSOPCol(lstDemand[iDemand]);

            aSBom := TSBom(aFGPriorityReader.FList.Objects[inumber]);
            if aSBom = nil then Continue;  // ���û��Bom�� ����

            dDemandQty := aSOPCol_demand.DemandQty;
            // ���һ�Σ� ������������
            if iper = 10 then
            begin
              //  ��ȥ���������� ���ϴ�δ��������������ǽ�ȥ
              dqty := dDemandQty - aSOPCol_demand.iQty_ok;
              aSOPCol_demand.iQty_calc := dDemandQty;
            end
            else  // �������� 10%
            begin
              dqty := Round(dDemandQty * 0.1);
              if aSOPCol_demand.iQty_calc + dqty > dDemandQty then
              begin
                //dqty := dDemandQty - aSOPCol_demand.iQty_calc;    
                aSOPCol_demand.iQty_calc := dDemandQty;
              end
              else
              begin
                aSOPCol_demand.iQty_calc := aSOPCol_demand.iQty_calc + dqty;
              end;
              //  ��ȥ���������� ���ϴ�δ��������������ǽ�ȥ
              dqty := aSOPCol_demand.iQty_calc - aSOPCol_demand.iQty_ok;
            end;

            // dqty ���μ��� ��������
            dqty_child_min := -9999;
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);

              if aSBomChild.dUsage = 0 then
              begin
                Continue;
              end;
                   
              // ȡ�ɹ�Ӧ����
              dqty_a := aSBomChild.GetQtyAvail(dtMemand - aSBomChild.FLT);

              if DoubleE( dqty_a , 0 ) then
              begin
                dqty_child_min := 0;
                Break;  // ���������Ϊ0�� ���������
              end;                                               
            
              if DoubleE(dqty_child_min , -9999) or DoubleG(dqty_child_min , Trunc( dqty_a / aSBomChild.dUsage) ) then
              begin
                dqty_child_min := Trunc(dqty_a / aSBomChild.dUsage);
              end;
            end;

            if DoubleLE(dqty_child_min , 0) then
            begin 
              Continue; // �������������������Ϊ0������������һSKU
            end;
                                     
            if DoubleG( dqty , dqty_child_min) then // ����������������
            begin
              dqty := dqty_child_min;
            end;

            // ��������ٿ��������� ����
            aSOPCol_demand.iQty_ok := aSOPCol_demand.iQty_ok + dqty; // ����������
            // ���ӹ�Ӧ���ѷ�����
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
              aSBomChild.AllocQty(aFGPriorityReader.FList[inumber],  dtMemand - aSBomChild.FLT, dqty * aSBomChild.dUsage);
            end;
          end;
        end;
      end;


      // Ƿ�Ϸ��䣬���㲻����ƻ���ȱ���ϣ� slFGNumber ��Ϊ�Ѱ����ȼ�˳������
      for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
      begin
        dtMemand := aSOPCol.dt2 - 2;
        if dtMemand < aSOPCol.dt1 then
          dtMemand := aSOPCol.dt1;
                     
          if (aFGPriorityReader.FList[inumber] = '93.56.3564201') then
          begin
            Sleep(1);
          end;  

        aSOPReader.GetDemands(aFGPriorityReader.FList[inumber], aSOPCol.dt1,
          dtpDemandBeginDate4.DateTime, lstDemand);
        for iDemand := 0 to lstDemand.Count - 1 do
        begin
          aSOPCol_demand := TSOPCol(lstDemand[iDemand]);

          aSBom := TSBom(aFGPriorityReader.FList.Objects[inumber]);
          if aSBom = nil then
          begin
            aSOPCol_demand.AddShortageICItem(aFGPriorityReader.FList[inumber] + ' ��BOM');
            Continue;  // ���û��Bom�� ����
          end;

          // һ�μ����꣬������    // ������ iQty�� ǰ�������Ϊ����ģ�����Ϊ����Ƿ��
          dqty := aSOPCol_demand.iQty - aSOPCol_demand.iQty_ok;
          if DoubleLE(dqty , 0) then Continue; // �����ˣ���Ƿ��

          // dqty ���μ��� ��������
          for ichild := 0 to aSBom.FList.Count - 1 do
          begin
            aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
            if aSBomChild.dUsage = 0 then
            begin
              Continue;
            end;

            dqty_a := aSBomChild.GetQtyAvail2(dtMemand - aSBomChild.FLT);

            if Trunc(dqty_a / aSBomChild.dUsage) < dqty then  // ��������¼��Ƿ����
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty_a);
              smsg := aSBomChild.FList[0] + '(' + aSBomChild.FGroup + '): ' + Format('%.0f', [dqty * aSBomChild.dUsage - dqty_a]);
              aSOPCol_demand.AddShortageICItem(smsg);
            end
            else   //  ��������
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty * aSBomChild.dUsage);
            end; 
          end;
        end;
      end;
      

    end;

    SaveMPS(sfile_save, aSOPReader, aSOPSimReader, aMrpMPSReader, aKeyICItemSupplyReader);

    
    MessageBox(Handle, '���', '��ʾ', 0);

  finally
    aFGPriorityReader.Free;
          
    aSOPReader.Free;
    
    slProjYear.Free;

    for idate := 0 to sldate.Count - 1 do
    begin
      aSOPCol := TSOPCol(sldate.Objects[idate]);
      aSOPCol.Free;
    end;
    sldate.Free;

                      
    aKeyICItemSupplyReader.Free;
    aSBomReader.Free;

    aSOPSimReader.Free;

    aMrpMPSReader.Free;

    lstDemand.Free;

    aSAPStockReader.Free;
  end;
end;

procedure TfrmMRPSimulation.btnSIM2SAP5Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSIM2SAP5.Text := sfile;
end;

procedure TfrmMRPSimulation.btnSIM2SAP5_saveClick(Sender: TObject);
  function GetInsertIdx(slweek: TStringList; ptrSopColHead: PSIMColHead): Integer;
  var
    i: Integer;
    p: PSIMColHead;
  begin
    if slweek.Count = 0 then
    begin
      Result := 0;
      Exit;
    end;
    
    Result := -1;
 
    for i := 0 to slweek.Count - 1 do
    begin
      p := PSIMColHead(slweek.Objects[i]);
      if p^.dt1 > ptrSopColHead^.dt1 then
      begin
        Break;
      end
      else if p^.dt1 = ptrSopColHead^.dt1 then
      begin
        Exit;
      end;
    end;
    Result := i;
  end;

  function IndexOfCol(slweek: TStringList; aSOPSimCol: TSOPSimCol): Integer;
  var
    i: Integer;   
    p: PSIMColHead;
  begin
    Result := -1;
    for i := 0 to slweek.Count - 1 do
    begin
      p := PSIMColHead(slweek.Objects[i]);
      if p^.dt1 = aSOPSimCol.dt1 then
      begin
        Result := i;
        Break;
      end;
    end;
  end;
var
  sfile: string;                   
  ExcelApp, WorkBook: Variant;

  slProjYear: TStringList;
  aSOPSimReader: TSOPSimReader;
  iproj: Integer;
  aSOPSimProj: TSOPSimProj;
  iline: Integer;
  aSOPSimLine: TSOPSimLine;
  irow: Integer;
  icol: Integer;
  imonth: Integer;
  slmonth: TStringList;
  iweek: Integer;
  aSopColHeadPtr: PSIMColHead;
  ptrSopColHead_new: PSIMColHead;
  slweek: TStringList;
  idx: Integer;
  aSOPSimCol: TSOPSimCol;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  slProjYear := TfrmProjYear.GetProjYears;
  aSOPSimReader := TSOPSimReader.Create(leSIM2SAP5.Text, slProjYear);

          
  slweek := TStringList.Create;

  try

    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.Message), '�����ʾ', 0);
        Exit;
      end;
    end;
                
    WorkBook := ExcelApp.WorkBooks.Add;
    ExcelApp.DisplayAlerts := False;


    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;


    try
      for iproj := 0 to aSOPSimReader.ProjCount - 1 do
      begin
        aSOPSimProj := aSOPSimReader.Projs[iproj];
        for imonth := 0 to aSOPSimProj.MonthCount - 1 do
        begin
          slmonth := aSOPSimProj.Months[imonth];
          for iweek := 0 to slmonth.Count - 1 do
          begin
            aSopColHeadPtr := PSIMColHead(slmonth.Objects[iweek]);

            idx := GetInsertIdx(slweek, aSopColHeadPtr);
            if idx >= 0 then
            begin
              ptrSopColHead_new := New(PSIMColHead);
              ptrSopColHead_new^ := aSopColHeadPtr^;
              slweek.InsertObject(idx, ptrSopColHead_new^.sdate, TObject(ptrSopColHead_new));
            end;
          end;
        end;
      end;

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'MATNR';
      ExcelApp.Cells[irow, 2].Value := 'BERID';

      icol :=  3;
      for iweek := 0 to slweek.Count -1 do
      begin
        ptrSopColHead_new := PSIMColHead(slweek.Objects[iweek]);
        ExcelApp.Cells[irow, icol + iweek].Value := FormatDateTime('YYYY', ptrSopColHead_new^.dt1) + Copy( IntToStr(100 + WeekOfTheYear(ptrSopColHead_new^.dt1)), 2, 2);
      end;


      irow := 2;
      for iproj := 0 to aSOPSimReader.ProjCount - 1 do
      begin
        aSOPSimProj := aSOPSimReader.Projs[iproj];
        for iline := 0 to aSOPSimProj.Count - 1 do
        begin
          aSOPSimLine := aSOPSimProj.Items[iline];
          ExcelApp.Cells[irow, 1].Value :=aSOPSimLine.snumber;
          ExcelApp.Cells[irow, 2].Value := aSOPSimLine.sarea;

          icol :=  3;
          for iweek := 0 to aSOPSimLine.Count - 1 do
          begin
            aSOPSimCol := aSOPSimLine.Items[iweek];
            idx := IndexOfCol(slweek, aSOPSimCol);
            if idx >= 0 then
            begin 
              ExcelApp.Cells[irow, icol + idx].Value := aSOPSimCol.qty_a;
            end;
          end;
          irow := irow + 1;
        end;
      end;


      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;      
  finally
    aSOPSimReader.Free;
    slProjYear.Free;

    for iweek := 0 to slweek.Count - 1 do
    begin
      aSopColHeadPtr := PSIMColHead(slweek.Objects[iweek]);
      Dispose(aSopColHeadPtr);
    end;
  end;

  MessageBox(Handle, '���', '��ʾ', 0);
end;

end.
