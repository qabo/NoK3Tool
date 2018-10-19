unit MRPSimulationWin;

(*
1、 库存可以满足需求的物料，SAP ZPPR020 报表不会提取。Load物料供应的时候需要需要提取库存


*)

(*
系统外做模拟MRP与齐套模拟步骤：
  1、导出调整后的要货计划  ZPPR028
  2、导出BOM ZPPR021， 料号从调整后的要货计划取
  3、导出库存 MB52
  4、BOM转SBOM  无需带库存
  5、调整后的要货计划转S&OP格式，注意SKU不要重复
  6、SKU优先级计算， SKU优先级文件里的才会参与齐套分析
  7、模拟MRP
  8、齐套分析 *******8
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
    scategory: string; //物料类别
    sg: string; //通用性
    snumber: string; //物料编码(主料)
    sname: string; //物料描述
    slt: string; //LT
    smc: string; //MC
    sproj: string; //项目
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

  Memo1.Lines.Add('开始保存EXCEL ');
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
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
      ExcelApp.Cells[irow, 1].Value := 'MRP区域';
      ExcelApp.Cells[irow, 2].Value := '项目';
      ExcelApp.Cells[irow, 3].Value := '物料编码';
      ExcelApp.Cells[irow, 4].Value := '标准制式';
      ExcelApp.Cells[irow, 5].Value := '颜色';
      ExcelApp.Cells[irow, 6].Value := '容量';
      ExcelApp.Cells[irow, 7].Value := '内容项';
            

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

        ExcelApp.Cells[irow,7].Value := '要货计划量';
        ExcelApp.Cells[irow + 1, 7].Value := '可供应量';
        ExcelApp.Cells[irow + 2, 7].Value := '要货计划与可供应量差异';
        ExcelApp.Cells[irow + 3, 7].Value := '欠料';

                            
        MergeCells(ExcelApp, irow, 1, irow + 3, 1);
        MergeCells(ExcelApp, irow, 2, irow + 3, 2);
        MergeCells(ExcelApp, irow, 3, irow + 3, 3);
        MergeCells(ExcelApp, irow, 4, irow + 3, 4);
        MergeCells(ExcelApp, irow, 5, irow + 3, 5);  
        MergeCells(ExcelApp, irow, 6, irow + 3, 6);

 
        dt0 := 0;     
        icol := 8;

        if iline = 0 then  // 第一个SKU编码，写列标题
        begin


          // 写标题列 /////////////////////////////////////////
          for idate := 0 to aSOPLine.FList.Count - 1 do
          begin
            aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
            if aSOPCol.dt1 < dtpDemandBeginDate4.DateTime then Continue;

            if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
            begin
              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(dt0)) + '月';
              MergeCells(ExcelApp, 1, icol, 2, icol);
              AddColor(ExcelApp, 1, icol, 2, icol, $CCFFFF);
              icol := icol + 1;
            end;
               
            ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
            ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;   
            icol := icol + 1;

            // 最后一个日期
            if idate = aSOPLine.FList.Count - 1 then
            begin
              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(aSOPCol.dt1)) + '月';
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

          // 最后一个日期
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

    Memo1.Lines.Add('物料供应分配sheet');
    // 把 物料供应 的分配结果显示
    ExcelApp.Sheets[aSOPReader.FProjs.Count + 1].Activate;
    ExcelApp.Sheets[aSOPReader.FProjs.Count + 1].Name := '关键物料供应';

    irow := 3;
    for iline := 0 to aKeyICItemSupplyReader.FList.Count - 1 do
    begin
      aKeyICItemSupplyLine := TKeyICItemSupplyLine(aKeyICItemSupplyReader.FList.Objects[iline]);

      if iline = 0 then
      begin
        ExcelApp.Cells[1, 1].Value := '物料编码';   
        ExcelApp.Cells[1, 2].Value := '物料名称';
        ExcelApp.Cells[1, 3].Value := '子物料';

        ExcelApp.Cells[1, 4].Value := '数量';
        ExcelApp.Cells[1, 5].Value := '半成品库存';
        ExcelApp.Cells[1, 6].Value := '可用库存';
        ExcelApp.Cells[1, 7].Value := '新外购入库';
        ExcelApp.Cells[1, 8].Value := 'MRP当天外购入库';

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
      ExcelApp.Cells[irow, 4].Value := '可用量';
      ExcelApp.Cells[irow + 1, 4].Value := '分配量';
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
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end;
  Memo1.Lines.Add('保存EXCEL 结束');  
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

  Memo1.Lines.Add('开始保存EXCEL ');
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
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
    ExcelApp.Sheets[1].Name := '供应能力';
  
    irow := 1; 
    ExcelApp.Cells[irow, 1].Value := 'MATNR';
    ExcelApp.Cells[irow, 2].Value := 'BERID';      
    ExcelApp.Cells[irow, 3].Value := '产品名称';
    ExcelApp.Cells[irow, 4].Value := '项目'; 

    irow := 2;
    for iline := 0 to aSAPS618Reader.Count - 1 do
    begin
      aSAPS618 := aSAPS618Reader.Items[iline];
        
      ExcelApp.Cells[irow, 1].Value := aSAPS618.FNumber;
      ExcelApp.Cells[irow, 2].Value := aSAPS618.FMrpArea;   
      ExcelApp.Cells[irow, 3].Value := aSAPS618.sname;

      icol := 5;
      if iline = 0 then  // 第一个SKU编码，写列标题
      begin 
        // 写标题列 /////////////////////////////////////////
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

      ExcelApp.Cells[irow, 4].Value := '调整后的要货 计划';
      ExcelApp.Cells[irow + 1, 4].Value := '供应能力';
      ExcelApp.Cells[irow + 2, 4].Value := '差异';
      ExcelApp.Cells[irow + 3, 4].Value := '欠料';

      MergeCells(ExcelApp, irow, 1, irow + 4, 1);
      MergeCells(ExcelApp, irow, 2, irow + 4, 2);
      MergeCells(ExcelApp, irow, 3, irow + 4, 3);
 
      irow := irow + 4;
    end;

  
    ExcelApp.Sheets[1].Activate;
      
    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end;
  Memo1.Lines.Add('保存EXCEL 结束');
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

  Memo1.Lines.Add('获取项目年度 ......');
  slProjYear := TfrmProjYear.GetProjYears;
  sldate := TStringList.Create;

  Memo1.Lines.Add('读取关键物料供应能力 ......');
  aKeyICItemSupplyReader := TKeyICItemSupplyReader.Create(leSupplyEval.Text);

  Memo1.Lines.Add('读取简易BOM ......');
  aSBomReader := TSBomReader.Create(leSBOM.Text);
          
  Memo3.Lines.Add('开始读取 要货计划  调整后的要货计划');
  aSAPS618Reader := TSAPS618Reader.Create(leDemand5.Text, '调整后的要货计划', OnLogEvent);


  Memo1.Lines.Add('读取SKU优先级 ......');
  aFGPriorityReader := TFGPriorityReader.Create(lePriority.Text);
                                          
  Memo1.Lines.Add('读取SAP库存 ......');
  aSAPStockReader := TSAPStockReader.Create(leStock.Text);

  aKeyICItemSupplyReader.SAPStock := aSAPStockReader; // 有些item有库存，所以没得出采购需求， 回复交期里面没有这个料。但是分析的时候没有这个料，程序会认为可供应为0.所以获取物料供应的时候，没有交期的，取库存
         
  try
                                        
    Memo1.Lines.Add('把 Bom 的子项物料， 跟供应能力联系上 ......');
    // 把 Bom 的子项物料， 跟供应能力联系上
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
                        
    Memo1.Lines.Add('匹配成品料号Bom ......');
    // 匹配成品料号Bom
    for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
    begin
      aFGPriorityReader.FList.Objects[inumber] := aSBomReader.GetBom(aFGPriorityReader.FList[inumber]);
      if aFGPriorityReader.FList.Objects[inumber] = nil then
      begin
        Memo1.Lines.Add(aFGPriorityReader.FList[inumber] + '  没有BOM');
      end;
    end;
    
    Memo1.Lines.Add('取日期列表 ......');
    // 取日期列表
//    aSOPReader.GetDateList(sldate);
    aSAPS618Reader.GetDateList(sldate);

    Memo1.Lines.Add('每一个日期计算 ......');
    // 每一个日期计算
    for idate := 0 to sldate.Count - 1 do
    begin     
      aSAPS618ColPtr := PSAPS618Col(sldate.Objects[idate]); // 日期
      if aSAPS618ColPtr^.dt1 < dtpDemandBeginDate4.DateTime then Continue;
 
      dtMemand := aSAPS618ColPtr^.dt2 - 2;
      if dtMemand < aSAPS618ColPtr^.dt1 then
        dtMemand := aSAPS618ColPtr^.dt1;

      Memo1.Lines.Add( FormatDateTime('yyyy-MM-dd', dtMemand) + '    ' + aSAPS618ColPtr^.sweek + ' ......');
    
      // 每一个日期循环计算10次， 每次满足10%需求，最后一次满足剩余的全部
      for iper := 1 to 10 do
      begin
        // 每个SKU料号循环， slFGNumber 认为已按优先级顺序排序
        for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
        begin 
          // 获取要货计划里， 此SKU料号， 日期的需求
          aSAPS618Reader.GetDemands(aFGPriorityReader.FList[inumber], aSAPS618ColPtr^.dt1,
            dtpDemandBeginDate4.DateTime, lstDemand);
 
          for iDemand := 0 to lstDemand.Count - 1 do
          begin
            aSAPS618ColPtr_demand := PSAPS618Col(lstDemand[iDemand]);

            aSBom := TSBom(aFGPriorityReader.FList.Objects[inumber]);
            if aSBom = nil then Continue;  // 如果没有Bom， 跳过

            dDemandQty := aSAPS618ColPtr_demand^.dDemandQty;
            // 最后一次， 满足余下所有
            if iper = 10 then
            begin
              //  减去齐套数量， 把上次未能满足的余量考虑进去
              dqty := dDemandQty - aSAPS618ColPtr_demand^.dQty_ok;
              aSAPS618ColPtr_demand^.dQty_calc := dDemandQty;
            end
            else  // 否则，满足 10%
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
              //  减去齐套数量， 把上次未能满足的余量考虑进去
              dqty := aSAPS618ColPtr_demand^.dQty_calc - aSAPS618ColPtr_demand^.dQty_ok;
            end;

            // dqty 本次计算 齐套数量
            dqty_child_min := -9999;
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
                                                                          
              // 取可供应数量
              dqty_a := aSBomChild.GetQtyAvail(dtMemand - aSBomChild.FLT);

              if DoubleE( dqty_a , 0 ) then
              begin
                dqty_child_min := 0;
                Break;  // 如果可用量为0， 不许继续了
              end;                                               
            
              if DoubleE(dqty_child_min , -9999) or DoubleG(dqty_child_min , Trunc( dqty_a / aSBomChild.dUsage) ) then
              begin
                dqty_child_min := Trunc(dqty_a / aSBomChild.dUsage);
              end;
            end;

            if DoubleLE(dqty_child_min , 0) then
            begin 
              Continue; // 可满足齐套需求的物料为0，继续计算下一SKU
            end;
                                     
            if DoubleG( dqty , dqty_child_min) then // 供给不能满足需求，
            begin
              dqty := dqty_child_min;
            end;

            // 计算出最少可齐套数， 分配
            aSAPS618ColPtr_demand^.dQty_ok := aSAPS618ColPtr_demand^.dQty_ok + dqty; // 齐套数增加
            // 增加供应的已分配量
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
              aSBomChild.AllocQty(aFGPriorityReader.FList[inumber], dtMemand - aSBomChild.FLT, dqty * aSBomChild.dUsage); 
            end;
          end;
        end;
      end;


      // 欠料分配，计算不满足计划所缺物料， slFGNumber 认为已按优先级顺序排序
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
            AddShortageICItem(aSAPS618ColPtr_demand, aFGPriorityReader.FList[inumber] + ' 无BOM');
            Continue;  // 如果没有Bom， 跳过
          end;

          // 一次计算完，不分批    // 这里用 iQty， 前面的需求为满足的，不算为当天欠料
          dqty := aSAPS618ColPtr_demand^.dQty - aSAPS618ColPtr_demand^.dQty_ok;
          if DoubleLE(dqty , 0) then Continue; // 满足了，无欠料

          // dqty 本次计算 齐套数量
          for ichild := 0 to aSBom.FList.Count - 1 do
          begin
            aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);

            dqty_a := aSBomChild.GetQtyAvail2(dtMemand - aSBomChild.FLT);

            if Trunc(dqty_a / aSBomChild.dUsage) < dqty then  // 不够，记录所欠数量
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty_a);
              smsg := aSBomChild.FList[0] + '(' + aSBomChild.FGroup + '): ' + Format('%.0f', [dqty * aSBomChild.dUsage - dqty_a]);
              AddShortageICItem(aSAPS618ColPtr_demand, smsg);
            end
            else   //  够，分配
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty * aSBomChild.dUsage);
            end; 
          end;
        end;
      end; 
    end;

    SaveSAP(sfile_save, aSAPS618Reader);

    
    MessageBox(Handle, '完成', '提示', 0);

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
  ExcelApp.DisplayAlerts := False;
//  while ExcelApp.Sheets.Count > 1 do
//  begin
//    ExcelApp.Sheets.Delete(1);
//  end;

  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := 'SKU优先顺序';

  irow := 1;
  ExcelApp.Cells[irow, 1].Value := '产品编码';
  ExcelApp.Cells[irow, 2].Value := '优先级';

  irow := 2;
  for inumber := 0 to slFGNumber.Count - 1 do
  begin
    ExcelApp.Cells[irow, 1].Value := slFGNumber[inumber];
    ExcelApp.Cells[irow, 2].Value := inumber + 1;
    irow := irow + 1;
  end;


  try

    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

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

  MessageBox(Handle, '完成', '提示', 0);

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
  
  MessageBox(Self.Handle, '完成', '提示', 0);
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
  MessageBox(Handle, '完成', '提示', 0);
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

  // 简化的MRP计算，不考虑低位码
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

   

  Memo3.Lines.Add('开始读取 库存');
  aSAPStockReader := TSAPStockReader.Create(leSAPStock3.Text, OnLogEvent);

  if not aSAPStockReader.ReadOk then
  begin
    aSAPStockReader.Free;
    MessageBox(Handle, '读取库存失败', '警告', 0);
    Exit;
  end;

  Memo3.Lines.Add('开始读取 BOM');
  aSAPBomReader := TSAPBomReader2.Create(leSAPBom3.Text, OnLogEvent);
  
  Memo3.Lines.Add('开始读取 要货计划');
  aSAPS618Reader := TSAPS618Reader.Create(leDemand3.Text, '调整后的要货计划', OnLogEvent);
                      
  Memo3.Lines.Add('开始读取 物料信息');
  aSAPMaterialReader2 := TSAPMaterialReader2.Create(leICItem.Text, OnLogEvent);
  


  aSAPStockSum := TSAPStockSum.Create;
  aSAPStockReader.SumTo(aSAPStockSum);   
  aSAPStockReader.Free;
  
  lstDemand := TStringList.Create;
  lstNumber := TStringList.Create;
                          
  Memo3.Lines.Add('整理要货计划需求');
  for iLine := 0 to aSAPS618Reader.Count - 1 do
  begin
    aSAPS618 := aSAPS618Reader.Items[iLine];
    for iDate := 0 to aSAPS618.Count - 1 do
    begin
      aSAPS618ColPtr := aSAPS618.Items[iDate];
      
      // 数量为0的不必添加
      if DoubleE( aSAPS618ColPtr^.dqty, 0 ) then Continue;
      //日期早于开始日期的不计算，跳过
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
                            
  Memo3.Lines.Add('开始模拟MRP计算');
  try
    bLoop := True;
    while bLoop do
    begin
      bLoop := False;
      
      //备份需求
      lstDemand_tmp := TStringList.Create;
      for iMrpUnit := 0 to lstDemand.Count - 1 do
      begin
        lstDemand_tmp.AddObject(lstDemand[iMrpUnit], lstDemand.Objects[iMrpUnit]);
      end;
      lstDemand.Clear;

      //排序，按日期
      lstDemand_tmp.CustomSort(StringListSortCompare_DateTime);
 
      for iMrpUnit := 0 to lstDemand_tmp.Count - 1 do
      begin
        aMRPUnitPtr := PMRPUnit(lstDemand_tmp.Objects[iMrpUnit]);
        if not aMRPUnitPtr^.bExpend then
        begin
          // 根节点  不考虑库存 ////////////////////////////////////////////////////
          if aMRPUnitPtr^.aParentBom = nil then
          begin   
            aMRPUnitPtr^.bExpend := True;      
            aMRPUnitPtr^.aBom := aSAPBomReader.GetSapBom(lstDemand_tmp[iMrpUnit], '');
            if aMRPUnitPtr^.aBom = nil then  // 没有BOM，异常，记录日志
            begin
              Memo3.Lines.Add(lstDemand_tmp[iMrpUnit] + ' 没有BOM');
              Continue;
            end;
            for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
            begin
              aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
              // BOM 上层产生的需求  ， 需减去上层已分配库存
//              dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapItemGroup.Items[0].dusage;     
              dQty := aMRPUnitPtr^.dQty * aSapItemGroup.Items[0].dusage;      // 每一层的需求，都是扣减了已分配库存的
              // 计算可用库存
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
              
              //展开需求到下层
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
                  aMRPUnitPtr_Dep^.dQty := dQty * aSapBomChild.dPer / 100;  // 半成品替代料按配比分
                end;
                
                aMRPUnitPtr_Dep^.dQtyStock := aSapBomChild.FStock;   
                aSapBomChild.FStock := 0;  // 赋值后清0，否则遗留的数值会影响下一轮计算
                aMRPUnitPtr_Dep^.dQtyStockParent := aMRPUnitPtr^.dQtyStock + aMRPUnitPtr^.dQtyStockParent;  // 父项库存
                aMRPUnitPtr_Dep^.bExpend := False;
                aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                lstDemand.AddObject(aSapBomChild.FNumber, TObject(aMRPUnitPtr_Dep));
              end;
              bLoop := True;
            end;
          end
          else
          // 非根节点  不考虑库存 //////////////////////////////////////////////////      
          begin
            // 叶子节点， 直接添加到需求////////////////////////////////////////////
            if aMRPUnitPtr^.aBom.ChildCount = 0 then
            begin
              aMRPUnitPtr_Dep := New(PMRPUnit);
              aMRPUnitPtr_Dep^ := aMRPUnitPtr^;
              lstDemand.AddObject(lstDemand_tmp[iMrpUnit], TObject(aMRPUnitPtr_Dep));
            end        
            // 非叶子节点， 非根节点，即是半成品，展开需求//////////////////////////
            else    // 这里需考虑半成品库存 ////////////////////////////////////////
            begin
              aMRPUnitPtr^.bExpend := True;
 
              for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
              begin
                aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];

                // BOM 上层产生的需求  ， 需减去上层已分配库存
                //dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapItemGroup.Items[0].dusage;
                dQty := aMRPUnitPtr^.dQty* aSapItemGroup.Items[0].dusage;  // 每一层的需求，都是扣减了已分配库存的

                // 计算可用库存
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
                    aMRPUnitPtr_Dep^.dQty := dQty;  // 半成品替代料按配比分
                  end
                  else
                  begin
                    aMRPUnitPtr_Dep^.dQty := dQty * aSapBomChild.dPer / 100;  // 半成品替代料按配比分
                  end;
                  aMRPUnitPtr_Dep^.dQtyStock := aSapBomChild.FStock;
                  aSapBomChild.FStock := 0;  // 赋值后清0，否则遗留的数值会影响下一轮计算

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
    
   
               
    //排序，按日期
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


    // 保存 //////////////////////////////////////////////////////////////////////

    Memo3.Lines.Add('开始保存模拟MRP计算结果，只保存关键物料');
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
    ExcelApp.DisplayAlerts := False;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '模拟MRP';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '产品编码';
    ExcelApp.Cells[irow, 2].Value := '日期';
    ExcelApp.Cells[irow, 3].Value := '数量';    
    ExcelApp.Cells[irow, 4].Value := '库存';

    ExcelApp.Columns[1].ColumnWidth := 17;
    ExcelApp.Columns[2].ColumnWidth := 13;
    ExcelApp.Columns[3].ColumnWidth := 9;
    ExcelApp.Columns[4].ColumnWidth := 9;

    ProgressBar1.Max := lstDemand.Count;
    ProgressBar1.Position := 0;

    sl := TStringList.Create;
    sline := '物料编码'#9'物料名称'#9'日期'#9'需求数量'#9'库存'#9'父项库存'#9'ABC标志'#9'父项物料'#9'父项名称';
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
        Format('%0.0f'#9'%0.0f'#9'%0.0f', [aMRPUnitPtr^.dQty, aMRPUnitPtr^.dQtyStock, aMRPUnitPtr^.dQtyStockParent]) + #9 +  // 3行
        sabc + #9 +
        snumber_p + #9 +
        sname_p;
      sl.Add(sline);

      // 只写关键物料
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
    ExcelApp.Sheets[iSheet].Name := '模拟MRP交期回复';
                  
    ExcelApp.Cells[1, 1].Value := 'MRP模拟--要货计划生成物料需求计划';
    ExcelApp.Cells[2, 1].Value := '模拟日期:';                   
    ExcelApp.Cells[2, 2].Value := FormatDateTime('yyyy-MM-dd HH:mm', Now);
    ExcelApp.Cells[3, 1].Value := '库存日期:';

    
    irow := 4;
    ExcelApp.Cells[irow, 1].Value := '物料编码';
    ExcelApp.Cells[irow, 2].Value := '物料名称';
    ExcelApp.Cells[irow, 3].Value := '父项分配库存';
    ExcelApp.Cells[irow, 4].Value := '分配库存';
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
      // 只写关键物料
      if aMRPUnitPtr^.aBom.abc = '' then
      begin
        ProgressBar1.Position := ProgressBar1.Position + 1;
        Continue;
      end;
            
      aSAPMaterialRecordPtr := aSAPMaterialReader2.GetSAPMaterialRecord(aMRPUnitPtr^.snumber);
      
      ExcelApp.Cells[irow, 1].Value := lstNumber[iNumber];
      ExcelApp.Cells[irow, 2].Value := aMRPUnitPtr^.sname;
      ExcelApp.Cells[irow, 3].Value := aMRPUnitPtr^.dQtyStockParent;  // 父项库存
      ExcelApp.Cells[irow, 4].Value := aSAPStockSum.GetStock(lstNumber[iNumber]);
      ExcelApp.Cells[irow, 5].Value := '要货计划需求量';
      ExcelApp.Cells[irow + 1, 5].Value := '可供应量';
      ExcelApp.Cells[irow + 2, 5].Value := '供应差异';



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
          ExcelApp.Cells[irow, icol].Value := GetDemand(lstDemand, lstNumber[iNumber], dt, dtNextMonday);  // 下周一之前
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);

          dt := dtNextMonday;   // 日期定位到下周1
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
    ExcelApp.Cells[4, icolMax + 2].Value := '采购';
    
    AddBorder(ExcelApp, 1, 1, irow - 1, icolMax + 2);
    AddColor(ExcelApp, 3, 6, 3, icolMax + 2, $E8DEB7);
    AddColor(ExcelApp, 4, 1, 4, icolMax + 2, $E8DEB7);

    ExcelApp.Cells[irow + 2, 1].Value := '说明：';
    ExcelApp.Cells[irow + 3, 1].Value := '要货计划需求量：指净需求量，比如成品需求100（成品库存不考虑，因提供的是净要货计划已经考虑过了），中间半成品库存20，原料本身库存10，则要货计划需求量=100-20-10=70；';
    ExcelApp.Cells[irow + 4, 1].Value := '可供应量：采购回复，库存日期以后供应商可以交货之数量；如：库存日期魅族库存=100，要货计划需求量为100（此时最上层需求实为200），供应商确认可交付80，则可供应量=80；';
    ExcelApp.Cells[irow + 5, 1].Value := '供应差异：即当天结余=前天剩余量+当天可供应量-当天要货计划需求量';

                  
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
    ExcelApp.Sheets[iSheet].Name := '调整后的要货计划';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
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

          if stitle <> '物料工厂MRP 范围物料描述计划类型字符订单测量单位合计' then
          begin        
            Memo1.Lines.Add(sSheet +'  不是  调整后的要货计划  格式');
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
        ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
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
    ExcelApp.Sheets[iSheet].Name := '可用库存';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
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

          if stitle <> '工厂库存地点仓储地点的描述物料物料描述非限制使用的库存中转和转移' then
          begin        
            Memo1.Lines.Add(sSheet +'  不是  SAP导出库存  格式');
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
        ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
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
    ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
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
                        
          if stitle <> '母件物料编码母件物料描述工厂用途代工厂母件L/T' then
          begin
            Memo1.Lines.Add(sSheet +'  不是SAP导出BOM格式');
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
        ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
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
    ExcelApp.Sheets[iSheet].Name := '计算过程日志';


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
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

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

  MessageBox(Handle, '完成', '提示', 0);
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

  Memo1.Lines.Add('获取项目年度 ......');
  slProjYear := TfrmProjYear.GetProjYears;
  sldate := TStringList.Create;

  aMrpMPSReader := TMrpMPSReader.Create(leLastMPS.Text);
  Memo1.Lines.Add('读取上周MRP录入MPS ......');

  Memo1.Lines.Add('读取上周S&OP模拟 ......');
  aSOPSimReader := TSOPSimReader.Create(leLastSOPAnalysis.Text, slProjYear);

  Memo1.Lines.Add('读取关键物料供应能力 ......');
  aKeyICItemSupplyReader := TKeyICItemSupplyReader.Create(leSupplyEval.Text);

  Memo1.Lines.Add('读取简易BOM ......');
  aSBomReader := TSBomReader.Create(leSBOM.Text);
             
  Memo1.Lines.Add('读取要货计划 ......');
  aSOPReader := TSOPReader.Create(slProjYear, leDemand5.Text);
 

  Memo1.Lines.Add('读取SKU优先级 ......');
  aFGPriorityReader := TFGPriorityReader.Create(lePriority.Text);
                                          
  Memo1.Lines.Add('读取SAP库存 ......');
  aSAPStockReader := TSAPStockReader.Create(leStock.Text);

  aKeyICItemSupplyReader.SAPStock := aSAPStockReader; // 有些item有库存，所以没得出采购需求， 回复交期里面没有这个料。但是分析的时候没有这个料，程序会认为可供应为0.所以获取物料供应的时候，没有交期的，取库存
         
  try
                                        
    Memo1.Lines.Add('把 Bom 的子项物料， 跟供应能力联系上 ......');
    // 把 Bom 的子项物料， 跟供应能力联系上
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
                        
    Memo1.Lines.Add('匹配成品料号Bom ......');
    // 匹配成品料号Bom
    for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
    begin
      aFGPriorityReader.FList.Objects[inumber] := aSBomReader.GetBom(aFGPriorityReader.FList[inumber]);
      if aFGPriorityReader.FList.Objects[inumber] = nil then
      begin
        Memo1.Lines.Add(aFGPriorityReader.FList[inumber] + '  没有BOM');
      end;
    end;
    
    Memo1.Lines.Add('取日期列表 ......');
    // 取日期列表
    aSOPReader.GetDateList(sldate);
//    aSAPS618Reader.GetDateList(sldate);

    Memo1.Lines.Add('每一个日期计算 ......');
    // 每一个日期计算
    for idate := 0 to sldate.Count - 1 do
    begin     
      aSOPCol := TSOPCol(sldate.Objects[idate]); // 日期
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
    
      // 每一个日期循环计算10次， 每次满足10%需求，最后一次满足剩余的全部
      for iper := 1 to 10 do
      begin
        // 每个SKU料号循环， slFGNumber 认为已按优先级顺序排序
        for inumber := 0 to aFGPriorityReader.FList.Count - 1 do
        begin
 

          // 获取要货计划里， 此SKU料号， 日期的需求
          aSOPReader.GetDemands(aFGPriorityReader.FList[inumber], aSOPCol.dt1,
            dtpDemandBeginDate4.DateTime, lstDemand);
 
          for iDemand := 0 to lstDemand.Count - 1 do
          begin
            aSOPCol_demand := TSOPCol(lstDemand[iDemand]);

            aSBom := TSBom(aFGPriorityReader.FList.Objects[inumber]);
            if aSBom = nil then Continue;  // 如果没有Bom， 跳过

            dDemandQty := aSOPCol_demand.DemandQty;
            // 最后一次， 满足余下所有
            if iper = 10 then
            begin
              //  减去齐套数量， 把上次未能满足的余量考虑进去
              dqty := dDemandQty - aSOPCol_demand.iQty_ok;
              aSOPCol_demand.iQty_calc := dDemandQty;
            end
            else  // 否则，满足 10%
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
              //  减去齐套数量， 把上次未能满足的余量考虑进去
              dqty := aSOPCol_demand.iQty_calc - aSOPCol_demand.iQty_ok;
            end;

            // dqty 本次计算 齐套数量
            dqty_child_min := -9999;
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);

              if aSBomChild.dUsage = 0 then
              begin
                Continue;
              end;
                   
              // 取可供应数量
              dqty_a := aSBomChild.GetQtyAvail(dtMemand - aSBomChild.FLT);

              if DoubleE( dqty_a , 0 ) then
              begin
                dqty_child_min := 0;
                Break;  // 如果可用量为0， 不许继续了
              end;                                               
            
              if DoubleE(dqty_child_min , -9999) or DoubleG(dqty_child_min , Trunc( dqty_a / aSBomChild.dUsage) ) then
              begin
                dqty_child_min := Trunc(dqty_a / aSBomChild.dUsage);
              end;
            end;

            if DoubleLE(dqty_child_min , 0) then
            begin 
              Continue; // 可满足齐套需求的物料为0，继续计算下一SKU
            end;
                                     
            if DoubleG( dqty , dqty_child_min) then // 供给不能满足需求，
            begin
              dqty := dqty_child_min;
            end;

            // 计算出最少可齐套数， 分配
            aSOPCol_demand.iQty_ok := aSOPCol_demand.iQty_ok + dqty; // 齐套数增加
            // 增加供应的已分配量
            for ichild := 0 to aSBom.FList.Count - 1 do
            begin
              aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
              aSBomChild.AllocQty(aFGPriorityReader.FList[inumber],  dtMemand - aSBomChild.FLT, dqty * aSBomChild.dUsage);
            end;
          end;
        end;
      end;


      // 欠料分配，计算不满足计划所缺物料， slFGNumber 认为已按优先级顺序排序
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
            aSOPCol_demand.AddShortageICItem(aFGPriorityReader.FList[inumber] + ' 无BOM');
            Continue;  // 如果没有Bom， 跳过
          end;

          // 一次计算完，不分批    // 这里用 iQty， 前面的需求为满足的，不算为当天欠料
          dqty := aSOPCol_demand.iQty - aSOPCol_demand.iQty_ok;
          if DoubleLE(dqty , 0) then Continue; // 满足了，无欠料

          // dqty 本次计算 齐套数量
          for ichild := 0 to aSBom.FList.Count - 1 do
          begin
            aSBomChild :=  TSBomChild(aSBom.FList.Objects[ichild]);
            if aSBomChild.dUsage = 0 then
            begin
              Continue;
            end;

            dqty_a := aSBomChild.GetQtyAvail2(dtMemand - aSBomChild.FLT);

            if Trunc(dqty_a / aSBomChild.dUsage) < dqty then  // 不够，记录所欠数量
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty_a);
              smsg := aSBomChild.FList[0] + '(' + aSBomChild.FGroup + '): ' + Format('%.0f', [dqty * aSBomChild.dUsage - dqty_a]);
              aSOPCol_demand.AddShortageICItem(smsg);
            end
            else   //  够，分配
            begin
              aSBomChild.AllocQty2( dtMemand - aSBomChild.FLT , dqty * aSBomChild.dUsage);
            end; 
          end;
        end;
      end;
      

    end;

    SaveMPS(sfile_save, aSOPReader, aSOPSimReader, aMrpMPSReader, aKeyICItemSupplyReader);

    
    MessageBox(Handle, '完成', '提示', 0);

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
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.Message), '金蝶提示', 0);
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
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

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

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
