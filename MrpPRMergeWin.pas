unit MrpPRMergeWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommUtils;

type
  TfrmMrpPRMerge = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leMrpPR: TLabeledEdit;
    btnMrpPR: TButton;
    Memo1: TMemo;
    GroupBox1: TGroupBox;
    lstManFile: TListBox;
    btnAdd: TButton;
    btnDel: TButton;
    ProgressBar1: TProgressBar;
    procedure btnMrpPRClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure lstManFileDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure lstManFileDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

const
//  CSPN = 'PN';
  CSSel = '选择';
  CSItemID = '内码';
  CSNumber = '物料编码';
  CSName = '名称';
  CSNeedDate = '需求日期';
  CSArriveDate = '到料日期';
  CSOrderDate = '下单日期';
  CSAdviceOrderDate = '建议下单日期';
  CSUnit = '单位';
  CSRoughNeed = '毛需求数';
  CSOPO = '可用未结订单';
  CSStock = '库存数';
  CSStockAvailble = '可用库存';
  CSNetNeed = '净需求';
  CSSPQ = 'SPQ';
  CSMOQ =	'MOQ';
  CSLT = 'L/T';
  CSPlanQty = '计划定货量';
  CSOldQty = 'old数量';
  CSQty = '数量';
  CSErpClass = '物料属性';
  CSNeedType = '需求类型';
  CSSignStatus = '签核状态';
  CSProj = '项目';
  CSPlanOrderNo = '计划订单号';
  CSMC = 'MC';

//  CSPN_man = 'PN';
  CSNumber99_man = '物料编码99';
  CSOrderDate_man = '下单日期';
  CSNetNeed_man = '净需求';
  CSNetNeedMan_man = '更新需求数量';
  CSSignStatus_man = '';
  CSNote_man = '备注';



type 
  TPRRecord = packed record
//    sPN: string;
    sSel: string;
    sItemID: string;
    sNumber: string;
    sName: string;
    sNeedDate: string;
    sArriveDate: string;
    sOrderDate: string;
    sAdviceOrderDate: string;
    sUnit: string;
    sRoughNeed: string;
    sOPO: string;
    sStock: string;
    sStockAvailble: string;
    sNetNeed: string;
    sNetNeedMan: string;
    sSPQ: string;
    sMOQ: string;
    sLT: string;
    sPlanQty: string;
    sOldQty: string;
    sQty: string;
    sErpClass: string;
    sNeedType: string;
    sSignStatus: string;
    sProj: string;
    sPlanOrderNo: string;
    sMC: string;
  end;
  PPRRecord = ^TPRRecord;

  TPRManRecord = packed record
//    sPN: string;
    sNnumber99: string;
    sOrderDate: string;
    sNetNeed: string;
    sNetNeedMan: string;
    sSignStatus: string;
    sNote: string; 
  end;
  PPRManRecord = ^TPRManRecord;

class procedure TfrmMrpPRMerge.ShowForm;
var
  frmHWCinSum: TfrmMrpPRMerge;
begin
  frmHWCinSum := TfrmMrpPRMerge.Create(nil);
  try
    frmHWCinSum.ShowModal;
  finally
    frmHWCinSum.Free;
  end;
end;

procedure TfrmMrpPRMerge.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMrpPRMerge.btnMrpPRClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMrpPR.Text := sfile;
end;
     
procedure TfrmMrpPRMerge.btnAddClick(Sender: TObject);
var
  sfile: string;
  sl: TStringList;
  i: Integer;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  sl := TStringList.Create;
  try
    sl.Text := StringReplace(sfile, ';', #13#10, [rfReplaceAll]);
    for i := 0 to sl.Count - 1 do
    begin
      lstManFile.Items.Add(sl[i]);
    end;
  finally
    sl.Free;
  end;
end;

procedure TfrmMrpPRMerge.btnDelClick(Sender: TObject);
begin
  if lstManFile.SelCount = 0 then Exit;
  if MessageBox(Handle, '确认删除选中项目？', '提示', MB_YESNO) <> MrYes then
  begin
    Exit;
  end;
  lstManFile.DeleteSelected;
end;

function IndexOfColAlert(ExcelApp: Variant; const s: string): Integer;
var
  icol: Integer;
  svalue: string;
begin
  Result := -1;
  for icol := 1 to 50 do
  begin
    svalue := ExcelApp.Cells[1, icol].Value;
    svalue := Trim(svalue);
    if UpperCase(svalue) = UpperCase(s) then
    begin
      Result := icol;
      Break;
    end;
  end;
  if Result = -1 then
  begin
    raise Exception.Create(s + ' 列不存在');
  end;
end;

function IndexOfMrpPR(lst: TList; aManPtr: PPRManRecord): Integer;
var
  i: Integer;
  aMrpPtr: PPRRecord;
begin
  Result := -1;
  for i := 0 to lst.Count - 1 do
  begin
    aMrpPtr := PPRRecord(lst[i]);
    if (aMrpPtr^.sNumber = aManPtr^.sNnumber99)
      and (aMrpPtr^.sOrderDate = aManPtr^.sOrderDate) then
//      and (aMrpPtr^.sSignStatus = aManPtr^.sSignStatus) then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure GetManPr(lstManPR: TList; aMrpPtr: PPRRecord; lst: TList);
var
  i: Integer;
  aManPrt: PPRManRecord;
begin
  lst.Clear;
  for i := 0 to lstManPR.Count - 1 do
  begin
    aManPrt := PPRManRecord(lstManPR[i]);
    if (aManPrt^.sNnumber99 = aMrpPtr^.sNumber)
      and (aManPrt^.sOrderDate = aMrpPtr^.sOrderDate) then
//      and (aManPrt^.sSignStatus = aMrpPtr^.sSignStatus) then
    begin
      lst.Add(aManPrt); 
    end;
  end;
end;


procedure TfrmMrpPRMerge.btnSave2Click(Sender: TObject);
const
  CSTitle = '日期单据编号收料仓库加工材料长代码加工材料名称实收数量备注订单单号';
var
  ExcelApp, WorkBook: Variant;       
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;


//  iPN: Integer; 
  iSel: Integer;
  iItemID: Integer;
  iNumber: Integer; 
  iName: Integer; 
  iNeedDate: Integer;
  iArriveDate: Integer;
  iOrderDate: Integer;
  iAdviceOrderDate: Integer; 
  iUnit: Integer; 
  iRoughNeed: Integer; 
  iOPO: Integer; 
  iStock: Integer; 
  iStockAvailble: Integer; 
  iNetNeed: Integer; 
  iSPQ: Integer; 
  iMOQ: Integer; 
  iLT: Integer; 
  iPlanQty: Integer; 
  iOldQty: Integer; 
  iQty: Integer; 
  iErpClass: Integer; 
  iNeedType: Integer; 
  iSignStatus: Integer; 
  iProj: Integer; 
  iPlanOrderNo: Integer; 
  iMC: Integer;


//  iPN_man: Integer;
  iNumber99_man: Integer;
  iOrderDate_man: Integer;
  iNetNeed_man: Integer;
  iNetNeedMan_man: Integer;
  iNote_man: Integer;

  lstManPR: TList;
  aManPtr: PPRManRecord;
                
  lst: TList;

  lstMrpPR: TList;
  aMrpPtr: PPRRecord;
  irow: Integer;
  snumber: string;
  ifile: Integer;
  dqty: Double;

  iMan: Integer;
  idx: Integer;

 
  i: Integer;   
  sfile: string;
begin

  if not ExcelSaveDialog(sfile) then Exit;

  if not FileExists(leMrpPR.Text) then
  begin
    MessageBox(Handle, 'Mrp PR 文件不存在', '金蝶提示', 0);
    Exit;
  end;

  lstMrpPR := TList.Create;
  lstManPR := TList.Create;
  try

    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';


    try
      WorkBook := ExcelApp.WorkBooks.Open(leMrpPR.Text);

      try
        iSheetCount := ExcelApp.Sheets.Count;
        for iSheet := 1 to iSheetCount do
        begin
          if not ExcelApp.Sheets[iSheet].Visible then Continue;

          ExcelApp.Sheets[iSheet].Activate;
                              
          sSheet := ExcelApp.Sheets[iSheet].Name;
          if sSheet <> 'PR Sum' then Continue;
  

          //iPN := IndexOfColAlert(ExcelApp, CSPN);
          iSel := IndexOfColAlert(ExcelApp, CSSel);
          iItemID := IndexOfColAlert(ExcelApp, CSItemID);
          iNumber := IndexOfColAlert(ExcelApp, CSNumber);
          iName := IndexOfColAlert(ExcelApp, CSName);
          iNeedDate := IndexOfColAlert(ExcelApp, CSNeedDate);
          iArriveDate := IndexOfColAlert(ExcelApp, CSArriveDate);
          iOrderDate := IndexOfColAlert(ExcelApp, CSOrderDate);
          iAdviceOrderDate := IndexOfColAlert(ExcelApp, CSAdviceOrderDate);
          iUnit := IndexOfColAlert(ExcelApp, CSUnit);
          iRoughNeed := IndexOfColAlert(ExcelApp, CSRoughNeed);
          iOPO := IndexOfColAlert(ExcelApp, CSOPO);
          iStock := IndexOfColAlert(ExcelApp, CSStock);
          iStockAvailble := IndexOfColAlert(ExcelApp, CSStockAvailble);
          iNetNeed := IndexOfColAlert(ExcelApp, CSNetNeed);
          iSPQ := IndexOfColAlert(ExcelApp, CSSPQ);
          iMOQ := IndexOfColAlert(ExcelApp, CSMOQ);
          iLT := IndexOfColAlert(ExcelApp, CSLT);
          iPlanQty := IndexOfColAlert(ExcelApp, CSPlanQty);
          iOldQty := IndexOfColAlert(ExcelApp, CSOldQty);
          iQty := IndexOfColAlert(ExcelApp, CSQty);
          iErpClass := IndexOfColAlert(ExcelApp, CSErpClass);
          iNeedType := IndexOfColAlert(ExcelApp, CSNeedType);
          iSignStatus := IndexOfColAlert(ExcelApp, CSSignStatus);
          iProj := IndexOfColAlert(ExcelApp, CSProj);
          iPlanOrderNo := IndexOfColAlert(ExcelApp, CSPlanOrderNo);
          iMC := IndexOfColAlert(ExcelApp, CSMC);

          irow := 2;
          snumber := ExcelApp.Cells[irow, iNumber].Value;
          while snumber <> '' do
          begin
            aMrpPtr := New(PPRRecord);
            lstMrpPR.Add(aMrpPtr);
 
//            aMrpPtr^.sPN := ExcelApp.Cells[irow, iPN].Value;
            aMrpPtr^.sSel := ExcelApp.Cells[irow, iSel].Value;
            aMrpPtr^.sItemID := ExcelApp.Cells[irow, iItemID].Value;
            aMrpPtr^.sNumber := snumber;
            aMrpPtr^.sName := ExcelApp.Cells[irow, iName].Value;
            aMrpPtr^.sNeedDate := ExcelApp.Cells[irow, iNeedDate].Value;
            aMrpPtr^.sArriveDate := ExcelApp.Cells[irow, iArriveDate].Value;
            aMrpPtr^.sOrderDate := ExcelApp.Cells[irow, iOrderDate].Value;
            aMrpPtr^.sAdviceOrderDate := ExcelApp.Cells[irow, iAdviceOrderDate].Value;
            aMrpPtr^.sUnit := ExcelApp.Cells[irow, iUnit].Value;
            aMrpPtr^.sRoughNeed := ExcelApp.Cells[irow, iRoughNeed].Value;
            aMrpPtr^.sOPO := ExcelApp.Cells[irow, iOPO].Value;
            aMrpPtr^.sStock := ExcelApp.Cells[irow, iStock].Value;
            aMrpPtr^.sStockAvailble := ExcelApp.Cells[irow, iStockAvailble].Value;
            aMrpPtr^.sNetNeed := ExcelApp.Cells[irow, iNetNeed].Value;
            aMrpPtr^.sNetNeedMan := aMrpPtr^.sNetNeed;
            aMrpPtr^.sSPQ := ExcelApp.Cells[irow, iSPQ].Value;
            aMrpPtr^.sMOQ := ExcelApp.Cells[irow, iMOQ].Value;
            aMrpPtr^.sLT := ExcelApp.Cells[irow, iLT].Value;
            aMrpPtr^.sPlanQty := ExcelApp.Cells[irow, iPlanQty].Value;
            aMrpPtr^.sOldQty := ExcelApp.Cells[irow, iOldQty].Value;
            aMrpPtr^.sQty := ExcelApp.Cells[irow, iQty].Value;
            aMrpPtr^.sErpClass := ExcelApp.Cells[irow, iErpClass].Value;
            aMrpPtr^.sNeedType := ExcelApp.Cells[irow, iNeedType].Value;
            aMrpPtr^.sSignStatus := ExcelApp.Cells[irow, iSignStatus].Value;
            aMrpPtr^.sProj := ExcelApp.Cells[irow, iProj].Value;
            aMrpPtr^.sPlanOrderNo := ExcelApp.Cells[irow, iPlanOrderNo].Value;
            aMrpPtr^.sMC := ExcelApp.Cells[irow, iMC].Value;

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, iNumber].Value;
          end;

        end;
      
      finally
        ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook.Close;
      end;

    finally
      ExcelApp.Visible := True;
      ExcelApp.Quit;
    end;



 
    for ifile := 0 to lstManFile.Items.Count - 1 do
    begin

      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';


      try
        WorkBook := ExcelApp.WorkBooks.Open(lstManFile.Items[ifile]);

        try
          iSheetCount := ExcelApp.Sheets.Count;
          for iSheet := 1 to iSheetCount do
          begin
            if not ExcelApp.Sheets[iSheet].Visible then Continue;

            ExcelApp.Sheets[iSheet].Activate;
                              
            sSheet := ExcelApp.Sheets[iSheet].Name;
            if sSheet <> '14DAYS PR Sum' then Continue;


//            iPN_man := IndexOfColAlert(ExcelApp, CSPN_man);
            iNumber99_man := IndexOfColAlert(ExcelApp, CSNumber99_man);
            iOrderDate_man := IndexOfColAlert(ExcelApp, CSOrderDate_man);
            iNetNeed_man := IndexOfColAlert(ExcelApp, CSNetNeed_man);
            iNetNeedMan_man := IndexOfColAlert(ExcelApp, CSNetNeedMan_man);
            iNote_man := IndexOfColAlert(ExcelApp, CSNote_man); 

            irow := 2;
            snumber := ExcelApp.Cells[irow, iNumber99_man].Value;
            while snumber <> '' do
            begin
              aManPtr := New(PPRManRecord);
              lstManPR.Add(aManPtr);

//              aManPtr^.sPN := ExcelApp.Cells[irow, iPN_man].Value;
              aManPtr^.sNnumber99 := ExcelApp.Cells[irow, iNumber99_man].Value;
              aManPtr^.sOrderDate := ExcelApp.Cells[irow, iOrderDate_man].Value;
              aManPtr^.sNetNeed := ExcelApp.Cells[irow, iNetNeed_man].Value;
              aManPtr^.sNetNeedMan := ExcelApp.Cells[irow, iNetNeedMan_man].Value;
              aManPtr^.sNote := ExcelApp.Cells[irow, iNote_man].Value;

              irow := irow + 1;
              snumber := ExcelApp.Cells[irow, iNumber99_man].Value;
            end;
              
          end;
        finally
          ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
          WorkBook.Close;
        end;

      finally
        ExcelApp.Visible := True;
        ExcelApp.Quit;
      end;
 
    end;


    // 合并 ////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    lst := TList.Create;
    try

      for i := 0 to lstMrpPR.Count - 1 do
      begin
        aMrpPtr := PPRRecord(lstMrpPR[i]);
        GetManPr(lstManPR, aMrpPtr, lst);
        if lst.Count = 0 then Continue;

        dqty := 0;
        for iMan := 0 to lst.Count - 1 do
        begin     
          aManPtr := PPRManRecord(lst[iMan]);
          if aManPtr^.sNetNeedMan = '' then Continue;
          dqty := dqty + StrToFloat(aManPtr^.sNetNeedMan);
        end;

        aMrpPtr^.sQty := FloatToStr(dqty);

        lst.Clear;
      end;
    finally
      lst.Free;
    end;

//    for iMan := 0 to lstManPR.Count - 1 do
//    begin
//      aManPtr := PPRManRecord(lstManPR[iMan]);
//      idx := IndexOfMrpPR(lstMrpPR, aManPtr);
//      if idx < 0 then Continue;
//
//      aMrpPtr := PPRRecord(lstMrpPR[idx]);
//      aMrpPtr^.sQty := aManPtr^.sNetNeedMan;
//    end;

    // 写文件 //////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


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
//      ExcelApp.Cells[irow, iPN].Value := CSPN;
      ExcelApp.Cells[irow, iSel].Value := CSSel;
      ExcelApp.Cells[irow, iItemID].Value := CSItemID;       
      ExcelApp.Cells[irow, iNumber].Value := CSNumber;
      ExcelApp.Cells[irow, iName].Value := CSName;
      ExcelApp.Cells[irow, iNeedDate].Value := CSNeedDate;
      ExcelApp.Cells[irow, iArriveDate].Value := CSArriveDate;
      ExcelApp.Cells[irow, iOrderDate].Value := CSOrderDate;
      ExcelApp.Cells[irow, iAdviceOrderDate].Value := CSAdviceOrderDate;
      ExcelApp.Cells[irow, iUnit].Value := CSUnit;
      ExcelApp.Cells[irow, iRoughNeed].Value := CSRoughNeed;
      ExcelApp.Cells[irow, iOPO].Value := CSOPO;
      ExcelApp.Cells[irow, iStock].Value := CSStock;
      ExcelApp.Cells[irow, iStockAvailble].Value := CSStockAvailble;
      ExcelApp.Cells[irow, iNetNeed].Value := CSNetNeed;
      ExcelApp.Cells[irow, iSPQ].Value := CSSPQ;
      ExcelApp.Cells[irow, iMOQ].Value := CSMOQ;
      ExcelApp.Cells[irow, iLT].Value := CSLT;
      ExcelApp.Cells[irow, iPlanQty].Value := CSPlanQty;
      ExcelApp.Cells[irow, iOldQty].Value := CSOldQty;
      ExcelApp.Cells[irow, iQty].Value := CSQty;
      ExcelApp.Cells[irow, iErpClass].Value := CSErpClass;
      ExcelApp.Cells[irow, iNeedType].Value := CSNeedType;
      ExcelApp.Cells[irow, iSignStatus].Value := CSSignStatus;
      ExcelApp.Cells[irow, iProj].Value := CSProj;
      ExcelApp.Cells[irow, iPlanOrderNo].Value := CSPlanOrderNo;
      ExcelApp.Cells[irow, iMC].Value := CSMC;

      irow := 2;

      ProgressBar1.Max := lstMrpPR.Count;
      ProgressBar1.Position := 0;

      for i := 0 to lstMrpPR.Count - 1 do
      begin
        aMrpPtr := PPRRecord(lstMrpPR[i]);

//        ExcelApp.Cells[irow, iPN].Value := aMrpPtr^.sPN;
        ExcelApp.Cells[irow, iSel].Value := aMrpPtr^.sSel;
        ExcelApp.Cells[irow, iItemID].Value := aMrpPtr^.sItemID;   
        ExcelApp.Cells[irow, iNumber].Value := aMrpPtr^.sNumber;
        ExcelApp.Cells[irow, iName].Value := aMrpPtr^.sName;
        ExcelApp.Cells[irow, iNeedDate].Value := aMrpPtr^.sNeedDate;
        ExcelApp.Cells[irow, iArriveDate].Value := aMrpPtr^.sArriveDate;
        ExcelApp.Cells[irow, iOrderDate].Value := aMrpPtr^.sOrderDate;
        ExcelApp.Cells[irow, iAdviceOrderDate].Value := aMrpPtr^.sAdviceOrderDate;
        ExcelApp.Cells[irow, iUnit].Value := aMrpPtr^.sUnit;
        ExcelApp.Cells[irow, iRoughNeed].Value := aMrpPtr^.sRoughNeed;
        ExcelApp.Cells[irow, iOPO].Value := aMrpPtr^.sOPO;
        ExcelApp.Cells[irow, iStock].Value := aMrpPtr^.sStock;
        ExcelApp.Cells[irow, iStockAvailble].Value := aMrpPtr^.sStockAvailble;
        ExcelApp.Cells[irow, iNetNeed].Value := aMrpPtr^.sNetNeed;
        ExcelApp.Cells[irow, iSPQ].Value := aMrpPtr^.sSPQ;
        ExcelApp.Cells[irow, iMOQ].Value := aMrpPtr^.sMOQ;
        ExcelApp.Cells[irow, iLT].Value := aMrpPtr^.sLT;
        ExcelApp.Cells[irow, iPlanQty].Value := aMrpPtr^.sPlanQty;
        ExcelApp.Cells[irow, iOldQty].Value := aMrpPtr^.sOldQty;
        ExcelApp.Cells[irow, iQty].Value := aMrpPtr^.sQty;
        ExcelApp.Cells[irow, iErpClass].Value := aMrpPtr^.sErpClass;
        ExcelApp.Cells[irow, iNeedType].Value := aMrpPtr^.sNeedType;
        ExcelApp.Cells[irow, iSignStatus].Value := aMrpPtr^.sSignStatus;
        ExcelApp.Cells[irow, iProj].Value := aMrpPtr^.sProj;
        ExcelApp.Cells[irow, iPlanOrderNo].Value := aMrpPtr^.sPlanOrderNo;
        ExcelApp.Cells[irow, iMC].Value := aMrpPtr^.sMC;
        
        irow := irow + 1;

        ProgressBar1.Position := ProgressBar1.Position + 1;
      end;

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, iMC] ].Interior.Color := $DBDCF2;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, iMC] ].HorizontalAlignment := xlCenter;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, iMC] ].Borders.LineStyle := 1; //加边框

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

  finally
    for i := 0 to lstMrpPR.Count - 1 do
    begin
      aMrpPtr := PPRRecord(lstMrpPR[i]);
      Dispose(aMrpPtr);
    end;
    lstMrpPR.Free;

    for i := 0 to lstManPR.Count - 1 do
    begin
      aManPtr := PPRManRecord(lstManPR[i]);
      Dispose(aManPtr);
    end;
    lstManPR.Free;
  end;


  MessageBox(Handle, '完成', '提示', 0);

 
end;

procedure TfrmMrpPRMerge.lstManFileDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := True;
end;

procedure TfrmMrpPRMerge.lstManFileDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  idx: Integer;
  iItemIndex: Integer;
begin
  idx := lstManFile.ItemAtPos(Point(X, Y), True);
  if idx < 0 then Exit;
  if idx = lstManFile.ItemIndex then Exit;  
  Memo1.Lines.Add('from:' + lstManFile.Items[lstManFile.ItemIndex]);

  iItemIndex := lstManFile.ItemIndex;
  
  lstManFile.Items.Insert(idx, lstManFile.Items[iItemIndex]);
  if iItemIndex > idx then
  begin
    iItemIndex := iItemIndex + 1;
  end;  
  lstManFile.Items.Delete(iItemIndex);

  Memo1.Lines.Add('to  :' + lstManFile.Items[idx]);
end;

procedure TfrmMrpPRMerge.FormCreate(Sender: TObject);
begin
//    lstManFile.Items.Add('C:\Users\qiujinbo\Desktop\0411\MRP.PR WK15---夏丽娟(原文件).xlsx');
//    lstManFile.Items.Add('C:\Users\qiujinbo\Desktop\0411\MRP.PR WK15---杨忠祥(原文件).xlsx');
//
//    leMrpPR.Text := 'C:\Users\qiujinbo\Desktop\0411\Finally All PR List 20170410(结果文件).xlsx';
//    lstManFile.Items.Add('sfile113');
//    lstManFile.Items.Add('sfile114');
end;

end.
