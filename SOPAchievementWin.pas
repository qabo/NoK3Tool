unit SOPAchievementWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommUtils,
  IniFiles;

type
  TfrmSOPAchievement = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    Memo1: TMemo;
    ProgressBar1: TProgressBar;
    leSOP: TLabeledEdit;
    btnSOP: TButton;
    leAchievement: TLabeledEdit;
    btnAchievement: TButton;
    leWeek: TLabeledEdit;
    leYear: TLabeledEdit;
    GroupBox1: TGroupBox;
    lstSchFile: TListBox;
    btnAdd: TButton;
    btnDel: TButton;
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSOPClick(Sender: TObject);
    procedure btnAchievementClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure lstSchFileDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure lstSchFileDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

const
  CSPN = 'PN';
  CSSel = 'ѡ��';
  CSItemID = '����';
  CSNumber = '���ϱ���';
  CSName = '����';
  CSNeedDate = '��������';
  CSArriveDate = '��������';
  CSOrderDate = '�µ�����';
  CSAdviceOrderDate = '�����µ�����';
  CSUnit = '��λ';
  CSRoughNeed = 'ë������';
  CSOPO = '����δ�ᶩ��';
  CSStock = '�����';
  CSStockAvailble = '���ÿ��';
  CSNetNeed = '������';
  CSSPQ = 'SPQ';
  CSMOQ =	'MOQ';
  CSLT = 'L/T';
  CSPlanQty = '�ƻ�������';
  CSOldQty = 'old����';
  CSQty = '����';
  CSErpClass = '��������';
  CSNeedType = '��������';
  CSSignStatus = 'ǩ��״̬';
  CSProj = '��Ŀ';
  CSPlanOrderNo = '�ƻ�������';
  CSMC = 'MC';

  CSPN_man = 'PN';
  CSNumber99_man = '���ϱ���99';
  CSOrderDate_man = '�µ�����';
  CSNetNeed_man = '������';
  CSNetNeedMan_man = '������������';
  CSSignStatus_man = '';
  CSNote_man = '��ע';



type
  TSchRecord = packed record
    dt: TDateTime;
    qty: Double;
  end;
  PSchRecord = ^TSchRecord;

  TRowRecord = packed record
    sdate: string;
    sbillno: string;
    sstock: string;  
    snumber: string;
    sname: string;
    dqty: Double;
    snote: string;
    sorderno: string;
  end;
  PRowRecord = ^TRowRecord;

  TPRRecord = packed record
    sPN: string;
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
    sPN: string;
    sNnumber99: string;
    sOrderDate: string;
    sNetNeed: string;
    sNetNeedMan: string;
    sSignStatus: string;
    sNote: string; 
  end;
  PPRManRecord = ^TPRManRecord;

  TProjSchs = class

  end;

class procedure TfrmSOPAchievement.ShowForm;
var
  frmHWCinSum: TfrmSOPAchievement;
begin
  frmHWCinSum := TfrmSOPAchievement.Create(nil);
  try
    frmHWCinSum.ShowModal;
  finally
    frmHWCinSum.Free;
  end;
end;

procedure TfrmSOPAchievement.btnExitClick(Sender: TObject);
begin
  Close;
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
    raise Exception.Create(s + ' �в�����');
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


function GetSchAct(slNumber_sch: TStringList; dt1, dt2: TDateTime; const snumber: string): Double;
var
  i: Integer;
  lst: TList;
  idate: Integer;     
  aSchRecPtr: PSchRecord;
begin
  Result := 0;
  for i := 0 to slNumber_sch.Count - 1 do
  begin
    if slNumber_sch[i] = snumber then
    begin
      lst := TList(slNumber_sch.Objects[i]);
      for idate := 0 to lst.Count - 1 do
      begin
        aSchRecPtr := PSchRecord(lst[idate]);
        if (aSchRecPtr^.dt >= dt1) and (aSchRecPtr^.dt <= dt2) then
        begin
          Result := Result + aSchRecPtr^.qty;
        end;
      end;
    end;
  end;
end;


procedure TfrmSOPAchievement.btnSave2Click(Sender: TObject);
const
  CSTitle = '���ڵ��ݱ�����ϲֿ�ӹ����ϳ�����ӹ���������ʵ��������ע��������';
var
  ExcelApp, WorkBook: Variant;       
  ExcelApp2, WorkBook2: Variant;
  sSheet: string;
  sSheet2: string;
  iSheet: Integer;
  iSheet2: Integer;
  iSheetCount: Integer;
  stitle, stitle1, stitle2, stitle3, stitle4, stitle5: string;
  irow, icol: Integer;
  icolWeek: Integer;
  icolSOP: Integer;
  snumber: string; 
  i: Integer;
  v: Variant;
  s: string;
  dQty: Double;
  sdate, sdate1, sdate2: string;
  dt1, dt2: TDateTime;

  slSOPProj: TStringList;
  slSOP: TStringList;
  idx: Integer;
  isch: Integer;
  sfile_sch: string;
  icolDate1, icolDate2: Integer;
  aProjSchs: TProjSchs;
  slNumber_sch: TStringList;
  aSchRecPtr: PSchRecord;
  lst: TList; 
begin
  sdate := leWeek.Text;
  sdate1 := Copy(sdate, 1, Pos('-', sdate) - 1);
  sdate2 := Copy(sdate, Pos('-', sdate) + 1, Length(sdate));
  sdate1 := leYear.Text + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
  sdate2 := leYear.Text + '-' + StringReplace(sdate2, '/', '-', [rfReplaceAll]);
  dt1 := myStrToDateTime(sdate1);
  dt2 := myStrToDateTime(sdate2);






  //////////////////////////////////////////////////////////////////////
  //////////////////////////////////////////////////////////////////////

  //  �� S&OP �ƻ�

               
  slSOPProj := TStringList.Create;
  aProjSchs := TProjSchs.Create;
  slNumber_sch := TStringList.Create;

  try



    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';


    try
      WorkBook2 := ExcelApp2.WorkBooks.Open(leSOP.Text);

      try
        iSheetCount := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet2 := ExcelApp2.Sheets[iSheet2].Name;
          if Pos(' ', sSheet2) > 0 then
            sSheet2 := Copy(sSheet2, 1, Pos(' ', sSheet2) - 1);
          sSheet2 := UpperCase(sSheet2);

          stitle1 := ExcelApp2.Cells[1, 1].Value;
          stitle2 := ExcelApp2.Cells[1, 2].Value;
          stitle3 := ExcelApp2.Cells[1, 3].Value;
          stitle4 := ExcelApp2.Cells[1, 4].Value;

          stitle := stitle1 + stitle2 + stitle3 + stitle4;
        
          if stitle <> '��ʽ���ϱ�����ɫ����' then
          begin
            Memo1.Lines.Add(sSheet2 + ' ��ʽ������');
            Continue;  
          end;

          icolSOP := 0;
          icol := 5;
          stitle1 := ExcelApp2.Cells[1, icol].Value;
          stitle2 := ExcelApp2.Cells[2, icol].Value;
          while stitle1 + stitle2 <> '' do
          begin
            if stitle2 = leWeek.Text then
            begin
              icolSOP := icol;
              Break;
            end;

            icol := icol + 1;    
            stitle1 := ExcelApp2.Cells[1, icol].Value;
            stitle2 := ExcelApp2.Cells[2, icol].Value;
          end;

          if icolSOP = 0 then
          begin
            Memo1.Lines.Add(sSheet2 + ' �Ҳ��� week ' + leWeek.Text);
            Continue;   // �Ҳ���week�� Ҳû��Ҫ������
          end;
                        

          Memo1.Lines.Add(sSheet2 + ' �ҵ� week ��  irow: ' + IntToStr(2) + '  icolWeek: ' + GetRef(icolSOP));

          idx := slSOPProj.IndexOf(sSheet2);
          if idx >= 0 then
          begin
            slSOP := TStringList(slSOPProj.Objects[idx]);
          end
          else
          begin
            slSOP := TStringList.Create;
            slSOPProj.AddObject(sSheet2, slSOP);
          end;

               
          irow := 3;
          snumber := ExcelApp2.Cells[irow, 2].Value;
          while snumber <> '' do
          begin
            if IsCellMerged(ExcelApp2, irow, 2, irow, 3) then Break;

            v := ExcelApp2.Cells[irow, icolSop].Value;
            s := v;
            if (s <> '') and  not VarIsNumeric(v) then
            begin
              MessageBox(Handle, PChar('S&OP  ' + sSheet2 + ' ��Ԫ��' + IntToStr(irow) + GetRef(icolSOP) + '������Ч����'), '��ʾ', 0);
              Break;
            end;
                  
            dqty := v;

            // ���ڶ���������� S&OP �� ����
            if slSOP.IndexOfName(snumber) >= 0 then
            begin
              slSOP.Values[snumber] := FloatToStr(StrToFloat(slSOP.Values[snumber]) + dqty);
            end
            else
            begin
              slSOP.Add(snumber + '=' + FloatToStr(dqty));
            end;  

            //ExcelApp.Cells[irow, icolWeek].Value := ExcelApp.Cells[irow, icolWeek].Value + dQty;
            irow := irow + 1;
            snumber := ExcelApp2.Cells[irow, 2].Value;

            Memo1.Lines.Add(IntToStr(irow));
          end;

        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;

          
    //////////////////////////////////////////////////////////////////////
    //////////////////////////////////////////////////////////////////////


    for isch := 0 to lstSchFile.Items.Count - 1 do
    begin
      sfile_sch := lstSchFile.Items[isch];


      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';


      try
        WorkBook := ExcelApp.WorkBooks.Open(sfile_sch);

        try
          iSheetCount := ExcelApp.Sheets.Count;
          for iSheet := 1 to iSheetCount do
          begin
            if not ExcelApp.Sheets[iSheet].Visible then Continue;

            ExcelApp.Sheets[iSheet].Activate;

            sSheet := ExcelApp.Sheets[iSheet].Name;
 
            if Copy(sSheet, 1, 3) <> 'CTB' then
            begin       
              Memo1.Lines.Add(sSheet + ' �������Ų�sheet������ CTB- ��ͷ');
              Continue;
            end;


            stitle1 := ExcelApp.Cells[2, 1].Value;
            stitle2 := ExcelApp.Cells[2, 2].Value;
            stitle3 := ExcelApp.Cells[2, 3].Value;
            stitle4 := ExcelApp.Cells[2, 4].Value;
            stitle5 := ExcelApp.Cells[2, 5].Value;

            stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
        
            if stitle <> '�����ͺ����ϱ�����ɫ������Ŀ' then
            begin
              Memo1.Lines.Add(sSheet + ' ��ʽ������');
              Continue;
            end;

            icolDate1 := 7;
            icolDate2 := 7;
            v := ExcelApp.Cells[3, icolDate2].Value;
            while VarIsType(v, varDate) do
            begin
              icolDate2 := icolDate2 + 1;                
              v := ExcelApp.Cells[3, icolDate2].Value;
            end;
            icolDate2 := icolDate2 - 1;

            irow := 4;
            stitle1 := ExcelApp.Cells[irow, 5].Value;
            stitle2 := ExcelApp.Cells[irow + 1, 5].Value;
            stitle := stitle1 + stitle2;
            while stitle = '�ƻ�ʵ��' do
            begin
              if not IsCellMerged(ExcelApp, irow, 2, irow, 3) then
              begin                                     
                snumber := ExcelApp.Cells[irow, 2].Value;
                lst := TList.Create;
                slNumber_sch.AddObject(snumber, lst);

                for icol := icolDate1 to icolDate2 do
                begin      
                  aSchRecPtr := New(PSchRecord);
                  lst.Add(aSchRecPtr);
                  aSchRecPtr^.dt := ExcelApp.Cells[3, icol].Value;
                  v := ExcelApp.Cells[irow + 1, icol].Value;
                  if VarIsNumeric(v) then
                    aSchRecPtr^.qty := v
                  else
                    aSchRecPtr^.qty := 0;
                end;                                                  
              end;  
              irow := irow + 2;
              stitle1 := ExcelApp.Cells[irow, 5].Value;
              stitle2 := ExcelApp.Cells[irow + 1, 5].Value;
              stitle := stitle1 + stitle2;
            end;
 
          end;      
        finally
          ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
          WorkBook.Close;
        end;

      finally
        ExcelApp.Visible := True;
        ExcelApp.Quit;
      end; 


    end;




    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';


    try
      WorkBook := ExcelApp.WorkBooks.Open(leAchievement.Text);

      try
        iSheetCount := ExcelApp.Sheets.Count;
        for iSheet := 1 to iSheetCount do
        begin
          if not ExcelApp.Sheets[iSheet].Visible then Continue;

          ExcelApp.Sheets[iSheet].Activate;

          sSheet := ExcelApp.Sheets[iSheet].Name;

          if Pos(' ', sSheet) > 0 then
            sSheet := Copy(sSheet, 1, Pos(' ', sSheet) - 1);
          sSheet := UpperCase(sSheet);

 
          idx := slSOPProj.IndexOf(sSheet);
          if idx < 0 then
          begin
            Memo1.Lines.Add(sSheet + ' û�� SOP');
            Continue;
          end;

          slSOP := TStringList(slSOPProj.Objects[idx]);

          stitle1 := ExcelApp.Cells[1, 1].Value;
          stitle2 := ExcelApp.Cells[1, 2].Value;
          stitle3 := ExcelApp.Cells[1, 3].Value;
          stitle4 := ExcelApp.Cells[1, 4].Value;

          stitle := stitle1 + stitle2 + stitle3 + stitle4;
        
          if stitle <> '��ʽ���ϱ�����ɫ����' then
          begin
            Memo1.Lines.Add(sSheet + ' ��ʽ������');
            Continue;
          end;

          icolWeek := 0;
          for icol := 5 to 500 do
          begin
            stitle1 := ExcelApp.Cells[2, icol].Value;
            stitle2 := ExcelApp.Cells[2, icol + 1].Value;
            stitle3 := ExcelApp.Cells[2, icol + 2].Value;

            stitle := stitle1 + stitle2 + stitle3;
            if stitle = '' then
            begin
              Break;
            end;

            if stitle = 'S&OP��Ӧ�ƻ�ʵ���������' then
            begin
              stitle := ExcelApp.Cells[1, icol].Value;
              stitle := Copy(stitle, Pos('(', stitle) + 1, Length(stitle));
              stitle := Copy(stitle, 1, Pos(')', stitle) - 1);
              if stitle = leWeek.Text then
              begin
                icolWeek := icol;
                Break;
              end;
            end;
          end;

          if icolWeek = 0 then
          begin
            Memo1.Lines.Add(sSheet + ' �Ҳ��� week ' + leWeek.Text);
            Continue;
          end;
        
                       
          Memo1.Lines.Add(sSheet + ' �ҵ� week ��  irow: ' + IntToStr(2) + '  icolWeek: ' + GetRef(icolWeek));
 
          irow := 3;
          snumber := ExcelApp.Cells[irow, 2].Value;
          while snumber <> '' do
          begin

            ExcelApp.Cells[irow, icolWeek].Value := slSOP.Values[snumber];  
            ExcelApp.Cells[irow, icolWeek + 1].Value := GetSchAct(slNumber_sch, dt1, dt2, snumber);
            ExcelApp.Cells[irow, icolWeek + 2].Value := '=IF(' + GetRef(icolWeek) + IntToStr(irow) + '=0,1,' + GetRef(icolWeek + 1) + IntToStr(irow) + '/' + GetRef(icolWeek) + IntToStr(irow) + ')';

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 2].Value;
          end;
 

        end;     

        WorkBook.Save;
      finally
        ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
        WorkBook.Close;
      end;

    finally
      ExcelApp.Visible := True;
      ExcelApp.Quit;
    end; 

  finally 
    for i := 0 to slSOPProj.Count - 1 do
    begin
      slSOP := TStringList(slSOPProj.Objects[i]);
      slSOP.Free;
    end;
    slSOPProj.Free;

    aProjSchs.Free;

    for i:= 0 to slNumber_sch.Count - 1 do
    begin
      lst := TList(slNumber_sch.Objects[i]);
      for icol := 0 to lst.Count - 1 do
      begin
        aSchRecPtr := PSchRecord(lst[icol]);
        Dispose(aSchRecPtr);
      end;
      lst.Free;
    end;
    slNumber_sch.Free;
  end;

  MessageBox(Handle, '���', '��ʾ', 0);
 
end;

procedure TfrmSOPAchievement.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  sl: TStringList;
  i: Integer;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  leYear.Text := ini.ReadString('history', leYear.Name, '2017');    
  leSOP.Text := ini.ReadString('history', leSOP.Name, '');
  leAchievement.Text := ini.ReadString('history', leAchievement.Name, '');

  sl := TStringList.Create;
  try
    ini.ReadSectionValues('schedule', sl);
    for i := 0 to sl.Count - 1 do
    begin
      lstSchFile.Items.Add(sl.ValueFromIndex[i]);
    end;
  finally
    sl.Free;
  end;

  ini.Free;
  
end;
    
procedure TfrmSOPAchievement.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  i: Integer;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  ini.WriteString('history', leYear.Name, leYear.Text);     
  ini.WriteString('history', leSOP.Name, leSOP.Text);
  ini.WriteString('history', leAchievement.Name, leAchievement.Text);
  ini.EraseSection('schedule');
  for i := 0 to lstSchFile.Items.Count - 1 do
  begin
    ini.WriteString('schedule', IntToStr(i), lstSchFile.Items[i]);
  end;
  ini.Free;
end;

procedure TfrmSOPAchievement.btnSOPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOP.Text := sfile;
end;

procedure TfrmSOPAchievement.btnAchievementClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leAchievement.Text := sfile;
end;

procedure TfrmSOPAchievement.btnAddClick(Sender: TObject);
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
      lstSchFile.Items.Add(sl[i]);
    end;
  finally
    sl.Free;
  end;
end;

procedure TfrmSOPAchievement.btnDelClick(Sender: TObject);
begin
  if lstSchFile.SelCount = 0 then Exit;
  if MessageBox(Handle, 'ȷ��ɾ��ѡ����Ŀ��', '��ʾ', MB_YESNO) <> MrYes then
  begin
    Exit;
  end;
  lstSchFile.DeleteSelected;
end;

procedure TfrmSOPAchievement.lstSchFileDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  idx: Integer;
  iItemIndex: Integer;
begin
  idx := lstSchFile.ItemAtPos(Point(X, Y), True);
  if idx < 0 then Exit;
  if idx = lstSchFile.ItemIndex then Exit;  
  Memo1.Lines.Add('from:' + lstSchFile.Items[lstSchFile.ItemIndex]);

  iItemIndex := lstSchFile.ItemIndex;
  
  lstSchFile.Items.Insert(idx, lstSchFile.Items[iItemIndex]);
  if iItemIndex > idx then
  begin
    iItemIndex := iItemIndex + 1;
  end;  
  lstSchFile.Items.Delete(iItemIndex);

  Memo1.Lines.Add('to  :' + lstSchFile.Items[idx]);
end;

procedure TfrmSOPAchievement.lstSchFileDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := True;
end;

end.
