unit FacAccountCheckWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SAPMB51Reader2, StdCtrls, ComObj, CommUtils, SAPCMSPushErrorReader2,
  SAPDailyAccountReader2, StockMZ2FacReader, Grids, ValEdit, Menus, IniFiles,
  ComCtrls, ToolWin, ImgList, ExtCtrls, Clipbrd, ICMO2FacReader;

type
  TfrmFacAccountCheck = class(TForm)
    Memo1: TMemo;
    vle_ml: TValueListEditor;
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave: TToolButton;
    ToolButton3: TToolButton;
    btnExit: TToolButton;
    tbOpen: TToolButton;
    ToolButton2: TToolButton;
    pm_ml: TPopupMenu;
    mmiWinB: TMenuItem;
    mmiwinR: TMenuItem;
    mmiCPIN: TMenuItem;
    rgFac: TRadioGroup;
    mmiDB: TMenuItem;
    mmiQin: TMenuItem;
    mmiDB_out: TMenuItem;
    mmiPPBom: TMenuItem;
    mmi03to01: TMenuItem;
    mmiA2B: TMenuItem;
    mmiSOut: TMenuItem;
    mmiDB_in: TMenuItem;
    mmiQout: TMenuItem;
    leStockFac2MZ: TLabeledEdit;
    leMB51: TLabeledEdit;
    btnStockFac2MZ: TButton;
    btnMB51: TButton;
    mmiSQ01PPBom: TMenuItem;
    mmiRTV: TMenuItem;
    mmiICMO2fac: TMenuItem;
    pm_wt: TPopupMenu;
    mmiWinB_wt: TMenuItem;
    mmiWinR_wt: TMenuItem;
    mmiCPIN_wt: TMenuItem;
    mmiQin_wt: TMenuItem;
    mmiA2B_wt: TMenuItem;
    mmiQout_wt: TMenuItem;
    mmiDB_wt: TMenuItem;
    mmiSOut_wt: TMenuItem;
    mmiPPBom_wt: TMenuItem;
    MenuItem17: TMenuItem;
    pm_yd: TPopupMenu;
    mmiWinB_yd: TMenuItem;
    mmiWinR_yd: TMenuItem;
    mmiCPIN_yd: TMenuItem;
    mmiQin_yd: TMenuItem;
    mmiA2B_yd: TMenuItem;
    mmiDB_yd: TMenuItem;
    mmiSOut_yd: TMenuItem;
    mmiPPBom_yd: TMenuItem;
    MenuItem10: TMenuItem;
    mmiPPBomChange_yd: TMenuItem;
    mmiDB_out_wt: TMenuItem;
    mmiDB_in_wt: TMenuItem;
    leICMO2Fac: TLabeledEdit;
    Button1: TButton;
    mmiDB_in_out_yd: TMenuItem;
    mmiSQ01PPBomChange_yd: TMenuItem;
    mmiWINDB: TMenuItem;
    procedure btnExitClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure mmiWinBClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure rgFacClick(Sender: TObject);
    procedure btnStockFac2MZClick(Sender: TObject);
    procedure btnMB51Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    procedure SaveAndLoadValues(idx1, idx2: Integer);
    procedure btnSaveClick_ml; 
    procedure btnSaveClick_wt;
    procedure btnSaveClick_yd;
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

function myTrim(const S: string): string;
var
  I, L: Integer;
begin
  L := Length(S);
  I := 1;
  while (I <= L) and not (((S[I] >= '0') and (S[I] <= '9')) or ((S[I] >= 'a') and (S[I] < 'z')) or ((S[I] >= 'A') and (S[I] <= 'Z'))) do
    Inc(I);
  if I > L then Result := '' else
  begin
    while not (((S[L] >= '0') and (S[L] <= '9')) or ((S[L] >= 'a') and (S[L] < 'z')) or ((S[L] >= 'A') and (S[L] <= 'Z'))) do Dec(L);

    Result := Copy(S, I, L - I + 1);
  end;
end;

class procedure TfrmFacAccountCheck.ShowForm;
var
  frmFacAccountCheck: TfrmFacAccountCheck;
begin
  frmFacAccountCheck := TfrmFacAccountCheck.Create(nil);
  try
    frmFacAccountCheck.ShowModal;
  finally
    frmFacAccountCheck.Free;
  end;
end;
       
procedure TfrmFacAccountCheck.btnExitClick(Sender: TObject);
begin
  Close;
end;
    
procedure TfrmFacAccountCheck.FormCreate(Sender: TObject); 
begin
  vle_ml.Strings.Clear;

  SaveAndLoadValues(-1, rgFac.ItemIndex);
end;
       
procedure TfrmFacAccountCheck.FormDestroy(Sender: TObject); 
begin
  SaveAndLoadValues(rgFac.ItemIndex, -1);
end;
   
procedure TfrmFacAccountCheck.SaveAndLoadValues(idx1, idx2: Integer);
var
  i: Integer;   
  ini: TIniFile;
  s: string;
  sl: TStringList;
begin
  rgFac.Tag := idx2;

  if idx1 <> -1 then
  begin
    ini := TIniFile.Create(AppIni);
    try
      s := StringReplace(vle_ml.Strings.Text, #13#10, '||', [rfReplaceAll]);
      ini.WriteString(self.ClassName, vle_ml.Name + rgFac.Items[idx1], s);

      ini.WriteString(self.ClassName, leStockFac2MZ.Name + rgFac.Items[idx1], leStockFac2MZ.Text); 
      ini.WriteString(self.ClassName, leMB51.Name + rgFac.Items[idx1], leMB51.Text);
//      ini.WriteString(self.ClassName, leCMSErrMsg.Name + rgFac.Items[idx1], leCMSErrMsg.Text);
      ini.WriteString(self.ClassName, leICMO2Fac.Name + rgFac.Items[idx1], leICMO2Fac.Text);
    finally
      ini.Free;
    end;
  end;

  vle_ml.Strings.Clear;

  case idx2 of
    0, 4:
    begin
      tbOpen.DropdownMenu := pm_ml;
    end;
    1:
    begin
      tbOpen.DropdownMenu := pm_wt;
    end;
    2:
    begin
      tbOpen.DropdownMenu := pm_yd;
    end; 
    else Exit; 
  end;

  for i := 0 to tbOpen.DropdownMenu.Items.Count - 1 do
  begin
    tbOpen.DropdownMenu.Items[i].OnClick := mmiWinBClick;
    vle_ml.Values[tbOpen.DropdownMenu.Items[i].Caption] := '';
  end;
 
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, vle_ml.Name + rgFac.Items[idx2], '');
    sl := TStringList.Create;
    try
      sl.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
      for i := 0 to sl.Count - 1 do
      begin                                                            
        if tbOpen.DropdownMenu.Items.Find(sl.Names[i]) = nil then Continue;
        vle_ml.Values[ sl.Names[i] ] := sl.ValueFromIndex[i];
      end;
    finally
      sl.Free;
    end;
 
    leStockFac2MZ.Text := ini.ReadString(self.ClassName, leStockFac2MZ.Name + rgFac.Items[idx2], '');
    leMB51.Text := ini.ReadString(self.ClassName, leMB51.Name + rgFac.Items[idx2], '');
//    leCMSErrMsg.Text := ini.ReadString(self.ClassName, leCMSErrMsg.Name + rgFac.Items[idx2], '');
    leICMO2Fac.Text := ini.ReadString(self.ClassName, leICMO2Fac.Name + rgFac.Items[idx2], '');
  finally
    ini.Free;
  end;

end;  
    
procedure TfrmFacAccountCheck.btnStockFac2MZClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leStockFac2MZ.Text := sfile;
end;

procedure TfrmFacAccountCheck.btnMB51Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMB51.Text := sfile;
end;

procedure TfrmFacAccountCheck.mmiWinBClick(Sender: TObject);
var
  mi: TMenuItem;
  sfile: string;
  s: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  mi := TMenuItem(Sender);
  s := mi.Caption;
  s := Copy(s, 1, Pos('(', s) - 1);
  vle_ml.Values[s] := sfile;
end;

procedure TfrmFacAccountCheck.rgFacClick(Sender: TObject);
begin
  SaveAndLoadValues(rgFac.Tag, rgFac.ItemIndex);
  vle_ml.TitleCaptions[0] := rgFac.Items[rgFac.ItemIndex];
  vle_ml.Invalidate;

  if rgFac.Items[ rgFac.ItemIndex ] = '����' then
  begin
    tbOpen.DropdownMenu := pm_ml;
  end 
  else if rgFac.Items[rgFac.ItemIndex] = '��̩' then
  begin
    tbOpen.DropdownMenu := pm_wt;
  end
  else if rgFac.Items[rgFac.ItemIndex] = '���' then
  begin
    tbOpen.DropdownMenu := pm_yd;
  end;
end;

procedure TfrmFacAccountCheck.btnSaveClick(Sender: TObject);
begin
  if rgFac.Items[ rgFac.ItemIndex ] = '����' then
  begin
    btnSaveClick_ml;
  end 
  else if rgFac.Items[rgFac.ItemIndex] = '��̩' then
  begin
    btnSaveClick_wt;
  end
  else if rgFac.Items[rgFac.ItemIndex] = '���' then
  begin
    btnSaveClick_yd;
  end;
end;

procedure TfrmFacAccountCheck.btnSaveClick_ml;
const
  CSBoolean: array[Boolean] of string = ('��', '��');
var
  ExcelApp, WorkBook: Variant;
  aSAPMB51Reader2: TSAPMB51Reader2;
  aSAPCMSPushErrorReader2: TSAPCMSPushErrorReader2;
  iSheet: Integer;
  irow: Integer;
  sfile: string;   
  aStockMZ2FacReader: TStockMZ2FacReader;
  
  aSAPDailyAccountReader2_winB: TSAPDailyAccountReader2_winB;
  aSAPDailyAccountReader2_winR: TSAPDailyAccountReader2_winB;
  aSAPDailyAccountReader2_RTV: TSAPDailyAccountReader2_RTV;
  aSAPDailyAccountReader2_cpin: TSAPDailyAccountReader2_cpin_ml;
  aSAPDailyAccountReader2_qin: TSAPDailyAccountReader2_qin_ml;
  aSAPDailyAccountReader2_a2b: TSAPDailyAccountReader2_qout_ml;
  aSAPDailyAccountReader2_03to01: TSAPDailyAccountReader2_03to01_ml;
  aSAPDailyAccountReader2_qout: TSAPDailyAccountReader2_qout_ml;
  aSAPDailyAccountReader2_DB: TSAPDailyAccountReader2_DB_ml;
  aSAPDailyAccountReader2_DB_out: TSAPDailyAccountReader2_DB_out_ml;
  aSAPDailyAccountReader2_DB_in: TSAPDailyAccountReader2_DB_in_ml; 
  aSAPDailyAccountReader2_sout: TSAPDailyAccountReader2_sout_ml;
  aSAPDailyAccountReader2_xout: TSAPDailyAccountReader2_xout_ml;

  aSAPDailyAccountReader2_coois: TSAPDailyAccountReader2_coois;
  //aSAPDailyAccountReader2_icmo_mz2fac: TSAPDailyAccountReader2_icmo_mz2fac;
  aSAPDailyAccountReader2_PPBom: TSAPDailyAccountReader2_PPBOM;


  aSAPDailyAccountReader2_winB_DB: TSAPDailyAccountReader2_winB;

  i_fac: Integer;
  i_mz: Integer;
  s_fac, s_fac2: string;
  s_mz, s_mz2: string;
  bFound: Boolean;

  iCountWinB_Fac: Integer;
  iCountWinB_DB_Fac: Integer;
  iCountWinR_Fac: Integer;
  iCountCPIN_Fac: Integer;
  iCountQIn_Fac: Integer;
  iCountA2B_Fac: Integer;     //�Ϻŵ���
  iCount03to01_Fac: Integer;  //�������ɢ��  
  iCountQout_Fac: Integer;    //���ϳ���
  iCountDB_Fac: Integer;
  iCountDB_in_Fac: Integer;
  iCountDB_out_Fac: Integer;
  iCountSout_Fac: Integer;
  iCountPPBom: Integer;


  iCountMatch_WinB: Integer;
  iCountMatch_WinB_DB: Integer;
  iCountMatch_WinR: Integer;
  iCountMatch_cpin: Integer;
  iCountMatch_qin: Integer;
  iCountMatch_A2B: Integer;
  iCountMatch_03to01: Integer;
  iCountMatch_qout: Integer;
  iCountMatch_DB: Integer;
  iCountMatch_DB_out: Integer;
  iCountMatch_DB_in: Integer;
  iCountMatch_Sout: Integer;
  iCountMatch_PPBom: Integer;

  aSAPMB51RecordPtr: PSAPMB51Record;
  aDailyAccount_winBPtr: PDailyAccount_winB;
  aDailyAccount_win_MatchBPtr: PDailyAccount_winB;
  aDailyAccount_RTVPtr: PDailyAccount_RTV;
  aDailyAccount_cpinPtr: PDailyAccount_cpin;
  aDailyAccountqinPtr: PDailyAccount_qin;
  aDailyAccountqoutPtr: PDailyAccount_qout;
  aDailyAccount_DBPtr: PDailyAccount_DB;
  aDailyAccount_DB_inPtr: PDailyAccount_DB_in;
  aDailyAccount_DB_outPtr: PDailyAccount_DB_out;
  aDailyAccount_OutAInBCPtr: PDailyAccount_OutAInBC;
  aDailyAccount_soutPtr: PDailyAccount_sout;
  aDailyAccount_xoutPtr: PDailyAccount_xout;
  ptrDailyAccount_PPBOM: PDailyAccount_PPBom;
  ptrDailyAccount_coois: PDailyAccount_coois;
 
  aCPINmz2facReader: TCPINmz2facReader;

  aSAPMB51RecordPtr_match: PSAPMB51Record;
  
  s: string;
  sfile_k3: string;                                                                

  sfile_sq01_ppbom: string;

  sbillno: string;
  idx: Integer;
  dDelta: Double;
  sl: TStringList;
  sline: string;

  dQtyMatchx: Double;
  dQtyMatch0: Double;

//  ptrDailyAccount_COOIS_Head: PDailyAccount_COOIS_Head;
begin
  if not ExcelSaveDialog(sfile) then Exit;
                                                                        
  aSAPMB51Reader2 := TSAPMB51Reader2.Create(leMB51.Text, nil);
  aStockMZ2FacReader := TStockMZ2FacReader_ml.Create(leStockFac2MZ.Text);
//  aSAPCMSPushErrorReader2 := TSAPCMSPushErrorReader2.Create(leCMSErrMsg.Text, nil);

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

    Memo1.Lines.Add('����');

    WorkBook := ExcelApp.WorkBooks.Add;
    ExcelApp.DisplayAlerts := False;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate; 
    ExcelApp.Sheets[iSheet].Name := '����';
                  

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    s := mmiWinB.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
                                                 
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winB := TSAPDailyAccountReader2_winB_ML.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_winB.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := '����';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '����';
        ExcelApp.Cells[irow, 8].Value := '�������';
        ExcelApp.Cells[irow, 9].Value := '��;';
        ExcelApp.Cells[irow, 10].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 11].Value := '��ע';
        ExcelApp.Cells[irow, 12].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 13].Value := 'ժҪ';
        ExcelApp.Cells[irow, 14].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 15].Value := '�رձ�־';
        ExcelApp.Cells[irow, 16].Value := '����';
        ExcelApp.Cells[irow, 17].Value := '���鷽ʽ';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';
        ExcelApp.Cells[irow, 19].Value := 'Դ������';
                         
        AddColor(ExcelApp, irow, 4, irow, 5, clYellow);
        AddColor(ExcelApp, irow, 6, irow, 6, clRed);


        irow := irow + 1;
        iCountWinB_Fac := aSAPDailyAccountReader2_winB.Count;
        iCountMatch_WinB := 0; 
        for i_fac := 0 to aSAPDailyAccountReader2_winB.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winB.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dQty;  

          ExcelApp.Cells[irow, 7].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.dtCheck;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.suse;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.ssupplier;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_winBPtr^.snote;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.sstock;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_winBPtr^.ssummary;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_winBPtr^.sbiller;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sclose;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_winBPtr^.sdept;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_winBPtr^.schecktype;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_winBPtr^.sedi;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_winBPtr^.ssourcebillno;

          s_fac := aDailyAccount_winBPtr^.sbillno +
            aDailyAccount_winBPtr^.snumber +
            aDailyAccount_winBPtr^.snote;   // �ɹ�����
 
          bFound := False;
          aSAPMB51RecordPtr_match := nil;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber +
              aSAPMB51RecordPtr^.sbillno_po;// �ɹ�����

            if s_fac = s_mz then
            begin                                              
              bFound := True;

              dQtyMatchx := aSAPMB51Reader2.GetMB51Qty101(aSAPMB51RecordPtr);
              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
                dQtyMatch0 := dQtyMatchx;
              end
              else
              begin
                if Abs(dQtyMatch0 - aDailyAccount_winBPtr^.dQty ) >
                  Abs(dQtyMatchx - aDailyAccount_winBPtr^.dQty ) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;      
                  dQtyMatch0 := dQtyMatchx;
                end;
              end;
              
              if DoubleE(dQtyMatch0 - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                Break;
              end; 
            end;
          end;    
 
          if bFound then
          begin
            ExcelApp.Cells[irow, 5].Value := dQtyMatch0;
            ExcelApp.Cells[irow, 6].Value := dQtyMatch0 - aDailyAccount_winBPtr^.dQty; 

            aSAPMB51Reader2.SetCalcFlag(aSAPMB51RecordPtr_match, s);

            if DoubleE(dQtyMatch0, aDailyAccount_winBPtr^.dQty) then
            begin
              iCountMatch_WinB := iCountMatch_WinB + 1;
            end;
          end
          else
          begin
            ExcelApp.Cells[irow, 5].Value := '0';
            ExcelApp.Cells[irow, 6].Value := - aDailyAccount_winBPtr^.dQty;
          end;
  
          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winB.Free;
      end;
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    s := mmiWinR.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
 
                                                      
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winR := TSAPDailyAccountReader2_winB_ML.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_winR.Count > 0 then
    begin
      try
    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := 'ʵ������';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '��ע';
        ExcelApp.Cells[irow, 8].Value := '����';
        ExcelApp.Cells[irow, 9].Value := '�������';
        ExcelApp.Cells[irow, 10].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 11].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := 'ժҪ';
        ExcelApp.Cells[irow, 14].Value := '��˱�־';
        ExcelApp.Cells[irow, 15].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 16].Value := 'EDI�ύ';
                 
        AddColor(ExcelApp, irow, 5, irow, 6, clYellow);
        AddColor(ExcelApp, irow, 7, irow, 7, clRed);
 
        irow := irow + 1;
        iCountWinR_Fac := aSAPDailyAccountReader2_winR.Count;
        iCountMatch_WinR := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_winR.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winR.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dQty;  

          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.dtCheck;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.ssupplier;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_winBPtr^.sstock;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_winBPtr^.ssummary;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_winBPtr^.schecktype;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sbiller;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_winBPtr^.sedi;

          s_fac := aDailyAccount_winBPtr^.sbillno +
            aDailyAccount_winBPtr^.snumber
             + aDailyAccount_winBPtr^.snote;

          bFound := False;               
          aSAPMB51RecordPtr_match := nil;
          dQtyMatchx := 0;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber
              + aSAPMB51RecordPtr^.sbillno_po;

            if s_fac = s_mz then
            begin                                              
              bFound := True;


              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else
              begin
                if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty ) >
                  Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_winBPtr^.dQty ) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
                end;
              end;
              
              if DoubleE(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                Break;
              end;

//
//
//              dQtyMatchx := dQtyMatchx + aSAPMB51RecordPtr^.dqty;
//              ExcelApp.Cells[irow, 5].Value := dQtyMatchx;
//              ExcelApp.Cells[irow, 6].Value := dQtyMatchx - aDailyAccount_winBPtr^.dQty;
//              
//              aSAPMB51RecordPtr^.bCalc := True;
//              aSAPMB51RecordPtr^.sMatchType := s;
//              
//              if DoubleE(dQtyMatchx - aDailyAccount_winBPtr^.dQty, 0) then
//              begin
//                iCountMatch_WinR := iCountMatch_WinR + 1;
//                Break;
//              end;
            end;
          end;     

//          if not bFound then
//          begin
//            ExcelApp.Cells[irow, 5].Value := '0';
//            ExcelApp.Cells[irow, 6].Value := aDailyAccount_winBPtr^.dQty;
//          end;


          if bFound then
          begin
            ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 6].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty;
            aSAPMB51RecordPtr^.bCalc := True;
            aSAPMB51RecordPtr^.sMatchType := s;

            if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_winBPtr^.dQty) then
            begin
              iCountMatch_WinB := iCountMatch_WinB + 1;
            end;
          end
          else
          begin
            ExcelApp.Cells[irow, 5].Value := '0';
            ExcelApp.Cells[irow, 6].Value := - aDailyAccount_winBPtr^.dQty;
          end;



          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winR.Free;
      end;
    end;          
                

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    s := mmiWINDB.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
                                                 
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winB_DB := TSAPDailyAccountReader2_winB_ML.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_winB_DB.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := '����';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '����';
        ExcelApp.Cells[irow, 8].Value := '�������';
        ExcelApp.Cells[irow, 9].Value := '��;';
        ExcelApp.Cells[irow, 10].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 11].Value := '��ע';
        ExcelApp.Cells[irow, 12].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 13].Value := 'ժҪ';
        ExcelApp.Cells[irow, 14].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 15].Value := '�رձ�־';
        ExcelApp.Cells[irow, 16].Value := '����';
        ExcelApp.Cells[irow, 17].Value := '���鷽ʽ';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';
        ExcelApp.Cells[irow, 19].Value := 'Դ������';
                         
        AddColor(ExcelApp, irow, 4, irow, 5, clYellow);
        AddColor(ExcelApp, irow, 6, irow, 6, clRed);


        irow := irow + 1;
        iCountWinB_DB_Fac := aSAPDailyAccountReader2_winB_DB.Count;
        iCountMatch_WinB_DB := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_winB_DB.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winB_DB.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dQty;  

          ExcelApp.Cells[irow, 7].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.dtCheck;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.suse;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.ssupplier;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_winBPtr^.snote;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.sstock;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_winBPtr^.ssummary;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_winBPtr^.sbiller;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sclose;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_winBPtr^.sdept;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_winBPtr^.schecktype;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_winBPtr^.sedi;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_winBPtr^.ssourcebillno;

          s_fac := aDailyAccount_winBPtr^.sbillno +
            aDailyAccount_winBPtr^.snumber;

//          s_fac2 := aDailyAccount_winBPtr^.sbillno +
//            aDailyAccount_winBPtr^.snumber +
//            aDailyAccount_winBPtr^.snote;   // �ɹ�����

          dQtyMatchx := MaxInt;
          aSAPMB51RecordPtr_match := nil;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr.bCalc then Continue;

            if aSAPMB51RecordPtr^.smovingtype <> '311' THEN Continue;

            if aSAPMB51RecordPtr^.dqty < 0 then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber + aSAPMB51RecordPtr^.sbillno_po;
                             
//            s_mz2 := sbillno +
//              aSAPMB51RecordPtr^.snumber + aSAPMB51RecordPtr^.sbillno_po +
//              aSAPMB51RecordPtr^.sbillno_po;// �ɹ�����

            if s_fac = s_mz then
            begin
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else
              begin
                if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty ) >
                  Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_winBPtr^.dQty ) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
                end;
              end;
              if DoubleE(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                Break;
              end;
            end;
          end;     

          if bFound then
          begin
            ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 6].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty;
            aSAPMB51RecordPtr^.bCalc := True;
            aSAPMB51RecordPtr^.sMatchType := s;

            if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_winBPtr^.dQty) then
            begin
              iCountMatch_WinB_DB := iCountMatch_WinB_DB + 1;
            end;
          end
          else
          begin
            ExcelApp.Cells[irow, 5].Value := '0';
            ExcelApp.Cells[irow, 6].Value := - aDailyAccount_winBPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winB_DB.Free;
      end;
    end;

         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    
    // ���ϵ�
    s := mmiRTV.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
 
                                                      
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_RTV := TSAPDailyAccountReader2_RTV_ML.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_RTV.Count > 0 then
    begin
      try
    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := 'ʵ������';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '��ע';
        ExcelApp.Cells[irow, 8].Value := '����';
        ExcelApp.Cells[irow, 9].Value := '�������';
        ExcelApp.Cells[irow, 10].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 11].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := 'ժҪ';
        ExcelApp.Cells[irow, 14].Value := '��˱�־';
        ExcelApp.Cells[irow, 15].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 16].Value := 'EDI�ύ';
                 
        AddColor(ExcelApp, irow, 5, irow, 6, clYellow);
        AddColor(ExcelApp, irow, 7, irow, 7, clRed);
 
        irow := irow + 1;
        iCountWinR_Fac := iCountWinR_Fac + aSAPDailyAccountReader2_RTV.Count; 
        for i_fac := 0 to aSAPDailyAccountReader2_RTV.Count - 1 do
        begin
          aDailyAccount_RTVPtr := aSAPDailyAccountReader2_RTV.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_RTVPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_RTVPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_RTVPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_RTVPtr^.dQty;

          ExcelApp.Cells[irow, 8].Value := aDailyAccount_RTVPtr^.dt;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_RTVPtr^.dtCheck;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_RTVPtr^.ssupplier;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_RTVPtr^.sstock;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_RTVPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_RTVPtr^.ssummary;
          ExcelApp.Cells[irow, 14].Value := ''; //aDailyAccount_RTVPtr^.scheckflag;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_RTVPtr^.sbiller;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_RTVPtr^.sedi;
            
          s_fac := aDailyAccount_RTVPtr^.sbillno +
            aDailyAccount_RTVPtr^.snumber
             + aDailyAccount_RTVPtr^.snote;

          bFound := False;
          dQtyMatchx := 0;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber
              + aSAPMB51RecordPtr^.sbillno_po;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              dQtyMatchx := dQtyMatchx + aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 5].Value := dQtyMatchx;
              ExcelApp.Cells[irow, 6].Value := dQtyMatchx - aDailyAccount_RTVPtr^.dQty;
              
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              
              if DoubleE(dQtyMatchx - aDailyAccount_RTVPtr^.dQty, 0) then
              begin
                iCountMatch_WinR := iCountMatch_WinR + 1;
                Break;
              end;
            end;
          end;     

          if not bFound then
          begin
            ExcelApp.Cells[irow, 5].Value := '0';
            ExcelApp.Cells[irow, 6].Value := aDailyAccount_RTVPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_RTV.Free;
      end;
    end;

    
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
     

    s := mmiCPIN.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);


    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_cpin := TSAPDailyAccountReader2_cpin_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_cpin.Count > 0 then
    begin
      s := mmiICMO2fac.Caption;
      if Pos('(', s) > 0 then
      begin
        s := Copy(s, 1, Pos('(', s) - 1);
      end;
      sfile_k3 := vle_ml.Values[s];
      Memo1.Lines.Add(s);
      aCPINmz2facReader := TCPINmz2facReader.Create(sfile_k3);

      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '�������񵥺�';
        ExcelApp.Cells[irow, 2].Value := '����';
        ExcelApp.Cells[irow, 3].Value := '�������';
        ExcelApp.Cells[irow, 4].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 5].Value := '��������';
        ExcelApp.Cells[irow, 6].Value := 'ʵ������';
        ExcelApp.Cells[irow, 7].Value := 'SAP����';
        ExcelApp.Cells[irow, 8].Value := '����';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := '�ջ��ֿ�';
        ExcelApp.Cells[irow, 11].Value := '���ݱ��';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := '������λ';
        ExcelApp.Cells[irow, 14].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 15].Value := '�����';
        ExcelApp.Cells[irow, 16].Value := '��˱�־';
        ExcelApp.Cells[irow, 17].Value := '�����־';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';


        irow := irow + 1;
        iCountCPIN_Fac := aSAPDailyAccountReader2_cpin.Count;
        iCountMatch_CPIN := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_cpin.Count - 1 do
        begin
          aDailyAccount_cpinPtr := aSAPDailyAccountReader2_cpin.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_cpinPtr^.sicmo;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_cpinPtr^.dt;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_cpinPtr^.dtcheck;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_cpinPtr^.snumber;  
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_cpinPtr^.sname;    
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_cpinPtr^.dqty;

          ExcelApp.Cells[irow, 9].Value := aDailyAccount_cpinPtr^.sbatchno;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_cpinPtr^.sstock;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_cpinPtr^.sbillno;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_cpinPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_cpinPtr^.sdept;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_cpinPtr^.sbiller;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_cpinPtr^.schecker;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_cpinPtr^.scheckflag;   
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_cpinPtr^.sbackflush;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_cpinPtr^.sedi;

          s_fac :=  aDailyAccount_cpinPtr.sbillno + aDailyAccount_cpinPtr^.snumber +
            aDailyAccount_cpinPtr^.sstock;
                    
          bFound := False;
          dDelta := 9999999999;
          idx := -1;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr.bCalc then Continue;

            if (aSAPMB51RecordPtr^.smovingtype <> '101') and
              (aSAPMB51RecordPtr^.smovingtype <> '102') then
            begin
              Continue;
            end;                        

            if aSAPMB51RecordPtr^.fstockname = ''  then // Ҫ�вֿ�����
            begin
              Continue;
            end;            
                     
            s_mz := aCPINmz2facReader.cpin_mz2fac(aSAPMB51RecordPtr^.sbillno) +  // ���幤���ţ�ת���ɴ����������� ���жԱ�
              aSAPMB51RecordPtr^.snumber + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True; 
              if dDelta > aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty then
              begin
                dDelta := Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty);
                idx := i_mz;
              end;
              if DoubleE(dDelta, 0) then Break;
            end;
          end;     

          if bFound then
          begin               
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[idx];
            ExcelApp.Cells[irow, 7].Value := aSAPMB51RecordPtr^.dqty;
            ExcelApp.Cells[irow, 8].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty;
            if DoubleE(dDelta, 0) then
            begin
              iCountMatch_CPIN := iCountMatch_CPIN + 1;            
            end;
            aSAPMB51RecordPtr^.bCalc := True;   
            aSAPMB51RecordPtr^.sMatchType := s;
          end
          else
          begin
            ExcelApp.Cells[irow, 7].Value := '0';
            ExcelApp.Cells[irow, 8].Value := - aDailyAccount_cpinPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_cpin.Free;
        aCPINmz2facReader.Free;
      end;
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    // ������ⵥ - Sample                       
    s := mmiQin.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                   
    Memo1.Lines.Add('���ļ��� ' + s);   
    aSAPDailyAccountReader2_qin := TSAPDailyAccountReader2_qin_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_qin.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := '����';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '����';
        ExcelApp.Cells[irow, 8].Value := '�������';
        ExcelApp.Cells[irow, 9].Value := '��;';
        ExcelApp.Cells[irow, 10].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 11].Value := '��ע';
        ExcelApp.Cells[irow, 12].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 13].Value := 'ժҪ';
        ExcelApp.Cells[irow, 14].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 15].Value := '�رձ�־';
        ExcelApp.Cells[irow, 16].Value := '����';
        ExcelApp.Cells[irow, 17].Value := '���鷽ʽ';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';
        ExcelApp.Cells[irow, 19].Value := 'Դ������';    


        irow := irow + 1;
        iCountQIn_Fac := aSAPDailyAccountReader2_qin.Count;
        iCountMatch_Qin := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_qin.Count - 1 do
        begin
          aDailyAccountqinPtr := aSAPDailyAccountReader2_qin.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqinPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccountqinPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqinPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqinPtr^.dqty;
        
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqinPtr^.dt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqinPtr^.dtcheck;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqinPtr^.suse;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqinPtr^.ssupplier;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqinPtr^.snote;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqinPtr^.sstock;
          ExcelApp.Cells[irow, 13].Value := aDailyAccountqinPtr^.ssummary;
          ExcelApp.Cells[irow, 14].Value := aDailyAccountqinPtr^.sbiller;
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqinPtr^.scloseflag;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqinPtr^.sdept;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqinPtr^.schecktype;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqinPtr^.sedit;      
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqinPtr^.ssourcebillno;

          s_fac := aDailyAccountqinPtr^.snumber +
            aDailyAccountqinPtr^.sbillno
            ; //+  aDailyAccountqinPtr^.sstock;

          aSAPMB51RecordPtr_match := nil;
          
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr^.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccountqinPtr^.dQty) >
                Abs(aSAPMB51RecordPtr^.dqty - aDailyAccountqinPtr^.dQty) then
              begin                                                                   
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end;
              if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccountqinPtr^.dQty) then Break; 
            end;
          end;

          if bFound then
          begin
            if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccountqinPtr^.dQty) then
            begin
              iCountMatch_Qin := iCountMatch_Qin + 1;
            end;

            ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 6].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccountqinPtr^.dQty;
 
            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;
          end
          else
          begin
            ExcelApp.Cells[irow, 5].Value := '0';
            ExcelApp.Cells[irow, 6].Value := - aDailyAccountqinPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_qin.Free;
      end;
    end;         
                     



    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                           
    s := mmiA2B.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                    
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_a2b := TSAPDailyAccountReader2_qout_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_a2b.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '��Ʒ������';
        ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 3].Value := '����';
        ExcelApp.Cells[irow, 4].Value := 'SAP����';
        ExcelApp.Cells[irow, 5].Value := '����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '�������';
        ExcelApp.Cells[irow, 8].Value := '�����ֿ�';
        ExcelApp.Cells[irow, 9].Value := '���ϲ���';
        ExcelApp.Cells[irow, 10].Value := '���ݱ��';
        ExcelApp.Cells[irow, 11].Value := '��;1';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 14].Value := '��λ';
        ExcelApp.Cells[irow, 15].Value := '��˱�־';
        ExcelApp.Cells[irow, 16].Value := '�������';
        ExcelApp.Cells[irow, 17].Value := '��;2';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';


        irow := irow + 1;
        iCountA2B_Fac := aSAPDailyAccountReader2_a2b.Count;
        iCountMatch_A2B := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_a2b.Count - 1 do
        begin
          aDailyAccountqoutPtr := aSAPDailyAccountReader2_a2b.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqoutPtr^.snumber;
          ExcelApp.Cells[irow, 2].Value := aDailyAccountqoutPtr^.sname;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqoutPtr^.dqty;
                                                                    
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqoutPtr^.dt;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqoutPtr^.dtcheck;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqoutPtr^.sstock;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqoutPtr^.sdetp;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqoutPtr^.sbillno;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqoutPtr^.suse1;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqoutPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccountqoutPtr^.sbiller;
          ExcelApp.Cells[irow, 14].Value := aDailyAccountqoutPtr^.sunit;
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqoutPtr^.scheckflag;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqoutPtr^.souttype;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqoutPtr^.suse2;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqoutPtr^.sedi;

          s_fac := aDailyAccountqoutPtr^.snumber +
            aDailyAccountqoutPtr^.sbillno;
            ; // + aDailyAccountqoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_A2B := iCountMatch_A2B + 1;
              ExcelApp.Cells[irow, 4].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr^.dqty - aDailyAccountqoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;     
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 4].Value := '0';
            ExcelApp.Cells[irow, 5].Value := - aDailyAccountqoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_a2b.Free;
      end;
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    Memo1.Lines.Add('�������ɢ��');

    s := mmi03to01.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + s);

    aSAPDailyAccountReader2_03to01 := TSAPDailyAccountReader2_03to01_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_03to01.Count > 0 then
    begin
      try    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
      
        ExcelApp.Cells[irow, 1].Value := '��Ʒ������';
        ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 3].Value := '����';
        ExcelApp.Cells[irow, 4].Value := 'SAP����';
        ExcelApp.Cells[irow, 5].Value := '����';
        ExcelApp.Cells[irow, 6].Value := '��ע';
        ExcelApp.Cells[irow, 7].Value := '����';
        ExcelApp.Cells[irow, 8].Value := '�������';
        ExcelApp.Cells[irow, 9].Value := '�����ֿ�';
        ExcelApp.Cells[irow, 10].Value := '���ϲ���';
        ExcelApp.Cells[irow, 11].Value := '���ݱ��';
        ExcelApp.Cells[irow, 12].Value := '��;1';
        ExcelApp.Cells[irow, 13].Value := '��ע';
        ExcelApp.Cells[irow, 14].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 15].Value := '��λ';
        ExcelApp.Cells[irow, 16].Value := '��˱�־';
        ExcelApp.Cells[irow, 17].Value := '�������';
        ExcelApp.Cells[irow, 18].Value := '��;2';
        ExcelApp.Cells[irow, 19].Value := 'EDI�ύ';


        irow := irow + 1;
        iCount03to01_Fac := aSAPDailyAccountReader2_03to01.Count;
        iCountMatch_03to01 := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_03to01.Count - 1 do
        begin
          aDailyAccount_OutAInBCPtr := aSAPDailyAccountReader2_03to01.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_OutAInBCPtr^.snumber;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_OutAInBCPtr^.sname;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_OutAInBCPtr^.dQty;

          ExcelApp.Cells[irow, 7].Value := aDailyAccount_OutAInBCPtr^.dt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_OutAInBCPtr^.dtCheck;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_OutAInBCPtr^.sstock_out;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_OutAInBCPtr^.sdept;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_OutAInBCPtr^.sbillno;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_OutAInBCPtr^.suse1;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_OutAInBCPtr^.snote;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_OutAInBCPtr^.sbiller;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_OutAInBCPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_OutAInBCPtr^.scheckflag;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_OutAInBCPtr^.souttype;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_OutAInBCPtr^.suse2;     
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_OutAInBCPtr^.sedi;

          s_fac := aDailyAccount_OutAInBCPtr^.snumber +
            aDailyAccount_OutAInBCPtr^.sbillno
            ; // + aDailyAccount_OutAInBCPtr^.sstock_out;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
                    
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_03to01 := iCountMatch_03to01 + 1;
              ExcelApp.Cells[irow, 4].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_OutAInBCPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;        
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 4].Value := '0';
            ExcelApp.Cells[irow, 5].Value := - aDailyAccount_OutAInBCPtr^.dQty;
          end;

          irow := irow + 1;      
        end;
      finally
        aSAPDailyAccountReader2_03to01.Free;
      end;                                
    end;      


             
                     



    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('���ϳ���');
                        
    s := mmiQout.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_qout := TSAPDailyAccountReader2_qout_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_qout.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '��Ʒ������';
        ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 3].Value := '����';
        ExcelApp.Cells[irow, 4].Value := 'SAP����';
        ExcelApp.Cells[irow, 5].Value := '����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '�������';
        ExcelApp.Cells[irow, 8].Value := '�����ֿ�';
        ExcelApp.Cells[irow, 9].Value := '���ϲ���';
        ExcelApp.Cells[irow, 10].Value := '���ݱ��';
        ExcelApp.Cells[irow, 11].Value := '��;1';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 14].Value := '��λ';
        ExcelApp.Cells[irow, 15].Value := '��˱�־';
        ExcelApp.Cells[irow, 16].Value := '�������';
        ExcelApp.Cells[irow, 17].Value := '��;2';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';


        irow := irow + 1;
        iCountQout_Fac := aSAPDailyAccountReader2_qout.Count;
        iCountMatch_qout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_qout.Count - 1 do
        begin
          aDailyAccountqoutPtr := aSAPDailyAccountReader2_qout.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqoutPtr^.snumber;
          ExcelApp.Cells[irow, 2].Value := aDailyAccountqoutPtr^.sname;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqoutPtr^.dqty;
                                                                    
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqoutPtr^.dt;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqoutPtr^.dtcheck;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqoutPtr^.sstock;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqoutPtr^.sdetp;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqoutPtr^.sbillno;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqoutPtr^.suse1;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqoutPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccountqoutPtr^.sbiller;
          ExcelApp.Cells[irow, 14].Value := aDailyAccountqoutPtr^.sunit;
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqoutPtr^.scheckflag;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqoutPtr^.souttype;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqoutPtr^.suse2;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqoutPtr^.sedi;

          s_fac := aDailyAccountqoutPtr^.snumber +
            aDailyAccountqoutPtr^.sbillno
            ; // + aDailyAccountqoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];     
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
                   
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_qout := iCountMatch_qout + 1;
              ExcelApp.Cells[irow, 4].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr^.dqty - aDailyAccountqoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;   
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 4].Value := '0';
            ExcelApp.Cells[irow, 5].Value := - aDailyAccountqoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_qout.Free;
      end;
    end;         


    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                  
    Memo1.Lines.Add('����');
              
    s := mmiDB.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                    
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB := TSAPDailyAccountReader2_DB_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_DB.Count > 0 then
    begin
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := '����';

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '����';
        ExcelApp.Cells[irow, 2].Value := '�������';
        ExcelApp.Cells[irow, 3].Value := '���ݱ��';
        ExcelApp.Cells[irow, 4].Value := '�����ֿ�';
        ExcelApp.Cells[irow, 5].Value := '����ֿ�';
        ExcelApp.Cells[irow, 6].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 7].Value := '��������';
        ExcelApp.Cells[irow, 8].Value := '��������';
        ExcelApp.Cells[irow, 9].Value := 'SAP����';
        ExcelApp.Cells[irow, 10].Value := '����';
        ExcelApp.Cells[irow, 11].Value := '��ע';
        ExcelApp.Cells[irow, 12].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 13].Value := '��˱�־';
        ExcelApp.Cells[irow, 14].Value := 'EDI�ύ';


        irow := irow + 1;
        iCountDB_Fac := aSAPDailyAccountReader2_DB.Count;
        iCountMatch_DB := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB.Count - 1 do
        begin
          aDailyAccount_DBPtr := aSAPDailyAccountReader2_DB.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DBPtr^.dt;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_DBPtr^.dtCheck;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DBPtr^.sbillno;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_DBPtr^.sstock_out;   
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_DBPtr^.sstock_in;   
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_DBPtr^.snumber;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_DBPtr^.sname;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_DBPtr^.dQty;
        
          ExcelApp.Cells[irow, 11].Value := '';

          ExcelApp.Cells[irow, 12].Value := aDailyAccount_DBPtr^.sbiller;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_DBPtr^.scheckflag;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_DBPtr^.sedi;
             
          if aDailyAccount_DBPtr^.sstock_out = aDailyAccount_DBPtr^.sstock_in then
          begin
            iCountDB_Fac := iCountDB_Fac - 1; // �����͵���һ�������������
            irow := irow + 1;
            Continue;
          end;

          s_fac := aDailyAccount_DBPtr^.sbillno +
            aDailyAccount_DBPtr^.snumber
            ; // + aDailyAccount_DBPtr^.sstock_in;

          aSAPMB51RecordPtr_match := nil;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];       
            if aSAPMB51RecordPtr^.bCalc then Continue;

            if aSAPMB51RecordPtr^.smovingtype <> '311' then Continue;

            if aSAPMB51RecordPtr^.dqty < 0 then Continue; // ֻ��������
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber;
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr^.dQty) > Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_DBPtr^.dQty) then
              begin                                     
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end;

              if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_DBPtr^.dQty) then
              begin
                Break;
              end;
            end;
          end;     

          if bFound then
          begin    
            if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_DBPtr^.dQty) then
            begin
              iCountMatch_DB := iCountMatch_DB + 1;
            end;

            ExcelApp.Cells[irow, 9].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 10].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr^.dQty;


            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;
          end
          else
          begin
            ExcelApp.Cells[irow, 9].Value := '0';
            ExcelApp.Cells[irow, 10].Value := - aDailyAccount_DBPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_DB.Free;
      end;
    end;


    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                      
                  
    Memo1.Lines.Add('���� ����');
          
    s := mmiDB_out.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                 
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB_out := TSAPDailyAccountReader2_DB_out_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_DB_out.Count > 0 then
    begin
      try    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '��Ʒ������';
        ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 3].Value := '����';
        ExcelApp.Cells[irow, 4].Value := 'SAP����';
        ExcelApp.Cells[irow, 5].Value := '����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '�������';
        ExcelApp.Cells[irow, 8].Value := '�����ֿ�';
        ExcelApp.Cells[irow, 9].Value := '���ϲ���';
        ExcelApp.Cells[irow, 10].Value := '���ݱ��';
        ExcelApp.Cells[irow, 11].Value := '��;1';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 14].Value := '��λ';
        ExcelApp.Cells[irow, 15].Value := '��˱�־';
        ExcelApp.Cells[irow, 16].Value := '�������';
        ExcelApp.Cells[irow, 17].Value := '��;2';
        ExcelApp.Cells[irow, 18].Value := 'EDI�ύ';
             

        irow := irow + 1;
        iCountDB_out_Fac := aSAPDailyAccountReader2_DB_out.Count;
        iCountMatch_DB_out := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB_out.Count - 1 do
        begin
          aDailyAccount_DB_outPtr := aSAPDailyAccountReader2_DB_out.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DB_outPtr^.snumber;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_DB_outPtr^.sname;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DB_outPtr^.dQty;

        
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_DB_outPtr^.dt;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_DB_outPtr^.dtCheck;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_DB_outPtr^.sstock_out;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_DB_outPtr^.sdept;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_DB_outPtr^.sbillno;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_DB_outPtr^.suse1;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_DB_outPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_DB_outPtr^.sbiller;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_DB_outPtr^.sunit;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_DB_outPtr^.scheckflag;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_DB_outPtr^.souttype;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_DB_outPtr^.suse2;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_DB_outPtr^.sedi;

          s_fac := aDailyAccount_DB_outPtr^.snumber +
            aDailyAccount_DB_outPtr^.sbillno
            ; // + aDailyAccount_DB_outPtr^.sstock_out;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];      
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // +  aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_DB_out := iCountMatch_DB_out + 1;
              ExcelApp.Cells[irow, 4].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_DB_outPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;   
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 4].Value := '0';
            ExcelApp.Cells[irow, 5].Value := - aDailyAccount_DB_outPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_DB_out.Free;
      end;
    end;             
                                                  

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                            
    Memo1.Lines.Add('���� ����');
         
    s := mmiDB_in.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                          
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB_in := TSAPDailyAccountReader2_DB_in_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_DB_in.Count > 0 then
    begin
      try    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := '����';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '��ע';
        ExcelApp.Cells[irow, 8].Value := '����';
        ExcelApp.Cells[irow, 9].Value := '�������';
        ExcelApp.Cells[irow, 10].Value := '��;';
        ExcelApp.Cells[irow, 11].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 14].Value := 'ժҪ';
        ExcelApp.Cells[irow, 15].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 16].Value := '�رձ�־';
        ExcelApp.Cells[irow, 17].Value := '����';
        ExcelApp.Cells[irow, 18].Value := '���鷽ʽ';
        ExcelApp.Cells[irow, 19].Value := 'EDI�ύ';
        ExcelApp.Cells[irow, 20].Value := 'Դ������';
             

        irow := irow + 1;
        iCountDB_in_Fac := aSAPDailyAccountReader2_DB_in.Count;
        iCountMatch_DB_in := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB_in.Count - 1 do
        begin
          aDailyAccount_DB_inPtr := aSAPDailyAccountReader2_DB_in.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DB_inPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_DB_inPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DB_inPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_DB_inPtr^.dQty;
        
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_DB_inPtr^.dt;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_DB_inPtr^.dtCheck;
        
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_DB_inPtr^.suse;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_DB_inPtr^.ssupplier;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_DB_inPtr^.snote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_DB_inPtr^.sstock_in;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_DB_inPtr^.ssummary;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_DB_inPtr^.sbiller; 
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_DB_inPtr^.scloseflag;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_DB_inPtr^.sdept;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_DB_inPtr^.schecktype;   
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_DB_inPtr^.sedi;
          //ExcelApp.Cells[irow, 20].Value :=

          s_fac := aDailyAccount_DB_inPtr^.sbillno +
            aDailyAccount_DB_inPtr^.snumber
            ; // + aDailyAccount_DB_inPtr^.sstock_in;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];      
            if aSAPMB51RecordPtr^.bCalc then Continue;

            if aSAPMB51RecordPtr^.dqty < 0 then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              ExcelApp.Cells[irow, 5].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 6].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_DB_inPtr^.dQty;
              if DoubleE( aSAPMB51RecordPtr^.dqty - aDailyAccount_DB_inPtr^.dQty, 0) then
              begin
                iCountMatch_DB_in := iCountMatch_DB_in + 1;
              end;
              aSAPMB51RecordPtr^.bCalc := True;      
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;     

          if not bFound then
          begin
            ExcelApp.Cells[irow, 5].Value := '0';
            ExcelApp.Cells[irow, 6].Value := - aDailyAccount_DB_inPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_DB_in.Free;
      end;
    end;
                                  
                
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                               
    Memo1.Lines.Add('���� Ͷ�ϵ�');
        
    s := mmiPPBom.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + sfile_k3);       
    aSAPDailyAccountReader2_PPBom := TSAPDailyAccountReader2_PPBOM_ml.Create( sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_PPBom.Count > 0 then
    begin

      s := mmiSQ01PPBom.Caption;
      if Pos('(', s) > 0 then
      begin
        s := Copy(s, 1, Pos('(', s) - 1);
      end;
      sfile_sq01_ppbom := vle_ml.Values[s];
      Memo1.Lines.Add(s);

      Memo1.Lines.Add('���ļ��� ' + sfile_sq01_ppbom);      
      aSAPDailyAccountReader2_coois := TSAPDailyAccountReader2_coois.Create(sfile_sq01_ppbom, 'Sheet1', aStockMZ2FacReader);

    
 
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
 
        ExcelApp.Cells[irow, 1].Value := '�Ƶ�����';
        ExcelApp.Cells[irow, 2].Value := '�������';
        ExcelApp.Cells[irow, 3].Value := '����/ί�ⶩ����';
        ExcelApp.Cells[irow, 4].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 5].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 6].Value := '��������';
        ExcelApp.Cells[irow, 7].Value := '��ע';
        ExcelApp.Cells[irow, 8].Value := '����Ͷ�ϵ���';
        ExcelApp.Cells[irow, 9].Value := '�������ϳ�����';
        ExcelApp.Cells[irow, 10].Value := '������������';
        ExcelApp.Cells[irow, 11].Value := '�ƻ�Ͷ������';
        ExcelApp.Cells[irow, 12].Value := 'SAP����';
        ExcelApp.Cells[irow, 13].Value := '����';
        ExcelApp.Cells[irow, 14].Value := 'Ӧ������';
        ExcelApp.Cells[irow, 15].Value := '�ֿ�';
        ExcelApp.Cells[irow, 16].Value := '��λ����';
        ExcelApp.Cells[irow, 17].Value := '��˱�־';
        ExcelApp.Cells[irow, 18].Value := '��������';
        ExcelApp.Cells[irow, 19].Value := 'EDI�ύ';


        irow := irow + 1;
        iCountPPBom := aSAPDailyAccountReader2_PPBom.Count;
        iCountMatch_PPBom := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_PPBom.Count - 1 do
        begin
          ptrDailyAccount_PPBOM := aSAPDailyAccountReader2_PPBom.Items[i_fac];
                                    
          ExcelApp.Cells[irow, 1].Value := ptrDailyAccount_PPBOM^.dtdate;
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_PPBOM^.dtCheck;
          ExcelApp.Cells[irow, 3].Value := ptrDailyAccount_PPBOM^.sicmobillno;
          ExcelApp.Cells[irow, 4].Value := ptrDailyAccount_PPBOM^.snumber;
          ExcelApp.Cells[irow, 5].Value := ptrDailyAccount_PPBOM^.sname;
          ExcelApp.Cells[irow, 6].Value := ptrDailyAccount_PPBOM^.dQty;    
          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_PPBOM^.snote;   
          ExcelApp.Cells[irow, 8].Value := ptrDailyAccount_PPBOM^.sppbombillno;   
          ExcelApp.Cells[irow, 9].Value := ptrDailyAccount_PPBOM^.snumber_item;
          ExcelApp.Cells[irow, 10].Value := ptrDailyAccount_PPBOM^.sname_item;
          ExcelApp.Cells[irow, 11].Value := ptrDailyAccount_PPBOM^.dqtyplan;
         
          ExcelApp.Cells[irow, 14].Value := ptrDailyAccount_PPBOM^.dqtyshould;
          ExcelApp.Cells[irow, 15].Value := ptrDailyAccount_PPBOM^.sstockname;
          ExcelApp.Cells[irow, 16].Value := ptrDailyAccount_PPBOM^.dusage;
          ExcelApp.Cells[irow, 17].Value := ptrDailyAccount_PPBOM^.scheckflag;
          ExcelApp.Cells[irow, 18].Value := ptrDailyAccount_PPBOM^.sworkshopname;  
          ExcelApp.Cells[irow, 19].Value := ptrDailyAccount_PPBOM^.sedi;                
          ExcelApp.Cells[irow, 20].Value := ptrDailyAccount_PPBOM^.sstockname_ml;
 

          s_fac := ptrDailyAccount_PPBOM^.sicmobillno + ptrDailyAccount_PPBOM^.snumber_item;

          bFound := False;
          for i_mz := 0 to aSAPDailyAccountReader2_coois.Count - 1 do
          begin
            ptrDailyAccount_coois := aSAPDailyAccountReader2_coois.Items[i_mz];      
            if ptrDailyAccount_coois^.bCalc then Continue;
          
            sbillno := ptrDailyAccount_coois^.sbillno_fac;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno + ptrDailyAccount_coois^.snumber_item;
            if s_fac = s_mz then
            begin                                              
              bFound := True;
              ExcelApp.Cells[irow, 12].Value := ptrDailyAccount_coois^.dqtyneed;
              ExcelApp.Cells[irow, 13].Value := ptrDailyAccount_coois^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan;
              if DoubleE( ptrDailyAccount_coois^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan, 0) then
              begin
                iCountMatch_PPBom := iCountMatch_PPBom + 1;
              end;
              ptrDailyAccount_coois^.bCalc := True;
              ptrDailyAccount_coois^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 12].Value := '0';
            ExcelApp.Cells[irow, 13].Value := - ptrDailyAccount_PPBOM^.dqtyplan;
            
            if DoubleE( ptrDailyAccount_PPBOM^.dqtyplan, 0) then
            begin
              iCountMatch_PPBom := iCountMatch_PPBom + 1;
            end;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_coois.Free;
  //      aSAPDailyAccountReader2_icmo_mz2fac.Free;
        aSAPDailyAccountReader2_PPBom.Free;
      end;

    end;        

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                                   
    Memo1.Lines.Add('���� ��������');
                    
    s := mmiSOut.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    
   
    Memo1.Lines.Add('���ļ��� ' + s);
            
    aSAPDailyAccountReader2_sout := TSAPDailyAccountReader2_sout_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_sout.Count > 0 then
    begin
      try    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        ExcelApp.Cells[irow, 1].Value := '�������񵥺�';
        ExcelApp.Cells[irow, 2].Value := '����';
        ExcelApp.Cells[irow, 3].Value := '�������';
        ExcelApp.Cells[irow, 4].Value := '�ɱ��������';
        ExcelApp.Cells[irow, 5].Value := '�ɱ�����';
        ExcelApp.Cells[irow, 6].Value := '��ע';
        ExcelApp.Cells[irow, 7].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 8].Value := '��������';
        ExcelApp.Cells[irow, 9].Value := 'ʵ������';
        ExcelApp.Cells[irow, 10].Value := 'SAP����';
        ExcelApp.Cells[irow, 11].Value := '����';
        ExcelApp.Cells[irow, 12].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 13].Value := '���ݱ��';
        ExcelApp.Cells[irow, 14].Value := '���ϲ���';
        ExcelApp.Cells[irow, 15].Value := '������;';
        ExcelApp.Cells[irow, 16].Value := '����';
        ExcelApp.Cells[irow, 17].Value := '�����';
        ExcelApp.Cells[irow, 18].Value := '��˱�־';
        ExcelApp.Cells[irow, 19].Value := '�Ƶ�';
        ExcelApp.Cells[irow, 20].Value := 'EDI�ύ';
 

        irow := irow + 1;
        iCountSout_Fac := aSAPDailyAccountReader2_sout.Count;
        iCountMatch_Sout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_sout.Count - 1 do
        begin
          aDailyAccount_soutPtr := aSAPDailyAccountReader2_sout.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_soutPtr^.sicmo;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_soutPtr^.dt;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_soutPtr^.dtCheck;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_soutPtr^.scostnumber;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_soutPtr^.scostname;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_soutPtr^.snote;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_soutPtr^.snumber;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_soutPtr^.sname;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_soutPtr^.dqty;

          ExcelApp.Cells[irow, 12].Value := aDailyAccount_soutPtr^.sstock;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_soutPtr^.sbillno;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_soutPtr^.sdept;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_soutPtr^.suse;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_soutPtr.sbatchno;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_soutPtr^.schecker;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_soutPtr^.scheckflag;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_soutPtr^.sbiller;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_soutPtr^.sedi;

          s_fac := aDailyAccount_soutPtr^.snumber +
            aDailyAccount_soutPtr^.sbillno
            ; // + aDailyAccount_soutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];          
            if aSAPMB51RecordPtr^.bCalc then Continue;

            if (aSAPMB51RecordPtr^.smovingtype <> 'X01') and
              (aSAPMB51RecordPtr^.smovingtype <> 'X02') then
              Continue;
              
            if aSAPMB51RecordPtr^.dqty < 0 then Continue;

          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_Sout := iCountMatch_Sout + 1;
              ExcelApp.Cells[irow, 10].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 11].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_soutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 10].Value := '0';
            ExcelApp.Cells[irow, 11].Value := aDailyAccount_soutPtr^.dQty;
          end;

          irow := irow + 1;      
        end;
      finally
        aSAPDailyAccountReader2_sout.Free;
      end;
    end;             
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////                                          

    (*
    Memo1.Lines.Add('���۳���');

    s := mmiXOut.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];


    Memo1.Lines.Add('���ļ��� ' + s);

    aSAPDailyAccountReader2_xout := TSAPDailyAccountReader2_xout_ml.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_xout.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '������λ';
        ExcelApp.Cells[irow, 3].Value := '�Ϻ�';
        ExcelApp.Cells[irow, 4].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 5].Value := '����';

        ExcelApp.Cells[irow, 6].Value := 'SAP';
        ExcelApp.Cells[irow, 7].Value := '����';


        ExcelApp.Cells[irow, 8].Value := '��������';
        ExcelApp.Cells[irow, 9].Value := '�����̼��';
        ExcelApp.Cells[irow, 10].Value := '��ݹ�˾';
        ExcelApp.Cells[irow, 11].Value := '���ӵ���';
        ExcelApp.Cells[irow, 12].Value := '������ע';
        ExcelApp.Cells[irow, 13].Value := '����ʱ��';
        ExcelApp.Cells[irow, 14].Value := '��λ';
        ExcelApp.Cells[irow, 15].Value := '����';
        ExcelApp.Cells[irow, 16].Value := '��ע';


        irow := irow + 1;
        iCountSout_Fac := aSAPDailyAccountReader2_xout.Count;
        iCountMatch_Sout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_xout.Count - 1 do
        begin
          aDailyAccount_xoutPtr := aSAPDailyAccountReader2_xout.Items[i_fac];


          ExcelApp.Cells[irow, 1].Value := aDailyAccount_xoutPtr^.sxoutbillno;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_xoutPtr^.sxoutdept;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_xoutPtr^.snumber;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_xoutPtr^.sname;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_xoutPtr^.dQty;

//          ExcelApp.Cells[irow, 6].Value := 'SAP';
//          ExcelApp.Cells[irow, 7].Value := '����';


          ExcelApp.Cells[irow, 8].Value := aDailyAccount_xoutPtr^.sorder;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_xoutPtr^.sproxy;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_xoutPtr^.sexp;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_xoutPtr^.sebillno;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_xoutPtr^.smnote;
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_xoutPtr^.sddate;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_xoutPtr^.sstock_fac;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_xoutPtr^.sdate;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_xoutPtr^.snote;



          s_fac := aDailyAccount_xoutPtr^.snumber +
            aDailyAccount_xoutPtr^.sbillno
            ; // + aDailyAccount_xoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr^.bCalc then Continue;

            if (aSAPMB51RecordPtr^.smovingtype <> 'X01') and
              (aSAPMB51RecordPtr^.smovingtype <> 'X02') then
              Continue;

            if aSAPMB51RecordPtr^.dqty < 0 then Continue;


            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'ML' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin
              bFound := True;
              iCountMatch_Sout := iCountMatch_Sout + 1;
              ExcelApp.Cells[irow, 10].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 11].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_xoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 10].Value := '0';
            ExcelApp.Cells[irow, 11].Value := aDailyAccount_xoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_xout.Free;
      end;
    end;
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    *)

    sl := TStringList.Create;
    try
      WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
      iSheet := iSheet + 1;
      ExcelApp.Sheets[iSheet].Activate;
      ExcelApp.Sheets[iSheet].Name := 'MB51';


      sline := '����ƾ֤'#9'ƾ֤����'#9'���ص�'#9'�ִ��ص������'#9'ƾ̧֤ͷ�ı�'#9'�ƶ�����'#9'���ϱ���'#9'��������'#9'��¼�뵥λ��ʾ������'#9'��������'#9'��������'#9'����ʱ��'#9'����'#9'�ɹ�����'#9'�Ƿ�ƥ��'#9'ƥ�䵥��';
      sl.Add(sline);

      for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
      begin
        aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
        sline := aSAPMB51RecordPtr^.sbillno + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.fdate) + #9
          + aSAPMB51RecordPtr^.fstockno + #9
          + aSAPMB51RecordPtr^.fstockname + #9
          + aSAPMB51RecordPtr^.fnote + #9
          + aSAPMB51RecordPtr^.smovingtype + #9         
          + aSAPMB51RecordPtr^.snumber + #9
          + aSAPMB51RecordPtr^.sname + #9
          + FloatToStr(aSAPMB51RecordPtr^.dqty) + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.fdate) + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.finputdate) + #9
          + FormatDateTime('HH:mm:ss', aSAPMB51RecordPtr^.finputtime) + #9
          + aSAPMB51RecordPtr^.spo + #9
          + aSAPMB51RecordPtr^.sbillno_po + #9
          +  CSBoolean[aSAPMB51RecordPtr^.bCalc] + #9
          +  aSAPMB51RecordPtr^.sMatchType;
        sl.Add(sline);
      end;

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;     
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select; 
                 
    finally
      sl.Free;
    end;
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Columns[1].ColumnWidth := 14.38;     
    ExcelApp.Columns[2].ColumnWidth := 21.63;
    ExcelApp.Columns[3].ColumnWidth := 13.63;
    ExcelApp.Columns[4].ColumnWidth := 12.38;
    ExcelApp.Columns[5].ColumnWidth := 16.50;
    ExcelApp.Columns[6].ColumnWidth := 15;
    ExcelApp.Columns[7].ColumnWidth := 21.50;
    ExcelApp.Columns[8].ColumnWidth := 78.75;

    irow := 1;
    
    AddHorizontalAlignment(ExcelApp, irow, 1, irow, 8, xlCenter);  
    AddHorizontalAlignment(ExcelApp, irow + 1, 1, irow + 12, 7, xlCenter);

    ExcelApp.Cells[irow, 1].Value := '����';
    ExcelApp.Cells[irow, 2].Value := '�嵥������';
    MergeCells(ExcelApp, irow, 2, irow, 3);
    ExcelApp.Cells[irow, 4].Value := '�����ᱨ����';
    ExcelApp.Cells[irow, 5].Value := 'SAP��ʽ����';
    ExcelApp.Cells[irow, 6].Value := '������SAP����';
    ExcelApp.Cells[irow, 7].Value := '��ע';
    ExcelApp.Cells[irow, 8].Value := '���촦�����';

		AddColor(ExcelApp, irow, 1, irow, 8, $B7B8E6);
		AddColor(ExcelApp, irow, 6, irow, 7, $DCCD92);

    irow := 2;
    ExcelApp.Cells[irow, 1].Value := FormatDateTime('yyyy/MM/dd', Now);
    MergeCells(ExcelApp, irow, 1, irow + 11, 1);

    ExcelApp.Cells[irow, 2].Value := '�⹺��ⵥ';
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);
    ExcelApp.Cells[irow, 3].Value := 'PO����';
    ExcelApp.Cells[irow + 1, 3].Value := 'PO����';
    AddColor(ExcelApp, irow, 3, irow, 8, $DAC0CC);  
    AddColor(ExcelApp, irow + 1, 3, irow + 1, 8, $DEF1EB);

    ExcelApp.Cells[irow, 4].Value := iCountWinB_Fac; 
    ExcelApp.Cells[irow, 5].Value := iCountMatch_WinB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);
                           
    ExcelApp.Cells[irow + 1, 4].Value := iCountWinR_Fac;
    ExcelApp.Cells[irow + 1, 5].Value := iCountMatch_WinR;
    ExcelApp.Cells[irow + 1, 6].Value := '=D' + IntToStr(irow + 1) + '-E' + IntToStr(irow + 1);

    irow := irow + 2;
    ExcelApp.Cells[irow, 2].Value := '��Ʒ���';  
    ExcelApp.Cells[irow, 4].Value := iCountcpin_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_cpin;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    AddColor(ExcelApp, irow, 6, irow + 8, 7, $F3EEDA);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '������ⵥ';
    ExcelApp.Cells[irow, 3].Value := 'Sample';
    ExcelApp.Cells[irow, 4].Value := iCountqin_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_qin;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '�Ϻŵ���';
    ExcelApp.Cells[irow, 4].Value := iCountA2B_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_a2b;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '�������ɢ��';
    ExcelApp.Cells[irow, 4].Value := iCount03to01_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_03to01;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '���ϳ���';
    ExcelApp.Cells[irow, 4].Value := iCountqout_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_qout;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '���ε�';
    MergeCells(ExcelApp, irow, 2, irow + 2, 2);
    ExcelApp.Cells[irow, 3].Value := '�������ڲ���';
    ExcelApp.Cells[irow + 1, 3].Value := '���루��������';
    ExcelApp.Cells[irow + 2, 3].Value := '��������������';

    ExcelApp.Cells[irow, 4].Value := iCountDB_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_DB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    ExcelApp.Cells[irow + 1, 4].Value := iCountDB_in_Fac;
    ExcelApp.Cells[irow + 1, 5].Value := iCountMatch_DB_in;
    ExcelApp.Cells[irow + 1, 6].Value := '=D' + IntToStr(irow + 1) + '-E' + IntToStr(irow + 1);

    ExcelApp.Cells[irow + 2, 4].Value := iCountDB_Out_Fac;
    ExcelApp.Cells[irow + 2, 5].Value := iCountMatch_DB_out;
    ExcelApp.Cells[irow + 2, 6].Value := '=D' + IntToStr(irow + 2) + '-E' + IntToStr(irow + 2);

    AddColor(ExcelApp, irow + 1, 3, irow + 1, 8, $B4D5FC);   
    AddColor(ExcelApp, irow + 2, 3, irow + 2, 8, $9BD7C4);

    irow := irow + 3;
    ExcelApp.Cells[irow, 2].Value := '����Ͷ�ϵ�';
    ExcelApp.Cells[irow, 4].Value := iCountPPBom;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_PPBom;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ϵ�';
    ExcelApp.Cells[irow, 4].Value := iCountSout_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_Sout;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);
       
    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�⹺�������-������';
    ExcelApp.Cells[irow, 4].Value := iCountWinB_DB_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_WinB_DB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);



    AddBorder(ExcelApp, 1, 1, 14, 8);
    
                



    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
    

  finally
    aSAPMB51Reader2.Free;
    aSAPCMSPushErrorReader2.Free;     
    aStockMZ2FacReader.Free;
  end;
         

  MessageBox(Handle, '���', '��ʾ', 0);
end;                    

procedure TfrmFacAccountCheck.btnSaveClick_wt;
const
  CSBoolean: array[Boolean] of string = ('��', '��');
var
  ExcelApp, WorkBook: Variant;
  aSAPMB51Reader2: TSAPMB51Reader2;
  aSAPCMSPushErrorReader2: TSAPCMSPushErrorReader2;
  aICMO2FacReader2: TICMO2FacReader2;
  iSheet: Integer;
  irow: Integer;
  sfile: string;   
  aStockMZ2FacReader: TStockMZ2FacReader;
  
  aSAPDailyAccountReader2_winB: TSAPDailyAccountReader2_winB;
  aSAPDailyAccountReader2_winR: TSAPDailyAccountReader2_winB;
  aSAPDailyAccountReader2_RTV: TSAPDailyAccountReader2_RTV;
  aSAPDailyAccountReader2_cpin: TSAPDailyAccountReader2_cpin;
  aSAPDailyAccountReader2_qin: TSAPDailyAccountReader2_qin;
  aSAPDailyAccountReader2_a2b: TSAPDailyAccountReader2_qout;
  aSAPDailyAccountReader2_03to01: TSAPDailyAccountReader2_03to01;
  aSAPDailyAccountReader2_qout: TSAPDailyAccountReader2_qout;
  aSAPDailyAccountReader2_DB: TSAPDailyAccountReader2_DB;      
  aSAPDailyAccountReader2_DB_in: TSAPDailyAccountReader2_DB_in;
  aSAPDailyAccountReader2_DB_out: TSAPDailyAccountReader2_DB_out;
  aSAPDailyAccountReader2_sout: TSAPDailyAccountReader2_sout;

  aSAPDailyAccountReader2_coois: TSAPDailyAccountReader2_coois;
  aSAPDailyAccountReader2_PPBom: TSAPDailyAccountReader2_PPBOM;

  i_fac: Integer;
  i_mz: Integer;
  s_fac: string;
  s_mz: string;
  bFound: Boolean;

  iCountWinB_Fac: Integer;      
  iCountWinR_Fac: Integer;
  iCountCPIN_Fac: Integer;
  iCountQIn_Fac: Integer;
  iCountA2B_Fac: Integer;     //�Ϻŵ���
  iCount03to01_Fac: Integer;  //�������ɢ��  
  iCountQout_Fac: Integer;    //���ϳ���
  iCountDB_Fac: Integer;
  iCountDB_in_Fac: Integer;
  iCountDB_out_Fac: Integer;
  iCountSout_Fac: Integer;  
  iCountPPBom: Integer;


  iCountMatch_WinB: Integer;
  iCountMatch_WinR: Integer;
  iCountMatch_cpin: Integer;
  iCountMatch_qin: Integer;
  iCountMatch_A2B: Integer;
  iCountMatch_03to01: Integer;
  iCountMatch_qout: Integer;
  iCountMatch_DB: Integer;
  iCountMatch_DB_out: Integer;
  iCountMatch_DB_in: Integer;
  iCountMatch_Sout: Integer;
  iCountMatch_PPBom: Integer;
  iCountMatch_PPBom_mz: Integer;

  aSAPMB51RecordPtr: PSAPMB51Record;
  aDailyAccount_winBPtr: PDailyAccount_winB;
  aDailyAccount_RTVPtr: PDailyAccount_RTV;
  aDailyAccount_cpinPtr: PDailyAccount_cpin;
  aDailyAccountqinPtr: PDailyAccount_qin;
  aDailyAccountqoutPtr: PDailyAccount_qout;
  aDailyAccount_DBPtr: PDailyAccount_DB;
  aDailyAccount_DBPtr2: PDailyAccount_DB;
  aDailyAccount_DB_inPtr: PDailyAccount_DB_in;
  aDailyAccount_DB_outPtr: PDailyAccount_DB_out;
  aDailyAccount_OutAInBCPtr: PDailyAccount_OutAInBC;
  aDailyAccount_soutPtr: PDailyAccount_sout;
  ptrDailyAccount_PPBOM: PDailyAccount_PPBom;
  ptrDailyAccount_coois: PDailyAccount_coois;

//  aCPINmz2facReader: TCPINmz2facReader;

  aSAPMB51RecordPtr_match: PSAPMB51Record;
  
  s: string;
  s2: string;
  sfile_k3: string;                                                                

  sfile_sq01_ppbom: string;

  sbillno: string;
  idx: Integer;
  dDelta: Double;
  sl: TStringList;
  sline: string;

  dQtyMatchx: Double;
  dQtyMatch0: Double;
  ptrDailyAccount_coois_match: PDailyAccount_coois;
begin
  if not ExcelSaveDialog(sfile) then Exit;
                                                                        
  aSAPMB51Reader2 := TSAPMB51Reader2.Create(leMB51.Text, nil);
  aStockMZ2FacReader := TStockMZ2FacReader_wt.Create(leStockFac2MZ.Text);
//  aSAPCMSPushErrorReader2 := TSAPCMSPushErrorReader2.Create(leCMSErrMsg.Text, nil);
  aICMO2FacReader2 := TICMO2FacReader2.Create(leICMO2Fac.Text, nil);

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

    Memo1.Lines.Add('����');

    WorkBook := ExcelApp.WorkBooks.Add;
    ExcelApp.DisplayAlerts := False;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate; 
    ExcelApp.Sheets[iSheet].Name := '����';
                  

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    s := mmiWinB_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winB := TSAPDailyAccountReader2_winB_wt.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_winB.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);
//        AddColor(ExcelApp, irow, 6, irow, 6, clRed);


        irow := irow + 1;
        iCountWinB_Fac := aSAPDailyAccountReader2_winB.Count;
        iCountMatch_WinB := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_winB.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winB.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_winBPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_winBPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_winBPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_winBPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_winBPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_winBPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_winBPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_winBPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_winBPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_winBPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_winBPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_winBPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_winBPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_winBPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_winBPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_winBPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_winBPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_winBPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_winBPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_winBPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_winBPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_winBPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_winBPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_winBPtr^.sstock_desc_wt;


          s_fac := aDailyAccount_winBPtr^.sbillno +
            aDailyAccount_winBPtr^.snumber +
            aDailyAccount_winBPtr^.sitemtext  ;       // �ɹ�����


          aSAPMB51RecordPtr_match := nil; 
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr^.smovingtype <> '101' then Continue;

            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
                          
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber +
              aSAPMB51RecordPtr^.sbillno_po;// �ɹ�����

            if s_fac = s_mz then
            begin
              bFound := True;
                      
              dQtyMatchx := aSAPMB51Reader2.GetMB51Qty101(aSAPMB51RecordPtr);
              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match :=  aSAPMB51RecordPtr;   
                dQtyMatch0 := dQtyMatchx;
              end
              else
              begin
                if Abs(dQtyMatch0 - aDailyAccount_winBPtr^.dQty) >
                  Abs(dQtyMatchx - aDailyAccount_winBPtr^.dQty) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;        
                  dQtyMatch0 := dQtyMatchx;
                end;
              end;  
              
              if DoubleE(dQtyMatch0 - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                Break;
              end; 
            end;
          end;

          if bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := dQtyMatch0;
            ExcelApp.Cells[irow, 14].Value := dQtyMatch0 - aDailyAccount_winBPtr^.dQty;
                     
            aSAPMB51Reader2.SetCalcFlag(aSAPMB51RecordPtr_match, s);
    
            if DoubleE(dQtyMatch0, aDailyAccount_winBPtr^.dQty) then
            begin
              iCountMatch_WinB := iCountMatch_WinB + 1;
            end;
          end
          else
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccount_winBPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winB.Free;
      end;
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    s := mmiWinR_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
 
                                                      
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winR := TSAPDailyAccountReader2_winB_wt.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_winR.Count > 0 then
    begin
      try
    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);
//        AddColor(ExcelApp, irow, 6, irow, 6, clRed);

 
        irow := irow + 1;
        iCountWinR_Fac := aSAPDailyAccountReader2_winR.Count;
        iCountMatch_WinR := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_winR.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winR.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_winBPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_winBPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_winBPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_winBPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_winBPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_winBPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_winBPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_winBPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_winBPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_winBPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_winBPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_winBPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_winBPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_winBPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_winBPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_winBPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_winBPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_winBPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_winBPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_winBPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_winBPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_winBPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_winBPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_winBPtr^.sstock_desc_wt;


          s_fac := aDailyAccount_winBPtr^.sbillno +
            aDailyAccount_winBPtr^.snumber +
            aDailyAccount_winBPtr^.sitemtext  ;       // �ɹ�����


          aSAPMB51RecordPtr_match := nil;
          bFound := False;
          dQtyMatchx := 0;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
              
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber
              + aSAPMB51RecordPtr^.sbillno_po;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else
              begin
                if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty) >
                  Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_winBPtr^.dQty) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
                end;
              end;
              if DoubleE(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                Break;
              end;
            end;
          end;     

          if bFound then
          begin 
            ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_winBPtr^.dQty;

            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;

            iCountMatch_WinR := iCountMatch_WinR + 1;
          end
          else
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := aDailyAccount_winBPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winR.Free;
      end;
    end;          
 
         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    s := mmiCPIN_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);


    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_cpin := TSAPDailyAccountReader2_cpin_wt.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_cpin.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';   
        ExcelApp.Cells[irow, 36].Value := '���幤����';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);


        irow := irow + 1;
        iCountCPIN_Fac := aSAPDailyAccountReader2_cpin.Count;
        iCountMatch_CPIN := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_cpin.Count - 1 do
        begin
          aDailyAccount_cpinPtr := aSAPDailyAccountReader2_cpin.Items[i_fac];
 
          ExcelApp.Cells[irow, 1].Value := aDailyAccount_cpinPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_cpinPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_cpinPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_cpinPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_cpinPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_cpinPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_cpinPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_cpinPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_cpinPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_cpinPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_cpinPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_cpinPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_cpinPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_cpinPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_cpinPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_cpinPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_cpinPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_cpinPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_cpinPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_cpinPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_cpinPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_cpinPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_cpinPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_cpinPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_cpinPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_cpinPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_cpinPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_cpinPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_cpinPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_cpinPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_cpinPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_cpinPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_cpinPtr^.sstock_desc_wt;
          ExcelApp.Cells[irow, 36].Value := aICMO2FacReader2.ICMOFac2MZ(aDailyAccount_cpinPtr^.sicmo);

          s_fac :=  aDailyAccount_cpinPtr.sdoc + aDailyAccount_cpinPtr^.snumber +
            aDailyAccount_cpinPtr^.sstock;
                    
          bFound := False;
          dDelta := 9999999999;
          idx := -1;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr.bCalc then Continue;

            if (aSAPMB51RecordPtr^.smovingtype <> '101') and
              (aSAPMB51RecordPtr^.smovingtype <> '102') then
            begin
              Continue;
            end;                        

            if aSAPMB51RecordPtr^.fstockname = ''  then // Ҫ�вֿ�����
            begin
              Continue;
            end;

            sbillno := aSAPMB51RecordPtr^.snote_entry;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;              
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
            sbillno := Copy(sbillno, 5, Length(sbillno) - 4);

            

//            s_mz := aCPINmz2facReader.cpin_mz2fac(aSAPMB51RecordPtr^.sbillno) +
//              aSAPMB51RecordPtr^.snumber + aSAPMB51RecordPtr^.fstockname;
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin
              bFound := True; 
              if dDelta > aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty then
              begin
                dDelta := Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty);
                idx := i_mz;
              end;
              if DoubleE(dDelta, 0) then Break;
            end;
          end;     

          if bFound then
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[idx];
            ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty;
            ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty;
            if DoubleE(dDelta, 0) then
            begin
              iCountMatch_CPIN := iCountMatch_CPIN + 1;            
            end;
            aSAPMB51RecordPtr^.bCalc := True;   
            aSAPMB51RecordPtr^.sMatchType := s;
          end
          else
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccount_cpinPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_cpin.Free;
//        aCPINmz2facReader.Free;
      end;
    end;
           

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    // ������ⵥ - Sample                       
    s := mmiQin_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
        
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_qin := TSAPDailyAccountReader2_qin_wt.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_qin.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);
//        AddColor(ExcelApp, irow, 6, irow, 6, clRed);


        irow := irow + 1;
        iCountQIn_Fac := aSAPDailyAccountReader2_qin.Count;
        iCountMatch_qin := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_qin.Count - 1 do
        begin
          aDailyAccountqinPtr := aSAPDailyAccountReader2_qin.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqinPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccountqinPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqinPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqinPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccountqinPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqinPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqinPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqinPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqinPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqinPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqinPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqinPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqinPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqinPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqinPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqinPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqinPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccountqinPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccountqinPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccountqinPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccountqinPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccountqinPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccountqinPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccountqinPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccountqinPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccountqinPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccountqinPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccountqinPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccountqinPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccountqinPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccountqinPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccountqinPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccountqinPtr^.sstock_desc_wt;


          s_fac := aDailyAccountqinPtr^.sbillno +
            aDailyAccountqinPtr^.snumber;

          aSAPMB51RecordPtr_match := nil;
          dQtyMatchx := 0;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr^.smovingtype <> '511' then Continue;

            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
              
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber;

            if s_fac = s_mz then
            begin
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else
              begin
                if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccountqinPtr^.dQty) >
                  Abs(aSAPMB51RecordPtr^.dqty - aDailyAccountqinPtr^.dQty) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
                end;  
              end;
              if DoubleE(aSAPMB51RecordPtr_match^.dqty - aDailyAccountqinPtr^.dQty, 0) then
              begin
                Break;
              end;
            end;
          end;

          if bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccountqinPtr^.dQty;

            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;
            iCountMatch_qin := iCountMatch_qin + 1;
          end
          else
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccountqinPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_qin.Free;
      end;
    end;
     
                     

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                           
    s := mmiA2B_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                    
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_a2b := TSAPDailyAccountReader2_qout_wt.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_a2b.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        
        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);


        irow := irow + 1;
        iCountA2B_Fac := aSAPDailyAccountReader2_a2b.Count;
        iCountMatch_A2B := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_a2b.Count - 1 do
        begin
          aDailyAccountqoutPtr := aSAPDailyAccountReader2_a2b.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqoutPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccountqoutPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqoutPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqoutPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccountqoutPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqoutPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqoutPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqoutPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqoutPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqoutPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqoutPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqoutPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqoutPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqoutPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqoutPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqoutPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqoutPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccountqoutPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccountqoutPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccountqoutPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccountqoutPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccountqoutPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccountqoutPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccountqoutPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccountqoutPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccountqoutPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccountqoutPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccountqoutPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccountqoutPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccountqoutPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccountqoutPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccountqoutPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccountqoutPtr^.sstock_desc_wt;          

          s_fac := aDailyAccountqoutPtr^.snumber +
            aDailyAccountqoutPtr^.sbillno;
            ; // + aDailyAccountqoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
                  
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_A2B := iCountMatch_A2B + 1;
              ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr^.dqty - aDailyAccountqoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccountqoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_a2b.Free;
      end;
    end; 
             
                     

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('���ϳ���');
                        
    s := mmiQout_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_qout := TSAPDailyAccountReader2_qout_wt.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_qout.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);




        irow := irow + 1;
        iCountQout_Fac := aSAPDailyAccountReader2_qout.Count;
        iCountMatch_qout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_qout.Count - 1 do
        begin
          aDailyAccountqoutPtr := aSAPDailyAccountReader2_qout.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqoutPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccountqoutPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqoutPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqoutPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccountqoutPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqoutPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqoutPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqoutPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqoutPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqoutPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqoutPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqoutPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqoutPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqoutPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqoutPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqoutPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqoutPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccountqoutPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccountqoutPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccountqoutPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccountqoutPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccountqoutPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccountqoutPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccountqoutPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccountqoutPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccountqoutPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccountqoutPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccountqoutPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccountqoutPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccountqoutPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccountqoutPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccountqoutPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccountqoutPtr^.sstock_desc_wt;          

          s_fac := aDailyAccountqoutPtr^.snumber +
            aDailyAccountqoutPtr^.sbillno
            ; // + aDailyAccountqoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];     
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
                      
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
             
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_qout := iCountMatch_qout + 1;
              ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr^.dqty - aDailyAccountqoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccountqoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_qout.Free;
      end;
    end;         

        


    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

                  
    Memo1.Lines.Add('����');
              
    s := mmiDB_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                    
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB := TSAPDailyAccountReader2_DB_wt.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_DB.Count > 0 then
    begin
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := '����';

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);



        irow := irow + 1;
        iCountDB_Fac := aSAPDailyAccountReader2_DB.Count;
        iCountMatch_DB := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB.Count - 1 do
        begin
          aDailyAccount_DBPtr := aSAPDailyAccountReader2_DB.Items[i_fac];

          if aDailyAccount_DBPtr^.dQty < 0 then Continue;
          
          if aDailyAccount_DBPtr^.bCalc = True then Continue;

          aDailyAccount_DBPtr^.bCalc := True;

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DBPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_DBPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DBPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_DBPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_DBPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_DBPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_DBPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_DBPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_DBPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_DBPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_DBPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_DBPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_DBPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_DBPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_DBPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_DBPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_DBPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_DBPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_DBPtr^.sstock_desc;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_DBPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_DBPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_DBPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_DBPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_DBPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_DBPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_DBPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_DBPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_DBPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_DBPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_DBPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_DBPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_DBPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_DBPtr^.sstock_desc_wt;


                                      
          aDailyAccount_DBPtr2 := TSAPDailyAccountReader2_DB_wt(aSAPDailyAccountReader2_DB).GetItem2(aDailyAccount_DBPtr);
          if aDailyAccount_DBPtr2 <> nil then
          begin
            aDailyAccount_DBPtr2^.bCalc := True;

            ExcelApp.Cells[irow + 1, 1].Value := aDailyAccount_DBPtr2^.sfacname;
            ExcelApp.Cells[irow + 1, 2].Value :=  aDailyAccount_DBPtr2^.sbillno;
            ExcelApp.Cells[irow + 1, 3].Value := aDailyAccount_DBPtr2^.sdoc;
            ExcelApp.Cells[irow + 1, 4].Value := aDailyAccount_DBPtr2^.dt;
            ExcelApp.Cells[irow + 1, 5].Value := aDailyAccount_DBPtr2^.smpn;
            ExcelApp.Cells[irow + 1, 6].Value := aDailyAccount_DBPtr2^.smpn_name;
            ExcelApp.Cells[irow + 1, 7].Value := aDailyAccount_DBPtr2^.smvt;
            ExcelApp.Cells[irow + 1, 8].Value := aDailyAccount_DBPtr2^.smvr;
            ExcelApp.Cells[irow + 1, 9].Value := aDailyAccount_DBPtr2^.snumber_wt;
            ExcelApp.Cells[irow + 1, 10].Value := aDailyAccount_DBPtr2^.snumber;
            ExcelApp.Cells[irow + 1, 11].Value := aDailyAccount_DBPtr2^.smodel;
            ExcelApp.Cells[irow + 1, 12].Value := aDailyAccount_DBPtr2^.dQty;

          
            ExcelApp.Cells[irow + 1, 15].Value := aDailyAccount_DBPtr2^.sunit;
            ExcelApp.Cells[irow + 1, 16].Value := aDailyAccount_DBPtr2^.stext;
            ExcelApp.Cells[irow + 1, 17].Value := aDailyAccount_DBPtr2^.swc;
            ExcelApp.Cells[irow + 1, 18].Value := aDailyAccount_DBPtr2^.sitemtext;
            ExcelApp.Cells[irow + 1, 19].Value := aDailyAccount_DBPtr2^.sitemno;
            ExcelApp.Cells[irow + 1, 20].Value := aDailyAccount_DBPtr2^.sstock_wt;
            ExcelApp.Cells[irow + 1, 21].Value := aDailyAccount_DBPtr2^.sstock_desc;
            ExcelApp.Cells[irow + 1, 22].Value := aDailyAccount_DBPtr2^.sfacno;
            ExcelApp.Cells[irow + 1, 23].Value := aDailyAccount_DBPtr2^.sitemgroupname;
            ExcelApp.Cells[irow + 1, 24].Value := aDailyAccount_DBPtr2^.smvr_desc;
            ExcelApp.Cells[irow + 1, 25].Value := aDailyAccount_DBPtr2^.sitemgroup;
            ExcelApp.Cells[irow + 1, 26].Value := aDailyAccount_DBPtr2^.sordertype;
            ExcelApp.Cells[irow + 1, 27].Value := aDailyAccount_DBPtr2^.dicmoqty;
            ExcelApp.Cells[irow + 1, 28].Value := aDailyAccount_DBPtr2^.sdoc_item;
            ExcelApp.Cells[irow + 1, 29].Value := aDailyAccount_DBPtr2^.smvt_desc;
            ExcelApp.Cells[irow + 1, 30].Value := aDailyAccount_DBPtr2^.sstatus;
            ExcelApp.Cells[irow + 1, 31].Value := aDailyAccount_DBPtr2^.dtbill;
            ExcelApp.Cells[irow + 1, 32].Value := aDailyAccount_DBPtr2^.dbillqty;
            ExcelApp.Cells[irow + 1, 33].Value := aDailyAccount_DBPtr2^.sfac;
            ExcelApp.Cells[irow + 1, 34].Value := aDailyAccount_DBPtr2^.sicmo;
            ExcelApp.Cells[irow + 1, 35].Value := aDailyAccount_DBPtr2^.sstock_desc_wt;
                     

            if aDailyAccount_DBPtr^.sstock_desc = aDailyAccount_DBPtr2^.sstock_desc then // �����ֿ������ֿ��Ӧ����ͬһ���ֿ�
            begin
              ExcelApp.Cells[irow, 36].Value := aDailyAccount_DBPtr^.sstock_desc;
              ExcelApp.Cells[irow + 1, 36].Value := aDailyAccount_DBPtr2^.sstock_desc;
              iCountMatch_DB := iCountMatch_DB + 2;
              irow := irow + 2;
              Continue;
            end;
          end
          else
          begin
            aDailyAccount_DBPtr2 := TSAPDailyAccountReader2_DB_wt(aSAPDailyAccountReader2_DB).GetItem2(aDailyAccount_DBPtr);
          end;

          s_fac := aDailyAccount_DBPtr^.sbillno +
            aDailyAccount_DBPtr^.snumber +
            aDailyAccount_DBPtr^.sstock_desc;

          if (aDailyAccount_DBPtr^.sbillno = 'D180809000114')
            and (aDailyAccount_DBPtr^.snumber = '83.68.36802905OS') then
          begin
            Sleep(12);
          end;

          aSAPMB51RecordPtr_match := nil;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];       
            if aSAPMB51RecordPtr^.bCalc then Continue;

            if aSAPMB51RecordPtr^.smovingtype <> '311' then Continue;

//            if aSAPMB51RecordPtr^.dqty < 0 then Continue; // ֻ��������

            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                    
            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end; 
          
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber +
              aSAPMB51RecordPtr^.fstockname;

          
            if (sbillno = 'D180809000114')
              and (aSAPMB51RecordPtr^.snumber = '83.68.36802905OS') then
            begin
              Sleep(12);
            end;

            if s_fac = s_mz then
            begin                                              
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr^.dQty) >

                Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_DBPtr^.dQty) then
              begin                                     
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end;

              if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_DBPtr^.dQty) then
              begin
                Break;
              end;
            end;
          end;     

          if bFound then
          begin    
            if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_DBPtr^.dQty) then
            begin
              iCountMatch_DB := iCountMatch_DB + 2;
            end;

            ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr^.dQty;
            if aDailyAccount_DBPtr2 <> nil then
            begin
              ExcelApp.Cells[irow + 1, 13].Value := -aSAPMB51RecordPtr_match^.dqty;
              ExcelApp.Cells[irow + 1, 14].Value := -aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr2^.dQty;
            end;


            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;
          end
          else
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccount_DBPtr^.dQty;    
            ExcelApp.Cells[irow + 1, 13].Value := '0';
            ExcelApp.Cells[irow + 1, 14].Value := aDailyAccount_DBPtr^.dQty;
          end;

          irow := irow + 2;
        end;
      
      finally
        aSAPDailyAccountReader2_DB.Free;
      end;
    end; 
                  


    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    Memo1.Lines.Add('����');

    s := mmiDB_in_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB_in := TSAPDailyAccountReader2_DB_in_wt.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_DB_in.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);
//        AddColor(ExcelApp, irow, 6, irow, 6, clRed);


        irow := irow + 1;
        iCountDB_in_Fac := aSAPDailyAccountReader2_DB_in.Count;
        iCountMatch_DB_in := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB_in.Count - 1 do
        begin
          aDailyAccount_DB_inPtr := aSAPDailyAccountReader2_DB_in.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DB_inPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_DB_inPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DB_inPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_DB_inPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_DB_inPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_DB_inPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_DB_inPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_DB_inPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_DB_inPtr^.snumber_wt;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_DB_inPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_DB_inPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_DB_inPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_DB_inPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_DB_inPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_DB_inPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_DB_inPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_DB_inPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_DB_inPtr^.sstock_wt;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_DB_inPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_DB_inPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_DB_inPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_DB_inPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_DB_inPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_DB_inPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_DB_inPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_DB_inPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_DB_inPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_DB_inPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_DB_inPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_DB_inPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_DB_inPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_DB_inPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_DB_inPtr^.sstock_desc_wt;


          s_fac := aDailyAccount_DB_inPtr^.sbillno +
            aDailyAccount_DB_inPtr^.snumber;


          dQtyMatchx := 0;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr^.smovingtype <> '311' then Continue;

            if DoubleL( aSAPMB51RecordPtr^.dqty, 0 ) then Continue;

            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
                        
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
 
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber;

            if s_fac = s_mz then
            begin
              bFound := True;

              dQtyMatchx := dQtyMatchx + aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 13].Value := dQtyMatchx;
              ExcelApp.Cells[irow, 14].Value := dQtyMatchx - aDailyAccount_DB_inPtr^.dQty;

              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;

              if DoubleE( dQtyMatchx - aDailyAccount_DB_inPtr^.dQty, 0) then
              begin
                iCountMatch_DB_in := iCountMatch_DB_in + 1;
                Break;
              end;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccount_DB_inPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_DB_in.Free;
      end;
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

 
                               
    Memo1.Lines.Add('Ͷ�ϵ�');
        
    s := mmiPPBom_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + sfile_k3);       
    aSAPDailyAccountReader2_PPBom := TSAPDailyAccountReader2_PPBOM_wt.Create( sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_PPBom.Count > 0 then
    begin

      s2 := mmiSQ01PPBom.Caption;
      if Pos('(', s2) > 0 then
      begin
        s2 := Copy(s2, 1, Pos('(', s2) - 1);
      end;
      sfile_sq01_ppbom := vle_ml.Values[s2];
      Memo1.Lines.Add(s2);

      Memo1.Lines.Add('���ļ��� ' + sfile_sq01_ppbom);      
      aSAPDailyAccountReader2_coois := TSAPDailyAccountReader2_coois.Create(sfile_sq01_ppbom, 'Sheet1', aStockMZ2FacReader);

    
 
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '�ӹ�������';
        ExcelApp.Cells[irow, 2].Value := '��������';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�´�����';
        ExcelApp.Cells[irow, 6].Value := '�᰸����';
        ExcelApp.Cells[irow, 7].Value := '������ʼ����';
        ExcelApp.Cells[irow, 8].Value := '�����������';
        ExcelApp.Cells[irow, 9].Value := '�ƻ�����';
        ExcelApp.Cells[irow, 10].Value := '����ƻ�����';
        ExcelApp.Cells[irow, 11].Value := '��̩���Ϻ�';
        ExcelApp.Cells[irow, 12].Value := '�ͻ����Ϻ�';
        ExcelApp.Cells[irow, 13].Value := '������Ŀ��ʶ';
        ExcelApp.Cells[irow, 14].Value := '��������';
        ExcelApp.Cells[irow, 15].Value := '��Ŀ����';
        ExcelApp.Cells[irow, 16].Value := '��������';
        ExcelApp.Cells[irow, 17].Value := '��ע1';
        ExcelApp.Cells[irow, 18].Value := '�������';
        ExcelApp.Cells[irow, 19].Value := '����Ŀ';
        ExcelApp.Cells[irow, 20].Value := '��̩�����ϱ���';
        ExcelApp.Cells[irow, 21].Value := '�ͻ������ϱ���';
        ExcelApp.Cells[irow, 22].Value := '��������';
        ExcelApp.Cells[irow, 23].Value := '������';


        ExcelApp.Cells[irow, 26].Value := '��Ͷ������';
        ExcelApp.Cells[irow, 27].Value := '��λ';
        ExcelApp.Cells[irow, 28].Value := '���ǰ����';
        ExcelApp.Cells[irow, 29].Value := '�����';
        ExcelApp.Cells[irow, 30].Value := '���ȼ�';
        ExcelApp.Cells[irow, 31].Value := '�������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '������λ';
        ExcelApp.Cells[irow, 34].Value := '��ע2';
        ExcelApp.Cells[irow, 35].Value := '������';     
        ExcelApp.Cells[irow, 36].Value := '���幤����';


        irow := irow + 1;
        iCountPPBom := aSAPDailyAccountReader2_PPBom.Count;
        iCountMatch_PPBom := 0;
        iCountMatch_PPBom_mz := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_PPBom.Count - 1 do
        begin
          ptrDailyAccount_PPBOM := aSAPDailyAccountReader2_PPBom.Items[i_fac];
                                    

          ExcelApp.Cells[irow, 1].Value := ptrDailyAccount_PPBOM^.sfacname;
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_PPBOM^.sfac;
          ExcelApp.Cells[irow, 3].Value := ptrDailyAccount_PPBOM^.sicmobillno;
          ExcelApp.Cells[irow, 4].Value := ptrDailyAccount_PPBOM^.sicmotye;
          ExcelApp.Cells[irow, 5].Value := ptrDailyAccount_PPBOM^.dtRelease;
          if ptrDailyAccount_PPBOM^.dtClose <> 0 then
          begin
            ExcelApp.Cells[irow, 6].Value := ptrDailyAccount_PPBOM^.dtClose;
          end;
          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_PPBOM^.dtBegin;
          ExcelApp.Cells[irow, 8].Value := ptrDailyAccount_PPBOM^.dtEnd;
          ExcelApp.Cells[irow, 9].Value := ptrDailyAccount_PPBOM^.splanbillno;
          ExcelApp.Cells[irow, 10].Value := ptrDailyAccount_PPBOM^.splanbillno_mz;
          ExcelApp.Cells[irow, 11].Value := ptrDailyAccount_PPBOM^.snumber_wt;
          ExcelApp.Cells[irow, 12].Value := ptrDailyAccount_PPBOM^.snumber;
          ExcelApp.Cells[irow, 13].Value := ptrDailyAccount_PPBOM^.svItemFlag;
          ExcelApp.Cells[irow, 14].Value := ptrDailyAccount_PPBOM^.sname;
          ExcelApp.Cells[irow, 15].Value := ptrDailyAccount_PPBOM^.sItemCode;
          ExcelApp.Cells[irow, 16].Value := ptrDailyAccount_PPBOM^.dICMOQty;
          ExcelApp.Cells[irow, 17].Value := ptrDailyAccount_PPBOM^.snote1;
          ExcelApp.Cells[irow, 18].Value := ptrDailyAccount_PPBOM^.iChangeCount;
          ExcelApp.Cells[irow, 19].Value := ptrDailyAccount_PPBOM^.irowitem;
          ExcelApp.Cells[irow, 20].Value := ptrDailyAccount_PPBOM^.snumber_item_wt;
          ExcelApp.Cells[irow, 21].Value := ptrDailyAccount_PPBOM^.snumber_item;
          ExcelApp.Cells[irow, 22].Value := ptrDailyAccount_PPBOM^.sname_item;
          ExcelApp.Cells[irow, 23].Value := ptrDailyAccount_PPBOM^.dqtyplan;


          ExcelApp.Cells[irow, 26].Value := ptrDailyAccount_PPBOM^.dqtyout;
          ExcelApp.Cells[irow, 27].Value := ptrDailyAccount_PPBOM^.sstockname;
          ExcelApp.Cells[irow, 28].Value := ptrDailyAccount_PPBOM^.dqty0;
          ExcelApp.Cells[irow, 29].Value := ptrDailyAccount_PPBOM^.sgroup;
          ExcelApp.Cells[irow, 30].Value := ptrDailyAccount_PPBOM^.sprioriry;
          ExcelApp.Cells[irow, 31].Value := ptrDailyAccount_PPBOM^.dper;
          ExcelApp.Cells[irow, 32].Value := ptrDailyAccount_PPBOM^.dqtyshould;
          ExcelApp.Cells[irow, 33].Value := ptrDailyAccount_PPBOM^.sunit;
          ExcelApp.Cells[irow, 34].Value := ptrDailyAccount_PPBOM^.snote2;
          ExcelApp.Cells[irow, 35].Value := ptrDailyAccount_PPBOM^.schangelog;
                                                
          ExcelApp.Cells[irow, 36].Value := aICMO2FacReader2.ICMOFac2MZ(ptrDailyAccount_PPBOM^.sicmobillno);

          if DoubleE( ptrDailyAccount_PPBOM^.dqtyplan, 0 ) then
          begin                
            iCountMatch_PPBom := iCountMatch_PPBom + 1;
            irow := irow + 1;
            Continue;
          end;
          
          s_fac := ptrDailyAccount_PPBOM^.sicmobillno + ptrDailyAccount_PPBOM^.snumber_item;
                 
          ptrDailyAccount_coois_match := nil;
        
          bFound := False;
          for i_mz := 0 to aSAPDailyAccountReader2_coois.Count - 1 do
          begin
            ptrDailyAccount_coois := aSAPDailyAccountReader2_coois.Items[i_mz];      
            if ptrDailyAccount_coois^.bCalc then Continue;
          
            sbillno := ptrDailyAccount_coois^.sbillno_fac;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                   
            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end; 
          
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno + ptrDailyAccount_coois^.snumber_item;
            if s_fac = s_mz then
            begin
              bFound := True;
              if ptrDailyAccount_coois_match = nil then
              begin
                ptrDailyAccount_coois_match := ptrDailyAccount_coois;
              end
              else if Abs(ptrDailyAccount_coois_match^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan)
               > Abs(ptrDailyAccount_coois^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan) then
              begin
                ptrDailyAccount_coois_match := ptrDailyAccount_coois;
              end; 
            end;
          end;     

          if bFound then
          begin 
            ExcelApp.Cells[irow, 24].Value := ptrDailyAccount_coois_match^.dqtyneed;
            ExcelApp.Cells[irow, 25].Value := ptrDailyAccount_coois_match^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan;

            if DoubleE( ptrDailyAccount_coois_match^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan, 0) then
            begin
              iCountMatch_PPBom := iCountMatch_PPBom + 1;
            end;
            ptrDailyAccount_coois_match^.bCalc := True;
            ptrDailyAccount_coois_match^.sMatchType := s;
          end
          else
          begin
            if ptrDailyAccount_PPBOM^.dqtyplan > 0 then
            begin
              ExcelApp.Cells[irow, 24].Value := '0';
              ExcelApp.Cells[irow, 25].Value := - ptrDailyAccount_PPBOM^.dqtyplan;
            end
            else
            begin                                       
              iCountMatch_PPBom := iCountMatch_PPBom + 1;
              ExcelApp.Cells[irow, 24].Value := '0';
              ExcelApp.Cells[irow, 25].Value := '0';
            end;
          end;

          irow := irow + 1;
        end;

        for i_mz := 0 to aSAPDailyAccountReader2_coois.Count - 1 do
        begin
          ptrDailyAccount_coois := aSAPDailyAccountReader2_coois.Items[i_mz];
          if ptrDailyAccount_coois^.bCalc then Continue;

          ExcelApp.Cells[irow, 1].Value := ''; //ptrDailyAccount_coois.sfacname;
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_coois.sfac;
          ExcelApp.Cells[irow, 3].Value := ptrDailyAccount_coois.sbillno_fac;
          ExcelApp.Cells[irow, 4].Value := '';//ptrDailyAccount_coois.sicmotye;
          ExcelApp.Cells[irow, 5].Value := ''; // ptrDailyAccount_coois.dtRelease;
          if ptrDailyAccount_coois.dtFinish <> 0 then
          begin
            ExcelApp.Cells[irow, 6].Value := ptrDailyAccount_coois.dtFinish
          end;
//          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_coois.dtBegin;
//          ExcelApp.Cells[irow, 8].Value := ptrDailyAccount_coois.dtEnd;
//          ExcelApp.Cells[irow, 9].Value := ptrDailyAccount_coois.splanbillno;
//          ExcelApp.Cells[irow, 10].Value := ptrDailyAccount_coois.splanbillno_mz;
//          ExcelApp.Cells[irow, 11].Value := ptrDailyAccount_coois.snumber_wt;
          ExcelApp.Cells[irow, 12].Value := ptrDailyAccount_coois.snumber;
//          ExcelApp.Cells[irow, 13].Value := ptrDailyAccount_coois.svItemFlag;
//          ExcelApp.Cells[irow, 14].Value := ptrDailyAccount_coois.sname;
//          ExcelApp.Cells[irow, 15].Value := ptrDailyAccount_coois.sItemCode;
          ExcelApp.Cells[irow, 16].Value := ptrDailyAccount_coois.dqtyorder;
//          ExcelApp.Cells[irow, 17].Value := ptrDailyAccount_coois.snote1;
//          ExcelApp.Cells[irow, 18].Value := ptrDailyAccount_coois.iChangeCount;
//          ExcelApp.Cells[irow, 19].Value := ptrDailyAccount_coois.irowitem;
//          ExcelApp.Cells[irow, 20].Value := ptrDailyAccount_coois.snumber_item_wt;
          ExcelApp.Cells[irow, 21].Value := ptrDailyAccount_coois.snumber_item;
//          ExcelApp.Cells[irow, 22].Value := ptrDailyAccount_coois.sname_item;
          ExcelApp.Cells[irow, 23].Value := ptrDailyAccount_coois.dQtyIn;


//          ExcelApp.Cells[irow, 26].Value := ptrDailyAccount_coois.dqtyout;
          ExcelApp.Cells[irow, 27].Value := ptrDailyAccount_coois.sstockname;
//          ExcelApp.Cells[irow, 28].Value := ptrDailyAccount_coois.dqty0;
//          ExcelApp.Cells[irow, 29].Value := ptrDailyAccount_coois.sgroup;
//          ExcelApp.Cells[irow, 30].Value := ptrDailyAccount_coois.sprioriry;
//          ExcelApp.Cells[irow, 31].Value := ptrDailyAccount_coois.dper;
//          ExcelApp.Cells[irow, 32].Value := ptrDailyAccount_coois.dqtyshould;
//          ExcelApp.Cells[irow, 33].Value := ptrDailyAccount_coois.sunit;
//          ExcelApp.Cells[irow, 34].Value := ptrDailyAccount_coois.snote2;
//          ExcelApp.Cells[irow, 35].Value := ptrDailyAccount_coois.schangelog;
          ExcelApp.Cells[irow, 36].Value := 'mz';                             
          ExcelApp.Cells[irow, 37].Value := ptrDailyAccount_coois.sbillno;
          iCountMatch_PPBom_mz := iCountMatch_PPBom_mz + 1;

          irow := irow + 1;
        end;
        
      finally
        aSAPDailyAccountReader2_coois.Free; 
        aSAPDailyAccountReader2_PPBom.Free;
      end;

    end;        
        

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    Memo1.Lines.Add('��������');
                    
    s := mmiSOut_wt.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    

    Memo1.Lines.Add('���ļ��� ' + s);
            
    aSAPDailyAccountReader2_sout := TSAPDailyAccountReader2_sout_WT.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_sout.Count > 0 then
    begin
      try    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '������';
        ExcelApp.Cells[irow, 2].Value := '������';
        ExcelApp.Cells[irow, 3].Value := '����';
        ExcelApp.Cells[irow, 4].Value := '��Ʒ�Ϻ�';
        ExcelApp.Cells[irow, 5].Value := 'MZ';
        ExcelApp.Cells[irow, 6].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 7].Value := '��������';
        ExcelApp.Cells[irow, 8].Value := '��ע1';
        ExcelApp.Cells[irow, 9].Value := '��������';
        ExcelApp.Cells[irow, 10].Value := '�����Ϻ�';
        ExcelApp.Cells[irow, 11].Value := '��������';
        ExcelApp.Cells[irow, 12].Value := '��������';


        ExcelApp.Cells[irow, 15].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 16].Value := 'BOM����';
        ExcelApp.Cells[irow, 17].Value := '��ע2';
        ExcelApp.Cells[irow, 18].Value := '��������';
        ExcelApp.Cells[irow, 19].Value := '���ݱ��';
        ExcelApp.Cells[irow, 20].Value := '���幤����';
 

        irow := irow + 1;
        iCountSout_Fac := aSAPDailyAccountReader2_sout.Count;
        iCountMatch_Sout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_sout.Count - 1 do
        begin
          aDailyAccount_soutPtr := aSAPDailyAccountReader2_sout.Items[i_fac];


          ExcelApp.Cells[irow, 1].Value := aDailyAccount_soutPtr^.sfac;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_soutPtr^.sicmo;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_soutPtr^.dt;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_soutPtr^.snumber_wt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_soutPtr^.snumber;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_soutPtr^.sname;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_soutPtr^.dicmoqty;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_soutPtr^.snote1;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_soutPtr^.dqtyout;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_soutPtr^.snumber_child;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_soutPtr^.sname_child;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_soutPtr^.dqtyout;


          ExcelApp.Cells[irow, 15].Value := aDailyAccount_soutPtr^.sstock_wt;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_soutPtr^.sbomusage;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_soutPtr^.snote2;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_soutPtr^.sicmotype;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_soutPtr^.sbillno;
          ExcelApp.Cells[irow, 20].Value := aICMO2FacReader2.ICMOFac2MZ(aDailyAccount_soutPtr^.sicmo);
 


          s_fac := aDailyAccount_soutPtr^.snumber_child +
            aDailyAccount_soutPtr^.sbillno
            ; // + aDailyAccount_soutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if (aDailyAccount_soutPtr^.dqtyout > 0) and (aSAPMB51RecordPtr^.dqty < 0) then Continue;
            if (aDailyAccount_soutPtr^.dqtyout < 0) and (aSAPMB51RecordPtr^.dqty > 0) then Continue;

            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end; 
          
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_Sout := iCountMatch_Sout + 1;
              ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_soutPtr^.dqtyout;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;    
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := aDailyAccount_soutPtr^.dqtyout;
          end;

          irow := irow + 1;      
        end;
      finally
        aSAPDailyAccountReader2_sout.Free;
      end;
    end;             
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////                                          

    sl := TStringList.Create;
    try
      WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
      iSheet := iSheet + 1;
      ExcelApp.Sheets[iSheet].Activate;
      ExcelApp.Sheets[iSheet].Name := 'MB51';


      sline := '����ƾ֤'#9'ƾ֤����'#9'���ص�'#9'�ִ��ص������'#9'ƾ̧֤ͷ�ı�'#9'�ƶ�����'#9'���ϱ���'#9'��������'#9'��¼�뵥λ��ʾ������'#9'��������'#9'��������'#9'����ʱ��'#9'����'#9'�ɹ�����'#9'�Ƿ�ƥ��'#9'ƥ�䵥��'#9'���ϱ���'#9'��������';
      sl.Add(sline);

      for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
      begin
        aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
        sline := aSAPMB51RecordPtr^.sbillno + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.fdate) + #9
          + aSAPMB51RecordPtr^.fstockno + #9
          + aSAPMB51RecordPtr^.fstockname + #9
          + aSAPMB51RecordPtr^.fnote + #9
          + aSAPMB51RecordPtr^.smovingtype + #9                  
          + aSAPMB51RecordPtr^.snumber + #9
          + aSAPMB51RecordPtr^.sname + #9
          + FloatToStr(aSAPMB51RecordPtr^.dqty) + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.fdate) + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.finputdate) + #9
          + FormatDateTime('HH:mm:ss', aSAPMB51RecordPtr^.finputtime) + #9
          + aSAPMB51RecordPtr^.spo + #9
          + aSAPMB51RecordPtr^.sbillno_po + #9
          + CSBoolean[aSAPMB51RecordPtr^.bCalc] + #9
          + aSAPMB51RecordPtr^.sMatchType + #9
          + aSAPMB51RecordPtr^.snumber + #9
          + aSAPMB51RecordPtr^.sname;
        sl.Add(sline);
      end;

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;     
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select; 
                 
    finally
      sl.Free;
    end;

    
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Columns[1].ColumnWidth := 14.38;     
    ExcelApp.Columns[2].ColumnWidth := 21.63;
    ExcelApp.Columns[3].ColumnWidth := 13.63;
    ExcelApp.Columns[4].ColumnWidth := 12.38;
    ExcelApp.Columns[5].ColumnWidth := 16.50;
    ExcelApp.Columns[6].ColumnWidth := 15;
    ExcelApp.Columns[7].ColumnWidth := 21.50;
    ExcelApp.Columns[8].ColumnWidth := 78.75;

    irow := 1;
    
    AddHorizontalAlignment(ExcelApp, irow, 1, irow, 8, xlCenter);  
    AddHorizontalAlignment(ExcelApp, irow + 1, 1, irow + 12, 7, xlCenter);

    ExcelApp.Cells[irow, 1].Value := '����';
    ExcelApp.Cells[irow, 2].Value := '���嵥������';
    MergeCells(ExcelApp, irow, 2, irow, 3);
    ExcelApp.Cells[irow, 4].Value := '��̩�ᱨ����';
    ExcelApp.Cells[irow, 5].Value := 'SAP��ʽ����';
    ExcelApp.Cells[irow, 6].Value := '��̩��SAP����';
    ExcelApp.Cells[irow, 7].Value := '��ע';
    ExcelApp.Cells[irow, 8].Value := '���촦�����';

		AddColor(ExcelApp, irow, 1, irow, 8, $B7B8E6);
		AddColor(ExcelApp, irow, 6, irow, 7, $DCCD92);

    irow := 2;
    ExcelApp.Cells[irow, 1].Value := FormatDateTime('yyyy/MM/dd', Now);
    MergeCells(ExcelApp, irow, 1, irow + 11, 1);

    ExcelApp.Cells[irow, 2].Value := '�⹺��ⵥ';
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);
    ExcelApp.Cells[irow, 3].Value := 'PO����';
    ExcelApp.Cells[irow + 1, 3].Value := 'PO����';
    AddColor(ExcelApp, irow, 3, irow, 8, $DAC0CC);  
    AddColor(ExcelApp, irow + 1, 3, irow + 1, 8, $DEF1EB);

    ExcelApp.Cells[irow, 4].Value := iCountWinB_Fac; 
    ExcelApp.Cells[irow, 5].Value := iCountMatch_WinB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);
                           
    ExcelApp.Cells[irow + 1, 4].Value := iCountWinR_Fac;
    ExcelApp.Cells[irow + 1, 5].Value := iCountMatch_WinR;
    ExcelApp.Cells[irow + 1, 6].Value := '=D' + IntToStr(irow + 1) + '-E' + IntToStr(irow + 1);

    irow := irow + 2;
    ExcelApp.Cells[irow, 2].Value := '��Ʒ���';  
    ExcelApp.Cells[irow, 4].Value := iCountcpin_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_cpin;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    AddColor(ExcelApp, irow, 6, irow + 8, 7, $F3EEDA);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '������ⵥ';
    ExcelApp.Cells[irow, 3].Value := 'Sample';
    ExcelApp.Cells[irow, 4].Value := iCountqin_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_qin;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '�Ϻŵ���';
    ExcelApp.Cells[irow, 4].Value := iCountA2B_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_a2b;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '�������ɢ��';
    ExcelApp.Cells[irow, 4].Value := iCount03to01_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_03to01;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '���ϳ���';
    ExcelApp.Cells[irow, 4].Value := iCountqout_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_qout;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '���ε�';
    MergeCells(ExcelApp, irow, 2, irow + 2, 2);
    ExcelApp.Cells[irow, 3].Value := '�������ڲ���';
    ExcelApp.Cells[irow + 1, 3].Value := '���루��������';
    ExcelApp.Cells[irow + 2, 3].Value := '��������������';

    ExcelApp.Cells[irow, 4].Value := iCountDB_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_DB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    ExcelApp.Cells[irow + 1, 4].Value := iCountDB_in_Fac;
    ExcelApp.Cells[irow + 1, 5].Value := iCountMatch_DB_in;
    ExcelApp.Cells[irow + 1, 6].Value := '=D' + IntToStr(irow + 1) + '-E' + IntToStr(irow + 1);

    ExcelApp.Cells[irow + 2, 4].Value := iCountDB_Out_Fac;
    ExcelApp.Cells[irow + 2, 5].Value := iCountMatch_DB_out;
    ExcelApp.Cells[irow + 2, 6].Value := '=D' + IntToStr(irow + 2) + '-E' + IntToStr(irow + 2);

    AddColor(ExcelApp, irow + 1, 3, irow + 1, 8, $B4D5FC);   
    AddColor(ExcelApp, irow + 2, 3, irow + 2, 8, $9BD7C4);

    irow := irow + 3;
    ExcelApp.Cells[irow, 2].Value := '����Ͷ�ϵ�';
    ExcelApp.Cells[irow, 4].Value := iCountPPBom;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_PPBom;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ϵ�';      
    ExcelApp.Cells[irow, 4].Value := iCountSout_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_Sout;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

                     
    AddBorder(ExcelApp, 1, 1, 13, 8);
    
                



    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
    

  finally
    aSAPMB51Reader2.Free;
    aSAPCMSPushErrorReader2.Free;     
    aStockMZ2FacReader.Free;
    aICMO2FacReader2.Free;
  end;
         

  MessageBox(Handle, '���', '��ʾ', 0);
end;



procedure TfrmFacAccountCheck.btnSaveClick_yd;
const
  CSBoolean: array[Boolean] of string = ('��', '��');
var
  ExcelApp, WorkBook: Variant;
  aSAPMB51Reader2: TSAPMB51Reader2;
  aSAPCMSPushErrorReader2: TSAPCMSPushErrorReader2;
  iSheet: Integer;
  irow: Integer;
  sfile: string;   
  aStockMZ2FacReader: TStockMZ2FacReader;
  
  aSAPDailyAccountReader2_winB: TSAPDailyAccountReader2_winB;
  aSAPDailyAccountReader2_winR: TSAPDailyAccountReader2_winB;
  aSAPDailyAccountReader2_RTV: TSAPDailyAccountReader2_RTV;
  aSAPDailyAccountReader2_cpin: TSAPDailyAccountReader2_cpin;
  aSAPDailyAccountReader2_qin: TSAPDailyAccountReader2_qin;
  aSAPDailyAccountReader2_a2b: TSAPDailyAccountReader2_qout;
  aSAPDailyAccountReader2_03to01: TSAPDailyAccountReader2_03to01;
  aSAPDailyAccountReader2_qout: TSAPDailyAccountReader2_qout;
  aSAPDailyAccountReader2_DB: TSAPDailyAccountReader2_DB;      
  aSAPDailyAccountReader2_DB_in: TSAPDailyAccountReader2_DB_in;
  aSAPDailyAccountReader2_DB_out: TSAPDailyAccountReader2_DB_out;
  aSAPDailyAccountReader2_sout: TSAPDailyAccountReader2_sout;

  aSAPDailyAccountReader2_coois: TSAPDailyAccountReader2_coois;
  aSAPDailyAccountReader2_PPBom: TSAPDailyAccountReader2_PPBOM;
  aSAPDailyAccountReader2_PPBomChange_yd: TSAPDailyAccountReader2_PPBOMChange_yd;
  aSAPDailyAccountReader2_PPBomChange_mz: TSAPDailyAccountReader2_PPBOMChange_mz;

  i_fac: Integer;
  i_mz: Integer;
  s_fac: string;
  s_mz: string;
  bFound: Boolean;

  iCountWinB_Fac: Integer;      
  iCountWinR_Fac: Integer;
  iCountCPIN_Fac: Integer;
  iCountQIn_Fac: Integer;
  iCountA2B_Fac: Integer;     //�Ϻŵ���
  iCount03to01_Fac: Integer;  //�������ɢ��  
  iCountQout_Fac: Integer;    //���ϳ���
  iCountDB_Fac: Integer;
  iCountDB_in_Fac: Integer;
  iCountDB_out_Fac: Integer;
  iCountSout_Fac: Integer;  
  iCountPPBom: Integer;
  iCountPPBomChange: Integer;


  iCountMatch_WinB: Integer;
  iCountMatch_WinR: Integer;
  iCountMatch_cpin: Integer;
  iCountMatch_qin: Integer;
  iCountMatch_A2B: Integer;
  iCountMatch_03to01: Integer;
  iCountMatch_qout: Integer;
  iCountMatch_DB: Integer;
  iCountMatch_DB_out: Integer;
  iCountMatch_DB_in: Integer;
  iCountMatch_Sout: Integer;
  iCountMatch_PPBom: Integer;
  iCountMatch_PPBom_mz: Integer;  
  iCountMatch_PPBom_Change: Integer;

  aSAPMB51RecordPtr: PSAPMB51Record;
  aDailyAccount_winBPtr: PDailyAccount_winB;
  aDailyAccount_RTVPtr: PDailyAccount_RTV;
  aDailyAccount_cpinPtr: PDailyAccount_cpin;
  aDailyAccountqinPtr: PDailyAccount_qin;
  aDailyAccountqoutPtr: PDailyAccount_qout;
  aDailyAccount_DBPtr: PDailyAccount_DB;
  aDailyAccount_DBPtr2: PDailyAccount_DB;
  aDailyAccount_DB_inPtr: PDailyAccount_DB_in;
  aDailyAccount_DB_outPtr: PDailyAccount_DB_out;
  aDailyAccount_OutAInBCPtr: PDailyAccount_OutAInBC;
  aDailyAccount_soutPtr: PDailyAccount_sout;
  ptrDailyAccount_PPBOM: PDailyAccount_PPBom;
  ptrDailyAccount_coois: PDailyAccount_coois;
  
  ptrDailyAccount_PPBomChange_yd: PDailyAccount_PPBomChange_yd; 
  ptrDailyAccount_PPBomChange_mz: PDailyAccount_PPBomChange_mz;

  //aCPINmz2facReader: TCPINmz2facReader;
               
  aSAPMB51RecordPtr_match: PSAPMB51Record;
  
  s: string;
  s2: string;
  sfile_k3: string;                                                                

  sfile_sq01_ppbom: string;

  sbillno: string;
  idx: Integer;
  dDelta: Double;
  sl: TStringList;
  sline: string;

  dQtyMatchx: Double;
begin
  if not ExcelSaveDialog(sfile) then Exit;
                                                                        
  aSAPMB51Reader2 := TSAPMB51Reader2.Create(leMB51.Text, nil);
  aStockMZ2FacReader := TStockMZ2FacReader_yd.Create(leStockFac2MZ.Text);
//  aSAPCMSPushErrorReader2 := TSAPCMSPushErrorReader2.Create(leCMSErrMsg.Text, nil);

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

    Memo1.Lines.Add('����');

    WorkBook := ExcelApp.WorkBooks.Add;
    ExcelApp.DisplayAlerts := False;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate; 
    ExcelApp.Sheets[iSheet].Name := '����';
                  

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    s := mmiWinB_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winB := TSAPDailyAccountReader2_winB_yd.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_winB.Count > 0 then
    begin
      try
        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;


        ExcelApp.Cells[irow, 1].Value := '���ݱ��';
        ExcelApp.Cells[irow, 2].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 3].Value := '��������';
        ExcelApp.Cells[irow, 4].Value := 'ʵ������';
        ExcelApp.Cells[irow, 5].Value := 'SAP����';
        ExcelApp.Cells[irow, 6].Value := '����';
        ExcelApp.Cells[irow, 7].Value := '����';
        ExcelApp.Cells[irow, 8].Value := '�������';
        ExcelApp.Cells[irow, 9].Value := '��Ӧ��';
        ExcelApp.Cells[irow, 10].Value := '���ϲֿ�';   
        ExcelApp.Cells[irow, 11].Value := '���ϲֿ�����';
        ExcelApp.Cells[irow, 12].Value := '��ע';
        ExcelApp.Cells[irow, 13].Value := 'ժҪ';
        ExcelApp.Cells[irow, 14].Value := '��˱�־';
        ExcelApp.Cells[irow, 15].Value := '�Ƶ�';

        AddColor(ExcelApp, irow, 5, irow, 6, clYellow); 


        irow := irow + 1;
        iCountWinB_Fac := aSAPDailyAccountReader2_winB.Count;
        iCountMatch_WinB := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_winB.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winB.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sname;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dQty;

          //ExcelApp.Cells[irow, 5].Value := '';
          //ExcelApp.Cells[irow, 6].Value := '';

          ExcelApp.Cells[irow, 7].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.dtCheck;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.ssupplier;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.sstock_yd;
          ExcelApp.Cells[irow, 11].Value := '';
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.snote;  // �ɹ�������
          ExcelApp.Cells[irow, 13].Value := aDailyAccount_winBPtr^.ssummary;
          ExcelApp.Cells[irow, 14].Value := aDailyAccount_winBPtr^.scheckflag;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sbiller;

          s_fac := //myTrim(aDailyAccount_winBPtr^.sbillno) +  
            aDailyAccount_winBPtr^.snumber  +
            aDailyAccount_winBPtr^.snote;     // �ɹ�������


          dQtyMatchx := 0;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr^.smovingtype <> '101' then Continue;

            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'SY' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;
 
            s_mz := //sbillno +
              aSAPMB51RecordPtr^.snumber +
              aSAPMB51RecordPtr^.sbillno_po;// �ɹ�����

            if s_fac = s_mz then
            begin
              bFound := True;

              dQtyMatchx := dQtyMatchx + aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 13].Value := dQtyMatchx;
              ExcelApp.Cells[irow, 14].Value := dQtyMatchx - aDailyAccount_winBPtr^.dQty;

              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;

              if DoubleE( dQtyMatchx - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                iCountMatch_WinB := iCountMatch_WinB + 1;
                Break;
              end;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccount_winBPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winB.Free;
      end;
    end;
    
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

  (*
    s := mmiWinR_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
 
    (*
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_winR := TSAPDailyAccountReader2_winB_yd.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_winR.Count > 0 then
    begin
      try
    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);
//        AddColor(ExcelApp, irow, 6, irow, 6, clRed);

 
        irow := irow + 1;
        iCountWinR_Fac := aSAPDailyAccountReader2_winR.Count;
        iCountMatch_WinR := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_winR.Count - 1 do
        begin
          aDailyAccount_winBPtr := aSAPDailyAccountReader2_winR.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_winBPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_winBPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_winBPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_winBPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_winBPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_winBPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_winBPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_winBPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_winBPtr^.snumber_yd;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_winBPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_winBPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_winBPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_winBPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_winBPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_winBPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_winBPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_winBPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_winBPtr^.sstock_yd;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_winBPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_winBPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_winBPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_winBPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_winBPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_winBPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_winBPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_winBPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_winBPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_winBPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_winBPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_winBPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_winBPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_winBPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_winBPtr^.sstock_desc_yd;


          s_fac := aDailyAccount_winBPtr^.sbillno +
            aDailyAccount_winBPtr^.snumber +
            aDailyAccount_winBPtr^.sitemtext  ;       // �ɹ�����


          bFound := False;
          dQtyMatchx := 0;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
 
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber
              + aSAPMB51RecordPtr^.sbillno_po;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              dQtyMatchx := dQtyMatchx + aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 13].Value := dQtyMatchx;
              ExcelApp.Cells[irow, 14].Value := dQtyMatchx - aDailyAccount_winBPtr^.dQty;
              
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              
              if DoubleE(dQtyMatchx - aDailyAccount_winBPtr^.dQty, 0) then
              begin
                iCountMatch_WinR := iCountMatch_WinR + 1;
                Break;
              end;
            end;
          end;     

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := aDailyAccount_winBPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_winR.Free;
      end;
    end;          
     *)
         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    s := mmiCPIN_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);


    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_cpin := TSAPDailyAccountReader2_cpin_yd.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_cpin.Count > 0 then
    begin 
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '������';
        ExcelApp.Cells[irow, 2].Value := '������';
        ExcelApp.Cells[irow, 3].Value := '���ݱ��';
        ExcelApp.Cells[irow, 4].Value := '����';
        ExcelApp.Cells[irow, 5].Value := '��Ʒ�Ϻ�';
        ExcelApp.Cells[irow, 6].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 7].Value := '�������';       
        ExcelApp.Cells[irow, 8].Value := 'SAP����';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := '�ջ��ֿ�';
        ExcelApp.Cells[irow, 11].Value := '�����ջ��ֿ�';

        AddColor(ExcelApp, irow, 8, irow, 9, clYellow);


        irow := irow + 1;
        iCountCPIN_Fac := aSAPDailyAccountReader2_cpin.Count;
        iCountMatch_CPIN := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_cpin.Count - 1 do
        begin
          aDailyAccount_cpinPtr := aSAPDailyAccountReader2_cpin.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_cpinPtr^.sicmo;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_cpinPtr^.sfacname;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_cpinPtr^.sbillno;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_cpinPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_cpinPtr^.snumber;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_cpinPtr^.sname;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_cpinPtr^.dqty;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_cpinPtr^.sstock_yd;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_cpinPtr^.sstock;

          s_fac :=  myTrim( aDailyAccount_cpinPtr.sbillno ) +
            aDailyAccount_cpinPtr^.snumber +
            aDailyAccount_cpinPtr^.sstock;
                    
          bFound := False;
          dDelta := 9999999999;
          idx := -1;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr.bCalc then Continue;

            if (aSAPMB51RecordPtr^.smovingtype <> '101') and
              (aSAPMB51RecordPtr^.smovingtype <> '102') then
            begin
              Continue;
            end;                        

            if aSAPMB51RecordPtr^.fstockname = ''  then // Ҫ�вֿ�����
            begin
              Continue;
            end;

            sbillno := aSAPMB51RecordPtr^.snote_entry;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 2) = 'SY' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber + aSAPMB51RecordPtr^.fstockno;

            if s_fac = s_mz then
            begin
              bFound := True; 
              if dDelta > aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty then
              begin
                dDelta := Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty);
                idx := i_mz;
              end;
              if DoubleE(dDelta, 0) then Break;
            end;
          end;     

          if bFound then
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[idx];
            ExcelApp.Cells[irow, 8].Value := aSAPMB51RecordPtr^.dqty;
            ExcelApp.Cells[irow, 9].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_cpinPtr^.dQty;
            if DoubleE(dDelta, 0) then
            begin
              iCountMatch_CPIN := iCountMatch_CPIN + 1;
            end;
            aSAPMB51RecordPtr^.bCalc := True;
            aSAPMB51RecordPtr^.sMatchType := s;
            ExcelApp.Cells[irow, 36].Value := aSAPMB51RecordPtr^.sbillno_po;
          end
          else
          begin
            ExcelApp.Cells[irow, 8].Value := '0';
            ExcelApp.Cells[irow, 9].Value := - aDailyAccount_cpinPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_cpin.Free;
//        aCPINmz2facReader.Free;
      end;
    end;
           
 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    (*
    // ������ⵥ - Sample        ��Ʒ���               
    s := mmiQin_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
        
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_qin := TSAPDailyAccountReader2_qin_yd.Create(sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_qin.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);
//        AddColor(ExcelApp, irow, 6, irow, 6, clRed);


        irow := irow + 1;
        iCountQIn_Fac := aSAPDailyAccountReader2_qin.Count;
        iCountMatch_qin := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_qin.Count - 1 do
        begin
          aDailyAccountqinPtr := aSAPDailyAccountReader2_qin.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqinPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccountqinPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqinPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqinPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccountqinPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqinPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqinPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqinPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqinPtr^.snumber_yd;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqinPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqinPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqinPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqinPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqinPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqinPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqinPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqinPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccountqinPtr^.sstock_yd;
          ExcelApp.Cells[irow, 21].Value := aDailyAccountqinPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccountqinPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccountqinPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccountqinPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccountqinPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccountqinPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccountqinPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccountqinPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccountqinPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccountqinPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccountqinPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccountqinPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccountqinPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccountqinPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccountqinPtr^.sstock_desc_yd;


          s_fac := aDailyAccountqinPtr^.sbillno +
            aDailyAccountqinPtr^.snumber;


          dQtyMatchx := 0;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr^.smovingtype <> '511' then Continue;

            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
 
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber;

            if s_fac = s_mz then
            begin
              bFound := True;

              dQtyMatchx := dQtyMatchx + aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 13].Value := dQtyMatchx;
              ExcelApp.Cells[irow, 14].Value := dQtyMatchx - aDailyAccountqinPtr^.dQty;

              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;

              if DoubleE( dQtyMatchx - aDailyAccountqinPtr^.dQty, 0) then
              begin
                iCountMatch_qin := iCountMatch_qin + 1;
                Break;
              end;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccountqinPtr^.dQty;  
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_qin.Free;
      end;
    end;
   *)  
                     

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    // �Ϻŵ���
    s := mmiA2B_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
    (*
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_a2b := TSAPDailyAccountReader2_qout_yd.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_a2b.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;
        
        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);


        irow := irow + 1;
        iCountA2B_Fac := aSAPDailyAccountReader2_a2b.Count;
        iCountMatch_A2B := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_a2b.Count - 1 do
        begin
          aDailyAccountqoutPtr := aSAPDailyAccountReader2_a2b.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqoutPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccountqoutPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqoutPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqoutPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccountqoutPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqoutPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqoutPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqoutPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqoutPtr^.snumber_yd;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqoutPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqoutPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqoutPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqoutPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqoutPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqoutPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqoutPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqoutPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccountqoutPtr^.sstock_yd;
          ExcelApp.Cells[irow, 21].Value := aDailyAccountqoutPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccountqoutPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccountqoutPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccountqoutPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccountqoutPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccountqoutPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccountqoutPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccountqoutPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccountqoutPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccountqoutPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccountqoutPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccountqoutPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccountqoutPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccountqoutPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccountqoutPtr^.sstock_desc_yd;          

          s_fac := aDailyAccountqoutPtr^.snumber +
            aDailyAccountqoutPtr^.sbillno;
            ; // + aDailyAccountqoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_A2B := iCountMatch_A2B + 1;
              ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr^.dqty - aDailyAccountqoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccountqoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_a2b.Free;
      end;
    end; 
             
    *)                 

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('���ϳ���');
   (*
    s := mmiQout_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_qout := TSAPDailyAccountReader2_qout_yd.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_qout.Count > 0 then
    begin
      try

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);




        irow := irow + 1;
        iCountQout_Fac := aSAPDailyAccountReader2_qout.Count;
        iCountMatch_qout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_qout.Count - 1 do
        begin
          aDailyAccountqoutPtr := aSAPDailyAccountReader2_qout.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccountqoutPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccountqoutPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccountqoutPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccountqoutPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccountqoutPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccountqoutPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccountqoutPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccountqoutPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccountqoutPtr^.snumber_yd;
          ExcelApp.Cells[irow, 10].Value := aDailyAccountqoutPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccountqoutPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccountqoutPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccountqoutPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccountqoutPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccountqoutPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccountqoutPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccountqoutPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccountqoutPtr^.sstock_yd;
          ExcelApp.Cells[irow, 21].Value := aDailyAccountqoutPtr^.sstock;
          ExcelApp.Cells[irow, 22].Value := aDailyAccountqoutPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccountqoutPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccountqoutPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccountqoutPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccountqoutPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccountqoutPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccountqoutPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccountqoutPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccountqoutPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccountqoutPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccountqoutPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccountqoutPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccountqoutPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccountqoutPtr^.sstock_desc_yd;          

          s_fac := aDailyAccountqoutPtr^.snumber +
            aDailyAccountqoutPtr^.sbillno
            ; // + aDailyAccountqoutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];     
            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
                   
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_qout := iCountMatch_qout + 1;
              ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr^.dqty - aDailyAccountqoutPtr^.dQty;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccountqoutPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      
      finally
        aSAPDailyAccountReader2_qout.Free;
      end;
    end;         

     *)   


    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    (*
    Memo1.Lines.Add('����');
              
    s := mmiDB_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);
                                    
    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB := TSAPDailyAccountReader2_DB_yd.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_DB.Count > 0 then
    begin
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := '����';

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '��������';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '����ƾ֤';
        ExcelApp.Cells[irow, 4].Value := '��������';
        ExcelApp.Cells[irow, 5].Value := '�����̴���';
        ExcelApp.Cells[irow, 6].Value := '����������';
        ExcelApp.Cells[irow, 7].Value := '�ƶ�����';
        ExcelApp.Cells[irow, 8].Value := '�ƶ�ԭ��';
        ExcelApp.Cells[irow, 9].Value := '����';
        ExcelApp.Cells[irow, 10].Value := 'MZ';
        ExcelApp.Cells[irow, 11].Value := '����ͺ�';
        ExcelApp.Cells[irow, 12].Value := '��������';
                                                        
        ExcelApp.Cells[irow, 13].Value := 'SAP����';
        ExcelApp.Cells[irow, 14].Value := '����';
        
        ExcelApp.Cells[irow, 15].Value := '����������λ';
        ExcelApp.Cells[irow, 16].Value := 'ƾ̧֤ͷ�ı�';
        ExcelApp.Cells[irow, 17].Value := '������������';
        ExcelApp.Cells[irow, 18].Value := '��Ŀ�ı�';
        ExcelApp.Cells[irow, 19].Value := '������Ŀ��';
        ExcelApp.Cells[irow, 20].Value := '���ص�';
        ExcelApp.Cells[irow, 21].Value := 'MZ';
        ExcelApp.Cells[irow, 22].Value := '�������';
        ExcelApp.Cells[irow, 23].Value := '����������';
        ExcelApp.Cells[irow, 24].Value := '�ƶ�ԭ������';
        ExcelApp.Cells[irow, 25].Value := '������';
        ExcelApp.Cells[irow, 26].Value := '��������';
        ExcelApp.Cells[irow, 27].Value := '������������';
        ExcelApp.Cells[irow, 28].Value := '����ƾ֤��Ŀ';
        ExcelApp.Cells[irow, 29].Value := '�ƶ������ı�';
        ExcelApp.Cells[irow, 30].Value := '�춯״��';
        ExcelApp.Cells[irow, 31].Value := '��������';
        ExcelApp.Cells[irow, 32].Value := '��������';
        ExcelApp.Cells[irow, 33].Value := '����';
        ExcelApp.Cells[irow, 34].Value := '����������';
        ExcelApp.Cells[irow, 35].Value := '�ִ��ص������';

 
        AddColor(ExcelApp, irow, 13, irow, 14, clYellow);



        irow := irow + 1;
        iCountDB_Fac := aSAPDailyAccountReader2_DB.Count;
        iCountMatch_DB := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB.Count - 1 do
        begin
          aDailyAccount_DBPtr := aSAPDailyAccountReader2_DB.Items[i_fac];
          
          if aDailyAccount_DBPtr^.bCalc = True then Continue;

          aDailyAccount_DBPtr^.bCalc := True;

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DBPtr^.sfacname;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_DBPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DBPtr^.sdoc;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_DBPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_DBPtr^.smpn;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_DBPtr^.smpn_name;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_DBPtr^.smvt;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_DBPtr^.smvr;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_DBPtr^.snumber_yd;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_DBPtr^.snumber;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_DBPtr^.smodel;
          ExcelApp.Cells[irow, 12].Value := aDailyAccount_DBPtr^.dQty;

          
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_DBPtr^.sunit;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_DBPtr^.stext;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_DBPtr^.swc;
          ExcelApp.Cells[irow, 18].Value := aDailyAccount_DBPtr^.sitemtext;
          ExcelApp.Cells[irow, 19].Value := aDailyAccount_DBPtr^.sitemno;
          ExcelApp.Cells[irow, 20].Value := aDailyAccount_DBPtr^.sstock_yd;
          ExcelApp.Cells[irow, 21].Value := aDailyAccount_DBPtr^.sstock_desc;
          ExcelApp.Cells[irow, 22].Value := aDailyAccount_DBPtr^.sfacno;
          ExcelApp.Cells[irow, 23].Value := aDailyAccount_DBPtr^.sitemgroupname;
          ExcelApp.Cells[irow, 24].Value := aDailyAccount_DBPtr^.smvr_desc;
          ExcelApp.Cells[irow, 25].Value := aDailyAccount_DBPtr^.sitemgroup;
          ExcelApp.Cells[irow, 26].Value := aDailyAccount_DBPtr^.sordertype;
          ExcelApp.Cells[irow, 27].Value := aDailyAccount_DBPtr^.dicmoqty;
          ExcelApp.Cells[irow, 28].Value := aDailyAccount_DBPtr^.sdoc_item;
          ExcelApp.Cells[irow, 29].Value := aDailyAccount_DBPtr^.smvt_desc;
          ExcelApp.Cells[irow, 30].Value := aDailyAccount_DBPtr^.sstatus;
          ExcelApp.Cells[irow, 31].Value := aDailyAccount_DBPtr^.dtbill;
          ExcelApp.Cells[irow, 32].Value := aDailyAccount_DBPtr^.dbillqty;
          ExcelApp.Cells[irow, 33].Value := aDailyAccount_DBPtr^.sfac;
          ExcelApp.Cells[irow, 34].Value := aDailyAccount_DBPtr^.sicmo;
          ExcelApp.Cells[irow, 35].Value := aDailyAccount_DBPtr^.sstock_desc_yd;


                                      
          aDailyAccount_DBPtr2 := TSAPDailyAccountReader2_DB_yd(aSAPDailyAccountReader2_DB).GetItem2(aDailyAccount_DBPtr);
          if aDailyAccount_DBPtr2 <> nil then
          begin
            aDailyAccount_DBPtr2^.bCalc := True;

            ExcelApp.Cells[irow + 1, 1].Value := aDailyAccount_DBPtr2^.sfacname;
            ExcelApp.Cells[irow + 1, 2].Value :=  aDailyAccount_DBPtr2^.sbillno;
            ExcelApp.Cells[irow + 1, 3].Value := aDailyAccount_DBPtr2^.sdoc;
            ExcelApp.Cells[irow + 1, 4].Value := aDailyAccount_DBPtr2^.dt;
            ExcelApp.Cells[irow + 1, 5].Value := aDailyAccount_DBPtr2^.smpn;
            ExcelApp.Cells[irow + 1, 6].Value := aDailyAccount_DBPtr2^.smpn_name;
            ExcelApp.Cells[irow + 1, 7].Value := aDailyAccount_DBPtr2^.smvt;
            ExcelApp.Cells[irow + 1, 8].Value := aDailyAccount_DBPtr2^.smvr;
            ExcelApp.Cells[irow + 1, 9].Value := aDailyAccount_DBPtr2^.snumber_yd;
            ExcelApp.Cells[irow + 1, 10].Value := aDailyAccount_DBPtr2^.snumber;
            ExcelApp.Cells[irow + 1, 11].Value := aDailyAccount_DBPtr2^.smodel;
            ExcelApp.Cells[irow + 1, 12].Value := aDailyAccount_DBPtr2^.dQty;

          
            ExcelApp.Cells[irow + 1, 15].Value := aDailyAccount_DBPtr2^.sunit;
            ExcelApp.Cells[irow + 1, 16].Value := aDailyAccount_DBPtr2^.stext;
            ExcelApp.Cells[irow + 1, 17].Value := aDailyAccount_DBPtr2^.swc;
            ExcelApp.Cells[irow + 1, 18].Value := aDailyAccount_DBPtr2^.sitemtext;
            ExcelApp.Cells[irow + 1, 19].Value := aDailyAccount_DBPtr2^.sitemno;
            ExcelApp.Cells[irow + 1, 20].Value := aDailyAccount_DBPtr2^.sstock_yd;
            ExcelApp.Cells[irow + 1, 21].Value := aDailyAccount_DBPtr2^.sstock_desc;
            ExcelApp.Cells[irow + 1, 22].Value := aDailyAccount_DBPtr2^.sfacno;
            ExcelApp.Cells[irow + 1, 23].Value := aDailyAccount_DBPtr2^.sitemgroupname;
            ExcelApp.Cells[irow + 1, 24].Value := aDailyAccount_DBPtr2^.smvr_desc;
            ExcelApp.Cells[irow + 1, 25].Value := aDailyAccount_DBPtr2^.sitemgroup;
            ExcelApp.Cells[irow + 1, 26].Value := aDailyAccount_DBPtr2^.sordertype;
            ExcelApp.Cells[irow + 1, 27].Value := aDailyAccount_DBPtr2^.dicmoqty;
            ExcelApp.Cells[irow + 1, 28].Value := aDailyAccount_DBPtr2^.sdoc_item;
            ExcelApp.Cells[irow + 1, 29].Value := aDailyAccount_DBPtr2^.smvt_desc;
            ExcelApp.Cells[irow + 1, 30].Value := aDailyAccount_DBPtr2^.sstatus;
            ExcelApp.Cells[irow + 1, 31].Value := aDailyAccount_DBPtr2^.dtbill;
            ExcelApp.Cells[irow + 1, 32].Value := aDailyAccount_DBPtr2^.dbillqty;
            ExcelApp.Cells[irow + 1, 33].Value := aDailyAccount_DBPtr2^.sfac;
            ExcelApp.Cells[irow + 1, 34].Value := aDailyAccount_DBPtr2^.sicmo;
            ExcelApp.Cells[irow + 1, 35].Value := aDailyAccount_DBPtr2^.sstock_desc_yd;
                     

            if aDailyAccount_DBPtr^.sstock_desc = aDailyAccount_DBPtr2^.sstock_desc then // �����ֿ������ֿ��Ӧ����ͬһ���ֿ�
            begin
              ExcelApp.Cells[irow, 36].Value := aDailyAccount_DBPtr^.sstock_desc;
              ExcelApp.Cells[irow + 1, 36].Value := aDailyAccount_DBPtr2^.sstock_desc;
              iCountMatch_DB := iCountMatch_DB + 2;
              irow := irow + 2;
              Continue;
            end;
          end;

          s_fac := aDailyAccount_DBPtr^.sbillno +
            aDailyAccount_DBPtr^.snumber +
            aDailyAccount_DBPtr^.sstock_desc;

          aSAPMB51RecordPtr_match := nil;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];       
            if aSAPMB51RecordPtr^.bCalc then Continue;

            if aSAPMB51RecordPtr^.smovingtype <> '311' then Continue;

//            if aSAPMB51RecordPtr^.dqty < 0 then Continue; // ֻ��������

            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                    
            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end; 
          
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber +
              aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else if Abs(aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr^.dQty) > Abs(aSAPMB51RecordPtr^.dqty - aDailyAccount_DBPtr^.dQty) then
              begin                                     
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end;

              if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_DBPtr^.dQty) then
              begin
                Break;
              end;
            end;
          end;     

          if bFound then
          begin    
            if DoubleE(aSAPMB51RecordPtr_match^.dqty, aDailyAccount_DBPtr^.dQty) then
            begin
              iCountMatch_DB := iCountMatch_DB + 2;
            end;

            ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 14].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr^.dQty;
            if aDailyAccount_DBPtr2 <> nil then
            begin
              ExcelApp.Cells[irow + 1, 13].Value := -aSAPMB51RecordPtr_match^.dqty;
              ExcelApp.Cells[irow + 1, 14].Value := -aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DBPtr2^.dQty;
            end;


            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;
          end
          else
          begin
            ExcelApp.Cells[irow, 13].Value := '0';
            ExcelApp.Cells[irow, 14].Value := - aDailyAccount_DBPtr^.dQty;    
            ExcelApp.Cells[irow + 1, 13].Value := '0';
            ExcelApp.Cells[irow + 1, 14].Value := aDailyAccount_DBPtr^.dQty;
          end;

          irow := irow + 2;
        end;
      
      finally
        aSAPDailyAccountReader2_DB.Free;
      end;
    end; 
                  
     *)

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    Memo1.Lines.Add('�������');

    s := mmiDB_in_out_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + s);
    aSAPDailyAccountReader2_DB_in := TSAPDailyAccountReader2_DB_in_yd.Create(sfile_k3, '����', aStockMZ2FacReader);

    if aSAPDailyAccountReader2_DB_in.Count > 0 then
    begin
      try


        Memo1.Lines.Add(s);

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '����';
        ExcelApp.Cells[irow, 2].Value := '���ݱ��';
        ExcelApp.Cells[irow, 3].Value := '�����ֿ�';
        ExcelApp.Cells[irow, 4].Value := '����ֿ�';
        ExcelApp.Cells[irow, 5].Value := '���ϳ�����';
        ExcelApp.Cells[irow, 6].Value := '��������';
        ExcelApp.Cells[irow, 7].Value := '��������';
        ExcelApp.Cells[irow, 8].Value := 'SAP����';
        ExcelApp.Cells[irow, 9].Value := '����';
        
        AddColor(ExcelApp, irow, 8, irow, 9, clYellow);

        irow := irow + 1;
        iCountDB_in_Fac := aSAPDailyAccountReader2_DB_in.Count;
        iCountMatch_DB_in := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_DB_in.Count - 1 do
        begin
          aDailyAccount_DB_inPtr := aSAPDailyAccountReader2_DB_in.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := aDailyAccount_DB_inPtr^.dt;
          ExcelApp.Cells[irow, 2].Value :=  aDailyAccount_DB_inPtr^.sbillno;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_DB_inPtr^.sstockno_out_yd;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_DB_inPtr^.sstockno_in_yd;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_DB_inPtr^.snumber;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_DB_inPtr^.sname;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_DB_inPtr^.dQty;

          s_fac := myTrim(aDailyAccount_DB_inPtr^.sbillno) +
            aDailyAccount_DB_inPtr^.snumber +
            aDailyAccount_DB_inPtr^.sstockno_in;

          aSAPMB51RecordPtr_match := nil;
          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

            if aSAPMB51RecordPtr^.smovingtype <> '311' then Continue;

            if DoubleL( aSAPMB51RecordPtr^.dqty, 0 ) then Continue;

            if aSAPMB51RecordPtr.bCalc then Continue;

            sbillno := aSAPMB51RecordPtr^.fnote;
            sbillno := UpperCase(sbillno);
          
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'SYM' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end;
                  
            if Copy(sbillno, 1, 2) = 'SY' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end;

            s_mz := sbillno +
              aSAPMB51RecordPtr^.snumber +
              aSAPMB51RecordPtr^.fstockno;
              
            if s_fac = s_mz then
            begin        
              bFound := True;

              if aSAPMB51RecordPtr_match = nil then
              begin
                aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
              end
              else
              begin
                if Abs(aDailyAccount_DB_inPtr^.dQty - aSAPMB51RecordPtr_match^.dqty)
                  > Abs(aDailyAccount_DB_inPtr^.dQty - aSAPMB51RecordPtr^.dqty) then
                begin
                  aSAPMB51RecordPtr_match := aSAPMB51RecordPtr;
                end;
              end;

              if DoubleE( aSAPMB51RecordPtr_match.dqty - aDailyAccount_DB_inPtr^.dQty, 0) then
              begin
                Break;
              end;
            end;
          end;

          if bFound then
          begin
            ExcelApp.Cells[irow, 8].Value := aSAPMB51RecordPtr_match^.dqty;
            ExcelApp.Cells[irow, 9].Value := aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DB_inPtr^.dQty;

            aSAPMB51RecordPtr_match^.bCalc := True;
            aSAPMB51RecordPtr_match^.sMatchType := s;

            if DoubleE( aSAPMB51RecordPtr_match^.dqty - aDailyAccount_DB_inPtr^.dQty, 0) then
            begin
              iCountMatch_DB_in := iCountMatch_DB_in + 1;
            end;
          end
          else
          begin
            ExcelApp.Cells[irow, 8].Value := '0';
            ExcelApp.Cells[irow, 9].Value := aDailyAccount_DB_inPtr^.dQty;
          end;

          irow := irow + 1;
        end;
      finally
        aSAPDailyAccountReader2_DB_in.Free;
      end;
    end;
 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

 

    Memo1.Lines.Add('Ͷ�ϵ�');
        
    s := mmiPPBom_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + sfile_k3);       
    aSAPDailyAccountReader2_PPBom := TSAPDailyAccountReader2_PPBOM_yd.Create( sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_PPBom.Count > 0 then
    begin

      s2 := mmiSQ01PPBom.Caption;
      if Pos('(', s2) > 0 then
      begin
        s2 := Copy(s2, 1, Pos('(', s2) - 1);
      end;
      sfile_sq01_ppbom := vle_ml.Values[s2];
      Memo1.Lines.Add(s2);

      Memo1.Lines.Add('���ļ��� ' + sfile_sq01_ppbom);      
      aSAPDailyAccountReader2_coois := TSAPDailyAccountReader2_coois.Create(sfile_sq01_ppbom, 'Sheet1', aStockMZ2FacReader);

    
 
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '�Ƶ�����';
        ExcelApp.Cells[irow, 2].Value := '�������';
        ExcelApp.Cells[irow, 3].Value := '����/ί�ⶩ����';
        ExcelApp.Cells[irow, 4].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 5].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 6].Value := '��������';
        ExcelApp.Cells[irow, 7].Value := '����Ͷ�ϵ���';
        ExcelApp.Cells[irow, 8].Value := '�������ϳ�����';
        ExcelApp.Cells[irow, 9].Value := '������������';
        ExcelApp.Cells[irow, 10].Value := '�ƻ�Ͷ������';

        ExcelApp.Cells[irow, 13].Value := 'Ӧ������';
        ExcelApp.Cells[irow, 14].Value := '�ֿ�';
        ExcelApp.Cells[irow, 15].Value := '��λ����';
        ExcelApp.Cells[irow, 16].Value := '��˱�־';
        ExcelApp.Cells[irow, 17].Value := '��������';
 
        irow := irow + 1;
        iCountPPBom := aSAPDailyAccountReader2_PPBom.Count;
        iCountMatch_PPBom := 0;
        iCountMatch_PPBom_mz := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_PPBom.Count - 1 do
        begin
          ptrDailyAccount_PPBOM := aSAPDailyAccountReader2_PPBom.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := ptrDailyAccount_PPBOM^.dtdate;
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_PPBOM^.dtCheck;
          ExcelApp.Cells[irow, 3].Value := ptrDailyAccount_PPBOM^.sicmobillno;
          ExcelApp.Cells[irow, 4].Value := ptrDailyAccount_PPBOM^.snumber;
          ExcelApp.Cells[irow, 5].Value := ptrDailyAccount_PPBOM^.sname;
          ExcelApp.Cells[irow, 6].Value := ptrDailyAccount_PPBOM^.dqty;
          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_PPBOM^.sppbombillno;
          ExcelApp.Cells[irow, 8].Value := ptrDailyAccount_PPBOM^.snumber_item;
          ExcelApp.Cells[irow, 9].Value := ptrDailyAccount_PPBOM^.sname_item;
          ExcelApp.Cells[irow, 10].Value := ptrDailyAccount_PPBOM^.dqtyplan;

          ExcelApp.Cells[irow, 13].Value := ptrDailyAccount_PPBOM^.dqtyshould;
          ExcelApp.Cells[irow, 14].Value := ptrDailyAccount_PPBOM^.sstockname_yd;
          ExcelApp.Cells[irow, 15].Value := ptrDailyAccount_PPBOM^.dusage;
          ExcelApp.Cells[irow, 16].Value := ptrDailyAccount_PPBOM^.scheckflag;
          ExcelApp.Cells[irow, 17].Value := ptrDailyAccount_PPBOM^.sworkshopname;
 

          s_fac := ptrDailyAccount_PPBOM^.sppbombillno + ptrDailyAccount_PPBOM^.snumber_item;

          bFound := False;
          for i_mz := 0 to aSAPDailyAccountReader2_coois.Count - 1 do
          begin
            ptrDailyAccount_coois := aSAPDailyAccountReader2_coois.Items[i_mz];      
            if ptrDailyAccount_coois^.bCalc then Continue;
          
            sbillno := ptrDailyAccount_coois^.sbillno_fac;
//            idx := Pos('-', sbillno);
//            if idx > 0 then
//            begin
//              sbillno := Copy(sbillno, 1, idx - 1);
//            end;

            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                   
            if Copy(sbillno, 1, 3) = 'NWT' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end; 
                        
            if Copy(sbillno, 1, 2) = 'WT' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            if Copy(sbillno, 1, 2) = 'SY' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno + ptrDailyAccount_coois^.snumber_item;
            if s_fac = s_mz then
            begin                                              
              bFound := True;
              ExcelApp.Cells[irow, 11].Value := ptrDailyAccount_coois^.dqtyneed;
              ExcelApp.Cells[irow, 12].Value := ptrDailyAccount_coois^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan;
              if DoubleE( ptrDailyAccount_coois^.dqtyneed - ptrDailyAccount_PPBOM^.dqtyplan, 0) then
              begin
                iCountMatch_PPBom := iCountMatch_PPBom + 1;
              end;
              ptrDailyAccount_coois^.bCalc := True;
              ptrDailyAccount_coois^.sMatchType := s;
              Break;
            end;
          end;     

          if not bFound then
          begin
            if ptrDailyAccount_PPBOM^.dqtyplan > 0 then
            begin
              ExcelApp.Cells[irow, 11].Value := '0';
              ExcelApp.Cells[irow, 12].Value := - ptrDailyAccount_PPBOM^.dqtyplan;
            end
            else
            begin                                       
              iCountMatch_PPBom := iCountMatch_PPBom + 1;
              ExcelApp.Cells[irow, 11].Value := '0';
              ExcelApp.Cells[irow, 12].Value := '0';
            end;
          end;

          irow := irow + 1;
        end;

        for i_mz := 0 to aSAPDailyAccountReader2_coois.Count - 1 do
        begin
          ptrDailyAccount_coois := aSAPDailyAccountReader2_coois.Items[i_mz];
          if ptrDailyAccount_coois^.bCalc then Continue;

          ExcelApp.Cells[irow, 1].Value := ptrDailyAccount_coois^.dtfac;
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_coois^.dtfac;
          ExcelApp.Cells[irow, 3].Value := ptrDailyAccount_coois^.sbillno_fac;
          ExcelApp.Cells[irow, 4].Value := ptrDailyAccount_coois^.snumber;
          ExcelApp.Cells[irow, 5].Value := '';
          ExcelApp.Cells[irow, 6].Value := ptrDailyAccount_coois^.dqtyorder;
          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_coois^.sbillno_fac;
          ExcelApp.Cells[irow, 8].Value := ptrDailyAccount_coois^.snumber_item;
          ExcelApp.Cells[irow, 9].Value := '';
          ExcelApp.Cells[irow, 10].Value := ptrDailyAccount_coois^.dqtyneed;

          ExcelApp.Cells[irow, 13].Value := '';
          ExcelApp.Cells[irow, 14].Value := ptrDailyAccount_coois^.sstockname;
          ExcelApp.Cells[irow, 15].Value := '';
          ExcelApp.Cells[irow, 16].Value := '';
          ExcelApp.Cells[irow, 17].Value := '';

          iCountMatch_PPBom_mz := iCountMatch_PPBom_mz + 1;

          irow := irow + 1;
        end;
        
      finally
        aSAPDailyAccountReader2_coois.Free; 
        aSAPDailyAccountReader2_PPBom.Free;
      end;

    end;        

 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

 

    Memo1.Lines.Add('Ͷ�ϱ����');
        
    s := mmiPPBomChange_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    Memo1.Lines.Add(s);

    Memo1.Lines.Add('���ļ��� ' + sfile_k3);       
    aSAPDailyAccountReader2_PPBomChange_yd := TSAPDailyAccountReader2_PPBOMChange_yd.Create( sfile_k3, s, aStockMZ2FacReader);

    if aSAPDailyAccountReader2_PPBomChange_yd.Count > 0 then
    begin

      s2 := mmiSQ01PPBomChange_yd.Caption;
      if Pos('(', s2) > 0 then
      begin
        s2 := Copy(s2, 1, Pos('(', s2) - 1);
      end;
      sfile_sq01_ppbom := vle_ml.Values[s2];
      Memo1.Lines.Add(s2);

      Memo1.Lines.Add('���ļ��� ' + sfile_sq01_ppbom);      
      aSAPDailyAccountReader2_PPBomChange_mz := TSAPDailyAccountReader2_PPBOMChange_mz.Create( sfile_sq01_ppbom, s, aStockMZ2FacReader);

    
 
      try
        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '�����־';
        ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 3].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 4].Value := '����Ͷ�ϵ���';
        ExcelApp.Cells[irow, 5].Value := '���ϴ���';
        ExcelApp.Cells[irow, 6].Value := '��������';
        ExcelApp.Cells[irow, 7].Value := '��׼����';
        ExcelApp.Cells[irow, 8].Value := '�ֿ�';
        ExcelApp.Cells[irow, 9].Value := '���ԭ��';
        ExcelApp.Cells[irow, 10].Value := '�Ƶ�����';
        ExcelApp.Cells[irow, 11].Value := '�������';
        ExcelApp.Cells[irow, 12].Value := '������';
        ExcelApp.Cells[irow, 13].Value := '�ƻ�Ͷ������';
        ExcelApp.Cells[irow, 14].Value := 'SAP����';
        ExcelApp.Cells[irow, 15].Value := '����';
 
        irow := irow + 1;
        iCountPPBomChange := aSAPDailyAccountReader2_PPBomChange_yd.Count;
 
        iCountMatch_PPBom_Change := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_PPBomChange_yd.Count - 1 do
        begin
          ptrDailyAccount_PPBomChange_yd := aSAPDailyAccountReader2_PPBomChange_yd.Items[i_fac];

          ExcelApp.Cells[irow, 1].Value := ptrDailyAccount_PPBomChange_yd^.sChangeFlag;
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_PPBomChange_yd^.snumber;
          ExcelApp.Cells[irow, 3].Value := ptrDailyAccount_PPBomChange_yd^.sname;
          ExcelApp.Cells[irow, 4].Value := ptrDailyAccount_PPBomChange_yd^.sppbombillno;
          ExcelApp.Cells[irow, 5].Value := ptrDailyAccount_PPBomChange_yd^.snumber_item;
          ExcelApp.Cells[irow, 6].Value := ptrDailyAccount_PPBomChange_yd^.sname_item;
          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_PPBomChange_yd^.susage;
          ExcelApp.Cells[irow, 8].Value := ptrDailyAccount_PPBomChange_yd^.sstock_fac;
          ExcelApp.Cells[irow, 9].Value := ptrDailyAccount_PPBomChange_yd^.sChangeReason;
          ExcelApp.Cells[irow, 10].Value := ptrDailyAccount_PPBomChange_yd^.sdt;
          ExcelApp.Cells[irow, 11].Value := ptrDailyAccount_PPBomChange_yd^.sdtCheck;
          ExcelApp.Cells[irow, 12].Value := ptrDailyAccount_PPBomChange_yd^.sChangeVer;
          ExcelApp.Cells[irow, 13].Value := ptrDailyAccount_PPBomChange_yd^.dQty;
//          ExcelApp.Cells[irow, 14].Value := 'SAP����';
          ExcelApp.Cells[irow, 15].Value := '=' + GetRef(14) + IntToStr(irow) + '=' + GetRef(13) + IntToStr(irow);


          s_fac := ptrDailyAccount_PPBomChange_yd^.sppbombillno + '-' +
            ptrDailyAccount_PPBomChange_yd^.sChangeVer +
            ptrDailyAccount_PPBomChange_yd^.snumber_item;

          bFound := False;
          for i_mz := 0 to aSAPDailyAccountReader2_PPBomChange_mz.Count - 1 do
          begin
            ptrDailyAccount_PPBomChange_mz := aSAPDailyAccountReader2_PPBomChange_mz.Items[i_mz];
            if ptrDailyAccount_PPBomChange_mz^.bCalc then Continue;
          
            sbillno := ptrDailyAccount_PPBomChange_mz^.schangebillno;
//            idx := Pos('-', sbillno);
//            if idx > 0 then
//            begin
//              sbillno := Copy(sbillno, 1, idx - 1);
//            end;
//
//            idx := Pos('/', sbillno);
//            if idx > 0 then
//            begin
//              sbillno := Copy(sbillno, 1, idx - 1);
//            end;
//                   
//            if Copy(sbillno, 1, 3) = 'NWT' then
//            begin
//              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
//            end; 

            if Copy(sbillno, 1, 2) = 'SY' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := sbillno + ptrDailyAccount_PPBomChange_mz^.snumber_item;
            if s_fac = s_mz then
            begin                                              
              bFound := True;
              ExcelApp.Cells[irow, 14].Value := ptrDailyAccount_PPBomChange_mz^.dqty;
               
              iCountMatch_PPBom_Change := iCountMatch_PPBom_Change + 1;

              ptrDailyAccount_PPBomChange_mz^.bCalc := True;
              ptrDailyAccount_PPBomChange_mz^.sMatchType := s;
            end;
          end;
          
          irow := irow + 1;
        end;

        for i_mz := 0 to aSAPDailyAccountReader2_PPBomChange_mz.Count - 1 do
        begin
          ptrDailyAccount_PPBomChange_mz := aSAPDailyAccountReader2_PPBomChange_mz.Items[i_mz];
          if ptrDailyAccount_PPBomChange_mz^.bCalc then Continue;

          ExcelApp.Cells[irow, 1].Value := '';
          ExcelApp.Cells[irow, 2].Value := ptrDailyAccount_PPBomChange_mz^.snumber;
          ExcelApp.Cells[irow, 3].Value := '';
          ExcelApp.Cells[irow, 4].Value := ptrDailyAccount_PPBomChange_mz.schangebillno;
          ExcelApp.Cells[irow, 5].Value := ptrDailyAccount_PPBomChange_mz^.snumber_item;
          ExcelApp.Cells[irow, 6].Value := '';
          ExcelApp.Cells[irow, 7].Value := ptrDailyAccount_PPBomChange_mz^.sunit;
          ExcelApp.Cells[irow, 8].Value := '';
          ExcelApp.Cells[irow, 9].Value := '';
          ExcelApp.Cells[irow, 10].Value := '';

          ExcelApp.Cells[irow, 13].Value := '';
          ExcelApp.Cells[irow, 14].Value := '';
          ExcelApp.Cells[irow, 15].Value := '';
          ExcelApp.Cells[irow, 16].Value := ptrDailyAccount_PPBomChange_mz^.dqty;
          ExcelApp.Cells[irow, 17].Value := '';

          irow := irow + 1;
        end;
        
      finally
        aSAPDailyAccountReader2_PPBomChange_mz.Free;
        aSAPDailyAccountReader2_PPBomChange_yd.Free;
      end;

    end;
     
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    Memo1.Lines.Add('��������');
                    
    s := mmiSOut_yd.Caption;
    if Pos('(', s) > 0 then
    begin
      s := Copy(s, 1, Pos('(', s) - 1);
    end;
    sfile_k3 := vle_ml.Values[s];
    
   
    Memo1.Lines.Add('���ļ��� ' + s);
            
    aSAPDailyAccountReader2_sout := TSAPDailyAccountReader2_sout_yd.Create(sfile_k3, s, aStockMZ2FacReader);
    if aSAPDailyAccountReader2_sout.Count > 0 then
    begin
      try    

        WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
        iSheet := iSheet + 1;
        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := s;

        irow := 1;

        ExcelApp.Cells[irow, 1].Value := '������';
        ExcelApp.Cells[irow, 2].Value := '������';
        ExcelApp.Cells[irow, 3].Value := '���ݱ��';
        ExcelApp.Cells[irow, 4].Value := '����';
        ExcelApp.Cells[irow, 5].Value := '��Ʒ�Ϻ�';
        ExcelApp.Cells[irow, 6].Value := '��Ʒ����';
        ExcelApp.Cells[irow, 7].Value := '��������';
        ExcelApp.Cells[irow, 8].Value := '��������';
        ExcelApp.Cells[irow, 9].Value := '�����Ϻ�';
        ExcelApp.Cells[irow, 10].Value := '��������';
        ExcelApp.Cells[irow, 11].Value := '��������';
        
        ExcelApp.Cells[irow, 14].Value := '���ϲֿ�';
        ExcelApp.Cells[irow, 15].Value := '��λ����';
        ExcelApp.Cells[irow, 16].Value := '��ע�����Ⱥ�飩';
        ExcelApp.Cells[irow, 17].Value := '��������';
 

        irow := irow + 1;
        iCountSout_Fac := aSAPDailyAccountReader2_sout.Count;
        iCountMatch_Sout := 0;
        for i_fac := 0 to aSAPDailyAccountReader2_sout.Count - 1 do
        begin
          aDailyAccount_soutPtr := aSAPDailyAccountReader2_sout.Items[i_fac];


          ExcelApp.Cells[irow, 1].Value := aDailyAccount_soutPtr^.sicmo;
          ExcelApp.Cells[irow, 2].Value := aDailyAccount_soutPtr^.sfac;
          ExcelApp.Cells[irow, 3].Value := aDailyAccount_soutPtr^.sbillno;
          ExcelApp.Cells[irow, 4].Value := aDailyAccount_soutPtr^.dt;
          ExcelApp.Cells[irow, 5].Value := aDailyAccount_soutPtr^.snumber;
          ExcelApp.Cells[irow, 6].Value := aDailyAccount_soutPtr^.sname;
          ExcelApp.Cells[irow, 7].Value := aDailyAccount_soutPtr^.dqty;
          ExcelApp.Cells[irow, 8].Value := aDailyAccount_soutPtr^.dqtyout;
          ExcelApp.Cells[irow, 9].Value := aDailyAccount_soutPtr^.snumber_child;
          ExcelApp.Cells[irow, 10].Value := aDailyAccount_soutPtr^.sname_child;
          ExcelApp.Cells[irow, 11].Value := aDailyAccount_soutPtr^.dqtyout;

          ExcelApp.Cells[irow, 14].Value := aDailyAccount_soutPtr^.sstock_yd;
          ExcelApp.Cells[irow, 15].Value := aDailyAccount_soutPtr^.dusage;
          ExcelApp.Cells[irow, 16].Value := aDailyAccount_soutPtr^.snote;
          ExcelApp.Cells[irow, 17].Value := aDailyAccount_soutPtr^.sicmotype;
  
          s_fac := aDailyAccount_soutPtr^.snumber_child +
            myTrim(aDailyAccount_soutPtr^.sbillno)
            ; // + aDailyAccount_soutPtr^.sstock;

          bFound := False;
          for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
          begin
            aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];

//            if (aDailyAccount_soutPtr^.dqtyout > 0) and (aSAPMB51RecordPtr^.dqty < 0) then Continue;
//            if (aDailyAccount_soutPtr^.dqtyout < 0) and (aSAPMB51RecordPtr^.dqty > 0) then Continue;

            if aSAPMB51RecordPtr^.bCalc then Continue;
          
            sbillno := aSAPMB51RecordPtr^.fnote;
            idx := Pos('-', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;
                 
            idx := Pos('/', sbillno);
            if idx > 0 then
            begin
              sbillno := Copy(sbillno, 1, idx - 1);
            end;

            if Copy(sbillno, 1, 3) = 'SYM' then
            begin
              sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
            end; 
          
            if Copy(sbillno, 1, 2) = 'SY' then
            begin
              sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
            end; 
          
            s_mz := aSAPMB51RecordPtr^.snumber +
              sbillno
              ; // + aSAPMB51RecordPtr^.fstockname;

            if s_fac = s_mz then
            begin                                              
              bFound := True;
              iCountMatch_Sout := iCountMatch_Sout + 1;
              ExcelApp.Cells[irow, 12].Value := aSAPMB51RecordPtr^.dqty;
              ExcelApp.Cells[irow, 13].Value := aSAPMB51RecordPtr^.dqty - aDailyAccount_soutPtr^.dqtyout;
              aSAPMB51RecordPtr^.bCalc := True;
              aSAPMB51RecordPtr^.sMatchType := s;
//              ExcelApp.Cells[irow, 20].Value := aDailyAccount_soutPtr^.sicmo;
              Break;
            end;
          end;

          if not bFound then
          begin
            ExcelApp.Cells[irow, 12].Value := '0';
            ExcelApp.Cells[irow, 13].Value := aDailyAccount_soutPtr^.dqtyout;
          end;

          irow := irow + 1;      
        end;
      finally
        aSAPDailyAccountReader2_sout.Free;
      end;
    end;

    (*  
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////                                          

    sl := TStringList.Create;
    try
      WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
      iSheet := iSheet + 1;
      ExcelApp.Sheets[iSheet].Activate;
      ExcelApp.Sheets[iSheet].Name := 'MB51';


      sline := '����ƾ֤'#9'ƾ֤����'#9'���ص�'#9'�ִ��ص������'#9'ƾ̧֤ͷ�ı�'#9'�ƶ�����'#9'���ϱ���'#9'��������'#9'��¼�뵥λ��ʾ������'#9'��������'#9'��������'#9'����ʱ��'#9'����'#9'�ɹ�����'#9'�Ƿ�ƥ��'#9'ƥ�䵥��'#9'���ϱ���'#9'��������';
      sl.Add(sline);

      for i_mz := 0 to aSAPMB51Reader2.Count - 1 do
      begin
        aSAPMB51RecordPtr := aSAPMB51Reader2.Items[i_mz];
        sline := aSAPMB51RecordPtr^.sbillno + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.fdate) + #9
          + aSAPMB51RecordPtr^.fstockno + #9
          + aSAPMB51RecordPtr^.fstockname + #9
          + aSAPMB51RecordPtr^.fnote + #9
          + aSAPMB51RecordPtr^.smovingtype + #9      
          + aSAPMB51RecordPtr^.snumber + #9
          + aSAPMB51RecordPtr^.sname + #9
          + FloatToStr(aSAPMB51RecordPtr^.dqty) + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.fdate) + #9
          + FormatDateTime('yyyy-MM-dd', aSAPMB51RecordPtr^.finputdate) + #9
          + FormatDateTime('HH:mm:ss', aSAPMB51RecordPtr^.finputtime) + #9
          + aSAPMB51RecordPtr^.spo + #9
          + aSAPMB51RecordPtr^.sbillno_po + #9
          + CSBoolean[aSAPMB51RecordPtr^.bCalc] + #9
          + aSAPMB51RecordPtr^.sMatchType + #9
          + aSAPMB51RecordPtr^.snumber + #9
          + aSAPMB51RecordPtr^.sname;
        sl.Add(sline);
      end;

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;     
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select; 
                 
    finally
      sl.Free;
    end;
     *)
    
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Columns[1].ColumnWidth := 14.38;
    ExcelApp.Columns[2].ColumnWidth := 21.63;
    ExcelApp.Columns[3].ColumnWidth := 13.63;
    ExcelApp.Columns[4].ColumnWidth := 12.38;
    ExcelApp.Columns[5].ColumnWidth := 16.50;
    ExcelApp.Columns[6].ColumnWidth := 15;
    ExcelApp.Columns[7].ColumnWidth := 21.50;
    ExcelApp.Columns[8].ColumnWidth := 78.75;

    irow := 1;
    
    AddHorizontalAlignment(ExcelApp, irow, 1, irow, 8, xlCenter);  
    AddHorizontalAlignment(ExcelApp, irow + 1, 1, irow + 13, 7, xlCenter);

    ExcelApp.Cells[irow, 1].Value := '����';
    ExcelApp.Cells[irow, 2].Value := '���嵥������';
    MergeCells(ExcelApp, irow, 2, irow, 3);
    ExcelApp.Cells[irow, 4].Value := '����ᱨ����';
    ExcelApp.Cells[irow, 5].Value := 'SAP��ʽ����';
    ExcelApp.Cells[irow, 6].Value := '�����SAP����';
    ExcelApp.Cells[irow, 7].Value := '��ע';
    ExcelApp.Cells[irow, 8].Value := '���촦�����';

		AddColor(ExcelApp, irow, 1, irow, 8, $B7B8E6);
		AddColor(ExcelApp, irow, 6, irow, 7, $DCCD92);

    irow := 2;
    ExcelApp.Cells[irow, 1].Value := FormatDateTime('yyyy/MM/dd', Now);
    MergeCells(ExcelApp, irow, 1, irow + 12, 1);

    ExcelApp.Cells[irow, 2].Value := '�⹺��ⵥ';
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);
    ExcelApp.Cells[irow, 3].Value := 'PO����';
    ExcelApp.Cells[irow + 1, 3].Value := 'PO����';
    AddColor(ExcelApp, irow, 3, irow, 8, $DAC0CC);  
    AddColor(ExcelApp, irow + 1, 3, irow + 1, 8, $DEF1EB);

    ExcelApp.Cells[irow, 4].Value := iCountWinB_Fac; 
    ExcelApp.Cells[irow, 5].Value := iCountMatch_WinB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);
                           
    ExcelApp.Cells[irow + 1, 4].Value := iCountWinR_Fac;
    ExcelApp.Cells[irow + 1, 5].Value := iCountMatch_WinR;
    ExcelApp.Cells[irow + 1, 6].Value := '=D' + IntToStr(irow + 1) + '-E' + IntToStr(irow + 1);

    irow := irow + 2;
    ExcelApp.Cells[irow, 2].Value := '��Ʒ���';  
    ExcelApp.Cells[irow, 4].Value := iCountcpin_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_cpin;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    AddColor(ExcelApp, irow, 6, irow + 8, 7, $F3EEDA);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '������ⵥ';
    ExcelApp.Cells[irow, 3].Value := 'Sample';
    ExcelApp.Cells[irow, 4].Value := iCountqin_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_qin;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '�Ϻŵ���';
    ExcelApp.Cells[irow, 4].Value := iCountA2B_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_a2b;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '�������ɢ��';
    ExcelApp.Cells[irow, 4].Value := iCount03to01_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_03to01;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ⵥ';
    ExcelApp.Cells[irow, 3].Value := '���ϳ���';
    ExcelApp.Cells[irow, 4].Value := iCountqout_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_qout;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '���ε�';
    MergeCells(ExcelApp, irow, 2, irow + 2, 2);
    ExcelApp.Cells[irow, 3].Value := '�������ڲ���';
    ExcelApp.Cells[irow + 1, 3].Value := '���루��������';
    ExcelApp.Cells[irow + 2, 3].Value := '��������������';

    ExcelApp.Cells[irow, 4].Value := iCountDB_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_DB;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    ExcelApp.Cells[irow + 1, 4].Value := iCountDB_in_Fac;
    ExcelApp.Cells[irow + 1, 5].Value := iCountMatch_DB_in;
    ExcelApp.Cells[irow + 1, 6].Value := '=D' + IntToStr(irow + 1) + '-E' + IntToStr(irow + 1);

    ExcelApp.Cells[irow + 2, 4].Value := iCountDB_Out_Fac;
    ExcelApp.Cells[irow + 2, 5].Value := iCountMatch_DB_out;
    ExcelApp.Cells[irow + 2, 6].Value := '=D' + IntToStr(irow + 2) + '-E' + IntToStr(irow + 2);

    AddColor(ExcelApp, irow + 1, 3, irow + 1, 8, $B4D5FC);   
    AddColor(ExcelApp, irow + 2, 3, irow + 2, 8, $9BD7C4);

    irow := irow + 3;
    ExcelApp.Cells[irow, 2].Value := '����Ͷ�ϵ�';
    ExcelApp.Cells[irow, 4].Value := iCountPPBom;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_PPBom;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := '�������ϵ�';
    ExcelApp.Cells[irow, 4].Value := iCountSout_Fac;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_Sout;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);

    irow := irow + 1;
    ExcelApp.Cells[irow, 2].Value := 'Ͷ�ϱ����';
    ExcelApp.Cells[irow, 4].Value := iCountPPBomChange;
    ExcelApp.Cells[irow, 5].Value := iCountMatch_PPBom_Change;
    ExcelApp.Cells[irow, 6].Value := '=D' + IntToStr(irow) + '-E' + IntToStr(irow);


    AddBorder(ExcelApp, 1, 1, 14, 8);
    
                



    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
    

  finally
    aSAPMB51Reader2.Free;
    aSAPCMSPushErrorReader2.Free;     
    aStockMZ2FacReader.Free;
  end;
         

  MessageBox(Handle, '���', '��ʾ', 0);
end;

procedure TfrmFacAccountCheck.Button1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leICMO2Fac.Text := sfile;
end;

end.
