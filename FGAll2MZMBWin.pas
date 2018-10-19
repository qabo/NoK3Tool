unit FGAll2MZMBWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, ProjNameNoWin, CommUtils, StdCtrls,
  ExtCtrls, IniFiles, ComObj;

type
  TfrmFGAll2MZMB = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    tbProjNameNo: TToolButton;
    ToolButton7: TToolButton;
    btnExit: TToolButton;
    ToolButton5: TToolButton;
    leCPIN: TLabeledEdit;
    leFGStock: TLabeledEdit;
    btnCPIN: TButton;
    btnFGStock: TButton;
    btnCPIN_s: TButton;
    btnFGStock_s: TButton;
    Memo1: TMemo;
    procedure tbProjNameNoClick(Sender: TObject);
    procedure btnCPINClick(Sender: TObject);
    procedure btnFGStockClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnCPIN_sClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnFGStock_sClick(Sender: TObject);
  private
    { Private declarations }
    procedure Log(const str: string);
    procedure SaveCPIN(const sfile_save, sfile_data: string;
      slProjNo: TStringList);
    procedure SaveFGStock(const sfile_save, sfile_data: string;
      slProjNo: TStringList);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

type
  TCPINSumRecord = packed record
    dt: string; //�������
    dtCheck: string; //	�������
    snumber: string; //	�Ϻ�
    sname: string; //	��Ʒ����
    dqty: string;	 // ����
    sfac: string; // ������
    sbatchno: string; //	����
    snote: string; //	��ע
    sstock: string; //	�ջ��ֿ�
    sicmo: string; //	������
    sbillno: string; //	���ݱ��
    snumber_fac: string;	// �����Ϻ�
    bOk: Boolean;
  end;
  PCPINSumRecord = ^TCPINSumRecord;
             
  TCPINRecord = packed record
    dt: string; //�������
    snumber: string; //	�Ϻ�
    sname: string; //	��Ʒ����
    dqty: string; //	 ����
    sfac: string; // 	������
    sbatchno: string; //	����
    snote: string; //	��ע
    sstock: string; //	�ջ��ֿ�
    sicmo: string; //	������
    sbillno: string; //	���ݱ��
    snumber_fac: string; //	�����Ϻ�
    bOk: Boolean;
  end;
  PCPINRecord = ^TCPINRecord;
               
  TUnOutRecord = packed record
    snumber: string; //�Ϻ�
    sname: string; //	��Ʒ����
    dqty: string; //	����
    snote: string; //	��ע
    sqty_mat:string; //	���Ͻ��
    bOk: Boolean;
  end;
  PUnOutRecord = ^TUnOutRecord;



  TFGStockRecord = packed record
    snumber: string; //�Ϻ�
    sname: string; //��Ʒ����
    dqty: string; //�������
    dqty_rework: string; //	����
    dqty_uncheck: string; //	������(��3���£�
    saddr: string; //		����ص�
    sbatchno: string; //	����
    snote: string; //	��ע
    dqty_ok: string; //	�ɷ������
    sncard: string; //	�±���
    dtEnd: string; //	��ֹ����
    bOk: Boolean;
  end;
  PFGStockRecord = ^TFGStockRecord;

class procedure TfrmFGAll2MZMB.ShowForm;
var
  frmFGAll2MZMB: TfrmFGAll2MZMB;
begin
  frmFGAll2MZMB := TfrmFGAll2MZMB.Create(nil);
  frmFGAll2MZMB.ShowModal;
  frmFGAll2MZMB.Free;
end;

procedure TfrmFGAll2MZMB.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leCPIN.Text := ini.ReadString(self.ClassName, leCPIN.Name, '');
    leFGStock.Text := ini.ReadString(self.ClassName, leFGStock.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmFGAll2MZMB.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leCPIN.Name, leCPIN.Text);
    ini.WriteString(self.ClassName, leFGStock.Name, leFGStock.Text);
  finally
    ini.Free;
  end;
end;
        
procedure TfrmFGAll2MZMB.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmFGAll2MZMB.btnCPINClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leCPIN.Text := sfile;
end;

procedure TfrmFGAll2MZMB.btnFGStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leFGStock.Text := sfile;
end;
      
procedure TfrmFGAll2MZMB.tbProjNameNoClick(Sender: TObject);
begin
  TfrmProjNameNo.ShowForm;
end;

procedure TfrmFGAll2MZMB.Log(const str: string);
begin
  Memo1.Lines.Add(str);
end;

function NumberInList(const snumber: string; sl: TStringList): Boolean;
var
  i: Integer;
  s: string;
begin
  Result := False;
  for i := 0 to sl.Count - 1 do
  begin
    s := Trim(sl.Names[i]);
    if s = '' then Continue;

    if UpperCase(Copy(snumber, 1, Length(s))) = UpperCase(s) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

procedure TfrmFGAll2MZMB.btnCPIN_sClick(Sender: TObject);
var
  sfile: string;       
  slProjNo_OEM: TStringList;
  slProjNo_ODM: TStringList;
  s: string;
begin                                  
  if not ExcelSaveDialog(sfile) then Exit;
      
  slProjNo_OEM := TfrmProjNameNo.GetProjNos_OEM;
  slProjNo_ODM := TfrmProjNameNo.GetProjNos_ODM;

  s := ChangeFileExt(sfile, '') + '-����' + ExtractFileExt(sfile);
  SaveCPIN(s, leCPIN.Text, slProjNo_OEM);
  s := ChangeFileExt(sfile, '') + '-����' + ExtractFileExt(sfile);
  SaveCPIN(s, leCPIN.Text, slProjNo_ODM);
   
  slProjNo_OEM.Free;
  slProjNo_ODM.Free;

  MessageBox(Handle, '���', '��ʾ', 0);
end;

procedure TfrmFGAll2MZMB.SaveCPIN(const sfile_save, sfile_data: string;
  slProjNo: TStringList);
var
  ExcelApp, WorkBook: Variant;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  s: string;
  s1, s2, s3, s4, s5, s6: string;
  irow: Integer;
  lst: TList;  
  iCount: Integer;
  aCPINRecordPtr: PCPINSumRecord;
  snumber: string;
begin

  lst := TList.Create;            
  
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(sfile_data);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        Log(sSheet);

        irow := 1;
        s1 := ExcelApp.Cells[irow, 1].Value;   
        s2 := ExcelApp.Cells[irow, 2].Value;
        s3 := ExcelApp.Cells[irow, 3].Value;
        s4 := ExcelApp.Cells[irow, 4].Value;
        s5 := ExcelApp.Cells[irow, 5].Value;
        s6 := ExcelApp.Cells[irow, 6].Value;
        s := s1 + s2 + s3 + s4 + s5 + s6;

        if s <> '���������������ϺŲ�Ʒ��������������' then
        begin
          Log('sheet ' + sSheet + ' ��ʽ����');
          Log('��ȷ:���������������ϺŲ�Ʒ��������������');
          Log('�ļ�:' +s);
          Continue;
        end;

        irow := irow + 1;
        snumber := ExcelApp.Cells[irow, 3].Value;
        while snumber <> '' do
        begin
          aCPINRecordPtr := New(PCPINSumRecord);
          aCPINRecordPtr.bOk := False;
          aCPINRecordPtr^.dt := ExcelApp.Cells[irow, 1].Value; // �������
          aCPINRecordPtr^.dtCheck := ExcelApp.Cells[irow, 2].Value; // �������
          aCPINRecordPtr^.snumber := ExcelApp.Cells[irow, 3].Value; // �Ϻ�
          aCPINRecordPtr^.sname := ExcelApp.Cells[irow, 4].Value; // ��Ʒ����
          aCPINRecordPtr^.dqty := ExcelApp.Cells[irow, 5].Value; // ����
          aCPINRecordPtr^.sfac := ExcelApp.Cells[irow, 6].Value; // ������
          aCPINRecordPtr^.sbatchno := ExcelApp.Cells[irow, 7].Value; // ����
          aCPINRecordPtr^.snote := ExcelApp.Cells[irow, 8].Value; // ��ע
          aCPINRecordPtr^.sstock := ExcelApp.Cells[irow, 9].Value; // �ջ��ֿ�
          aCPINRecordPtr^.sicmo := ExcelApp.Cells[irow, 10].Value; // ������
          aCPINRecordPtr^.sbillno := ExcelApp.Cells[irow, 11].Value; // ���ݱ��
          aCPINRecordPtr^.snumber_fac := ExcelApp.Cells[irow, 12].Value; // �����Ϻ�
          lst.Add(aCPINRecordPtr);

          irow := irow + 1;   
          snumber := ExcelApp.Cells[irow, 3].Value;
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

  ///////////////////////////////////////////////////////////////////////

       
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
                     
  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := '��Ʒ������';
  irow := 1;
  ExcelApp.Cells[irow, 1].Value := '�������';
  ExcelApp.Cells[irow, 2].Value := '�������';
  ExcelApp.Cells[irow, 3].Value := '�Ϻ�';
  ExcelApp.Cells[irow, 4].Value := '��Ʒ����';
  ExcelApp.Cells[irow, 5].Value := '����';
  ExcelApp.Cells[irow, 6].Value := '������';
  ExcelApp.Cells[irow, 7].Value := '����';
  ExcelApp.Cells[irow, 8].Value := '��ע';
  ExcelApp.Cells[irow, 9].Value := '�ջ��ֿ�';
  ExcelApp.Cells[irow, 10].Value := '������';
  ExcelApp.Cells[irow, 11].Value := '���ݱ��';
  ExcelApp.Cells[irow, 12].Value := '�����Ϻ�';

  ExcelApp.Columns[1].ColumnWidth:= 10.29;
  ExcelApp.Columns[2].ColumnWidth:= 10.29;
  ExcelApp.Columns[3].ColumnWidth:= 17.43;
  ExcelApp.Columns[4].ColumnWidth:= 17.29;
  ExcelApp.Columns[5].ColumnWidth:= 6.869;
  ExcelApp.Columns[6].ColumnWidth:= 7.29;
  ExcelApp.Columns[7].ColumnWidth:= 4;
  ExcelApp.Columns[8].ColumnWidth:= 16;
  ExcelApp.Columns[9].ColumnWidth:= 10.43;
  ExcelApp.Columns[10].ColumnWidth:= 11;
  ExcelApp.Columns[11].ColumnWidth:= 11;
  ExcelApp.Columns[12].ColumnWidth:= 16.57;

  AddColor(ExcelApp, irow, 1, irow, 12, $DBDCF2);

  irow := irow + 1;
  for iCount := 0 to lst.Count - 1 do
  begin                           
    aCPINRecordPtr := PCPINSumRecord(lst[iCount]);
    if not NumberInList(aCPINRecordPtr.snumber, slProjNo) then Continue;
    ExcelApp.Cells[irow, 1].Value := aCPINRecordPtr^.dt; //'�������';
    ExcelApp.Cells[irow, 2].Value := aCPINRecordPtr^.dtCheck; //'�������';
    ExcelApp.Cells[irow, 3].Value := aCPINRecordPtr^.snumber; //'�Ϻ�';
    ExcelApp.Cells[irow, 4].Value := aCPINRecordPtr^.sname; //'��Ʒ����';
    ExcelApp.Cells[irow, 5].Value := aCPINRecordPtr^.dqty; //'����';
    ExcelApp.Cells[irow, 6].Value := aCPINRecordPtr^.sfac; //'������';
    ExcelApp.Cells[irow, 7].Value := aCPINRecordPtr^.sbatchno; //'����';
    ExcelApp.Cells[irow, 8].Value := aCPINRecordPtr^.snote; //'��ע';
    ExcelApp.Cells[irow, 9].Value := aCPINRecordPtr^.sstock; //'�ջ��ֿ�';
    ExcelApp.Cells[irow, 10].Value := aCPINRecordPtr^.sicmo; //'������';
    ExcelApp.Cells[irow, 11].Value := aCPINRecordPtr^.sbillno; //'���ݱ��';
    ExcelApp.Cells[irow, 12].Value := aCPINRecordPtr^.snumber_fac; //'�����Ϻ�';
    aCPINRecordPtr.bOk := True;
    irow := irow + 1;
  end;

  AddBorder(ExcelApp, 1, 1, irow - 1, 12);

  WorkBook.SaveAs(sfile_save);
  ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  WorkBook.Close;
  ExcelApp.Quit;

            
        
  for iCount := 0 to lst.Count - 1 do
  begin
    aCPINRecordPtr := PCPINSumRecord(lst[iCount]);
    if not aCPINRecordPtr^.bOk then
    begin
//      Log( 'ûд�� ' + aCPINRecordPtr^.snumber );
    end;
    Dispose(aCPINRecordPtr);
  end;

  lst.Free;    
  
end;

procedure TfrmFGAll2MZMB.SaveFGStock(const sfile_save, sfile_data: string;
  slProjNo: TStringList);
type
  TSheetType = (stCPIN, stStock, stUnOut);
var
  ExcelApp, WorkBook: Variant;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  s: string;
  s1, s2, s3, s4, s5, s6: string;
  irow: Integer;
  slSheet: TStringList;  
  slSheetOk: TStringList;
  slSheetCPIN: TStringList;
  slSheetUnOut: TStringList;
  lst: TList;
  iCount: Integer;
  aFGStockRecordPtr: PFGStockRecord;
  aCPINRecordPtr: PCPINRecord;
  aUnOutRecordPtr: PUnOutRecord;
  snumber: string;   
  aSheetType: TSheetType;
  bHaveQtyOk: Boolean;
  icol: Integer;
begin
  slSheet := TStringList.Create;
  slSheetOk := TStringList.Create;
  slSheetCPIN := TStringList.Create;
  slSheetUnOut := TStringList.Create;

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(sfile_data);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        Log(sSheet);

        irow := 1;
        s1 := ExcelApp.Cells[irow, 1].Value;   
        s2 := ExcelApp.Cells[irow, 2].Value;
        s3 := ExcelApp.Cells[irow, 3].Value;
        s4 := ExcelApp.Cells[irow, 4].Value;
        s5 := ExcelApp.Cells[irow, 5].Value;
        s := s1 + s2 + s3 + s4 + s5;
        if s = '�ϺŲ�Ʒ���ƿ����������' then
        begin
          aSheetType := stStock;
        end
        else if s = '��������ϺŲ�Ʒ��������������' then
        begin
          aSheetType := stCPIN;
        end
        else if s = '�ϺŲ�Ʒ����������ע���Ͻ��' then
        begin
          aSheetType := stUnOut;
        end
        else
        begin
          Log('sheet ' + sSheet + ' ��ʽ����'); 
          Continue;
        end;
                  
        lst := TList.Create;

        case aSheetType of
          stStock:
          begin
                
            bHaveQtyOk := False;
            for icol := 1 to 20 do
            begin
              s := ExcelApp.Cells[irow, icol].Value;
              if s = '�ɷ������' then
              begin
                bHaveQtyOk := True;
                Break;
              end;
            end;

            if bHaveQtyOk then
            begin
              slSheetOk.AddObject(sSheet, lst);
            end
            else
            begin
              slSheet.AddObject(sSheet, lst);
            end;

            irow := irow + 2;
            snumber := ExcelApp.Cells[irow, 1].Value;
            while snumber <> '' do
            begin
              aFGStockRecordPtr := New(PFGStockRecord);
              aFGStockRecordPtr^.bOk := False;

              aFGStockRecordPtr^.snumber := ExcelApp.Cells[irow, 1].Value; //�Ϻ�
              aFGStockRecordPtr^.sname := ExcelApp.Cells[irow, 2].Value; //��Ʒ����
              aFGStockRecordPtr^.dqty := ExcelApp.Cells[irow, 3].Value; //�������
              aFGStockRecordPtr^.dqty_rework := ExcelApp.Cells[irow, 4].Value; //	����
              aFGStockRecordPtr^.dqty_uncheck := ExcelApp.Cells[irow, 5].Value; //	������(��3���£�
              aFGStockRecordPtr^.saddr := ExcelApp.Cells[irow, 6].Value; //		����ص�
              aFGStockRecordPtr^.sbatchno := ExcelApp.Cells[irow, 7].Value; //	����
              aFGStockRecordPtr^.snote := ExcelApp.Cells[irow, 8].Value; //	��ע
              if bHaveQtyOk then
              begin
                aFGStockRecordPtr^.dqty_ok := ExcelApp.Cells[irow, 9].Value; //	�ɷ������
                aFGStockRecordPtr^.sncard := ExcelApp.Cells[irow, 10].Value; //	�±���
                aFGStockRecordPtr^.dtEnd := ExcelApp.Cells[irow, 11].Value; //	��ֹ����
              end
              else
              begin
                aFGStockRecordPtr^.sncard := ExcelApp.Cells[irow, 9].Value; //	�±���
                aFGStockRecordPtr^.dtEnd := ExcelApp.Cells[irow, 10].Value; //	��ֹ����              
              end;
              lst.Add(aFGStockRecordPtr);

              irow := irow + 1;   
              snumber := ExcelApp.Cells[irow, 1].Value;
            end;
          end;
          stCPIN:
          begin
            slSheetCPIN.AddObject(sSheet, lst);

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 2].Value;
            while snumber <> '' do
            begin
              aCPINRecordPtr := New(PCPINRecord);
              aCPINRecordPtr^.bOk := False;
 
              aCPINRecordPtr^.dt := ExcelApp.Cells[irow, 1].Value;
              aCPINRecordPtr^.snumber := ExcelApp.Cells[irow, 2].Value;
              aCPINRecordPtr^.sname := ExcelApp.Cells[irow, 3].Value;
              aCPINRecordPtr^.dqty := ExcelApp.Cells[irow, 4].Value;
              aCPINRecordPtr^.sfac := ExcelApp.Cells[irow, 5].Value;
              aCPINRecordPtr^.sbatchno := ExcelApp.Cells[irow, 6].Value;
              aCPINRecordPtr^.snote := ExcelApp.Cells[irow, 7].Value;
              aCPINRecordPtr^.sstock := ExcelApp.Cells[irow, 8].Value;
              aCPINRecordPtr^.sicmo := ExcelApp.Cells[irow, 9].Value;
              aCPINRecordPtr^.sbillno := ExcelApp.Cells[irow, 10].Value;
              aCPINRecordPtr^.snumber_fac := ExcelApp.Cells[irow, 11].Value;
                   
              lst.Add(aCPINRecordPtr);

              irow := irow + 1;   
              snumber := ExcelApp.Cells[irow, 2].Value;
            end;
          end;
          stUnOut:
          begin
            slSheetUnOut.AddObject(sSheet, lst);

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 2].Value;
            while snumber <> '' do
            begin
              aUnOutRecordPtr := New(PUnOutRecord);
              aUnOutRecordPtr^.bOk := False;
 
              aUnOutRecordPtr^.snumber := ExcelApp.Cells[irow, 1].Value;
              aUnOutRecordPtr^.sname := ExcelApp.Cells[irow, 2].Value;
              aUnOutRecordPtr^.sqty_mat := ExcelApp.Cells[irow, 3].Value;
              aUnOutRecordPtr^.snote := ExcelApp.Cells[irow, 4].Value;
              aUnOutRecordPtr^.sqty_mat := ExcelApp.Cells[irow, 5].Value;
                   
              lst.Add(aUnOutRecordPtr);

              irow := irow + 1;   
              snumber := ExcelApp.Cells[irow, 2].Value;
            end;
          end;
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

  ///////////////////////////////////////////////////////////////////////


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



                    
  for iSheet := 0 to slSheetCPIN.Count - 1 do
  begin
    lst := TList(slSheetCPIN.Objects[iSheet]);

    if iSheet > 0 then
    begin
      ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[ExcelApp.Sheets.Count]);
    end;

    ExcelApp.Sheets[ExcelApp.Sheets.Count].Activate;
    ExcelApp.Sheets[ExcelApp.Sheets.Count].Name := slSheetCPIN[iSheet];
    irow := 1; 
    ExcelApp.Cells[irow, 1].Value := '�������';
    ExcelApp.Cells[irow, 2].Value := '�Ϻ�';
    ExcelApp.Cells[irow, 3].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 4].Value := '����';
    ExcelApp.Cells[irow, 5].Value := '������';
    ExcelApp.Cells[irow, 6].Value := '����';
    ExcelApp.Cells[irow, 7].Value := '��ע';
    ExcelApp.Cells[irow, 8].Value := '�ջ��ֿ�';
    ExcelApp.Cells[irow, 9].Value := '������';
    ExcelApp.Cells[irow, 10].Value := '���ݱ��';
    ExcelApp.Cells[irow, 11].Value := '�����Ϻ�';

    ExcelApp.Columns[1].ColumnWidth := 10.75;
    ExcelApp.Columns[2].ColumnWidth := 17.75;
    ExcelApp.Columns[3].ColumnWidth := 18.63;
    ExcelApp.Columns[4].ColumnWidth := 7.13;
    ExcelApp.Columns[5].ColumnWidth := 6.5;
    ExcelApp.Columns[6].ColumnWidth := 4.38;
    ExcelApp.Columns[7].ColumnWidth := 14.38;
    ExcelApp.Columns[8].ColumnWidth := 10.38;
    ExcelApp.Columns[9].ColumnWidth := 10.88;
    ExcelApp.Columns[10].ColumnWidth := 9.88;
    ExcelApp.Columns[11].ColumnWidth := 17;


    AddColor(ExcelApp, irow, 1, irow, 11, $DEF1EB);

    irow := irow + 1;
    for iCount := 0 to lst.Count - 1 do
    begin                           
      aCPINRecordPtr := PCPINRecord(lst[iCount]);
      if not NumberInList(aCPINRecordPtr.snumber, slProjNo) then Continue;

      ExcelApp.Cells[irow, 1].Value := aCPINRecordPtr^.dt; // '�������';
      ExcelApp.Cells[irow, 2].Value := aCPINRecordPtr^.snumber; // '�Ϻ�';
      ExcelApp.Cells[irow, 3].Value := aCPINRecordPtr^.sname; // '��Ʒ����';
      ExcelApp.Cells[irow, 4].Value := aCPINRecordPtr^.dqty; // '����';
      ExcelApp.Cells[irow, 5].Value := aCPINRecordPtr^.sfac; // '������';
      ExcelApp.Cells[irow, 6].Value := aCPINRecordPtr^.sbatchno; // '����';
      ExcelApp.Cells[irow, 7].Value := aCPINRecordPtr^.snote; // '��ע';
      ExcelApp.Cells[irow, 8].Value := aCPINRecordPtr^.sstock; // '�ջ��ֿ�';
      ExcelApp.Cells[irow, 9].Value := aCPINRecordPtr^.sicmo; // '������';
      ExcelApp.Cells[irow, 10].Value := aCPINRecordPtr^.sbillno; // '���ݱ��';
      ExcelApp.Cells[irow, 11].Value := aCPINRecordPtr^.snumber_fac; //'�����Ϻ�';
 
      aCPINRecordPtr.bOk := True;
      irow := irow + 1;
    end;

    AddBorder(ExcelApp, 1, 1, irow - 1, 11);
 
  end;

////////////////////////////////////////////////////////////////////////////////


  for iSheet := 0 to slSheetOk.Count - 1 do
  begin
    lst := TList(slSheetOk.Objects[iSheet]);
                                   
    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[ExcelApp.Sheets.Count]);

    ExcelApp.Sheets[ExcelApp.Sheets.Count].Activate;
    ExcelApp.Sheets[ExcelApp.Sheets.Count].Name := slSheetOk[iSheet];
    irow := 1; 
    ExcelApp.Cells[irow, 1].Value := '�Ϻ�';
    ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 3].Value := '�������';
    ExcelApp.Cells[irow, 4].Value := '����';     
    ExcelApp.Cells[irow + 1, 4].Value := '����';
    ExcelApp.Cells[irow + 1, 5].Value := '������(��3���£�';
    ExcelApp.Cells[irow, 6].Value := '����ص�';
    ExcelApp.Cells[irow, 7].Value := '����';
    ExcelApp.Cells[irow, 8].Value := '��ע';  
    ExcelApp.Cells[irow, 9].Value := '�ɷ������';
    ExcelApp.Cells[irow, 10].Value := '�±���';
    ExcelApp.Cells[irow, 11].Value := '��ֹ����';

    MergeCells(ExcelApp, irow, 1, irow + 1, 1);
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);
    MergeCells(ExcelApp, irow, 3, irow + 1, 3);
    
    MergeCells(ExcelApp, irow, 6, irow + 1, 6);
    MergeCells(ExcelApp, irow, 7, irow + 1, 7);
    MergeCells(ExcelApp, irow, 8, irow + 1, 8);   
    MergeCells(ExcelApp, irow, 9, irow + 1, 9);
    MergeCells(ExcelApp, irow, 10, irow + 1, 10);
    MergeCells(ExcelApp, irow, 11, irow + 1, 11);

    ExcelApp.Columns[1].ColumnWidth := 16.5;
    ExcelApp.Columns[2].ColumnWidth := 21;
    ExcelApp.Columns[3].ColumnWidth := 9;
    ExcelApp.Columns[4].ColumnWidth := 8.5;
    ExcelApp.Columns[5].ColumnWidth := 14.6;
    ExcelApp.Columns[6].ColumnWidth := 12.88;
    ExcelApp.Columns[7].ColumnWidth := 12.88;
    ExcelApp.Columns[8].ColumnWidth := 7.88;  
    ExcelApp.Columns[9].ColumnWidth := 7.88;
    ExcelApp.Columns[10].ColumnWidth := 8.38;
    ExcelApp.Columns[11].ColumnWidth := 12.88;

    

    AddColor(ExcelApp, irow, 1, irow + 1, 11, $DEF1EB);

    irow := irow + 2;
    for iCount := 0 to lst.Count - 1 do
    begin                           
      aFGStockRecordPtr := PFGStockRecord(lst[iCount]);
      if not NumberInList(aFGStockRecordPtr.snumber, slProjNo) then Continue;
      ExcelApp.Cells[irow, 1].Value := aFGStockRecordPtr^.snumber;  // '�Ϻ�';
      ExcelApp.Cells[irow, 2].Value := aFGStockRecordPtr^.sname;  // '��Ʒ����';
      ExcelApp.Cells[irow, 3].Value := aFGStockRecordPtr^.dqty;  // '�������';
      ExcelApp.Cells[irow, 4].Value := aFGStockRecordPtr^.dqty_rework;  // '����';
      ExcelApp.Cells[irow, 5].Value := aFGStockRecordPtr^.dqty_uncheck;  // '������(��3���£�';
      ExcelApp.Cells[irow, 6].Value := aFGStockRecordPtr^.saddr;  // '����ص�';
      ExcelApp.Cells[irow, 7].Value := aFGStockRecordPtr^.sbatchno;  // '����';
      ExcelApp.Cells[irow, 8].Value := aFGStockRecordPtr^.snote;  // '��ע';    
      ExcelApp.Cells[irow, 9].Value := aFGStockRecordPtr.dqty_ok;
      ExcelApp.Cells[irow, 10].Value := aFGStockRecordPtr^.sncard;  // '�±���';
      ExcelApp.Cells[irow, 11].Value := aFGStockRecordPtr^.dtEnd;  // '��ֹ����';
      aFGStockRecordPtr.bOk := True;
      irow := irow + 1;
    end;

    AddBorder(ExcelApp, 1, 1, irow - 1, 11);
  end;



////////////////////////////////////////////////////////////////////////////////


  for iSheet := 0 to slSheet.Count - 1 do
  begin
    lst := TList(slSheet.Objects[iSheet]);
                                   
    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[ExcelApp.Sheets.Count]);

    ExcelApp.Sheets[ExcelApp.Sheets.Count].Activate;
    ExcelApp.Sheets[ExcelApp.Sheets.Count].Name := slSheet[iSheet];
    irow := 1; 
    ExcelApp.Cells[irow, 1].Value := '�Ϻ�';
    ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 3].Value := '�������';
    ExcelApp.Cells[irow, 4].Value := '����';     
    ExcelApp.Cells[irow + 1, 4].Value := '����';
    ExcelApp.Cells[irow + 1, 5].Value := '������(��3���£�';
    ExcelApp.Cells[irow, 6].Value := '����ص�';
    ExcelApp.Cells[irow, 7].Value := '����';
    ExcelApp.Cells[irow, 8].Value := '��ע';   
    ExcelApp.Cells[irow, 9].Value := '�±���';
    ExcelApp.Cells[irow, 10].Value := '��ֹ����';

    MergeCells(ExcelApp, irow, 1, irow + 1, 1);
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);
    MergeCells(ExcelApp, irow, 3, irow + 1, 3);
    
    MergeCells(ExcelApp, irow, 6, irow + 1, 6);
    MergeCells(ExcelApp, irow, 7, irow + 1, 7);
    MergeCells(ExcelApp, irow, 8, irow + 1, 8);
    MergeCells(ExcelApp, irow, 9, irow + 1, 9);
    MergeCells(ExcelApp, irow, 10, irow + 1, 10);

    ExcelApp.Columns[1].ColumnWidth := 16.5;
    ExcelApp.Columns[2].ColumnWidth := 21;
    ExcelApp.Columns[3].ColumnWidth := 9;
    ExcelApp.Columns[4].ColumnWidth := 8.5;
    ExcelApp.Columns[5].ColumnWidth := 14.6;
    ExcelApp.Columns[6].ColumnWidth := 12.88;
    ExcelApp.Columns[7].ColumnWidth := 12.88;
    ExcelApp.Columns[8].ColumnWidth := 7.88;
    ExcelApp.Columns[9].ColumnWidth := 8.38;
    ExcelApp.Columns[10].ColumnWidth := 12.88;

    

    AddColor(ExcelApp, irow, 1, irow + 1, 10, $DEF1EB);

    irow := irow + 2;
    for iCount := 0 to lst.Count - 1 do
    begin                           
      aFGStockRecordPtr := PFGStockRecord(lst[iCount]);
      if not NumberInList(aFGStockRecordPtr.snumber, slProjNo) then Continue;
      ExcelApp.Cells[irow, 1].Value := aFGStockRecordPtr^.snumber;  // '�Ϻ�';
      ExcelApp.Cells[irow, 2].Value := aFGStockRecordPtr^.sname;  // '��Ʒ����';
      ExcelApp.Cells[irow, 3].Value := aFGStockRecordPtr^.dqty;  // '�������';
      ExcelApp.Cells[irow, 4].Value := aFGStockRecordPtr^.dqty_rework;  // '����';
      ExcelApp.Cells[irow, 5].Value := aFGStockRecordPtr^.dqty_uncheck;  // '������(��3���£�';
      ExcelApp.Cells[irow, 6].Value := aFGStockRecordPtr^.saddr;  // '����ص�';
      ExcelApp.Cells[irow, 7].Value := aFGStockRecordPtr^.sbatchno;  // '����';
      ExcelApp.Cells[irow, 8].Value := aFGStockRecordPtr^.snote;  // '��ע';
      ExcelApp.Cells[irow, 9].Value := aFGStockRecordPtr^.sncard;  // '�±���';
      ExcelApp.Cells[irow, 10].Value := aFGStockRecordPtr^.dtEnd;  // '��ֹ����';
      aFGStockRecordPtr.bOk := True;
      irow := irow + 1;
    end;

    AddBorder(ExcelApp, 1, 1, irow - 1, 10);
  end;


       

  for iSheet := 0 to slSheetUnOut.Count - 1 do
  begin
    lst := TList(slSheetUnOut.Objects[iSheet]);
                 
    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[ExcelApp.Sheets.Count]);

    ExcelApp.Sheets[ExcelApp.Sheets.Count].Activate;
    ExcelApp.Sheets[ExcelApp.Sheets.Count].Name := slSheetUnOut[iSheet];
    irow := 1; 
    ExcelApp.Cells[irow, 1].Value := '�Ϻ�';
    ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 3].Value := '����';
    ExcelApp.Cells[irow, 4].Value := '��ע';
    ExcelApp.Cells[irow, 5].Value := '���Ͻ��';


    ExcelApp.Columns[1].ColumnWidth := 21.25;
    ExcelApp.Columns[2].ColumnWidth := 44;
    ExcelApp.Columns[3].ColumnWidth := 13;
    ExcelApp.Columns[4].ColumnWidth := 13;
    ExcelApp.Columns[5].ColumnWidth := 21.25; 


    AddColor(ExcelApp, irow, 1, irow, 5, $DEF1EB);

    irow := irow + 1;
    for iCount := 0 to lst.Count - 1 do
    begin                           
      aUnOutRecordPtr := PUnOutRecord(lst[iCount]);
      if not NumberInList(aUnOutRecordPtr.snumber, slProjNo) then Continue;

      ExcelApp.Cells[irow, 1].Value := aUnOutRecordPtr^.snumber;
      ExcelApp.Cells[irow, 2].Value := aUnOutRecordPtr^.sname;
      ExcelApp.Cells[irow, 3].Value := aUnOutRecordPtr^.dqty;
      ExcelApp.Cells[irow, 4].Value := aUnOutRecordPtr^.snote;
      ExcelApp.Cells[irow, 5].Value := aUnOutRecordPtr^.sqty_mat;
 
      aUnOutRecordPtr.bOk := True;
      irow := irow + 1;
    end;

    AddBorder(ExcelApp, 1, 1, irow - 1, 5);
 
  end;




  WorkBook.SaveAs(sfile_save);
  ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  WorkBook.Close;
  ExcelApp.Quit;

            
  for iSheet := 0 to slSheet.Count - 1 do
  begin
    lst := TList(slSheet.Objects[iSheet]);
    for iCount := 0 to lst.Count - 1 do
    begin
      aFGStockRecordPtr := PFGStockRecord(lst[iCount]);
      if not aFGStockRecordPtr^.bOk then
      begin
//        Log( slSheet[iSheet] + 'ûд�� ' + aFGStockRecordPtr^.snumber );
      end;
      Dispose(aFGStockRecordPtr);
    end;
  end;

  for iSheet := 0 to slSheetOk.Count - 1 do
  begin
    lst := TList(slSheetOk.Objects[iSheet]);
    for iCount := 0 to lst.Count - 1 do
    begin
      aFGStockRecordPtr := PFGStockRecord(lst[iCount]);
      if not aFGStockRecordPtr^.bOk then
      begin
//        Log( slSheetOk[iSheet] + 'ûд�� ' + aFGStockRecordPtr^.snumber );
      end;
      Dispose(aFGStockRecordPtr);
    end;
  end;

  for iSheet := 0 to slSheetCPIN.Count - 1 do
  begin
    lst := TList(slSheetCPIN.Objects[iSheet]);
    for iCount := 0 to lst.Count - 1 do
    begin
      aCPINRecordPtr := PCPINRecord(lst[iCount]);
      if not aCPINRecordPtr^.bOk then
      begin
//        Log(slSheetCPIN[iSheet] + 'ûд�� ' + aCPINRecordPtr^.snumber );
      end;
      Dispose(aCPINRecordPtr);
    end;
  end;

  for iSheet := 0 to slSheetUnOut.Count - 1 do
  begin
    lst := TList(slSheetUnOut.Objects[iSheet]);
    for iCount := 0 to lst.Count - 1 do
    begin
      aUnOutRecordPtr := PUnOutRecord(lst[iCount]);
      if not aUnOutRecordPtr^.bOk then
      begin
//        Log(slSheetUnOut[iSheet] + 'ûд�� ' + aCPINRecordPtr^.snumber );
      end;
      Dispose(aUnOutRecordPtr);
    end;
  end;
 
  slSheet.Free;
  slSheetOk.Free;
  slSheetCPIN.Free;
  slSheetUnOut.Free;
 
end;

procedure TfrmFGAll2MZMB.btnFGStock_sClick(Sender: TObject);
var
  sfile: string;       
  slProjNo_OEM: TStringList;
  slProjNo_ODM: TStringList;
  s: string;
begin                                  
  if not ExcelSaveDialog(sfile) then Exit;
      
  slProjNo_OEM := TfrmProjNameNo.GetProjNos_OEM;
  slProjNo_ODM := TfrmProjNameNo.GetProjNos_ODM;

  s := ChangeFileExt(sfile, '') + '-����' + ExtractFileExt(sfile);
  SaveFGStock(s, leFGStock.Text, slProjNo_OEM);
  s := ChangeFileExt(sfile, '') + '-����' + ExtractFileExt(sfile);
  SaveFGStock(s, leFGStock.Text, slProjNo_ODM);
   
  slProjNo_OEM.Free;
  slProjNo_ODM.Free;

  MessageBox(Handle, '���', '��ʾ', 0);
end;

end.
