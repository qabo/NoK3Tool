unit Excel2MrpBomWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, RawMPSReader, NumberCheckReader,
  ExcelBomReader, StdCtrls, ExtCtrls, IniFiles, CommUtils, SAPMaterialReader,
  SAPMaterialReader2, ComObj;

type
  TfrmExcel2MrpBom = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leExcelBom: TLabeledEdit;
    leBomCheck: TLabeledEdit;
    leRawMps: TLabeledEdit;
    btnExcelBom: TButton;
    btnBomCheck: TButton;
    btnRawMps: TButton;
    leMMList: TLabeledEdit;
    btnMMList: TButton;
    Memo1: TMemo;
    ToolButton1: TToolButton;
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnExcelBomClick(Sender: TObject);
    procedure btnBomCheckClick(Sender: TObject);
    procedure btnRawMpsClick(Sender: TObject);
    procedure btnMMListClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;
 
implementation

{$R *.dfm}

class procedure TfrmExcel2MrpBom.ShowForm;  
var
  frmExcel2MrpBom: TfrmExcel2MrpBom;
begin
  frmExcel2MrpBom := TfrmExcel2MrpBom.Create(nil);
  try
    frmExcel2MrpBom.ShowModal;
  finally
    frmExcel2MrpBom.Free;
  end;
end;
     
procedure TfrmExcel2MrpBom.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leExcelBom.Text := ini.ReadString(self.ClassName, leExcelBom.Name, '');
    leBomCheck.Text := ini.ReadString(self.ClassName, leBomCheck.Name, '');
    leRawMps.Text := ini.ReadString(self.ClassName, leRawMps.Name, '');
    leMMList.Text := ini.ReadString(self.ClassName, leMMList.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmExcel2MrpBom.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leExcelBom.Name, leExcelBom.Text);
    ini.WriteString(self.ClassName, leBomCheck.Name, leBomCheck.Text);
    ini.WriteString(self.ClassName, leRawMps.Name, leRawMps.Text);
    ini.WriteString(self.ClassName, leMMList.Name, leMMList.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmExcel2MrpBom.btnSave2Click(Sender: TObject);    
var
  aRawMPSReader: TRawMPSReader;
//  aNumberCheckReader: TNumberCheckReader;
  aExcelBomReader: TExcelBomReader;
  aSAPMaterialReader: TSAPMaterialReader2;

  sfile: string;
  iBom: Integer;
  aRawBomPtr: PRawBom;

  ExcelApp, WorkBook: Variant;
  irow: Integer;
  sbom: string;
  igroup: Integer;
  inumber: Integer;
  aBomLoc: TBomLoc;
  aBomLocLinePtr: PBomLocLine;
  ic: Integer;

  aSAPMaterialRecordPtr: PSAPMaterialRecord;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  aRawMPSReader := TRawMPSReader.Create(leRawMps.Text);

//  aNumberCheckReader := TNumberCheckReader.Create( leBomCheck.Text);

  aExcelBomReader := TExcelBomReader.Create(leExcelBom.Text);
  
  aSAPMaterialReader := TSAPMaterialReader2.Create(leMMList.Text);
  try
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

    ExcelApp.Cells[irow, 1].Value := '母件物料编码';
    ExcelApp.Cells[irow, 2].Value := '母件物料描述';
    ExcelApp.Cells[irow, 3].Value := '工厂';
    ExcelApp.Cells[irow, 4].Value := '用途';
    ExcelApp.Cells[irow, 5].Value := '代工厂';
    ExcelApp.Cells[irow, 6].Value := '母件L/T';
    ExcelApp.Cells[irow, 7].Value := '基本数量';
    ExcelApp.Cells[irow, 8].Value := '状态';
    ExcelApp.Cells[irow, 9].Value := '上层物料编码';
    ExcelApp.Cells[irow, 10].Value := '层级';
    ExcelApp.Cells[irow, 11].Value := 'No';
    ExcelApp.Cells[irow, 12].Value := '项目类别';
    ExcelApp.Cells[irow, 13].Value := '子件物料编码';
    ExcelApp.Cells[irow, 14].Value := '子件物料描述';
    ExcelApp.Cells[irow, 15].Value := '采购类型';
    ExcelApp.Cells[irow, 16].Value := 'ABC标识';
    ExcelApp.Cells[irow, 17].Value := '关键物料';
    ExcelApp.Cells[irow, 18].Value := 'L/T';
    ExcelApp.Cells[irow, 19].Value := '子件用量';
    ExcelApp.Cells[irow, 20].Value := '单位';
    ExcelApp.Cells[irow, 21].Value := '替代项目组';
    ExcelApp.Cells[irow, 22].Value := '优先级';
    ExcelApp.Cells[irow, 23].Value := '使用可能性';
    ExcelApp.Cells[irow, 24].Value := '排序字符串';
    ExcelApp.Cells[irow, 25].Value := '位号1';
    ExcelApp.Cells[irow, 26].Value := '位号2';
    ExcelApp.Cells[irow, 27].Value := '位号3';
    ExcelApp.Cells[irow, 28].Value := '位号4';
 
    try

      irow := irow + 1;
      for iBom := 0 to aRawMPSReader.slBomNumber.Count - 1 do
      begin
        aRawBomPtr := PRawBom(aRawMPSReader.slBomNumber.Objects[iBom]);
        sbom := aRawBomPtr^.sbom;

        ic := 1;
        for igroup := 0 to aExcelBomReader.FList.Count - 1 do
        begin
          aBomLoc := TBomLoc(aExcelBomReader.FList.Objects[igroup]);
          for inumber := 0 to aBomLoc.FList.Count - 1 do
          begin
            aBomLocLinePtr := PBomLocLine(aBomLoc.FList.Objects[inumber]);
            if ((Pos(aRawBomPtr^.sver, aBomLocLinePtr^.sver) > 0) or (aBomLocLinePtr^.sver = '通用')) and
              ((Pos(aRawBomPtr^.scap, aBomLocLinePtr^.scap) > 0) or (aBomLocLinePtr^.scap = '通用')) and
              ((Pos(aRawBomPtr^.scol, aBomLocLinePtr^.scol) > 0) or (aBomLocLinePtr^.scol = '通用')) then
            begin
              aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aBomLocLinePtr^.snumber);
              if sbom <> '' then
              begin
                ExcelApp.Cells[irow, 1].Value := sbom;  //'母件物料编码';
                ExcelApp.Cells[irow, 2].Value := sbom;  //'母件物料描述';
                ExcelApp.Cells[irow, 3].Value := '1001';//'工厂';
                ExcelApp.Cells[irow, 4].Value := '1';   //'用途';
                ExcelApp.Cells[irow, 5].Value := 'ML';  //'代工厂';
                sbom := '';
              end;
              ExcelApp.Cells[irow, 6].Value := 7; // '母件L/T';
              ExcelApp.Cells[irow, 7].Value := 1; // '基本数量';
              ExcelApp.Cells[irow, 8].Value := 1; // '状态';
              ExcelApp.Cells[irow, 9].Value := aRawBomPtr^.sbom; //'上层物料编码';
              ExcelApp.Cells[irow, 10].Value := '1.' + IntToStr(ic);   //'层级';
              ExcelApp.Cells[irow, 11].Value := Copy(IntToStr(10000 + ic * 10), 2, 4); // 'No';
              ExcelApp.Cells[irow, 12].Value := 'L'; //'项目类别';
              ExcelApp.Cells[irow, 13].Value := aBomLocLinePtr^.snumber; // '子件物料编码';
              ExcelApp.Cells[irow, 14].Value := aBomLocLinePtr^.sname; // '子件物料描述';
              if aSAPMaterialRecordPtr <> nil then
              begin
                ExcelApp.Cells[irow, 15].Value := aSAPMaterialRecordPtr^.sPType; //'F'; // '采购类型';
              end
              else
              begin
                if (Copy(aBomLocLinePtr^.snumber, 1, 2) = '01') or
                  (Copy(aBomLocLinePtr^.snumber, 1, 2) = '04') then
                begin
                  ExcelApp.Cells[irow, 15].Value := 'F';
                end
                else
                begin
                  ExcelApp.Cells[irow, 15].Value := 'E';
                end;
              end;
              ExcelApp.Cells[irow, 16].Value := '';  // 'ABC标识';
              ExcelApp.Cells[irow, 17].Value := '';  //'关键物料';
              if aSAPMaterialRecordPtr <> nil then
              begin
                ExcelApp.Cells[irow, 18].Value := aSAPMaterialRecordPtr^.dLT_PD; // 'L/T';
              end
              else
              begin
                Memo1.Lines.Add('物料清单找不到 ' + aBomLocLinePtr^.snumber);
              end;
              ExcelApp.Cells[irow, 19].Value := aBomLoc.susage; // '子件用量';
              ExcelApp.Cells[irow, 20].Value := 'ST'; // '单位';
              if aBomLoc.FList.Count > 1 then
              begin
                ExcelApp.Cells[irow, 21].Value := 'A' + IntToStr(1000 + igroup); // '替代项目组';
              end;
              ExcelApp.Cells[irow, 22].Value := IntToStr(inumber + 1); //'优先级';
              ExcelApp.Cells[irow, 23].Value := 0; //'使用可能性';
              ExcelApp.Cells[irow, 24].Value := '';  //'排序字符串';
              ExcelApp.Cells[irow, 25].Value := '';  //'位号1';
              ExcelApp.Cells[irow, 26].Value := '';  //'位号2';
              ExcelApp.Cells[irow, 27].Value := '';  //'位号3';
              ExcelApp.Cells[irow, 28].Value := '';  //'位号4';

              ic := ic + 1;

              irow := irow + 1;
            end;  
          end;
        end;
      end;
            
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

             
    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

    MessageBox(self.Handle, '完成', '提示', 0);



  finally
    aRawMPSReader.Free;
//    aNumberCheckReader.Free;
    aExcelBomReader.Free;
    aSAPMaterialReader.Free;
  end;
end;

procedure TfrmExcel2MrpBom.btnExcelBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  leExcelBom.Text := sfile;
end;

procedure TfrmExcel2MrpBom.btnBomCheckClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  leBomCheck.Text := sfile;
end;

procedure TfrmExcel2MrpBom.btnRawMpsClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  leRawMps.Text := sfile;
end;

procedure TfrmExcel2MrpBom.btnMMListClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  leMMList.Text := sfile;
end;

procedure TfrmExcel2MrpBom.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmExcel2MrpBom.ToolButton1Click(Sender: TObject);
var
  aRawMPSReader: TRawMPSReader;
//  aSAPMaterialReader: TSAPMaterialReader2;

  sfile: string;
  imps: Integer;
  aRawBomPtr: PRawBom;

  ExcelApp, WorkBook: Variant;
  irow: Integer;
  icol: Integer;
  sbom: string;
  igroup: Integer;
  inumber: Integer;
  aBomLoc: TBomLoc;
  aBomLocLinePtr: PBomLocLine;
  ic: Integer;

  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  iweek: Integer;
  aRawMPSLine: TRawMPSLine;
  aRawMPSColPtr: PRawMPSCol;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  aRawMPSReader := TRawMPSReader.Create(leRawMps.Text);
 
//  aSAPMaterialReader := TSAPMaterialReader2.Create(leMMList.Text);
  try
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

    try

      icol := 13;
      for iweek := 0 to aRawMPSReader.slWeek.Count - 1 do
      begin
        ExcelApp.Cells[irow, icol + iweek].Value := 'W' + aRawMPSReader.slWeek[iweek];
      end;

      irow := irow + 1;
      for imps := 0 to aRawMPSReader.FList.Count - 1 do
      begin
        aRawMPSLine := TRawMPSLine(aRawMPSReader.FList.Objects[imps]);

        ExcelApp.Cells[irow, 1].Value := aRawMPSLine.sbom;  // 物料
        ExcelApp.Cells[irow, 2].Value := aRawMPSLine.sbom;  // 物料描述
        ExcelApp.Cells[irow, 3].Value := 0;       //'物料组';
        ExcelApp.Cells[irow, 4].Value := 0;       //'物料组描述';
        ExcelApp.Cells[irow, 5].Value := '1001';  //'工厂';
        ExcelApp.Cells[irow, 6].Value := 'BSF';   //'需求类型';
        ExcelApp.Cells[irow, 7].Value := '00';    //'版本';
        ExcelApp.Cells[irow, 8].Value := 'X';     //'Act';
        ExcelApp.Cells[irow, 9].Value := '';      //'需求计划号';
        ExcelApp.Cells[irow, 10].Value := aRawMPSLine.sarea; //'MRP 范围';
        ExcelApp.Cells[irow, 11].Value := 'A00';  //'MRP控制者';
        ExcelApp.Cells[irow, 12].Value := 'PC';   //'基本计量单位';


        icol := 13;
        for iweek := 0 to aRawMPSLine.FList.Count - 1 do
        begin
          aRawMPSColPtr := PRawMPSCol(aRawMPSLine.FList[iweek]);
          ExcelApp.Cells[irow, icol + iweek].Value := aRawMPSColPtr^.iQty;
        end;

        irow := irow + 1;
      end;
            
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

             
    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

    MessageBox(self.Handle, '完成', '提示', 0);



  finally
    aRawMPSReader.Free;
//    aSAPMaterialReader.Free;
  end;
end;

end.
