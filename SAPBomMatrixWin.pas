unit SAPBomMatrixWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, CommUtils, StdCtrls, ExtCtrls, ComCtrls, ToolWin,
  ImgList, ProjNameReader, BomD2Reader, BomD2Reader2, ComObj;

type
  TfrmSAPBomMatrix = class(TForm)
    leProjName: TLabeledEdit;
    leSapBom: TLabeledEdit;
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    tbSave: TToolButton;
    ToolButton1: TToolButton;
    tbQuit: TToolButton;
    btnProjName: TButton;
    btnSapBom: TButton;
    GroupBox1: TGroupBox;
    mmoNumber: TMemo;
    ProgressBar1: TProgressBar;
    Memo1: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbQuitClick(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure btnProjNameClick(Sender: TObject);
    procedure btnSapBomClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;



implementation

{$R *.dfm}

{ TfrmSAPBomMatrix }

class procedure TfrmSAPBomMatrix.ShowForm;
var
  frmSAPBomMatrix: TfrmSAPBomMatrix;
begin
  frmSAPBomMatrix := TfrmSAPBomMatrix.Create(nil);
  try
    frmSAPBomMatrix.ShowModal;;
  finally
    frmSAPBomMatrix.Free;
  end;
end;

procedure TfrmSAPBomMatrix.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    leProjName.Text := ini.ReadString(self.ClassName, leProjName.Name, '');
    leSapBom.Text := ini.ReadString(self.ClassName, leSapBom.Name, '');
    s := ini.ReadString(self.ClassName, mmoNumber.Name, '');
    mmoNumber.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;

procedure TfrmSAPBomMatrix.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leProjName.Name, leProjName.Text);
    ini.WriteString(self.ClassName, leSapBom.Name, leSapBom.Text);
    s := StringReplace(mmoNumber.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(Self.ClassName, mmoNumber.Name, s);
  finally
    ini.Free;
  end;
end;
  
procedure TfrmSAPBomMatrix.btnProjNameClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leProjName.Text := sfile;
end;

procedure TfrmSAPBomMatrix.btnSapBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSapBom.Text := sfile;
end;

procedure TfrmSAPBomMatrix.tbQuitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSAPBomMatrix.tbSaveClick(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;
  sfile: string;
  aProjNameReader: TProjNameReader;
  aBomD2Reader: TBomD2Reader2;
  icol: Integer;
  icol2: Integer;
  icol3: Integer;
  irow: Integer;
  iline: Integer;
  sproj: string;
  inumber: Integer;
  ptrBomD2Item: PBomD2Item;
  ptrBomD2Item_child: PBomD2Item;
  aBomD2: TBomD2;
  dwTick: DWORD;
  slNumber: TStringList;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  dwTick := GetTickCount;

  slNumber := TStringList.Create;
  slNumber.Text := mmoNumber.Text;
  
  aProjNameReader := TProjNameReader.Create(leProjName.Text);
  aBomD2Reader := TBomD2Reader2.Create(leSapBom.Text, aProjNameReader);
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

    ExcelApp.Sheets[1].Activate;
//    ExcelApp.Sheets[1].Name := 'WhereUser';

    try
      irow := 2;
      ExcelApp.Cells[irow, 1].Value := '子件物料编码';
      ExcelApp.Cells[irow, 2].Value := '子件物料描述';
      ExcelApp.Cells[irow, 3].Value := 'ABC标识';
      ExcelApp.Cells[irow, 4].Value := '采购类型';
      ExcelApp.Cells[irow, 5].Value := 'L/T';
      ExcelApp.Cells[irow, 6].Value := '所用机种';

      ExcelApp.Columns[1].ColumnWidth := 12.5;   
      ExcelApp.Columns[2].ColumnWidth := 38.63;
      ExcelApp.Columns[3].ColumnWidth := 8.38;
      ExcelApp.Columns[4].ColumnWidth := 8.38;
      ExcelApp.Columns[5].ColumnWidth := 8.38;
      ExcelApp.Columns[6].ColumnWidth := 14.88;

      icol := 7;    
      AddColor(ExcelApp, irow - 1, icol, irow - 1, icol + slNumber.Count - 1, $C0C0C0);
      for iline := 0 to slNumber.Count - 1 do
      begin
        sproj := aProjNameReader.ProjOfNumber(Trim(slNumber[iline]));
        ExcelApp.Cells[irow - 1, icol].Value := sproj;
        ExcelApp.Cells[irow, icol].Value := Trim(slNumber[iline]);
        ExcelApp.Columns[icol].ColumnWidth := 12.5;
        icol := icol + 1;
      end;


      ExcelApp.Cells[irow, icol].Value := '使用可能性';
      icol := icol + 1;      
      AddColor(ExcelApp, irow - 1, icol, irow - 1, icol + slNumber.Count - 1, $C0C0C0);
      for iline := 0 to slNumber.Count - 1 do
      begin
        sproj := aProjNameReader.ProjOfNumber(Trim(slNumber[iline]));
        ExcelApp.Cells[irow - 1, icol].Value := sproj;   
        ExcelApp.Cells[irow, icol].Value := Trim(slNumber[iline]);
        ExcelApp.Columns[icol].ColumnWidth := 12.5;
        icol := icol + 1;
      end;  
                     
      ExcelApp.Cells[irow, icol].Value := '替代项目组';
      icol := icol + 1;            
      AddColor(ExcelApp, irow - 1, icol, irow - 1, icol + slNumber.Count - 1, $C0C0C0);
      for iline := 0 to slNumber.Count - 1 do
      begin
        sproj := aProjNameReader.ProjOfNumber(Trim(slNumber[iline]));
        ExcelApp.Cells[irow - 1, icol].Value := sproj;  
        ExcelApp.Cells[irow, icol].Value := Trim(slNumber[iline]);
        ExcelApp.Columns[icol].ColumnWidth := 12.5;
        icol := icol + 1;
      end;       

      AddColor(ExcelApp, irow, 1, irow, icol - 1, $C0C0C0);

      for iline := 0 to slNumber.Count - 1 do
      begin
        aBomD2 := aBomD2Reader.BomByNumber(Trim(slNumber[iline]));
        if aBomD2 = nil then
        begin
          Memo1.Lines.Add(Trim(slNumber[iline]) + '  没有 BOM 11');
        end;
        slNumber.Objects[iline] := aBomD2;
      end;

      ProgressBar1.Max := aBomD2Reader.FNumbers.Count;
      ProgressBar1.Position := 0;

      irow := 3;
      for inumber := 0 to aBomD2Reader.FNumbers.Count - 1 do
      begin
        ptrBomD2Item := PBomD2Item(aBomD2Reader.FNumbers.Objects[inumber]);

        ExcelApp.Cells[irow, 1].Value := ptrBomD2Item^.snumber_child;
        ExcelApp.Cells[irow, 2].Value := ptrBomD2Item^.sname_child;
        ExcelApp.Cells[irow, 3].Value := ptrBomD2Item^.sabc;
        ExcelApp.Cells[irow, 4].Value := ptrBomD2Item^.sptype;
        ExcelApp.Cells[irow, 5].Value := ptrBomD2Item^.slt;
        ExcelApp.Cells[irow, 6].Value := aBomD2Reader.GetWhereUse(ptrBomD2Item^.snumber_child);
 
        icol := 7;
        for iline := 0 to slNumber.Count - 1 do
        begin
          aBomD2 := TBomD2(slNumber.Objects[iline]);
          if aBomD2 <> nil then
          begin
            ptrBomD2Item_child := aBomD2.ChildByNumber(ptrBomD2Item^.snumber_child);
            if ptrBomD2Item_child <> nil then
            begin
              ExcelApp.Cells[irow, icol].Value := ptrBomD2Item_child^.dusage;

              icol2 := icol + slNumber.Count + 1;
              ExcelApp.Cells[irow, icol2].Value := ptrBomD2Item_child^.dper;
              if not aBomD2.CheckAlloc(ptrBomD2Item_child) then
              begin
                aBomD2.CheckAlloc(ptrBomD2Item_child) ;
                AddColor(ExcelApp, irow, icol2, irow, icol2, clRed);
              end;
                      
              icol3 := icol2 + slNumber.Count + 1;
              if ptrBomD2Item_child^.sgroup <> '' then
              begin
                ExcelApp.Cells[irow, icol3].Value := ptrBomD2Item_child^.sgroup;
              end;   

            end;
          end;
          icol := icol + 1;
        end;
 
        irow := irow + 1;
        ProgressBar1.Position := ProgressBar1.Position + 1;
      end;
     
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存 
    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;

  finally
    aProjNameReader.Free;
    aBomD2Reader.Free;
    slNumber.Free;
  end;

  dwTick := GetTickCount - dwTick;

  MessageBox(Handle, PChar('完成  ' + Format('%0.2f 秒', [dwTick / 1000])), '提示', 0);
end;

end.
