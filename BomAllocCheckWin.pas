unit BomAllocCheckWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, CommUtils, ExtCtrls, StdCtrls, ImgList, ComCtrls,
  ToolWin, ExcelBomReader, RawMPSReader, NumberCheckReader, SAPBomReader,
  ComObj, CommVars;

type
  TfrmBomAllocCheck = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ImageList1: TImageList;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    btnSapFcst: TButton;
    leSapFcst: TLabeledEdit;
    Memo1: TMemo;
    leExcelBom: TLabeledEdit;
    btnExcelom: TButton;
    TabSheet2: TTabSheet;
    leSAPBom: TLabeledEdit;
    btnSAPBom: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnExcelomClick(Sender: TObject);
    procedure btnSapFcstClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSAPBomClick(Sender: TObject);
  private
    { Private declarations }
    procedure OnLogEvent(const s: string);
    procedure btnSave2Click0;     
    procedure btnSave2Click1;
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

type
  TBomLocCheckerLine = packed record
    snumber: string;
    sname: string;
    bXls: Boolean;
    bSap: Boolean;
    dPer: Double;
  end;
  PBomLocCheckerLine = ^TBomLocCheckerLine;

  TBomLocChecker = class
  private
    FList: TList;
    susage_xls: string;
    susage_sap: string;
    procedure Clear;
    function FindBomLocCheckerLine(const sNumber: string): PBomLocCheckerLine;
    procedure AddFromSap(aSapBom: TSapBom);
  public
    constructor Create;
    destructor Destroy; override;
    procedure CopyFromXls(aBomLoc: TBomLoc);
    procedure AssignSapBom(aSapBom: TSapBom);
  end;

constructor TBomLocChecker.Create;
begin
  FList := TList.Create;
  inherited;
end;

destructor TBomLocChecker.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TBomLocChecker.Clear;
var
  i: Integer;
  p: PBomLocCheckerLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomLocCheckerLine(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TBomLocChecker.FindBomLocCheckerLine(const sNumber: string): PBomLocCheckerLine;
var
  i: Integer;
  p: PBomLocCheckerLine;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomLocCheckerLine(FList[i]);
    if p^.snumber = sNumber then
    begin
      Result := p;
      Break;
    end;
  end;
end;

procedure TBomLocChecker.CopyFromXls(aBomLoc: TBomLoc);
var
  i: Integer;
  aBomLocCheckerLinePtr: PBomLocCheckerLine;
  aBomLocLinePtr: PBomLocLine;
begin
  susage_xls := aBomLoc.susage;
  for i := 0 to aBomLoc.FList.Count - 1 do
  begin
    aBomLocLinePtr := PBomLocLine(aBomLoc.FList.Objects[i]);

    aBomLocCheckerLinePtr := New(PBomLocCheckerLine);
    aBomLocCheckerLinePtr^.snumber := aBomLocLinePtr^.snumber;
    aBomLocCheckerLinePtr^.sname := aBomLocLinePtr^.sname;
    aBomLocCheckerLinePtr^.bXls := True;       
    aBomLocCheckerLinePtr^.bSap := False;     
    aBomLocCheckerLinePtr^.dPer := 0;
    FList.Add(aBomLocCheckerLinePtr);
  end;
end;

procedure TBomLocChecker.AddFromSap(aSapBom: TSapBom);
var
  aBomLocCheckerLinePtr: PBomLocCheckerLine;
begin
  aBomLocCheckerLinePtr := New(PBomLocCheckerLine);
  aBomLocCheckerLinePtr^.snumber := aSapBom.FNumber;
  aBomLocCheckerLinePtr^.sname := aSapBom.FName;   
  aBomLocCheckerLinePtr^.bXls := False;
  aBomLocCheckerLinePtr^.bSap := True;
  aBomLocCheckerLinePtr^.dPer := aSapBom.dPer;
  FList.Add(aBomLocCheckerLinePtr);
end;

procedure TBomLocChecker.AssignSapBom(aSapBom: TSapBom);
var
  iChild: Integer;
  aSapItemGroup: TSapItemGroup;
  iItem: Integer;
  aSapBomItem: TSapBom;
  bFound: Boolean;
  aBomLocCheckerLinePtr: PBomLocCheckerLine;
begin
  for iChild := 0 to aSapBom.ChildCount - 1 do
  begin
    aSapItemGroup := aSapBom.Childs[iChild];
    if aSapItemGroup.ItemCount = 0 then Continue; // 正常情况不会等于0
    
    aSapBomItem := aSapItemGroup.Items[0];
    if aSapBomItem.ChildCount = 0 then // 叶子节点
    begin
      bFound := False;
      for iItem := 0 to aSapItemGroup.ItemCount - 1 do
      begin
        aSapBomItem := aSapItemGroup.Items[iItem];
        aBomLocCheckerLinePtr := FindBomLocCheckerLine(aSapBomItem.FNumber);
        if aBomLocCheckerLinePtr <> nil then
        begin
          bFound := True;
          Break;
        end;  
      end;
      
      if bFound then
      begin
        susage_sap := FloatToStr(aSapBom.dusage);
        
        for iItem := 0 to aSapItemGroup.ItemCount - 1 do
        begin
          aSapBomItem := aSapItemGroup.Items[iItem];
          aBomLocCheckerLinePtr := FindBomLocCheckerLine(aSapBomItem.FNumber);
          if aBomLocCheckerLinePtr = nil then
          begin
            AddFromSap(aSapBomItem);
          end
          else
          begin
            aBomLocCheckerLinePtr.bSap := True;
            aBomLocCheckerLinePtr.dPer := aSapBomItem.dPer;
          end;  
        end;
      end;
    end
    else
    begin
      for iItem := 0 to aSapItemGroup.ItemCount - 1 do
      begin
        aSapBomItem := aSapItemGroup.Items[iItem];
        AssignSapBom(aSapBomItem);
      end;
    end;   
  end;
end;

{ TfrmBomAllocCheck }

class procedure TfrmBomAllocCheck.ShowForm;
var
  frmBomAllocCheck: TfrmBomAllocCheck;
begin
  frmBomAllocCheck := TfrmBomAllocCheck.Create(nil);
  try
    frmBomAllocCheck.ShowModal;
  finally
    frmBomAllocCheck.Free;
  end;
end;

procedure TfrmBomAllocCheck.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leExcelBom.Text := ini.ReadString(self.ClassName, leExcelBom.Name, '');
    leSapFcst.Text := ini.ReadString(self.ClassName, leSapFcst.Name, '');
    leSAPBom.Text := ini.ReadString(self.ClassName, leSAPBom.Name, '');
    PageControl1.ActivePageIndex := ini.ReadInteger(self.ClassName, PageControl1.Name, 0);
  finally
    ini.Free;
  end;
end;

procedure TfrmBomAllocCheck.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leExcelBom.Name, leExcelBom.Text);
    ini.WriteString(self.ClassName, leSapFcst.Name, leSapFcst.Text);
    ini.WriteString(self.ClassName, leSAPBom.Name, leSAPBom.Text);
    ini.WriteInteger(self.ClassName, PageControl1.Name, PageControl1.ActivePageIndex);
  finally
    ini.Free;
  end;
end;
          
procedure TfrmBomAllocCheck.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmBomAllocCheck.btnExcelomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leExcelBom.Text := sfile;
end;

procedure TfrmBomAllocCheck.btnSapFcstClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSapFcst.Text := sfile;
end;

procedure TfrmBomAllocCheck.btnSAPBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;

procedure TfrmBomAllocCheck.btnSave2Click0;
var
//  sfile: string;    
  aExcelBomReader: TExcelBomReader;
  aRawMPSReader: TRawMPSReader;
  slcol: TStringList;
  slcap: TStringList;
  slver: TStringList;
  iBomLoc: Integer;
  aBomLoc: TBomLoc;
  iBomLocLine: Integer;
  aBomLocLinePtr: PBomLocLine;

  iRawBom: Integer;
  aRawBomPtr: PRawBom;
  i: Integer;
  s: string;
  sl: TStringList;
  iccMPS: Integer;
  iccBom: Integer;
  idx: Integer;
 
begin
//  if not ExcelSaveDialog(sfile) then Exit;
  Memo1.Clear;

  aExcelBomReader := TExcelBomReader.Create(leExcelBom.Text);
  aRawMPSReader := TRawMPSReader.Create(leSapFcst.Text);
  try
    slcol := TStringList.Create;
    slcap := TStringList.Create;
    slver := TStringList.Create;

    for iBomLoc := 0 to aExcelBomReader.FList.Count - 1 do
    begin
      aBomLoc := TBomLoc(aExcelBomReader.FList.Objects[iBomLoc]);
      for iBomLocLine := 0 to aBomLoc.FList.Count - 1 do
      begin
        aBomLocLinePtr := PBomLocLine(aBomLoc.FList.Objects[iBomLocLine]);
                                    
        s := aBomLocLinePtr^.scol;
        s := StringReplace(s, ',', #13#10, [rfReplaceAll]);
        s := StringReplace(s, '，', #13#10, [rfReplaceAll]);
        sl := TStringList.Create;
        sl.Text := s;
        for i := 0 to sl.Count - 1 do
        begin
          if slcol.IndexOf(sl[i]) < 0 then
          begin
            slcol.Add(sl[i]);
          end;
        end;
        sl.Free;
                                  
        s := aBomLocLinePtr^.scap;
        s := StringReplace(s, ',', #13#10, [rfReplaceAll]);
        s := StringReplace(s, '，', #13#10, [rfReplaceAll]);    
        sl := TStringList.Create;
        sl.Text := s;
        for i := 0 to sl.Count - 1 do
        begin
          if slcap.IndexOf(sl[i]) < 0 then
          begin
            slcap.Add(sl[i]);
          end;
        end;
        sl.Free;
                                                 
        s := aBomLocLinePtr^.sver;
        s := StringReplace(s, ',', #13#10, [rfReplaceAll]);
        s := StringReplace(s, '，', #13#10, [rfReplaceAll]);     
        sl := TStringList.Create;
        sl.Text := s;
        for i := 0 to sl.Count - 1 do
        begin
          if slver.IndexOf(sl[i]) < 0 then
          begin
            slver.Add(sl[i]);
          end;
        end;
        sl.Free;
      end;
    end;

    iccMPS := 0;
    for iRawBom := 0 to aRawMPSReader.slBomNumber.Count - 1 do
    begin
      aRawBomPtr := PRawBom(aRawMPSReader.slBomNumber.Objects[iRawBom]);
      idx := slcol.IndexOf(aRawBomPtr^.scol);
      if idx < 0 then
      begin
        idx := slcol.IndexOfName(aRawBomPtr^.scol);
      end;
      if idx < 0 then
      begin
        Memo1.Lines.Add('MPS 颜色在BOM找不到: ' + aRawBomPtr^.scol);
        iccMPS := iccMPS + 1;
      end
      else
      begin
        slcol[idx] := aRawBomPtr^.scol + '=1';
      end;
    end;

    for iRawBom := 0 to aRawMPSReader.slBomNumber.Count - 1 do
    begin
      aRawBomPtr := PRawBom(aRawMPSReader.slBomNumber.Objects[iRawBom]);
      idx :=slcap.IndexOf(aRawBomPtr^.scap);
      if idx < 0 then
      begin
        idx := slcap.IndexOfName(aRawBomPtr^.scap);
      end;
      if idx < 0 then
      begin
        Memo1.Lines.Add('MPS 容量在BOM找不到: ' + aRawBomPtr^.scap);
        iccMPS := iccMPS + 1;
      end
      else
      begin
        slcap[idx] := aRawBomPtr^.scap + '=1'; 
      end;  
    end;           
      
    for iRawBom := 0 to aRawMPSReader.slBomNumber.Count - 1 do
    begin
      aRawBomPtr := PRawBom(aRawMPSReader.slBomNumber.Objects[iRawBom]);
      idx := slver.IndexOf(aRawBomPtr^.sver);
      if idx < 0 then
      begin
        idx := slver.IndexOfName(aRawBomPtr^.sver);
      end;
      if idx < 0 then
      begin
        Memo1.Lines.Add('MPS 版本在BOM找不到: ' + aRawBomPtr^.sver);
        iccMPS := iccMPS + 1;
      end
      else
      begin
        slver[idx] := aRawBomPtr^.sver + '=1';
      end;  
    end;

    Memo1.Lines.Add('');

    iccBom := 0;
    for i := 0 to slcol.Count - 1 do
    begin
      if slcol[i] = '通用' then Continue;
      if slcol.ValueFromIndex[i] <> '1' then
      begin
        Memo1.Lines.Add('Excel Bom 颜色在MPS找不到: ' + slcol[i]);
        iccBom := iccBom + 1;      
      end;
    end;

    for i := 0 to slcap.Count - 1 do
    begin                                
      if slcap[i] = '通用' then Continue;
      if slcap.ValueFromIndex[i] <> '1' then
      begin
        Memo1.Lines.Add('Excel Bom 容量在MPS找不到: ' + slcap[i]);
        iccBom := iccBom + 1;
      end;
    end;

    for i := 0 to slver.Count - 1 do
    begin     
      if slver[i] = '通用' then Continue;
      if slver.ValueFromIndex[i] <> '1' then
      begin
        Memo1.Lines.Add('Excel Bom 版本在MPS找不到: ' + slver[i]);
        iccBom := iccBom + 1;
      end;
    end;  

    slcol.Free;
    slcap.Free;
    slver.Free;
  finally
    aExcelBomReader.Free;
    aRawMPSReader.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

procedure TfrmBomAllocCheck.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;  

procedure TfrmBomAllocCheck.btnSave2Click1;
var
  sfile: string;
  aExcelBomReader: TExcelBomReader;
  aSAPBomReader: TSAPBomReader;
  aList: TList;
  iBomLoc: Integer;
  aBomLoc: TBomLoc;
  i: Integer;
  aBomLocChecker: TBomLocChecker;

  iSapBom: Integer;
  aSapBom: TSapBom;
  
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  iLine: Integer;
  aBomLocCheckerLinePtr: PBomLocCheckerLine;
  dPerSum: Double;
begin
  if not ExcelSaveDialog(sfile) then Exit;


  aList := TList.Create;

  aExcelBomReader := TExcelBomReader.Create(leExcelBom.Text);
  aSAPBomReader := TSAPBomReader.Create(leSAPBom.Text, OnLogEvent);
  try
    for iBomLoc := 0 to aExcelBomReader.FList.Count - 1 do
    begin
      aBomLoc := TBomLoc(aExcelBomReader.FList.Objects[iBomLoc]);

      aBomLocChecker := TBomLocChecker.Create;
      aList.Add(aBomLocChecker);

      aBomLocChecker.CopyFromXls(aBomLoc);
    end;

    for i := 0 to aList.Count - 1 do
    begin
      aBomLocChecker := TBomLocChecker(aList[i]);
      for iSapBom := 0 to aSAPBomReader.FList.Count - 1 do
      begin
        aSapBom := TSapBom(aSAPBomReader.FList.Objects[iSapBom]);
        aBomLocChecker.AssignSapBom(aSapBom);
      end;
    end;




    // 开始保存 Excel
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := True;
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
      ExcelApp.Cells[irow, 1].Value := '物料编码';
      ExcelApp.Cells[irow, 2].Value := '物料名称';
      ExcelApp.Cells[irow, 3].Value := 'Excel用量';
      ExcelApp.Cells[irow, 4].Value := 'Sap用量';
      ExcelApp.Cells[irow, 5].Value := 'Excel';
      ExcelApp.Cells[irow, 6].Value := 'Sap';   
      ExcelApp.Cells[irow, 7].Value := '配比';
      ExcelApp.Cells[irow, 8].Value := '配比总和';

      ExcelApp.Columns[1].ColumnWidth := 21;
      ExcelApp.Columns[2].ColumnWidth := 34;
//      ExcelApp.Cells[irow, 3].Value := 'Excel用量';
//      ExcelApp.Cells[irow, 4].Value := 'Sap用量';
//      ExcelApp.Cells[irow, 5].Value := 'Excel';
//      ExcelApp.Cells[irow, 6].Value := 'Sap';   
//      ExcelApp.Cells[irow, 7].Value := '配比';
//      ExcelApp.Cells[irow, 8].Value := '配比总和';

      irow := irow + 1;

      for i := 0 to aList.Count - 1 do
      begin
        aBomLocChecker := TBomLocChecker(aList[i]);

        if aBomLocChecker.FList.Count = 0 then Continue;

        dPerSum := 0;
        for iLine := 0 to aBomLocChecker.FList.Count - 1 do
        begin                                          
          aBomLocCheckerLinePtr := PBomLocCheckerLine(aBomLocChecker.FList[iLine]);
          ExcelApp.Cells[irow + iLine, 1].Value := aBomLocCheckerLinePtr^.snumber;
          ExcelApp.Cells[irow + iLine, 2].Value := aBomLocCheckerLinePtr^.sname;
          ExcelApp.Cells[irow + iLine, 5].Value := CSBoolean[aBomLocCheckerLinePtr^.bXls];
          ExcelApp.Cells[irow + iLine, 6].Value := CSBoolean[aBomLocCheckerLinePtr^.bSap];
          ExcelApp.Cells[irow + iLine, 7].Value := aBomLocCheckerLinePtr^.dPer;
          dPerSum := dPerSum + aBomLocCheckerLinePtr^.dPer;
        end;    
         
        ExcelApp.Cells[irow, 3].Value := aBomLocChecker.susage_xls;
        ExcelApp.Cells[irow, 4].Value := aBomLocChecker.susage_sap;
        ExcelApp.Cells[irow, 8].Value := dPerSum;

        MergeCells(ExcelApp, irow, 3, irow + aBomLocChecker.FList.Count - 1, 3); 
        MergeCells(ExcelApp, irow, 4, irow + aBomLocChecker.FList.Count - 1, 4);
        MergeCells(ExcelApp, irow, 8, irow + aBomLocChecker.FList.Count - 1, 8);

        irow := irow + aBomLocChecker.FList.Count;
      end;

      AddBorder(ExcelApp, 1, 1, irow - 1, 8);
        
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

  finally
    aExcelBomReader.Free;
    aSAPBomReader.Free;

    for i := 0 to aList.Count - 1 do
    begin
      aBomLocChecker := TBomLocChecker(aList[i]);
      aBomLocChecker.Free;
    end;
    aList.Clear;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

procedure TfrmBomAllocCheck.btnSave2Click(Sender: TObject);
begin
  case PageControl1.ActivePageIndex of
    0: btnSave2Click0();
    1: btnSave2Click1();
  end;
end;

end.
