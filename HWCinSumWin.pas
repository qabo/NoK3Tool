unit HWCinSumWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommUtils;

type
  TfrmHWCinSum = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leCin: TLabeledEdit;
    btnCin: TButton;
    Memo1: TMemo;
    procedure btnCinClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

type
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

class procedure TfrmHWCinSum.ShowForm; 
var
  frmHWCinSum: TfrmHWCinSum;
begin
  frmHWCinSum := TfrmHWCinSum.Create(nil);
  try
    frmHWCinSum.ShowModal;
  finally
    frmHWCinSum.Free;
  end;
end;

procedure TfrmHWCinSum.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmHWCinSum.btnCinClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leCin.Text := sfile;
end;

procedure TfrmHWCinSum.btnSave2Click(Sender: TObject);
const
  CSTitle = '日期单据编号收料仓库加工材料长代码加工材料名称实收数量备注订单单号';
var
  ExcelApp, WorkBook: Variant;       
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;
  //sFile: string;
  stitle, stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7, stitle8: string;
  irow: Integer;
  snumber: string;
  sdate: string;
  sbillno: string;
  sstock: string;
  sname: string;
  dqty: Double;
  snote: string;
  sorderno: string;
  lst: TList;
  i: Integer;
  p: PRowRecord;
  pSum: PRowRecord;
  lstSum: TStringList;
  isum: Integer;
  sfile: string;
begin
  
  if not ExcelSaveDialog(sfile) then Exit;

  if not FileExists(leCin.Text) then
  begin
    //MessageBox(Form1.Handle, '文件不存在', '金蝶提示', 0);
    Exit;
  end;        
  lst := TList.Create;
  lstSum := TStringList.Create;
 
                                           
  p := New(PRowRecord);
  p^.dqty := 0;
  lstSum.AddObject('蓝色16GB', TObject(p));

  p := New(PRowRecord);
  p^.dqty := 0;
  lstSum.AddObject('黑色16GB', TObject(p));

  p := New(PRowRecord);
  p^.dqty := 0;
  lstSum.AddObject('金色16GB', TObject(p));
             
  p := New(PRowRecord);
  p^.dqty := 0;
  lstSum.AddObject('蓝色32GB', TObject(p));

  p := New(PRowRecord);
  p^.dqty := 0;
  lstSum.AddObject('黑色32GB', TObject(p));
      
  p := New(PRowRecord);
  p^.dqty := 0;
  lstSum.AddObject('金色32GB', TObject(p));


  
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';


    try
      WorkBook := ExcelApp.WorkBooks.Open(leCin.Text);

      try
        iSheetCount := ExcelApp.Sheets.Count;
        for iSheet := 1 to iSheetCount do
        begin
          if not ExcelApp.Sheets[iSheet].Visible then Continue;

          ExcelApp.Sheets[iSheet].Activate;
                              
          sSheet := ExcelApp.Sheets[iSheet].Name;

          stitle1 := ExcelApp.Cells[1, 1].Value;
          stitle2 := ExcelApp.Cells[1, 2].Value;
          stitle3 := ExcelApp.Cells[1, 3].Value;
          stitle4 := ExcelApp.Cells[1, 4].Value;
          stitle5 := ExcelApp.Cells[1, 5].Value;
          stitle6 := ExcelApp.Cells[1, 6].Value;
          stitle7 := ExcelApp.Cells[1, 7].Value;
          stitle8 := ExcelApp.Cells[1, 8].Value;

          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8;

          if stitle <> CSTitle then
          begin
            Memo1.Lines.Add(sSheet + ' 格式不符合');
            Continue;
          end;

          Memo1.Lines.Add(sSheet + ' 开始读取数据');

          irow := 2;
        
          snumber := ExcelApp.Cells[irow, 4].Value;
          while snumber <> '' do
          begin
            sdate := ExcelApp.Cells[irow, 1].Value;
            sbillno := ExcelApp.Cells[irow, 2].Value;
            sstock := ExcelApp.Cells[irow, 3].Value;
            sname := ExcelApp.Cells[irow, 5].Value;
            dqty := ExcelApp.Cells[irow, 6].Value;
            snote := ExcelApp.Cells[irow, 7].Value;
            sorderno := ExcelApp.Cells[irow, 8].Value;

            p := New(PRowRecord);
            p^.sdate := sdate;
            p^.sbillno := sbillno;
            p^.sstock := sstock;
            p^.snumber := snumber;
            p^.sname := sname;
            p^.dqty := dqty;
            p^.snote := snote;
            p^.sorderno := sorderno;
            lst.Add(p);

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 4].Value;
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


    for i := 0 to lst.Count - 1 do
    begin
      p := PRowRecord(lst[i]);
      if Pos('-', p^.snumber) > 0 then    // 带 '-'
      begin
        if Pos('-T', p^.snumber) = 0 then  // 非 'T'
        begin
          for isum := 0 to lstSum.Count - 1 do
          begin
            if Pos(lstSum[isum], p^.sname) > 0 then
            begin
              pSum := PRowRecord(lstSum.Objects[isum]);
              pSum^.dqty := pSum^.dqty + p^.dqty;
              Break;
            end; 
          end;
        end;
      end
      else     // 不不不带 '-'
      begin
        if Pos('欧规', p^.sname) > 0 then  // 非 'T'
        begin
          for isum := 0 to lstSum.Count - 1 do
          begin
            if Pos(lstSum[isum], p^.sname) > 0 then
            begin
              pSum := PRowRecord(lstSum.Objects[isum]);
              pSum^.dqty := pSum^.dqty + p^.dqty;
              Break;
            end; 
          end;
        end;
      end;
    end;


     


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
      ExcelApp.Cells[irow, 1].Value := '颜色容量';
      ExcelApp.Cells[irow, 2].Value := '数量';

      irow := 2;

      for isum := 0 to lstSum.Count - 1 do
      begin               
        pSum := PRowRecord(lstSum.Objects[isum]);
        ExcelApp.Cells[irow, 1].Value := lstSum[isum];
        ExcelApp.Cells[irow, 2].Value := psum^.dqty;
        irow := irow + 1;
      end;

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 2] ].Interior.Color := $DBDCF2;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 2] ].HorizontalAlignment := xlCenter;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, 2] ].Borders.LineStyle := 1; //加边框

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

    MessageBox(self.Handle, '完成', '提示', 0);
    
  finally
    for i := 0 to lst.Count - 1 do
    begin
      p := PRowRecord(lst[i]);
      Dispose(p);
    end;  
    lst.Free;

    for i := 0 to lstSum.Count - 1 do
    begin
      p := PRowRecord(lstSum.Objects[i]);
      Dispose(p);
    end;
    lstSum.Free;
  end;
end;

end.
