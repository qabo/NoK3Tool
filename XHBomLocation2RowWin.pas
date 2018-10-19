unit XHBomLocation2RowWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XHBomReader, ImgList, ComCtrls, ToolWin, CommUtils, StdCtrls,
  ExtCtrls, ComObj;

type
  TfrmXHBomLocation2Row = class(TForm)
    ToolBar1: TToolBar;
    ToolButton5: TToolButton;
    tbClose: TToolButton;
    ToolButton1: TToolButton;
    ImageList1: TImageList;
    tbSave: TToolButton;
    leBom: TLabeledEdit;
    btnBom: TButton;
    procedure tbCloseClick(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure btnBomClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmXHBomLocation2Row.ShowForm;
var
  frmXHBomLocation2Row: TfrmXHBomLocation2Row;
begin
  frmXHBomLocation2Row := TfrmXHBomLocation2Row.Create(nil);
  try
    frmXHBomLocation2Row.ShowModal;
  finally
    frmXHBomLocation2Row.Free;
  end;
end;

procedure TfrmXHBomLocation2Row.tbCloseClick(Sender: TObject);
begin
  Close;
end;
    
procedure TfrmXHBomLocation2Row.btnBomClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leBom.Text := sfile;
end;

procedure TfrmXHBomLocation2Row.tbSaveClick(Sender: TObject);
var                    
  ExcelApp, WorkBook: Variant;
  sfile: string;
  i: Integer;
  aXHBomReader: TXHBomReader;
  ptrXHBomChilcRecord: PXHBomChilcRecord;
  irow: Integer;
  sl: TStringList;
  ilac: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  aXHBomReader := TXHBomReader.Create(leBom.Text);
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

    sl := TStringList.Create;

    try
      irow := 1;

      ExcelApp.Cells[irow, 1].Value :=  '位置号';
      ExcelApp.Cells[irow, 2].Value :=  '子项物料代码';
      ExcelApp.Cells[irow, 3].Value :=  '物料名称';
      ExcelApp.Cells[irow, 4].Value :=  '规格型号';
              
      irow := 2;
      for i := 0 to aXHBomReader.Count - 1 do
      begin
        ptrXHBomChilcRecord := aXHBomReader.Items[i];

        sl.Text := StringReplace(ptrXHBomChilcRecord^.slocation, ',', #13#10, [rfReplaceAll]);
        if sl.Count = 0 then
        begin
          ExcelApp.Cells[irow, 1].Value :=  '';
          ExcelApp.Cells[irow, 2].Value :=  ptrXHBomChilcRecord^.snumber;
          ExcelApp.Cells[irow, 3].Value :=  ptrXHBomChilcRecord^.sname;
          ExcelApp.Cells[irow, 4].Value :=  ptrXHBomChilcRecord^.smodel;
          irow := irow + 1;
        end
        else
        begin
          for ilac := 0 to sl.Count -1 do
          begin
            ExcelApp.Cells[irow, 1].Value :=  sl[ilac];
            ExcelApp.Cells[irow, 2].Value :=  ptrXHBomChilcRecord^.snumber;
            ExcelApp.Cells[irow, 3].Value :=  ptrXHBomChilcRecord^.sname;
            ExcelApp.Cells[irow, 4].Value :=  ptrXHBomChilcRecord^.smodel;
            irow := irow + 1;
          end;
        end; 
      end;



      ExcelApp.Sheets[1].Activate;
      
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;    
  finally
    aXHBomReader.Free;

    sl.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
