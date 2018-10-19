unit BSFormatWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommUtils, ComObj,
  BSDemandReader, IniFiles;

type
  TfrmBSFormat = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leBSDemand: TLabeledEdit;
    btnBSDemand: TButton;
    Memo1: TMemo;
    ProgressBar1: TProgressBar;
    procedure btnExitClick(Sender: TObject);
    procedure btnBSDemandClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation



{$R *.dfm}

class procedure TfrmBSFormat.ShowForm;
var
  frmBSFormat: TfrmBSFormat;
begin
  frmBSFormat := TfrmBSFormat.Create(nil);
  try
    frmBSFormat.ShowModal;
  finally
    frmBSFormat.Free;
  end;
end;

procedure TfrmBSFormat.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmBSFormat.btnBSDemandClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leBSDemand.Text := sfile;
end;

procedure TfrmBSFormat.btnSave2Click(Sender: TObject);
var
  sfile: string;
  aBSDemandReader: TBSDemandReader;
  aBSDemand: TBSDemand;
  i: Integer;     
  ExcelApp, WorkBook: Variant;
  irow: Integer;
begin

  if not ExcelSaveDialog(sfile) then Exit;

  aBSDemandReader := TBSDemandReader.Create(leBSDemand.Text);
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

    ExcelApp.Sheets[1].Activate;
    ExcelApp.Sheets[1].Name := '产品预测单';    

    try
      ProgressBar1.Max := aBSDemandReader.FList.Count;
      ProgressBar1.Position := 1;

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '编号*';
      ExcelApp.Cells[irow, 2].Value := '需求类型*';
      ExcelApp.Cells[irow, 3].Value := '物料长编码*';
      ExcelApp.Cells[irow, 4].Value := '单位*';
      ExcelApp.Cells[irow, 5].Value := '数量*';
      ExcelApp.Cells[irow, 6].Value := '预测开始日期*';
      ExcelApp.Cells[irow, 7].Value := '预测截止日期*';
      ExcelApp.Cells[irow, 8].Value := '均化周期类型*';
      ExcelApp.Cells[irow, 9].Value := '源单类型*';
      ExcelApp.Cells[irow, 10].Value := '源单号*';
      ExcelApp.Cells[irow, 11].Value := '源单行号*';
      ExcelApp.Cells[irow, 12].Value := '备注*';
      ExcelApp.Cells[irow, 13].Value := '备注2';


      irow := 2;
      for i := 0 to aBSDemandReader.FList.Count -1 do
      begin
        aBSDemand := TBSDemand(aBSDemandReader.FList[i]);

        ExcelApp.Cells[irow, 1].Value := '';
        ExcelApp.Cells[irow, 2].Value := '国内量产';
        ExcelApp.Cells[irow, 3].Value := aBSDemand.FNumber99;
        ExcelApp.Cells[irow, 4].Value := 'PCS';
        ExcelApp.Cells[irow, 5].Value := aBSDemand.FQty;
        ExcelApp.Cells[irow, 6].Value := aBSDemand.FDate;
        ExcelApp.Cells[irow, 7].Value := aBSDemand.FDate;
        ExcelApp.Cells[irow, 8].Value := '不均化';
        ExcelApp.Cells[irow, 9].Value := '';
        ExcelApp.Cells[irow, 10].Value := '';
        ExcelApp.Cells[irow, 11].Value := '';
        ExcelApp.Cells[irow, 12].Value := 'MLBS需求';
        ExcelApp.Cells[irow, 13].Value := '';

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
    aBSDemandReader.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

procedure TfrmBSFormat.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try                                    
    leBSDemand.Text := ini.ReadString(Self.Name, leBSDemand.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmBSFormat.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try                                                     
    ini.WriteString(Self.Name, leBSDemand.Name, leBSDemand.Text); 
  finally
    ini.Free;
  end;
end;

end.
