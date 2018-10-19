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

    ExcelApp.Sheets[1].Activate;
    ExcelApp.Sheets[1].Name := '��ƷԤ�ⵥ';    

    try
      ProgressBar1.Max := aBSDemandReader.FList.Count;
      ProgressBar1.Position := 1;

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '���*';
      ExcelApp.Cells[irow, 2].Value := '��������*';
      ExcelApp.Cells[irow, 3].Value := '���ϳ�����*';
      ExcelApp.Cells[irow, 4].Value := '��λ*';
      ExcelApp.Cells[irow, 5].Value := '����*';
      ExcelApp.Cells[irow, 6].Value := 'Ԥ�⿪ʼ����*';
      ExcelApp.Cells[irow, 7].Value := 'Ԥ���ֹ����*';
      ExcelApp.Cells[irow, 8].Value := '������������*';
      ExcelApp.Cells[irow, 9].Value := 'Դ������*';
      ExcelApp.Cells[irow, 10].Value := 'Դ����*';
      ExcelApp.Cells[irow, 11].Value := 'Դ���к�*';
      ExcelApp.Cells[irow, 12].Value := '��ע*';
      ExcelApp.Cells[irow, 13].Value := '��ע2';


      irow := 2;
      for i := 0 to aBSDemandReader.FList.Count -1 do
      begin
        aBSDemand := TBSDemand(aBSDemandReader.FList[i]);

        ExcelApp.Cells[irow, 1].Value := '';
        ExcelApp.Cells[irow, 2].Value := '��������';
        ExcelApp.Cells[irow, 3].Value := aBSDemand.FNumber99;
        ExcelApp.Cells[irow, 4].Value := 'PCS';
        ExcelApp.Cells[irow, 5].Value := aBSDemand.FQty;
        ExcelApp.Cells[irow, 6].Value := aBSDemand.FDate;
        ExcelApp.Cells[irow, 7].Value := aBSDemand.FDate;
        ExcelApp.Cells[irow, 8].Value := '������';
        ExcelApp.Cells[irow, 9].Value := '';
        ExcelApp.Cells[irow, 10].Value := '';
        ExcelApp.Cells[irow, 11].Value := '';
        ExcelApp.Cells[irow, 12].Value := 'MLBS����';
        ExcelApp.Cells[irow, 13].Value := '';

        irow := irow + 1;
        ProgressBar1.Position := ProgressBar1.Position + 1;
      end;
      

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end; 
    
  finally
    aBSDemandReader.Free;
  end;

  MessageBox(Handle, '���', '��ʾ', 0);
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
