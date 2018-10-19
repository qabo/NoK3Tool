unit ICMrpResultExecWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, StdCtrls, ExtCtrls, Buttons, ComObj, DateUtils,
  CommUtils, DB, ADODB;

type 
  TfrmICMrpResultExec = class(TForm)
    ToolBar1: TToolBar;
    tbExport: TToolButton;
    SaveDialog1: TSaveDialog;
    ProgressBar1: TProgressBar;
    ADOQuery1: TADOQuery;
    mmoSql: TMemo;
    leRunID: TLabeledEdit;
    procedure tbExportClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }   
  public
    { Public declarations }
    class procedure ShowForm;
  end;



implementation

uses dm;

{$R *.dfm}
              
const
  xlCenter = -4108;
  
var
  fmst: TFormatSettings;
 
class procedure TfrmICMrpResultExec.ShowForm;
var
  frmMPSMergePC: TfrmICMrpResultExec;
begin
  frmMPSMergePC := TfrmICMrpResultExec.Create(nil);
  frmMPSMergePC.ShowModal;
end; 
  
procedure TfrmICMrpResultExec.tbExportClick(Sender: TObject);
var
  ExcelApp, WorkBook: Variant; 
  irow: Integer;
begin
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := 'MRP����' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;

  ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(mmoSql.Text);
  ADOQuery1.SQL.Add(' where t1.FHeadSelfJ0550=' + leRunID.Text + '  ');
  ADOQuery1.Open;

  try

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

    try

      irow := 1;

      ExcelApp.Cells[irow, 1].Value := '������';
      ExcelApp.Cells[irow, 2].Value := '���ϱ���';
      ExcelApp.Cells[irow, 3].Value := '��������';
      ExcelApp.Cells[irow, 4].Value := '�ɹ�Ա';
      ExcelApp.Cells[irow, 5].Value := 'MC';
      ExcelApp.Cells[irow, 6].Value := '�ƻ���������';
      ExcelApp.Cells[irow, 7].Value := '�ƻ�������';
      ExcelApp.Cells[irow, 8].Value := '�ƻ���������';
      ExcelApp.Cells[irow, 9].Value := '�ɹ����뵥����';
      ExcelApp.Cells[irow, 10].Value := '�ɹ����뵥��';
      ExcelApp.Cells[irow, 11].Value := '�ɹ����뵥��¼';
      ExcelApp.Cells[irow, 12].Value := '�ɹ���������';
      ExcelApp.Cells[irow, 13].Value := '�ɹ���������';
      ExcelApp.Cells[irow, 14].Value := '�ɹ�������';
      ExcelApp.Cells[irow, 15].Value := '�ɹ�������¼';
      ExcelApp.Cells[irow, 16].Value := '�ɹ���������';
      ExcelApp.Cells[irow, 17].Value := 'LT';
      ExcelApp.Cells[irow, 18].Value := 'MOQ';
      ExcelApp.Cells[irow, 19].Value := 'SPQ';


      ProgressBar1.Max := ADOQuery1.RecordCount;
      ProgressBar1.Position := 1;
      irow := 2;
      ADOQuery1.First;
      while not ADOQuery1.Eof do
      begin

        ExcelApp.Cells[irow, 1].Value := ADOQuery1.FieldByName('������').AsString;
        ExcelApp.Cells[irow, 2].Value := ADOQuery1.FieldByName('���ϱ���').AsString;
        ExcelApp.Cells[irow, 3].Value := ADOQuery1.FieldByName('��������').AsString;
        ExcelApp.Cells[irow, 4].Value := ADOQuery1.FieldByName('�ɹ�Ա').AsString;
        ExcelApp.Cells[irow, 5].Value := ADOQuery1.FieldByName('MC').AsString;
        ExcelApp.Cells[irow, 6].Value := ADOQuery1.FieldByName('�ƻ���������').AsString;
        ExcelApp.Cells[irow, 7].Value := ADOQuery1.FieldByName('�ƻ�������').AsString;
        ExcelApp.Cells[irow, 8].Value := ADOQuery1.FieldByName('�ƻ���������').AsString;
        ExcelApp.Cells[irow, 9].Value := ADOQuery1.FieldByName('�ɹ����뵥����').AsString;
        ExcelApp.Cells[irow, 10].Value := ADOQuery1.FieldByName('�ɹ����뵥��').AsString;
        ExcelApp.Cells[irow, 11].Value := ADOQuery1.FieldByName('�ɹ����뵥��¼').AsString;
        ExcelApp.Cells[irow, 12].Value := ADOQuery1.FieldByName('�ɹ���������').AsString;
        ExcelApp.Cells[irow, 13].Value := ADOQuery1.FieldByName('�ɹ���������').AsString;
        ExcelApp.Cells[irow, 14].Value := ADOQuery1.FieldByName('�ɹ�������').AsString;
        ExcelApp.Cells[irow, 15].Value := ADOQuery1.FieldByName('�ɹ�������¼').AsString;
        ExcelApp.Cells[irow, 16].Value := ADOQuery1.FieldByName('�ɹ���������').AsString;
        ExcelApp.Cells[irow, 17].Value := ADOQuery1.FieldByName('LT').AsString;
        ExcelApp.Cells[irow, 18].Value := ADOQuery1.FieldByName('MOQ').AsString;
        ExcelApp.Cells[irow, 19].Value := ADOQuery1.FieldByName('SPQ').AsString;
 
        ProgressBar1.Position := ProgressBar1.Position + 1;
        ADOQuery1.Next;
        irow := irow + 1;
      end;
    
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 19] ].Borders.LineStyle := 1; //�ӱ߿�

      WorkBook.SaveAs(SaveDialog1.FileName);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
 
    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
  finally
    ADOQuery1.Close;
  end;

  MessageBox(Handle, '���', '�����ʾ', 0);
end;

procedure TfrmICMrpResultExec.FormCreate(Sender: TObject);
begin
  GetLocaleFormatSettings(0, fmst);
  fmst.DateSeparator := '-'; 
end;                           

end.
