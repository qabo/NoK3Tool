unit SelectFGDemandWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ADODB, StdCtrls, DB, ComCtrls;

type
  TfrmSelectFGDemand = class(TForm)
    lbWeeks: TListBox;
    ADOQuery1: TADOQuery;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    cbProjs: TComboBox;
    DateTimePicker1: TDateTimePicker;
    GroupBox2: TGroupBox;
    btnOK: TButton;
    btnCancel: TButton;
    procedure FormCreate(Sender: TObject);
    procedure cbProjsChange(Sender: TObject);
  private
    { Private declarations }
    procedure InitList(ADOConnection1: TADOConnection);
  public
    { Public declarations }
    class function GetFGPlans(ADOConnection1: TADOConnection;
      var sproj: string; var dt: TDateTime; slFGPlans: TStringList): Boolean;
  end;


implementation

{$R *.dfm}

class function TfrmSelectFGDemand.GetFGPlans(ADOConnection1: TADOConnection;
  var sproj: string; var dt: TDateTime; slFGPlans: TStringList): Boolean;
var
  frmSelectFGDemand: TfrmSelectFGDemand;
  i: Integer;
  mr: TModalResult;
begin
  Result := False;
  
  slFGPlans.Clear;

  frmSelectFGDemand := TfrmSelectFGDemand.Create(nil);
  try
    frmSelectFGDemand.InitList(ADOConnection1);
    mr := frmSelectFGDemand.ShowModal;
    if mr <> mrOk then Exit;
    Result := True;
    sproj := frmSelectFGDemand.cbProjs.Text;
    dt := frmSelectFGDemand.DateTimePicker1.DateTime;
    for i := 0 to frmSelectFGDemand.lbWeeks.Items.Count - 1 do
    begin
      if frmSelectFGDemand.lbWeeks.Selected[i] then
      slFGPlans.AddObject(frmSelectFGDemand.lbWeeks.Items[i], frmSelectFGDemand.lbWeeks.Items.Objects[i]);
    end;
  finally
    frmSelectFGDemand.Free;
  end;
end;

procedure TfrmSelectFGDemand.InitList(ADOConnection1: TADOConnection);
begin
  cbProjs.Clear;

  ADOQuery1.Close;
  ADOQuery1.Connection := ADOConnection1;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(' select distinct fproj from fgdemand_entry order by fproj ');
  ADOQuery1.Open;
  ADOQuery1.First;
  while not ADOQuery1.Eof do
  begin
    cbProjs.Items.Add(ADOQuery1.FieldByName('fproj').AsString);
    ADOQuery1.Next;
  end;
  ADOQuery1.Close;

  if cbProjs.Items.Count = 0 then Exit;
  cbProjs.ItemIndex := 0;

  cbProjsChange(cbProjs);
end;

procedure TfrmSelectFGDemand.FormCreate(Sender: TObject);
begin
  DateTimePicker1.DateTime := Now;
  ModalResult := mrCancel;
end;

procedure TfrmSelectFGDemand.cbProjsChange(Sender: TObject);
begin       
  lbWeeks.Clear;
  
  ADOQuery1.Close; 
  ADOQuery1.SQL.Clear;

  ADOQuery1.SQL.Add(' select distinct t1.fid, t1.fname                 ');
  ADOQuery1.SQL.Add(' from fgdemand t1                                 ');
  ADOQuery1.SQL.Add(' inner join fgdemand_entry t2 on t1.fid=t2.fid    ');
  ADOQuery1.SQL.Add(' where t2.fproj=''' + cbProjs.Text + '''          ');
  ADOQuery1.SQL.Add(' order by t1.fid                                  ');


  ADOQuery1.Open;
  ADOQuery1.First;
  while not ADOQuery1.Eof do
  begin
    lbWeeks.Items.AddObject(ADOQuery1.FieldByName('fname').AsString, TObject(ADOQuery1.FieldByName('fid').AsInteger));
    ADOQuery1.Next;
  end;
  ADOQuery1.Close;
end;

end.
