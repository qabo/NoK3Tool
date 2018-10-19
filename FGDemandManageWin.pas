unit FGDemandManageWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ImgList, ComCtrls, ToolWin, DB, ADODB;

type
  TfrmFGDemandManage = class(TForm)
    lbWeeks: TListBox;
    ToolBar1: TToolBar;
    ToolButton7: TToolButton;
    btnManage: TToolButton;
    ToolButton10: TToolButton;
    btnExit: TToolButton;
    ImageList1: TImageList;
    ADOQuery1: TADOQuery;
    procedure btnManageClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm(ADOConnection1: TADOConnection);
  end;


implementation

{$R *.dfm}

class procedure TfrmFGDemandManage.ShowForm(ADOConnection1: TADOConnection); 
var
  frmFGDemandManage: TfrmFGDemandManage;
begin
  frmFGDemandManage := TfrmFGDemandManage.Create(nil);
  try
    frmFGDemandManage.ADOQuery1.Connection := ADOConnection1;
    frmFGDemandManage.ShowModal;
  finally
    frmFGDemandManage.Free;
  end;
end;

procedure TfrmFGDemandManage.btnManageClick(Sender: TObject);
var
  i: Integer;
  cc: Integer;
begin
  if MessageBox(Handle, '确定删除销售计划？', '提示', MB_YESNO) <> IDYES then Exit;

  cc := 0;
  ADOQuery1.Close;
  ADOQuery1.SQL.Clear; 
  for i := 0 to lbWeeks.Items.Count - 1 do
  begin
    if not lbWeeks.Selected[i] then Continue;
    ADOQuery1.SQL.Add('Delete from fgdemand where fid=' + IntToStr(Integer( lbWeeks.Items.Objects[i] )));   
    ADOQuery1.SQL.Add('Delete from fgdemand_entry where fid=' + IntToStr(Integer( lbWeeks.Items.Objects[i] )));
    cc := cc + 1;
  end;

  if cc = 0 then
  begin
    MessageBox(Handle, '请选择要删除的week', '提示', 0);
    Exit;
  end;

  ADOQuery1.Connection.Connected := True;
  ADOQuery1.ExecSQL;    
  ADOQuery1.Connection.Connected := False;

  lbWeeks.DeleteSelected;

  MessageBox(Handle, '删除成功', '提示', 0);
end;

procedure TfrmFGDemandManage.FormShow(Sender: TObject);
var
  sname: string;
  id: Integer;
begin
  lbWeeks.Clear;

  ADOQuery1.Connection.Connected := True;

  ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(' select fid, fname from fgdemand order by fid ');
  ADOQuery1.Open;
  ADOQuery1.First;
  while not ADOQuery1.Eof do
  begin
    sname := ADOQuery1.FieldByName('fname').AsString;
    id := ADOQuery1.FieldByName('fid').AsInteger;
    lbWeeks.Items.AddObject(sname, TObject(id));
    ADOQuery1.Next;
  end;
  ADOQuery1.Close;

  ADOQuery1.Connection.Connected := False;
end;

procedure TfrmFGDemandManage.btnExitClick(Sender: TObject);
begin
  Close;
end;

end.
