unit ProjOfICItemWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, CommUtils, ComCtrls, ToolWin, ImgList, ComObj;

type
  TfrmProjOfICItem = class(TForm)
    leProjOfICItem: TLabeledEdit;
    leProjEODate: TLabeledEdit;
    btnProjOfICItem: TButton;
    btnProjEODate: TButton;
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    Memo1: TMemo;
    procedure btnProjOfICItemClick(Sender: TObject);
    procedure btnProjEODateClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}


type
  TProjOfICItemReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
  public
    FList: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
  end;

  TProjEODateReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
  public
    FList: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
  end;

{ TProjOfICItemReader }

constructor TProjOfICItemReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TProjOfICItemReader.Destroy;
begin
  FList.Free;
end;

procedure TProjOfICItemReader.Clear;
begin
  FList.Clear;
end;

procedure TProjOfICItemReader.Log(const str: string);
begin

end;

procedure TProjOfICItemReader.Open;
const
  CINumber = 1;   //物料编码
  CIName = 2;     //	物料名称
  CIWhereUse = 3; //	Where USE

var
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  irow: Integer;
  snumber: string;
  sWhereUse: string;
begin
  Clear;


  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        Log(sSheet);

        irow := 2;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        while snumber <> '' do
        begin
          sWhereUse := ExcelApp.Cells[irow, CIWhereUse].Value;

          FList.Add(snumber + '=' + sWhereUse);

          irow := irow + 1;    
          snumber := ExcelApp.Cells[irow, CINumber].Value;
        end;
        Break;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end;  
end;

{ TProjEODateReader }

constructor TProjEODateReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TProjEODateReader.Destroy;
begin
  FList.Free;
end;

procedure TProjEODateReader.Log(const str: string);
begin
  FList.Clear;
end;

procedure TProjEODateReader.Clear;
begin

end;
   
procedure TProjEODateReader.Open;
const
  CIProj = 1; //  项目归属
  CIDate = 4; // 	FCST Date

var
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  irow: Integer;
  sproj: string;
  dt: TDateTime;
  sdate: string;
begin
  Clear;


  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        Log(sSheet);

        irow := 2;
        sproj := ExcelApp.Cells[irow, CIProj].Value;
        while sproj <> '' do
        begin
          dt := ExcelApp.Cells[irow, CIDate].Value;
          sdate := FormatDateTime('yyyy-MM-dd', dt);
          FList.Add(sproj + '=' + sdate);

          irow := irow + 1;     
          sproj := ExcelApp.Cells[irow, CIProj].Value;
        end;
        Break;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end;  
end;

{ TfrmProjOfICItem }

class procedure TfrmProjOfICItem.ShowForm;
var
  frmProjOfICItem: TfrmProjOfICItem;
begin
  frmProjOfICItem := TfrmProjOfICItem.Create(nil);
  try
    frmProjOfICItem.ShowModal;
  finally
    frmProjOfICItem.Free;
  end;
end;
       
procedure TfrmProjOfICItem.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmProjOfICItem.btnProjOfICItemClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leProjOfICItem.Text := sfile;
end;

procedure TfrmProjOfICItem.btnProjEODateClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leProjEODate.Text := sfile;
end;

function StringListSortCompare(List: TStringList; Index1, Index2: Integer): Integer;
var
  dt1, dt2: TDateTime;
begin
  dt1 := myStrToDateTime(List.ValueFromIndex[Index1]);    
  dt2 := myStrToDateTime(List.ValueFromIndex[Index2]);
  Result := Round(dt1 - dt2);
end;

procedure TfrmProjOfICItem.btnSave2Click(Sender: TObject);
var
  sfile: string;
  aProjOfICItemReader: TProjOfICItemReader;
  aProjEODateReader: TProjEODateReader;
  inumber: Integer;
  iproj: Integer;
  swhereuse: string;
  sproj: string;  
  sLastProj: string;
       
  ExcelApp, WorkBook: Variant;
  irow: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  Memo1.Lines.Add(sfile);

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
   
  try

    WorkBook := ExcelApp.WorkBooks.Add;
                 
    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '物料编码';
    ExcelApp.Cells[irow, 2].Value := '归属项目';
  

    aProjOfICItemReader := TProjOfICItemReader.Create(leProjOfICItem.Text);
    aProjEODateReader := TProjEODateReader.Create(leProjEODate.Text);
    try
//      aProjEODateReader.FList.CustomSort( StringListSortCompare );
      Memo1.Lines.Add(aProjEODateReader.FList.Text);

      for inumber := 0 to aProjOfICItemReader.FList.Count - 1 do
      begin
        swhereuse := aProjOfICItemReader.FList.ValueFromIndex[inumber];

        sLastProj := '';
      
        for iproj := 0 to aProjEODateReader.FList.Count - 1 do
        begin
          sproj := aProjEODateReader.FList.Names[iproj];
          if Pos(sproj, swhereuse) > 0 then
          begin
            sLastProj :=sproj;
            Break;
          end;
        end;

        irow := irow + 1;
        ExcelApp.Cells[irow, 1].Value := aProjOfICItemReader.FList.Names[inumber];
        ExcelApp.Cells[irow, 2].Value := sLastProj;
      end;


    finally
      aProjOfICItemReader.Free;
      aProjEODateReader.Free;
    end;

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;

  MessageBox(Self.Handle, '完成', '提示', 0);
 
end;

end.
