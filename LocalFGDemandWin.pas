unit LocalFGDemandWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SOPReaderUnit, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin,
  CommUtils, ProjYearWin, SelectFGDemandWin, DBGridEhGrouping,
  ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, GridsEh, DBAxisGridsEh, IniFiles,
  DBGridEh, Provider, DBClient, ComObj, ExcelConsts,
  FGDemandManageWin;

type
  TfrmLocalFGDemand = class(TForm)
    ToolBar1: TToolBar;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    btnSave: TToolButton;
    Memo1: TMemo;
    btnProjYear: TToolButton;
    ToolButton2: TToolButton;
    mmosql: TMemo;
    GroupBox1: TGroupBox;
    ToolButton10: TToolButton;
    mmoFiles: TMemo;
    btnAdd: TToolButton;
    dtpCurrentWeek: TDateTimePicker;
    Label1: TLabel;
    Button1: TButton;
    btnSaveDiff: TToolButton;
    ToolButton3: TToolButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnProjYearClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnSaveDiffClick(Sender: TObject);
  private
    { Private declarations } 
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}
 
class procedure TfrmLocalFGDemand.ShowForm;
var
  frmFGDemand: TfrmLocalFGDemand;
begin
  frmFGDemand := TfrmLocalFGDemand.Create(nil);
  try
    frmFGDemand.ShowModal;
  finally
    frmFGDemand.Free;
  end;
end;

procedure TfrmLocalFGDemand.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  sfile: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    dtpCurrentWeek.DateTime := ini.ReadDateTime(self.ClassName, dtpCurrentWeek.Name, 0);
    sfile := ini.ReadString(self.ClassName, mmoFiles.Name, '');
    sfile := Trim(sfile);
    mmoFiles.Lines.Add( StringReplace( sfile , ';', #13#10, [rfReplaceAll]) ); 
  finally
    ini.Free;
  end;
end;

procedure TfrmLocalFGDemand.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  sfile: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    sfile := StringReplace( Trim( mmoFiles.Lines.Text) , #13#10, ';', [rfReplaceAll]) ;
    ini.WriteString(self.ClassName, mmoFiles.Name, sfile);
    ini.WriteDateTime(self.ClassName, dtpCurrentWeek.Name, dtpCurrentWeek.DateTime);
  finally
    ini.Free;
  end;
end;
 
procedure AddDateToList(sldate: TStringList; aSOPCol: TSOPCol);
var
  idate: Integer;
  aSOPCol0: TSOPCol;
begin
  for idate := 0 to sldate.Count - 1 do
  begin
    aSOPCol0 := TSOPCol(sldate.Objects[idate]);
    if aSOPCol0.dt1 = aSOPCol.dt1 then Exit;

    if aSOPCol0.dt1 > aSOPCol.dt1 then
    begin
      sldate.InsertObject(idate, aSOPCol.sDate, aSOPCol);
      Exit;
    end;
  end;
  sldate.AddObject(aSOPCol.sDate, aSOPCol);
end;
    
procedure TfrmLocalFGDemand.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmLocalFGDemand.btnProjYearClick(Sender: TObject);
begin
  TfrmProjYear.ShowForm;
end;
 
procedure TfrmLocalFGDemand.btnAddClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoFiles.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;

procedure TfrmLocalFGDemand.btnSaveClick(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;
 
  
  slReaders: TStringList;
  ifile: Integer; 
  sname: string;

  aSOPReader: TSOPReader;

  sproj: string;
  slproj: TStringList;

  sldate: TStringList;
  slnumber: TStringList;

  idx: Integer;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iLine: Integer;
  aSOPLine: TSOPLine;
  aSOPLine2: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  slProjYear: TStringList;   

  sfile: string;
  irow: Integer;
  irow1: Integer;

  iweek0: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;



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
 

  slReaders := TStringList.Create;
  slproj := TStringList.Create;
  sldate := TStringList.Create;
  slnumber := TStringList.Create;

  try       
    slProjYear := TfrmProjYear.GetProjYears;
    try
      for ifile := 0 to mmoFiles.Lines.Count - 1 do
      begin
        if Trim(mmoFiles.Lines[ifile] ) = '' then Continue;
        
        sname := ChangeFileExt( ExtractFileName(mmoFiles.Lines[ifile]), '' );

        Memo1.Lines.Add('读取销售计划');
        aSOPReader := TSOPReader.Create(slProjYear, mmoFiles.Lines[ifile]); 
        slReaders.AddObject(sname, aSOPReader);

        for iproj := 0 to aSOPReader.FProjs.Count - 1 do
        begin
          sproj := aSOPReader.FProjs[iproj];
          if slproj.IndexOf(sproj) < 0 then
          begin
            slproj.Add(sproj);
          end;
        end;

      end;
    finally
      slProjYear.Free;
    end;

    while ExcelApp.Sheets.Count < slproj.Count do
    begin
      ExcelApp.Sheets.Add;
    end;

    for iproj := 0 to slproj.Count - 1 do
    begin
      sproj := slproj[iproj];

      ExcelApp.Sheets[iproj + 1].Activate; 
      ExcelApp.Sheets[iproj + 1].Name := sproj;

      for ifile := 0 to slReaders.Count - 1 do
      begin
        aSOPReader := TSOPReader(slReaders.Objects[ifile]);
        idx := aSOPReader.FProjs.IndexOf(sproj);
        if idx < 0 then Continue;

        aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[idx]);
        if aSOPProj.FList.Count = 0 then Continue;

        aSOPLine := TSOPLine( aSOPProj.FList.Objects[0] );
         
        for idate := 0 to aSOPLine.FList.Count - 1 do
        begin
          aSOPCol := TSOPCol( aSOPLine.FList.Objects[idate]);
          AddDateToList(sldate, aSOPCol);
        end;

        for iLine := 0 to aSOPProj.FList.Count - 1 do
        begin
          aSOPLine := TSOPLine( aSOPProj.FList.Objects[iLine] );
          if slnumber.IndexOf(aSOPLine.sNumber) >= 0 then Continue;
          slnumber.AddObject(aSOPLine.sNumber, aSOPLine);
        end;

      end;


      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'Week'; 
      ExcelApp.Cells[irow, 2].Value := '版本';
      ExcelApp.Cells[irow, 3].Value := '产品编码';
      ExcelApp.Cells[irow, 4].Value := '颜色';
      ExcelApp.Cells[irow, 5].Value := '容量';
                   
      ExcelApp.Columns[1].ColumnWidth := 45;
      ExcelApp.Columns[2].ColumnWidth := 12;
      ExcelApp.Columns[3].ColumnWidth := 15;
      ExcelApp.Columns[4].ColumnWidth := 7;
      ExcelApp.Columns[5].ColumnWidth := 7;

      MergeCells(ExcelApp, irow, 1, irow + 2, 1);
      MergeCells(ExcelApp, irow, 2, irow + 2, 2);
      MergeCells(ExcelApp, irow, 3, irow + 2, 3);
      MergeCells(ExcelApp, irow, 4, irow + 2, 4);
      MergeCells(ExcelApp, irow, 5, irow + 2, 5);

      for idate := 0 to sldate.Count - 1 do
      begin
        aSOPCol := TSOPCol(sldate.Objects[idate]);
        ExcelApp.Cells[irow,     idate * 3 + 6].Value := aSOPCol.sWeek;
        ExcelApp.Cells[irow + 1, idate * 3 + 6].Value := aSOPCol.sDate;  
        ExcelApp.Cells[irow + 2, idate * 3 + 6].Value := 'Qty';
        ExcelApp.Cells[irow + 2, idate * 3 + 1 + 6].Value := 'Delta';
        ExcelApp.Cells[irow + 2, idate * 3 + 2 + 6].Value := '%';

        MergeCells(ExcelApp, irow, idate * 3 + 6, irow, idate * 3 + 2 + 6);
        MergeCells(ExcelApp, irow + 1, idate * 3 + 6, irow + 1, idate * 3 + 2 + 6);
      end;


      AddColor(ExcelApp, irow, 1, irow + 2, sldate.Count * 3 + 5, $DBDCF2);
      ExcelApp.Range[ ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow + 2, sldate.Count * 3 + 5] ].HorizontalAlignment := xlCenter;

      irow := 4;
      
      for iLine := 0 to slnumber.Count - 1 do
      begin
        aSOPLine := TSOPLine(slnumber.Objects[iLine]);

        irow1 := irow;

        for ifile := 0 to slReaders.Count - 1 do
        begin
          aSOPReader := TSOPReader(slReaders.Objects[ifile]);
          idx := aSOPReader.FProjs.IndexOf(sproj);
          if idx < 0 then Continue;

          ExcelApp.Cells[irow, 1].Value := slReaders[ifile]; 
          ExcelApp.Cells[irow, 2].Value := aSOPLine.sVer;
          ExcelApp.Cells[irow, 3].Value := aSOPLine.sNumber;
          ExcelApp.Cells[irow, 4].Value := aSOPLine.sColor;
          ExcelApp.Cells[irow, 5].Value := aSOPLine.sCap;
          
          aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[idx]);
          aSOPLine2 := aSOPProj.GetLine(aSOPLine.sVer, aSOPLine.sNumber,
            aSOPLine.sColor, aSOPLine.sCap);
          if aSOPLine2 = nil then Continue;

          iweek0 := -1;
          
          for idate := 0 to sldate.Count - 1 do
          begin
            aSOPCol := aSOPLine2.GetCol( sldate[idate] );
            if aSOPCol = nil then Continue;

            if iweek0 = -1 then
            begin
              if aSOPCol.dt1 >= dtpCurrentWeek.DateTime then
                iweek0 := idate;
            end;

            ExcelApp.Cells[irow, idate * 3 + 6].Value := aSOPCol.iQty;
            if irow = irow1 then
            begin
              ExcelApp.Cells[irow, idate * 3 + 1 + 6].Value := 0;
              ExcelApp.Cells[irow, idate * 3 + 2 + 6].Value := 0;
            end
            else
            begin
              ExcelApp.Cells[irow, idate * 3 + 1 + 6].Value := '=' + GetRef( idate * 3 + 6 ) + IntToStr(irow) + '-' + GetRef( idate * 3 + 6 ) + IntToStr(irow - 1);
              ExcelApp.Cells[irow, idate * 3 + 2 + 6].Value := '=IF(' + GetRef( idate * 3 + 6 ) + IntToStr(irow - 1) + '=0,IF(' + GetRef( idate * 3 + 6 ) + IntToStr(irow) + '=0,0,1),(' + GetRef( idate * 3 + 6 ) + IntToStr(irow) + '-' + GetRef( idate * 3 + 6 ) + IntToStr(irow - 1) + ')/' + GetRef( idate * 3 + 6 ) + IntToStr(irow - 1) + ')';
              ExcelApp.Cells[irow, idate * 3 + 2 + 6].NumberFormatLocal := '0.0%';

              if (iweek0 = -1) or (idate - iweek0 <= 1) then
              begin
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlLess, '=-0.0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[1].Interior.Color := $0000FF;
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlGreater, '=0.0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[2].Interior.Color := $0000FF;
              end
              else if idate - iweek0 <= 2 then
              begin
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlLess, '=-0.1', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[1].Interior.Color := $0000FF;
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlGreater, '=0.1', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[2].Interior.Color := $0000FF;
              end
              else if idate - iweek0 <= 4 then
              begin
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlLess, '=-0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[1].Interior.Color := $0000FF;
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlGreater, '=0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[2].Interior.Color := $0000FF;
              end
              else if idate - iweek0 <= 8 then
              begin
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlLess, '=-0.3', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[1].Interior.Color := $0000FF;
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlGreater, '=0.3', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[2].Interior.Color := $0000FF;
              end
              else if idate - iweek0 <= 10 then
              begin
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlLess, '=-0.4', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[1].Interior.Color := $0000FF;
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlGreater, '=0.4', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[2].Interior.Color := $0000FF;
              end
              else if idate - iweek0 <= 13 then
              begin
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlLess, '=-0.5', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[1].Interior.Color := $0000FF;
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions.Add(xlCellValue, xlGreater, '=0.5', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                ExcelApp.Range[ExcelApp.Cells[irow, idate * 3 + 2 + 6], ExcelApp.Cells[irow, idate * 3 + 2 + 6]].FormatConditions[2].Interior.Color := $0000FF;                
              end;

            end;
          end;

          irow := irow + 1;
        end;
      end;  

      AddBorder(ExcelApp, 1, 1, irow - 1, sldate.Count * 3 + 5);

      sldate.Clear;
      slnumber.Clear;
    end;
    

  finally
    for ifile := 0 to slReaders.Count - 1 do
    begin
      aSOPReader := TSOPReader(slReaders.Objects[ifile]);
      aSOPReader.Free;
    end;
    slReaders.Free;

    slproj.Free;
    sldate.Free;
    slnumber.Free;
  end;
    

  ExcelApp.Sheets[1].Activate;
    
  try
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;

  MessageBox(Handle, '完成', '提示', 0);

end;

procedure TfrmLocalFGDemand.Button1Click(Sender: TObject);
begin
  mmoFiles.Clear;
end;

procedure TfrmLocalFGDemand.btnSaveDiffClick(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;
 
  
  slReaders: TStringList;
  ifile: Integer; 
  sname: string;

  aSOPReader: TSOPReader;

  sproj: string;
  slproj: TStringList;

  sldate: TStringList;
  slnumber: TStringList;

  idx: Integer;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iLine: Integer;
  aSOPLine: TSOPLine;
  aSOPLine2: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  aSOPCol0: TSOPCol;
  slProjYear: TStringList;   

  sfile: string;
  irow: Integer;
  irow1: Integer;
  icol: Integer;
  bNeedWrite: Boolean;
  bHeadWritten: Boolean;

  iweek0: Integer;
  dqty0, dqty: Double;
  dt: TDateTime;
begin
  if not ExcelSaveDialog(sfile) then Exit;



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
 

  slReaders := TStringList.Create;
  slproj := TStringList.Create;
  sldate := TStringList.Create;
  slnumber := TStringList.Create;

  try       
    slProjYear := TfrmProjYear.GetProjYears;
    try
      for ifile := 0 to mmoFiles.Lines.Count - 1 do
      begin
        if Trim(mmoFiles.Lines[ifile] ) = '' then Continue;
        
        sname := ChangeFileExt( ExtractFileName(mmoFiles.Lines[ifile]), '' );

        Memo1.Lines.Add('读取销售计划');
        aSOPReader := TSOPReader.Create(slProjYear, mmoFiles.Lines[ifile]); 
        slReaders.AddObject(sname, aSOPReader);

        for iproj := 0 to aSOPReader.FProjs.Count - 1 do
        begin
          sproj := aSOPReader.FProjs[iproj];
          if slproj.IndexOf(sproj) < 0 then
          begin
            slproj.Add(sproj);
          end;
        end;

      end;
    finally
      slProjYear.Free;
    end;

    while ExcelApp.Sheets.Count < slproj.Count do
    begin
      ExcelApp.Sheets.Add;
    end;

    for iproj := 0 to slproj.Count - 1 do
    begin
      sproj := slproj[iproj];

      ExcelApp.Sheets[iproj + 1].Activate; 
      ExcelApp.Sheets[iproj + 1].Name := sproj;

      for ifile := 0 to slReaders.Count - 1 do
      begin
        aSOPReader := TSOPReader(slReaders.Objects[ifile]);
        idx := aSOPReader.FProjs.IndexOf(sproj);
        if idx < 0 then Continue;

        aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[idx]);
        if aSOPProj.FList.Count = 0 then Continue;

        aSOPLine := TSOPLine( aSOPProj.FList.Objects[0] );
         
        for idate := 0 to aSOPLine.FList.Count - 1 do
        begin
          aSOPCol := TSOPCol( aSOPLine.FList.Objects[idate]);
          AddDateToList(sldate, aSOPCol);
        end;

        for iLine := 0 to aSOPProj.FList.Count - 1 do
        begin
          aSOPLine := TSOPLine( aSOPProj.FList.Objects[iLine] );
          if slnumber.IndexOf(aSOPLine.sNumber) >= 0 then Continue;
          slnumber.AddObject(aSOPLine.sNumber, aSOPLine);
        end;

      end;


      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'Week'; 
      ExcelApp.Cells[irow, 2].Value := '版本';
      ExcelApp.Cells[irow, 3].Value := '产品编码';
      ExcelApp.Cells[irow, 4].Value := '颜色';
      ExcelApp.Cells[irow, 5].Value := '容量';
                   
      ExcelApp.Columns[1].ColumnWidth := 45;
      ExcelApp.Columns[2].ColumnWidth := 12;
      ExcelApp.Columns[3].ColumnWidth := 15;
      ExcelApp.Columns[4].ColumnWidth := 7;
      ExcelApp.Columns[5].ColumnWidth := 7;

      MergeCells(ExcelApp, irow, 1, irow + 2, 1);
      MergeCells(ExcelApp, irow, 2, irow + 2, 2);
      MergeCells(ExcelApp, irow, 3, irow + 2, 3);
      MergeCells(ExcelApp, irow, 4, irow + 2, 4);
      MergeCells(ExcelApp, irow, 5, irow + 2, 5);

      irow := 4;
      
      for iLine := 0 to slnumber.Count - 1 do
      begin
        aSOPLine := TSOPLine(slnumber.Objects[iLine]);

        for ifile := 0 to slReaders.Count - 1 do
        begin
          aSOPReader := TSOPReader(slReaders.Objects[ifile]);
          idx := aSOPReader.FProjs.IndexOf(sproj);
          if idx < 0 then Continue;

          ExcelApp.Cells[irow, 1].Value := slReaders[ifile];
          ExcelApp.Cells[irow, 2].Value := aSOPLine.sVer;
          ExcelApp.Cells[irow, 3].Value := aSOPLine.sNumber;
          ExcelApp.Cells[irow, 4].Value := aSOPLine.sColor;
          ExcelApp.Cells[irow, 5].Value := aSOPLine.sCap;

          irow := irow + 1;
        end;
      end;


      // 按 列 写  /////////////////////////////////////////////////////////////
      iweek0 := -1;
      icol := 6;

      for idate := 0 to sldate.Count - 1 do
      begin          
        bNeedWrite := False;
       
        for iLine := 0 to slnumber.Count - 1 do
        begin
          aSOPLine := TSOPLine(slnumber.Objects[iLine]);
 
          for ifile := 1 to slReaders.Count - 1 do
          begin
            aSOPReader := TSOPReader(slReaders.Objects[ifile]);
            aSOPCol := nil;
            aSOPCol0 := nil;
            idx := aSOPReader.FProjs.IndexOf(sproj);
            if idx >= 0 then
            begin
              aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[idx]);
              aSOPLine2 := aSOPProj.GetLine(aSOPLine.sVer, aSOPLine.sNumber,
                aSOPLine.sColor, aSOPLine.sCap);
              if aSOPLine2 <> nil then
              begin
                aSOPCol := aSOPLine2.GetCol( sldate[idate] );
              end;
            end;

            /////////////////////////////////////////////////////////////

            aSOPReader := TSOPReader(slReaders.Objects[ifile - 1]);
            idx := aSOPReader.FProjs.IndexOf(sproj);
            if idx >= 0 then
            begin
              aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[idx]);
              aSOPLine2 := aSOPProj.GetLine(aSOPLine.sVer, aSOPLine.sNumber,
                aSOPLine.sColor, aSOPLine.sCap);
              if aSOPLine2 <> nil then
              begin
                aSOPCol0 := aSOPLine2.GetCol( sldate[idate] );
              end;
            end;

            dqty0 := 0;
            dqty := 0;
 
            if aSOPCol <> nil then
            begin
              dt := aSOPCol.dt1;
            end
            else
            begin
              if aSOPCol0 <> nil then
              begin
                dt := aSOPCol0.dt1
              end
              else Continue; // 都为nil，无法获取日期
            end;

            if aSOPCol <> nil then dqty := aSOPCol.iQty;
            if aSOPCol0 <> nil then dqty0 := aSOPCol0.iQty;

            if iweek0 = -1 then
            begin
              if dt >= dtpCurrentWeek.DateTime then
                iweek0 := idate;
            end;                   

            if (iweek0 = -1) or (idate - iweek0 <= 1) then
            begin
//                ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//                ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
//                ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//                ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;

              if dqty <> dqty0 then
              begin
                bNeedWrite := True;
                Break;
              end;
            end
            else if idate - iweek0 <= 2 then
            begin
              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.1', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.1', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;

              if dqty0 = 0 then
              begin
                if dqty > 0 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end
              else
              begin
                if Abs(dqty - dqty0) / dqty0 > 0.1 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end;            
            end
            else if idate - iweek0 <= 4 then
            begin
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;

              if dqty0 = 0 then
              begin
                if dqty > 0 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end
              else
              begin
                if Abs(dqty - dqty0) / dqty0 > 0.2 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end;
            end
            else if idate - iweek0 <= 8 then
            begin
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.3', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.3', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;

              if dqty0 = 0 then
              begin
                if dqty > 0 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end
              else
              begin
                if Abs(dqty - dqty0) / dqty0 > 0.3 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end;
            end
            else if idate - iweek0 <= 10 then
            begin
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.4', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.4', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;

              if dqty0 = 0 then
              begin
                if dqty > 0 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end
              else
              begin
                if Abs(dqty - dqty0) / dqty0 > 0.4 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end;
            end
            else if idate - iweek0 <= 13 then
            begin
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.5', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.5', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//              ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;

              if dqty0 = 0 then
              begin
                if dqty > 0 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end
              else
              begin
                if Abs(dqty - dqty0) / dqty0 > 0.5 then
                begin
                  bNeedWrite := True;
                  Break;
                end;
              end;
            end; 
          end;
        end;
                  

        if bNeedWrite then
        begin
          bHeadWritten := False;
          irow := 4;

          for iLine := 0 to slnumber.Count - 1 do
          begin
            aSOPLine := TSOPLine(slnumber.Objects[iLine]);

            irow1 := irow;

            for ifile := 0 to slReaders.Count - 1 do
            begin
              aSOPReader := TSOPReader(slReaders.Objects[ifile]);
              aSOPCol := nil;
              
              idx := aSOPReader.FProjs.IndexOf(sproj);
              if idx >= 0 then
              begin
                aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[idx]);
                aSOPLine2 := aSOPProj.GetLine(aSOPLine.sVer, aSOPLine.sNumber,
                  aSOPLine.sColor, aSOPLine.sCap);
                if aSOPLine2 <> nil then
                begin
                  aSOPCol := aSOPLine2.GetCol( sldate[idate] );
                end;
              end;

              dqty := 0;
              if aSOPCol <> nil then
              begin
                dqty := aSOPCol.iQty;

                if not bHeadWritten then
                begin
                  bHeadWritten := True;
                  ExcelApp.Cells[1,     icol].Value := aSOPCol.sWeek;
                  ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;
                  ExcelApp.Cells[3, icol].Value := 'Qty';
                  ExcelApp.Cells[3, icol + 1].Value := 'Delta';
                  ExcelApp.Cells[3, icol + 2].Value := '%';

                  MergeCells(ExcelApp, 1, icol, 1, icol + 2);
                  MergeCells(ExcelApp, 2, icol, 2, icol + 2);
                end; 
              end;


              ExcelApp.Cells[irow, icol].Value := dqty;
              if irow = irow1 then
              begin
                ExcelApp.Cells[irow, icol + 1].Value := 0;
                ExcelApp.Cells[irow, icol + 2].Value := 0;
              end
              else
              begin
                ExcelApp.Cells[irow, icol + 1].Value := '=' + GetRef( icol ) + IntToStr(irow) + '-' + GetRef( icol ) + IntToStr(irow - 1);
                ExcelApp.Cells[irow, icol + 2].Value := '=IF(' + GetRef( icol ) + IntToStr(irow - 1) + '=0,IF(' + GetRef( icol ) + IntToStr(irow) + '=0,0,1),(' + GetRef( icol ) + IntToStr(irow) + '-' + GetRef( icol ) + IntToStr(irow - 1) + ')/' + GetRef( icol ) + IntToStr(irow - 1) + ')';
                ExcelApp.Cells[irow, icol + 2].NumberFormatLocal := '0.0%';

                if (iweek0 = -1) or (idate - iweek0 <= 1) then
                begin
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;
                end
                else if idate - iweek0 <= 2 then
                begin
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.1', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.1', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;
                end
                else if idate - iweek0 <= 4 then
                begin
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;
                end
                else if idate - iweek0 <= 8 then
                begin
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.3', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.3', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;
                end
                else if idate - iweek0 <= 10 then
                begin
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.4', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.4', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;
                end
                else if idate - iweek0 <= 13 then
                begin
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlLess, '=-0.5', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[1].Interior.Color := $0000FF;
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions.Add(xlCellValue, xlGreater, '=0.5', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
                  ExcelApp.Range[ExcelApp.Cells[irow, icol + 2], ExcelApp.Cells[irow, icol + 2]].FormatConditions[2].Interior.Color := $0000FF;
                end;

              end;


              irow := irow + 1;
            end;
          end;

          icol := icol + 3;
        end;
      end;

      AddColor(ExcelApp, 1, 1, 3, icol - 1, $DBDCF2);
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[3, icol - 1] ].HorizontalAlignment := xlCenter;


      

      AddBorder(ExcelApp, 1, 1, irow - 1, icol - 1);

      sldate.Clear;
      slnumber.Clear;
    end;
    

  finally
    for ifile := 0 to slReaders.Count - 1 do
    begin
      aSOPReader := TSOPReader(slReaders.Objects[ifile]);
      aSOPReader.Free;
    end;
    slReaders.Free;

    slproj.Free;
    sldate.Free;
    slnumber.Free;
  end;
    

  ExcelApp.Sheets[1].Activate;
    
  try
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
