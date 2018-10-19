unit SWaterfall;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, IniFiles, ComObj,
  SWaterfallReader, Spin;

type
  TfrmSWaterfall = class(TForm)
    leWaterFall: TLabeledEdit;
    leSOPSum: TLabeledEdit;
    btnWaterFall: TButton;
    btnSOPSum: TButton;
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    tbSave: TToolButton;
    Memo1: TMemo;
    SpinEdit1: TSpinEdit;
    SpinEdit2: TSpinEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure btnWaterFallClick(Sender: TObject);
    procedure btnSOPSumClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

uses CommUtils, SOPReaderUnit, ProjYearWin;

{$R *.dfm}

class procedure TfrmSWaterfall.ShowForm;
var
  frmSWaterfall: TfrmSWaterfall;
begin
  frmSWaterfall := TfrmSWaterfall.Create(nil);
  try
    frmSWaterfall.ShowModal;
  finally
    frmSWaterfall.Free;
  end;
end;

procedure TfrmSWaterfall.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leWaterFall.Text := ini.ReadString(self.ClassName, leWaterFall.Name, '');
    leSOPSum.Text := ini.ReadString(self.ClassName, leSOPSum.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSWaterfall.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leWaterFall.Name, leWaterFall.Text);
    ini.WriteString(self.ClassName, leSOPSum.Name, leSOPSum.Text);
  finally
    ini.Free;
  end;
end;
    
procedure TfrmSWaterfall.btnWaterFallClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leWaterFall.Text := sfile;
end;

procedure TfrmSWaterfall.btnSOPSumClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOPSum.Text := sfile;
end;

function GetDateSum(aSOPProj: TSOPProj; iDate: Integer): Double;
var
  iLine: Integer;
  aSOPLine: TSOPLine;
begin
  Result := 0;
  for iLine := 0 to aSOPProj.LineCount - 1 do
  begin
    aSOPLine := aSOPProj.Lines[iLine];
    Result := Result + aSOPLine.Dates[iDate].iQty;
  end;
end;

function IndexOfWeek(slDate: TStringList; dt: TDateTime): Integer;
var
  i: Integer;
  sdate: string;
begin
  Result := -1;
  for i := 0 to slDate.Count - 1 do
  begin
    sdate := slDate.ValueFromIndex[i];
    if sdate = FormatDateTime('yyyy-MM-dd', dt) then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function ExtractSOPDate(const sfile: string): TDateTime;
var
  s: string;
  idx: Integer;
begin
  Result := 0;
  
  s := sfile;

  idx := Pos('(', s);
  if idx < 0 then Exit;
  s := Copy(s, idx + 1, Length(s));

  idx := Pos(')', s);
  if idx < 0 then Exit;     
  s := Copy(s, 1, idx - 1);
     
  idx := Pos(' ', s);
  if idx < 0 then Exit;     
  s := Copy(s, 1, idx - 1);

  Result := myStrToDateTime(s);
end;

function ExtractSOPName(const sfile: string): string;
var
  s: string;
  idx: Integer;
begin
  Result := '';
  
  s := sfile;

  idx := Pos('(', s);
  if idx < 0 then Exit;
  s := Copy(s, idx + 1, Length(s));

  idx := Pos(')', s);
  if idx < 0 then Exit;     
  s := Copy(s, 1, idx - 1);

  Result := s;
end;

procedure TfrmSWaterfall.tbSaveClick(Sender: TObject);
var                          
  ExcelApp, WorkBook: Variant;
  aSOPReader: TSOPReader;
  slProjYear: TStringList;
  iProj: Integer;
  aSOPProj: TSOPProj;
  aSOPLine: TSOPLine;
  aSOPCol: TSOPCol;
  iLine: Integer;
  iDate: Integer;
  dQty: Double;
  sfile: string;
  irow: Integer;

  aSWaterfallReader: TSWaterfallReader;
  aSWFProj: TSWFProj;
  aSWFLine: TSWFLine;
  aSWFCol: TSWFCol;

  slDate: TStringList;
  sdate: string;

  idx: Integer;
  ilt: Integer;
  iPayCol: Integer;

  dtSOP: TDateTime;
  sSOPName: string;
  iMaxRow: Integer;
  dMaxRow: Double;

  iCompareRow: Integer;
  iRow0: Integer;
  iRowSOP: Integer;
  iRowSOPPayCol: Integer;
  iMaxRowPayCol: Integer;  
  iMaxRowData: Integer;    
  iMaxRowDataPayCol: Integer;

  iSheet: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  ilt := (SpinEdit1.Value + SpinEdit2.Value) * 7;

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

  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;
  iSheet := 1;

  //////////////// summary sheet //////////////////////////////////////////////////////

  ExcelApp.Sheets[iSheet].Activate;
  ExcelApp.Sheets[iSheet].Name := 'Summary';

  ExcelApp.Cells[2, 1].Value := '项目';
  ExcelApp.Cells[2, 2].Value := '采购LT';
  ExcelApp.Cells[2, 3].Value := '制造LT';
  ExcelApp.Cells[2, 4].Value := '责任总量变化量';      

  ExcelApp.Columns[4].ColumnWidth := 16;


  slDate := TStringList.Create;

  try
                                               
    aSWaterfallReader := TSWaterfallReader.Create(leWaterFall.Text);

    slProjYear := TfrmProjYear.GetProjYears;
    aSOPReader := TSOPReader.Create(slProjYear, leSOPSum.Text);

    dtSOP := ExtractSOPDate(leSOPSum.Text);
    sSOPName := ExtractSOPName(leSOPSum.Text);

    try
      for iProj := 0 to aSOPReader.ProjCount - 1 do
      begin
        aSOPProj := aSOPReader.Projs[iProj];
        aSWFProj := aSWaterfallReader.GetProj(aSOPProj.FName);
                                
        ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[iSheet]); 
        iSheet := iSheet + 1;

        ExcelApp.Sheets[iSheet].Activate;
        ExcelApp.Sheets[iSheet].Name := aSOPProj.FName;

        irow := 3;
        ExcelApp.Cells[irow, 2].Value := 'MPS版本';
        ExcelApp.Cells[irow, 3].Value := '需求日期';
        ExcelApp.Cells[irow, 4].Value := '责任总量';

        ExcelApp.Columns[1].ColumnWidth := 2;
        ExcelApp.Columns[2].ColumnWidth := 18;
        ExcelApp.Columns[3].ColumnWidth := 10;
        ExcelApp.Columns[4].ColumnWidth := 9;

        // 统计有多少个日期列   /////////////////////////////
        if aSWFProj <> nil then
        begin
          for iLine := 0 to aSWFProj.LineCount - 1 do
          begin
            aSWFLine := aSWFProj.Lines[iLine];
            for iDate := 0 to aSWFLine.DateCount - 1 do
            begin
              aSWFCol := aSWFLine.Dates[iDate];
              sdate := FormatDateTime('yyyy-MM-dd', aSWFCol.FDate);
              slDate.Add(aSWFCol.FWeek + '=' + sdate);
            end;
            Break;
          end;
        end;


        for iLine := 0 to aSOPProj.LineCount - 1 do
        begin
          aSOPLine := aSOPProj.Lines[iLine]; 
          for iDate := 0 to aSOPLine.DateCount - 1 do
          begin
            aSOPCol := aSOPLine.Dates[iDate];

            // 对准日期列写
            idx := IndexOfWeek(slDate, aSOPCol.dt1);
            if idx < 0 then
            begin              
              sdate := FormatDateTime('yyyy-MM-dd', aSOPCol.dt1);
              slDate.Add(aSOPCol.sWeek + '=' + sdate);
            end; 
          end; 
        end;     


        irow := 3;
        for iDate := 0 to slDate.Count - 1 do
        begin                                                                 
          ExcelApp.Cells[irow - 1, iDate + 5].Value := slDate.Names[iDate];
          ExcelApp.Cells[irow, iDate + 5].Value := slDate.ValueFromIndex[iDate];      
          ExcelApp.Columns[iDate + 5].ColumnWidth := 10;

        end;

        AddColor(ExcelApp, irow - 1, 5, irow - 1, slDate.Count + 5 - 1, $DBDCF2 );
        AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $DBDCF2 );


        ExcelApp.Cells[irow, slDate.Count + 5].Value := '合计';
        

        // 开始写数据

        iMaxRow := 0;
        iMaxRowPayCol := 0;
        dMaxRow := 0;
        iRowSOP := 0;
        iRowSOPPayCol := 0;

        irow := 4;

        if aSWFProj <> nil then
        begin
          for iLine := 0 to aSWFProj.LineCount - 1 do
          begin
            aSWFLine := aSWFProj.Lines[iLine];
            ExcelApp.Cells[irow, 2].Value := aSWFLine.FName;
            ExcelApp.Cells[irow, 3].Value := aSWFLine.FDate;

            iPayCol := 0;
            for iDate := 0 to aSWFLine.DateCount - 1 do
            begin
              aSWFCol := aSWFLine.Dates[iDate];
              ExcelApp.Cells[irow, idate + 5].Value := aSWFCol.FQty;

              if aSWFCol.FDate <= aSWFLine.FDate + ilt then
              begin
                iPayCol := iDate;
              end; 
            end;

            AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $F2F2F2);

            if iPayCol > 0 then
            begin
              ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iPayCol + 5) + IntToStr(irow) + ')';
              AddColor(ExcelApp, irow, 5, irow, iPayCol + 5, $FFFF);
            end;

            aSWFLine.FRow := irow;
            aSWFLine.FPayCol := iPayCol;

            dQty := ExcelApp.Cells[irow, 4].Value;
            if dMaxRow < dQty then
            begin
              dMaxRow := dQty;
              iMaxRow := irow;   
              iMaxRowPayCol := iPayCol;
            end;

            ExcelApp.Cells[irow, slDate.Count + 5].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(slDate.Count + 5 - 1) + IntToStr(irow) + ')';
          
            irow := irow + 1;
          end;
        end;

        // 写SOP /////////////////////////
        for iLine := 0 to aSOPProj.LineCount - 1 do
        begin
          aSOPLine := aSOPProj.Lines[iLine];
                
          ExcelApp.Cells[irow, 2].Value := sSOPName;
          ExcelApp.Cells[irow, 3].Value := dtSOP;

          for iDate := 0 to aSOPLine.DateCount - 1 do
          begin
            aSOPCol := aSOPLine.Dates[iDate];

            // 对准日期列写
            idx := IndexOfWeek(slDate, aSOPCol.dt1); 
            dQty := GetDateSum(aSOPProj, iDate);
            ExcelApp.Cells[irow, idx + 5].Value := dQty;
          end;

          iPayCol := 0;
          for iDate := 0 to slDate.Count - 1 do
          begin
            if myStrToDateTime( slDate.ValueFromIndex[iDate] ) <= dtSOP + ilt then
            begin
              iPayCol := iDate;
            end;
          end;

          AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $F2F2F2);

          if iPayCol > 0 then
          begin
            ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iPayCol + 5) + IntToStr(irow) + ')';
            AddColor(ExcelApp, irow, 5, irow, iPayCol + 5, $FFFF);
            iRowSOPPayCol := iPayCol;
          end;


          iRowSOP := irow;

          dQty := ExcelApp.Cells[irow, 4].Value;
          if dMaxRow < dQty then
          begin 
            iMaxRow := irow;
            iMaxRowPayCol := iPayCol;
          end;
               
          ExcelApp.Cells[irow, slDate.Count + 5].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(slDate.Count + 5 - 1) + IntToStr(irow) + ')';
          
          irow := irow + 1;

          Break;  // 一个项目写一行， 所有行数量加总
        end;

        ExcelApp.Range[ ExcelApp.Cells[4, 4], ExcelApp.Cells[irow - 1, slDate.Count + 5] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

        AddBorder(ExcelApp, 2, 2, irow - 1, slDate.Count + 5);

        // 最大值设置颜色
        if iMaxRow > 0 then
        begin
          ExcelApp.Cells[iMaxRow, 4].Font.Color := $06009C;     
          ExcelApp.Cells[iMaxRow, 4].Interior.Color := $CEC7FF;
        end;

        // 下面是比较 ///////////////////////////////////////////////////////////

        iMaxRowData := iMaxRow;
        iMaxRowDataPayCol := iMaxRowPayCol;

        irow := irow + 1;
        iCompareRow := irow;

        irow := irow + 1;

        for iDate := 0 to slDate.Count - 1 do
        begin                                                                 
          ExcelApp.Cells[irow - 1, iDate + 5].Value := slDate.Names[iDate];
          ExcelApp.Cells[irow, iDate + 5].Value := slDate.ValueFromIndex[iDate];     
        end;

        AddColor(ExcelApp, irow - 1, 5, irow - 1, slDate.Count + 5 - 1, $DBDCF2 );
        AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $DBDCF2 );
 
        ExcelApp.Cells[irow, slDate.Count + 5].Value := '合计';
        ExcelApp.Cells[irow, 2].Value := 'MPS版本';
        ExcelApp.Cells[irow, 3].Value := '需求日期';
        ExcelApp.Cells[irow, 4].Value := '责任总量变化';

        iRow0 := 0;
        irow := irow + 1;
       
        
        dMaxRow := 0;
        iMaxRow := 0;

        if aSWFProj <> nil then
        begin
          for iLine := 0 to aSWFProj.LineCount - 1 do
          begin
            aSWFLine := aSWFProj.Lines[iLine];
            ExcelApp.Cells[irow, 2].Value := aSWFLine.FName;
            ExcelApp.Cells[irow, 3].Value := aSWFLine.FDate;
            ExcelApp.Cells[irow, slDate.Count + 5].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(slDate.Count + 5 - 1) + IntToStr(irow) + ')';
          
            for iDate := 0 to aSWFLine.DateCount - 1 do
            begin
              if iRow0 = 0 then
              begin
                ExcelApp.Cells[irow, idate + 5].Value := 0;
              end
              else
              begin
                ExcelApp.Cells[irow, idate + 5].Value := '=' + GetRef(iDate + 5) + IntToStr(aSWFLine.FRow) + '-' + GetRef(iDate + 5) + IntToStr(iRow0);
              end;
            end;  

            AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $F2F2F2);

            if aSWFLine.FPayCol > 0 then
            begin
              ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(aSWFLine.FPayCol + 5) + IntToStr(irow) + ')';
              AddColor(ExcelApp, irow, 5, irow, aSWFLine.FPayCol + 5, $FFFF);
            end;

            dQty := ExcelApp.Cells[irow, 4].Value;
            if dMaxRow < dQty then
            begin
              dMaxRow := dQty;
              iMaxRow := irow;     
              iMaxRowPayCol := aSWFLine.FPayCol;
            end;

            if iMaxRowData = aSWFLine.FRow then
            begin
              iMaxRowData := irow;
              ExcelApp.Cells[irow, 4].Interior.Color := $FF;
            end;

            iRow0 := aSWFLine.FRow;
            irow := irow + 1;
          end;
        end;


        ExcelApp.Cells[irow, 2].Value := sSOPName;
        ExcelApp.Cells[irow, 3].Value := dtSOP;    
        ExcelApp.Cells[irow, slDate.Count + 5].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(slDate.Count + 5 - 1) + IntToStr(irow) + ')';

        for iDate := 0 to slDate.Count - 1 do
        begin
          if irow0 = 0 then
          begin
            ExcelApp.Cells[irow, iDate + 5].Value := 0;
          end
          else
          begin
            ExcelApp.Cells[irow, iDate + 5].Value := '=' + GetRef(iDate + 5) + IntToStr(iRowSOP) + '-' + GetRef(iDate + 5) + IntToStr(iRow0);
          end;
        end;
                  
        AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $F2F2F2);

        if iRowSOPPayCol > 0 then
        begin
          ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef( iRowSOPPayCol  + 5) + IntToStr(irow) + ')';
          AddColor(ExcelApp, irow, 5, irow, iRowSOPPayCol + 5, $FFFF);
        end;

        dQty := ExcelApp.Cells[irow, 4].Value;
        if dMaxRow < dQty then
        begin
          iMaxRow := irow;     
          iMaxRowPayCol := iRowSOPPayCol;
        end;
            
        iPayCol := iRowSOPPayCol;
               
        if iMaxRowData = iRowSOP then
        begin
          iMaxRowData := irow;       
          ExcelApp.Cells[irow, 4].Interior.Color := $FF;
        end;
 
        irow := irow + 1;
      

        if iMaxRow > 0 then
        begin
          ExcelApp.Cells[iMaxRow, 4].Font.Color := $06009C;     
          ExcelApp.Cells[iMaxRow, 4].Interior.Color := $CEC7FF;
        end;
                     
        ///  最后汇总行处理   /////////////////////////////////////////////////////

        ExcelApp.Cells[irow, 3].Value := '责任总量变化量';
        ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iPayCol + 5) + IntToStr(irow) + ')';
        for iDate := 0 to slDate.Count - 1 do
        begin
          if iDate > iPayCol then Break;

          if iDate <= iMaxRowPayCol then
          begin
            if iMaxRowData + 1 < irow then
            begin
              ExcelApp.Cells[irow, iDate + 5].Value := '=SUM(' + GetRef(iDate + 5) + IntToStr(iMaxRowData + 1) + ':' + GetRef(iDate + 5) + IntToStr(irow - 1) + ')';
            end;
          end
          else
          begin
            ExcelApp.Cells[irow, iDate + 5].Value := '=' + GetRef(iDate + 5) + IntToStr(iRowSOP);
          end;  
        end;

        AddColor(ExcelApp, irow, 2, irow, slDate.Count + 5, $ECDFE4);  
        AddColor(ExcelApp, irow, 5, irow, iPayCol + 5, $50D092);
        AddColor(ExcelApp, irow, 5, irow, iMaxRowPayCol + 5, $00C0FF);


        ExcelApp.Range[ ExcelApp.Cells[iCompareRow + 1, 4], ExcelApp.Cells[irow, slDate.Count + 5] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

        AddBorder(ExcelApp, iCompareRow, 2, irow, slDate.Count + 5);

        // 回写 summary //////////////////////////////////////
        ExcelApp.Sheets[1].Activate;
        ExcelApp.Cells[iProj + 3, 1].Value := aSOPProj.FName;
        ExcelApp.Cells[iProj + 3, 2].Value := 6;
        ExcelApp.Cells[iProj + 3, 3].Value := 2; 
        ExcelApp.Cells[iProj + 3, 4].Value := '=''' + aSOPProj.FName + '''!D' + IntToStr(irow);

      end;  // loop of proj

      AddBorder(ExcelApp, 2, 1, aSOPReader.ProjCount + 2, 4);
      
      ExcelApp.Range[ ExcelApp.Cells[3, 4], ExcelApp.Cells[aSWaterfallReader.ProjCount + 2, 4] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

    finally
      slProjYear.Free;
      aSOPReader.Free;
       
      aSWaterfallReader.Free;
    end;

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;

    slDate.Free;
  end; 
  MessageBox(Handle, '完成', '提示', 0);
end;

end.
