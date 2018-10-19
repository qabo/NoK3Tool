unit SOPVerCompareWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, CommUtils, StdCtrls, ExtCtrls, ComObj,
  SOPReaderUnit, ProjYearWin, DateUtils, IniFiles, ExcelConsts;

type
  TfrmSOPVerCompare = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    tbClose: TToolButton;
    ToolButton1: TToolButton;
    tbSave: TToolButton;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ComboBox1: TComboBox;
    Memo1: TMemo;
    procedure tbCloseClick(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmSOPVerCompare.ShowForm;
var
  frmSOPVerCompare: TfrmSOPVerCompare;
begin
  frmSOPVerCompare := TfrmSOPVerCompare.Create(nil);
  try
    frmSOPVerCompare.ShowModal;
  finally
    frmSOPVerCompare.Free;
  end;
end;
          
procedure TfrmSOPVerCompare.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    LabeledEdit1.Text := ini.ReadString(self.ClassName, LabeledEdit1.Name, ''); 
    LabeledEdit2.Text := ini.ReadString(self.ClassName, LabeledEdit2.Name, '');
  finally
    ini.Free;
  end;
  ComboBox1Change(Sender);
end;

procedure TfrmSOPVerCompare.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, LabeledEdit1.Name, LabeledEdit1.Text);  
    ini.WriteString(self.ClassName, LabeledEdit2.Name, LabeledEdit2.Text);
  finally
    ini.Free;
  end;
end;
     
procedure TfrmSOPVerCompare.ComboBox1Change(Sender: TObject);
begin
  Memo1.Clear;
  Memo1.Lines.Add( StringReplace(ComboBox1.Text, '||', #13#10, [rfReplaceAll]) );
end;

procedure TfrmSOPVerCompare.tbCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSOPVerCompare.Button1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  LabeledEdit1.Text := sfile;
end;

procedure TfrmSOPVerCompare.Button2Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  LabeledEdit2.Text := sfile;
end;

procedure TfrmSOPVerCompare.ToolButton2Click(Sender: TObject);
begin
   TfrmProjYear.ShowForm;
end;

function StringListSortCompare_date(List: TStringList; Index1, Index2: Integer): Integer;
var
  aSOPCol1, aSOPCol2: TSOPCol;
begin
  aSOPCol1 := TSOPCol(List.Objects[Index1]);
  aSOPCol2 := TSOPCol(List.Objects[Index2]);
  if DoubleG( aSOPCol1.dt1 , aSOPCol2.dt1 ) then
    Result := 1
  else if DoubleE(aSOPCol1.dt1 , aSOPCol2.dt1) then
    Result := 0
  else Result := -1;
end;

procedure TfrmSOPVerCompare.tbSaveClick(Sender: TObject);
  function IndexOfDate(sldate: TStringList; dt1: TDateTime): Integer;
  var
    iCount: Integer;
    aSOPCol: TSOPCol;
  begin
    Result := -1;
    for iCount := 0 to sldate.Count - 1 do
    begin
      aSOPCol := TSOPCol(sldate.Objects[iCount]);
      if dt1 = aSOPCol.dt1 then
      begin
        Result := iCount;
        Break;
      end;
    end;
  end;
var
  ExcelApp, WorkBook: Variant;
  sfile: string;
  sop1, sop2: TSOPReader;
  slProjYear: TStringList;
  sldate1, sldate2: TStringList;
  aSOPCol: TSOPCol;
  idate: Integer;
  irow: Integer;
  icol: Integer;
  icolMax: Integer;
  dt0: TDateTime;
  iproj: Integer;
  aSOPProj1, aSOPProj2: TSOPProj;
  slVer1, slVer2: TStringList;
  slCap1, slCap2: TStringList;
  slColor1, slColor2: TStringList;
  iver: Integer;
  icap: Integer;
  icolor: Integer;
  irow1_proj: Integer;
  icol1: Integer;
  iSheet: Integer;
  iline: Integer;
  sver0: string;
  irow1_ver: Integer;
  aSOPLine: TSOPLine;
  lstMonth: TList;
  iMonth: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  slProjYear := TfrmProjYear.GetProjYears;    
  sop1 := TSOPReader.Create(slProjYear, LabeledEdit1.Text);
  sop2 := TSOPReader.Create(slProjYear, LabeledEdit2.Text);
  lstMonth := TList.Create;
  try



    // 开始保存 Excel
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := True;
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
      ExcelApp.Sheets[1].Delete;
    end;

              
    ExcelApp.Sheets[1].Name := 'Summary';

    sldate1 := TStringList.Create;
    sldate2 := TStringList.Create;

    sop1.GetDateList(sldate1);
    sop2.GetDateList(sldate2);


   for idate := sldate2.Count - 1 downto 0 do
   begin                                         
     aSOPCol := TSOPCol(sldate2.Objects[idate]);
     if IndexOfDate(sldate1, aSOPCol.dt1) < 0 then
     begin
       sldate1.AddObject(aSOPCol.sDate, aSOPCol);
     end
     else
     begin
       aSOPCol.Free;
     end;
     sldate2.Delete(idate);
   end;

   sldate1.CustomSort(StringListSortCompare_date);

   irow := 1;
   ExcelApp.Cells[irow, 1].Value := '版本:W' + IntToStr(YearOf(Now)) + IntToStr(WeekOf(Now));

   irow := 2;
   ExcelApp.Cells[irow, 1].Value := '机型';
   ExcelApp.Cells[irow, 2].Value := '版本\容量\颜色';
   ExcelApp.Cells[irow, 3].Value := '项目';

   MergeCells(ExcelApp, irow, 1, irow + 1, 1);
   MergeCells(ExcelApp, irow, 2, irow + 1, 2);
   MergeCells(ExcelApp, irow, 3, irow + 1, 3);

   ExcelApp.Columns[1].ColumnWidth := 8;
   ExcelApp.Columns[2].ColumnWidth := 25;
   ExcelApp.Columns[3].ColumnWidth := 15;

   icol := 4;
   dt0 := 0;

   for idate := 0 to sldate1.Count - 1 do
   begin
     aSOPCol := TSOPCol(sldate1.Objects[idate]);

     if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
     begin
       ExcelApp.Cells[irow, icol].Value := IntToStr(MonthOf(dt0)) + '月';
       MergeCells(ExcelApp, irow, icol, irow + 1, icol);
       ExcelApp.Columns[icol].ColumnWidth := 10;
       lstMonth.Add(Pointer(icol));
       icol := icol + 1;
     end;
     
     ExcelApp.Cells[irow, icol].Value := aSOPCol.sWeek;
     ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.sDate;      
     ExcelApp.Columns[icol].ColumnWidth := 10;
     icol := icol + 1;


     dt0 := aSOPCol.dt1;
   end;

   ExcelApp.Cells[irow, icol].Value := IntToStr(MonthOf(dt0)) + '月';    
   lstMonth.Add(Pointer(icol));
   MergeCells(ExcelApp, irow, icol, irow + 1, icol);

   icolMax := icol;

   irow := irow + 2;


   for iproj := 0 to sop1.ProjCount - 1 do
   begin
     aSOPProj1 := sop1.Projs[iproj];
     aSOPProj2 := sop2.GetProj(aSOPProj1.FName);
         
     ExcelApp.Cells[irow, 1].Value := aSOPProj1.FName;
     irow1_proj := irow;

     slVer1 := TStringList.Create;
     slCap1 := TStringList.Create;
     slColor1 := TStringList.Create;
             
     slVer2 := TStringList.Create;
     slCap2 := TStringList.Create;
     slColor2 := TStringList.Create;

     aSOPProj1.GetVerList(slVer1);   
     aSOPProj1.GetCapList(slCap1);
     aSOPProj1.GetColorList(slColor1);
     if aSOPProj2 <> nil then
     begin
       aSOPProj2.GetVerList(slVer2);   
       aSOPProj2.GetCapList(slCap2);
       aSOPProj2.GetColorList(slColor2);
     end;

     for iver := 0 to slVer2.Count - 1 do
     begin
       if slVer1.IndexOf(slVer2[iver]) < 0 then
       begin
         slVer1.Add(slVer2[iver]);
       end;
     end;
           
     for iver := 0 to slVer1.Count - 1 do
     begin
       ExcelApp.Cells[irow, 2].Value := slVer1[iver];
       MergeCells(ExcelApp, irow, 2, irow + 2, 2);
       ExcelApp.Cells[irow, 3].Value := Memo1.Lines[0];
       ExcelApp.Cells[irow + 1, 3].Value := Memo1.Lines[1];
       ExcelApp.Cells[irow + 2, 3].Value := Memo1.Lines[2];


                 
       dt0 := 0;
       icol := 4;
       icol1 := icol;

       for idate := 0 to sldate1.Count - 1 do
       begin
         aSOPCol := TSOPCol(sldate1.Objects[idate]);

         if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
         begin
           ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';  
           ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
           ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1);
           icol := icol + 1;
           icol1 := icol;
         end;

         ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumVer(slVer1[iver], aSOPCol.dt1);
         if aSOPProj2 <> nil then
         begin
           ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumVer(slVer1[iver], aSOPCol.dt1);
         end;
         ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
         icol := icol + 1;


         dt0 := aSOPCol.dt1;
       end;

       ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';
       ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
       ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 

       ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
       ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;

       irow := irow + 3;
     end;
        
     for icap := 0 to slCap2.Count - 1 do
     begin
       if slCap1.IndexOf(slCap2[icap]) < 0 then
       begin
         slCap1.Add(slCap2[icap]);
       end;
     end;
           
     for icap := 0 to slCap1.Count - 1 do
     begin
       ExcelApp.Cells[irow, 2].Value := slCap1[icap];
       MergeCells(ExcelApp, irow, 2, irow + 2, 2);
       ExcelApp.Cells[irow, 3].Value := Memo1.Lines[0];
       ExcelApp.Cells[irow + 1, 3].Value := Memo1.Lines[1];
       ExcelApp.Cells[irow + 2, 3].Value := Memo1.Lines[2];


                 
       dt0 := 0;
       icol := 4;
       icol1 := icol;

       for idate := 0 to sldate1.Count - 1 do
       begin
         aSOPCol := TSOPCol(sldate1.Objects[idate]);

         if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
         begin
           ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';  
           ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
           ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1);
           icol := icol + 1;
           icol1 := icol;
         end;

         ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumCap(slCap1[icap], aSOPCol.dt1);
         if aSOPProj2 <> nil then
         begin
           ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumCap(slCap1[icap], aSOPCol.dt1);
         end;
         ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
         icol := icol + 1;


         dt0 := aSOPCol.dt1;
       end;

       ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';
       ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
       ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 

       ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
       ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;

        
       irow := irow + 3; 
     end;
         
     for icolor := 0 to slColor2.Count - 1 do
     begin
       if slColor1.IndexOf(slColor2[icolor]) < 0 then
       begin
         slColor1.Add(slColor2[icolor]);
       end;
     end;
           
     for icolor := 0 to slColor1.Count - 1 do
     begin
       ExcelApp.Cells[irow, 2].Value := slColor1[icolor];
       MergeCells(ExcelApp, irow, 2, irow + 2, 2);
       ExcelApp.Cells[irow, 3].Value := Memo1.Lines[0];
       ExcelApp.Cells[irow + 1, 3].Value := Memo1.Lines[1];
       ExcelApp.Cells[irow + 2, 3].Value := Memo1.Lines[2];


                 
       dt0 := 0;
       icol := 4;
       icol1 := icol;

       for idate := 0 to sldate1.Count - 1 do
       begin
         aSOPCol := TSOPCol(sldate1.Objects[idate]);

         if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
         begin
           ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';  
           ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
           ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1);
           icol := icol + 1;
           icol1 := icol;
         end;

         ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumColor(slColor1[icolor], aSOPCol.dt1);
         if aSOPProj2 <> nil then
         begin
           ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumColor(slColor1[icolor], aSOPCol.dt1);
         end;
         ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
         icol := icol + 1;


         dt0 := aSOPCol.dt1;
       end;

       ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';
       ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
       ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 

       ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
       ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;


               
       irow := irow + 3; 
     end;

     ExcelApp.Cells[irow, 2].Value := 'TOTAL';
     MergeCells(ExcelApp, irow, 2, irow + 2, 2);
     ExcelApp.Cells[irow, 3].Value := Memo1.Lines[0];
     ExcelApp.Cells[irow + 1, 3].Value := Memo1.Lines[1];
     ExcelApp.Cells[irow + 2, 3].Value := Memo1.Lines[2];
 
     ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
     ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;

                 
     dt0 := 0;
     icol := 4;
     icol1 := icol;

     for idate := 0 to sldate1.Count - 1 do
     begin
       aSOPCol := TSOPCol(sldate1.Objects[idate]);

       if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
       begin
         ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';  
         ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
         ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1);
         icol := icol + 1;
         icol1 := icol;
       end;

       ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumAll(aSOPCol.dt1);
       if aSOPProj2 <> nil then
       begin
         ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumAll(aSOPCol.dt1);
       end;
       ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
       icol := icol + 1;


       dt0 := aSOPCol.dt1;
     end;

     ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + '' + IntToStr(irow) + ')';
     ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + '' + IntToStr(irow + 1) + ')';
     ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 

 

     irow := irow + 3;

     MergeCells(ExcelApp, irow1_proj, 1, irow - 1, 1);

     slVer1.Free;
     slCap1.Free;
     slColor1.Free;

     slVer2.Free;
     slCap2.Free;
     slColor2.Free;
   end;



    AddBorder(ExcelApp, 2, 1, irow - 1, icolMax);

    for iMonth := 0 to lstMonth.Count - 1 do
    begin
      AddColor(ExcelApp, 2, Integer(lstMonth[iMonth]), irow - 1, Integer(lstMonth[iMonth]), $00FFFF);
      ExcelApp.Range[ExcelApp.Cells[1, Integer(lstMonth[iMonth])], ExcelApp.Cells[irow - 1, Integer(lstMonth[iMonth])]].Font.Bold := True;
    end;
         
    ExcelApp.Range[ ExcelApp.Cells[4, 4], ExcelApp.Cells[irow - 1, icolMax  ] ].NumberFormatLocal := '0_ ';


    sldate1.Free;
    sldate2.Free;


    ExcelApp.Range[ ExcelApp.Cells[4, 4], ExcelApp.Cells[4, 4] ].Select;
    ExcelApp.ActiveWindow.FreezePanes := True;

   ////////////////////////////////////////////////////////////////////////////
   ////////////////////////////////////////////////////////////////////////////

    iSheet := 1;
    
    for iproj := 0 to sop1.ProjCount - 1 do
    begin
      aSOPProj1 := sop1.Projs[iproj];
      aSOPProj2 := sop2.GetProj(aSOPProj1.FName);
           
      lstMonth.Clear;

      ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[iSheet]);
      iSheet := iSheet + 1;
      
      ExcelApp.Sheets[iSheet].Name := aSOPProj1.FName;
      ExcelApp.Sheets[iSheet].Activate;

      sldate1 := TStringList.Create;
      sldate2 := TStringList.Create;
      
      aSOPProj1.GetDateList(sldate1);
      if aSOPProj2 <> nil then
      begin
        aSOPProj2.GetDateList(sldate2);
      end;
 

      for idate := sldate2.Count - 1 downto 0 do
      begin                                         
        aSOPCol := TSOPCol(sldate2.Objects[idate]);
        if IndexOfDate(sldate1, aSOPCol.dt1) < 0 then
        begin
          sldate1.AddObject(aSOPCol.sDate, aSOPCol);
        end
        else
        begin
          aSOPCol.Free;
        end;
        sldate2.Delete(idate);
      end;

      sldate1.CustomSort(StringListSortCompare_date);

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '制式';
      ExcelApp.Cells[irow, 2].Value:= '物料编码';
      ExcelApp.Cells[irow, 3].Value := '颜色';
      ExcelApp.Cells[irow, 4].Value := '容量';
      ExcelApp.Cells[irow, 5].Value := '项目';

      MergeCells(ExcelApp, irow, 1, irow + 1, 1);
      MergeCells(ExcelApp, irow, 2, irow + 1, 2);
      MergeCells(ExcelApp, irow, 3, irow + 1, 3);
      MergeCells(ExcelApp, irow, 4, irow + 1, 4);
      MergeCells(ExcelApp, irow, 5, irow + 1, 5);


      ExcelApp.Columns[1].ColumnWidth := 12;
      ExcelApp.Columns[2].ColumnWidth := 16;
      ExcelApp.Columns[3].ColumnWidth := 7;
      ExcelApp.Columns[4].ColumnWidth := 6;
      ExcelApp.Columns[5].ColumnWidth := 15;

      icol := 6;
      dt0 := 0;

      for idate := 0 to sldate1.Count - 1 do
      begin
        aSOPCol := TSOPCol(sldate1.Objects[idate]);

        if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
        begin
          ExcelApp.Cells[irow, icol].Value := IntToStr(MonthOf(dt0)) + '月';
          MergeCells(ExcelApp, irow, icol, irow + 1, icol);
          ExcelApp.Columns[icol].ColumnWidth := 10;
          lstMonth.Add(Pointer(icol));
          icol := icol + 1;
        end;
     
        ExcelApp.Cells[irow, icol].Value := aSOPCol.sWeek;
        ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.sDate;
        ExcelApp.Columns[icol].ColumnWidth := 10;
        icol := icol + 1;


        dt0 := aSOPCol.dt1;
      end;

      ExcelApp.Cells[irow, icol].Value := IntToStr(MonthOf(dt0)) + '月';
      MergeCells(ExcelApp, irow, icol, irow + 1, icol);       
      lstMonth.Add(Pointer(icol));

      icolMax := icol;

      irow := irow + 2;



      sver0 := '';
      irow1_ver := irow;
      
      for iline := 0 to aSOPProj1.LineCount - 1 do
      begin
        aSOPLine := aSOPProj1.Lines[iline];
        if sver0 <> aSOPLine.sVer then
        begin
          ExcelApp.Cells[irow, 1].Value := aSOPLine.sVer;
          if sver0 <> '' then
          begin
            MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1);
            irow1_ver := irow;
          end;
        end;
         

        ExcelApp.Cells[irow, 2].Value := aSOPLine.sNumber;
        ExcelApp.Cells[irow, 3].Value := aSOPLine.sColor;
        ExcelApp.Cells[irow, 4].Value := aSOPLine.sCap;
        ExcelApp.Cells[irow, 5].Value := Memo1.Lines[0];  
        ExcelApp.Cells[irow + 1, 5].Value := Memo1.Lines[1];
        ExcelApp.Cells[irow + 2, 5].Value := Memo1.Lines[2];

        MergeCells(ExcelApp, irow, 2, irow + 2, 2);
        MergeCells(ExcelApp, irow, 3, irow + 2, 3);
        MergeCells(ExcelApp, irow, 4, irow + 2, 4);
 

        icol := 6;
        icol1 := icol;
        dt0 := 0;

        for idate := 0 to sldate1.Count - 1 do
        begin
          aSOPCol := TSOPCol(sldate1.Objects[idate]);

          if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
            ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
            icol := icol + 1;
            icol1 := icol;
          end;
     
          ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetNumberQty(
            aSOPLine.sNumber, aSOPLine.sVer, aSOPLine.sColor, aSOPLine.sCap, aSOPCol.dt1);
          if aSOPProj2 <> nil then
          begin
            ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetNumberQty(
              aSOPLine.sNumber, aSOPLine.sVer, aSOPLine.sColor, aSOPLine.sCap, aSOPCol.dt1);
          end;
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
          icol := icol + 1;


          dt0 := aSOPCol.dt1;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 

        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;


        sver0 := aSOPLine.sVer;

        irow := irow + 3;
      end;

      if sver0 <> '' then
      begin
        MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1); 
      end;

      ExcelApp.Cells[irow, 1].Value := aSOPProj1.FName + #13#10 + 'TOTAL';

      slVer1 := TStringList.Create;
      slCap1 := TStringList.Create;
      slColor1 := TStringList.Create;

      slVer2 := TStringList.Create;
      slCap2 := TStringList.Create;
      slColor2 := TStringList.Create;
                                    
      aSOPProj1.GetVerList(slVer1); 
      aSOPProj1.GetVerList(slCap1);
      aSOPProj1.GetVerList(slColor1);

      if aSOPProj2 <> nil then
      begin
        aSOPProj2.GetVerList(slVer2);
        aSOPProj2.GetVerList(slCap2);
        aSOPProj2.GetVerList(slColor2);
      end;
      

      irow1_ver := irow;
      for iver := 0 to slVer1.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slVer1[iver];
        MergeCells(ExcelApp, irow, 2, irow + 2, 4);
        ExcelApp.Cells[irow, 5].Value := Memo1.Lines[0]; 
        ExcelApp.Cells[irow + 1, 5].Value := Memo1.Lines[1];
        ExcelApp.Cells[irow + 2, 5].Value := Memo1.Lines[2];



        icol := 6;
        icol1 := icol;
        dt0 := 0;

        for idate := 0 to sldate1.Count - 1 do
        begin
          aSOPCol := TSOPCol(sldate1.Objects[idate]);

          if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
            ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
            icol := icol + 1;
            icol1 := icol;
          end;
     
          ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumVer(slVer1[iver], aSOPCol.dt1);
          if aSOPProj2 <> nil then
          begin
            ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumVer(slVer1[iver], aSOPCol.dt1);
          end;
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
          icol := icol + 1;


          dt0 := aSOPCol.dt1;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
             
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;


        irow := irow + 3;
      end;




 
      for iver := 0 to slCap1.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slCap1[iver];
        MergeCells(ExcelApp, irow, 2, irow + 2, 4);
        ExcelApp.Cells[irow, 5].Value := Memo1.Lines[0]; 
        ExcelApp.Cells[irow + 1, 5].Value := Memo1.Lines[1];
        ExcelApp.Cells[irow + 2, 5].Value := Memo1.Lines[2];



        icol := 6;
        icol1 := icol;
        dt0 := 0;

        for idate := 0 to sldate1.Count - 1 do
        begin
          aSOPCol := TSOPCol(sldate1.Objects[idate]);

          if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
            ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
            icol := icol + 1;
            icol1 := icol;
          end;

          ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumCap(slCap1[iver], aSOPCol.dt1);
          if aSOPProj2 <> nil then
          begin
            ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumCap(slCap1[iver], aSOPCol.dt1);
          end;
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
          icol := icol + 1;


          dt0 := aSOPCol.dt1;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
               
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;


        irow := irow + 3;
      end;

               

      for iver := 0 to slColor1.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slColor1[iver];
        MergeCells(ExcelApp, irow, 2, irow + 2, 4);
        ExcelApp.Cells[irow, 5].Value := Memo1.Lines[0]; 
        ExcelApp.Cells[irow + 1, 5].Value := Memo1.Lines[1];
        ExcelApp.Cells[irow + 2, 5].Value := Memo1.Lines[2];



        icol := 6;
        icol1 := icol;
        dt0 := 0;

        for idate := 0 to sldate1.Count - 1 do
        begin
          aSOPCol := TSOPCol(sldate1.Objects[idate]);

          if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
            ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
            icol := icol + 1;
            icol1 := icol;
          end;
     
          ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumColor(slColor1[iver], aSOPCol.dt1);
          if aSOPProj2 <> nil then
          begin
            ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumColor(slColor1[iver], aSOPCol.dt1);
          end;
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
          icol := icol + 1;


          dt0 := aSOPCol.dt1;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
              
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;


        irow := irow + 3;
      end;


      ExcelApp.Cells[irow, 2].Value := 'TOTAL';
      MergeCells(ExcelApp, irow, 2, irow + 2, 4);
      ExcelApp.Cells[irow, 5].Value := Memo1.Lines[0]; 
      ExcelApp.Cells[irow + 1, 5].Value := Memo1.Lines[1];
      ExcelApp.Cells[irow + 2, 5].Value := Memo1.Lines[2];
              
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;

      icol := 6;
      icol1 := icol;
      dt0 := 0;

      for idate := 0 to sldate1.Count - 1 do
      begin
        aSOPCol := TSOPCol(sldate1.Objects[idate]);

        if not DoubleE(dt0, 0) and ( MonthOf(dt0) <> MonthOf(aSOPCol.dt1) ) then
        begin
          ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
          ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
          ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
          icol := icol + 1;
          icol1 := icol;
        end;
     
        ExcelApp.Cells[irow, icol].Value := aSOPProj1.GetSumAll(aSOPCol.dt1);
        if aSOPProj2 <> nil then
        begin
          ExcelApp.Cells[irow + 1, icol].Value := aSOPProj2.GetSumAll(aSOPCol.dt1);
        end;
        ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow) + '-' + GetRef(icol) + IntToStr(irow + 1); 
        icol := icol + 1;


        dt0 := aSOPCol.dt1;
      end;

      ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';       
      ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
      ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')'; 
           
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icolMax] ].FormatConditions[1].Font.Color := $0000FF;


      irow := irow + 3;

      

      slVer1.Free;
      slCap1.Free;
      slColor1.Free;

      slVer2.Free;
      slCap2.Free;
      slColor2.Free;      


      MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1);
     
      sldate1.Free;
      sldate2.Free;

      AddBorder(ExcelApp, 1, 1, irow - 1, icolMax);

      for iMonth := 0 to lstMonth.Count - 1 do
      begin
        AddColor(ExcelApp, 1, Integer(lstMonth[iMonth]), irow - 1, Integer(lstMonth[iMonth]), $00FFFF);
        ExcelApp.Range[ExcelApp.Cells[1, Integer(lstMonth[iMonth])], ExcelApp.Cells[irow - 1, Integer(lstMonth[iMonth])]].Font.Bold := True;
      end;
         
      ExcelApp.Range[ ExcelApp.Cells[4, 6], ExcelApp.Cells[irow - 1, icolMax  ] ].NumberFormatLocal := '0_ ';


      ExcelApp.Range[ ExcelApp.Cells[3, 6], ExcelApp.Cells[3, 6] ].Select;
      ExcelApp.ActiveWindow.FreezePanes := True;

    end;


              
    ExcelApp.Sheets[1].Activate;
    
    try
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

 

  finally
    sop1.Free;
    sop2.Free;
    slProjYear.Free;
    lstMonth.Free;
  end;

  MessageBox(Handle, '完成', 'OK', 0);
end;

end.
