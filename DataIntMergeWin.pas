unit DataIntMergeWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ComCtrls, ToolWin, StdCtrls, IniFiles, CommUtils, ComObj,
  DateUtils;

type
  TWeekDate = packed record
    sweek: string;
    sdate: string;
    dt1: TDateTime;
    dt2: TDateTime;
    newIdx: Integer;
    oldCol: Integer;
  end;
  PWeekDate = ^TWeekDate;
  
  TfrmDataIntMerge = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    mmoFiles: TMemo;
    Label1: TLabel;
    tbAdd: TToolButton;
    ToolButton2: TToolButton;
    mmoYears: TMemo;
    Label2: TLabel;
    Memo1: TMemo;
    Label3: TLabel;
    mmoMode: TMemo;
    procedure btnExitClick(Sender: TObject);
    procedure tbAddClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    procedure ReadWeeks(const sfile: string; slweeks: TStringList;
      const syear: string; var rc: Integer);
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmDataIntMerge.ShowForm;
var
  frmDataIntMerge: TfrmDataIntMerge;
begin
  frmDataIntMerge := TfrmDataIntMerge.Create(nil);
  try
    frmDataIntMerge.ShowModal;
  finally
    frmDataIntMerge.Free;
  end;
end;
   
procedure TfrmDataIntMerge.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    mmoYears.Text := StringReplace( ini.ReadString(self.ClassName, mmoYears.Name, ''), '||', #13#10, [rfReplaceAll] );
    mmoFiles.Text := StringReplace( ini.ReadString(self.ClassName, mmoFiles.Name, ''), '||', #13#10, [rfReplaceAll] );
    mmoMode.Text := StringReplace( ini.ReadString(self.ClassName, mmoMode.Name, ''), '||', #13#10, [rfReplaceAll] );
  finally
    ini.Free;
  end;
end;

procedure TfrmDataIntMerge.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, mmoYears.Name, StringReplace(mmoYears.Text, #13#10, '||', [rfReplaceAll] ));
    ini.WriteString(self.ClassName, mmoFiles.Name, StringReplace(mmoFiles.Text, #13#10, '||', [rfReplaceAll] ));
    ini.WriteString(self.ClassName, mmoMode.Name, StringReplace(mmoMode.Text, #13#10, '||', [rfReplaceAll] ));
  finally
    ini.Free;
  end;
end;

procedure TfrmDataIntMerge.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmDataIntMerge.tbAddClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoFiles.Lines.Add(StringReplace(sfile, ';', #13#10, [rfReplaceAll]));
end;

procedure TfrmDataIntMerge.ReadWeeks(const sfile: string; slweeks: TStringList;
  const syear: string; var rc: Integer);
var
  ExcelApp, WorkBook: Variant;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  sweek, sdate: string;
  irow: Integer;
  icol: Integer;
  iyear: Integer;
  sdt1, sdt2: string;
  dt0: TDateTime;
  dt1, dt2: TDateTime;
  idx: Integer;
  aWeekDatePtr: PWeekDate;
begin
  rc := 0;
  slweeks.Clear;

  iyear := StrToInt(syear);

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(sfile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;

        if stitle <> 'week物料编码颜色容量制式计划' then
        begin
        
          Continue;
        end;

        dt0 := 0;

        icol := 7;
        sweek := ExcelApp.Cells[irow, icol].Value;
        sdate := ExcelApp.Cells[irow + 1, icol].Value;
        while (sweek <> '') and (sdate <> '') do
        begin
          if dt0 = 0 then  // 第一列
          begin
            idx := Pos('-', sdate);
            if idx < 0 then // 单个日期
            begin
              sdt1 := StringReplace(sdate, '/', '-', [rfReplaceAll]);
              sdt1 := IntToStr(iyear) + '-' + sdt1;
              dt1 := myStrToDateTime(sdt1);
              dt2 := dt1;
            end
            else // 日期期间
            begin
              sdt1 := Copy(sdate, 1, idx - 1);
              sdt2 := Copy(sdate, idx + 1, Length(sdate) - idx);
              
              sdt1 := StringReplace(sdt1, '/', '-', [rfReplaceAll]);
              sdt2 := StringReplace(sdt2, '/', '-', [rfReplaceAll]);

              sdt1 := IntToStr(iyear) + '-' + sdt1;
              sdt2 := IntToStr(iyear) + '-' + sdt2;
              
              dt1 := myStrToDateTime(sdt1);
              dt2 := myStrToDateTime(sdt2);
            end;
          end
          else   // 第 n + 1 列
          begin
            idx := Pos('-', sdate);
            if idx < 0 then // 单个日期
            begin
              sdt1 := StringReplace(sdate, '/', '-', [rfReplaceAll]);
              sdt1 := IntToStr(iyear) + '-' + sdt1;
              dt1 := myStrToDateTime(sdt1);

              if dt1 < dt0 then
              begin
                iyear := iyear + 1;
                dt1 := EncodeDate(iyear, MonthOf(dt1), DayOf(dt1));
              end;
              
              dt2 := dt1;
            end
            else // 日期期间
            begin
              sdt1 := Copy(sdate, 1, idx - 1);
              sdt1 := StringReplace(sdt1, '/', '-', [rfReplaceAll]);  
              sdt1 := IntToStr(iyear) + '-' + sdt1;    
              dt1 := myStrToDateTime(sdt1);
              
              if dt1 < dt0 then
              begin
                iyear := iyear + 1;
                dt1 := EncodeDate(iyear, MonthOf(dt1), DayOf(dt1));
              end;
              
              sdt2 := Copy(sdate, idx + 1, Length(sdate) - idx);
              sdt2 := StringReplace(sdt2, '/', '-', [rfReplaceAll]);
              sdt2 := IntToStr(iyear) + '-' + sdt2; 
              dt2 := myStrToDateTime(sdt2);
            end;
          end;   
          dt0 := dt1;

          aWeekDatePtr := New(PWeekDate);
          aWeekDatePtr^.sweek := sweek;
          aWeekDatePtr^.sdate := sdate;
          aWeekDatePtr^.dt1 := dt1;
          aWeekDatePtr^.dt2 := dt2;
          aWeekDatePtr^.oldCol := icol;
                                                                      
          slweeks.AddObject(sweek + '=' + sdate, TObject(aWeekDatePtr));

          icol := icol + 1;
          sweek := ExcelApp.Cells[irow, icol].Value;
          sdate := ExcelApp.Cells[irow + 1, icol].Value;
        end;

        irow := 3;
        sweek := ExcelApp.Cells[irow, 1].Value;
        while sweek <> '' do
        begin
          rc := rc + 1;                        
          irow := irow + 1;
          sweek := ExcelApp.Cells[irow, 1].Value;
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

function ExtractProjName(const sfile: string): string;
begin
  Result := ExtractFileName(sfile);
  Result := StringReplace(Result, '-', ' ', [rfReplaceAll]);
  Result := Copy(Result, 1, Pos(' ', Result) - 1);
end;

function StringListSortCompare(List: TStringList; Index1, Index2: Integer): Integer;
var
  slweeks1: TStringList;
  slweeks2: TStringList;
  aWeekDatePtr1: PWeekDate;
  aWeekDatePtr2: PWeekDate;
begin
  slweeks1 := TStringList(List.Objects[Index1]);
  slweeks2 := TStringList(List.Objects[Index2]);

  if slweeks1.Count = 0 then
  begin
//    aWeekDatePtr1 := nil;
    Result := -1;
    Exit;
  end;

  if slweeks2.Count = 0 then
  begin
//    aWeekDatePtr2 := nil;
    Result := 1;
    Exit;
  end;
  
  aWeekDatePtr1 := PWeekDate(slweeks1.Objects[0]);
  aWeekDatePtr2 := PWeekDate(slweeks2.Objects[0]);

  if aWeekDatePtr1^.dt1 > aWeekDatePtr2^.dt1 then
    Result := 1
  else if aWeekDatePtr1^.dt1 = aWeekDatePtr2^.dt1 then
    Result := 0
  else
    Result := -1; 
end;

function AddToList(slweeks_all: TStringList; aWeekDatePtr: PWeekDate): Boolean;
var
  iweek: Integer;
  p: PWeekDate;
begin
  Result := True;
  for iweek := 0 to slweeks_all.Count - 1 do
  begin
    p := PWeekDate(slweeks_all.Objects[iweek]);
    if p^.dt1 > aWeekDatePtr^.dt1 then
    begin
      slweeks_all.InsertObject(iweek, aWeekDatePtr.sdate, TObject(aWeekDatePtr));
      Exit;
    end
    else if p^.dt1 = aWeekDatePtr^.dt1 then
    begin
      if p^.dt2 <> aWeekDatePtr^.dt2 then // 日期区间不一致
      begin
        Result := False;
      end;
      Exit;
    end;
  end;
  slweeks_all.AddObject(aWeekDatePtr.sdate, TObject(aWeekDatePtr));
end;

procedure IndexOfWeek(slweeks_all: TStringList; aWeekDatePtr: PWeekDate);
var
  i: Integer;
  p: PWeekDate;
begin
  aWeekDatePtr^.newIdx := -1;
  for i := 0 to slweeks_all.Count - 1 do
  begin
    p := PWeekDate(slweeks_all.Objects[i]);
    if p^.dt1 = aWeekDatePtr^.dt1 then
    begin
      aWeekDatePtr^.newIdx := i;
      Break;
    end;
  end;
end;

procedure TfrmDataIntMerge.btnSave2Click(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;  
  ExcelApp2, WorkBook2: Variant;
  sfile: string;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  iFile: Integer;
  slProjs: TStringList;
  slweeks: TStringList;
  sproj: string;
  syear: string;
  idx: Integer;             
  p: PWeekDate;
  aWeekDatePtr: PWeekDate;   
  aWeekDatePtr1: PWeekDate;
  aWeekDatePtr2: PWeekDate;
  iweek: Integer;
  dtMin: TDateTime;
  slweeks_all: TStringList;
  rc: Integer;
  irow: Integer;   
  irow2: Integer;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  sfile_dest: string;
  smode: string;
  iline: Integer;
begin
  if not ExcelSaveDialog(sfile_dest) then Exit;
  
  slProjs := TStringList.Create;
  slweeks_all := TStringList.Create;
  try
    for ifile := 0 to mmoFiles.Lines.Count - 1 do
    begin
      sfile := mmoFiles.Lines[ifile];
      sproj := ExtractProjName(sfile);
      syear := '2016';
      idx := mmoYears.Lines.IndexOfName(sproj);
      if idx >= 0 then
      begin
        syear := mmoYears.Lines.ValueFromIndex[idx];
      end;
      slweeks := TStringList.Create;

      ReadWeeks(sfile, slweeks, syear, rc); 
      slProjs.AddObject(sproj + '=' + IntToStr(rc), slweeks);
    end;

    slProjs.CustomSort(StringListSortCompare);

    for ifile := 0 to slProjs.Count - 1 do
    begin
      slweeks := TStringList(slProjs.Objects[ifile]);
      aWeekDatePtr := PWeekDate(slweeks.Objects[0]);
      Memo1.Lines.Add(slProjs[ifile] + '   ' + FormatDateTime('yyyy-MM-dd', aWeekDatePtr^.dt1));

      for iweek := 0 to slweeks.Count - 1 do
      begin           
        aWeekDatePtr := PWeekDate(slweeks.Objects[iweek]);
        if not AddToList(slweeks_all, aWeekDatePtr) then
        begin
          Memo1.Lines.Add('文件 ' + slProjs[ifile] + '   ' + aWeekDatePtr^.sweek + '  ' + aWeekDatePtr^.sdate + '  日期期间不对' );
        end;
      end;
    end;


    Memo1.Lines.Add(''); 
    Memo1.Lines.Add('');
    Memo1.Lines.Add('');
    Memo1.Lines.Add('');
    Memo1.Lines.Add('');

    
    for ifile := 0 to slProjs.Count - 1 do
    begin
      slweeks := TStringList(slProjs.Objects[ifile]);
      Memo1.Lines.Add(slProjs[ifile]);
      for iweek := 0 to slweeks.Count - 1 do
      begin           
        aWeekDatePtr := PWeekDate(slweeks.Objects[iweek]);
        IndexOfWeek(slweeks_all, aWeekDatePtr);    
        Memo1.Lines.Add(aWeekDatePtr^.sweek + '  ' + aWeekDatePtr^.sdate + '  ' + IntToStr(aWeekDatePtr^.newIdx));
      end;
    end;    

  

    // 开始保存 Excel

    for ifile := 0 to slProjs.Count - 1 do
    begin
      slweeks := TStringList(slProjs.Objects[ifile]);
      Memo1.Lines.Add(slProjs[ifile]);
      for iweek := 0 to slweeks.Count - 1 do
      begin
        aWeekDatePtr := PWeekDate(slweeks.Objects[iweek]);
        IndexOfWeek(slweeks_all, aWeekDatePtr);    
        Memo1.Lines.Add(aWeekDatePtr^.sweek + '  ' + aWeekDatePtr^.sdate + '  ' + IntToStr(aWeekDatePtr^.newIdx));
      end;
    end;

        
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
      ExcelApp.Sheets[1].Delete;
    end;

    ExcelApp.Sheets[1].Activate;
    ExcelApp.Sheets[1].Name := '集成汇总';    

    try 
      irow := 1; 
      ExcelApp.Cells[irow, 1].Value := '模式';
      ExcelApp.Cells[irow, 2].Value := '项目';
      ExcelApp.Cells[irow, 3].Value := 'week';
      ExcelApp.Cells[irow, 4].Value := 'week';
      ExcelApp.Cells[irow, 5].Value := '物料编码';
      ExcelApp.Cells[irow, 6].Value := '颜色';
      ExcelApp.Cells[irow, 7].Value := '容量';
      ExcelApp.Cells[irow, 8].Value := '制式';
      ExcelApp.Cells[irow, 9].Value := '计划';
      
      MergeCells(ExcelApp, irow, 1, irow + 1, 1);
      MergeCells(ExcelApp, irow, 2, irow + 1, 2);
      MergeCells(ExcelApp, irow, 3, irow + 1, 3);
      MergeCells(ExcelApp, irow, 4, irow + 1, 4);
      MergeCells(ExcelApp, irow, 5, irow + 1, 5);
      MergeCells(ExcelApp, irow, 6, irow + 1, 6);
      MergeCells(ExcelApp, irow, 7, irow + 1, 7);
      MergeCells(ExcelApp, irow, 8, irow + 1, 8);
      MergeCells(ExcelApp, irow, 9, irow + 1, 9);

      for iweek := 0 to slweeks_all.Count - 1 do
      begin
        aWeekDatePtr := PWeekDate(slweeks_all.Objects[iweek]);
        ExcelApp.Cells[irow, iweek + 10].Value := aWeekDatePtr^.sweek;
        ExcelApp.Cells[irow + 1, iweek + 10].Value := aWeekDatePtr^.sdate;
      end;

      irow := 3;
      for ifile := 0 to slProjs.Count - 1 do
      begin
        slweeks := TStringList(slProjs.Objects[ifile]);
        rc := StrToInt( slProjs.ValueFromIndex[ifile] )  ;
        smode := mmoMode.Lines.Values[slProjs.Names[ifile]];


        ExcelApp2 := CreateOleObject('Excel.Application' );
        ExcelApp2.Visible := False;
        ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
        try

          WorkBook2 := ExcelApp2.WorkBooks.Open(mmoFiles.Lines[ifile]);

          try
            iSheetCount := ExcelApp2.Sheets.Count;
            for iSheet := 1 to iSheetCount do
            begin
              if not ExcelApp2.Sheets[iSheet].Visible then Continue;

              ExcelApp2.Sheets[iSheet].Activate;

              sSheet := ExcelApp2.Sheets[iSheet].Name;

              irow2 := 1;
              stitle1 := ExcelApp2.Cells[irow2, 1].Value;
              stitle2 := ExcelApp2.Cells[irow2, 2].Value;
              stitle3 := ExcelApp2.Cells[irow2, 3].Value;
              stitle4 := ExcelApp2.Cells[irow2, 4].Value;
              stitle5 := ExcelApp2.Cells[irow2, 5].Value;
              stitle6 := ExcelApp2.Cells[irow2, 6].Value;
              stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;

              if stitle <> 'week物料编码颜色容量制式计划' then
              begin
        
                Continue;
              end;

              ExcelApp2.Range[ ExcelApp2.Cells[3, 1], ExcelApp2.Cells[rc + 2, 6] ].Copy;
              ExcelApp.Cells[irow, 4].Select;
              ExcelApp.ActiveSheet.Paste;
                                  
              for iweek := 0 to slweeks.Count - 1 do
              begin
                p := PWeekDate(slweeks.Objects[iweek]);
                ExcelApp2.Range[ ExcelApp2.Cells[3, p^.oldCol], ExcelApp2.Cells[rc + 2, p^.oldCol] ].Copy;
                ExcelApp.Cells[irow, p^.newIdx + 10].Select;
                ExcelApp.Paste;
              end;               
              ExcelApp2.Range[ ExcelApp2.Cells[1, 1], ExcelApp2.Cells[1, 1] ].Copy;


              for iline := 0 to rc - 1 do
              begin
                ExcelApp.Cells[irow + iline, 1].Value := smode;
                ExcelApp.Cells[irow + iline, 2].Value := slProjs.Names[ifile];
                ExcelApp.Cells[irow + iline, 3].Value := ChangeFileExt( ExtractFileName(mmoFiles.Lines[ifile]), '') ;
              end;

              irow := irow + rc;
              
              Break;
            end;



          finally
            ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
            WorkBook2.Close;
          end;

        finally
          ExcelApp2.Visible := True;
          ExcelApp2.Quit; 
        end;  
      end;

      AddBorder(ExcelApp, 1, 1, irow - 1, slweeks_all.Count + 9);
      AddColor(ExcelApp, 1, 1, 2, slweeks_all.Count + 9, $DBDCF2);   
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[2, slweeks_all.Count + 9] ].HorizontalAlignment := xlCenter;

      WorkBook.SaveAs(sfile_dest);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end; 
    
  finally
    for ifile := 0 to slProjs.Count - 1 do
    begin
      slweeks := TStringList(slProjs.Objects[ifile]);
      for iweek := 0 to slweeks.Count - 1 do
      begin
        aWeekDatePtr := PWeekDate(slweeks.Objects[iweek]);
        Dispose(aWeekDatePtr);
      end;
      slweeks.Free;
    end;
    slProjs.Free;
    slweeks_all.Free;
  end;
  MessageBox(Handle, '完成', '提示', 0);
end;

end.
