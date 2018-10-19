unit MergeS620ImportTemplateWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SAPImportS620Reader, StdCtrls, ComCtrls, ToolWin, ImgList, CommUtils,
  ComObj, IniFiles;

type
  TfrmMergeS620ImportTemplate = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    tbSave: TToolButton;
    mmofiles: TMemo;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    Memo1: TMemo;
    procedure ToolButton4Click(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    procedure OnLog(const str: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmMergeS620ImportTemplate.ShowForm;
var
  frmMergeS620ImportTemplate: TfrmMergeS620ImportTemplate;
begin
  frmMergeS620ImportTemplate := TfrmMergeS620ImportTemplate.Create(nil);
  try
    frmMergeS620ImportTemplate.ShowModal;
  finally
    frmMergeS620ImportTemplate.Free;
  end;
end;
    
procedure TfrmMergeS620ImportTemplate.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, mmofiles.Name, '');
    mmofiles.Lines.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);
  finally
    ini.Free;
  end;
end;

procedure TfrmMergeS620ImportTemplate.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := StringReplace(mmofiles.Text, #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmofiles.Name, s);
  finally
    ini.Free;
  end;
end;

procedure TfrmMergeS620ImportTemplate.ToolButton4Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmMergeS620ImportTemplate.ToolButton1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmofiles.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );
end;
 
procedure TfrmMergeS620ImportTemplate.OnLog(const str: string);
begin
  Memo1.Lines.Add(str);
end;

procedure TfrmMergeS620ImportTemplate.tbSaveClick(Sender: TObject);
  function ColOfDate(sldate: TStringList; const sdt: string): Integer;
  var
    idx: Integer;
  begin
    Result := -1;
    idx := sldate.IndexOf(sdt);
    if idx >= 0 then
    begin
      Result := Integer(sldate.Objects[idx]);
    end;
  end;
var
  sfile: string;   
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  icol: Integer;
  lst: TList;
  i: Integer;
  aSAPImportS620Reader: TSAPImportS620Reader;
  sldate: TStringList;
  idate: Integer;
  iline: Integer;
  aSAPImportS620Line: TSAPImportS620Line;
  ptrSAPImportS620Col: PSAPImportS620Col;
  icolExtra: Integer;
  dOE, dOD: Double;
  dIn, dOut: Double;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  sldate := TStringList.Create;
  lst := TList.Create;

  for i := 0 to mmofiles.Lines.Count - 1 do
  begin
    aSAPImportS620Reader := TSAPImportS620Reader.Create(mmofiles.Lines[i], OnLog);
    lst.Add(aSAPImportS620Reader);

    for idate := 0 to aSAPImportS620Reader.FDates.Count - 1 do
    begin
      if sldate.IndexOf(aSAPImportS620Reader.FDates[idate]) < 0 then
      begin
        sldate.Add(aSAPImportS620Reader.FDates[idate]);
      end;
    end;
  end;


  dIn := 0;
  dOut := 0;


  sldate.Sort;

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

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    try

      ExcelApp.Sheets[1].Activate;
      ExcelApp.Sheets[1].Name := '国内';

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'MATNR';
      ExcelApp.Cells[irow, 2].Value := 'BERID';

      icol := 3;
      for idate := 0 to sldate.Count - 1 do
      begin
        sldate.Objects[idate] := TObject(icol);
        ExcelApp.Cells[irow, icol].Value := sldate[idate];
        icol := icol + 1;
      end;

      icolExtra := icol + 3;    
      ExcelApp.Cells[irow, icolExtra].Value := '产品编码';
      ExcelApp.Cells[irow, icolExtra + 1].Value := '版本';
      ExcelApp.Cells[irow, icolExtra + 2].Value := '颜色';
      ExcelApp.Cells[irow, icolExtra + 3].Value := '容量';
      ExcelApp.Cells[irow, icolExtra + 4].Value := '项目';


      irow := 2;
      for i := 0 to lst.Count - 1 do
      begin
        aSAPImportS620Reader := TSAPImportS620Reader(lst[i]);
        for iline := 0 to aSAPImportS620Reader.Count - 1 do
        begin
          aSAPImportS620Line := aSAPImportS620Reader.Items[iline];

          if Copy(aSAPImportS620Line.sMATNR, 1, 3) = '90.' then Continue;

          ExcelApp.Cells[irow, 1].Value := aSAPImportS620Line.sMATNR;
          ExcelApp.Cells[irow, 2].Value := aSAPImportS620Line.sBERID;
                                                                               
          ExcelApp.Cells[irow, icolExtra - 1].Value := '=' + GetRef(1) + IntToStr(irow) + '=' + GetRef(icolExtra) + IntToStr(irow);
          
          ExcelApp.Cells[irow, icolExtra].Value := aSAPImportS620Line.snumber;
          ExcelApp.Cells[irow, icolExtra + 1].Value := aSAPImportS620Line.sver;
          ExcelApp.Cells[irow, icolExtra + 2].Value := aSAPImportS620Line.scolor;
          ExcelApp.Cells[irow, icolExtra + 3].Value := aSAPImportS620Line.scap;
          ExcelApp.Cells[irow, icolExtra + 4].Value := aSAPImportS620Line.sproj;

          for idate := 0 to aSAPImportS620Line.Count - 1 do
          begin
            ptrSAPImportS620Col := aSAPImportS620Line.Items[idate];
            icol := ColOfDate(sldate, ptrSAPImportS620Col^.sdt);
            if icol = -1 then
            begin
              raise Exception.Create('col of date not found ' + ptrSAPImportS620Col^.sdt);
            end;
            ExcelApp.Cells[irow, icol].Value := ptrSAPImportS620Col^.dQty;
            dIn := dIn + ptrSAPImportS620Col^.dQty;
          end;

          irow := irow + 1;
        end;
      end;

      
      //////////////////////////////////////////////////////////////////////////
      //////////////////////////////////////////////////////////////////////////
      //////////////////////////////////////////////////////////////////////////
      

      WorkBook.Sheets.Add(after:=WorkBook.Sheets[1]);
          
      ExcelApp.Sheets[2].Activate;
      ExcelApp.Sheets[2].Name := '海外';

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'MATNR';
      ExcelApp.Cells[irow, 2].Value := 'BERID';

      icol := 3;
      for idate := 0 to sldate.Count - 1 do
      begin
        sldate.Objects[idate] := TObject(icol);
        ExcelApp.Cells[irow, icol].Value := sldate[idate];
        icol := icol + 1;
      end;                

      icolExtra := icol + 3;
      ExcelApp.Cells[irow, icolExtra].Value := '产品编码';
      ExcelApp.Cells[irow, icolExtra + 1].Value := '版本';
      ExcelApp.Cells[irow, icolExtra + 2].Value := '颜色';
      ExcelApp.Cells[irow, icolExtra + 3].Value := '容量';
      ExcelApp.Cells[irow, icolExtra + 4].Value := '项目';



      irow := 2;
      for i := 0 to lst.Count - 1 do
      begin
        aSAPImportS620Reader := TSAPImportS620Reader(lst[i]);
        for iline := 0 to aSAPImportS620Reader.Count - 1 do
        begin
          aSAPImportS620Line := aSAPImportS620Reader.Items[iline];

          if Copy(aSAPImportS620Line.sMATNR, 1, 3) <> '90.' then Continue;

          ExcelApp.Cells[irow, 1].Value := aSAPImportS620Line.sMATNR;
          ExcelApp.Cells[irow, 2].Value := aSAPImportS620Line.sBERID;
                                                                                                                                         
          ExcelApp.Cells[irow, icolExtra - 1].Value := '=' + GetRef(1) + IntToStr(irow) + '=' + GetRef(icolExtra) + IntToStr(irow);
          
          ExcelApp.Cells[irow, icolExtra].Value := aSAPImportS620Line.snumber;
          ExcelApp.Cells[irow, icolExtra + 1].Value := aSAPImportS620Line.sver;
          ExcelApp.Cells[irow, icolExtra + 2].Value := aSAPImportS620Line.scolor;
          ExcelApp.Cells[irow, icolExtra + 3].Value := aSAPImportS620Line.scap;
          ExcelApp.Cells[irow, icolExtra + 4].Value := aSAPImportS620Line.sproj;

          for idate := 0 to aSAPImportS620Line.Count - 1 do
          begin
            ptrSAPImportS620Col := aSAPImportS620Line.Items[idate];
            icol := ColOfDate(sldate, ptrSAPImportS620Col^.sdt);
            if icol = -1 then
            begin
              raise Exception.Create('col of date not found ' + ptrSAPImportS620Col^.sdt);
            end;
            ExcelApp.Cells[irow, icol].Value := ptrSAPImportS620Col^.dQty;
            dOut := dOut + ptrSAPImportS620Col^.dQty;
          end;

          irow := irow + 1;
        end;
      end;

      sfile := ExtractFilePath(sfile) + '国内+海外 ' + FormatDateTime('yy.M.D', Now) + '.xlsx';
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;


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

    try

      ExcelApp.Sheets[1].Activate;
      ExcelApp.Sheets[1].Name := '国内';

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'MATNR';
      ExcelApp.Cells[irow, 2].Value := 'BERID';

      icol := 3;
      for idate := 0 to sldate.Count - 1 do
      begin
        sldate.Objects[idate] := TObject(icol);
        ExcelApp.Cells[irow, icol].Value := sldate[idate];
        icol := icol + 1;
      end;

      icolExtra := icol + 3;    
      ExcelApp.Cells[irow, icolExtra].Value := '产品编码';
      ExcelApp.Cells[irow, icolExtra + 1].Value := '版本';
      ExcelApp.Cells[irow, icolExtra + 2].Value := '颜色';
      ExcelApp.Cells[irow, icolExtra + 3].Value := '容量';
      ExcelApp.Cells[irow, icolExtra + 4].Value := '项目';


      irow := 2;
      for i := 0 to lst.Count - 1 do
      begin
        aSAPImportS620Reader := TSAPImportS620Reader(lst[i]);
        for iline := 0 to aSAPImportS620Reader.Count - 1 do
        begin
          aSAPImportS620Line := aSAPImportS620Reader.Items[iline];

          if Copy(aSAPImportS620Line.sMATNR, 1, 3) = '90.' then Continue;

          ExcelApp.Cells[irow, 1].Value := aSAPImportS620Line.sMATNR;
          ExcelApp.Cells[irow, 2].Value := aSAPImportS620Line.sBERID;
                                                                               
          ExcelApp.Cells[irow, icolExtra - 1].Value := '=' + GetRef(1) + IntToStr(irow) + '=' + GetRef(icolExtra) + IntToStr(irow);
          
          ExcelApp.Cells[irow, icolExtra].Value := aSAPImportS620Line.snumber;
          ExcelApp.Cells[irow, icolExtra + 1].Value := aSAPImportS620Line.sver;
          ExcelApp.Cells[irow, icolExtra + 2].Value := aSAPImportS620Line.scolor;
          ExcelApp.Cells[irow, icolExtra + 3].Value := aSAPImportS620Line.scap;
          ExcelApp.Cells[irow, icolExtra + 4].Value := aSAPImportS620Line.sproj;

          for idate := 0 to aSAPImportS620Line.Count - 1 do
          begin
            ptrSAPImportS620Col := aSAPImportS620Line.Items[idate];
            icol := ColOfDate(sldate, ptrSAPImportS620Col^.sdt);
            if icol = -1 then
            begin
              raise Exception.Create('col of date not found ' + ptrSAPImportS620Col^.sdt);
            end;
            ExcelApp.Cells[irow, icol].Value := ptrSAPImportS620Col^.dQty;
          end;

          irow := irow + 1;
        end;
      end;

      sfile := ExtractFilePath(sfile) + 'OEM ODM 国内  ' + FormatDateTime('yy.M.D', Now) + '.xlsx';
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


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

    try
 
      ExcelApp.Sheets[1].Activate;
      ExcelApp.Sheets[1].Name := '海外';

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'MATNR';
      ExcelApp.Cells[irow, 2].Value := 'BERID';

      icol := 3;
      for idate := 0 to sldate.Count - 1 do
      begin
        sldate.Objects[idate] := TObject(icol);
        ExcelApp.Cells[irow, icol].Value := sldate[idate];
        icol := icol + 1;
      end;                

      icolExtra := icol + 3;
      ExcelApp.Cells[irow, icolExtra].Value := '产品编码';
      ExcelApp.Cells[irow, icolExtra + 1].Value := '版本';
      ExcelApp.Cells[irow, icolExtra + 2].Value := '颜色';
      ExcelApp.Cells[irow, icolExtra + 3].Value := '容量';
      ExcelApp.Cells[irow, icolExtra + 4].Value := '项目';



      irow := 2;
      for i := 0 to lst.Count - 1 do
      begin
        aSAPImportS620Reader := TSAPImportS620Reader(lst[i]);
        for iline := 0 to aSAPImportS620Reader.Count - 1 do
        begin
          aSAPImportS620Line := aSAPImportS620Reader.Items[iline];

          if Copy(aSAPImportS620Line.sMATNR, 1, 3) <> '90.' then Continue;

          ExcelApp.Cells[irow, 1].Value := aSAPImportS620Line.sMATNR;
          ExcelApp.Cells[irow, 2].Value := aSAPImportS620Line.sBERID;
                                                                                                                                         
          ExcelApp.Cells[irow, icolExtra - 1].Value := '=' + GetRef(1) + IntToStr(irow) + '=' + GetRef(icolExtra) + IntToStr(irow);
          
          ExcelApp.Cells[irow, icolExtra].Value := aSAPImportS620Line.snumber;
          ExcelApp.Cells[irow, icolExtra + 1].Value := aSAPImportS620Line.sver;
          ExcelApp.Cells[irow, icolExtra + 2].Value := aSAPImportS620Line.scolor;
          ExcelApp.Cells[irow, icolExtra + 3].Value := aSAPImportS620Line.scap;
          ExcelApp.Cells[irow, icolExtra + 4].Value := aSAPImportS620Line.sproj;

          for idate := 0 to aSAPImportS620Line.Count - 1 do
          begin
            ptrSAPImportS620Col := aSAPImportS620Line.Items[idate];
            icol := ColOfDate(sldate, ptrSAPImportS620Col^.sdt);
            if icol = -1 then
            begin
              raise Exception.Create('col of date not found ' + ptrSAPImportS620Col^.sdt);
            end;
            ExcelApp.Cells[irow, icol].Value := ptrSAPImportS620Col^.dQty;
          end;

          irow := irow + 1;
        end;
      end;

      sfile := ExtractFilePath(sfile) + 'OEM ODM 海外 ' + FormatDateTime('yy.M.D', Now) + '.xlsx';
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;
             
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


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

    try
 
      ExcelApp.Sheets[1].Activate;
      ExcelApp.Sheets[1].Name := 'Check';

      ExcelApp.Cells[1, 1].Value := '国内';
      ExcelApp.Cells[2, 1].Value := '海外';
       
      ExcelApp.Cells[1, 3].Value := dIn;
      ExcelApp.Cells[2, 3].Value := dOut;  
      ExcelApp.Cells[3, 3].Value := '=SUM(' + GetRef(3) + '1:' + GetRef(3) + '2)';

      for i := 0 to lst.Count - 1 do
      begin
        aSAPImportS620Reader := TSAPImportS620Reader(lst[i]);
        ExcelApp.Cells[i + 1, 5].Value := aSAPImportS620Reader.FSum;
      end;
      ExcelApp.Cells[lst.Count + 1, 5].Value := '=SUM(' + GetRef(5) + '1:' + GetRef(5) + IntToStr(lst.Count) + ')';
                      
      sfile := ExtractFilePath(sfile) + 'Check' + FormatDateTime('yy.M.D', Now) + '.xlsx';
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit; 
    end;

  finally
    for i := 0 to lst.Count - 1 do
    begin
      aSAPImportS620Reader := TSAPImportS620Reader(lst[i]);
      aSAPImportS620Reader.Free;
    end;
    lst.Free;

    sldate.Free;
  end;


  MessageBox(Handle, '完成', '提示', 0);
end;

end.
