unit SAP2SOPWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin, ComObj, DateUtils,
  Buttons, IniFiles, CommUtils;

type 
  TfrmSAP2SOP = class(TForm)
    leSOP: TLabeledEdit;
    btnSOP: TButton;
    OpenDialog1: TOpenDialog;
    ToolBar1: TToolBar;
    ImageList1: TImageList;
    SaveDialog1: TSaveDialog;
    tbSave: TToolButton;
    ToolButton1: TToolButton;
    tbQuit: TToolButton;
    GroupBox1: TGroupBox;
    Memo1: TMemo;
    leS618: TLabeledEdit;
    btnS618: TButton;
    cbPlan: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure btnSOPClick(Sender: TObject);
    procedure tbQuitClick(Sender: TObject);
    procedure btnS618Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

uses SOPReaderUnit, SAPS618Reader;

{$R *.dfm}
            
class procedure TfrmSAP2SOP.ShowForm;
var
  frmSOP2SAP: TfrmSAP2SOP;
begin
  frmSOP2SAP := TfrmSAP2SOP.Create(nil);
  try
    frmSOP2SAP.ShowModal;
  finally
    frmSOP2SAP.Free;
  end;
end;
   
procedure TfrmSAP2SOP.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leSOP.Text := ini.ReadString(self.ClassName, leSOP.Name, '');
    leS618.Text := ini.ReadString(self.ClassName, leS618.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSAP2SOP.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leSOP.Name, leSOP.Text);
    ini.WriteString(self.ClassName, leS618.Name, leS618.Text);
  finally
    ini.Free;
  end;
end; 
                
procedure TfrmSAP2SOP.btnSOPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOP.Text := sfile;
end;
  
procedure TfrmSAP2SOP.btnS618Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leS618.Text := sfile;
end;

procedure TfrmSAP2SOP.tbQuitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSAP2SOP.tbSaveClick(Sender: TObject);
var
  sfile: string;        
  ExcelApp, WorkBook: Variant;
  aSOPReader: TSOPReader;
  iProj: Integer;
  aSOPProj: TSOPProj;
  irow: Integer;
  iLine: Integer;
  aSOPLine: TSOPLine;
  iweek: Integer;
  aSOPCol: TSOPCol;
  slDate: TStringList;
  sdate: string;
  iQty: Integer; 
  idx: Integer;
  bBlank: Boolean;
  slNumberDetail: TStringList;
  slProjYear: TStringList;

  aSAPS618Reader: TSAPS618Reader;
  sver: string;
  irow0: Integer;
  sweek: string;
  sdt1, sdt2: string;
  dt1: TDateTime;
  aSAPS618: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
  irow1, irow2: Integer;
  slver: TStringList;
  slcap: TStringList;
  slcol: TStringList;
  svalue: string;
  iCell: integer;
  iMonth0: Integer;
  icol: Integer;
  icol1: Integer;
  icolMax: Integer;
  slMonth: TStringList;
begin
  sfile := '调整后的要货计划' + FormatDateTime('yyyyMMdd', Now) + '.xlsx';
  if not ExcelSaveDialog(sfile) then Exit;

  slProjYear := TStringList.Create;

  aSAPS618Reader := TSAPS618Reader.Create(leS618.Text, cbPlan.Text);

  slver := TStringList.Create;
  slcap := TStringList.Create;
  slcol := TStringList.Create;

  slMonth := TStringList.Create;

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
      ExcelApp.Sheets[2].Delete;
    end;

    ExcelApp.Sheets[1].Activate;
    //ExcelApp.Sheets[1].Name := '产品预测单';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := 'MATNR';

    ExcelApp.Columns[1].ColumnWidth := 16;

    slDate := TStringList.Create;
    slNumberDetail := TStringList.Create;
    try
      aSOPReader := TSOPReader.Create(slProjYear, leSOP.Text);
      try
        for iProj := 0 to aSOPReader.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader.Projs[iProj];

          slMonth.Clear;

          if ExcelApp.Sheets.Count < iProj + 1 then
          begin
            ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[iProj]);
          end;
          ExcelApp.Sheets[iProj + 1].Name := aSOPProj.FName;
          
          ExcelApp.Cells[1, 1].Value := '项目';
          ExcelApp.Cells[1, 2].Value := '整机/裸机';
          ExcelApp.Cells[1, 3].Value := '包装';
          ExcelApp.Cells[1, 4].Value := '标准制式';
          ExcelApp.Cells[1, 5].Value := 'MRP区域';
          ExcelApp.Cells[1, 6].Value := '制式';
          ExcelApp.Cells[1, 7].Value := '物料编码';
          ExcelApp.Cells[1, 8].Value := '颜色';
          ExcelApp.Cells[1, 9].Value := '容量';
          MergeCells(ExcelApp, 1, 1, 2, 1);
          MergeCells(ExcelApp, 1, 2, 2, 2);
          MergeCells(ExcelApp, 1, 3, 2, 3);
          MergeCells(ExcelApp, 1, 4, 2, 4); 
          MergeCells(ExcelApp, 1, 5, 2, 5);
          MergeCells(ExcelApp, 1, 6, 2, 6);
          MergeCells(ExcelApp, 1, 7, 2, 7);
          MergeCells(ExcelApp, 1, 8, 2, 8);
          MergeCells(ExcelApp, 1, 9, 2, 9);

          iMonth0 := 0;
          icol := 10;
          icol1 := icol;
          for iweek := 0 to aSAPS618Reader.slWeek.Count - 1 do
          begin
            sweek := aSAPS618Reader.slWeek.Names[iweek];
            sdt1 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
            sweek := Copy(sweek,  Pos('-', sweek) + 1, Length(sweek));
            sdt2 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
            sweek := sdt1 + '-' + sdt2;

            sdt1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sdt1, '/', '-', [rfReplaceAll]);
            dt1 := myStrToDateTime(sdt1);


            if iMonth0 = 0 then
            begin
              iMonth0 := MonthOf(dt1);
            end
            else
            begin
              if iMonth0 <> MonthOf(dt1) then
              begin
                ExcelApp.Cells[1, icol].Value := IntToStr(iMonth0) + '月';
                MergeCells(ExcelApp, 1, icol, 2, icol);
                slMonth.Add(IntToStr(icol));
                iMonth0 := MonthOf(dt1);
                icol := icol + 1;
                icol1 := icol;
              end;
            end;
                                                                     
            ExcelApp.Cells[1, icol].Value := 'WK' + aSAPS618Reader.slWeek.ValueFromIndex[iweek];
            ExcelApp.Cells[2, icol].Value := sweek;

            icol := icol + 1;
            
          end;

          if icol1 <> icol then
          begin
            ExcelApp.Cells[1, icol].Value := IntToStr(iMonth0) + '月';
            MergeCells(ExcelApp, 1, icol, 2, icol);
            slMonth.Add(IntToStr(icol));
            icol := icol + 1;
          end;
          icolMax := icol - 1;

                           
          irow := 3;
          sver := '';
          irow0 := 0;
          
          for iLine := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iLine];

            ExcelApp.Cells[irow, 1].Value := aSOPProj.FName;  
            ExcelApp.Cells[irow, 2].Value := aSOPLine.sFG;
            ExcelApp.Cells[irow, 3].Value := aSOPLine.sPkg;
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sStdVer;
            ExcelApp.Cells[irow, 5].Value := aSOPLine.sMRPArea;
            // 新的一个版本
            if sver <> aSOPLine.sVer then
            begin
              if sver <> '' then
              begin
                MergeCells(ExcelApp, irow0, 6, irow - 1, 6);
              end;
              ExcelApp.Cells[irow, 6].Value := aSOPLine.sVer;
              sver := aSOPLine.sVer;
              irow0 := irow;
            end;

            if aSOPLine.sNumber = '83.68.36810002CN' then
            begin
              Sleep(1);
            end;

            ExcelApp.Cells[irow, 7].Value := aSOPLine.sNumber;
            ExcelApp.Cells[irow, 8].Value := aSOPLine.sColor;
            ExcelApp.Cells[irow, 9].Value := aSOPLine.sCap;

            aSAPS618 := aSAPS618Reader.GetItem(aSOPLine.sNumber, aSOPLine.sMRPArea); // GetItem里是新Create的，调用者要负责释放掉
            if aSAPS618 <> nil then
            begin           
              
              iMonth0 := 0;
              icol := 10;
              icol1 := icol;
              for iweek := 0 to aSAPS618.Count - 1 do
              begin
                aSAPS618ColPtr := aSAPS618.Items[iweek];
                   

                if iMonth0 = 0 then
                begin
                  iMonth0 := MonthOf(aSAPS618ColPtr^.dt1);
                end
                else
                begin
                  if iMonth0 <> MonthOf(aSAPS618ColPtr^.dt1) then
                  begin
                    ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                    iMonth0 := MonthOf(aSAPS618ColPtr^.dt1);
                    icol := icol + 1;
                    icol1 := icol;
                  end;
                end;


                ExcelApp.Cells[irow, icol].Value := aSAPS618ColPtr^.dqty;
                icol := icol + 1;

              end;
              if icol <> icol1 then
              begin
                ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
              end;
              aSAPS618.Free; // GetItem里面是新Create的，要释放掉
            end;

            idx := slver.IndexOf(aSOPLine.sVer);
            if idx < 0 then
            begin
              idx := slver.AddObject(aSOPLine.sVer, TStringList.Create);
            end;    
            TStringList( slver.Objects[idx] ).Add(IntToStr(irow));
                  
            idx := slcap.IndexOf(aSOPLine.sCap);
            if idx < 0 then
            begin
              idx := slcap.AddObject(aSOPLine.sCap, TStringList.Create);
            end;    
            TStringList( slcap.Objects[idx] ).Add(IntToStr(irow));
                           
            idx := slcol.IndexOf(aSOPLine.sColor);
            if idx < 0 then
            begin
              idx := slcol.AddObject(aSOPLine.sColor, TStringList.Create);
            end;    
            TStringList( slcol.Objects[idx] ).Add(IntToStr(irow));
          
            irow := irow + 1;
          end;
          if sver <> '' then
          begin
            MergeCells(ExcelApp, irow0, 6, irow - 1, 6);
          end;

          irow1 := 3;
          irow2 := irow - 1;
                                     
          AddBorder(ExcelApp, 1, 1, irow - 1, icolMax);

          irow0 := irow;

          ExcelApp.Cells[irow, 6].Value := aSOPProj.FName + #13#10'TOTAL';

          for idx := 0 to slver.Count - 1 do
          begin
            ExcelApp.Cells[irow, 7].Value := slver[idx];
            MergeCells(ExcelApp, irow, 7, irow, 9);

            iMonth0 := 0;
            icol := 10;
            icol1 := icol;
            for iweek := 0 to aSAPS618Reader.slWeek.Count - 1 do
            begin

              sweek := aSAPS618Reader.slWeek[iweek];
              sdt1 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
              sweek := Copy(sweek,  Pos('-', sweek) + 1, Length(sweek));
              sdt2 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
              sweek := sdt1 + '-' + sdt2;

              sdt1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sdt1, '/', '-', [rfReplaceAll]);
              dt1 := myStrToDateTime(sdt1);

              if iMonth0 = 0 then
              begin
                iMonth0 := MonthOf(dt1);
              end
              else
              begin
                if iMonth0 <> MonthOf(dt1) then
                begin 
                  ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                  iMonth0 := MonthOf(dt1);
                  icol := icol + 1;
                  icol1 := icol;
                end;
              end;
                          
              svalue := '=0';
              for iCell := 0 to TStringList(slver.Objects[idx]).Count - 1 do
              begin
                svalue := svalue + '+' + GetRef(icol) + TStringList(slver.Objects[idx])[iCell];
              end;
              ExcelApp.Cells[irow, icol].Value := svalue;
              icol := icol + 1;

              //////////////////////////////////////////////////////////////////
                          
            end;

            if icol1 <> icol then
            begin
              ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            end;

            irow := irow + 1;
          end;

              
          for idx := 0 to slcap.Count - 1 do
          begin
            ExcelApp.Cells[irow, 7].Value := slcap[idx];
            MergeCells(ExcelApp, irow, 7, irow, 9);

            iMonth0 := 0;
            icol := 10;
            icol1 := icol;
            for iweek := 0 to aSAPS618Reader.slWeek.Count - 1 do
            begin

              sweek := aSAPS618Reader.slWeek[iweek];
              sdt1 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
              sweek := Copy(sweek,  Pos('-', sweek) + 1, Length(sweek));
              sdt2 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
              sweek := sdt1 + '-' + sdt2;

              sdt1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sdt1, '/', '-', [rfReplaceAll]);
              dt1 := myStrToDateTime(sdt1);

              if iMonth0 = 0 then
              begin
                iMonth0 := MonthOf(dt1);
              end
              else
              begin
                if iMonth0 <> MonthOf(dt1) then
                begin 
                  ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                  iMonth0 := MonthOf(dt1);
                  icol := icol + 1;
                  icol1 := icol;
                end;
              end;

                          
              svalue := '=0';
              for iCell := 0 to TStringList(slcap.Objects[idx]).Count - 1 do
              begin
                svalue := svalue + '+' + GetRef(icol) + TStringList(slcap.Objects[idx])[iCell];
              end;
              ExcelApp.Cells[irow, icol].Value := svalue;
              icol := icol + 1;

              //////////////////////////////////////////////////////////////////
                              
            end;

            if icol1 <> icol then
            begin
              ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            end;

            irow := irow + 1;
          end;
            
          for idx := 0 to slcol.Count - 1 do
          begin
            ExcelApp.Cells[irow, 7].Value := slcol[idx];
            MergeCells(ExcelApp, irow, 7, irow, 9);
                   
            iMonth0 := 0;
            icol := 10;
            icol1 := icol;
            for iweek := 0 to aSAPS618Reader.slWeek.Count - 1 do
            begin

              sweek := aSAPS618Reader.slWeek[iweek];
              sdt1 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
              sweek := Copy(sweek,  Pos('-', sweek) + 1, Length(sweek));
              sdt2 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
              sweek := sdt1 + '-' + sdt2;

              sdt1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sdt1, '/', '-', [rfReplaceAll]);
              dt1 := myStrToDateTime(sdt1);

              if iMonth0 = 0 then
              begin
                iMonth0 := MonthOf(dt1);
              end
              else
              begin
                if iMonth0 <> MonthOf(dt1) then
                begin 
                  ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

                  iMonth0 := MonthOf(dt1);
                  icol := icol + 1;
                  icol1 := icol;
                end;
              end;

                          
              svalue := '=0';
              for iCell := 0 to TStringList(slcol.Objects[idx]).Count - 1 do
              begin
                svalue := svalue + '+' + GetRef(icol) + TStringList(slcol.Objects[idx])[iCell];
              end;
              ExcelApp.Cells[irow, icol].Value := svalue;   
              icol := icol + 1;

              //////////////////////////////////////////////////////////////////
              
            end;
                
            if icol1 <> icol then
            begin
              ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            end;

            irow := irow + 1;
          end;





          MergeCells(ExcelApp, irow0, 6, irow, 6);

          ExcelApp.Cells[irow, 7].Value := 'TOTAL';
          MergeCells(ExcelApp, irow, 7, irow, 9);
                   
          iMonth0 := 0;
          icol := 10;
          icol1 := icol;
          for iweek := 0 to aSAPS618Reader.slWeek.Count - 1 do
          begin


            sweek := aSAPS618Reader.slWeek[iweek];
            sdt1 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
            sweek := Copy(sweek,  Pos('-', sweek) + 1, Length(sweek));
            sdt2 := Copy(sweek, 1, 2) + '/' + Copy(sweek, 3, 2);
            sweek := sdt1 + '-' + sdt2;

            sdt1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sdt1, '/', '-', [rfReplaceAll]);
            dt1 := myStrToDateTime(sdt1);

            if iMonth0 = 0 then
            begin
              iMonth0 := MonthOf(dt1);
            end
            else
            begin
              if iMonth0 <> MonthOf(dt1) then
              begin 
                ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

                iMonth0 := MonthOf(dt1);  
                icol := icol + 1;
                icol1 := icol;
              end;
            end;

                      
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';

            icol := icol + 1;

            //////////////////////////////////////////////////////////////////
              
          end;
              
          if icol1 <> icol then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
          end;

          AddBorder(ExcelApp, irow0, 6, irow, icolMax);

          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, icolMax]].Font.Name := '微软雅黑';
          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, icolMax]].Font.Size := 9;      
          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, icolMax]].HorizontalAlignment := xlCenter;

          for idx := 0 to slMonth.Count - 1 do
          begin
            icol := StrToInt(slMonth[idx]);
            ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[irow, icol]].Font.Size := 10; 
            ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[irow, icol]].Font.Bold  := True;  
            ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[irow, icol]].Interior.Color := $CDFFFF;
          end;

          ExcelApp.Range[ ExcelApp.Cells[3, 10], ExcelApp.Cells[3, 10] ].Select;
          ExcelApp.ActiveWindow.FreezePanes := True;


          for idx := 0 to slver.Count - 1 do
          begin
            slver.Objects[idx].Free;
          end;
          slver.Clear;

          for idx := 0 to slcap.Count - 1 do
          begin
            slcap.Objects[idx].Free;
          end;
          slcap.Clear;

          for idx := 0 to slcol.Count - 1 do
          begin
            slcol.Objects[idx].Free;
          end;
          slcol.Clear;


        end; 
      finally
        aSOPReader.Free;
      end;

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;

      slDate.Free;
      slNumberDetail.Free;

      slProjYear.Free;
    end;
  finally
    aSAPS618Reader.Free;
 
    slver.Free;
    slcap.Free;
    slcol.Free;

    slMonth.Free;

  end;
  MessageBox(Handle, '完成', '提示', 0);
end;

end.


