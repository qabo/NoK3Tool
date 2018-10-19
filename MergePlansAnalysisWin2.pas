unit MergePlansAnalysisWin2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, ComObj, StdCtrls, ExtCtrls, DateUtils,
  CommUtils, IniFiles, Clipbrd, Excel2000, ADODB;
  
const
  xlCenter = -4108;

type
  TfrmMergePlansAnalysis2 = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    ToolButton5: TToolButton;
    tbOEM: TToolButton;
    ProgressBar1: TProgressBar;
    ToolButton7: TToolButton;
    btnExit: TToolButton;
    Memo1: TMemo;
    leSUM: TLabeledEdit;
    btnSUM: TButton;
    leAnalysis: TLabeledEdit;
    Button1: TButton;
    leWeek: TLabeledEdit;
    tbODM: TToolButton;
    ToolButton2: TToolButton;
    procedure btnExitClick(Sender: TObject);
    procedure tbOEMClick(Sender: TObject);
    procedure btnSUMClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbODMClick(Sender: TObject);
  private
    { Private declarations }  
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
 
type
  
  TWeekRecord = packed record
    sweek: string;
    sqty: string;
    scomment: string
  end;
  PWeekRecord = ^TWeekRecord;

  TPlan = class
  private
    smode: string;
    sproj: string;
    sweek: string;
    snumber: string;
    scolor: string;
    scap: string;
    sver: string;
    splan: string;
    slDemand: TStringList;
    slSOP: TStringList;
    slMPS : TStringList;
    slSch: TStringList;
    slAct: TStringList;
    slStk: TStringList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
  end;
   
  TSOPvsDemand = class
  private
    smode: string;
    sproj: string;
    sweek: string;
    snumber: string;
    scolor: string;
    scap: string;
    sver: string;
    splan: string;
    slDemand: TStringList;
    slSOP: TStringList;   
    slStk: TStringList;
    FReasons_OEM: array[TOEMSOPvsDemandSet_OEM] of TStringList;
    FReasons_ODM: array[TOEMSOPvsDemandSet_ODM] of TStringList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  TACTvsDemand = class
  private
    smode: string;
    sproj: string;
    sweek: string;
    snumber: string;
    scolor: string;
    scap: string;
    sver: string;
    splan: string;   
    slDemand: TStringList; 
    slAct: TStringList;     
    FReasons_OEM: array[TOEMACTvsDemandSet_OEM] of TStringList;
    FReasons_ODM: array[TOEMACTvsDemandSet_ODM] of TStringList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
  end;

  TACTvsSch = class
  private
    smode: string;
    sproj: string;
    sweek: string;
    snumber: string;
    scolor: string;
    scap: string;
    sver: string;
    splan: string;   
    slSch: TStringList;  
    slAct: TStringList;
    FReasons_OEM: array[TOEMACTvsSchSet_OEM] of TStringList;
    FReasons_ODM: array[TOEMACTvsSchSet_ODM] of TStringList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
  end;

procedure ClearList(sl: TStringList);
var
  i: Integer;
  p: PWeekRecord;
begin
  for i := 0 to sl.Count - 1 do
  begin
    p := PWeekRecord(sl.Objects[i]);
    Dispose(p);
  end;
  sl.Clear;
end;

{ TPlan }

constructor TPlan.Create;
begin
  slDemand := TStringList.Create;
  slSOP := TStringList.Create;
  slMPS := TStringList.Create;
  slSch := TStringList.Create;
  slAct := TStringList.Create;
  slStk := TStringList.Create;
end;

destructor TPlan.Destroy;
begin
  Clear;
  slDemand.Free;
  slSOP.Free;
  slMPS.Free;
  slSch.Free;
  slAct.Free;
  slStk.Free;
end;
      
procedure TPlan.Clear;
begin
  ClearList(slDemand);
  ClearList(slSOP);
  ClearList(slMPS);
  ClearList(slSch);
  ClearList(slAct);
  ClearList(slStk);
end;









{ TSOPvsDemand } 


constructor TSOPvsDemand.Create;
var
  s_OEM: TOEMSOPvsDemandSet_OEM;
  s_ODM: TOEMSOPvsDemandSet_ODM;
begin
  slDemand := TStringList.Create;
  slSOP := TStringList.Create;
  slStk := TStringList.Create;
  for s_OEM := Low(TOEMSOPvsDemandSet_OEM) to High(TOEMSOPvsDemandSet_OEM) do
  begin
    FReasons_OEM[s_OEM] := TStringList.Create;
  end;
  for s_ODM := Low(TOEMSOPvsDemandSet_ODM) to High(TOEMSOPvsDemandSet_ODM) do
  begin
    FReasons_ODM[s_ODM] := TStringList.Create;
  end;
end;

destructor TSOPvsDemand.Destroy;
var
  s_OEM: TOEMSOPvsDemandSet_OEM;
  s_ODM: TOEMSOPvsDemandSet_ODM;
begin
  Clear;
  slDemand.Free;
  slSOP.Free;
  slStk.Free;
  for s_OEM := Low(TOEMSOPvsDemandSet_OEM) to High(TOEMSOPvsDemandSet_OEM) do
  begin
    FReasons_OEM[s_OEM].Free;
  end;
  for s_ODM := Low(TOEMSOPvsDemandSet_ODM) to High(TOEMSOPvsDemandSet_ODM) do
  begin
    FReasons_ODM[s_ODM].Free;
  end;
end;

procedure TSOPvsDemand.Clear;   
var
  s_OEM: TOEMSOPvsDemandSet_OEM;
  s_ODM: TOEMSOPvsDemandSet_ODM;
begin
  ClearList(slDemand);
  ClearList(slSOP);
  ClearList(slStk);
  for s_OEM := Low(TOEMSOPvsDemandSet_OEM) to High(TOEMSOPvsDemandSet_OEM) do
  begin
    ClearList(FReasons_OEM[s_OEM]);
  end;
  for s_ODM := Low(TOEMSOPvsDemandSet_ODM) to High(TOEMSOPvsDemandSet_ODM) do
  begin
    ClearList(FReasons_ODM[s_ODM]);
  end;
end;

{ TACTvsDemand }

constructor TACTvsDemand.Create;
var
  sOEM: TOEMACTvsDemandSet_OEM;
  sODM: TOEMACTvsDemandSet_ODM;
begin
  slDemand := TStringList.Create;
  slAct := TStringList.Create;
  for sOEM := Low(TOEMACTvsDemandSet_OEM) to High(TOEMACTvsDemandSet_OEM) do
  begin
    FReasons_OEM[sOEM] := TStringList.Create;
  end;
  
  for sODM := Low(TOEMACTvsDemandSet_ODM) to High(TOEMACTvsDemandSet_ODM) do
  begin
    FReasons_ODM[sODM] := TStringList.Create;
  end;
end;

destructor TACTvsDemand.Destroy;
var
  sOEM: TOEMACTvsDemandSet_OEM;
  sODM: TOEMACTvsDemandSet_ODM;
begin            
  Clear;
  slDemand.Free;
  slAct.Free;
  for sOEM := Low(TOEMACTvsDemandSet_OEM) to High(TOEMACTvsDemandSet_OEM) do
  begin
    FReasons_OEM[sOEM].Free;
  end;
  for sODM := Low(TOEMACTvsDemandSet_ODM) to High(TOEMACTvsDemandSet_ODM) do
  begin
    FReasons_ODM[sODM].Free;
  end;
end;
   

procedure TACTvsDemand.Clear;
var
  sOEM: TOEMACTvsDemandSet_OEM;
  sODM: TOEMACTvsDemandSet_ODM;
begin
  ClearList(slDemand);
  ClearList(slAct); 
  for sOEM := Low(TOEMACTvsDemandSet_OEM) to High(TOEMACTvsDemandSet_OEM) do
  begin
    ClearList(FReasons_OEM[sOEM]);
  end;
  for sODM := Low(TOEMACTvsDemandSet_ODM) to High(TOEMACTvsDemandSet_ODM) do
  begin
    ClearList(FReasons_ODM[sODM]);
  end;
end;

{ TACTvsSch }

constructor TACTvsSch.Create;
var
  sOEM: TOEMACTvsSchSet_OEM;
  sODM: TOEMACTvsSchSet_ODM;
begin         
  slSch := TStringList.Create;
  slAct := TStringList.Create;
  for sOEM := Low(TOEMACTvsSchSet_OEM) to High(TOEMACTvsSchSet_OEM) do
  begin
    FReasons_OEM[sOEM] := TStringList.Create;
  end;
  for sODM := Low(TOEMACTvsSchSet_ODM) to High(TOEMACTvsSchSet_ODM) do
  begin
    FReasons_ODM[sODM] := TStringList.Create;
  end;
end;

destructor TACTvsSch.Destroy;
var
  sOEM: TOEMACTvsSchSet_OEM;
  sODM: TOEMACTvsSchSet_ODM;
begin
  Clear;
  slSch.Free;
  slAct.Free;
  for sOEM := Low(TOEMACTvsSchSet_OEM) to High(TOEMACTvsSchSet_OEM) do
  begin
    FReasons_OEM[sOEM].Free;
  end;    
  for sODM := Low(TOEMACTvsSchSet_ODM) to High(TOEMACTvsSchSet_ODM) do
  begin
    FReasons_ODM[sODM].Free;
  end;
end;
   
procedure TACTvsSch.Clear;
var
  sOEM: TOEMACTvsSchSet_OEM;
  sODM: TOEMACTvsSchSet_ODM;
begin
  ClearList(slSch);
  ClearList(slAct);
  for sOEM := Low(TOEMACTvsSchSet_OEM) to High(TOEMACTvsSchSet_OEM) do
  begin
    ClearList(FReasons_OEM[sOEM]);
  end; 
  for sODM := Low(TOEMACTvsSchSet_ODM) to High(TOEMACTvsSchSet_ODM) do
  begin
    ClearList(FReasons_ODM[sODM]);
  end;
end;












var
  frmMergePlans: TfrmMergePlansAnalysis2;

class procedure TfrmMergePlansAnalysis2.ShowForm;
begin
  frmMergePlans := TfrmMergePlansAnalysis2.Create(nil);
  frmMergePlans.ShowModal;
  frmMergePlans.Free;
end;
     
procedure TfrmMergePlansAnalysis2.btnSUMClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSUM.Text := sfile;
end;

procedure TfrmMergePlansAnalysis2.Button1Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leAnalysis.Text := sfile;
end;
 
procedure TfrmMergePlansAnalysis2.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMergePlansAnalysis2.tbOEMClick(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;
  iSheet, iSheetCount: Integer;
  sSheet: string;
  irow: Integer;
  irow1, irow2: Integer; 
  sweek, sweek0: string;
  slweek: TStringList;
  iweek: Integer;
  icol: Integer;

  aPlan: TPlan;
  aSOPvsDemand: TSOPvsDemand;
  aACTvsDemand: TACTvsDemand;
  aACTvsSch: TACTvsSch;

  slPlan: TStringList;
  slSOPvsDemand: TStringList;
  slACTvsDemand: TStringList;
  slACTvsSch: TStringList;

  sa: TOEMSOPvsDemandSet_OEM;
  sb: TOEMACTvsDemandSet_OEM;
  sc: TOEMACTvsSchSet_OEM;

  i: Integer;   
  p: PWeekRecord;
  vComment: Variant;
  
  idx: Integer;

  sfile: string;
  dwTick: DWORD;

  va: Variant;
  splan, splan1, splan2, splan3, splan4, splan5, splan6: string;
  serr: string;
begin

  if not ExcelSaveDialog(sfile) then Exit;
                  
  Memo1.Lines.Add('----------------------------------------------------------');
  for sa := Low(TOEMSOPvsDemandSet_OEM) to High(TOEMSOPvsDemandSet_OEM) do
  begin
    Memo1.Lines.Add(CSOEMSOPvsDemand_OEM[sa]);
  end;

  Memo1.Lines.Add('----------------------------------------------------------');
  for sb := Low(TOEMACTvsDemandSet_OEM) to High(TOEMACTvsDemandSet_OEM) do
  begin
    Memo1.Lines.Add(CSOEMACTvsDemand_OEM[sb]);
  end;
                                        
  Memo1.Lines.Add('----------------------------------------------------------');
  for sc := Low(TOEMACTvsSchSet_OEM) to High(TOEMACTvsSchSet_OEM) do
  begin
    Memo1.Lines.Add(CSOEMACTvsSch_OEM[sc]);
  end;

  dwTick := GetTickCount;

  slPlan := TStringList.Create;
  slSOPvsDemand := TStringList.Create;
  slACTvsDemand := TStringList.Create;
  slACTvsSch := TStringList.Create;

  slweek := TStringList.Create;

  try
  
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';

    ExcelApp.ScreenUpdating := False;
    //ExcelApp.Calculation := xlCalculationManual;

    try
      WorkBook := ExcelApp.WorkBooks.Open(leSUM.Text);

      try
        iSheetCount := ExcelApp.Sheets.Count;
        for iSheet := 1 to iSheetCount do
        begin
          if not ExcelApp.Sheets[iSheet].Visible then Continue;

          ExcelApp.Sheets[iSheet].Activate;
        
          sSheet := ExcelApp.Sheets[iSheet].Name;

          if (sSheet <> 'OEM&ODM数据集成') and (sSheet <> '集成汇总') then
          begin
            Memo1.Lines.Add('数据集成汇总文件 sheet ' + sSheet + ' 名称不对（sheet名称要是"OEM&ODM数据集成"或者"集成汇总"）');
            Continue;
          end;
 

          /////  取 week 数量
          irow := 2;
          icol := 10;
          sweek := ExcelApp.Cells[irow, icol].Value;
          while sweek <> '' do
          begin
            slweek.Add(sweek);
            icol := icol + 1;
            sweek := ExcelApp.Cells[irow, icol].Value;
          end;

          irow := 3;
          sweek := ExcelApp.Cells[irow, 3].Value;
          while sweek <> '' do
          begin   
            splan1 := ExcelApp.Cells[irow, 9].Value;
            splan2 := ExcelApp.Cells[irow + 1, 9].Value;
            splan3 := ExcelApp.Cells[irow + 2, 9].Value;
            splan4 := ExcelApp.Cells[irow + 3, 9].Value;
            splan5 := ExcelApp.Cells[irow + 4, 9].Value;
            splan6 := ExcelApp.Cells[irow + 5, 9].Value;

            splan := splan1 + splan2 + splan3 + splan4 + splan5 + splan6;
            if splan <> '销售计划S&OP供应计划MPS排产计划实际产出期初库存' then
            begin      
              Memo1.Lines.Add(splan);
              serr := '第 ' + IntToStr(irow) + ' 行数据格式不对';
              Memo1.Lines.Add(serr);
              raise Exception.Create(serr);
            end;

            if sweek = leWeek.Text then
            begin 
              aPlan := TPlan.Create;
              aPlan.smode := ExcelApp.Cells[irow, 1].Value;
              aPlan.sproj := ExcelApp.Cells[irow, 2].Value;
              aPlan.sweek := ExcelApp.Cells[irow, 3].Value;
              aPlan.snumber := ExcelApp.Cells[irow, 5].Value;
              aPlan.scolor := ExcelApp.Cells[irow, 6].Value;
              aPlan.scap := ExcelApp.Cells[irow, 7].Value;
              aPlan.sver := ExcelApp.Cells[irow, 8].Value;
              aPlan.splan := ExcelApp.Cells[irow, 9].Value;
              slPlan.AddObject(aPlan.snumber, aPlan);

              for iweek := 0 to slweek.Count - 1 do
              begin
                icol := iweek + 10;
                  
                p := New(PWeekRecord);
                p^.sweek := slweek[iweek];
                p^.sqty := ExcelApp.Cells[irow, icol].Value;
                aPlan.slDemand.AddObject( p^.sweek, TObject(p) );
                                     
                p := New(PWeekRecord);
                p^.sweek := slweek[iweek];
                p^.sqty := ExcelApp.Cells[irow + 1, icol].Value;
                aPlan.slSOP.AddObject( p^.sweek, TObject(p)  );
                                                 
                p := New(PWeekRecord);
                p^.sweek := slweek[iweek];
                p^.sqty := ExcelApp.Cells[irow + 2, icol].Value;
                aPlan.slMPS.AddObject(  p^.sweek , TObject(p) );
                                                                      
                p := New(PWeekRecord);
                p^.sweek := slweek[iweek];
                p^.sqty := ExcelApp.Cells[irow + 3, icol].Value;
                aPlan.slSch.AddObject(  p^.sweek, TObject(p) );
                                                                        
                p := New(PWeekRecord);
                p^.sweek := slweek[iweek];
                p^.sqty := ExcelApp.Cells[irow + 4, icol].Value;
                aPlan.slAct.AddObject(  p^.sweek , TObject(p) );  
                                                                        
                p := New(PWeekRecord);
                p^.sweek := slweek[iweek];
                p^.sqty := ExcelApp.Cells[irow + 5, icol].Value;
                aPlan.slStk.AddObject(  p^.sweek , TObject(p) );
              end;
            end;
            
            irow := irow + 6;      
            sweek := ExcelApp.Cells[irow, 3].Value;

          end;

        end;
      finally
        ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook.Close;
      end;

    finally
      ExcelApp.Visible := True;
      ExcelApp.Quit; 
    end;

    Memo1.Lines.Add('读数据耗时： ' + IntToStr(GetTickCount - dwTick));

    if FileExists(leAnalysis.Text) then
    begin
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';
      try
        WorkBook := ExcelApp.WorkBooks.Open(leAnalysis.Text);

        try
          iSheetCount := ExcelApp.Sheets.Count;
          for iSheet := 1 to iSheetCount do
          begin
            if not ExcelApp.Sheets[iSheet].Visible then Continue;

            ExcelApp.Sheets[iSheet].Activate;

            sSheet := ExcelApp.Sheets[iSheet].Name;

            if sSheet = 'KPI分析-S&OP供应计划 VS 销售计划' then
            begin
              slweek.Clear;         
                    
              /////  取 week 数量
              icol := 9;
              sweek := ExcelApp.Cells[1, icol].Value;
              while sweek <> '' do
              begin
                slweek.Add(sweek);
                icol := icol + 1;
                sweek := ExcelApp.Cells[1, icol].Value;
              end;

              irow := 2;    
              sweek := ExcelApp.Cells[irow, 3].Value;
              while sweek <> '' do
              begin

                aSOPvsDemand := TSOPvsDemand.Create;
                aSOPvsDemand.smode := ExcelApp.Cells[irow, 1].Value;
                aSOPvsDemand.sproj := ExcelApp.Cells[irow, 2].Value;
                aSOPvsDemand.sweek := ExcelApp.Cells[irow, 3].Value;
                aSOPvsDemand.snumber := ExcelApp.Cells[irow, 4].Value;
                aSOPvsDemand.scolor := ExcelApp.Cells[irow, 5].Value;
                aSOPvsDemand.scap := ExcelApp.Cells[irow, 6].Value;
                aSOPvsDemand.sver := ExcelApp.Cells[irow, 7].Value;
                aSOPvsDemand.splan := ExcelApp.Cells[irow, 8].Value;
                slSOPvsDemand.AddObject(aSOPvsDemand.sproj + aSOPvsDemand.snumber, aSOPvsDemand);

                for iweek := 0 to slweek.Count - 1 do
                begin
                  icol := iweek + 9;             
                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow, icol].Value;
                  aSOPvsDemand.slDemand.AddObject( p^.sweek, TObject(p) );
                       
                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow + 1, icol].Value;
                  aSOPvsDemand.slStk.AddObject( p^.sweek , TObject(p));
                                   
                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow + 2, icol].Value;
                  aSOPvsDemand.slSOP.AddObject( p^.sweek , TObject(p));

                  for sa := Low(TOEMSOPvsDemandSet_OEM) to High(TOEMSOPvsDemandSet_OEM) do
                  begin
                    splan := ExcelApp.Cells[irow + 4 + Ord(sa), 8].Value;
                    if splan <> CSOEMSOPvsDemand_OEM[sa] then
                    begin
                      raise Exception.Create(sSheet + '   行 ' + IntToStr(irow + 4 + Ord(sa)) + ' 列 8 计划列数值错误，当前是' + splan + '正确值应该是 ' + CSOEMSOPvsDemand_OEM[sa]);
                    end;
                  
                    p := New(PWeekRecord);
                    p^.sweek := slweek[iweek];
                    p^.sqty := ExcelApp.Cells[irow + 4 + Ord(sa), icol].Value;
                    vComment := ExcelApp.Cells[irow + 4 + Ord(sa), icol].Comment;


                    if FindVarData(vComment)^.VDispatch <> nil then
                    begin
                      p^.scomment := vComment.Text;
                    end
                    else
                    begin
                      p^.scomment := '';
                    end;
                    aSOPvsDemand.FReasons_OEM[sa].AddObject( p^.sweek,  TObject(p)  );
                  end;
                end;
                    
                irow := irow + 4 + Length(aSOPvsDemand.FReasons_OEM);  
                sweek := ExcelApp.Cells[irow, 3].Value;

              end;

              Memo1.Lines.Add('读分析结果111耗时： ' + IntToStr(GetTickCount - dwTick));
            end;



                  
            if sSheet = 'KPI分析-实际产出 VS S&OP供应计划' then
            begin
              slweek.Clear;         
               
              /////  取 week 数量
              icol := 9;
              sweek := ExcelApp.Cells[1, icol].Value;
              while sweek <> '' do
              begin
                slweek.Add(sweek);
                icol := icol + 1;
                sweek := ExcelApp.Cells[1, icol].Value;
              end;

              irow := 2;
              sweek := ExcelApp.Cells[irow, 3].Value;
              while sweek <> '' do
              begin
                 
                aACTvsDemand := TACTvsDemand.Create;
                aACTvsDemand.smode := ExcelApp.Cells[irow, 1].Value;
                aACTvsDemand.sproj := ExcelApp.Cells[irow, 2].Value;
                aACTvsDemand.sweek := ExcelApp.Cells[irow, 3].Value;
                aACTvsDemand.snumber := ExcelApp.Cells[irow, 4].Value;
                aACTvsDemand.scolor := ExcelApp.Cells[irow, 5].Value;
                aACTvsDemand.scap := ExcelApp.Cells[irow, 6].Value;
                aACTvsDemand.sver := ExcelApp.Cells[irow, 7].Value;
                aACTvsDemand.splan := ExcelApp.Cells[irow, 8].Value;
                slACTvsDemand.AddObject(aACTvsDemand.sproj + aACTvsDemand.snumber, aACTvsDemand);

                for iweek := 0 to slweek.Count - 1 do
                begin
                  icol := iweek + 9;

                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow, icol].Value;
                  aACTvsDemand.slDemand.AddObject( p^.sweek, TObject(p) );

                
                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow + 1, icol].Value;
                  aACTvsDemand.slACT.AddObject( p^.sweek, TObject(p) );

                  for sb := Low(TOEMACTvsDemandSet_OEM) to High(TOEMACTvsDemandSet_OEM) do
                  begin
                    splan := ExcelApp.Cells[irow + 3 + Ord(sb), 8].Value;
                    if splan <> CSOEMACTvsDemand_OEM[sb] then
                    begin
                      raise Exception.Create(sSheet + ' 行 ' + IntToStr(irow + 3 + Ord(sb)) + ' 列 8 计划列数值错误，当前是' + splan + '正确值应该是 ' + CSOEMACTvsDemand_OEM[sb]);
                    end;

                    p := New(PWeekRecord);
                    p^.sweek := slweek[iweek];
                    p^.sqty := ExcelApp.Cells[irow + 3 + Ord(sb), icol].Value;      
                    vComment := ExcelApp.Cells[irow + 3 + Ord(sb), icol].Comment;


                    if FindVarData(vComment)^.VDispatch <> nil then
                    begin
                      p^.scomment := vComment.Text;
                    end
                    else
                    begin
                      p^.scomment := '';
                    end;
                    aACTvsDemand.FReasons_OEM[sb].AddObject( p^.sweek, TObject(p) );
                  end;
                end;


                irow := irow + 3 + Length(aACTvsDemand.FReasons_OEM);
                sweek := ExcelApp.Cells[irow, 3].Value;
              end;
                   

              Memo1.Lines.Add('读分析结果222耗时： ' + IntToStr(GetTickCount - dwTick));
            end;



             
                  
            if sSheet = 'KPI分析-实际产出 VS 排产计划' then
            begin
              slweek.Clear;
                   

              /////  取 week 数量
              icol := 9;
              sweek := ExcelApp.Cells[1, icol].Value;
              while sweek <> '' do
              begin
                slweek.Add(sweek);
                icol := icol + 1;
                sweek := ExcelApp.Cells[1, icol].Value;
              end;
            
              irow := 2;
              sweek := ExcelApp.Cells[irow, 3].Value;
              while sweek <> '' do
              begin
                   
                aACTvsSch := TACTvsSch.Create;
                aACTvsSch.smode := ExcelApp.Cells[irow, 1].Value;
                aACTvsSch.sproj := ExcelApp.Cells[irow, 2].Value;
                aACTvsSch.sweek := ExcelApp.Cells[irow, 3].Value;
                aACTvsSch.snumber := ExcelApp.Cells[irow, 4].Value;
                aACTvsSch.scolor := ExcelApp.Cells[irow, 5].Value;
                aACTvsSch.scap := ExcelApp.Cells[irow, 6].Value;
                aACTvsSch.sver := ExcelApp.Cells[irow, 7].Value;
                aACTvsSch.splan := ExcelApp.Cells[irow, 8].Value;
                slACTvsSch.AddObject(aACTvsSch.sproj + aACTvsSch.snumber, aACTvsSch);

                for iweek := 0 to slweek.Count - 1 do
                begin
                  icol := iweek + 9;
                
                
                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow, icol].Value;
                  aACTvsSch.slSch.AddObject( p^.sweek , TObject(p));
                
                
                  p := New(PWeekRecord);
                  p^.sweek := slweek[iweek];
                  p^.sqty := ExcelApp.Cells[irow + 1, icol].Value;
                  aACTvsSch.slACT.AddObject( p^.sweek , TObject(p));

                  for sc := Low(TOEMACTvsSchSet_OEM) to High(TOEMACTvsSchSet_OEM) do
                  begin                 
                    splan := ExcelApp.Cells[irow + 3 + Ord(sc), 8].Value;
                    if splan <> CSOEMACTvsSch_OEM[sc] then
                    begin                   
                      raise Exception.Create(sSheet + ' 行 ' + IntToStr(irow + 3 + Ord(sc)) + ' 列 8 计划列数值错误，当前是' + splan + '正确值应该是 ' + CSOEMACTvsSch_OEM[sc]);
                    end;

                    p := New(PWeekRecord);
                    p^.sweek := slweek[iweek];
                    p^.sqty := ExcelApp.Cells[irow + 3 + Ord(sc), icol].Value;     
                    vComment := ExcelApp.Cells[irow + 3 + Ord(sc), icol].Comment;


                    if FindVarData(vComment)^.VDispatch <> nil then
                    begin
                      p^.scomment := vComment.Text;
                    end
                    else
                    begin
                      p^.scomment := '';
                    end;
                    aACTvsSch.FReasons_OEM[sc].AddObject(p^.sweek, TObject(p) );
                  end;
                end;

            
                irow := irow + 3 + Length(aACTvsSch.FReasons_OEM);
                sweek := ExcelApp.Cells[irow, 3].Value;    

                //Memo1.Lines.Add('读分析结果333耗时： ' + IntToStr(GetTickCount - dwTick));
              end;


            end;



          
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
    while ExcelApp.Sheets.Count < 3 do
    begin
      ExcelApp.Sheets.Add;
    end;



    /////////////////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////////////
    ProgressBar1.Max := slPlan.Count * 3;
    ProgressBar1.Position := 0;

    ExcelApp.Sheets[1].Activate;
    ExcelApp.Sheets[1].Name := 'KPI分析-S&OP供应计划 VS 销售计划';
    
    ExcelApp.Cells[1, 1].Value := '模式';
    ExcelApp.Cells[1, 2].Value := '项目';
    ExcelApp.Cells[1, 3].Value := 'week';
    ExcelApp.Cells[1, 4].Value := '物料编码';
    ExcelApp.Cells[1, 5].Value := '颜色';
    ExcelApp.Cells[1, 6].Value := '容量';
    ExcelApp.Cells[1, 7].Value := '制式';
    ExcelApp.Cells[1, 8].Value := '计划';
                                             
    ExcelApp.Columns[4].ColumnWidth := 16;
    ExcelApp.Columns[7].ColumnWidth := 12;  
    ExcelApp.Columns[8].ColumnWidth := 25;

    if slPlan.Count > 0 then
    begin
      aPlan := TPlan(slPlan.Objects[0]);
      for iweek := 0 to aPlan.slDemand.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
        ExcelApp.Cells[1, iweek + 9].Value := p^.sweek;
      end;

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].Interior.Color := $DBDCF2;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].HorizontalAlignment := xlCenter;
    end;



    slSOPvsDemand.Sorted := True;
    slACTvsDemand.Sorted := True;
    slACTvsSch.Sorted := True;

    irow := 2;
    for i := 0 to slPlan.Count - 1 do
    begin          
      aPlan := TPlan(slPlan.Objects[i]);
      ExcelApp.Cells[irow, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow, 8].Value := '销售计划';
      
      ExcelApp.Cells[irow + 1, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 1, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 1, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 1, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 1, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 1, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 1, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 1, 8].Value := 'S&OP';

      ExcelApp.Cells[irow + 2, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 2, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 2, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 2, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 2, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 2, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 2, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 2, 8].Value := '期初库存';
                       
      ExcelApp.Cells[irow + 3, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 3, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 3, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 3, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 3, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 3, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 3, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 3, 8].Value := 'S&OP VS 销售计划';

      for iweek := 0 to aPlan.slDemand.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
        ExcelApp.Cells[irow, iweek + 9].Value := p^.sqty;
      end;

      for iweek := 0 to aPlan.slSOP.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slSOP.Objects[iweek]);
        ExcelApp.Cells[irow + 1, iweek + 9].Value := p^.sqty;
        ExcelApp.Cells[irow + 3, iweek + 9].Value := '=' + GetRef(iweek + 9) + IntToStr(irow + 1) + '+' + GetRef(iweek + 9) + IntToStr(irow + 2) + '-' + GetRef(iweek + 9) + IntToStr(irow);
      end;
                
      for iweek := 0 to aPlan.slStk.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slStk.Objects[iweek]);
        ExcelApp.Cells[irow + 2, iweek + 9].Value := p^.sqty;
      end;
        
      ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].Interior.Color := $9DE476;
      ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].NumberFormatLocal := '0_ ';
      ExcelApp.Range[ ExcelApp.Cells[irow + 3, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ ExcelApp.Cells[irow + 3, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].FormatConditions[1].Font.Color := $0000FF;


      idx := slSOPvsDemand.IndexOf(aPlan.sproj + aPlan.snumber);
      if idx >= 0 then
      begin
        aSOPvsDemand := TSOPvsDemand(slSOPvsDemand.Objects[idx]);
      end
      else
      begin
        aSOPvsDemand := nil;
      end;


      for sa := Low(TOEMSOPvsDemandSet_OEM) to High(TOEMSOPvsDemandSet_OEM) do
      begin
     
        ExcelApp.Cells[irow + Ord(sa) + 4, 1].Value := aPlan.smode;
        ExcelApp.Cells[irow + Ord(sa) + 4, 2].Value := aPlan.sproj;
        ExcelApp.Cells[irow + Ord(sa) + 4, 3].Value := aPlan.sweek;
        ExcelApp.Cells[irow + Ord(sa) + 4, 4].Value := aPlan.snumber;
        ExcelApp.Cells[irow + Ord(sa) + 4, 5].Value := aPlan.scolor;
        ExcelApp.Cells[irow + Ord(sa) + 4, 6].Value := aPlan.scap;
        ExcelApp.Cells[irow + Ord(sa) + 4, 7].Value := aPlan.sver;



        ExcelApp.Cells[irow + Ord(sa) + 4, 8].Value := CSOEMSOPvsDemand_OEM[sa];


        if aSOPvsDemand <> nil then
        begin
          for iweek := 0 to aSOPvsDemand.FReasons_OEM[sa].Count - 1 do
          begin
            p := PWeekRecord(aSOPvsDemand.FReasons_OEM[sa].Objects[iweek]);
            ExcelApp.Cells[irow + Ord(sa) + 4, iweek + 9].Value := p^.sqty;
            if p^.scomment <> '' then
            begin
              ExcelApp.Cells[irow + Ord(sa) + 4, iweek + 9].AddComment(p^.scomment);
            end;
          end;
        end;
      end;

      irow := irow + Length(aSOPvsDemand.FReasons_OEM) + 4;

      ProgressBar1.Position := ProgressBar1.Position + 1;    
      Memo1.Lines.Add('11 irow: ' + IntToStr(irow));
    end;

    if slPlan.Count > 0 then
    begin
      aPlan := TPlan(slPlan.Objects[0]);
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, aPlan.slDemand.Count + 8] ].Borders.LineStyle := 1; //加边框
    end;

         

    Memo1.Lines.Add('写写写分析结果111耗时： ' + IntToStr(GetTickCount - dwTick));



    /////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////

    ExcelApp.Sheets[2].Activate;
    ExcelApp.Sheets[2].Name := 'KPI分析-实际产出 VS S&OP供应计划';


    ExcelApp.Cells[1, 1].Value := '模式';
    ExcelApp.Cells[1, 2].Value := '项目';
    ExcelApp.Cells[1, 3].Value := 'week';
    ExcelApp.Cells[1, 4].Value := '物料编码';
    ExcelApp.Cells[1, 5].Value := '颜色';
    ExcelApp.Cells[1, 6].Value := '容量';
    ExcelApp.Cells[1, 7].Value := '制式';
    ExcelApp.Cells[1, 8].Value := '计划';
                                             
    ExcelApp.Columns[4].ColumnWidth := 16;
    ExcelApp.Columns[7].ColumnWidth := 12;  
    ExcelApp.Columns[8].ColumnWidth := 25;

    if slPlan.Count > 0 then
    begin
      aPlan := TPlan(slPlan.Objects[0]);
      for iweek := 0 to aPlan.slDemand.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
        ExcelApp.Cells[1, iweek + 9].Value := p^.sweek;
      end;
      
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].Interior.Color := $DBDCF2;  
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].HorizontalAlignment := xlCenter;
    end;

     

    irow := 2;
    for i := 0 to slPlan.Count - 1 do
    begin          
      aPlan := TPlan(slPlan.Objects[i]);
      ExcelApp.Cells[irow, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow, 8].Value := 'S&OP供应计划';
      
      ExcelApp.Cells[irow + 1, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 1, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 1, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 1, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 1, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 1, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 1, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 1, 8].Value := '实际产出';

      ExcelApp.Cells[irow + 2, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 2, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 2, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 2, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 2, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 2, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 2, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 2, 8].Value := '实际产出 VS S&OP供应计划';

      for iweek := 0 to aPlan.slSOP.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slSOP.Objects[iweek]);
        ExcelApp.Cells[irow, iweek + 9].Value := p^.sqty;
      end;
         
      for iweek := 0 to aPlan.slAct.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slAct.Objects[iweek]);
        ExcelApp.Cells[irow + 1, iweek + 9].Value := p^.sqty;   
        ExcelApp.Cells[irow + 2, iweek + 9].Value := '=' + GetRef(iweek + 9) + IntToStr(irow + 1) + '-' + GetRef(iweek + 9) + IntToStr(irow);
      end;
         
      ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].Interior.Color := $9DE476;   
      ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].NumberFormatLocal := '0_ ';
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions[1].Font.Color := $0000FF;



      idx := slACTvsDemand.IndexOf(aPlan.sproj + aPlan.snumber);
      if idx >= 0 then
      begin
        aACTvsDemand := TACTvsDemand(slACTvsDemand.Objects[idx]);
      end
      else
      begin
        aACTvsDemand := nil;
      end;
      for sb := Low(TOEMACTvsDemandSet_OEM) to High(TOEMACTvsDemandSet_OEM) do
      begin
        ExcelApp.Cells[irow + Ord(sb) + 3, 1].Value := aPlan.smode;
        ExcelApp.Cells[irow + Ord(sb) + 3, 2].Value := aPlan.sproj;
        ExcelApp.Cells[irow + Ord(sb) + 3, 3].Value := aPlan.sweek;
        ExcelApp.Cells[irow + Ord(sb) + 3, 4].Value := aPlan.snumber;
        ExcelApp.Cells[irow + Ord(sb) + 3, 5].Value := aPlan.scolor;
        ExcelApp.Cells[irow + Ord(sb) + 3, 6].Value := aPlan.scap;
        ExcelApp.Cells[irow + Ord(sb) + 3, 7].Value := aPlan.sver;
        ExcelApp.Cells[irow + Ord(sb) + 3, 8].Value := CSOEMACTvsDemand_OEM[sb];

        if aACTvsDemand <> nil then
        begin
          for iweek := 0 to aACTvsDemand.FReasons_OEM[sb].Count - 1 do
          begin
            p := PWeekRecord(aACTvsDemand.FReasons_OEM[sb].Objects[iweek]);
            ExcelApp.Cells[irow + Ord(sb) + 3, iweek + 9].Value := p^.sqty;
            if p^.scomment <> '' then
            begin
              ExcelApp.Cells[irow + Ord(sb) + 3, iweek + 9].AddComment(p^.scomment);
            end;
          end;
        end;                                  
      end;

      irow := irow + Length(aACTvsDemand.FReasons_OEM) + 3;

      ProgressBar1.Position := ProgressBar1.Position + 1;    
      Memo1.Lines.Add('22 irow: ' + IntToStr(irow));
    end;

    if slPlan.Count > 0 then
    begin
      aPlan := TPlan(slPlan.Objects[0]);
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, aPlan.slDemand.Count + 8] ].Borders.LineStyle := 1; //加边框
    end;


                

    Memo1.Lines.Add('写写写分析结果222耗时： ' + IntToStr(GetTickCount - dwTick));



    /////////////////////////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////

    ExcelApp.Sheets[3].Activate;
    ExcelApp.Sheets[3].Name := 'KPI分析-实际产出 VS 排产计划';

 
    ExcelApp.Cells[1, 1].Value := '模式';
    ExcelApp.Cells[1, 2].Value := '项目';
    ExcelApp.Cells[1, 3].Value := 'week';
    ExcelApp.Cells[1, 4].Value := '物料编码';
    ExcelApp.Cells[1, 5].Value := '颜色';
    ExcelApp.Cells[1, 6].Value := '容量';
    ExcelApp.Cells[1, 7].Value := '制式';
    ExcelApp.Cells[1, 8].Value := '计划';
                                             
    ExcelApp.Columns[4].ColumnWidth := 16;
    ExcelApp.Columns[7].ColumnWidth := 12;  
    ExcelApp.Columns[8].ColumnWidth := 25;

    if slPlan.Count > 0 then
    begin
      aPlan := TPlan(slPlan.Objects[0]);
      for iweek := 0 to aPlan.slDemand.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
        ExcelApp.Cells[1, iweek + 9].Value := p^.sweek;
      end;
      
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].Interior.Color := $DBDCF2;  
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].HorizontalAlignment := xlCenter;
    end;

     

    irow := 2;
    for i := 0 to slPlan.Count - 1 do
    begin          
      aPlan := TPlan(slPlan.Objects[i]);
      ExcelApp.Cells[irow, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow, 8].Value := '排产计划';
      
      ExcelApp.Cells[irow + 1, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 1, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 1, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 1, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 1, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 1, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 1, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 1, 8].Value := '实际产出';

      ExcelApp.Cells[irow + 2, 1].Value := aPlan.smode;
      ExcelApp.Cells[irow + 2, 2].Value := aPlan.sproj;
      ExcelApp.Cells[irow + 2, 3].Value := aPlan.sweek;
      ExcelApp.Cells[irow + 2, 4].Value := aPlan.snumber;
      ExcelApp.Cells[irow + 2, 5].Value := aPlan.scolor;
      ExcelApp.Cells[irow + 2, 6].Value := aPlan.scap;
      ExcelApp.Cells[irow + 2, 7].Value := aPlan.sver;
      ExcelApp.Cells[irow + 2, 8].Value := '实际产出 VS 排产计划';

      for iweek := 0 to aPlan.slSch.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slSch.Objects[iweek]);
        ExcelApp.Cells[irow, iweek + 9].Value := p^.sqty;
      end;
         
      for iweek := 0 to aPlan.slAct.Count - 1 do
      begin
        p := PWeekRecord(aPlan.slAct.Objects[iweek]);
        ExcelApp.Cells[irow + 1, iweek + 9].Value := p^.sqty;   
        ExcelApp.Cells[irow + 2, iweek + 9].Value := '=' + GetRef(iweek + 9) + IntToStr(irow + 1) + '-' + GetRef(iweek + 9) + IntToStr(irow);
      end;
         
      ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].Interior.Color := $9DE476;      
      ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].NumberFormatLocal := '0_ ';
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions[1].Font.Color := $0000FF;



      idx := slACTvsSch.IndexOf(aPlan.sproj + aPlan.snumber);
      if idx >= 0 then
      begin
        aACTvsSch := TACTvsSch(slACTvsSch.Objects[idx]);
      end
      else
      begin
        aACTvsSch := nil;
      end;
      for sc := Low(TOEMACTvsSchSet_OEM) to High(TOEMACTvsSchSet_OEM) do
      begin               
        ExcelApp.Cells[irow + Ord(sc) + 3, 1].Value := aPlan.smode;
        ExcelApp.Cells[irow + Ord(sc) + 3, 2].Value := aPlan.sproj;
        ExcelApp.Cells[irow + Ord(sc) + 3, 3].Value := aPlan.sweek;
        ExcelApp.Cells[irow + Ord(sc) + 3, 4].Value := aPlan.snumber;
        ExcelApp.Cells[irow + Ord(sc) + 3, 5].Value := aPlan.scolor;
        ExcelApp.Cells[irow + Ord(sc) + 3, 6].Value := aPlan.scap;
        ExcelApp.Cells[irow + Ord(sc) + 3, 7].Value := aPlan.sver;
        ExcelApp.Cells[irow + Ord(sc) + 3, 8].Value := CSOEMACTvsSch_OEM[sc];

        if aACTvsSch <> nil then
        begin
          for iweek := 0 to aACTvsSch.FReasons_OEM[sc].Count - 1 do
          begin
            p := PWeekRecord(aACTvsSch.FReasons_OEM[sc].Objects[iweek]);
            ExcelApp.Cells[irow + Ord(sc) + 3, iweek + 9].Value := p^.sqty;
            if p^.scomment <> '' then
            begin
              ExcelApp.Cells[irow + Ord(sc) + 3, iweek + 9].AddComment(p^.scomment);
            end;
          end;
        end;                                  
      end;

      irow := irow + Length(aACTvsSch.FReasons_OEM) + 3;

      ProgressBar1.Position := ProgressBar1.Position + 1;
      Memo1.Lines.Add('33 irow: ' + IntToStr(irow));
    end;
            


    Memo1.Lines.Add('写写写分析结果333耗时： ' + IntToStr(GetTickCount - dwTick));


    if slPlan.Count > 0 then
    begin
      aPlan := TPlan(slPlan.Objects[0]);
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, aPlan.slDemand.Count + 8] ].Borders.LineStyle := 1; //加边框
    end;


             


              
    ExcelApp.Sheets[1].Activate;
    
    try
      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

 
        
    Memo1.Lines.Add('完成耗时： ' + IntToStr(GetTickCount - dwTick));




  finally

    for i := 0 to slPlan.Count - 1 do
    begin
      aPlan := TPlan(slPlan.Objects[i]);
      aPlan.Free;
    end;
    slPlan.Free;


    for i := 0 to slSOPvsDemand.Count - 1 do
    begin
      aSOPvsDemand := TSOPvsDemand(slSOPvsDemand.Objects[i]);
      aSOPvsDemand.Free;
    end;
    slSOPvsDemand.Free;
             
    for i := 0 to slACTvsDemand.Count - 1 do
    begin
      aACTvsDemand := TACTvsDemand(slACTvsDemand.Objects[i]);
      aACTvsDemand.Free;
    end;
    slACTvsDemand.Free;

    for i := 0 to slACTvsSch.Count - 1 do
    begin
      aACTvsSch := TACTvsSch(slACTvsSch.Objects[i]);
      aACTvsSch.Free;
    end;
    slACTvsSch.Free;
 
    slweek.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;
 
procedure TfrmMergePlansAnalysis2.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try                                    
    leWeek.Text := ini.ReadString(Self.Name, leWeek.Name, '');  
    leSUM.Text := ini.ReadString(Self.Name, leSUM.Name, '');   
    leAnalysis.Text := ini.ReadString(Self.Name, leAnalysis.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmMergePlansAnalysis2.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try                                                     
    ini.WriteString(Self.Name, leWeek.Name, leWeek.Text);
    ini.WriteString(Self.Name, leSUM.Name, leSUM.Text);
    ini.WriteString(Self.Name, leAnalysis.Name, leAnalysis.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmMergePlansAnalysis2.tbODMClick(Sender: TObject);
var
  ExcelApp, WorkBook: Variant;
  iSheet, iSheetCount: Integer;
  sSheet: string;
  irow: Integer;
  irow1, irow2: Integer; 
  sweek, sweek0: string;
  slweek: TStringList;
  iweek: Integer;
  icol: Integer;

  aPlan: TPlan;
  aSOPvsDemand: TSOPvsDemand;
  aACTvsDemand: TACTvsDemand;
  aACTvsSch: TACTvsSch;

  slPlan: TStringList;
  slSOPvsDemand: TStringList;
  slACTvsDemand: TStringList;
  slACTvsSch: TStringList;

  sa: TOEMSOPvsDemandSet_ODM;
  sb: TOEMACTvsDemandSet_ODM;
  sc: TOEMACTvsSchSet_ODM;

  i: Integer;   
  p: PWeekRecord;
  vComment: Variant;
  
  idx: Integer;

  sfile: string;
  dwTick: DWORD; 

  va: Variant;
  splan, splan1, splan2, splan3, splan4, splan5, splan6: string;
  serr: string;     
  
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;

  sl: TStringList;
  sline: string;
begin

  if not ExcelSaveDialog(sfile) then Exit;

  Memo1.Lines.Add('----------------------------------------------------------');
  for sa := Low(TOEMSOPvsDemandSet_ODM) to High(TOEMSOPvsDemandSet_ODM) do
  begin             
    Memo1.Lines.Add(CSOEMSOPvsDemand_ODM[sa]);
  end;

  Memo1.Lines.Add('----------------------------------------------------------');     
  for sb := Low(TOEMACTvsDemandSet_ODM) to High(TOEMACTvsDemandSet_ODM) do
  begin
    Memo1.Lines.Add(CSOEMACTvsDemand_ODM[sb]);
  end;

  Memo1.Lines.Add('----------------------------------------------------------');
  for sc := Low(TOEMACTvsSchSet_ODM) to High(TOEMACTvsSchSet_ODM) do
  begin
    Memo1.Lines.Add(CSOEMACTvsSch_ODM[sc]);
  end;

  dwTick := GetTickCount;

  slPlan := TStringList.Create;
  slSOPvsDemand := TStringList.Create;
  slACTvsDemand := TStringList.Create;
  slACTvsSch := TStringList.Create;

  slweek := TStringList.Create;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + leSUM.Text + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'集成汇总'+'$]';

    ADOTabXLS.Active:=true;
                     
    /////  取 week 数量
    for icol := 9 to ADOTabXLS.FieldDefList.Count - 1 do
    begin
      sweek := ADOTabXLS.FieldDefList[icol].Name;
      if UpperCase(Copy(sweek, 1, 1)) = 'F' then Break;
      slweek.AddObject(sweek, TObject(icol));
    end;
 
    irow := 2;
    ADOTabXLS.First;

    irow := irow + 1;
    ADOTabXLS.Next;  // 跳过第二行
    
    while not ADOTabXLS.Eof do
    begin                     
      sweek := ADOTabXLS.Fields[2].AsString;
      if sweek <> leWeek.Text then
      begin
        irow := irow + 1;
        ADOTabXLS.Next;
        Continue;
      end;

      aPlan := TPlan.Create;
                      
      aPlan.smode := ADOTabXLS.Fields[0].AsString;
      aPlan.sproj := ADOTabXLS.Fields[1].AsString;
      aPlan.sweek := sweek;

      aPlan.snumber := ADOTabXLS.Fields[4].AsString;
      aPlan.scolor := ADOTabXLS.Fields[5].AsString;
      aPlan.scap := ADOTabXLS.Fields[6].AsString;
      aPlan.sver := ADOTabXLS.Fields[7].AsString;
      aPlan.splan := ADOTabXLS.Fields[8].AsString;
      slPlan.AddObject(aPlan.snumber, aPlan);

      splan1 := ADOTabXLS.Fields[8].AsString; 
      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer( slweek.Objects[iweek] ) ].AsString;
        aPlan.slDemand.AddObject( p^.sweek, TObject(p) );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      splan2 := ADOTabXLS.Fields[8].AsString; 
      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer( slweek.Objects[iweek] ) ].AsString;
        aPlan.slSOP.AddObject( p^.sweek, TObject(p)  );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      splan3 := ADOTabXLS.Fields[8].AsString; 
      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer( slweek.Objects[iweek] ) ].AsString;
        aPlan.slMPS.AddObject( p^.sweek, TObject(p)  );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      splan4 := ADOTabXLS.Fields[8].AsString; 
      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer( slweek.Objects[iweek] ) ].AsString;
        aPlan.slSch.AddObject( p^.sweek, TObject(p)  );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      splan5 := ADOTabXLS.Fields[8].AsString;
      for iweek := 0 to slweek.Count - 1 do
      begin
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer( slweek.Objects[iweek] ) ].AsString;
        aPlan.slAct.AddObject( p^.sweek, TObject(p)  );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      splan6 := ADOTabXLS.Fields[8].AsString;
      for iweek := 0 to slweek.Count - 1 do
      begin
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer( slweek.Objects[iweek] ) ].AsString;
        aPlan.slStk.AddObject( p^.sweek, TObject(p)  );
      end;


      splan := splan1 + splan2 + splan3 + splan4 + splan5 + splan6;
      if splan <> '销售计划S&OP供应计划MPS排产计划实际产出期初库存' then
      begin
        serr := '第 ' + IntToStr(irow - 5) + ' 行数据格式不对( 销售计划S&OP供应计划MPS排产计划实际产出期初库存 )';
        Memo1.Lines.Add(serr);
        raise Exception.Create(serr);
      end;


                  

      irow := irow + 1;
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;

          
  Memo1.Lines.Add('读数据耗时： ' + IntToStr(GetTickCount - dwTick));





  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + leAnalysis.Text + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'KPI分析-S&OP供应计划 VS 销售计划'+'$]';

    ADOTabXLS.Active:=true;

    slweek.Clear;
                    
    /////  取 week 数量
    for icol := 8 to ADOTabXLS.FieldDefList.Count - 1 do
    begin          
      sweek := ADOTabXLS.FieldDefList[icol].Name;
      slweek.AddObject(sweek, TObject(icol));
    end; 

    ADOTabXLS.First;     
    
    while not ADOTabXLS.Eof do
    begin       
      aSOPvsDemand := TSOPvsDemand.Create;
      aSOPvsDemand.smode := ADOTabXLS.Fields[0].AsString;
      aSOPvsDemand.sproj := ADOTabXLS.Fields[1].AsString;
      aSOPvsDemand.sweek := ADOTabXLS.Fields[2].AsString;
      aSOPvsDemand.snumber := ADOTabXLS.Fields[3].AsString;
      aSOPvsDemand.scolor := ADOTabXLS.Fields[4].AsString;
      aSOPvsDemand.scap := ADOTabXLS.Fields[5].AsString;
      aSOPvsDemand.sver := ADOTabXLS.Fields[6].AsString;
      aSOPvsDemand.splan := ADOTabXLS.Fields[7].AsString;
      slSOPvsDemand.AddObject(aSOPvsDemand.sproj + aSOPvsDemand.snumber, aSOPvsDemand);

      for iweek := 0 to slweek.Count - 1 do
      begin
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aSOPvsDemand.slDemand.AddObject( p^.sweek, TObject(p) );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      for iweek := 0 to slweek.Count - 1 do
      begin
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aSOPvsDemand.slSOP.AddObject( p^.sweek, TObject(p) );
      end;   

      irow := irow + 1;
      ADOTabXLS.Next;
      for iweek := 0 to slweek.Count - 1 do
      begin
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aSOPvsDemand.slStk.AddObject( p^.sweek, TObject(p) );
      end;

      ADOTabXLS.Next;  //跳过公式行

      for sa := Low(TOEMSOPvsDemandSet_ODM) to High(TOEMSOPvsDemandSet_ODM) do
      begin                     
        irow := irow + 1;
        ADOTabXLS.Next;
                  
        splan := ADOTabXLS.Fields[ 7 ].AsString;
        if splan <> CSOEMSOPvsDemand_ODM[sa] then
        begin
          raise Exception.Create(sSheet + '   行 ' + IntToStr(irow + 4 + Ord(sa)) + ' 列 8 计划列数值错误，当前是' + splan + '正确值应该是 ' + CSOEMSOPvsDemand_ODM[sa]);
        end;
                                      
        for iweek := 0 to slweek.Count - 1 do
        begin

          p := New(PWeekRecord);
          p^.sweek := slweek[iweek];
          p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
          p^.scomment := '';
          aSOPvsDemand.FReasons_ODM[sa].AddObject( p^.sweek,  TObject(p)  );                  
                  
        end;
      end;
                
 
    
      ADOTabXLS.Next;
    end;
        

    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end; 

  Memo1.Lines.Add('读数据耗时： ' + IntToStr(GetTickCount - dwTick));










  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + leAnalysis.Text + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'KPI分析-实际产出 VS S&OP供应计划'+'$]';

    ADOTabXLS.Active:=true;

    slweek.Clear;
                    
    /////  取 week 数量
    for icol := 8 to ADOTabXLS.FieldDefList.Count - 1 do
    begin          
      sweek := ADOTabXLS.FieldDefList[icol].Name;
      slweek.AddObject(sweek, TObject(icol));
    end; 

    irow := 2;
    ADOTabXLS.First;     
    
    while not ADOTabXLS.Eof do
    begin

      aACTvsDemand := TACTvsDemand.Create;
      aACTvsDemand.smode := ADOTabXLS.Fields[ 0 ].AsString;
      aACTvsDemand.sproj := ADOTabXLS.Fields[ 1 ].AsString;
      aACTvsDemand.sweek := ADOTabXLS.Fields[ 2 ].AsString;
      aACTvsDemand.snumber := ADOTabXLS.Fields[ 3 ].AsString;
      aACTvsDemand.scolor := ADOTabXLS.Fields[ 4 ].AsString;
      aACTvsDemand.scap := ADOTabXLS.Fields[ 5 ].AsString;
      aACTvsDemand.sver := ADOTabXLS.Fields[ 6 ].AsString;
      aACTvsDemand.splan := ADOTabXLS.Fields[ 7 ].AsString;
      slACTvsDemand.AddObject(aACTvsDemand.sproj + aACTvsDemand.snumber, aACTvsDemand);

      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aACTvsDemand.slDemand.AddObject( p^.sweek, TObject(p) );
      end;

      irow := irow + 1;
      ADOTabXLS.Next;
      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aACTvsDemand.slACT.AddObject( p^.sweek, TObject(p) );
      end;
                              
      irow := irow + 1;
      ADOTabXLS.Next;  //跳过公式行

      for sb := Low(TOEMACTvsDemandSet_ODM) to High(TOEMACTvsDemandSet_ODM) do
      begin                
        irow := irow + 1;
        ADOTabXLS.Next;
        
        splan := ADOTabXLS.Fields[ 7 ].AsString;
        if splan <> CSOEMACTvsDemand_ODM[sb] then
        begin
          raise Exception.Create(sSheet + ' 行 ' + IntToStr(irow + 3 + Ord(sb)) + ' 列 8 计划列数值错误，当前是' + splan + '正确值应该是 ' + CSOEMACTvsDemand_ODM[sb]);
        end;
                       
        for iweek := 0 to slweek.Count - 1 do
        begin
          p := New(PWeekRecord);
          p^.sweek := slweek[iweek];
          p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
          p^.scomment := '';
          aACTvsDemand.FReasons_ODM[sb].AddObject( p^.sweek, TObject(p) );
        end;     
      end;


      irow := irow + 1;
      ADOTabXLS.Next;
    end;
        

    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end; 

  Memo1.Lines.Add('读数据耗时： ' + IntToStr(GetTickCount - dwTick));









  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + leAnalysis.Text + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'KPI分析-实际产出 VS 排产计划'+'$]';

    ADOTabXLS.Active:=true;

    slweek.Clear;
                    
    /////  取 week 数量
    for icol := 8 to ADOTabXLS.FieldDefList.Count - 1 do
    begin          
      sweek := ADOTabXLS.FieldDefList[icol].Name;
      slweek.AddObject(sweek, TObject(icol));
    end; 

    irow := 2;
    ADOTabXLS.First;     
    
    while not ADOTabXLS.Eof do
    begin

      aACTvsSch := TACTvsSch.Create;
      aACTvsSch.smode := ADOTabXLS.Fields[ 0 ].AsString;
      aACTvsSch.sproj := ADOTabXLS.Fields[ 1 ].AsString;
      aACTvsSch.sweek := ADOTabXLS.Fields[ 2 ].AsString;
      aACTvsSch.snumber := ADOTabXLS.Fields[ 3 ].AsString;
      aACTvsSch.scolor := ADOTabXLS.Fields[ 4 ].AsString;
      aACTvsSch.scap := ADOTabXLS.Fields[ 5 ].AsString;
      aACTvsSch.sver := ADOTabXLS.Fields[ 6 ].AsString;
      aACTvsSch.splan := ADOTabXLS.Fields[ 7 ].AsString;
      slACTvsSch.AddObject(aACTvsSch.sproj + aACTvsSch.snumber, aACTvsSch);

      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aACTvsSch.slSch.AddObject( p^.sweek , TObject(p)); 
      end;

      irow := irow + 1;
      ADOTabXLS.Next; 
      for iweek := 0 to slweek.Count - 1 do
      begin 
        p := New(PWeekRecord);
        p^.sweek := slweek[iweek];
        p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
        aACTvsSch.slACT.AddObject( p^.sweek , TObject(p));

      end;     
                         
      irow := irow + 1;
      ADOTabXLS.Next;  // 跳过公式行


      for sc := Low(TOEMACTvsSchSet_ODM) to High(TOEMACTvsSchSet_ODM) do
      begin   
        irow := irow + 1;
        ADOTabXLS.Next;
                      
        splan := ADOTabXLS.Fields[ 7 ].AsString;
        if splan <> CSOEMACTvsSch_ODM[sc] then
        begin                   
          if (sc = scFacMan_ODM) and (splan <> '代工厂执行力') then
            raise Exception.Create(sSheet + ' 行 ' + IntToStr(irow + 3 + Ord(sc)) + ' 列 8 计划列数值错误，当前是' + splan + '正确值应该是 ' + CSOEMACTvsSch_ODM[sc]);
        end;
                                        
        for iweek := 0 to slweek.Count - 1 do
        begin
          p := New(PWeekRecord);
          p^.sweek := slweek[iweek];
          p^.sqty := ADOTabXLS.Fields[ Integer(slweek.Objects[iweek]) ].AsString;
          p^.scomment := '';
          aACTvsSch.FReasons_ODM[sc].AddObject(p^.sweek, TObject(p) );
        end;
      end;                

                

      irow := irow + 1;
      ADOTabXLS.Next;
    end;
        

    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end; 

  Memo1.Lines.Add('读数据耗时： ' + IntToStr(GetTickCount - dwTick));





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
  while ExcelApp.Sheets.Count < 3 do
  begin
    ExcelApp.Sheets.Add;
  end;

  sl := TStringList.Create;

  /////////////////////////////////////////////////////////////////////////////////////////////
  /////////////////////////////////////////////////////////////////////////////////////////////
  ProgressBar1.Max := slPlan.Count * 3;
  ProgressBar1.Position := 0;

  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := 'KPI分析-S&OP供应计划 VS 销售计划';

  sline := '模式'#9'项目'#9'week'#9'物料编码'#9'颜色'#9'容量'#9'制式'#9'计划';
                                             
  ExcelApp.Columns[4].ColumnWidth := 16;
  ExcelApp.Columns[7].ColumnWidth := 12;  
  ExcelApp.Columns[8].ColumnWidth := 25;

  if slPlan.Count > 0 then
  begin
    aPlan := TPlan(slPlan.Objects[0]);
    for iweek := 0 to aPlan.slDemand.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
      sline := sline + #9 + p^.sweek;
    end;

    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].Interior.Color := $DBDCF2;
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].HorizontalAlignment := xlCenter;
  end;
  sl.Add(sline);



  slSOPvsDemand.Sorted := True;
  slACTvsDemand.Sorted := True;
  slACTvsSch.Sorted := True;

  irow := 2;
  for i := 0 to slPlan.Count - 1 do
  begin          
    aPlan := TPlan(slPlan.Objects[i]);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '销售计划';

    for iweek := 0 to aPlan.slDemand.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
      sline := sline + #9 + p^.sqty;
    end;      
    sl.Add(sline);


    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      'S&OP';
    for iweek := 0 to aPlan.slSOP.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slSOP.Objects[iweek]);
      sline := sline + #9 + p^.sqty;           
    end; 
    sl.Add(sline);


    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '期初库存';

    for iweek := 0 to aPlan.slDemand.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slStk.Objects[iweek]);
      sline := sline + #9 + p^.sqty;
    end;
    sl.Add(sline);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      'S&OP VS 销售计划';
      
    for iweek := 0 to aPlan.slSOP.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slSOP.Objects[iweek]); 
      sline := sline + #9 + '=' + GetRef(iweek + 9) + IntToStr(irow + 1) + '+' + GetRef(iweek + 9) + IntToStr(irow + 2) + '-' + GetRef(iweek + 9) + IntToStr(irow);        
    end;
    sl.Add(sline);

    ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].Interior.Color := $9DE476;
    ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].NumberFormatLocal := '0_ ';
    ExcelApp.Range[ ExcelApp.Cells[irow + 3, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    ExcelApp.Range[ ExcelApp.Cells[irow + 3, 9], ExcelApp.Cells[irow + 3, 9 + aPlan.slDemand.Count - 1] ].FormatConditions[1].Font.Color := $0000FF;

    idx := slSOPvsDemand.IndexOf(aPlan.sproj + aPlan.snumber);
    if idx >= 0 then
    begin
      aSOPvsDemand := TSOPvsDemand(slSOPvsDemand.Objects[idx]);
    end
    else
    begin
      aSOPvsDemand := nil;
    end;    

    for sa := Low(TOEMSOPvsDemandSet_ODM) to High(TOEMSOPvsDemandSet_ODM) do
    begin
      sline := aPlan.smode + #9 +
        aPlan.sproj + #9 +
        aPlan.sweek + #9 +
        aPlan.snumber + #9 +
        aPlan.scolor + #9 +
        aPlan.scap + #9 +
        aPlan.sver + #9 + CSOEMSOPvsDemand_ODM[sa];

      if aSOPvsDemand <> nil then
      begin
        for iweek := 0 to aSOPvsDemand.FReasons_ODM[sa].Count - 1 do
        begin
          p := PWeekRecord(aSOPvsDemand.FReasons_ODM[sa].Objects[iweek]);
          sline := sline + #9 + p^.sqty;
        end;
      end;

      sl.Add(sline);
    end;

    irow := irow + Length(aSOPvsDemand.FReasons_ODM) + 4;

    ProgressBar1.Position := ProgressBar1.Position + 1;    
    Memo1.Lines.Add('11 irow: ' + IntToStr(irow));
  end;

                           
  ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
  Clipboard.SetTextBuf(PChar(sl.Text));
  ExcelApp.ActiveSheet.Paste;      
           

  if slPlan.Count > 0 then
  begin
    aPlan := TPlan(slPlan.Objects[0]);
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[sl.Count, aPlan.slDemand.Count + 8] ].Borders.LineStyle := 1; //加边框
  end;

         

  Memo1.Lines.Add('写写写分析结果111耗时： ' + IntToStr(GetTickCount - dwTick));



  /////////////////////////////////////////////////////////////////////////////////
  /////////////////////////////////////////////////////////////////////////////////    
  /////////////////////////////////////////////////////////////////////////////////
  /////////////////////////////////////////////////////////////////////////////////

  ExcelApp.Sheets[2].Activate;
  ExcelApp.Sheets[2].Name := 'KPI分析-实际产出 VS S&OP供应计划';

  sl.Clear;

  sline := '模式'#9'项目'#9'week'#9'物料编码'#9'颜色'#9'容量'#9'制式'#9'计划';
                                             
  ExcelApp.Columns[4].ColumnWidth := 16;
  ExcelApp.Columns[7].ColumnWidth := 12;  
  ExcelApp.Columns[8].ColumnWidth := 25;

  if slPlan.Count > 0 then
  begin
    aPlan := TPlan(slPlan.Objects[0]);
    for iweek := 0 to aPlan.slDemand.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
      sline := sline + #9 + p^.sweek;
    end;
      
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].Interior.Color := $DBDCF2;  
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].HorizontalAlignment := xlCenter;
  end;
  sl.Add(sline);

     

  irow := 2;
  for i := 0 to slPlan.Count - 1 do
  begin          
    aPlan := TPlan(slPlan.Objects[i]);
    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      'S&OP供应计划';
    for iweek := 0 to aPlan.slSOP.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slSOP.Objects[iweek]);
      sline := sline + #9 + p^.sqty;
    end;
    sl.Add(sline);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '实际产出';

    for iweek := 0 to aPlan.slAct.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slAct.Objects[iweek]);
      sline := sline + #9 + p^.sqty;    
    end;
    sl.Add(sline);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '实际产出 VS S&OP供应计划';

    for iweek := 0 to aPlan.slAct.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slAct.Objects[iweek]);
      sline := sline + #9 + '=' + GetRef(iweek + 9) + IntToStr(irow + 1) + '-' + GetRef(iweek + 9) + IntToStr(irow);
    end;
    sl.Add(sline);
         
    ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].Interior.Color := $9DE476;    
    ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].NumberFormatLocal := '0_ ';
    ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions[1].Font.Color := $0000FF;


    idx := slACTvsDemand.IndexOf(aPlan.sproj + aPlan.snumber);
    if idx >= 0 then
    begin
      aACTvsDemand := TACTvsDemand(slACTvsDemand.Objects[idx]);
    end
    else
    begin
      aACTvsDemand := nil;
    end;
    
    for sb := Low(TOEMACTvsDemandSet_ODM) to High(TOEMACTvsDemandSet_ODM) do
    begin
      sline := aPlan.smode + #9 +
        aPlan.sproj + #9 +
        aPlan.sweek + #9 +
        aPlan.snumber + #9 +
        aPlan.scolor + #9 +
        aPlan.scap + #9 +
        aPlan.sver + #9 +
        CSOEMACTvsDemand_ODM[sb];

      if aACTvsDemand <> nil then
      begin
        for iweek := 0 to aACTvsDemand.FReasons_ODM[sb].Count - 1 do
        begin
          p := PWeekRecord(aACTvsDemand.FReasons_ODM[sb].Objects[iweek]);
          sline := sline + #9 + p^.sqty;
        end;
      end;
      sl.Add(sline);                                  
    end;

    irow := irow + Length(aACTvsDemand.FReasons_ODM) + 3;

    ProgressBar1.Position := ProgressBar1.Position + 1;    
    Memo1.Lines.Add('22 irow: ' + IntToStr(irow));
  end;

  if slPlan.Count > 0 then
  begin
    aPlan := TPlan(slPlan.Objects[0]);
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, aPlan.slDemand.Count + 8] ].Borders.LineStyle := 1; //加边框
  end;

                            
  ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
  Clipboard.SetTextBuf(PChar(sl.Text));
  ExcelApp.ActiveSheet.Paste;      
           

                

  Memo1.Lines.Add('写写写分析结果222耗时： ' + IntToStr(GetTickCount - dwTick));



  /////////////////////////////////////////////////////////////////////////////////
  /////////////////////////////////////////////////////////////////////////////////

  ExcelApp.Sheets[3].Activate;
  ExcelApp.Sheets[3].Name := 'KPI分析-实际产出 VS 排产计划';

          
  sl.Clear;

  sline := '模式'#9'项目'#9'week'#9'物料编码'#9'颜色'#9'容量'#9'制式'#9'计划';
                                             
  ExcelApp.Columns[4].ColumnWidth := 16;
  ExcelApp.Columns[7].ColumnWidth := 12;  
  ExcelApp.Columns[8].ColumnWidth := 25;

  if slPlan.Count > 0 then
  begin
    aPlan := TPlan(slPlan.Objects[0]);
    for iweek := 0 to aPlan.slDemand.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slDemand.Objects[iweek]);
      sline := sline + #9 + p^.sweek;
    end;
      
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].Interior.Color := $DBDCF2;  
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, aPlan.slDemand.Count + 8] ].HorizontalAlignment := xlCenter;
  end;
  sl.Add(sline);

     

  irow := 2;
  for i := 0 to slPlan.Count - 1 do
  begin          
    aPlan := TPlan(slPlan.Objects[i]);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '排产计划';

    for iweek := 0 to aPlan.slSch.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slSch.Objects[iweek]);
      sline := sline + #9 + p^.sqty;
    end;

    sl.Add(sline);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '实际产出';
      
    for iweek := 0 to aPlan.slAct.Count - 1 do
    begin
      p := PWeekRecord(aPlan.slAct.Objects[iweek]);
      sline := sline + #9 + p^.sqty;
    end;
    sl.Add(sline);

    sline := aPlan.smode + #9 +
      aPlan.sproj + #9 +
      aPlan.sweek + #9 +
      aPlan.snumber + #9 +
      aPlan.scolor + #9 +
      aPlan.scap + #9 +
      aPlan.sver + #9 +
      '实际产出 VS 排产计划'; 
         
    for iweek := 0 to aPlan.slAct.Count - 1 do
    begin 
      sline := sline + #9 + '=' + GetRef(iweek + 9) + IntToStr(irow + 1) + '-' + GetRef(iweek + 9) + IntToStr(irow);
    end;
    sl.Add(sline);
         
    ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].Interior.Color := $9DE476;       
    ExcelApp.Range[ ExcelApp.Cells[irow, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].NumberFormatLocal := '0_ ';
    ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    ExcelApp.Range[ ExcelApp.Cells[irow + 2, 9], ExcelApp.Cells[irow + 2, 9 + aPlan.slDemand.Count - 1] ].FormatConditions[1].Font.Color := $0000FF;



    idx := slACTvsSch.IndexOf(aPlan.sproj + aPlan.snumber);
    if idx >= 0 then
    begin
      aACTvsSch := TACTvsSch(slACTvsSch.Objects[idx]);
    end
    else
    begin
      aACTvsSch := nil;
    end;
    
    for sc := Low(TOEMACTvsSchSet_ODM) to High(TOEMACTvsSchSet_ODM) do
    begin
      sline := aPlan.smode + #9 +
        aPlan.sproj + #9 +
        aPlan.sweek + #9 +
        aPlan.snumber + #9 +
        aPlan.scolor + #9 +
        aPlan.scap + #9 +
        aPlan.sver + #9 +
        CSOEMACTvsSch_ODM[sc];

      if aACTvsSch <> nil then
      begin
        for iweek := 0 to aACTvsSch.FReasons_ODM[sc].Count - 1 do
        begin
          p := PWeekRecord(aACTvsSch.FReasons_ODM[sc].Objects[iweek]);
          sline := sline + #9 + p^.sqty;
        end;
      end;
      
      sl.Add(sline);                                  
    end;

    irow := irow + Length(aACTvsSch.FReasons_ODM) + 3;   

    ProgressBar1.Position := ProgressBar1.Position + 1;
    Memo1.Lines.Add('33 irow: ' + IntToStr(irow));
  end;
         
                            
  ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
  Clipboard.SetTextBuf(PChar(sl.Text));
  ExcelApp.ActiveSheet.Paste;      
           
    


  Memo1.Lines.Add('写写写分析结果333耗时： ' + IntToStr(GetTickCount - dwTick));


  if slPlan.Count > 0 then
  begin
    aPlan := TPlan(slPlan.Objects[0]);
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, aPlan.slDemand.Count + 8] ].Borders.LineStyle := 1; //加边框
  end;


             


              
  ExcelApp.Sheets[1].Activate;
    
  try
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;



  sl.Free;
 
        
  Memo1.Lines.Add('完成耗时： ' + IntToStr(GetTickCount - dwTick));





  for i := 0 to slPlan.Count - 1 do
  begin
    aPlan := TPlan(slPlan.Objects[i]);
    aPlan.Free;
  end;
  slPlan.Free;


  for i := 0 to slSOPvsDemand.Count - 1 do
  begin
    aSOPvsDemand := TSOPvsDemand(slSOPvsDemand.Objects[i]);
    aSOPvsDemand.Free;
  end;
  slSOPvsDemand.Free;
             
  for i := 0 to slACTvsDemand.Count - 1 do
  begin
    aACTvsDemand := TACTvsDemand(slACTvsDemand.Objects[i]);
    aACTvsDemand.Free;
  end;
  slACTvsDemand.Free;

  for i := 0 to slACTvsSch.Count - 1 do
  begin
    aACTvsSch := TACTvsSch(slACTvsSch.Objects[i]);
    aACTvsSch.Free;
  end;
  slACTvsSch.Free;
 
  slweek.Free;


  MessageBox(Handle, '完成', '提示', 0);
end;

end.
