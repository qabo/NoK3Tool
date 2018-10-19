unit SAPBomReader;

interface

uses
  Classes, ComObj, ActiveX, SysUtils, Windows, CommUtils, SAPStockReader,
  SAPMaterialReader;

type
  TSapBom = class;
  
  TSapItemGroup = class
  private
    function GetItemCount: Integer;
    function GetItems(i: Integer): TSapBom;
  public
    FGroupNo: string;
    FList: TStringList;

    FParent: TSapBom;

    constructor Create(const sgroup: string);
    destructor Destroy; override;
    procedure Clear;
    property ItemCount: Integer read GetItemCount;
    property Items[i: Integer]: TSapBom read GetItems;

    procedure Save(ExcelApp: Variant; var irow: Integer); 
    procedure SaveSBom(ExcelApp: Variant; var irow: Integer; root: TSapBom;
      dUsage: Double; lt: Double; bWriteHeader: Boolean; dStock: Double;
      aSAPStockReader: TSAPStockReader);
  end;

  TSapBom = class
  private
    function GetGroups(i: Integer): TSapItemGroup;
    function GetChildCount: Integer;
  public
    FNumber: string;
    FName: string;
    FList: TList;
    FACount: Integer;

    slevel: string;
    sptype: string;
    abc: string;
    sfac: string;
    lt: Double;
    sgroup: string;
    spriority: string;
    dusage: Double;
    dPer: Double;
    iLowestCode: Integer;

    FParent: TSapItemGroup;
 
    FStock: Double;

    constructor Create(const snumber: string);
    destructor Destroy; override;
    procedure Clear;
    property Childs[i: Integer]: TSapItemGroup read GetGroups;
    property ChildCount: Integer read GetChildCount;

    procedure Save(ExcelApp: Variant; var irow: Integer);
    procedure SaveSBom(ExcelApp: Variant; var irow: Integer; root: TSapBom;
      dUsage: Double; lt: Double; dStock: Double; aSAPStockReader: TSAPStockReader);
    function GetSapBom(const snumber: string): TSapBom;

    procedure GetLowestCode(aSAPMaterialRecordPtr: PSAPMaterialRecord);
    function ChildByGroup(const sgroup: string): TSapItemGroup;
    function GetPlanNumber: string;
  end;

  TSAPBomReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    
    procedure Open;
    procedure Log(const str: string);
  public
    FList: TStringList;
    FNumbers: TStringList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    procedure Save(const sfile: string);
    procedure SaveSBom(const sfile: string; aSAPStockReader: TSAPStockReader);
    function GetSapBom(const snumber: string): TSapBom;
    procedure GetLowestCode(aSAPMaterialRecordPtr: PSAPMaterialRecord);
  end;

implementation

{ TSapItemGroup }

constructor TSapItemGroup.Create(const sgroup: string);
begin
  FGroupNo := sgroup;
  FList := TStringList.Create;
  FParent := nil;
end;
destructor TSapItemGroup.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSapItemGroup.Clear;
var
  i: Integer;
  aSapBom: TSapBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSapBom := TSapBom(FList.Objects[i]);
    aSapBom.Free;
  end;
  FList.Clear;
end;

function TSapItemGroup.GetItemCount: Integer;
begin
  Result := FList.Count;
end;

function TSapItemGroup.GetItems(i: Integer): TSapBom;
begin
  Result := TSapBom(FList.Objects[i]);
end;

procedure TSapItemGroup.Save(ExcelApp: Variant; var irow: Integer);
var
  i: Integer;
begin
  for i := 0 to self.ItemCount - 1 do
  begin
    self.Items[i].Save(ExcelApp, irow);
  end;
end;

procedure TSapItemGroup.SaveSBom(ExcelApp: Variant; var irow: Integer;
  root: TSapBom; dUsage: Double; lt: Double; bWriteHeader: Boolean; dStock: Double;
  aSAPStockReader: TSAPStockReader);
var
  i: Integer;
  aSapBom: TSapBom;
  dParentStock: Double;
begin
  for i := 0 to self.ItemCount - 1 do
  begin
    aSapBom := Items[i];
    if aSapBom.sgroup = '' then
    begin      
      dParentStock := dStock * aSapBom.dusage;
    end
    else
    begin   
      dParentStock := (dStock * aSapBom.dPer / 100) * aSapBom.dusage;
    end;
    aSapBom.SaveSBom(ExcelApp, irow, root, dUsage, lt, dParentStock, aSAPStockReader);
  end;
end;  

{ TSapBom }

constructor TSapBom.Create(const snumber: string);
begin
  FNumber := snumber;
  FACount := 0; 
  FList := TList.Create;
  FParent := nil;
  dusage := 1;
end;

destructor TSapBom.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSapBom.Clear;
var
  i: Integer;
  aSapItemGroup: TSapItemGroup;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSapItemGroup := TSapItemGroup(FList[i]);
    aSapItemGroup.Free;
  end;
  FList.Clear;
end;

function TSapBom.GetSapBom(const snumber: string): TSapBom;
var
  i: Integer;
  aSapItemGroup: TSapItemGroup;
  iChildItem: Integer;
  aSapBom: TSapBom;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSapItemGroup := TSapItemGroup(FList[i]);
    for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
    begin
      aSapBom := aSapItemGroup.Items[iChildItem];
      if aSapBom.FNumber = snumber then
      begin
        Result := aSapBom;
        Exit;
      end;
    end;
  end;
end;

procedure TSapBom.GetLowestCode(aSAPMaterialRecordPtr: PSAPMaterialRecord);
var
  i: Integer;
  aSapItemGroup: TSapItemGroup;     
  iChildItem: Integer;
  aSapBom: TSapBom;
begin
  if Self.FNumber = aSAPMaterialRecordPtr^.sNumber then
  begin
    if aSAPMaterialRecordPtr^.iLowestCode < self.iLowestCode then
    begin
      aSAPMaterialRecordPtr^.iLowestCode := self.iLowestCode;
    end;
  end
  else
  begin
    for i := 0 to FList.Count - 1 do
    begin
      aSapItemGroup := TSapItemGroup(FList[i]);
      for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
      begin
        aSapBom := aSapItemGroup.Items[iChildItem];
        aSapBom.GetLowestCode(aSAPMaterialRecordPtr);
      end;
    end;
  end;
end;  

function TSapBom.ChildByGroup(const sgroup: string): TSapItemGroup;
var
  i: Integer;
begin
  Result := nil;
  if sgroup = '' then Exit;
  for i := 0 to self.ChildCount - 1 do
  begin
    if Childs[i].FGroupNo = sgroup then
    begin
      Result := Childs[i];
      Break;
    end;
  end;
end;

function TSapBom.GetPlanNumber: string;
var
  i, j: Integer;
  aSapItemGroup: TSapItemGroup;
  aSapBom: TSapBom;
begin
  Result := '';
  for i := 0 to self.ChildCount - 1 do
  begin
    aSapItemGroup := Childs[i];
    for j := 0 to aSapItemGroup.ItemCount - 1 do
    begin
      aSapBom := aSapItemGroup.Items[j];
      if Pos('裸机', aSapBom.FName) > 0 then
      begin
        Result := aSapBom.FNumber;
        Result := '90' + Copy(Result, 3, Length(Result));
        Break;
      end;
    end;
    if Result <> '' then Break; 
  end;
end;

function TSapBom.GetGroups(i: Integer): TSapItemGroup;
begin
  Result := TSapItemGroup(FList[i]);
end;

function TSapBom.GetChildCount: Integer;
begin
  Result := FList.Count;
end;

procedure TSapBom.Save(ExcelApp: Variant; var irow: Integer);
var
  i: Integer;
begin
  if FParent <> nil then   // 不是根节点，写自制件自身
  begin
    ExcelApp.Cells[irow, 4].Value := self.FNumber;
    ExcelApp.Cells[irow, 5].Value := self.FName;
    ExcelApp.Cells[irow, 6].Value := '''' + self.slevel;
    ExcelApp.Cells[irow, 7].Value := self.sptype;
    ExcelApp.Cells[irow, 8].Value := self.abc;
    ExcelApp.Cells[irow, 9].Value := self.sgroup;
    ExcelApp.Cells[irow, 10].Value := self.lt;
    irow := irow + 1;
  end;

  for i := 0 to Self.ChildCount - 1 do
  begin
    Self.Childs[i].Save(ExcelApp, irow);
  end;

end;

(*
dStock: 上级库存
*)
procedure TSapBom.SaveSBom(ExcelApp: Variant; var irow: Integer; root: TSapBom;
  dUsage: Double; lt: Double; dStock: Double; aSAPStockReader: TSAPStockReader);
var
  i: Integer;
  dStockParent: Double;
begin
  if ChildCount = 0 then // 是叶子节点
  begin
    if UpperCase( Self.abc ) <> 'A' then Exit; // 不是关键物料，不写
    ExcelApp.Cells[irow, 3].Value := self.FNumber;
    ExcelApp.Cells[irow, 4].Value := self.FName;
    ExcelApp.Cells[irow, 5].Value := lt; 
    if self.sgroup <> '' then
    begin
      ExcelApp.Cells[irow, 6].Value := Self.FParent.FParent.FNumber + '-' + self.sgroup;
    end;
    ExcelApp.Cells[irow, 7].Value := dStock;
    ExcelApp.Cells[irow, 8].Value := aSAPStockReader.AllocStockSum(FNumber);
    ExcelApp.Cells[irow, 9].Value := self.dusage * dUsage;
    ExcelApp.Cells[irow, 10].Value := self.spriority;
    irow := irow + 1;
  end
  else
  begin
    for i := 0 to ChildCount - 1 do
    begin
      if FParent = nil then // 根节点
      begin
        dStockParent := 0;
        FStock := 0;
      end
      else
      begin
        if self.sgroup <> '' then
        begin
          dStockParent := dStock * dusage * dPer / 100;
        end
        else
        begin
          dStockParent := dStock * dusage;
        end;
        FStock := aSAPStockReader.AllocStockSum(FNumber);
      end;
                  
      if self.sgroup <> '' then
      begin
        Childs[i].SaveSBom(ExcelApp, irow, root, dUsage * dusage * dPer / 100, self.lt + lt, i = 0,
          dStockParent + FStock, aSAPStockReader);
      end
      else
      begin
        Childs[i].SaveSBom(ExcelApp, irow, root, dUsage * dusage, self.lt + lt, i = 0,
          dStockParent + FStock, aSAPStockReader);
      end;
    end;
  end;
end;

{ TSAPBomReader }

constructor TSAPBomReader.Create(const sfile: string; aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  FNumbers := TStringList.Create;
  Open;
end;

destructor TSAPBomReader.Destroy;
begin
  Clear;
  FList.Free;
  FNumbers.Free;
  inherited;
end;

procedure TSAPBomReader.Clear;
var
  i: Integer;
  aSapBom: TSapBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSapBom := TSapBom(FList.Objects[i]);
    aSapBom.Free;
  end;
  FList.Clear;

  FNumbers.Clear;
end;

function TSAPBomReader.GetSapBom(const snumber: string): TSapBom;
var
  i: Integer;
  aSapBom: TSapBom;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSapBom := TSapBom(FList.Objects[i]);
    if aSapBom.FNumber = snumber then
    begin
      Result := aSapBom;
      Break;
    end;  
  end;
end;

procedure TSAPBomReader.GetLowestCode(aSAPMaterialRecordPtr: PSAPMaterialRecord);
var
  i: Integer;
  aSapBom: TSapBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSapBom := TSapBom(FList.Objects[i]);
    aSapBom.GetLowestCode(aSAPMaterialRecordPtr);
  end;
end;

procedure TSAPBomReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;

function DotCount(const str: string): Integer;
var
  s: string;
begin
  Result := 0;
  s := str;
  while Pos('.', s) > 0 do
  begin
    Result := Result + 1;
    s := Copy(s, Pos('.', s) + 1, Length(s));
  end;
end;

function SaveLevel(const slevel0, slevel: string): Integer;
var
  pc, pc0: Integer; 
begin
  Result := 0;

  pc := DotCount(slevel);
         
  pc0 := DotCount(slevel0);

  if pc0 = 0 then Exit;
  if pc0 = pc then Exit;

  if pc0 > pc then
    Result := -1
  else Result := 1;
end;  

procedure TSAPBomReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  snumber_child: string;
  snumber: string;
  aSapBom: TSapBom;
  aSapBomChild: TSapBom;
  aSapItemGroup: TSapItemGroup;
  aSapBomLast: TSapBom;
  sgroup0: string;
  slevel0: string;
  ilevel: Integer; 
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

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4;
        if stitle <> '母件物料编码母件物料描述工厂用途' then
        begin
          Log(sSheet +'  不是SAP导出BOM格式');
          Continue;
        end;


        slevel0 := '';
        sgroup0 := '';
        aSapBom := nil;
        aSapBomLast := nil;
        aSapItemGroup := nil;
        irow := 2;
        snumber_child := ExcelApp.Cells[irow, 13].Value;
        while snumber_child <> '' do
        begin
          if FNumbers.IndexOf(snumber_child) < 0 then
          begin
            FNumbers.Add(snumber_child);
          end;
          
          snumber := ExcelApp.Cells[irow, 1].Value;
          if snumber <> '' then
          begin
            if FNumbers.IndexOf(snumber) < 0 then
            begin
              FNumbers.Add(snumber);
            end;

            aSapBom := TSapBom.Create(snumber);
            aSapBom.FName := ExcelApp.Cells[irow, 2].Value;    
            aSapBom.sfac := ExcelApp.Cells[irow, 5].Value;
            aSapBom.lt := ExcelApp.Cells[irow, 6].Value;
            aSapBom.iLowestCode := 0;
            FList.AddObject(snumber, aSapBom);
            slevel0 := '';
            sgroup0 := '';
          end;                    

          aSapBomChild := TSapBom.Create(snumber_child);
          aSapBomChild.slevel := ExcelApp.Cells[irow, 10].Value;
          aSapBomChild.FName := ExcelApp.Cells[irow, 14].Value; 
          aSapBomChild.sptype := ExcelApp.Cells[irow, 15].Value; 
          aSapBomChild.abc := ExcelApp.Cells[irow, 16].Value;
          aSapBomChild.lt := ExcelApp.Cells[irow, 18].Value;
          aSapBomChild.dusage := ExcelApp.Cells[irow, 19].Value;
          aSapBomChild.sgroup := ExcelApp.Cells[irow, 21].Value;
          aSapBomChild.spriority := ExcelApp.Cells[irow, 22].Value;
          aSapBomChild.dPer := ExcelApp.Cells[irow, 23].Value;
          aSapBomChild.iLowestCode := DotCount( aSapBomChild.slevel );
 
 
          // 相同层级
          ilevel := SaveLevel(slevel0, aSapBomChild.slevel);  //判断BOM层级，只能以点的数量判断，不能以字符串长度判断
          case ilevel of
            -1:  // 上一级
            begin                                      
              aSapBom := aSapBomLast.FParent.FParent;
              //while Length(aSapBom.slevel) >= Length(aSapBomChild.slevel) do
              while DotCount(aSapBom.slevel) >= DotCount(aSapBomChild.slevel) do  //判断BOM层级，只能以点的数量判断，不能以字符串长度判断
              begin
                aSapBom := aSapBom.FParent.FParent;
              end;
              aSapItemGroup := aSapBom.Childs[aSapBom.ChildCount - 1];
              aSapBomLast := aSapItemGroup.Items[aSapItemGroup.ItemCount - 1];
              slevel0 := aSapBomLast.slevel;
              sgroup0 := aSapBomLast.sgroup;
              // 替代组为空， 或者替代组不同， 新建一个替代l组项
              if (sgroup0 = '') or (sgroup0 <> aSapBomChild.sgroup) then     //  半成品也有可能替代，这里也考虑了。
              begin
                aSapItemGroup := TSapItemGroup.Create(aSapBomChild.sgroup);
                aSapItemGroup.FParent := aSapBom;
                aSapBom.FList.Add(aSapItemGroup);
              end;              
            end;
            0:   // 同级
            begin
              // 替代组为空， 或者替代组不同， 新建一个替代组项
              if (sgroup0 = '') or (sgroup0 <> aSapBomChild.sgroup) then
              begin
                aSapItemGroup := TSapItemGroup.Create(aSapBomChild.sgroup);
                aSapItemGroup.FParent := aSapBom;                          
                aSapBom.FList.Add(aSapItemGroup);
              end;
            end;
            1:   // 下一级
            begin           
              aSapBom := aSapBomLast;
              aSapItemGroup := TSapItemGroup.Create(aSapBomChild.sgroup);
              aSapItemGroup.FParent := aSapBom;
              aSapBom.FList.Add(aSapItemGroup);
              slevel0 := '';
              sgroup0 := '';
            end;  
          end;

          aSapItemGroup.FList.AddObject(snumber_child, aSapBomChild);
          aSapBomChild.FParent := aSapItemGroup;

          aSapBomLast := aSapBomChild;
          slevel0 := aSapBomLast.slevel;
          sgroup0 := aSapBomLast.sgroup;

 
          irow := irow + 1;
          snumber_child := ExcelApp.Cells[irow, 13].Value;
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

procedure TSAPBomReader.Save(const sfile: string);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  i: Integer;
  aSapBom: TSapBom;
begin


  // 开始保存 Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;

  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;

  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := 'BOM';
  try
    irow := 1;
    ExcelApp.Cells[1, 1].Value := '母件物料编码';
    ExcelApp.Cells[1, 2].Value := '母件物料描述';    
    ExcelApp.Cells[1, 3].Value := '母件L/T';
    ExcelApp.Cells[1, 4].Value := '子件物料编码';
    ExcelApp.Cells[1, 5].Value := '子件物料描述';
    ExcelApp.Cells[1, 6].Value := '层级';
    ExcelApp.Cells[1, 7].Value := '采购类型';
    ExcelApp.Cells[1, 8].Value := 'ABC标识';
    ExcelApp.Cells[1, 9].Value := '替代项目组';         
    ExcelApp.Cells[1, 10].Value := 'L/T';


    irow := 2;
    for i := 0 to FList.Count - 1 do
    begin
      aSapBom := TSapBom(FList.Objects[i]);
      ExcelApp.Cells[irow, 1].Value := aSapBom.FNumber;
      ExcelApp.Cells[irow, 2].Value := aSapBom.FName;   
      ExcelApp.Cells[irow, 3].Value := 0;
      aSapBom.Save(ExcelApp, irow);
    end;

  
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end; 
  MessageBox(0, '完成', '提示', 0);

end;    

procedure TSAPBomReader.SaveSBom(const sfile: string;
  aSAPStockReader: TSAPStockReader);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  i: Integer;
  aSapBom: TSapBom;
  irow1: Integer;
begin
  // 开始保存 Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := True;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;

  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;

  ExcelApp.Sheets[1].Activate;
  ExcelApp.Sheets[1].Name := 'BOM';
  try
    irow := 1;
    ExcelApp.Cells[1, 1].Value := '产品编码';
    ExcelApp.Cells[1, 2].Value := '产品名称';
    ExcelApp.Cells[1, 3].Value := '物料编码';
    ExcelApp.Cells[1, 4].Value := '物料名称';
    ExcelApp.Cells[1, 5].Value := '提前期';   
    ExcelApp.Cells[1, 6].Value := '替代组';   
    ExcelApp.Cells[1, 7].Value := '半成品库存';
    ExcelApp.Cells[1, 8].Value := '库存';          
    ExcelApp.Cells[1, 9].Value := '用量';
    ExcelApp.Cells[1, 10].Value := '优先级';

    irow := 2;
    for i := 0 to FList.Count - 1 do
    begin
      aSapBom := TSapBom(FList.Objects[i]);

      irow1 := irow;
      aSapBom.SaveSBom(ExcelApp, irow, aSapBom, 1, 0, 0, aSAPStockReader);

      if irow1 = irow then Continue; // 没有关键物料
            
      ExcelApp.Cells[irow1, 1].Value := aSapBom.FNumber;
      ExcelApp.Cells[irow1, 2].Value := aSapBom.FName;
      
      MergeCells(ExcelApp, irow1, 1, irow - 1, 1);
      MergeCells(ExcelApp, irow1, 2, irow - 1, 2);
    end;

  
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end; 

end;

end.
