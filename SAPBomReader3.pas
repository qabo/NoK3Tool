unit SAPBomReader3;

interface

uses
  Classes, ComObj, ActiveX, SysUtils, Windows, CommUtils, SAPStockReader,
  ADODB, SAPMaterialReader, SAPBomReader;

type 
  TSAPBomReader3 = class
  private
    FFile: string;
//    ExcelApp, WorkBook: Variant;
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
    function GetSapBom(const snumber, sbomfac: string): TSapBom;
    procedure GetLowestCode(aSAPMaterialRecordPtr: PSAPMaterialRecord);
  end;

implementation
 
{ TSAPBomReader3 }

constructor TSAPBomReader3.Create(const sfile: string; aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  FNumbers := TStringList.Create;
  Open;
end;

destructor TSAPBomReader3.Destroy;
begin
  Clear;
  FList.Free;
  FNumbers.Free;
  inherited;
end;

procedure TSAPBomReader3.Clear;
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

function TSAPBomReader3.GetSapBom(const snumber, sbomfac: string): TSapBom;
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
      if sbomfac = '' then
      begin
        Break;
      end
      else
      begin
        if sbomfac =aSapBom.sfac then
        begin
          Break;
        end;
      end;  
    end;  
  end;
end;

procedure TSAPBomReader3.GetLowestCode(aSAPMaterialRecordPtr: PSAPMaterialRecord);
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

procedure TSAPBomReader3.Log(const str: string);
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

procedure TSAPBomReader3.Open;
var
//  iSheetCount, iSheet: Integer;
//  sSheet: string;
//  stitle1, stitle2, stitle3, stitle4: string;
//  stitle: string;
//  irow: Integer;
  snumber_child: string;
  snumber: string;
  aSapBom: TSapBom;
  aSapBomChild: TSapBom;
  aSapItemGroup: TSapItemGroup;
  aSapBomLast: TSapBom;
  //sgroup: string;
  slevel0: string;
  ilevel: Integer;
  
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
begin
  Clear;

        
  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;


  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'Sheet1'+'$]';

    ADOTabXLS.Active:=true;


 
    slevel0 := '';
    //sgroup0 := '';
    aSapBom := nil;
    aSapBomLast := nil;
    aSapItemGroup := nil;
//    irow := 2;
    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin                 
      snumber_child := ADOTabXLS.FieldByName('�Ӽ����ϱ���').AsString; // ExcelApp.Cells[irow, 13].Value;
      snumber_child := Trim(snumber_child);
      
      if snumber_child = '' then Break;

      if FNumbers.IndexOf(snumber_child) < 0 then
      begin
        FNumbers.Add(snumber_child);
      end;
          
      snumber := ADOTabXLS.FieldByName('ĸ�����ϱ���').AsString; //ExcelApp.Cells[irow, 1].Value;
      snumber := Trim(snumber);
      if snumber <> '' then
      begin
        if FNumbers.IndexOf(snumber) < 0 then
        begin
          FNumbers.Add(snumber);
        end;

        aSapBom := TSapBom.Create(snumber);
        aSapBom.FName := ADOTabXLS.FieldByName('ĸ����������').AsString; // ExcelApp.Cells[irow, 2].Value;
        aSapBom.lt := ADOTabXLS.FieldByName('ĸ��L/T').AsFloat; //  ExcelApp.Cells[irow, 6].Value;
        aSapBom.iLowestCode := 0;
        aSapBom.sfac := ADOTabXLS.FieldByName('������').AsString;
        FList.AddObject(snumber, aSapBom);
        slevel0 := '';
        //sgroup0 := '';
      end;                    

      aSapBomChild := TSapBom.Create(snumber_child);
      aSapBomChild.slevel := ADOTabXLS.FieldByName('�㼶').AsString; //  ExcelApp.Cells[irow, 10].Value;
      aSapBomChild.FName := ADOTabXLS.FieldByName('�Ӽ���������').AsString; //   ExcelApp.Cells[irow, 14].Value;
      aSapBomChild.sptype := ADOTabXLS.FieldByName('�ɹ�����').AsString; //   ExcelApp.Cells[irow, 15].Value;
      aSapBomChild.abc := ADOTabXLS.FieldByName('ABC��ʶ').AsString; //   ExcelApp.Cells[irow, 16].Value;
      aSapBomChild.lt := ADOTabXLS.FieldByName('L/T').AsFloat; //  ExcelApp.Cells[irow, 18].Value;
      aSapBomChild.dusage := ADOTabXLS.FieldByName('�Ӽ�����').AsFloat; //  ExcelApp.Cells[irow, 19].Value;
      aSapBomChild.sgroup := ADOTabXLS.FieldByName('�����Ŀ��').AsString; //  ExcelApp.Cells[irow, 21].Value;
      aSapBomChild.spriority := ADOTabXLS.FieldByName('���ȼ�').AsString; //  ExcelApp.Cells[irow, 22].Value;
      aSapBomChild.dPer := ADOTabXLS.FieldByName('ʹ�ÿ�����').AsFloat; //  ExcelApp.Cells[irow, 23].Value;
      aSapBomChild.iLowestCode := DotCount( aSapBomChild.slevel );

 
      // ��ͬ�㼶
      ilevel := SaveLevel(slevel0, aSapBomChild.slevel);  //�ж�BOM�㼶��ֻ���Ե�������жϣ��������ַ��������ж�
      case ilevel of
        -1:  // ��һ��
        begin                                      
          aSapBom := aSapBomLast.FParent.FParent;
          //while Length(aSapBom.slevel) >= Length(aSapBomChild.slevel) do
          while DotCount(aSapBom.slevel) >= DotCount(aSapBomChild.slevel) do  //�ж�BOM�㼶��ֻ���Ե�������жϣ��������ַ��������ж�
          begin
            aSapBom := aSapBom.FParent.FParent;
          end;
          aSapItemGroup := aSapBom.Childs[aSapBom.ChildCount - 1];
          aSapBomLast := aSapItemGroup.Items[aSapItemGroup.ItemCount - 1];
          slevel0 := aSapBomLast.slevel;
          //sgroup0 := aSapBomLast.sgroup;
          // �����Ϊ�գ� �����Ҳ�������飬 �½�һ�����l����
          aSapItemGroup := aSapBom.ChildByGroup(aSapBomChild.sgroup);
          if aSapItemGroup = nil then     //  ���ƷҲ�п������������Ҳ�����ˡ�
          begin
            aSapItemGroup := TSapItemGroup.Create(aSapBomChild.sgroup);
            aSapItemGroup.FParent := aSapBom;  
            aSapBom.FList.Add(aSapItemGroup);
          end
        end;
        0:   // ͬ��
        begin
          // �����Ϊ�գ� �����Ҳ�������飬 �½�һ�����l����
          aSapItemGroup := aSapBom.ChildByGroup(aSapBomChild.sgroup);
          if aSapItemGroup = nil then     //  ���ƷҲ�п������������Ҳ�����ˡ�
          begin
            aSapItemGroup := TSapItemGroup.Create(aSapBomChild.sgroup);
            aSapItemGroup.FParent := aSapBom;      
            aSapBom.FList.Add(aSapItemGroup);
          end;
        end;
        1:   // ��һ��
        begin           
          aSapBom := aSapBomLast;
          aSapItemGroup := TSapItemGroup.Create(aSapBomChild.sgroup);
          aSapItemGroup.FParent := aSapBom;
          aSapBom.FList.Add(aSapItemGroup);
          slevel0 := '';
//          sgroup0 := '';
        end;
      end;

      aSapItemGroup.FList.AddObject(snumber_child, aSapBomChild);
      aSapBomChild.FParent := aSapItemGroup;

      aSapBomLast := aSapBomChild;
      slevel0 := aSapBomLast.slevel;
//      sgroup0 := aSapBomLast.sgroup;
 
//      irow := irow + 1;   
      ADOTabXLS.Next;
      snumber_child := ADOTabXLS.FieldByName('�Ӽ����ϱ���').AsString; //  ExcelApp.Cells[irow, 13].Value;
      snumber_child := Trim(snumber_child);
    end; 


       
    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;

  FList.Sort;
end;

procedure TSAPBomReader3.Save(const sfile: string);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  i: Integer;
  aSapBom: TSapBom;
begin


  // ��ʼ���� Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '�����ʾ', 0);
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
    ExcelApp.Cells[1, 1].Value := 'ĸ�����ϱ���';
    ExcelApp.Cells[1, 2].Value := 'ĸ����������';    
    ExcelApp.Cells[1, 3].Value := 'ĸ��L/T';
    ExcelApp.Cells[1, 4].Value := '�Ӽ����ϱ���';
    ExcelApp.Cells[1, 5].Value := '�Ӽ���������';
    ExcelApp.Cells[1, 6].Value := '�㼶';
    ExcelApp.Cells[1, 7].Value := '�ɹ�����';
    ExcelApp.Cells[1, 8].Value := 'ABC��ʶ';
    ExcelApp.Cells[1, 9].Value := '�����Ŀ��';         
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
    ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end; 
  MessageBox(0, '���', '��ʾ', 0);

end;    

procedure TSAPBomReader3.SaveSBom(const sfile: string;
  aSAPStockReader: TSAPStockReader);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  i: Integer;
  aSapBom: TSapBom;
  irow1: Integer;
begin
  // ��ʼ���� Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := True;
    ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '�����ʾ', 0);
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
    ExcelApp.Cells[1, 1].Value := '��Ʒ����';
    ExcelApp.Cells[1, 2].Value := '��Ʒ����';
    ExcelApp.Cells[1, 3].Value := '���ϱ���';
    ExcelApp.Cells[1, 4].Value := '��������';
    ExcelApp.Cells[1, 5].Value := '��ǰ��';   
    ExcelApp.Cells[1, 6].Value := '�����';   
    ExcelApp.Cells[1, 7].Value := '���Ʒ���';
    ExcelApp.Cells[1, 8].Value := '���';          
    ExcelApp.Cells[1, 9].Value := '����';
    ExcelApp.Cells[1, 10].Value := '���ȼ�';

    irow := 2;
    for i := 0 to FList.Count - 1 do
    begin
      aSapBom := TSapBom(FList.Objects[i]);

      irow1 := irow;
      aSapBom.SaveSBom(ExcelApp, irow, aSapBom, 1, 0, 0, aSAPStockReader);

      if irow1 = irow then Continue; // û�йؼ�����
            
      ExcelApp.Cells[irow1, 1].Value := aSapBom.FNumber;
      ExcelApp.Cells[irow1, 2].Value := aSapBom.FName;
      
      MergeCells(ExcelApp, irow1, 1, irow - 1, 1);
      MergeCells(ExcelApp, irow1, 2, irow - 1, 2);
    end;

  
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end; 

end;

end.
