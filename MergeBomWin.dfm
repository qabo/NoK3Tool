object frmMergeBom: TfrmMergeBom
  Left = 380
  Top = 322
  Width = 866
  Height = 469
  Caption = 'frmMergeBom'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 16
    Top = 16
    Width = 177
    Height = 41
    Caption = #27983#35272
    TabOrder = 0
    OnClick = Button1Click
  end
  object vleBomsFOX: TValueListEditor
    Left = 208
    Top = 16
    Width = 625
    Height = 193
    Hint = #21491#38190#21487#20197#21024#38500'Bom'
    ParentShowHint = False
    PopupMenu = PopupMenu1
    ShowHint = True
    TabOrder = 1
    TitleCaptions.Strings = (
      'Bom'
      #23500#22763#24247#29983#20135#25968#37327)
    ColWidths = (
      523
      96)
  end
  object Button2: TButton
    Left = 16
    Top = 168
    Width = 177
    Height = 41
    Caption = #23548#20986
    TabOrder = 2
    OnClick = Button2Click
  end
  object vleBomsML: TValueListEditor
    Left = 208
    Top = 224
    Width = 625
    Height = 193
    Hint = #21491#38190#21487#20197#21024#38500'Bom'
    ParentShowHint = False
    ShowHint = False
    TabOrder = 3
    TitleCaptions.Strings = (
      'Bom'
      #39749#21147#29983#20135#25968#37327)
    ColWidths = (
      523
      96)
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = 'xlsx'
    Filter = 'Excel Files|*.xls;*.xlsx'
    Left = 24
    Top = 232
  end
  object SaveDialog1: TSaveDialog
    Left = 72
    Top = 232
  end
  object PopupMenu1: TPopupMenu
    Left = 304
    Top = 128
    object N1: TMenuItem
      Caption = #21024#38500
      OnClick = N1Click
    end
  end
  object ADOQuery1: TADOQuery
    CursorType = ctStatic
    CommandTimeout = 300
    Parameters = <>
    SQL.Strings = (
      'declare @number1 varchar(50), @number2 varchar(50)'
      'select @number1='#39'01.01.1090196B'#39', @number2='#39'01.01.1090196B'#39
      ''
      'declare @CurrentYear int'
      'declare @CurrentPeriod int'
      
        ' select @CurrentYear=FValue from t_SystemProfile where FCategory' +
        '='#39'IC'#39' and FKey='#39'CurrentYear'#39
      
        ' select @CurrentPeriod=FValue from t_SystemProfile where FCatego' +
        'ry='#39'IC'#39' and FKey='#39'CurrentPeriod'#39
      ''
      
        '--SELECT CONVERT(DATETIME,FStartDate) as FStartDate, DATEADD(DAY' +
        ',1,CONVERT(DATETIME,FEndDate)) as FEndDate FROM t_PeriodDate WHE' +
        'RE FYear = 2008 AND FPeriod = 2'
      ''
      'Set NoCount On '
      ' SET ANSI_WARNINGS OFF '
      ' Create Table #Happen2('
      '        FItemID int Null, '
      '        FStockID int Null, '
      '        FBatchNo NVARCHAR(200), '
      '        FQty decimal(28,10) Null, '
      '        FCUUnitQty decimal(28,10) Null, '
      '        FQty1 Decimal(28,10), '
      '        FCUUnitQty1 Decimal(28,10), '
      '        FAmount1 Decimal(28,10), '
      '        FQty2 Decimal(28,10), '
      '        FCUUnitQty2 Decimal(28,10), '
      '        FAmount2 Decimal(28,10), '
      '        FQty3 Decimal(28,10), '
      '        FCUUnitQty3 Decimal(28,10), '
      '        FAmount3 Decimal(28,10), '
      '        FQty4 Decimal(28,10), '
      '        FCUUnitQty4 Decimal(28,10), '
      '        FAmount4 Decimal(28,10), '
      '        FQty5 Decimal(28,10), '
      '        FCUUnitQty5 Decimal(28,10), '
      '        FAmount5 Decimal(28,10), '
      '        FQty6 Decimal(28,10), '
      '        FCUUnitQty6 Decimal(28,10), '
      '        FAmount6 Decimal(28,10), '
      '        FQty7 Decimal(28,10), '
      '        FCUUnitQty7 Decimal(28,10), '
      '        FAmount7 Decimal(28,10), '
      
        '        FTemp bit ,FPrice Decimal(28,10),FCUPrice Decimal(28,10)' +
        ',FAmount Decimal(28,10))'
      ' Create Table #Happen('
      '        FItemID int Null, '
      '        FStockID int Null, '
      '        FBatchNo NVARCHAR(200), '
      '        FQty decimal(28,10) Null, '
      '        FCUUnitQty decimal(28,10) Null, '
      '        FQty1 Decimal(28,10), '
      '        FCUUnitQty1 Decimal(28,10), '
      '        FAmount1 Decimal(28,10), '
      '        FQty2 Decimal(28,10), '
      '        FCUUnitQty2 Decimal(28,10), '
      '        FAmount2 Decimal(28,10), '
      '        FQty3 Decimal(28,10), '
      '        FCUUnitQty3 Decimal(28,10), '
      '        FAmount3 Decimal(28,10), '
      '        FQty4 Decimal(28,10), '
      '        FCUUnitQty4 Decimal(28,10), '
      '        FAmount4 Decimal(28,10), '
      '        FQty5 Decimal(28,10), '
      '        FCUUnitQty5 Decimal(28,10), '
      '        FAmount5 Decimal(28,10), '
      '        FQty6 Decimal(28,10), '
      '        FCUUnitQty6 Decimal(28,10), '
      '        FAmount6 Decimal(28,10), '
      '        FQty7 Decimal(28,10), '
      '        FCUUnitQty7 Decimal(28,10), '
      '        FAmount7 Decimal(28,10), '
      
        '        FTemp bit ,FPrice Decimal(28,10),FCUPrice Decimal(28,10)' +
        ',FAmount Decimal(28,10))'
      ' Create Table #Happen1('
      '        FItemID int Null, '
      '        FStockID int Null, '
      '        FBatchNo NVARCHAR(200), '
      '        FQty decimal(28,10) Null, '
      '        FCUUnitQty decimal(28,10) Null, '
      'FQty1 Decimal(28,10), '
      'FCUUnitQty1 Decimal(28,10), '
      'FQty2 Decimal(28,10), '
      'FCUUnitQty2 Decimal(28,10), '
      'FQty3 Decimal(28,10), '
      'FCUUnitQty3 Decimal(28,10), '
      'FQty4 Decimal(28,10), '
      'FCUUnitQty4 Decimal(28,10), '
      'FQty5 Decimal(28,10), '
      'FCUUnitQty5 Decimal(28,10), '
      'FQty6 Decimal(28,10), '
      'FCUUnitQty6 Decimal(28,10), '
      'FQty7 Decimal(28,10), '
      'FCUUnitQty7 Decimal(28,10), '
      
        ' FTemp bit ,FPrice Decimal(28,10),FCUPrice Decimal(28,10),FAmoun' +
        't Decimal(28,10) )'
      '  Insert Into #Happen1'
      '  Select t1.FItemID,t2.FItemID As FStockID,t6.FBatchNo,0,0,'
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-30,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Var' +
        'char,DATEADD(dd,-0,Getdate()),101)) Then t5.FRob*t6.FQty Else 0 ' +
        'End)  As FQty1,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-30,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Va' +
        'rchar,DATEADD(dd,-0,Getdate()),101)) Then CAST(t5.FRob*t6.FQty A' +
        'S DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQty1,'
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-90,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Var' +
        'char,DATEADD(dd,-31,Getdate()),101)) Then t5.FRob*t6.FQty Else 0' +
        ' End)  As FQty2,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-90,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Va' +
        'rchar,DATEADD(dd,-31,Getdate()),101)) Then CAST(t5.FRob*t6.FQty ' +
        'AS DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQty2,'
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-180,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Va' +
        'rchar,DATEADD(dd,-91,Getdate()),101)) Then t5.FRob*t6.FQty Else ' +
        '0 End)  As FQty3,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-180,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(V' +
        'archar,DATEADD(dd,-91,Getdate()),101)) Then CAST(t5.FRob*t6.FQty' +
        ' AS DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQty3' +
        ','
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-365,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Va' +
        'rchar,DATEADD(dd,-181,Getdate()),101)) Then t5.FRob*t6.FQty Else' +
        ' 0 End)  As FQty4,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-365,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(V' +
        'archar,DATEADD(dd,-181,Getdate()),101)) Then CAST(t5.FRob*t6.FQt' +
        'y AS DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQty' +
        '4,'
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-730,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(Va' +
        'rchar,DATEADD(dd,-366,Getdate()),101)) Then t5.FRob*t6.FQty Else' +
        ' 0 End)  As FQty5,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-730,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(V' +
        'archar,DATEADD(dd,-366,Getdate()),101)) Then CAST(t5.FRob*t6.FQt' +
        'y AS DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQty' +
        '5,'
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-1095,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(V' +
        'archar,DATEADD(dd,-731,Getdate()),101)) Then t5.FRob*t6.FQty Els' +
        'e 0 End)  As FQty6,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-1095,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(' +
        'Varchar,DATEADD(dd,-731,Getdate()),101)) Then CAST(t5.FRob*t6.FQ' +
        'ty AS DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQt' +
        'y6,'
      
        '(Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(dd' +
        ',-1096,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(V' +
        'archar,DATEADD(dd,-1096,Getdate()),101)) Then t5.FRob*t6.FQty El' +
        'se 0 End)  As FQty7,'
      
        '((Case When t5.FDate>=Convert(DateTime,Convert(Varchar,DATEADD(d' +
        'd,-1096,Getdate()),101)) And t5.FDate<=Convert(DateTime,Convert(' +
        'Varchar,DATEADD(dd,-1096,Getdate()),101)) Then CAST(t5.FRob*t6.F' +
        'Qty AS DECIMAL(28,10)) Else 0 End))/t7.FCoefficient  As FCUUnitQ' +
        'ty7,'
      '1  ,ISNULL(((SELECT t12.FBegBal/t12.FBegQty FROM t_ICItem t11'
      
        ' INNER JOIN (SELECT FItemID,SUM(FBegBal) AS FBegBal,SUM(FBegQty)' +
        ' AS FBegQty FROM ICBal t13'
      
        ' Where t13.FItemID = t6.FItemID And t13.FYear = @CurrentYear And' +
        ' t13.FPeriod = @CurrentPeriod'
      
        ' GROUP BY t13.FItemID) t12 ON t11.FItemID=t12.FItemID AND t12.FB' +
        'egBal>0 AND t12.FBegQty>0) ),0),0,0'
      ' From t_ICItem t1'
      
        'Join ICStockBill t5 On (t5.FStatus > 0 Or (t5.FUpStockWhenSave >' +
        ' 0 And t5.FCancellation <1 ))'
      'Join ICStockBillEntry t6 On t5.FInterID=t6.FInterID'
      'Left Join t_MeasureUnit t7 On t1.FStoreUnitID=t7.FMeasureUnitID'
      
        'Left Join t_Stock t2 On t2.FItemID = (case when t5.ftrantype=24 ' +
        'then t6.FSCStockID else t6.FDCStockID end) '
      ' Where t1.FItemID = t6.FItemID '
      
        ' And ((t5.FTrantype In (1,2,5,10,40) And t5.FRob =1) Or (t5.FTra' +
        'ntype In(21,24,28,29) And t5.FRob=-1))'
      ' AND (NOT (t5.FTrantype In (1) AND t5.FPOMode = 36681 ))'
      
        ' AND (NOT (t5.FTranType=1 and t6.FSourceInterID > 0 and EXISTS(S' +
        'ELECT 1 FROM ICHookRelations t8 where t6.FinterID=t8.fIBInterID ' +
        'and t8.FIBTag=4 )))'
      
        ' And t1.FNumber>=@number1 And t1. FNumber<=@number2 And t2.FType' +
        'ID NOT IN (504)'
      ''
      ''
      '  Insert Into #Happen1'
      '  Select t1.FItemID,t2.FItemID As FStockID,t6.FBatchNo,0,0,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-30,Getdate()),101)) And t6.FStockInDate<=Convert(DateTi' +
        'me,Convert(Varchar,DATEADD(dd,-0,Getdate()),101)) Then t6.FBegQt' +
        'y Else 0 End)  As FQty1,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-30,Getdate()),101)) And t6.FStockInDate<=Convert(DateT' +
        'ime,Convert(Varchar,DATEADD(dd,-0,Getdate()),101)) Then t6.FBegQ' +
        'ty Else 0 End))/t7.FCoefficient  As FCUUnitQty1,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-90,Getdate()),101)) And t6.FStockInDate<=Convert(DateTi' +
        'me,Convert(Varchar,DATEADD(dd,-31,Getdate()),101)) Then t6.FBegQ' +
        'ty Else 0 End)  As FQty2,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-90,Getdate()),101)) And t6.FStockInDate<=Convert(DateT' +
        'ime,Convert(Varchar,DATEADD(dd,-31,Getdate()),101)) Then t6.FBeg' +
        'Qty Else 0 End))/t7.FCoefficient  As FCUUnitQty2,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-180,Getdate()),101)) And t6.FStockInDate<=Convert(DateT' +
        'ime,Convert(Varchar,DATEADD(dd,-91,Getdate()),101)) Then t6.FBeg' +
        'Qty Else 0 End)  As FQty3,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-180,Getdate()),101)) And t6.FStockInDate<=Convert(Date' +
        'Time,Convert(Varchar,DATEADD(dd,-91,Getdate()),101)) Then t6.FBe' +
        'gQty Else 0 End))/t7.FCoefficient  As FCUUnitQty3,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-365,Getdate()),101)) And t6.FStockInDate<=Convert(DateT' +
        'ime,Convert(Varchar,DATEADD(dd,-181,Getdate()),101)) Then t6.FBe' +
        'gQty Else 0 End)  As FQty4,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-365,Getdate()),101)) And t6.FStockInDate<=Convert(Date' +
        'Time,Convert(Varchar,DATEADD(dd,-181,Getdate()),101)) Then t6.FB' +
        'egQty Else 0 End))/t7.FCoefficient  As FCUUnitQty4,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-730,Getdate()),101)) And t6.FStockInDate<=Convert(DateT' +
        'ime,Convert(Varchar,DATEADD(dd,-366,Getdate()),101)) Then t6.FBe' +
        'gQty Else 0 End)  As FQty5,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-730,Getdate()),101)) And t6.FStockInDate<=Convert(Date' +
        'Time,Convert(Varchar,DATEADD(dd,-366,Getdate()),101)) Then t6.FB' +
        'egQty Else 0 End))/t7.FCoefficient  As FCUUnitQty5,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-1095,Getdate()),101)) And t6.FStockInDate<=Convert(Date' +
        'Time,Convert(Varchar,DATEADD(dd,-731,Getdate()),101)) Then t6.FB' +
        'egQty Else 0 End)  As FQty6,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-1095,Getdate()),101)) And t6.FStockInDate<=Convert(Dat' +
        'eTime,Convert(Varchar,DATEADD(dd,-731,Getdate()),101)) Then t6.F' +
        'BegQty Else 0 End))/t7.FCoefficient  As FCUUnitQty6,'
      
        '(Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DAT' +
        'EADD(dd,-1096,Getdate()),101)) And t6.FStockInDate<=Convert(Date' +
        'Time,Convert(Varchar,DATEADD(dd,-1096,Getdate()),101)) Then t6.F' +
        'BegQty Else 0 End)  As FQty7,'
      
        '((Case When t6.FStockInDate>=Convert(DateTime,Convert(Varchar,DA' +
        'TEADD(dd,-1096,Getdate()),101)) And t6.FStockInDate<=Convert(Dat' +
        'eTime,Convert(Varchar,DATEADD(dd,-1096,Getdate()),101)) Then t6.' +
        'FBegQty Else 0 End))/t7.FCoefficient  As FCUUnitQty7,'
      '1   ,ISNULL(((SELECT t12.FBegBal/t12.FBegQty FROM t_ICItem t11'
      
        ' INNER JOIN (SELECT FItemID,SUM(FBegBal) AS FBegBal,SUM(FBegQty)' +
        ' AS FBegQty FROM ICBal t13'
      
        ' Where t13.FItemID = t6.FItemID And t13.FYear = @CurrentYear And' +
        ' t13.FPeriod = @CurrentPeriod'
      
        ' GROUP BY t13.FItemID) t12 ON t11.FItemID=t12.FItemID AND t12.FB' +
        'egBal>0 AND t12.FBegQty>0) ),0),0,0'
      ''
      'From t_ICItem t1'
      'Join ICInvInitial t6 On t1.FItemID = t6.FItemID'
      'Left Join t_MeasureUnit t7 On t1.FStoreUnitID=t7.FMeasureUnitID'
      'Left Join t_Stock t2 On t2.FItemID = t6.FStockID'
      
        ' Where 1=1  And t1.FNumber>=@number1 And t1. FNumber<=@number2 A' +
        'nd t2.FTypeID NOT IN (504)'
      ''
      ''
      ' Insert Into #Happen1 '
      
        'Select t1.FItemID,t2.FItemID As FStockID,t3.FBatchNo,(t3.FQTY) A' +
        's FQTY,CAST(t3.FQTY AS DECIMAL(28,10))/t7.FCoefficient As FCUUni' +
        'tQty,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1   ,ISNULL(((SELECT t12.FBegBa' +
        'l/t12.FBegQty FROM t_ICItem t11'
      
        ' INNER JOIN (SELECT FItemID,SUM(FBegBal) AS FBegBal,SUM(FBegQty)' +
        ' AS FBegQty FROM ICBal t13'
      
        ' Where t13.FItemID = t3.FItemID And t13.FYear = @CurrentYear And' +
        ' t13.FPeriod = @CurrentPeriod'
      
        ' GROUP BY t13.FItemID) t12 ON t11.FItemID=t12.FItemID AND t12.FB' +
        'egBal>0 AND t12.FBegQty>0) ),0),0,0'
      ''
      'From t_ICItem t1'
      ' Join ICINVENTORY t3 On t1.FItemID = t3.FItemID '
      'Left Join t_MeasureUnit t7 On t1.FStoreUnitID=t7.FMeasureUnitID'
      'Left Join t_Stock t2 On t2.FItemID = t3.FStockID'
      'Where 1=1'
      
        ' And t1.FNumber>=@number1 And t1. FNumber<=@number2 And t2.FType' +
        'ID NOT IN (504)'
      ''
      ''
      ' Insert Into #HAPPEN2'
      
        'Select t1.FITEMID,FStockID,t1.FBatchNo,Sum(FQTY)As FQTY,Sum(FCUU' +
        'nitQTY)As FCUUnitQTY,'
      
        'Sum(fqty1) As FQty1 ,Sum(fCUUnitqty1) As FCUUnitQty1 ,0,Sum(fqty' +
        '2) As FQty2 ,Sum(fCUUnitqty2) As FCUUnitQty2 ,0,Sum(fqty3) As FQ' +
        'ty3 ,Sum(fCUUnitqty3) As FCUUnitQty3 ,0,Sum(fqty4) As FQty4 ,Sum' +
        '(fCUUnitqty4) As FCUUnitQty4 ,0,Sum(fqty5) As FQty5 ,Sum(fCUUnit' +
        'qty5) As FCUUnitQty5 ,0,Sum(fqty6) As FQty6 ,Sum(fCUUnitqty6) As' +
        ' FCUUnitQty6 ,0,Sum(fqty7) As FQty7 ,Sum(fCUUnitqty7) As FCUUnit' +
        'Qty7 ,0,1,Min(FPrice),case Sum(FCUUnitQTY) when 0 then 0 else (M' +
        'in(FPrice)*Sum(FQTY))/Sum(FCUUnitQTY) end,Min(FPrice)*Sum(FQTY) '
      
        'From #HAPPEN1 t1 INNER JOIN t_ICItem t2 ON t1.FItemID=t2.FItemID' +
        '  '
      
        'GROUP By t1.FITEMID,t1.FBatchNo,FStockID Update #Happen2 Set FQt' +
        'y1= FQty,FQty2=0'
      ',FQty3=0'
      ',FQty4=0'
      ',FQty5=0'
      ',FQty6=0'
      ',FQty7=0'
      ' Where  FQty-FQty1<0 '
      ' Update #Happen2 Set FQty2= FQty-FQty1,FQty3=0'
      ',FQty4=0'
      ',FQty5=0'
      ',FQty6=0'
      ',FQty7=0'
      ' Where  FQty-FQty1-FQty2<0 '
      ' Update #Happen2 Set FQty3= FQty-FQty1-FQty2,FQty4=0'
      ',FQty5=0'
      ',FQty6=0'
      ',FQty7=0'
      ' Where  FQty-FQty1-FQty2-FQty3<0 '
      ' Update #Happen2 Set FQty4= FQty-FQty1-FQty2-FQty3,FQty5=0'
      ',FQty6=0'
      ',FQty7=0'
      ' Where  FQty-FQty1-FQty2-FQty3-FQty4<0 '
      ' Update #Happen2 Set FQty5= FQty-FQty1-FQty2-FQty3-FQty4,FQty6=0'
      ',FQty7=0'
      ' Where  FQty-FQty1-FQty2-FQty3-FQty4-FQty5<0 '
      
        ' Update #Happen2 Set FQty6= FQty-FQty1-FQty2-FQty3-FQty4-FQty5,F' +
        'Qty7=0'
      ' Where  FQty-FQty1-FQty2-FQty3-FQty4-FQty5-FQty6<0 '
      
        ' Update #Happen2 Set FQty7= FQty-FQty1-FQty2-FQty3-FQty4-FQty5-F' +
        'Qty6'
      'Update #Happen2 Set FAmount1=FPrice*FQty1'
      'Update #Happen2 Set FAmount2=FPrice*FQty2'
      'Update #Happen2 Set FAmount3=FPrice*FQty3'
      'Update #Happen2 Set FAmount4=FPrice*FQty4'
      'Update #Happen2 Set FAmount5=FPrice*FQty5'
      'Update #Happen2 Set FAmount6=FPrice*FQty6'
      'Update #Happen2 Set FAmount7=FPrice*FQty7'
      ''
      ' Update #Happen2 Set FCUUnitQty1= FCUUnitQty,FCUUnitQty2=0'
      ',FCUUnitQty3=0'
      ',FCUUnitQty4=0'
      ',FCUUnitQty5=0'
      ',FCUUnitQty6=0'
      ',FCUUnitQty7=0'
      ' Where  FCUUnitQty-FCUUnitQty1<0 '
      ' Update #Happen2 Set FCUUnitQty2= FCUUnitQty-FCUUnitQty1'
      ',FCUUnitQty3=0'
      ',FCUUnitQty4=0'
      ',FCUUnitQty5=0'
      ',FCUUnitQty6=0'
      ',FCUUnitQty7=0'
      ' Where  FCUUnitQty-FCUUnitQty1-FCUUnitQty2<0 '
      ' Update #Happen2 Set FCUUnitQty3= FCUUnitQty-FCUUnitQty1'
      '-FCUUnitQty2'
      ',FCUUnitQty4=0'
      ',FCUUnitQty5=0'
      ',FCUUnitQty6=0'
      ',FCUUnitQty7=0'
      ' Where  FCUUnitQty-FCUUnitQty1-FCUUnitQty2-FCUUnitQty3<0 '
      ' Update #Happen2 Set FCUUnitQty4= FCUUnitQty-FCUUnitQty1'
      '-FCUUnitQty2'
      '-FCUUnitQty3'
      ',FCUUnitQty5=0'
      ',FCUUnitQty6=0'
      ',FCUUnitQty7=0'
      
        ' Where  FCUUnitQty-FCUUnitQty1-FCUUnitQty2-FCUUnitQty3-FCUUnitQt' +
        'y4<0 '
      ' Update #Happen2 Set FCUUnitQty5= FCUUnitQty-FCUUnitQty1'
      '-FCUUnitQty2'
      '-FCUUnitQty3'
      '-FCUUnitQty4'
      ',FCUUnitQty6=0'
      ',FCUUnitQty7=0'
      
        ' Where  FCUUnitQty-FCUUnitQty1-FCUUnitQty2-FCUUnitQty3-FCUUnitQt' +
        'y4-FCUUnitQty5<0 '
      ' Update #Happen2 Set FCUUnitQty6= FCUUnitQty-FCUUnitQty1'
      '-FCUUnitQty2'
      '-FCUUnitQty3'
      '-FCUUnitQty4'
      '-FCUUnitQty5'
      ',FCUUnitQty7=0'
      
        ' Where  FCUUnitQty-FCUUnitQty1-FCUUnitQty2-FCUUnitQty3-FCUUnitQt' +
        'y4-FCUUnitQty5-FCUUnitQty6<0 '
      ' Update #Happen2 Set FCUUnitQty7= FCUUnitQty-FCUUnitQty1'
      '-FCUUnitQty2'
      '-FCUUnitQty3'
      '-FCUUnitQty4'
      '-FCUUnitQty5'
      '-FCUUnitQty6'
      ''
      ' Insert Into #HAPPEN'
      
        'Select t1.FITEMID,t1.FStockID,t1.FBatchNo,Sum(FQTY)As FQTY,Sum(F' +
        'CUUnitQTY)As FCUUnitQTY,'
      
        'Sum(fqty1) As FQty1 ,Sum(fCUUnitqty1) As FCUUnitQty1 ,Sum(FAmoun' +
        't1) As FAmount1,Sum(fqty2) As FQty2 ,Sum(fCUUnitqty2) As FCUUnit' +
        'Qty2 ,Sum(FAmount2) As FAmount2,Sum(fqty3) As FQty3 ,Sum(fCUUnit' +
        'qty3) As FCUUnitQty3 ,Sum(FAmount3) As FAmount3,Sum(fqty4) As FQ' +
        'ty4 ,Sum(fCUUnitqty4) As FCUUnitQty4 ,Sum(FAmount4) As FAmount4,' +
        'Sum(fqty5) As FQty5 ,Sum(fCUUnitqty5) As FCUUnitQty5 ,Sum(FAmoun' +
        't5) As FAmount5,Sum(fqty6) As FQty6 ,Sum(fCUUnitqty6) As FCUUnit' +
        'Qty6 ,Sum(FAmount6) As FAmount6,Sum(fqty7) As FQty7 ,Sum(fCUUnit' +
        'qty7) As FCUUnitQty7 ,Sum(FAmount7) As FAmount7,1,Min(FPrice),ca' +
        'se Sum(FCUUnitQTY) when 0 then 0 else (Min(FPrice)*Sum(FQTY))/Su' +
        'm(FCUUnitQTY) end,Min(FPrice)*Sum(FQTY) '
      
        'From #HAPPEN2 t1 INNER JOIN t_ICItem t2 ON t1.FItemID=t2.FItemID' +
        ' '
      'GROUP By t1.FITEMID,t1.FBatchNo,t1.FStockID'
      ''
      'HAVING (SUM(FQTY)>=0) '
      'SET NOCOUNT ON'
      'CREATE TABLE #ItemLevel( '
      ' FNumber1 Varchar(355),'
      ' FName1 Varchar(355),'
      ' FNumber2 Varchar(355),'
      ' FName2 Varchar(355),'
      ' FItemID int,'
      ' FNumber Varchar(355))'
      ''
      ' INSERT INTO #ItemLevel SELECT  '
      
        ' CASE WHEN CHARINDEX('#39'.'#39',FFullNumber)-1= -1 or FLevel<2 THEN NUL' +
        'L ELSE SUBSTRING(FNumber, 1,CHARINDEX('#39'.'#39',FFullNumber)-1)  END, '
      ' '#39#39','
      
        ' CASE WHEN CHARINDEX('#39'.'#39',FFullNumber,CHARINDEX('#39'.'#39',FFullNumber)+' +
        '1)-1= -1 or FLevel<3 THEN NULL ELSE SUBSTRING(FNumber, 1,CHARIND' +
        'EX('#39'.'#39',FFullNumber,CHARINDEX('#39'.'#39',FFullNumber)+1)-1)  END, '
      ' '#39#39','
      ' FItemID,FNumber FROM t_Item'
      ' WHERE FItemClassID=4'
      
        ' AND FDetail=1 AND FNumber>=@number1 AND FNumber<=@number2  AND ' +
        'FItemID In (Select Distinct FItemID from #Happen )'
      ' UPDATE t0 SET t0.FName1=t1.FName,t0.FName2=t2.FName'
      
        '  FROM #ItemLevel t0 left join t_Item t1 On t0.FNumber1=t1.FNumb' +
        'er  AND t1.FItemClassID=4 AND t1.FDetail=0 '
      
        ' left join t_Item t2 On t0.FNumber2=t2.FNumber  AND t2.FItemClas' +
        'sID=4 AND t2.FDetail=0 '
      ''
      'CREATE TABLE #DATA('
      'FName1 Varchar(355) Null,'
      'FName2 Varchar(355) Null,'
      ''
      'FStockName Varchar(355) Null,     FNumber  Varchar(355) null,'
      '     FShortNumber  Varchar(355) null,'
      '     FName  Varchar(355) null,'
      '     FModel  Varchar(355) null,'
      '     FUnitName  Varchar(355) null,'
      '     FCUUnitName  Varchar(355) null,'
      '     FQtyDecimal smallint null, '
      '     FPriceDecimal smallint null, '
      '     FQty Decimal(28,10) Null, '
      '     FCUUnitQty Decimal(28,10) Null, '
      '     FPrice Decimal(28,10) NULL, '
      '     FCUPrice Decimal(28,10) Null, '
      '     FAmount Decimal(28,10) Null, '
      ' FQty1 Decimal(28,10),'
      ' FCUUnitQty1 Decimal(28,10),'
      ' FAmount1 Decimal(28,10),'
      ' FQty2 Decimal(28,10),'
      ' FCUUnitQty2 Decimal(28,10),'
      ' FAmount2 Decimal(28,10),'
      ' FQty3 Decimal(28,10),'
      ' FCUUnitQty3 Decimal(28,10),'
      ' FAmount3 Decimal(28,10),'
      ' FQty4 Decimal(28,10),'
      ' FCUUnitQty4 Decimal(28,10),'
      ' FAmount4 Decimal(28,10),'
      ' FQty5 Decimal(28,10),'
      ' FCUUnitQty5 Decimal(28,10),'
      ' FAmount5 Decimal(28,10),'
      ' FQty6 Decimal(28,10),'
      ' FCUUnitQty6 Decimal(28,10),'
      ' FAmount6 Decimal(28,10),'
      ' FQty7 Decimal(28,10),'
      ' FCUUnitQty7 Decimal(28,10),'
      ' FAmount7 Decimal(28,10),'
      '     FSumSort smallint not null Default(0),'
      'Flevel0 Decimal(10,3), '
      'Flevel1 Decimal(10,3), '
      'Flevel2 Decimal(10,3), '
      'Flevel3 Decimal(10,3), '
      'Flevel4 Decimal(10,3), '
      'Flevel5 Decimal(10,3), '
      'Flevel6 Decimal(10,3), '
      ''
      '     FID int IDENTITY)'
      'INSERT INTO #DATA '
      
        'SELECT tt1.FName1,tt1.FName2,t2.FName,t1.FNumber,'#39#39','#39#39','#39#39','#39#39','#39#39',' +
        'MAX(t1.FQtyDecimal),MAX(t1.FPriceDecimal),Sum(FQty),Sum(FCUUnitQ' +
        'ty), case Sum(FQty) when 0 then 0 else Sum(FAmount)/Sum(FQty) en' +
        'd,(CASE Sum(FCUUnitQty) WHEN 0 THEN 0 ELSE Sum(FAmount)/Sum(FCUU' +
        'nitQty) END), sum(FAmount), SUM(FQty1),'
      ' SUM(FCUUnitQty1),'
      ' SUM(FAmount1),'
      ' SUM(FQty2),'
      ' SUM(FCUUnitQty2),'
      ' SUM(FAmount2),'
      ' SUM(FQty3),'
      ' SUM(FCUUnitQty3),'
      ' SUM(FAmount3),'
      ' SUM(FQty4),'
      ' SUM(FCUUnitQty4),'
      ' SUM(FAmount4),'
      ' SUM(FQty5),'
      ' SUM(FCUUnitQty5),'
      ' SUM(FAmount5),'
      ' SUM(FQty6),'
      ' SUM(FCUUnitQty6),'
      ' SUM(FAmount6),'
      ' SUM(FQty7),'
      ' SUM(FCUUnitQty7),'
      ' SUM(FAmount7),'
      'CASE   WHEN   Grouping(tt1.FName1)=1 THEN 106'
      '  WHEN   Grouping(tt1.FName2)=1 THEN 107'
      
        '  WHEN   Grouping(t2.FName)=1 THEN  108  WHEN   Grouping(t1.FNum' +
        'ber)=1 THEN 109  ELSE   0 END '
      ',0 ,0 ,0 ,0 ,0 ,0 ,0  FROM #Happen v2'
      ' Inner Join t_ICItem t1 On v2.FItemID=t1.FItemID'
      ' Left Join t_Stock t2 On v2.FStockID=t2.FItemID'
      'Inner Join #ItemLevel tt1'
      ''
      ' On t1.FItemID=tt1.FItemID Where 1=1'
      ''
      ''
      '  Group By tt1.FName1,tt1.FName2,t2.FName,t1.FNumber '
      '  --WITH ROLLUP '
      '  Having Sum(FQty)>0  '
      ''
      
        ' Update t1 Set t1.FName=t2.FName,t1.FShortNumber=t2.FShortNumber' +
        ',t1.FModel=t2.FModel, t1.FUnitName=t3.FName,t1.FCUUnitName=t4.FN' +
        'ame ,t1.FQtyDecimal=t2.FQtyDecimal,t1.FPriceDecimal=t2.FPriceDec' +
        'imal  From #DATA t1,t_ICItem t2,t_MeasureUnit t3,t_MeasureUnit t' +
        '4  Where t1.FNumber=t2.FNumber  And t2.FUnitGroupID=t3.FUnitGrou' +
        'pID  And t2.FStoreUnitID=t4.FMeasureUnitID  And t3.FStandard=1'
      ''
      ''
      'SELECT IDENTITY(INT,1,1) AS FLevel0,FNAME1 '
      ' INTO #LEVEL0 FROM #DATA'
      'Where FSumSort=107'
      ''
      ''
      'UPDATE a3 SET a3.FLevel0=lv.FLevel0'
      'FROM #DATA a3'
      'INNER JOIN #LEVEL0 lv ON ISNULL(lv.FNAME1,0)=ISNULL(a3.FNAME1,0)'
      ''
      'Where  Not( FSumSort=106)'
      'DROP TABLE #LEVEL0 '
      
        ' Update  #DATA set FLevel0=(Select max(ISNULL(FLevel0,0))+1 From' +
        ' #DATA) Where 1=1 AND FSumSort=106'
      ''
      'SELECT IDENTITY(INT,1,1) AS FLevel1,FNAME1 , FNAME2 '
      ' INTO #LEVEL1 FROM #DATA'
      'Where FSumSort=108'
      ''
      ''
      'ORDER BY  FLEVEL0 '
      'UPDATE a3 SET a3.FLevel1=lv.FLevel1'
      'FROM #DATA a3'
      
        'INNER JOIN #LEVEL1 lv ON ISNULL(lv.FNAME1,0)=ISNULL(a3.FNAME1,0)' +
        ' AND ISNULL(lv.FNAME2,0)=ISNULL(a3.FNAME2,0)'
      ''
      'Where  Not( FSumSort=106)'
      'DROP TABLE #LEVEL1 '
      ''
      'SELECT IDENTITY(INT,1,1) AS FLevel2,FNAME1 ,FNAME2 , FSTOCKNAME'
      ' INTO #LEVEL2 FROM #DATA'
      'Where FSumSort=109'
      ''
      ''
      'ORDER BY  FLEVEL0 ,  FLEVEL1'
      'UPDATE a3 SET a3.FLevel2=lv.FLevel2'
      'FROM #DATA a3'
      
        'INNER JOIN #LEVEL2 lv ON ISNULL(lv.FNAME1,0)=ISNULL(a3.FNAME1,0)' +
        ' AND ISNULL(lv.FNAME2,0)=ISNULL(a3.FNAME2,0) AND ISNULL(lv.FSTOC' +
        'KNAME,0)=ISNULL(a3.FSTOCKNAME,0)'
      ''
      'Where  Not( FSumSort=106)'
      'DROP TABLE #LEVEL2 '
      ''
      ' Update #DATA Set FLevel0=FLevel0+0.9   Where FSumSort=107'
      ' Update #DATA Set FLevel1=FLevel1+0.9   Where FSumSort=108'
      ' Update #DATA Set FLevel2=FLevel2+0.9   Where FSumSort=109'
      
        'Update #Data Set  FName1=isnull(FName1,'#39#39')+'#39'('#23567#35745')'#39'  WHERE FSumSor' +
        't=107'
      
        'Update #Data Set  FName2=isnull(FName2,'#39#39')+'#39'('#23567#35745')'#39'  WHERE FSumSor' +
        't=108'
      
        'Update #Data Set FStockName=Rtrim(ISNULL(FStockName,'#39#39'))+'#39'('#23567#35745')'#39' ' +
        'WHERE FSumSort=109'
      'Update #Data Set FName1='#39#21512#35745#39' WHERE FSumSort=106'
      'Update #Data Set FSumSort=101   WHERE FSumSort=106'
      ''
      'SELECT * FROM #DATA  '
      ' Order By  FLevel0,  FLevel1,  FLevel2'
      'DROP TABLE #DATA  '
      'DROP TABLE #ItemLevel'
      ' Drop Table #Happen'
      ' Drop Table #Happen1'
      ' Drop Table #Happen2')
    Left = 88
    Top = 296
  end
end
