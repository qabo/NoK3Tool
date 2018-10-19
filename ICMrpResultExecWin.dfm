object frmICMrpResultExec: TfrmICMrpResultExec
  Left = 548
  Top = 296
  Width = 772
  Height = 407
  Caption = 'MRP'#35745#21010#35746#21333#19979#25512
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ToolBar1: TToolBar
    Left = 0
    Top = 0
    Width = 756
    Height = 40
    AutoSize = True
    ButtonHeight = 36
    ButtonWidth = 31
    Caption = 'ToolBar1'
    Images = DataModule1.ImageList1
    ShowCaptions = True
    TabOrder = 0
    object tbExport: TToolButton
      Left = 0
      Top = 2
      Caption = #23548#20986
      ImageIndex = 2
      OnClick = tbExportClick
    end
  end
  object ProgressBar1: TProgressBar
    Left = 16
    Top = 192
    Width = 721
    Height = 17
    TabOrder = 1
  end
  object mmoSql: TMemo
    Left = 312
    Top = 72
    Width = 185
    Height = 89
    Lines.Strings = (
      
        'select t1.FHeadSelfJ0550 '#35745#31639#32534#21495', t101.FNumber '#29289#26009#32534#30721', t101.FName '#29289#26009#21517 +
        #31216', t101a.FName '#37319#36141#21592', t101b.FName MC,'
      
        #9't1.FBillDate '#35745#21010#35746#21333#26085#26399', t1.FBillNo '#35745#21010#35746#21333#21495', t1.FPlanQty '#35745#21010#35746#21333#25968#37327', t2.'#37319 +
        #36141#30003#35831#21333#26085#26399','
      #9't2.'#37319#36141#30003#35831#21333#21495', t2.'#37319#36141#30003#35831#21333#20998#24405', t2.'#37319#36141#30003#35831#25968#37327','
      
        #9't2.'#37319#36141#35746#21333#26085#26399', t2.'#37319#36141#35746#21333#21495', t2.'#37319#36141#35746#21333#20998#24405', t2.'#37319#36141#35746#21333#25968#37327', t101.FFixLeadTime LT' +
        ', t101.FQtyMin MOQ, t101.FBatchAppendQty SPQ'
      'from ICMrpResult t1'
      'left join ('
      
        #9'select a.FDate '#37319#36141#30003#35831#21333#26085#26399', b.FPlanOrderInterID, a.FBillNo '#37319#36141#30003#35831#21333#21495', ' +
        'b.FEntryID '#37319#36141#30003#35831#21333#20998#24405', b.FQty '#37319#36141#30003#35831#25968#37327','
      #9#9'c.'#37319#36141#35746#21333#21495', c.'#37319#36141#35746#21333#20998#24405', c.'#37319#36141#35746#21333#26085#26399', c.'#37319#36141#35746#21333#25968#37327
      #9'from PORequest a'
      #9'inner join PORequestEntry b on a.FInterID=b.FInterID'
      #9'left join  ('
      
        #9#9'select o1.FDate '#37319#36141#35746#21333#26085#26399', o1.FBillNo '#37319#36141#35746#21333#21495', o2.FEntryID '#37319#36141#35746#21333#20998#24405', ' +
        'o2.FSourceInterId, '
      #9#9#9'o2.FSourceEntryID,'
      #9#9#9'o2.FQty '#37319#36141#35746#21333#25968#37327
      #9#9'from POOrder o1'
      #9#9'inner join POOrderEntry o2 on o1.FInterID=o2.FInterID'
      
        #9') c on b.FInterID=c.FSourceInterId and b.FEntryID=c.FSourceEntr' +
        'yID'
      ') t2 on t1.FInterID=t2.FPlanOrderInterID'
      'inner join t_ICItem t101 on t1.FItemID=t101.FItemID'
      'left join t_Emp t101a on t101.FOrderRector=t101a.FItemID'
      'left join t_Emp t101b on t101.F_102=t101b.FItemID'
      '/*'
      'where t1.FHeadSelfJ0550=1215 '
      ''
      '*/')
    TabOrder = 2
    Visible = False
    WordWrap = False
  end
  object leRunID: TLabeledEdit
    Left = 16
    Top = 72
    Width = 121
    Height = 21
    EditLabel.Width = 69
    EditLabel.Height = 13
    EditLabel.Caption = #35745#31639#32534#21495#65306'   '
    TabOrder = 3
  end
  object SaveDialog1: TSaveDialog
    Left = 208
    Top = 8
  end
  object ADOQuery1: TADOQuery
    Connection = DataModule1.gadoc
    Parameters = <>
    Left = 256
    Top = 8
  end
end
