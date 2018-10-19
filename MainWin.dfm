object frmMain: TfrmMain
  Left = 787
  Top = 243
  Width = 994
  Height = 702
  Caption = 'frmMain'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object StatusBar1: TStatusBar
    Left = 0
    Top = 625
    Width = 978
    Height = 19
    Panels = <
      item
        Width = 250
      end
      item
        Width = 50
      end
      item
        Width = 50
      end>
  end
  object MainMenu1: TMainMenu
    Left = 136
    Top = 120
    object N22: TMenuItem
      Caption = #22522#30784#36164#26009
      object N23: TMenuItem
        Caption = #35745#21010#29289#26009#32500#25252
        OnClick = N23Click
      end
      object PC2: TMenuItem
        Caption = 'PC'#29289#26009#32500#25252
        OnClick = PC2Click
      end
    end
    object PC1: TMenuItem
      Caption = 'PC'
      object N1: TMenuItem
        Caption = #25968#25454#38598#25104
        OnClick = N1Click
      end
      object N7: TMenuItem
        Caption = #25968#25454#38598#25104#39033#30446#21512#24182
        OnClick = N7Click
      end
      object asdf1: TMenuItem
        Caption = #25968#25454#38598#25104#20998#26512
        OnClick = asdf1Click
      end
      object N3: TMenuItem
        Caption = #25968#25454#38598#25104#20998#26512'2'
        OnClick = N3Click
      end
      object N5: TMenuItem
        Caption = #25968#25454#38598#25104#20998#26512#32467#26524#26680#23545
        OnClick = N5Click
      end
      object N2: TMenuItem
        Caption = #28023#22806#20135#20986#27719#24635
        OnClick = N2Click
      end
      object SOP1: TMenuItem
        Caption = 'S&&OP'#36798#25104#29575
        OnClick = SOP1Click
      end
      object VS1: TMenuItem
        Caption = 'S&&OP'#35745#21010'VS'#23454#38469
        OnClick = VS1Click
      end
      object N4: TMenuItem
        Caption = #35201#36135#35745#21010
        OnClick = N4Click
      end
      object N11: TMenuItem
        Caption = #26412#22320#35201#36135#35745#21010
        OnClick = N11Click
      end
      object ODM1: TMenuItem
        Caption = 'ODM'#39033#30446#29983#20135#24322#24120#20998#26512'_'#36319#36827'_'#34917#36135#36861#36394#34920
        OnClick = ODM1Click
      end
      object SOPVSMPS1: TMenuItem
        Caption = 'S&OP VS MPS'
        OnClick = SOPVSMPS1Click
      end
      object SOP3: TMenuItem
        Caption = 'S&&OP'#36716#20135#21697#39044#27979#27169#26495
        OnClick = SOP3Click
      end
      object N13: TMenuItem
        Caption = #28023#22806#20849#29992#21253#26448
        OnClick = N13Click
      end
      object SAPtoSOP1: TMenuItem
        Caption = 'SAP to S&&OP'
        OnClick = SAPtoSOP1Click
      end
      object SOP5: TMenuItem
        Caption = 'S&&OP'#35745#21010#29256#26412#23545#27604
        OnClick = SOP5Click
      end
      object N17: TMenuItem
        Caption = #25104#21697#20837#24211#19982#24211#23384#25253#34920
        OnClick = N17Click
      end
      object S6201: TMenuItem
        Caption = #23548#20837'S620'#27169#26495#21512#24182
        OnClick = S6201Click
      end
      object test1: TMenuItem
        Caption = #38144#21806#35745#21010#27719#24635
        OnClick = test1Click
      end
      object Waterfall1: TMenuItem
        Caption = #38144#21806#35745#21010'Waterfall'
        Visible = False
        OnClick = Waterfall1Click
      end
      object SOP6: TMenuItem
        Caption = 'S&&OP'#25968#25454#27719#24635
        OnClick = SOP6Click
      end
    end
    object MC1: TMenuItem
      Caption = 'MC'
      object MRP1: TMenuItem
        Caption = 'MRP'#38656#27714#25163#24037#35843#25972#21512#24182
        OnClick = MRP1Click
      end
      object BS1: TMenuItem
        Caption = 'B/S '#29289#26009#38656#27714#26684#24335#36716#25442
        OnClick = BS1Click
      end
      object MPSWaterfall1: TMenuItem
        Caption = 'MPS Waterfall'
        OnClick = MPSWaterfall1Click
      end
      object SimpleWaterfall1: TMenuItem
        Caption = 'Simple Waterfall'
        OnClick = SimpleWaterfall1Click
      end
      object N8: TMenuItem
        Caption = #29942#39048#29289#26009#40784#22871#20998#26512
        Visible = False
      end
      object SOP4: TMenuItem
        Caption = 'S&&OP'#27169#25311#20998#26512
        OnClick = SOP4Click
      end
      object SOP2: TMenuItem
        Caption = 'S&OP'#35745#21010#27719#24635
        OnClick = SOP2Click
      end
      object N6: TMenuItem
        Caption = #25163#24037#29289#26009#38656#27714#35745#21010#65288'MC'#65289
        OnClick = N6Click
      end
      object N12: TMenuItem
        Caption = #39033#30446#24402#23646
        OnClick = N12Click
      end
      object nSAPBom2SBom: TMenuItem
        Caption = 'SAP Bom to SBom'
        OnClick = nSAPBom2SBomClick
      end
      object nLTP_CMS2MRPSim: TMenuItem
        Caption = 'LTP_CMS to MRP Sim'
        OnClick = nLTP_CMS2MRPSimClick
      end
      object MRP2: TMenuItem
        Caption = #32447#19979'CTB'
        OnClick = MRP2Click
      end
      object MRP3: TMenuItem
        Caption = #32447#19979'MRP'
        OnClick = MRP3Click
      end
      object mmiMRPAreaStockCheck: TMenuItem
        Caption = 'MRP'#21306#22495#19982#20179#24211#26816#26597
        OnClick = mmiMRPAreaStockCheckClick
      end
      object SAPBomMatrix1: TMenuItem
        Caption = 'SAP Bom Matrix'
        OnClick = SAPBomMatrix1Click
      end
      object MRPLog1: TMenuItem
        Caption = #25552#21462'MRP Log'
        OnClick = MRPLog1Click
      end
      object MRP21: TMenuItem
        Caption = #32447#19979'MRP 2'
        OnClick = MRP21Click
      end
      object ExcelBOM1: TMenuItem
        Caption = 'Excel_BOM'#21512#24182
        OnClick = ExcelBOM1Click
      end
      object miMrpSimDemand: TMenuItem
        Caption = 'MRP'#27169#25311#38656#27714#29983#25104
        OnClick = miMrpSimDemandClick
      end
      object ExcelBomSAPBom1: TMenuItem
        Caption = 'Excel Bom '#36716' SAP Bom'
        OnClick = ExcelBomSAPBom1Click
      end
      object WhereUse1: TMenuItem
        Caption = 'Where Use'
        OnClick = WhereUse1Click
      end
      object BOM1: TMenuItem
        Caption = 'BOM'#19982#37197#27604#26816#26597
        OnClick = BOM1Click
      end
      object N24: TMenuItem
        Caption = #20379#24212#19982#40784#22871
        OnClick = N24Click
      end
    end
    object N14: TMenuItem
      Caption = #36134#21153
      object N15: TMenuItem
        Caption = #25104#21697#25253#34920#21512#25104#65288#20998#39749#26063#39749#34013#65289
        OnClick = N15Click
      end
      object N20: TMenuItem
        Caption = #25104#21697#25253#34920#21512#25104#65288#24635#65289
        OnClick = N20Click
      end
      object N16: TMenuItem
        Caption = #20195#24037#21378#27599#26085#36134#21153#26680#23545
        OnClick = N16Click
      end
    end
    object N9: TMenuItem
      Caption = #36890#29992
      Visible = False
      object N10: TMenuItem
        Caption = #35774#32622
        OnClick = N10Click
      end
    end
    object SOPtoSAP1: TMenuItem
      Caption = 'S&&OP to SAP'
      OnClick = SOPtoSAP1Click
    end
    object SAPtoSOP2: TMenuItem
      Caption = 'SAP to S&&OP'
      OnClick = SAPtoSOP1Click
    end
    object N18: TMenuItem
      Caption = #25991#20214
      object N19: TMenuItem
        Caption = #20301#21495#20998#34892
        OnClick = N19Click
      end
      object N21: TMenuItem
        Caption = #25104#21697#20837#24211#24211#23384#20998#39749#34013#39749#26063
        OnClick = N21Click
      end
    end
  end
end
