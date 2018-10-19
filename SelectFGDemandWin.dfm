object frmSelectFGDemand: TfrmSelectFGDemand
  Left = 466
  Top = 210
  Width = 569
  Height = 571
  Caption = #36873#25321#35201#36135#35745#21010
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
  object lbWeeks: TListBox
    Left = 0
    Top = 57
    Width = 437
    Height = 476
    Align = alClient
    ItemHeight = 13
    MultiSelect = True
    TabOrder = 0
  end
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 553
    Height = 57
    Align = alTop
    TabOrder = 1
    object Label1: TLabel
      Left = 40
      Top = 20
      Width = 39
      Height = 13
      Caption = #39033#30446#65306' '
    end
    object Label2: TLabel
      Left = 272
      Top = 20
      Width = 36
      Height = 13
      Caption = #26085#26399#65306
    end
    object cbProjs: TComboBox
      Left = 80
      Top = 20
      Width = 145
      Height = 21
      ItemHeight = 13
      TabOrder = 0
      OnChange = cbProjsChange
    end
    object DateTimePicker1: TDateTimePicker
      Left = 320
      Top = 20
      Width = 153
      Height = 21
      Date = 42892.494052812500000000
      Format = 'yyyy-MM-dd'
      Time = 42892.494052812500000000
      TabOrder = 1
    end
  end
  object GroupBox2: TGroupBox
    Left = 437
    Top = 57
    Width = 116
    Height = 476
    Align = alRight
    TabOrder = 2
    object btnOK: TButton
      Left = 24
      Top = 24
      Width = 75
      Height = 25
      Caption = #30830#23450
      ModalResult = 1
      TabOrder = 0
    end
    object btnCancel: TButton
      Left = 24
      Top = 80
      Width = 75
      Height = 25
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
    end
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 232
    Top = 272
  end
end
