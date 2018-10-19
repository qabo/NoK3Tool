object frmFGDemandConfig: TfrmFGDemandConfig
  Left = 915
  Top = 388
  Width = 383
  Height = 233
  Caption = 'frmFGDemandConfig'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object leUpperLimit: TLabeledEdit
    Left = 72
    Top = 32
    Width = 169
    Height = 21
    EditLabel.Width = 39
    EditLabel.Height = 13
    EditLabel.Caption = #19978#38480#65306' '
    LabelPosition = lpLeft
    TabOrder = 0
  end
  object leLowerLimit: TLabeledEdit
    Left = 72
    Top = 88
    Width = 169
    Height = 21
    EditLabel.Width = 39
    EditLabel.Height = 13
    EditLabel.Caption = #19979#38480#65306' '
    LabelPosition = lpLeft
    TabOrder = 1
  end
  object btnCancel: TButton
    Left = 120
    Top = 144
    Width = 75
    Height = 25
    Caption = #30830#23450
    ModalResult = 1
    TabOrder = 2
  end
  object btnOk: TButton
    Left = 216
    Top = 144
    Width = 75
    Height = 25
    Caption = #21462#28040
    ModalResult = 2
    TabOrder = 3
  end
  object pnlUpperBrush: TPanel
    Left = 256
    Top = 24
    Width = 41
    Height = 41
    TabOrder = 4
    OnClick = pnlUpperBrushClick
  end
  object pnlLowerBrush: TPanel
    Left = 256
    Top = 80
    Width = 41
    Height = 41
    TabOrder = 5
    OnClick = pnlUpperBrushClick
  end
  object pnlUpperFont: TPanel
    Left = 312
    Top = 24
    Width = 41
    Height = 41
    Caption = 'A'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -32
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    OnClick = pnlUpperFontClick
  end
  object pnlLowerFont: TPanel
    Left = 312
    Top = 80
    Width = 41
    Height = 41
    Caption = 'A'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -32
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 7
    OnClick = pnlUpperFontClick
  end
  object ColorDialog1: TColorDialog
    Left = 232
  end
  object FontDialog1: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 304
  end
end
