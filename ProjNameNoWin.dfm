object frmProjNameNo: TfrmProjNameNo
  Left = 642
  Top = 292
  Width = 980
  Height = 659
  Caption = #39033#30446#32534#30721#21517#31216
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object Label1: TLabel
    Left = 30
    Top = 20
    Width = 33
    Height = 16
    Caption = #39749#26063'   '
  end
  object Label2: TLabel
    Left = 276
    Top = 20
    Width = 33
    Height = 16
    Caption = #39749#34013'   '
  end
  object Label3: TLabel
    Left = 492
    Top = 20
    Width = 69
    Height = 16
    Caption = #24573#30053#32534#30721#65306'   '
  end
  object Label4: TLabel
    Left = 699
    Top = 20
    Width = 105
    Height = 16
    Caption = #27719#24635#24573#30053#20851#38190#23383#65306'   '
  end
  object mmoOEM: TMemo
    Left = 20
    Top = 39
    Width = 227
    Height = 523
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 0
  end
  object btnCancel: TButton
    Left = 551
    Top = 581
    Width = 93
    Height = 31
    Caption = #21462#28040
    TabOrder = 1
    OnClick = btnCancelClick
  end
  object btnOk: TButton
    Left = 423
    Top = 581
    Width = 93
    Height = 31
    Caption = #30830#23450
    TabOrder = 2
    OnClick = btnOkClick
  end
  object mmoODM: TMemo
    Left = 266
    Top = 39
    Width = 208
    Height = 523
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 3
  end
  object mmoIgnoreNo: TMemo
    Left = 492
    Top = 39
    Width = 189
    Height = 523
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 4
  end
  object mmoIgnoreName4Sum: TMemo
    Left = 699
    Top = 39
    Width = 257
    Height = 523
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 5
  end
end
