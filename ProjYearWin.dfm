object frmProjYear: TfrmProjYear
  Left = 929
  Top = 289
  Width = 457
  Height = 498
  Caption = #39033#30446#24180#24230
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
  TextHeight = 13
  object Memo1: TMemo
    Left = 0
    Top = 0
    Width = 320
    Height = 460
    Align = alClient
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 0
  end
  object GroupBox1: TGroupBox
    Left = 320
    Top = 0
    Width = 121
    Height = 460
    Align = alRight
    TabOrder = 1
    object btnOk: TButton
      Left = 24
      Top = 16
      Width = 75
      Height = 25
      Caption = #30830#23450
      TabOrder = 0
      OnClick = btnOkClick
    end
    object btnCancel: TButton
      Left = 24
      Top = 64
      Width = 75
      Height = 25
      Caption = #21462#28040
      TabOrder = 1
      OnClick = btnCancelClick
    end
  end
end
