object frmDBConfig: TfrmDBConfig
  Left = 840
  Top = 421
  Width = 384
  Height = 242
  Caption = 'frmDBConfig'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object leServer: TLabeledEdit
    Left = 106
    Top = 32
    Width = 175
    Height = 21
    EditLabel.Width = 51
    EditLabel.Height = 13
    EditLabel.Caption = #26381#21153#22120#65306' '
    LabelPosition = lpLeft
    TabOrder = 0
  end
  object leUser: TLabeledEdit
    Left = 106
    Top = 64
    Width = 175
    Height = 21
    EditLabel.Width = 39
    EditLabel.Height = 13
    EditLabel.Caption = #29992#25143#65306' '
    LabelPosition = lpLeft
    TabOrder = 1
  end
  object lePwd: TLabeledEdit
    Left = 106
    Top = 96
    Width = 175
    Height = 21
    EditLabel.Width = 39
    EditLabel.Height = 13
    EditLabel.Caption = #23494#30721#65306' '
    LabelPosition = lpLeft
    PasswordChar = '*'
    TabOrder = 2
  end
  object btnOk: TButton
    Left = 144
    Top = 144
    Width = 75
    Height = 25
    Caption = #30830#23450
    ModalResult = 1
    TabOrder = 3
  end
  object btnCancel: TButton
    Left = 232
    Top = 144
    Width = 75
    Height = 25
    Caption = #21462#28040
    ModalResult = 2
    TabOrder = 4
  end
end
