object MaintainBox: TMaintainBox
  Left = 452
  Top = 396
  Width = 283
  Height = 93
  Caption = 'Maintain EPG Database'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 24
    Width = 57
    Height = 13
    Caption = 'Year events'
  end
  object edtYear1: TEdit
    Left = 80
    Top = 80
    Width = 65
    Height = 21
    TabOrder = 0
    Visible = False
  end
  object Button1: TButton
    Left = 168
    Top = 20
    Width = 89
    Height = 25
    Caption = 'Maintain NOW!'
    TabOrder = 1
    OnClick = Button1Click
  end
  object edtYear: TComboBox
    Left = 80
    Top = 24
    Width = 81
    Height = 21
    ItemHeight = 13
    TabOrder = 2
  end
end
