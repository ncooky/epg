object frmExEPG: TfrmExEPG
  Left = 409
  Top = 211
  Width = 297
  Height = 279
  Caption = 'Export EPG'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 2
    Top = 64
    Width = 287
    Height = 177
    Caption = 'Export EPG'
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 29
      Width = 39
      Height = 13
      Caption = 'Channel'
    end
    object Label2: TLabel
      Left = 8
      Top = 58
      Width = 48
      Height = 13
      Caption = 'Start Date'
    end
    object Label3: TLabel
      Left = 8
      Top = 88
      Width = 45
      Height = 13
      Caption = 'End Date'
    end
    object dtpAwal: TDateTimePicker
      Left = 96
      Top = 52
      Width = 186
      Height = 21
      Date = 39350.332592407410000000
      Time = 39350.332592407410000000
      TabOrder = 0
    end
    object dtpAkhir: TDateTimePicker
      Left = 96
      Top = 80
      Width = 186
      Height = 21
      Date = 39350.332694768520000000
      Time = 39350.332694768520000000
      TabOrder = 1
    end
    object cbExEPGChannel: TComboBox
      Left = 96
      Top = 24
      Width = 185
      Height = 21
      ItemHeight = 13
      TabOrder = 2
    end
  end
end
