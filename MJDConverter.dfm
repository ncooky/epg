object frmMJD: TfrmMJD
  Left = 586
  Top = 120
  Width = 299
  Height = 185
  Caption = 'frmMJD'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 22
    Height = 13
    Caption = 'Year'
  end
  object Label2: TLabel
    Left = 56
    Top = 8
    Width = 30
    Height = 13
    Caption = 'Month'
  end
  object Label3: TLabel
    Left = 96
    Top = 8
    Width = 23
    Height = 13
    Caption = 'Date'
  end
  object Label4: TLabel
    Left = 160
    Top = 8
    Width = 23
    Height = 13
    Caption = 'Hour'
  end
  object Label5: TLabel
    Left = 200
    Top = 8
    Width = 32
    Height = 13
    Caption = 'Minute'
  end
  object Label6: TLabel
    Left = 240
    Top = 8
    Width = 37
    Height = 13
    Caption = 'Second'
  end
  object LblRes: TLabel
    Left = 56
    Top = 96
    Width = 42
    Height = 13
    Caption = 'Result = '
  end
  object LblRes2: TLabel
    Left = 56
    Top = 120
    Width = 39
    Height = 13
    Caption = 'LblRes2'
  end
  object EdtYear: TEdit
    Left = 16
    Top = 32
    Width = 33
    Height = 21
    TabOrder = 0
  end
  object EdtMonth: TEdit
    Left = 56
    Top = 32
    Width = 33
    Height = 21
    TabOrder = 1
  end
  object EdtDate: TEdit
    Left = 96
    Top = 32
    Width = 33
    Height = 21
    TabOrder = 2
  end
  object EdtHour: TEdit
    Left = 160
    Top = 32
    Width = 33
    Height = 21
    TabOrder = 3
  end
  object EdtMin: TEdit
    Left = 200
    Top = 32
    Width = 33
    Height = 21
    TabOrder = 4
  end
  object EdtSec: TEdit
    Left = 240
    Top = 32
    Width = 33
    Height = 21
    TabOrder = 5
  end
  object Button1: TButton
    Left = 104
    Top = 64
    Width = 75
    Height = 25
    Caption = 'Calculate'
    TabOrder = 6
    OnClick = Button1Click
  end
end
