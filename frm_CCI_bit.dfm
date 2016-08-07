object frmCCIBits: TfrmCCIBits
  Left = 602
  Top = 301
  Hint = '1'
  BorderIcons = [biHelp]
  BorderStyle = bsSingle
  Caption = 'CCI Bits Input'
  ClientHeight = 92
  ClientWidth = 229
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  ShowHint = True
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 16
    Width = 55
    Height = 16
    Caption = 'Channel :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 19
    Top = 40
    Width = 52
    Height = 16
    Caption = 'CCI Bits :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblChannel: TLabel
    Left = 80
    Top = 19
    Width = 5
    Height = 13
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object edtCCI: TEdit
    Left = 80
    Top = 40
    Width = 25
    Height = 21
    MaxLength = 2
    TabOrder = 0
    OnKeyPress = edtCCIKeyPress
  end
  object Button1: TButton
    Left = 16
    Top = 64
    Width = 75
    Height = 25
    Caption = 'Save'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 144
    Top = 64
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 2
    OnClick = Button2Click
  end
end
