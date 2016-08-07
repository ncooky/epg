object frmEditCaEvent: TfrmEditCaEvent
  Left = 479
  Top = 289
  BorderStyle = bsNone
  Caption = 'Edit CA Package'
  ClientHeight = 219
  ClientWidth = 441
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object ngEditCAEvent: TNextGrid
    Left = 0
    Top = 0
    Width = 441
    Height = 185
    Options = [goGrid, goHeader]
    TabOrder = 0
    TabStop = True
    OnMouseUp = ngEditCAEventMouseUp
    OnSelectCell = ngEditCAEventSelectCell
    object NxTextColumn1: TNxTextColumn
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Header.Caption = 'No.'
      Position = 0
      SortType = stAlphabetic
    end
    object NxTextColumn2: TNxTextColumn
      DefaultWidth = 190
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Header.Caption = 'Channel'
      Position = 1
      SortType = stAlphabetic
      Width = 190
    end
    object NxTextColumn3: TNxTextColumn
      DefaultWidth = 150
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Header.Caption = 'CA Package'
      Position = 2
      SortType = stAlphabetic
      Width = 150
    end
    object NxTextColumn4: TNxTextColumn
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Position = 3
      SortType = stAlphabetic
      Visible = False
    end
    object NxTextColumn5: TNxTextColumn
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Position = 4
      SortType = stAlphabetic
      Visible = False
    end
    object NxTextColumn6: TNxTextColumn
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Position = 5
      SortType = stAlphabetic
      Visible = False
    end
  end
  object Button1: TButton
    Left = 352
    Top = 190
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 8
    Top = 192
    Width = 75
    Height = 25
    Caption = 'Save'
    TabOrder = 2
    Visible = False
  end
  object ppmEditSch: TPopupMenu
    Left = 176
    Top = 184
    object InsertRow1: TMenuItem
      Caption = 'Insert Row'
      OnClick = InsertRow1Click
    end
    object Delete1: TMenuItem
      Caption = 'Delete Row'
      OnClick = Delete1Click
    end
    object Delete2: TMenuItem
      Caption = 'Delete'
    end
  end
end
