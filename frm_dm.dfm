object dm: Tdm
  Left = 612
  Top = 208
  Width = 443
  Height = 304
  Caption = 'dm'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object EPG_DB: TADOConnection
    Provider = 'MSDAORA.1'
    Left = 56
    Top = 32
  end
  object DDL: TADOQuery
    Connection = EPG_DB
    Parameters = <>
    Left = 88
    Top = 32
  end
  object dml: TADOQuery
    Connection = EPG_DB
    Parameters = <>
    Left = 120
    Top = 32
  end
  object EPG_Access_DB: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Project XL to DB' +
      '\DataBase\Longdesc.mdb;Persist Security Info=False'
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 56
    Top = 64
  end
  object AccDDL: TADOQuery
    Connection = EPG_Access_DB
    Parameters = <>
    Left = 88
    Top = 64
  end
  object Accdml: TADOQuery
    Connection = EPG_Access_DB
    Parameters = <>
    Left = 120
    Top = 64
  end
  object EPG_DB_2: TADOConnection
    Provider = 'MSDAORA.1'
    Left = 56
    Top = 96
  end
  object DDL2: TADOQuery
    Connection = EPG_DB_2
    Parameters = <>
    Left = 88
    Top = 96
  end
  object dml2: TADOQuery
    Connection = EPG_DB_2
    Parameters = <>
    Left = 120
    Top = 96
  end
  object dmlTanggal: TADOQuery
    Connection = EPG_DB_2
    Parameters = <>
    Left = 120
    Top = 136
  end
  object DDLTanggal: TADOQuery
    Connection = EPG_DB_2
    Parameters = <>
    Left = 88
    Top = 136
  end
  object DDLIDTable: TADOQuery
    Connection = EPG_DB
    Parameters = <>
    Left = 88
    Top = 168
  end
  object DDLDateTime: TADOQuery
    Connection = EPG_DB_2
    Parameters = <>
    Left = 120
    Top = 168
  end
  object DDLPush: TADOQuery
    Connection = EPG_DB
    Parameters = <>
    Left = 88
  end
end
