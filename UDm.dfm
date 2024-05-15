object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 141
  Width = 232
  object IbSql: TIBQuery
    Database = ibBanco
    Transaction = ibConexao
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 32
    Top = 48
  end
  object ibBanco: TIBDatabase
    DatabaseName = 'c:\digifarma\dados\digifarma6.fdb'
    Params.Strings = (
      'user_name=digifarma'
      'password=saci')
    LoginPrompt = False
    DefaultTransaction = ibConexao
    ServerType = 'IBServer'
    Left = 96
    Top = 48
  end
  object ibConexao: TIBTransaction
    DefaultDatabase = ibBanco
    Left = 168
    Top = 48
  end
end
