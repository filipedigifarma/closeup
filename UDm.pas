unit UDm;

interface

uses
  System.SysUtils, System.Classes,System.INIFiles, ibx.IBDatabase, Data.DB, ibx.IBCustomDataSet,
  ibx.IBQuery, ibx.IBEvents, Vcl.Dialogs, System.Win.Registry, Winapi.Windows, Winapi.WinInet;

type
  TDM = class(TDataModule)
    IbSql: TIBQuery;
    ibBanco: TIBDatabase;
    ibConexao: TIBTransaction;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure p_gravaLog(msg:String;flag:integer);
    procedure Conexao;
    function CONECTADO:boolean;
    function f_Conectado(bMSG: Boolean = True): Boolean;
    function bd_ibQuery_init(var IbSql: TIBQuery): Boolean;

  end;

var
  DM: TDM;
  IsGlobalOffline: Boolean;

implementation

{ %CLASSGROUP 'Vcl.Controls.TControl' }

{$R *.dfm}

function TDM.CONECTADO:boolean;
var
  estado: Dword;
begin
  result:=false;

  if not InternetGetConnectedState(@estado, 0) then
  IsGlobalOffline := False
  else
  begin
    if (estado and INTERNET_CONNECTION_LAN <> 0) OR
      (estado and INTERNET_CONNECTION_MODEM <> 0) or
      (estado and INTERNET_CONNECTION_PROXY <> 0) then
      IsGlobalOffline := True;
      result:=true;
  end;

  if (IsGlobalOffline=false) then
  begin
//    if MessageDlg('Atenção você está sem internet, isto pode ocasionar mal funcionamento deste programa '
//    + 'deseja tentar uma nova conexão?',mtWarning,[mbYes,mbNo],0)=IDYES then
//    begin
//      Sleep(10000);
//      CONECTADO;
//    end
//    else
    result:=false;
  end;
end;

function TDM.bd_ibQuery_init(var IbSql: TIBQuery): Boolean;
begin

  Result := False;

  try
    if (f_Conectado = False) then
    begin
      Exit;
    end;
    IbSql := TIBQuery.Create(nil);
    IbSql.Database := ibBanco;
    IbSql.Transaction := ibConexao;
  except
    Result := False;
  end;

  Result := True;

end;

function TDM.f_Conectado(bMSG: Boolean = True): Boolean;
var
  bFlag: Boolean;
Begin

  bFlag := False;

  try
    ibBanco.Connected:=true;
    bFlag := ibBanco.TestConnected;
  except
    Result := False;
  end;

  try
    if (bFlag = true) then
    begin

      try
        ibBanco.ForceClose;
      finally

      end;

      ibBanco.Connected := True;
      ibConexao.Active := True;
      Result := True;
    end;
  except

    On E: Exception do
    Begin

//      if (bMSG = True) then
//      begin
//        MessageDlg
//          ('Erro na tentativa de abrir o banco de dados. Entre em contato com o suporte técnico Digifarma.'
//          + #13#10 + E.Message, mtError, [mbOK], 0);
//      end;
      Result := False;
    End;
  end;


//  Result := True;

End;

procedure TDM.p_gravaLog(msg:String;flag:integer);
var
Txt:TextFile;
mStr:TMemoryStream;
slLog,slLog2:TStringList;
begin

  try
    if not DirectoryExists('c:\digifarma\Closeup\Log') then
    ForceDirectories('c:\digifarma\Closeup\Log');


    if not (FileExists('c:\digifarma\Closeup\Log\Closeup_Log.txt')) then
    begin
      AssignFile(txt,'c:\digifarma\Closeup\Log\Closeup_Log.txt');
      RewRite(txt);
      CloseFile(txt);
    end;



    slLog:=TStringList.Create;
    slLog2:=TStringList.Create;

    mStr:=TMemoryStream.Create;
    mStr.LoadFromFile('C:\Digifarma\Closeup\Log\Closeup_Log.txt');
    if((mStr.Size/1024)>300)then
    begin
      freeandnil(mStr);
      slLog.LoadFromFile('C:\Digifarma\Closeup\Log\Closeup_Log.txt');
      slLog.Clear;
      slLog.Insert(0,'O log foi apagado automaticamente ao atingir o tamanho maximo');
      slLog.SaveToFile('C:\Digifarma\Closeup\Log\Closeup_Log.txt');
      slLog.Clear;
    end;


    if (flag=0) then
     slLog2.Add('---------------------- Integração Digifarma - Closeup ------------------------');

     slLog2.Add(FormatDateTime('dd/mm/yyyy hh:mm:ss',now)+ ' - '+msg);

     if (flag=1) then
     slLog2.Add('------------------------------------------------------------------------------');

     slLog.LoadFromFile('C:\Digifarma\Closeup\Log\Closeup_Log.txt');
     slLog.Insert(0,slLog2.Text);

     slLog.SaveToFile('C:\Digifarma\Closeup\Log\Closeup_Log.txt');

     if assigned(slLog) then
     freeandnil(slLog);

     if assigned(slLog2) then
     freeandnil(slLog2);

     if assigned(mStr) then
     freeandnil(mStr);
  except
   on ex:exception do
   begin
         exit;
   end;


  end;

end;


procedure TDM.Conexao;
Var
  Registro: TRegistry;
  PATH_BANCO: string;
begin

  try

    Registro := TRegistry.Create;
    Registro.RootKey := HKEY_CURRENT_USER;

    if Registro.OpenKey('DIGIFARMA', False) then
    begin
      if Registro.ValueExists('Database') then
        PATH_BANCO := Registro.ReadString('Database')
      else
      begin
        MessageDlg
          ('Banco de dados Digifarma não configurado. Entre em contato com o suporte técnico Digifarma.',
          mtError, [mbOK], 0);
        Abort;
      end;
      Registro.CloseKey;
      Registro.Free;
    end;
//
    ibBanco.DatabaseName := PATH_BANCO;
//    ibBanco.DatabaseName := 'c:\digifarma\dados\digifarma6.fdb';
  except
    MessageDlg('Um erro ocorreu na tentativa de ler as configrações do sistema. Entre em contato com o suporte técnico Digifarma.',mtError, [mbOK], 0);
    Abort;
  end;

  try
    if (ibBanco.Connected = True) then
    Begin
      ibBanco.Connected := False;
    End;
  except

  end;

  try
    if (ibConexao.Active = True) then
    Begin
      ibConexao.Active := False;
    End;
  except

  end;

  if (f_Conectado = False) then
  begin
    p_gravaLog('Banco de dados Offline aguardando 2 minutos', -1);
    Sleep(120000);
    Conexao;
  end;

end;

procedure TDM.DataModuleCreate(Sender: TObject);
var
REGISTRO:TRegistry;
Arquivo: TINIFile;
begin
  p_gravaLog('Iniciando Rotina Closeup', 1);
  p_gravaLog('Iniciando Conexão com Banco de dados', -1);

  Conexao;

  p_gravaLog('Acessando registro do Windows', -1);
  REGISTRO:=TRegistry.Create(KEY_WRITE or $0100);
  registro.RootKey:=HKEY_LOCAL_MACHINE;
  if(registro.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Run',false))then
  begin
    if (registro.KeyExists('Integração Digifarma-CloseUp')=false)then
    try
      registro.WriteString('Integração Digifarma-CloseUp','c:\Digifarma\CloseUp\CloseUp.exe');
    except
      on ex:exception do
      MessageDlg('Erro ao inserir no Registro do Windows',mterror,[mbok],0);
    end;
  end;

end;

end.
