unit uFuncoes;

interface

uses  SysUtils,System.Classes,winapi.windows,Winapi.ShellAPI,Winapi.ShlObj, Registry,
      Vcl.Forms,IdHashMessageDigest, Vcl.StdCtrls
  , Winapi.TlHelp32
  , Winapi.Messages
  , System.StrUtils
  , System.DateUtils
  , Vcl.Dialogs
  , IdSNTP
  , System.Generics.Collections
  , Winapi.WinInet
  , IdIcmpClient
  , XSBuiltIns
  , ActiveX
  , ComObj
  , Variants
  , JclSysInfo
  , EncdDecd
  , System.RegularExpressions
  , Winapi.WinSock
  , IdHTTP
  , Json
  , IBX.IBQuery
  , IBX.IBDatabase ;

 type TCheckInternet = (ciNO_INTERNET_CONNECTION, ciINTERNET_CONNECTION_MODEM, ciINTERNET_CONNECTION_LAN, ciINTERNET_CONNECTION_PROXY, ciINTERNET_CONNECTION_MODEM_BUSY);

 type
    TCampos = class
  private
      FCampos: TStringList;
      FTexto: String;
      procedure GravaCampo(NomeCampo: String; Value: String);
      function LeCampo(NomeCampo: String): String;
      function LeTexto: String;
      procedure GravaTexto(const Texto: String);
  public
    constructor Create;virtual;
    Property Campo[NomeCampo: String] : String read LeCampo write GravaCampo;
    property Texto: String read LeTexto write GravaTexto;
  end;

 type
    TChamaDLLErros = class
      bErro: Boolean;
      bArquivoExiste: Boolean;
      bMD5Executando: Boolean;
      bMD5Confere: Boolean;
      bCarregaDLL: Boolean;
      bChamaFuncao: Boolean;
      bRetornoPonteiro: Boolean;
  end;

 type
    TBina = class
      PORTA_ID: Integer;
      NOME: String;
      DESCRICAO: string;
  end;

type
  PTOKEN_USER = ^TOKEN_USER;
  _TOKEN_USER = record
    User: TSidAndAttributes;
  end;
  TOKEN_USER = _TOKEN_USER;

type
 ThwdInfo = class
   OS_NAME:String;
   OS_ARC:String;
   OS_VERSAO:String;
   OS_DATINS:String;
   PRODUTOID:String;
   MACHINEID:String;
   CPUID:String;
   CPUNAME:String;
   MEMORIA:string;
   MOBO:string;
   BIOS:string;
 end;

function iif(TesteLogico: Boolean; Verdadeiro: string; Falso: String=''): string;overload;
function iif(TesteLogico: Boolean; Verdadeiro: integer; Falso: integer): integer;overload;

function GetWMIObject(const objectName: String): IDispatch;

function net_CheckInternet : TCheckInternet;
Function net_GetHost_Name():String;
Function net_GetIPAddress():String;

function b64_Encode64(const s_xml: string): AnsiString;
Function b64_DecBase64(Value: String): String;


function date_utc_to_datetime(utc_date_time: String) : TDateTime;
function date_mysql_to_date(date_time: String) : TDate;
function date_mysql_to_datetime(date_time: String) : TDateTime;
function date_ntp: TDateTime;

procedure NetStat(var Conexao: TStringList; Porta: Integer=0);

function num_ComboBox_Index(Str: String; Obj: TComboBox; bItemData: Boolean; bLike: Boolean = False) : Integer;
function num_FileSize(fileName : wideString) : Int64;

function bd_Caminho_BD_Digifarma: String;

function hinfo_GetBios: string;
function hinfo_GetMachineGuid: string;
function hinfo_GetMoboSerial: string;
function hinfo_GetProductID: string;
function hinfo_guid:string; overload;
function hinfo_guid(var oGuid: ThwdInfo):String;overload;

function hwnd_chamaDLL(arquivo: string; Metodo: string; var oErro: TChamaDLLErros; var DLLHandle: THandle; bMD5: Boolean = False): Pointer;


function lnk_DeletaLink(Habilitados: TStringList): Boolean;
function lnk_BuscaLink(lpLink: String; Tipo: Integer = 0): String;
function App_RegisterDigifarma(AppName: String; IsDLL: Boolean = False):Boolean;

function str_CutStr(Var CutString: String; Start, Len: Integer): string;
function str_Bina: string;
function str_BrowseDialog(const Title: string; const Flag: integer): string;
Function str_Criptografia(Action, Src: String): String;
Function str_Data_Hora_NTP: String;
function str_FileLastModified(const TheFile: string): string;
function str_FormatByteSize(const bytes: Longint): string;
function str_GetDosOutput(CommandLine: string; Work: string = 'C:\'): string;
function str_GetExeNameByHandle(WindowHandle: THandle): String;
function str_GetFileVer(const FileName: string): string;
function str_GetTempDir: string; stdcall;
function str_GetWinDir: string; stdcall;
function str_GetSystemDir: string; stdcall;
function str_HexToString(H: String): String;
Function str_Html(Texto :String; bAcentos: Boolean = True): String;
//function str_HtmlChars(Texto: String): String;
function str_IdentifyingNumber_wmi: String;
function str_Mid_Delimitado(Campos: String; Delimiter: Char;Indice:Integer):string;
function str_NomeMes(mes:Integer):string;
function str_SerialPlacaMae:string;
function str_uuid:string;
function str_uuid_wmi:string;
function str_Mask_Telefone(const texto: string): string;
function str_MD5(const texto: string): string;
function str_MD5_File(const fileName : string) : string;
function str_PathLinux(Texto: String): String;
function str_Preenche_Espaco(valor:string;tamanho:Integer):string;
function str_Preenche_Zero(valor:string;tamanho:Integer):string;
function str_Preenche_caractere(valor:string;caractere:string;tamanho:Integer):string;
function str_Propriedade_Arquivo(Arquivo,Propriedade: String): String;
function str_Proxy: string;
function str_Quote(Texto: String): String;
function str_Retira_Acentos(const Texto: String): String;
function str_SaveToFile(Const Texto:String; Const Arquivo:String):String;
function str_SoNumero(Const Texto:String):String;
function num_SoNumero(Const Texto:String):Integer;
function str_SoLetras(Const Texto:String):String;
function str_StringToHex(input: string):string;
function str_GetCurrentUser: string;

function stl_Gera_Lista(Input: string; const Delimiter: Char; var Lista: TStringList): Boolean;
function stl_Split(Input: string; const Delimiter: Char; var Lista: TStringList): Boolean;
function stl_lista_processos(var stList:TStringList):boolean;
function stl_Localizar_Arquivo(Local,extensao: string; bEXT: Boolean = True):TStringList;
function stl_Localizar_Arquivo_Dir(Local,extensao: string):TStringList;
function stl_Localizar_Diretorio(Local: string):TStringList;
function stl_SerialPort(var Portas: TStringList) : Boolean;

function wmi_antivirus_name:string;
function wmi_process_xml:string;
function wmi_smart_card(var s_status, s_leitora, s_servico: String): Boolean;
function wmi_uninstall(programa: string):string;
function wmi_windows_version:string;

function xFindStrinList(Lista: TStringList; Item: string; var Index: Integer): Boolean;
function xChecaInternet():Boolean;
function xCopyFiles(FromDir: string; Destination: string): Boolean;
function xDeleteDir(dir: string): Boolean;
function xDeleteFiles(FileName:string; MandarPraLixeira: Boolean = False): Boolean;
function xFileExecute(ahWnd: Cardinal; const aFileName, aParams, aStartDir: string; aShowCmd: Integer; aWait: Boolean): Integer;
function xFindWindowLike ( const Titulo: string; var completo: string ): Integer;
Function xIsDate ( const DateString: string; formato: String = 'DD/MM/YYYY'): boolean;
function xIsNumeric ( const NumericString: string ): boolean;
function xKillDll(aDllName: string): Boolean;
function xRunAsAdmin(hWnd: HWND; filename: string; Parameters: string): Boolean;
function xSaveToFile(const Texto, Caminho: String; append: Boolean = False):boolean;
function xValidarChaveNFe(const ChaveNFe: string):boolean;
function xPing(IP: String): boolean;
function xPingICMP( host: String): Boolean;
function xProcessoExiste(ExeFileName: string): boolean;
function xVersao(Arquivo: string; Var VMaior, VMenor, VRelease, VCompilacao: String): Boolean;
function xValidaEAN(CodBar: string): Boolean;
function xValidaCNPJ(num: string): boolean;
function xValidaCPF(cpf:string):Boolean;

procedure xFileEnDeCrypto(Arquivo_Origem, Arquivo_Saida: string; Chave: Word);
procedure xKillProcess;
Procedure xMemoryTrim;

implementation

uses uOleVariantEnum, uMySql;

type   //CLASSE UTILIZADA PELA FUNCAO QUE RETORNA A VERSAO DO ARQUIVO (str_GetFileVer)
  TVerInfo = packed record
    vMajor, vMinor, vRelease, vBuild: word;
  end;


constructor TCampos.Create;
begin
  FTexto := '';
  FCampos := TStringList.Create;
end;

function TCampos.LeCampo(NomeCampo: String): String;
Begin
  Result := FCampos.Values[NomeCampo];
End;

procedure TCampos.GravaCampo(NomeCampo: String; Value: String);
Begin

  if ( Trim(NomeCampo) ='') then
  Begin
    MessageDlg('Nome de campo inválido!', TMsgDlgType.mtError,[mbOK],0);
    Exit;
  End;

  if ((Value = EmptyStr) or (Value = '')) then
  Begin
    FCampos.Values[NomeCampo] := ' ';
  End
  else
  Begin
    FCampos.Values[NomeCampo] := Value;
  End;
  FTexto := FCampos.Text;
End;

function TCampos.LeTexto: String;
Begin
  Result := FTexto;
End;

procedure TCampos.GravaTexto(Const Texto: String);
Begin
  FTexto := Texto;
  FCampos := TStringList.Create;
  FCampos.Delimiter :=';';
  FCampos.StrictDelimiter := True;
  FCampos.DelimitedText:=FTexto;
End;

function iif(TesteLogico: Boolean; Verdadeiro: string; Falso: String = ''): string; overload;
begin
  if (TesteLogico=False) then
    Result := Falso
  else
    Result := Verdadeiro;
end;

function iif(TesteLogico: Boolean; Verdadeiro: integer; Falso: integer): integer; overload;
begin
  if (TesteLogico=False) then
    Result := Falso
  else
    Result := Verdadeiro;
end;

function GetUserAndDomainFromPID(ProcessId: DWORD;  var User, Domain: string): Boolean;
var
  hToken: THandle;
  cbBuf: Cardinal;
  ptiUser: PTOKEN_USER;
  snu: SID_NAME_USE;
  ProcessHandle: THandle;
  UserSize, DomainSize: DWORD;
  bSuccess: Boolean;
begin
  Result := False;
  ProcessHandle := OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId);
  if ProcessHandle <> 0 then
  begin
  //  EnableProcessPrivilege(ProcessHandle, 'SeSecurityPrivilege', True);
    if OpenProcessToken(ProcessHandle, TOKEN_QUERY, hToken) then
    begin
      bSuccess := GetTokenInformation(hToken, TokenUser, nil, 0, cbBuf);
      ptiUser  := nil;
      while (not bSuccess) and (GetLastError = ERROR_INSUFFICIENT_BUFFER) do
      begin
        ReallocMem(ptiUser, cbBuf);
        bSuccess := GetTokenInformation(hToken, TokenUser, ptiUser, cbBuf, cbBuf);
      end;
      CloseHandle(hToken);

      if not bSuccess then
      begin
        Exit;
      end;

      UserSize := 0;
      DomainSize := 0;
      LookupAccountSid(nil, ptiUser.User.Sid, nil, UserSize, nil, DomainSize, snu);
      if (UserSize <> 0) and (DomainSize <> 0) then
      begin
        SetLength(User, UserSize);
        SetLength(Domain, DomainSize);
        if LookupAccountSid(nil, ptiUser.User.Sid, PChar(User), UserSize,
          PChar(Domain), DomainSize, snu) then
        begin
          Result := True;
          User := StrPas(PChar(User));
          Domain := StrPas(PChar(Domain));
        end;
      end;

      if bSuccess then
      begin
        FreeMem(ptiUser);
      end;
    end;
    CloseHandle(ProcessHandle);
  end;
end;


function GetWMIObject(const objectName: String): IDispatch;
var
  chEaten: Integer;
  BindCtx: IBindCtx;
  Moniker: IMoniker;
begin
  OleCheck(CreateBindCtx(0, bindCtx));
  OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), chEaten, Moniker));
  OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));
end;

{ RETORNA CONVERSÃO DE DATA  }
{ Autor : Rafael                                                                  }
{ Data : 16/07/2014                                                               }
function date_mysql_to_date(date_time: String) : TDate;
  var l_ano, l_mes, l_dia, l_hora: string;
begin

  if (Length(date_time) =10) then
  Begin
    l_ano := LeftStr(date_time,4);
    l_mes := MidStr(date_time,5,2);
    l_dia := RightStr(date_time,2);
    Result := StrToDate(l_dia + '/' + l_mes + '/' + l_ano);
  End;

  if (Length(date_time) =19) then
  Begin
    l_ano := LeftStr(date_time,4);
    l_mes := MidStr(date_time,6,2);
    l_dia := MidStr(date_time,9,2);
    l_hora := RightStr(date_time,9);
    Result := StrToDateTime(l_dia + '/' + l_mes + '/' + l_ano);
  End;

end;

function date_mysql_to_datetime(date_time: String) : TDateTime;
  var l_ano, l_mes, l_dia, l_hora: string;
begin

  if (Length(date_time) =10) then
  Begin
    l_ano := LeftStr(date_time,4);
    l_mes := MidStr(date_time,5,2);
    l_dia := RightStr(date_time,2);
    Result := StrToDate(l_dia + '/' + l_mes + '/' + l_ano + ' 00:00:00');
  End;

  if (Length(date_time) =19) then
  Begin
    l_ano := LeftStr(date_time,4);
    l_mes := MidStr(date_time,6,2);
    l_dia := MidStr(date_time,9,2);
    l_hora := RightStr(date_time,9);
    Result := StrToDateTime(l_dia + '/' + l_mes + '/' + l_ano + ' ' + l_hora);
  End;

end;

{ RETORNA CONVERSÃO DE DATA  }
{ Autor : Rafael                                                                  }
{ Data : 16/07/2014                                                               }
function date_utc_to_datetime(utc_date_time: String) : TDateTime;
var xsData: TXSDateTime;
    input, output:  string;
    date: TDateTime;

begin
  try
    xsData := TXSDateTime.Create;
    xsData.XSToNative(utc_date_time);
    date := xsData.AsDateTime;
    date := IncHour(date,3);
    Result := date;
  except
    Result := Unassigned;
  end;

end;

  //FUNCAO UTILIZADA PELA FUNCAO QUE RETORNA A VERSAO DO ARQUIVO  (str_GetFileVer)
function GetFileVerNumbers(const FileName: string): TVerInfo;
var
  len, dummy: cardinal;
  verdata: pointer;
  verstruct: pointer;
const
  InvalidVersion: TVerInfo = (vMajor: 0; vMinor: 0; vRelease: 0; vBuild: 0);
begin
  len := GetFileVersionInfoSize(PWideChar(FileName), dummy);
  if len = 0 then
    Exit(InvalidVersion);
  GetMem(verdata, len);
  try
    GetFileVersionInfo(PWideChar(FileName), 0, len, verdata);
    VerQueryValue(verdata, '\', verstruct, dummy);
    result.vMajor := HiWord(TVSFixedFileInfo(verstruct^).dwFileVersionMS);
    result.vMinor := LoWord(TVSFixedFileInfo(verstruct^).dwFileVersionMS);
    result.vRelease := HiWord(TVSFixedFileInfo(verstruct^).dwFileVersionLS);
    result.vBuild := LoWord(TVSFixedFileInfo(verstruct^).dwFileVersionLS);
  finally
    FreeMem(verdata);
  end;
end;

{ RETORNA O CAMINHO DO BANCO DE DADOS DO DIGIFARMA QUE CONSTA NO REGISTRO DO WIN  }
{ Autor : Rafael                                                                  }
{ Data : 26/11/2012                                                               }
function bd_Caminho_BD_Digifarma: String;
  var Registro: TRegistry;
begin

  Result := '';

  try
    Registro := TRegistry.Create(KEY_READ or $0100);
    Registro.RootKey:=HKEY_CURRENT_USER;
    // Somente abre se a chave existir
    if Registro.OpenKey ('DIGIFARMA', False) then
    begin
      // Envia as informações ao form, vendo se os valores existem, primeiramente...
      if Registro.ValueExists ('Database') then
        Result := Registro.ReadString('Database');
      // Fecha a chave e o objeto
      Registro.CloseKey;
      Registro.Free;
    end;
  except
    Result := 'c:\digifarma\dados\Digifarma6.fdb';
  end;

end;

{ CHECA CONEXAO DO WINDOWS COM A INTERNET                                         }
{ Autor : Rafael                                                                  }
{ Data : 29/07/2014                                                               }

function net_CheckInternet : TCheckInternet;
  var
     origin : cardinal;
begin
   result := TCheckInternet(InternetGetConnectedState(@origin,0));
end;

{ DECODIFICA UMA STRING EM BASE 64                                                }
{ Autor : Rafael                                                                  }
{ Data : 29/07/2014                                                               }

Function b64_DecBase64(Value: String): String;
  Var tamanho: Integer;
      Buff: TBytes;
      s_arq: TBytesStream;
      Enco: TEncoding;
Begin
  Result := '';
  tamanho :=0;
  Enco := nil;

  Try
    s_arq := TBytesStream.Create(decodeBase64(Value));
    tamanho := s_arq.Size - s_arq.Position;
    SetLength(Buff, tamanho);
    s_arq.ReadBuffer(Pointer(Buff)^, tamanho);
    tamanho := TEncoding.GetBufferEncoding(Buff, Enco);
    Result := Enco.GetString(s_arq.Bytes, tamanho, Length(s_arq.Bytes) - tamanho);
    Buff := Nil;
    Enco := nil;
    FreeAndNil(s_arq);
  except
    Buff := Nil;
    Enco := nil;
    FreeAndNil(s_arq);
    Result := '';
  End;
End;

{ CODIFICA UMA STRING EM BASE 64                                                  }
{ Autor : Rafael                                                                  }
{ Data : 29/07/2014                                                               }
function b64_Encode64(const s_xml: string): AnsiString;
var
  stream: TMemoryStream;
  XML: TStringStream;
begin
  stream := TMemoryStream.Create;
  XML := TStringStream.Create(s_xml);
  try
    stream.LoadFromStream(xml);
    result := EncodeBase64(stream.Memory, stream.Size);
  finally
    stream.Free;
  end;
end;

{ PEGA O IP DA REDE LOCAL DO COMPUTADOR                                           }
{ Autor : Rafael                                                                  }
{ Data : 29/07/2014                                                               }

Function net_GetIPAddress():String;
type
  pu_long = ^u_long;
var
  varTWSAData : TWSAData;
  varPHostEnt : PHostEnt;
  varTInAddr : TInAddr;
  namebuf : array[0..100] of AnsiChar;
begin
  If WSAStartup($101,varTWSAData) <> 0 Then
    Result := '0.0.0.0'
  Else
  Begin
    gethostname(namebuf,sizeof(namebuf));
    varPHostEnt := gethostbyname(namebuf);
    varTInAddr.S_addr := u_long(pu_long(varPHostEnt^.h_addr_list^)^);
    Result := inet_ntoa(varTInAddr);
  End;
  WSACleanup;
end;

{ PEGA O HOST NAME DO COMPUTADOR                                                  }
{ Autor : Rafael                                                                  }
{ Data : 29/07/2014                                                               }

Function net_GetHost_Name():String;
type
  pu_long = ^u_long;
var
  varTWSAData : TWSAData;
  varPHostEnt : PHostEnt;
  varTInAddr : TInAddr;
  namebuf : array[0..100] of AnsiChar;
begin
  If WSAStartup($101,varTWSAData) <> 0 Then
    Result := '0.0.0.0'
  Else
  Begin
    gethostname(namebuf,sizeof(namebuf));
    varPHostEnt := gethostbyname(namebuf);
    Result := varPHostEnt.h_name;
  End;
  WSACleanup;
end;



function hwnd_chamaDLL(arquivo: string; Metodo: string; var oErro: TChamaDLLErros; var DLLHandle: THandle; bMD5: Boolean = False): Pointer;
  var Retorno: ^Integer;
      DataXML: TDataXML;
begin

  oErro := TChamaDLLErros.Create;


  oErro.bErro := True;

  if (bMD5 = True) then
  Begin
    oErro.bMD5Executando := True;
    oErro.bMD5Confere := True;
    try
      DataXML := TDataXML.Create;
      DataXML.ServidorWS('www.digifarma.com.br');
      DataXML.SQL := 'select arquivo, local, md5 from sistema.check_files where arquivo = ' + QuotedStr(AnsiLowerCase(ExtractFileName(arquivo))) + ';';
      DataXML.ExecuteZ;

      if (DataXML.Erro = 0) then
      begin
        if (DataXML.RecordCount > 0) then
        begin
          if (str_MD5_file(arquivo) = DataXML.FieldByName('md5')) then
          begin
            oErro.bMD5Executando := False;
            oErro.bMD5Confere := False;
          end;
        end;
      end;
      oErro.bMD5Executando := False;
    except
      oErro.bMD5Executando := True;
      Exit;
    end;
  End;

  try

    oErro.bCarregaDLL := True;    {Erro: Carregando DLL}
    DLLHandle := LoadLibrary(PWideChar(arquivo));

    if DLLHandle <> 0 then
    begin
      oErro.bCarregaDLL := False;    {Erro: Carregando DLL}

      oErro.bChamaFuncao := True;    {Erro: Carregando a Função da DLL}
      Retorno := GetProcAddress(DLLHandle, PWideChar(Metodo));
      oErro.bChamaFuncao := False;    {Erro: Carregando a Função da DLL}


      oErro.bRetornoPonteiro := True;    {Erro: Retornando o ponteiro da DLL}
      Result := Retorno;
      oErro.bRetornoPonteiro := False;    {Erro: Retornando o ponteiro da DLL}
      oErro.bErro := False;
      Exit;
    end;
  except
    oErro.bErro := True;
    Exit;
  end;



end;

function App_RegisterDigifarma(AppName: String; IsDLL: Boolean = False):Boolean;
var
  oBanco:TDataXml;
  AppID, sCNPJ, sUF, sToken, mResp, mAppsId: String;
  mTabela: TIBQuery;
  ibBanco: IBX.IBDatabase.TIBDatabase;
  ibConexao: IBX.IBDatabase.TIBTransaction;
  oJson: TJSONObject;
  IdHTTP: TIdHTTP;
  params: TStringList;
  I: Integer;
begin

  try

    try
     {$REGION 'INICIA E CRIA AS CONEXÕES COM O BANCO DE DADOS EM MEMÓRIA'}
      ibBanco := IBX.IBDatabase.TIBDatabase.Create(Application);
      ibConexao := IBX.IBDatabase.TIBTransaction.Create(Application);
      ibBanco.LoginPrompt := False;
      ibConexao.Params.Add('read_committed');
      ibConexao.Params.Add('rec_version');
      ibConexao.Params.Add('nowait');
      ibConexao.DefaultDatabase := ibBanco;
      ibBanco.DefaultTransaction := ibConexao;
      ibBanco.DatabaseName := bd_Caminho_BD_Digifarma;
      ibBanco.params.add('password=saci');
      ibBanco.params.add('user_name=digifarma');
      mTabela := IBX.IBQuery.TIBQuery.Create(nil);
      mTabela.Database := ibBanco;
      mTabela.Transaction := ibConexao;
      {$ENDREGION}

     {$REGION 'PEGA O CNPJ E A UF NO BANCO LOCAL'}
      mTabela.SQL.Clear;
      mTabela.SQL.Add('select * from config');
      mTabela.Open;

      while not mTabela.EOF do
      begin
        sCNPJ := mTabela.FieldByName('cnpj').AsString;
        sUF := mTabela.FieldByName('estado').AsString;
        mTabela.Next;
      end;
     {$ENDREGION}

     {$REGION 'CADASTRA O APLICATIVO NA TABELA APP E PEGA O ID DO MESMO'}
      oBanco := TDataXml.Create;
      oBanco.ServidorWS('www.digifarma.com.br');
      oBanco.SQL := 'select app_id from  aplicativos.app where aplicativo = ' + QuotedStr(AppName);
      oBanco.ExecuteZ;

      if (oBanco.Erro <> 0) and (oBanco.Erro <> 1) then
      begin
        ShowMessage('Erro select :' + oBanco.Error);
      end;

      if oBanco.EOF then
      begin
        oBanco.close;
        oBanco.Free;

        oBanco := TDataXml.Create;
        oBanco.SQL := 'insert into aplicativos.app (aplicativo) values (:app)';
        oBanco.ParamName['app'] := AppName;
        oBanco.Execute;
        AppID := IntToStr(oBanco.LastID);

        if (oBanco.Erro <> 0) and (oBanco.Erro <> 1) then
        begin
          ShowMessage('Erro insert :' + oBanco.Error);
        end;
        oBanco.close;
        oBanco.Free;
      end
      else
      begin
        AppID := oBanco.FieldByName('app_id');
        oBanco.close;
        oBanco.Free;
      end;

      if (oBanco.Erro<>0) and (oBanco.Erro<>1) then
      begin
        ShowMessage('Erro:' + oBanco.Error);
      end;
    {$ENDREGION}

     {$REGION 'CONSULTA BLACKLIST'}
      oBanco := TDataXml.Create;
      oBanco.SQL := 'select * from aplicativos.app_blacklist where cnpj=' + IntToStr(StrToInt64(sCNPJ));
      oBanco.Execute;

      if not oBanco.EOF then
      begin
        mAppsId := oBanco.FieldByName('app_id');
        stl_Split(mAppsId, ',', params);
        if (params.Count > 0) then
        begin
          for I := 0 to params.Count - 1 do
          begin
            if (AppID = params[I]) then
            begin
              MessageDlg('ERRO: 009 - ENTRE EM CONTATO COM A DIGIFARMA', TMsgDlgType.mtWarning, [mbok], 0);
              Application.Terminate;
              xKillProcess;
            end;
          end;
        end
        else
        begin
          MessageDlg('ERRO: 009 - ENTRE EM CONTATO COM A DIGIFARMA', TMsgDlgType.mtWarning, [mbok], 0);
          Application.Terminate;
          xKillProcess;
        end;
      end;

      if (oBanco.Erro<>0) and (oBanco.Erro<>1) then
      begin
        ShowMessage('Erro:' + oBanco.Error);
      end;

      oBanco.close;
      FreeAndNil(oBanco);
    {$ENDREGION}

     {$REGION 'ENVIA PARA A API PRA REGISTRAR A FREQUENCIA'}
      Result := True;

      oJson := TJSONObject.Create(nil);
      oJson.AddPair('app_id', AppID);
      oJson.AddPair('cnpj', sCNPJ);

      if IsDLL then
      oJson.AddPair('versao', str_GetFileVer(GetModuleName(HInstance)))
      else
      oJson.AddPair('versao', str_GetFileVer(Application.ExeName));

      oJson.AddPair('UF', sUF);

      sToken := b64_Encode64(IntToHex(StrToInt(FormatDateTime('ddmmyyyy', Now)), 2) );
      params := TStringList.Create;
      params.Add('token=' + sToken);
      params.Add('json=' + oJson.ToJSON);

      IdHTTP := TIdHTTP.Create(Application);
      IdHTTP.HTTPOptions := [hoNoProtocolErrorException, hoWantProtocolErrorContent];
      IdHTTP.HandleRedirects := True;
      IdHTTP.Request.BasicAuthentication := True;
      IdHTTP.Request.UserAgent := 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0';
      IdHTTP.Request.ContentEncoding := 'multipart/form-data';

      mResp := IdHTTP.Post('http://www.digifarma.com.br/servicos/post/aplicativos/', params);
      {$ENDREGION}
    except
      result := False;
    end;

  finally
    FreeAndNil(oBanco);
    FreeAndNil(mTabela);
    FreeAndNil(params);
    FreeAndNil(IdHTTP);
    Result := True;

    try
      ibBanco.Close
    except
      result := False;
    end;

    FreeAndNil(ibBanco);

    try
      FreeAndNil(ibConexao);
    except
      result := False;
    end;
  end;
end;



{ RETORNA OS DADOS DO LINK PASSADO COMO PARAMETRO                                 }
{ Autor : Rafael                                                                  }
{ Data : 11/03/2012                                                               }

function lnk_BuscaLink(lpLink: String; Tipo: Integer = 0): String;
var
  ShellLink: IShellLink;
  PersistFile: IPersistFile;
  AnObj: IUnknown;
  FullPathAndNameOfFileToExecute, FullPathAndNameOfLinkFile, Description, ParamStringsOfFileToExecute,
  FullPathAndNameOfWorkingDirectroy, FullPathAndNameOfFileContiningIcon: array[0..MAX_PATH-1] of char;
  FindData: TWIN32FINDDATA;
  IconIndex, ShowCommand: Integer;
  HotKey: Word;

begin

  try
  // access to the two interfaces of the object
    AnObj       := CreateComObject(CLSID_ShellLink);
    ShellLink   := AnObj as IShellLink;
    PersistFile := AnObj as IPersistFile;
    // Opens the specified file and initializes an object from the file contents.
    PersistFile.Load(PWChar(WideString(lpLink)), 0);
  except
    Result := '';
    Exit;

  end;

  with ShellLink do
  begin
    try
      if (Tipo = 0) then
      Begin
        // Retrieves the path and file name of a Shell link object.
        FillChar(FullPathAndNameOfFileToExecute,MAX_PATH,#0);
        GetPath(FullPathAndNameOfFileToExecute, MAX_PATH,  FindData, SLGP_UNCPRIORITY);
        Result :=  StrPas(FullPathAndNameOfFileToExecute);
      End
      else if (Tipo = 1) then
      Begin
        // Retrieves the description string for a Shell link object.
        FillChar(Description,MAX_PATH,#0);
        GetDescription(Description, SizeOf(Description));
        Result :=  StrPas(Description);
      End
      else if (Tipo = 2) then
      Begin

        // Retrieves the command-line arguments associated with a Shell link object.
        FillChar(ParamStringsOfFileToExecute,MAX_PATH,#0);
        GetArguments(ParamStringsOfFileToExecute, SizeOf(ParamStringsOfFileToExecute));
        Result :=  StrPas(ParamStringsOfFileToExecute);
      End
      else if (Tipo = 3) then
      Begin

        // Retrieves the name of the working directory for a Shell link object.
        FillChar(FullPathAndNameOfWorkingDirectroy,MAX_PATH,#0);
        GetWorkingDirectory(FullPathAndNameOfWorkingDirectroy, SizeOf(FullPathAndNameOfWorkingDirectroy));
        Result :=  StrPas(FullPathAndNameOfWorkingDirectroy);
      End
      else if (Tipo = 4) then
      Begin

        // Retrieves the location (path and index) of the icon for a Shell link object.
        FillChar(FullPathAndNameOfFileContiningIcon,MAX_PATH,#0);
        GetIconLocation(FullPathAndNameOfFileContiningIcon, SizeOf(FullPathAndNameOfFileContiningIcon), IconIndex);
        Result :=  StrPas(FullPathAndNameOfFileContiningIcon);
      End
      else if (Tipo = 5) then
      Begin

        // Retrieves the hot key for a Shell link object.
        GetHotKey(HotKey);
        Result :=  IntToStr(HotKey);
      End
      else if (Tipo = 6) then
      Begin
        // Retrieves the show (SW_) command for a Shell link object.
        GetShowCmd(ShowCommand);

        Result :=  IntToStr(ShowCommand);

      end;
    except
      Result := '';
      Exit;
    end;
  end;

end;


{ DELETA LINKS DA AREA DE TRABALHO PRESERVANDO UMA LISTA DE EXCLUSÃO              }
{ Autor : Rafael                                                                  }
{ Data : 11/03/2012                                                               }
function lnk_DeletaLink(Habilitados: TStringList): Boolean;
  var Arquivo1,Arquivo2,Arquivos: TStringList;
      I,J: Integer;
      PathLnk,PathDigi: String;
      Arquivo: String;
begin

  Result := False;

  try
    Arquivo1 := TStringList.Create;
    Arquivo2 := TStringList.Create;
    Arquivos := TStringList.Create;

    Arquivo1 := stl_Localizar_Arquivo_Dir('C:\Users\Rafael\Desktop','*.lnk');
    Arquivo2 := stl_Localizar_Arquivo_Dir('C:\Users\Rafael\Desktop','*.lnk');

    Arquivo1.Delimiter := ';';
    Arquivo2.Delimiter := ';';
    Arquivos.Delimiter := ';';

    Arquivos.DelimitedText := Arquivo1.DelimitedText + ';' + Arquivo2.DelimitedText;

    I:=0;
    for I := 0 to (Arquivos.count-1) do
    Begin
      PathLnk :=  Arquivos[I];
      PathDigi := lowercase(lnk_BuscaLink(PathLnk));

      if (LowerCase(IncludeTrailingBackslash(ExtractFilePath(PathDigi))) = 'c:\digifarma\') then
      begin

        J :=0;
        if (xFindStrinList(Habilitados, ExtractFileName(PathDigi),J) = False) then
          DeleteFile(pwidechar(PathLnk));
      end;
    End;

    Result := True;
  except
    Result := False;
  end;


end;

{ RETORNA TABELA DE CONEXOES DO COMPUTADOR (PID;LocalIP;Port;RemoteIP;Port;Status_id;Status) }
{ Autor : Rafael                                                                   }
{ Data : 15/01/2015                                                                }

procedure NetStat(var Conexao: TStringList; Porta: Integer=0);
const
   ANY_SIZE = 1;
   iphlpapi = 'iphlpapi.dll';
   TCP_TABLE_OWNER_PID_ALL = 5;

   MIB_TCP_STATE:
   array[1..12] of string = ('CLOSED', 'LISTEN', 'SYN-SENT ','SYN-RECEIVED', 'ESTABLISHED', 'FIN-WAIT-1',
                             'FIN-WAIT-2', 'CLOSE-WAIT', 'CLOSING','LAST-ACK', 'TIME-WAIT', 'delete TCB');

type
  TCP_TABLE_CLASS = Integer;

  PMibTcpRowOwnerPid = ^TMibTcpRowOwnerPid;
  TMibTcpRowOwnerPid  = packed record
    dwState     : DWORD;
    dwLocalAddr : DWORD;
    dwLocalPort : DWORD;
    dwRemoteAddr: DWORD;
    dwRemotePort: DWORD;
    dwOwningPid : DWORD;
    end;

  PMIB_TCPTABLE_OWNER_PID  = ^MIB_TCPTABLE_OWNER_PID;
  MIB_TCPTABLE_OWNER_PID = packed record
   dwNumEntries: DWORD;
   table: Array [0..ANY_SIZE - 1] of TMibTcpRowOwnerPid;
  end;

var
   GetExtendedTcpTable:function  (pTcpTable: Pointer; dwSize: PDWORD; bOrder: BOOL; lAf: ULONG; TableClass: TCP_TABLE_CLASS; Reserved: ULONG): DWord; stdcall;


var
   hModule    : THandle;
   Error      : DWORD;
   TableSize  : DWORD;
   i          : integer;
   IpAddress  : in_addr;
   PID        : String;
   RemoteIp   : String;
   RemotePort : String;
   LocalIp    : String;
   LocalPort  : String;
   status_id  : String;
   Status     : String;
   pTcpTable  : PMIB_TCPTABLE_OWNER_PID;
begin
  TableSize := 0;

  Conexao := TStringList.Create;
  Conexao.Add('PID;LocalIP;Port;RemoteIP;Port;Status_id;Status');

  hModule             := LoadLibrary(iphlpapi);
  GetExtendedTcpTable := GetProcAddress(hModule, 'GetExtendedTcpTable');

  Error := GetExtendedTcpTable(nil, @TableSize, False, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0);
  if Error <> ERROR_INSUFFICIENT_BUFFER then exit;

  GetMem(pTcpTable, TableSize);
  try

    if GetExtendedTcpTable(pTcpTable, @TableSize, TRUE, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0) = NO_ERROR then
    Begin

      if (pTcpTable.dwNumEntries <=0) then
      Begin
        FreeMem(pTcpTable);
        Conexao.Clear;
        Exit;
      End;

      for i := 0 to pTcpTable.dwNumEntries - 1 do
      begin
        if ( ( (Porta >0) and (porta = htons(pTcpTable.Table[i].dwRemotePort)) ) or (porta<=0) )  then
        Begin
          PID              := IntToStr(pTcpTable.table[i].dwOwningPid);
          IpAddress.s_addr := pTcpTable.Table[i].dwRemoteAddr;
          RemoteIp         := string(inet_ntoa(IpAddress));
          RemotePort       := IntToStr( htons(pTcpTable.Table[i].dwRemotePort));
          IpAddress.s_addr := pTcpTable.Table[i].dwLocalAddr;
          LocalIp          := string(inet_ntoa(IpAddress));
          LocalPort        := IntToStr( htons(pTcpTable.Table[i].dwLocalPort));
          Status           := MIB_TCP_STATE[pTcpTable.Table[i].dwState];
          Status_id        := IntToStr(pTcpTable.Table[i].dwState);
          Conexao.Add(PID+';'+LocalIp+';'+LocalPort+';'+RemoteIp+';'+RemotePort+';'+Status_id+';'+Status);
        End;
      end;
   End;
  finally
     FreeMem(pTcpTable);
  end;
end;

{ RETORNA O INDICE DO COMBOBOX ONDE ESTÁ O ITEM STRING PASSADO COMO PARAMETRO     }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function num_ComboBox_Index(Str: String; Obj: TComboBox; bItemData: Boolean; bLike: Boolean = False) : Integer;
Var
    Item,A,I: Integer;
Begin

  A := (Obj.Items.Count -1);
  if (bItemData = False) then
  Begin
    for I := 0 to A do
    begin
      if (bLike=True) then
      begin
        if ( Pos(Str,Obj.Items[I]) <> 0) then
        Begin
          Result := I;
          Exit;
        End;
      end
      Else
      Begin
        if (Str = Obj.Items[I]) then
        Begin
          Result := I;
          Exit;
        End;
      End;
    end;
  End
  else
  Begin
    for I := 0 to A do
    begin
      Item := integer(Obj.Items.Objects[I]);
      if (StrToInt(Str) = Item) then
      Begin
        Result := I;
        Exit;
      End;
    end;
  End;
  Result := -1;

End;

{ RETORNA O TAMANHO DO ARQUIVO EM BYTES OU -1 QUANDO NÃO ENCONTRADO               }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function num_FileSize(fileName : wideString) : Int64;
  var
      sr : TSearchRec;
begin
  if SysUtils.FindFirst(fileName, faAnyFile, sr ) = 0 then
    result := Int64(sr.FindData.nFileSizeHigh) shl Int64(32) + Int64(sr.FindData.nFileSizeLow)
  else
    result := -1;

  SysUtils.FindClose(sr) ;
end;

{ Mostra a janela "abrir" ou "salvar" para selecionar o diretório e nao um arquivo especifico                                                                                }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }

function str_CutStr(Var CutString: String; Start, Len: Integer): string;
Begin

  if (Length(CutString)+1 >= (Start+Len) ) then
  Begin

    Result := MidStr(CutString, Start, Len);

    CutString := MidStr(CutString, 1, (Start-1)) +  MidStr(CutString, Start+Len, MaxInt);


  End;



End;

function str_Bina: string;
  Var Registro : TRegistry;
begin

  Result := 'icbox';

  try
    Registro := TRegistry.Create(KEY_READ or $0100);
    Registro.RootKey:=HKEY_CURRENT_USER;

    if Registro.OpenKey ('DIGIFARMA', False) then
    begin
      if Registro.ValueExists ('Bina') then
      Begin
        Result := Registro.ReadString('Bina');
      End;
      Registro.CloseKey;
      Registro.Free;
    end;

  except
    Result := 'icbox';
    Application.ProcessMessages;
  end;

end;

function str_BrowseDialog (const Title: string; const Flag: integer): string;
{*******************************************************************************
*  Exemplo:  BrowseDialog(Titulo,Flag);                                        *
*  Flags:                                                                      *
*  BIF_RETURNONLYFSDIRS   = Mostra pastas                                      *
*  BIF_BROWSEINCLUDEFILES = Mostra pastas e arquivos                           *
*  BIF_BROWSEFORCOMPUTER  = Mostra Computadores                                *
*  BIF_BROWSEFORPRINTER   = Mostra Impressoras                                 *
*  BIF_NEWDIALOGSTYLE    = $00000040                                           *
*  Edit1.text:=BrowseDialog('Selecione arquivo ou pasta',BIF_RETURNONLYFSDIRS);*
*******************************************************************************}
var
  lpItemID : PItemIDList;
  BrowseInfo : TBrowseInfo;
  DisplayName : array[0..MAX_PATH] of char;
  TempPath : array[0..MAX_PATH] of char;
begin
  Result:='';
  FillChar(BrowseInfo, sizeof(TBrowseInfo), #0);
  with BrowseInfo do begin
    hwndOwner := Application.Handle;
    pszDisplayName := @DisplayName;
    lpszTitle := PChar(Title);
    ulFlags := Flag;
  end;
  lpItemID := SHBrowseForFolder(BrowseInfo);
  if lpItemId <> nil then begin
    SHGetPathFromIDList(lpItemID, TempPath);
    Result := TempPath;
    GlobalFreePtr(lpItemID);
  end;
end;

{ GERA UMA STRING CRIPTOGRAFADA;/DECRIPTOGRAFADA DO TEXTO PASSADO COMO PARAMETRO. }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }

Function str_Criptografia(Action, Src: String): String;
Label Fim;
var KeyLen : Integer;
  KeyPos : Integer;
  OffSet : Integer;
  Dest, Key : String;
  SrcPos : Integer;
  SrcAsc : Integer;
  TmpSrcAsc : Integer;
  Range : Integer;
begin
  if (Src = '') Then
  begin
    Result:= '';
    Goto Fim;
  end;
  Key := 'YUQL23KL23DF90WI5E1JAS467NMCXXL6JAOAUWWMCL0AOMM4A4VZYW9KHJUI2347EJHJKDF3424SKL K3LAKDJSL9RTIKJ';
  Dest := '';
  KeyLen := Length(Key);
  KeyPos := 0;
  SrcPos := 0;
  SrcAsc := 0;
  Range := 256;
  if (Action = UpperCase('C')) then
  begin
    Randomize;
    OffSet := Random(Range);
    //OffSet := 170;
    Dest := Format('%1.2x',[OffSet]);
    for SrcPos := 1 to Length(Src) do
    begin
      SrcAsc := (Ord(Src[SrcPos]) + OffSet) Mod 255;
      if KeyPos < KeyLen then KeyPos := KeyPos + 1 else KeyPos := 1;
      SrcAsc := SrcAsc Xor Ord(Key[KeyPos]);
      Dest := Dest + Format('%1.2x',[SrcAsc]);
      OffSet := SrcAsc;
    end;
  end
  Else if (Action = UpperCase('D')) then
  begin
    OffSet := StrToInt('$'+ copy(Src,1,2));
    SrcPos := 3;
  repeat
    SrcAsc := StrToInt('$'+ copy(Src,SrcPos,2));
    if (KeyPos < KeyLen) Then KeyPos := KeyPos + 1 else KeyPos := 1;
    TmpSrcAsc := SrcAsc Xor Ord(Key[KeyPos]);
    if TmpSrcAsc <= OffSet then TmpSrcAsc := 255 + TmpSrcAsc - OffSet
    else TmpSrcAsc := TmpSrcAsc - OffSet;
    Dest := Dest + Chr(TmpSrcAsc);
    OffSet := SrcAsc;
    SrcPos := SrcPos + 2;
  until (SrcPos >= Length(Src));
  end;
  Result:= Dest;
  Fim:
end;


{ PEGA A DATA DO SERVIDOR http://www.ntp.br/ DA INTERNET                          }
{ Autor : Rafael                                                                  }
{ Data : 22/04/2013                                                               }
function str_Data_Hora_NTP: string;
var
  SNTPClient: TIdSNTP;
begin

  Result := '';
  try
    SNTPClient := TIdSNTP.Create(nil);
    SNTPClient.Host := 'pool.ntp.br';
    SNTPClient.SyncTime;
    Result := DateToStr(SNTPClient.DateTime);
    SNTPClient.Free;
  Except
    Result := 'ERRO';
  end;
end;


{ PEGA A DATA DO SERVIDOR http://www.ntp.br/ DA INTERNET                          }
{ Autor : Rafael                                                                  }
{ Data : 19/08/2014                                                               }
function date_ntp: TDateTime;
var
  SNTPClient: TIdSNTP;
begin

  Result := 0;
  try
    SNTPClient := TIdSNTP.Create(nil);
    SNTPClient.Host := 'a.st1.ntp.br';
 //   SNTPClient.ReceiveTimeout := 1200000;
 //   SNTPClient.Active := True;
 //   SNTPClient.Connect;
 //   SNTPClient.SyncTime;
    Result := SNTPClient.DateTime;
    SNTPClient.Free;
  Except
    Result := 0;
  end;
end;




{ PEGA A DATA DA ULTIMA ATUALIZAÇÃO DO ARQUIVO                                    }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function str_FileLastModified(const TheFile: string): string;
var
  FileH: THandle;
  LocalFT: TFileTime;
  DosFT: DWORD;
  LastAccessedTime: TDateTime;
  FindData: TWin32FindData;
begin
  Result := '';
  FileH := FindFirstFile(PwideChar(TheFile), FindData);
  if FileH <> INVALID_HANDLE_VALUE then
  begin
   //Windows.FindClose(Handle) ;
   if (FindData.dwFileAttributes AND
       FILE_ATTRIBUTE_DIRECTORY) = 0 then
    begin
     FileTimeToLocalFileTime(FindData.ftLastWriteTime, LocalFT);
     FileTimeToDosDateTime(LocalFT,LongRec(DosFT).Hi, LongRec(DosFT).Lo);
     LastAccessedTime := FileDateToDateTime(DosFT);
     Result := DateTimeToStr(LastAccessedTime);
    end;
  end;
end;

{ FORMATA O TAMANHO DO ARQUIVO PARA KB, MG OU GB. TAMANHO PASSADO COMO bytes      }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function str_FormatByteSize(const bytes: Longint): string;
  const
      B = 1; //byte
      KB = 1024 * B; //kilobyte
      MB = 1024 * KB; //megabyte
      GB = 1024 * MB; //gigabyte
begin
  if bytes > GB then
    result := FormatFloat('#.## GB', bytes / GB)
  else
    if bytes > MB then
      result := FormatFloat('#.## MB', bytes / MB)
    else
    if bytes > KB then
      result := FormatFloat('#.## KB', bytes / KB)
    else
      result := FormatFloat('#.## bytes', bytes) ;
end;

{ RETORNA O VALOR COM ZEROS A ESQUERDA PARA USO EM CAMPOS DE TAMANHO FIXO         }
{ Autor : Leo                                                                     }
{ Data : 14/03/2013                                                               }
function str_preenche_zero(valor:string;tamanho:Integer):string;
var
  x: Integer;
begin
  for x := Length(valor) to tamanho-1 do
    valor:='0'+valor;
  result:=valor;
end;

{ RETORNA O VALOR COM O CARACTERE ESCOLHIDO A ESQUERDA PARA USO EM CAMPOS DE TAMANHO FIXO         }
{ Autor : Filipe                                                                    }
{ Data : 29/04/2016                                                               }
function str_Preenche_caractere(valor:string;caractere:string;tamanho:Integer):string;
var
  x: Integer;
begin
  for x := Length(valor) to tamanho-1 do
    valor:=caractere+valor;
  result:=valor;
end;

function str_PathLinux(Texto: String): String;
  var aux: String;
Begin

  while ( (RightStr(Texto,1) = '/') or (RightStr(Texto,1) = '\') ) do
    Texto := LeftStr(Texto, Length(Texto)-1);

  if (InStr(1,Texto,'\')>0) then
  Begin
    aux := ReplaceText(Texto,'\\','//');
    aux := ReplaceText(aux,'\','/');
  End
  else
  Begin
    aux := ReplaceText(Texto,'//','\\');
    aux := ReplaceText(aux,'/','\');
  End;
  result := IncludeTrailingBackslash(aux);
End;

{ RETORNA O VALOR COM ESPAÇOS A DIREITA PARA USO EM CAMPOS DE TAMANHO FIXO        }
{ Autor : Leo                                                                     }
{ Data : 14/03/2013                                                               }
function str_Preenche_Espaco(valor:string;tamanho:Integer):string;
var
  x: Integer;
begin
  for x := Length(valor) to tamanho-1 do
    valor:=valor+' ';
  result:=valor;
end;

function stl_SerialPort(var Portas: TStringList) : Boolean;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  oBina         : TBina;
begin

  Result := False;
  Portas := TStringList.Create;
  Portas.Clear;

  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('select * from Win32_SerialPort','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    while oEnum.Next(1, colItem, iValue) = 0 do
    Begin

      oBina := TBina.Create;

      oBina.PORTA_ID := StrToInt(str_SoNumero(colItem.DeviceID));
      oBina.NOME := colItem.Description;
      oBina.DESCRICAO  := colItem.Name;

      Portas.AddObject(colItem.DeviceID, oBina);

//      Portas.Add(colItem.Caption + ';' + str_SoNumero(colItem.DeviceID));
      Result := true;
    end;
  except
    Result:= False;
  end;
End;

{ RETORNA O NUMERO DE SERIE DA PLACA MAE                                          }
{ Autor : Rafael                                                                  }
{ Data : 30/10/2012                                                               }
function str_serialPlacaMae: string;
var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;

begin
  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('SELECT * FROM Win32_BaseBoard','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    Result:=VarToStr(colItem.SerialNumber);
  Except
    Result:='';
  end;
End;

function str_uuid: String;
  var Uid: TGuid;
      lResult: HResult;
begin
  lResult := CreateGuid(Uid);
  if lResult = S_OK then
    Result := (GuidToString(Uid));
end;

function str_uuid_wmi: String;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;

  x: TCpuInfo;

begin
  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('Select * from Win32_ComputerSystemProduct','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    Result:=VarToStr(colItem.UUID);





  except
    Result:='';
  end;
End;

{ CONVERTE HEXADECIMAL PARA STRING                                                }
{ Autor : Rafael                                                                  }
{ Data : 12/04/2013                                                               }
function str_HexToString(H: String): String;
var I : Integer;
begin
  Result:= '';
  for I := 1 to length (H) div 2 do
    Result:= Result+Char(StrToInt('$'+Copy(H,(I-1)*2+1,2)));
end;

{ FORMAT STRING PADRAO HTML PARA ENVIO EM WEBSERVICE  ERRO: 'xmlParseEntityRef: no name  }
{ Autor : Rafael                                                                  }
{ Data : 16/06/2013                                                               }

Function str_Html(Texto :String; bAcentos: Boolean = True): String;
  Var aSub, aChar: TStringList;
      lsChar, lsSub: string;
//      aCHAR : String, aSUB: String;
      I: Integer;
Begin

  aSub := TStringList.Create;
  aChar := TStringList.Create;

  stl_Gera_Lista('&,<,>,",'',“,”', ',',  aCHAR);
  stl_Gera_Lista('&amp;,&lt;,&gt;,&quot;,&#39;,&quot;,&quot;',',', aSub);



//  aCHAR = Split(lsCHAR, ",");
//  aSUB = Split(lsSub, ",");

  For I := 0 To (aCHAR.Count-1) do
  Begin
    lsChar := aChar[I];
    If (AnsiContainsStr( Texto, lsChar)) Then
    Begin
      lsSub := aSUB[I];
      Texto :=  AnsiReplaceStr(Texto, lsCHAR, lsSub )
    End;
  end;

  If bAcentos = True Then
    Result := str_retira_Acentos(Texto)
  Else
    Result := Texto;


End;

{ CONVERTE CARACTERES ESPECIAIS EM CARACTERES HTML                                }
{ Autor : Rafael                                                                  }
{ Data : 24/04/2014                                                               }
{
function str_HtmlChars(Texto: String): String;
  var dicHtml: TDictionary<Integer,string>;
  I: Integer;
begin

  dicHtml.Add(Ord('Á'),'Á.&Aacute');  dicHtml.Add(Ord('á'),'á.&aacute'); dicHtml.Add(Ord('Â'),'Â.&Acirc');  dicHtml.Add(Ord('â'),'â.&acirc');
  dicHtml.Add(Ord('À'),'À.&Agrave');  dicHtml.Add(Ord('à'),'à.&agrave'); dicHtml.Add(Ord('Å'),'Å.&Aring');  dicHtml.Add(Ord('å'),'å.&aring');
  dicHtml.Add(Ord('Ã'),'Ã.&Atilde');  dicHtml.Add(Ord('ã'),'ã.&atilde'); dicHtml.Add(Ord('Ä'),'Ä.&Auml');   dicHtml.Add(Ord('ä'),'ä.&auml');
  dicHtml.Add(Ord('Æ'),'Æ.&AElig');   dicHtml.Add(Ord('æ'),'æ.&aelig');  dicHtml.Add(Ord('É'),'É.&Eacute'); dicHtml.Add(Ord('é'),'é.&eacute');
  dicHtml.Add(Ord('Ê'),'Ê.&Ecirc');   dicHtml.Add(Ord('ê'),'ê.&ecirc');  dicHtml.Add(Ord('È'),'È.&Egrave'); dicHtml.Add(Ord('è'),'è.&egrave');
  dicHtml.Add(Ord('Ë'),'Ë.&Euml');    dicHtml.Add(Ord('ë'),'ë.&euml');   dicHtml.Add(Ord('Í'),'Í.&Iacute'); dicHtml.Add(Ord('í'),'í.&iacute');
  dicHtml.Add(Ord('Î'),'Î.&Icirc');   dicHtml.Add(Ord('î'),'î.&icirc');  dicHtml.Add(Ord('Ì'),'Ì.&Igrave'); dicHtml.Add(Ord('ì'),'ì.&igrave');
  dicHtml.Add(Ord('Ï'),'Ï.&Iuml');    dicHtml.Add(Ord('ï'),'ï.&iuml');   dicHtml.Add(Ord('Ó'),'Ó.&Oacute'); dicHtml.Add(Ord('ó'),'ó.&oacute');
  dicHtml.Add(Ord('Ô'),'Ô.&Ocirc');   dicHtml.Add(Ord('ô'),'ô.&ocirc');  dicHtml.Add(Ord('Ò'),'Ò.&Ograve'); dicHtml.Add(Ord('ò'),'ò.&ograve');
  dicHtml.Add(Ord('Ø'),'Ø.&Oslash');  dicHtml.Add(Ord('ø'),'ø.&oslash'); dicHtml.Add(Ord('Õ'),'Õ.&Otilde'); dicHtml.Add(Ord('õ'),'õ.&otilde');
  dicHtml.Add(Ord('Ö'),'Ö.&Ouml');    dicHtml.Add(Ord('ö'),'ö.&ouml');   dicHtml.Add(Ord('Ú'),'Ú.&Uacute'); dicHtml.Add(Ord('ú'),'ú.&uacute');
  dicHtml.Add(Ord('Û'),'Û.&Ucirc');   dicHtml.Add(Ord('û'),'û.&ucirc');  dicHtml.Add(Ord('Ù'),'Ù.&Ugrave'); dicHtml.Add(Ord('ù'),'ù.&ugrave');
  dicHtml.Add(Ord('Ü'),'Ü.&Uuml');    dicHtml.Add(Ord('ü'),'ü.&uuml');   dicHtml.Add(Ord('Ç'),'Ç.&Ccedil'); dicHtml.Add(Ord('ç'),'ç.&ccedil');
  dicHtml.Add(Ord('Ñ'),'Ñ.&Ntilde');  dicHtml.Add(Ord('ñ'),'ñ.&ntilde'); dicHtml.Add(Ord('<'),'<.&lt');     dicHtml.Add(Ord('>'),'>.&gt');
  dicHtml.Add(Ord('&'),'&.&amp');     dicHtml.Add(Ord('"'),'".&quot');   dicHtml.Add(Ord('®'),'®.&reg');    dicHtml.Add(Ord('©'),'©.&copy');
  dicHtml.Add(Ord('Ý'),'Ý.&Yacute');  dicHtml.Add(Ord('ý'),'ý.&yacute'); dicHtml.Add(Ord('Þ'),'Þ.&THORN');  dicHtml.Add(Ord('þ'),'þ.&thorn');
  dicHtml.Add(Ord('ß'),'ß.&szlig');   dicHtml.Add(Ord('Ð'),'Ð.&ETH');    dicHtml.Add(Ord('ð'),'ð.&eth');

  for I := 0 to (Texto.Length-1) do
  begin
    if (dicHtml.ContainsKey[Ord(Texto[I])] = True) then
    begin

    end;

  end;

end;
}

function str_IdentifyingNumber_wmi: String;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
begin
  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('Select * from Win32_ComputerSystemProduct','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    Result:=VarToStr(colItem.IdentifyingNumber);
  except
    Result:='';
  end;
End;

{ RETORNA O USUÁRIO LOGADO NO WINDOWS                                             }
{ Autor : Rafael                                                                  }
{ Data : 09/04/2015                                                               }

function str_GetCurrentUser(): String;
  var s_u, s_d: String;
      n_pid: Integer;
begin

  n_pid := GetCurrentProcessId;

  GetUserAndDomainFromPID(n_pid, s_u, s_d);

end;

{ RETORNA O CONTEUDO DA JANELA DO MS-DOS                                          }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }

function str_GetDosOutput(CommandLine: string; Work: string = 'C:\'): string;
  var
    SA: TSecurityAttributes;
    SI: TStartupInfo;
    PI: TProcessInformation;
    StdOutPipeRead, StdOutPipeWrite: THandle;
    WasOK: Boolean;
    Buffer: array[0..255] of AnsiChar;
    BytesRead: Cardinal;
    WorkDir: string;
    Handle: Boolean;
begin
  Result := '';
  with SA do begin
    nLength := SizeOf(SA);
    bInheritHandle := True;
    lpSecurityDescriptor := nil;
  end;
  CreatePipe(StdOutPipeRead, StdOutPipeWrite, @SA, 0);
  try
    with SI do
    begin
      FillChar(SI, SizeOf(SI), 0);
      cb := SizeOf(SI);
      dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
      wShowWindow := SW_HIDE;
      hStdInput := GetStdHandle(STD_INPUT_HANDLE); // don't redirect stdin
      hStdOutput := StdOutPipeWrite;
      hStdError := StdOutPipeWrite;
    end;
    WorkDir := Work;
    Handle := CreateProcess(nil, PChar('cmd.exe /C ' + CommandLine),
                            nil, nil, True, 0, nil,
                            PChar(WorkDir), SI, PI);
    CloseHandle(StdOutPipeWrite);
    if Handle then
      try
        repeat
          WasOK := ReadFile(StdOutPipeRead, Buffer, 255, BytesRead, nil);
          if BytesRead > 0 then
          begin
            Buffer[BytesRead] := #0;
            Result := Result + Buffer;
          end;
        until not WasOK or (BytesRead = 0);
        WaitForSingleObject(PI.hProcess, INFINITE);
      finally
        CloseHandle(PI.hThread);
        CloseHandle(PI.hProcess);
      end;
  finally
    CloseHandle(StdOutPipeRead);
  end;
end;

{RETORNA O NOME DO EXECUTAVEL PELO HANDLE DA JANELA                            }
{ Autor : Leonardo                                                             }
{ Data : 08/11/2012                                                            }
function str_GetExeNameByHandle(WindowHandle: THandle): String;
var
ContinueLoop: BOOL;
FSnapshotHandle: THandle;
FProcessEntry32: TProcessEntry32;
ProcessID: PDWORD;
begin
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  ProcessID := AllocMem(SizeOf(DWORD));
  GetWindowThreadProcessId(WindowHandle, ProcessID);
  while (Integer(ContinueLoop) <> 0) do
  begin
    if FProcessEntry32.th32ProcessID = ProcessID^ then
    begin
      Result := FProcessEntry32.szExeFile;
      Break;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
  FreeMem(ProcessID);
end;


{RETORNA A VERSAO DO ARQUIVO}
{ Autor : Rafael    }
{ Data : 30/10/2012 }
function str_GetFileVer(const FileName: string): string;
begin
  with GetFileVerNumbers(FileName) do
    result := IntToStr(vMajor) + '.' +
      IntToStr(vMinor) + '.' + IntToStr(vRelease) + '.' +
      IntToStr(vBuild);
end;

{RETORNA O PRODUTO ID DA BIOS}
{ Autor : Rafael    }
{ Data : 03/02/2014 }
function hinfo_GetBios: string;
  var Registro: TRegistry;
      key: HKEY;
      valueType: DWORD;
      valueLen: DWORD;
      p, buffer: PChar;
      objWMIService : OLEVariant;
      colItems      : OLEVariant;
      colItem       : OLEVariant;
      oEnum         : IEnumvariant;
      iValue        : LongWord;
      Retorno       : string;
begin

  Result := '';

  try
    Registro := TRegistry.Create(KEY_READ or $0100);
    Registro.RootKey:=HKEY_LOCAL_MACHINE;
    if RegOpenKeyEx(HKEY_LOCAL_MACHINE, PChar('HARDWARE\DESCRIPTION\System'), 0, KEY_READ, key) = ERROR_SUCCESS then
    begin

      SetLastError(RegQueryValueEx(key, PChar('SystemBiosVersion'), nil, @valueType, nil, @valueLen));

      if ( (GetLastError = ERROR_SUCCESS) and (valueType = REG_MULTI_SZ) ) then
      begin
        GetMem(buffer, valueLen);
        try
          // receive the value's data (in an array).
          // Ein Array von Null-terminierten Strings
          // wird zurückgegeben
          RegQueryValueEx(key, PChar('SystemBiosVersion'), nil, nil, PBYTE(buffer), @valueLen);
          // Add values to stringlist
          // Werte in String Liste einfügen
          p := buffer;
          Result := String(p);
        finally
          FreeMem(buffer);
        end
      end
      else
      if (valueType = REG_SZ) then
      begin
        Retorno := Registro.ReadString('SystemBiosVersion');
        Result := Retorno;
      end;

      Registro.CloseKey;
      Registro.Free;
    end;
  except
    Result := '';
  end;

  try
    if (Result = '') then
    begin
      objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
      colItems      := objWMIService.ExecQuery('select * from Win32_BIOS','WQL',0);
      oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
      if oEnum.Next(1, colItem, iValue) = 0 then
      Result        :=VarToStr(colItem.SerialNumber);
    end;
  except
    Result        :=GetBIOSExtendedInfo;
  end;


end;

{RETORNA O PRODUTO ID DO WINDOWS}
{ Autor : Rafael    }
{ Data : 03/02/2014 }
function hinfo_GetMachineGuid: string;
  var Registro: TRegistry;

begin

  Result := '';

  try
    Registro := TRegistry.Create(KEY_READ or $0100);
    Registro.RootKey:=HKEY_LOCAL_MACHINE;
    if Registro.OpenKey ('SOFTWARE\Microsoft\Cryptography', False) then
    begin
      if ( (Registro.ValueExists('MachineGuid') = True) or (Registro.ValueExists('machineguid') = True) ) then
      Begin
        Result := Registro.ReadString('MachineGuid');
      End;
      Registro.CloseKey;
      Registro.Free;
    end;
  except
    Result := 'none';
  end;


end;

{RETORNA O ID DA PLACA MAE}
{ Autor : Rafael    }
{ Data : 03/02/2014 }

function hinfo_GetMoboSerial: string;
  Var ChaveMOB, fNome, fMD5: String;
      Arquivo: TStringList;
      a,b,c,d: LongWord;

  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
Begin

  asm
    push EAX
    push EBX
    push ECX
    push EDX

    mov eax, 1
    db $0F, $A2
    mov a, EAX
    mov b, EBX
    mov c, ECX
    mov d, EDX

    pop EDX
    pop ECX
    pop EBX
    pop EAX
  end;

  Result := '';

  ChaveMOB:=IntToHex(a,8) + '-' + IntToHex(b,8) + '-' + IntToHex(c,8) + '-' + IntToHex(d,8);

  if ( Length(Trim(ChaveMOB)) < 10  ) then
  Begin
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('SELECT SerialNumber FROM Win32_BaseBoard','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then


    ChaveMOB :=VarToStr(colItem.SerialNumber);
  end;
  Result := ChaveMOB;
End;

{RETORNA O PRODUTO ID DO WINDOWS}
{ Autor : Rafael    }
{ Data : 03/02/2014 }

function hinfo_GetProductID: string;
  var Registro: TRegistry;
begin

  Result := '';

  try
    Registro := TRegistry.Create(KEY_READ or $0100);
    Registro.RootKey:=HKEY_LOCAL_MACHINE;
    // Somente abre se a chave existir
    if Registro.OpenKey ('Software\Microsoft\Windows\CurrentVersion', False) then
    begin
      // Envia as informações ao form, vendo se os valores existem, primeiramente...
      if Registro.ValueExists ('ProductId') then
        Result := Registro.ReadString('ProductId');
      Registro.CloseKey;
      Registro.Free;
    end;
  except
    Result := '';
  end;

  if (Result = '') then

    try
      Registro := TRegistry.Create(KEY_READ or $0100);
      Registro.RootKey:=HKEY_LOCAL_MACHINE;
      // Somente abre se a chave existir
      if Registro.OpenKey ('Software\Microsoft\Windows NT\CurrentVersion', False) then
      begin
        // Envia as informações ao form, vendo se os valores existem, primeiramente...
        if Registro.ValueExists ('ProductID') then Result := Registro.ReadString('ProductID');
        if Registro.ValueExists ('ProductId') then Result := Registro.ReadString('ProductId');
        Registro.CloseKey;
        Registro.Free;
      end;
    except
      Result := 'NONE';
    end;



end;

{ RETORNA O PATH DA PASTA TEMP DO WINDOWS                                          }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
Function str_GetTempDir: String; stdcall;

var
  Buffer: array[0..MAX_PATH] of Char;
  lsTempPath, lsWindowsPath: String;
begin
  GetTempPath(MAX_PATH,  Buffer);
  lsTempPath :=  StrPas(Buffer);

  if (lsTempPath = '') then
  Begin
    lsWindowsPath := str_GetWinDir;
    lsTempPath := lsWindowsPath + '\temp';
    if (DirectoryExists(lsTempPath) = False) then
    Begin
      CreateDir(lsWindowsPath + '\temp');
      lsTempPath :=lsWindowsPath + '\temp\';
    End;
  End;

  Result := lsTempPath;
end;

{ RETORNA O PATH DA PASTA WINDOWS                                                 }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
Function str_GetWinDir: string; stdcall;
var
  dir: array [0..MAX_PATH] of Char;
begin
  GetWindowsDirectory(dir, MAX_PATH);
  Result := StrPas(dir);
end;

{ RETORNA O PATH DA PASTA WINDOWS                                                 }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
Function str_GetSystemDir: string; stdcall;
var
  dir: array [0..MAX_PATH] of Char;
begin
   GetSystemDirectory(dir, MAX_PATH);
  Result := StrPas(dir);
end;

{ GERA MASKARA DE TELEFONE PARA 9 DIGITOS                                         }
{ Autor : Rafael                                                                  }
{ Data : 09/10/2014                                                               }
function str_Mask_Telefone(const texto: string): string;
  var s_Numero, s_ddd, s_t1, s_t2: String;
Begin

  s_Numero := str_SoNumero(Texto);

  s_ddd := '('+str_CutStr(s_Numero,1,2) + ')';
  s_t2 := str_CutStr(s_Numero, Length(s_Numero)-3 ,4);
  s_t1 := s_Numero;

  Result := s_ddd + ' ' + s_t1 + '-' + s_t2;

end;

function str_MD5(const texto: string): string; {GERA UM MD5 DE UMA STRING}
var
  idMD5: TIdHashMessageDigest5;
begin
  idMD5 := TIdHashMessageDigest5.Create;
  try
    Result := idMD5.HashStringAsHex(texto);
  finally
    idMD5.Free;
  end;
end;


{ GERA O MD5 DO ARQUIVO ESPECIFICADO                                              }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function str_MD5_File(const fileName : string) : string;
var
  idmd5 : TIdHashMessageDigest5;
  fs : TFileStream;
begin
  idmd5 := TIdHashMessageDigest5.Create;

  fs := TFileStream.Create(fileName, fmOpenRead OR fmShareDenyWrite) ;
  try
    result := idmd5.HashStreamAsHex(fs);
  finally
    fs.Free;
    idmd5.Free;
  end;
end;

{ RETORNA CAMPO DE UMA LISTA DE ARGUMENTOS DELIMITADOS DE ACORDO COM SEU INDICE   }
{ Autor : RAFAEL                                                                  }
{ Data : 25/06/2013                                                               }
function str_Mid_Delimitado(Campos: String; Delimiter: Char;Indice:Integer):string;
  var Retorno: Boolean;
      Lista: TStringList;
Begin
  Result :='';

  try
    Lista := TStringList.Create;
    Lista.Clear;
    Lista.QuoteChar := '|';
    Lista.Delimiter := Delimiter;
    Lista.StrictDelimiter := True;
    Lista.DelimitedText := Campos;

    Result := Lista[Indice-1];
  except
    Result := '';
  end;


End;


{ RETORNO O NOME DO MÊS EM PORTUGUES                                              }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function str_NomeMes(mes:Integer):string;
begin
  if(mes =1)then
    Result:='JANEIRO'
  else
  if(mes =2)then
    Result:='FEVEREIRO'
  else
  if(mes =3)then
    Result:='MARÇO'
  else
  if(mes =4)then
    Result:='ABRIL'
  else
  if(mes =5)then
    Result:='MAIO'
  else
  if(mes =6)then
    Result:='JUNHO'
  else
  if(mes =7)then
    Result:='JULHO'
  else
  if(mes =8)then
    Result:='AGOSTO'
  else
  if(mes =9)then
    Result:='SETEMBRO'
  else
  if(mes =10)then
    Result:='OUTUBRO'
  else
  if(mes =11)then
    Result:='NOVEMBRO'
  else
  if(mes =12)then
    Result:='DEZEMBRO'
end;


{ FUNÇÃO RETORNA SERVIDOR PROXY : PORTA DO SERVIDOR PROXY                         }
{ Autor : Rafael                                                                  }
{ Data : 07/10/2014                                                               }
Function str_Proxy: string;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  Reg.RootKey := HKEY_CURRENT_USER;
  Reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\Internet Settings',false);
  Result :=Reg.ReadString('ProxyServer');
  Reg.Free;
end;

{ FUNÇÃO EXECUTE O QUOTEDSTR TROCANDO O APOSTROFO POR ´                           }
{ Autor : Rafael                                                                  }
{ Data : 19/06/2015                                                               }
function str_Quote(Texto: String): String;
Begin
  Texto := ReplaceText(Texto, '''', '´');
  Result := QuotedStr(Texto);
End;

{ FUNÇÃO TIRAR ACENTOS                                                            }
{ Autor : Rafael                                                                  }
{ Data : 16/06/2014                                                               }
Function str_Retira_Acentos(const Texto: String): String;
  var I, vPos: Integer;
  var s_a: string;
  Const vComAcento = 'ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜàáâãäåçèéêëìíîïòóôõöùúûü~^´`ªº°:!?';
  Const vSemAcento = 'AAAAAACEEEEIIIIOOOOOUUUUaaaaaaceeeeiiiiooooouuuu          ';

Begin

  s_a := Texto;

  For I := 1 To Length(s_a) do
  Begin
    vPos := AnsiContainsStr( vComAcento, AnsiMidStr(s_a, I, 1)).ToInteger;

    If vPos > 0 Then
      s_a := AnsiReplaceStr(s_a, AnsiMidStr(s_a, I, 1), AnsiMidStr(vSemAcento, (vPos-1), 1));
  End;
  Result := s_a;
End;


{ SALVA O CONTEUDO DO PARAMETRO TEXTO PARA UM ARQUIVO DEFINIDO NO PARAMETRO ARQUIVO   }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function str_SaveToFile (Const Texto:String; Const Arquivo:String):String;
  Var stl: TStringList;
Begin

  stl := TStringList.Create;
  stl.Add(Texto);
  stl.SaveToFile(Arquivo);
  FreeAndNil(stl);

End;

{ RETORNA SOMENTE CARCTERES DE A-Z e a-z                                          }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function str_SoLetras(Const Texto:String):String;
// Remove numeros de uma string deixando apenas letras
var
I: integer;
S: string;
begin
  S := '';
  for I := 1 To Length(Texto) Do
  begin
    if (Texto[I] in ['A'..'z']) then
    begin
      S := S + Copy(Texto, I, 1);
    end;
  end;
  result := S;
end;

{CONVERTE STRING PARA HEXADECIMAL                                                 }
{ Autor : Leonardo                                                                }
{ Data : 12/04/2013                                                               }

function str_StringtoHex(input: string): string;
var
  i, i2: Integer;
  s: string;
begin
  i2 := 1;
  for i := 1 to Length(input) do
  begin
    Inc(i2);
    if i2 = 2 then
    begin
      s  := s ;
      i2 := 1;
    end;
    s := s + IntToHex(Ord(input[i]), 2);
  end;
  Result := s;
end;


{ RETORNA SOMENTE NUMEROS                                                         }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }

function str_SoNumero(Const Texto:String):String;
// Remove caracteres de uma string deixando apenas numeros
var
I: integer;
S: string;
begin
  S := '';
  for I := 1 To Length(Texto) Do
  begin
    if (Texto[I] in ['0'..'9']) then
    begin
      S := S + Copy(Texto, I, 1);
    end;
  end;
  result := S;
end;

function num_SoNumero(Const Texto:String):Integer;
begin
  Result := StrToInt(iif(str_SoNumero(Texto)='','0',str_SoNumero(Texto)));
end;

{ RETORNA UMA PROPRIEDADE DO ARQUIVO                                              }
{ Autor : Leonardo                                                                }
{ Data : 21/08/2012                                                               }
{PARAMETROS :                                                                     }
{ * CompanyName                                                                   }
{  * FileDescription                                                              }
{  * FileVersion                                                                  }
{  * InternalName                                                                 }
{  * LegalCopyright                                                               }
{  * LegalTradeMarks                                                              }
{  * OriginalFilename                                                             }
{  * ProductName                                                                  }
{  * ProductVersion                                                               }
{  * Comments                                                                     }
function str_Propriedade_Arquivo(Arquivo,Propriedade: String): String;
const
  InfoNum           = 10;
  InfoStr           : Array[1..InfoNum] of String =
    ('CompanyName', 'FileDescription', 'FileVersion', 'InternalName',
    'LegalCopyright', 'LegalTradeMarks', 'OriginalFilename',
    'ProductName', 'ProductVersion', 'Comments');
var
  fCompanyName      : String;
  fFileDescription  : String;
  fFileVersion      : String;
  fInternalName     : String;
  fLegalCopyright   : String;
  fLegalTradeMark   : String;
  fOriginalFileName : String;
  fProductName      : String;
  fProductVersion   : String;
  fComments         : String;
  S                 : String;
  Len               : Cardinal;
  n                 : Cardinal;
  Buf               : PChar;
  Value             : PChar;
begin
  S := Arquivo;
  n := GetFileVersionInfoSize(PChar(S), n);
  if n > 0 then begin
     Buf := AllocMem(n);
     try
       GetFileVersionInfo(PChar(S), 0, n, Buf);
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[1]), Pointer(Value), Len) then
         fCompanyName := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[2]), Pointer(Value), Len) then
         fFileDescription := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[3]), Pointer(Value), Len) then
         fFileVersion := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[4]), Pointer(Value), Len) then
         fInternalName := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[5]), Pointer(Value), Len) then
         fLegalCopyright := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[6]), Pointer(Value), Len) then
         fLegalTradeMark := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[7]), Pointer(Value), Len) then
         fOriginalFileName := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[8]), Pointer(Value), Len) then
         fProductName := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[9]), Pointer(Value), Len) then
         fProductVersion := Value;
       if VerQueryValue(Buf, PChar('StringFileInfo\041604E4\' +
         InfoStr[10]), Pointer(Value), Len) then
         fComments := Value;
     finally
       FreeMem(Buf, n);
     end;
  end
  else begin
      fCompanyName := '';
      fFileDescription := '';
      fFileVersion := '';
      fInternalName := '';
      fLegalCopyright := '';
      fLegalTradeMark := '';
      fOriginalFileName := '';
      fProductName := '';
      fProductVersion := '';
      fComments := '';
  end;
  result := '?????';
  if Propriedade = 'CompanyName' then result := fCompanyName;
  if Propriedade = 'FileDescription' then result := fFileDescription;
  if Propriedade = 'FileVersion' then result := fFileVersion;
  if Propriedade = 'InternalName' then result := fInternalName;
  if Propriedade = 'LegalCopyright' then result := fLegalCopyright;
  if Propriedade = 'LegalTradeMarks' then result := fLegalTradeMark;
  if Propriedade = 'OriginalFilename' then result := fOriginalFileName;
  if Propriedade = 'ProductName' then result := fProductName;
  if Propriedade = 'ProductVersion' then result := fProductVersion;
  if Propriedade = 'Comments' then result := fComments;
end;


{ GERA UMA STRINGLIST DE UM TEXTO COM DELIMITAÇÕES                                }
{ Autor : Rafael                                                       }
{ Data : 17/06/2014                                                               }
function stl_Gera_Lista(Input: string; const Delimiter: Char; var Lista: TStringList): Boolean;
  var buf:string;
      x: Integer;
begin

	Result := False;

  Lista.Clear;
	Lista.Clear;
	Lista.QuoteChar := '|';
	Lista.Delimiter := Delimiter;
	Lista.StrictDelimiter := True;
	Lista.DelimitedText := Input;

  Result := True;

end;

{ GERA UMA STRINGLIST DE UM TEXTO COM DELIMITAÇÕES                                }
{ Autor : Leonardo & Rafael                                                       }
{ Data : 02/08/2012                                                               }
{Data da alteração: 10/04/2013 - LEO                                              }
function stl_Split (Input: string; const Delimiter: Char; var Lista: TStringList): Boolean;
  var buf:string;
      x: Integer;
begin
	Result := False;
  Lista.Clear;
  if(Trim(Input)<>'')then
  begin
    x:=1;
    if(Input[1]=Delimiter)then
      x:=2;
    repeat
      begin
        if(Input[x]<>Delimiter)then
          buf:=buf+Input[x]
        else
        begin
          Lista.Add(buf);
          buf:='';
        end;
      end;
      x:=x+1;
    until x>Length(Input);
    if(buf<>'')then//CASO O ULTIMO CAMPO NAO TENHA DELIMITADOR ADICIONA A LISTA AQUI
      Lista.Add(buf);
    buf:='';
    Result := True;
  end;
	{Lista.Clear;
	Lista.QuoteChar := '|';
	Lista.Delimiter := Delimiter;
	Lista.StrictDelimiter := True;
	Lista.DelimitedText := Input; }
end;

{ GERA UMA STRINGLIST COM O NOME DOS PROCESSOS EM EXECUÇÃO NO MOMENTO             }
{ Autor : Leonardo                                                                }
{ Data : 22/04/2013                                                               }
function stl_lista_processos(var stList:TStringList):boolean;
const
PROCESS_TERMINATE=$0001;
var
ContinueLoop: BOOL;
FSnapshotHandle: THandle;
FProcessEntry32: TProcessEntry32{declarar Uses Tlhelp32};
begin
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle,FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    stList.Add(FProcessEntry32.szExeFile);
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
  result:=True;
end;

{ GERA UMA STRINGLIST NOM O NOME DOS ARQUIVOS LOCALIZADOS NO PATH                 }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
{ ALTERAÇÃO : RAFAEL                                                              }
{ Data : 27/08/2013                                                               }
{ Data : INCLUSÃO DO PARAMETRO bEXT PARA OPÇÃO DE CARREGAR A LISTA DE ARQUIVOS SEM A EXTENSÃO DO MESMO }
function stl_Localizar_Arquivo(Local,extensao: string; bEXT: Boolean = True):TStringList;
var
SearchRec: TSearchRec;
SL:TStringList;
begin
  SL:=TStringList.Create;
  Local:= Trim(Local);
  if Local[length(Local)] <> '' then
  Local:= Local + '';
  SetCurrentDir(Local);
  if FindFirst(extensao, faArchive , SearchRec) = 0 then
  Repeat

    if (bEXT = False) then

      SL.Add( StringReplace(SearchRec.Name, ExtractFileExt(SearchRec.Name),'',[rfReplaceAll]) )
    else
      SL.Add(SearchRec.Name);

  Until FindNext(SearchRec) <> 0;
  System.SysUtils.FindClose(SearchRec);
  //FindClose(SearchRec);
  SetCurrentDir('c:\digifarma');
  result:=SL;
end;


{ GERA UMA STRINGLIST NOM O NOME DOS ARQUIVOS LOCALIZADOS NO PATH                 }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function stl_Localizar_Arquivo_Dir(Local,extensao: string):TStringList;
var SearchRec: TSearchRec;
    SL:TStringList;
    Path: String;
begin

  Path := IncludeTrailingBackslash(Local);

  SL:=TStringList.Create;
  Local:= Trim(Local);
  if Local[length(Local)] <> '' then
  Local:= Local + '';
  SetCurrentDir(Local);
  if FindFirst(extensao, faArchive , SearchRec) = 0 then
  Repeat
    SL.Add(Path + SearchRec.Name);
  Until FindNext(SearchRec) <> 0;
  System.SysUtils.FindClose(SearchRec);

  SetCurrentDir('c:\digifarma');
  //FreeAndNil(SearchRec);
  result:=SL;
end;

{ GERA UMA STRINGLIST NOM O NOME DOS SUBDIRETORIOS LOCALIZADOS NO PATH            }
{ Autor : Leonardo                                                                }
{ Data : 05/03/2013                                                               }
function stl_Localizar_Diretorio(Local: string):TStringList;
var
SearchRec: TSearchRec;
SL:TStringList;
begin
  SL:=TStringList.Create;
  Local:= IncludeTrailingBackslash(Trim(Local));
  if Local[length(Local)] <> '' then
  Local:= Local + '';
  SetCurrentDir(Local);
  if FindFirst(Local+'*.*', faDirectory , SearchRec) = 0 then
  Repeat
    if(SearchRec.Name<>'.')and(SearchRec.Name<>'..')and(DirectoryExists(local+'\'+SearchRec.Name)) then
      SL.Add(SearchRec.Name);
  Until FindNext(SearchRec) <> 0;
  //FreeAndNil(SearchRec);
  SetCurrentDir('c:\digifarma');
  result:=SL;
end;

{ TENTA DE VARIAS MANEIRAS SABER SE TEM CONEXAO COM A INTERNET                    }
{ Autor : Leonardo                                                                }
{ Data : 04/08/2015                                                               }
function xChecaInternet():Boolean;
var
s:TMemoryStream;
idHttp1:TIdHTTP;
begin
  result:=false;
  idHttp1:=TIdHTTP.Create(Application);
  s:=TMemoryStream.Create;
  try
    IdHTTP1.Get('http://www.digifarma.com.br',s);
    if(s.SIZE>0)then
    begin
      result:=True;
      s.Clear;
      Exit;
    end;
  except
    on e:exception do
    begin
      result:=false;
      s.Clear;
    end;
  end;
  try
    IdHTTP1.Get('http://www.digifarmaonline.com.br',s);
    if(s.SIZE>0)then
    begin
      s.Clear;
      result:=True;
      Exit;
    end;
  except
    on e:exception do
    begin
      s.Clear;
      result:=false;
    end;
  end;
end;

function xCopyFiles(FromDir: String; Destination: String): Boolean;
  var OpStruct: TSHFileOpStruct;
Begin

  try
    FromDir := FromDir;
    Destination := Destination;

    FillChar (OpStruct, SizeOf (OpStruct), #0);

    With OpStruct Do
    Begin
      Wnd:=0;
      wFunc:=FO_COPY;
      pFrom:= PCHAR(FromDir+ #0#0);
      pTo:= PCHAR(Destination+ #0#0);
      fFlags:=FOF_ALLOWUNDO or FOF_SILENT or FOF_NOCONFIRMATION;
    End;
    if not DirectoryExists(Destination) then ForceDirectories(Destination);
      SHFileOperation(OpStruct);

    Result := True;

  except
    Result := False;
  end;
End;

function xDeleteDir(dir: string): Boolean;
var
  fos: TSHFileOpStruct;
begin
  ZeroMemory(@fos, SizeOf(fos));
  with fos do
  begin
    wFunc  := FO_DELETE;
    fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
    pFrom  := PChar(dir + #0);
  end;
  Result := (0 = ShFileOperation(fos));
end;


{ DELETA TODOS OS ARQUIVOS DA PASTA PASSADA COMO PARAMETRO                        }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function xDeleteFiles(FileName:string; MandarPraLixeira: Boolean = False): boolean;
  var
  fos : TSHFileOpStruct;
begin
  FillChar(fos, SizeOf(fos), 0);
  with fos do begin
    wFunc := FO_DELETE;
    pFrom := PChar(FileName);
    if (MandarPraLixeira = True) then
      fFlags := FOF_NOERRORUI or FOF_ALLOWUNDO or FOF_NOCONFIRMATION or FOF_SILENT
    Else
      fFlags := FOF_NOERRORUI or FOF_NOCONFIRMATION or FOF_SILENT;
  end;
  Result := (ShFileOperation(fos)=0);
end;


procedure xFileEnDeCrypto(Arquivo_Origem, Arquivo_Saida: string; Chave: Word) ;
 var
   InMS, OutMS: TMemoryStream;
   cnt: Integer;
   C: byte;
 begin
   InMS := TMemoryStream.Create;
   OutMS := TMemoryStream.Create;
   try
     InMS.LoadFromFile(Arquivo_Origem) ;
     InMS.Position := 0;
     for cnt := 0 to InMS.Size - 1 do
       begin
         InMS.Read(C, 1) ;
         C := (C xor not (ord(chave shr cnt))) ;
         OutMS.Write(C, 1) ;
       end;
     OutMS.SaveToFile(Arquivo_Saida) ;
   finally
     InMS.Free;
     OutMS.Free;
   end;
 end;


{ EXECUTA A LINHA DE COMANDO COM A OPÇÃO DE AGUARDAR O FIM DA EXECUÇÃO            }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function xFileExecute(ahWnd: Cardinal; const aFileName, aParams, aStartDir: string; aShowCmd: Integer; aWait: Boolean): Integer;
var
  Info: TShellExecuteInfo;
  ExitCode: DWORD;
begin

  Result := -1;
  FillChar(Info, SizeOf(Info), 0);
  Info.cbSize := SizeOf(TShellExecuteInfo);
  with Info do begin
    fMask := SEE_MASK_NOCLOSEPROCESS;
    Wnd := ahWnd;
    lpFile := PChar(aFileName);
    lpParameters := PChar(aParams);
    lpDirectory := PChar(aStartDir);
    nShow := aShowCmd;
  end;

  if ShellExecuteEx(@Info) then
  begin
    if aWait then
    begin
      repeat
        Sleep(1);
        Application.ProcessMessages;
        GetExitCodeProcess(Info.hProcess, ExitCode);
      until (ExitCode <> STILL_ACTIVE) or Application.Terminated;
      Result := ExitCode;
    end;
  end
end;


{                     PROCURA UM ITEM NA LISTA                                    }
{ Autor : Rafael                                                                  }
{ Data : 08/03/2013                                                               }
function xFindStrinList(Lista: TStringList; Item: string; var Index: Integer): Boolean;
  Var I: Integer;
Begin

  Result := False;

  for I := 0 to (Lista.Count-1) do
  begin
    if ( LowerCase(Trim(Lista[I])) = LowerCase(Trim(Item))) then
    Begin
      Index := I;
      Result := True;
    End;
  end;

End;

{ VERIFICA SE É UMA DATA VALIDA                                                   }
{ Autor : Rafael                                                                  }
{ Data : 29/08/2012                                                               }
Function xFindWindowLike ( const Titulo: string; var completo: string ): Integer;
  Var r, hWindow: Integer;
      Tit: string;
      sl: TStringList;
begin
  hWindow := GetForegroundWindow;

  while hWindow <> 0 do
  begin
    Tit := StringOfChar(Chr(32), 255);
    r := GetWindowText(hWindow, PWideChar(Tit), 255);
    Tit :=  LeftStr(Tit,r);
    if (Pos(LowerCase(Titulo), LowerCase(Tit))<>0) then
    Begin
      Result := hWindow;
      Completo := Tit;
    End;
    hWindow := GetWindow(hWindow, GW_HWNDNEXT);
  end;

end;


{ VERIFICA SE É UMA DATA VALIDA                                                   }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
Function xIsDate ( const DateString: string; formato: String = 'DD/MM/YYYY'): boolean;
begin
  try
    if (Length(DateString) <> Length(formato)) then
    Begin
      result := false;
      Exit;
    End
    else
    Begin
      if (Length(DateString) > 10) then
        StrToDateTime ( DateString )
      else
        StrToDate ( DateString )
    End;

    result := true;
  except
    result := false;
  end;
end;

{ VERIFICA SE É UM NÚMERO VÁLIDO                                                  }
{ Autor : Rafael                                                                  }
{ Data : 08/10/2012                                                               }
Function xIsNumeric ( const NumericString: string ): boolean;
  var I: Integer;
begin
  try
    Result := TryStrToInt(NumericString, I);
  except
    result := false;
  end;
end;

{ RETIRA A DLL INFORMADA DA MEMÓRIA                                               }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function xKillDll(aDllName: string): Boolean;
var
  hDLL: THandle;
  aName: array[0..10] of char;
  FoundDLL: Boolean;
begin
  StrPCopy(aName, aDllName);
  FoundDLL := False;
  repeat
    hDLL := GetModuleHandle(aName);
    if hDLL = 0 then
      Break;
    FoundDLL := True;
    FreeLibrary(hDLL);
  until False;
end;

function xRunAsAdmin(hWnd: HWND; filename: string; Parameters: string): Boolean;
{ EXECUTA UM APLICATIVO COMO ADMINISTRADOR (COMPATIVEL COM VISTA,7,8,2008 SERVER  }
{ Autor : Leo                                                                     }
{ Data : 06/05/2013                                                               }
{
    See Step 3: Redesign for UAC Compatibility (UAC)
    http://msdn.microsoft.com/en-us/library/bb756922.aspx
}
var
    sei: TShellExecuteInfo;
begin
    ZeroMemory(@sei, SizeOf(sei));
    sei.cbSize := SizeOf(TShellExecuteInfo);
    sei.Wnd := hwnd;
    sei.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
    sei.lpVerb := PChar('runas');
    sei.lpFile := PChar(Filename); // PAnsiChar;
    if parameters <> '' then
    	sei.lpParameters := PChar(parameters); // PAnsiChar;
    sei.nShow := SW_SHOWNORMAL; //Integer;

    Result := ShellExecuteEx(@sei);
end;

{ SALVA TEXTO EM LOCAL ESPECIFICADO AMBOS PASSADOR POR PARAMETRO                  }
{ Autor : Rafael                                                                  }
{ Data : 15/06/2015                                                               }
function xSaveToFile(const Texto, Caminho: String; append: Boolean = False):boolean;
  var st_temp: TStringList;
begin

  st_temp := TStringList.Create;

  if (append = True) then
    if (FileExists(caminho) = True) then
      st_temp.LoadFromFile(caminho);

  st_temp.Add(Texto);
  st_temp.SaveToFile(Caminho);

  FreeAndNil(st_temp);

end;

{ VALIDA A CHAVE DA NF-e                                                          }
{ Autor : Leonardo                                                                }
{ Data : 02/08/2012                                                               }
function xValidarChaveNFe(const ChaveNFe: string):boolean;
const
  PESO : Array[0..43] of Integer = (4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2, 0);
var
  Retorno : boolean;
  aChave  : Array[0..43] of Char;
  Soma    : Integer;
  Verif   : Integer;
  I       : Integer;
begin
  Retorno := false;
  try
    try
      if Length(ChaveNFe) < 44 then
        exit;

      StrPCopy(aChave,StringReplace(ChaveNFe,' ', '',[rfReplaceAll]));
      Soma := 0;
      for I := Low(aChave) to High(aChave) do
        Soma := Soma + (StrToInt(aChave[i]) * PESO[i]);

      if Soma = 0 then
        raise Exception.Create('');

      Soma := Soma - (11 * (Trunc(Soma / 11)));
      if (Soma = 0) or (Soma = 1) then
        Verif := 0
      else
        Verif := 11 - Soma;

      Retorno := Verif = StrToInt(aChave[43]);
    except
      Retorno := false;
    end;
  finally
    Result := Retorno;
  end;
end;

function xVersao(Arquivo: string; Var VMaior, VMenor, VRelease, VCompilacao: String): Boolean;
  var Lista: TStringList;
      s_versao: string;
begin

  Lista := TStringList.Create;
  s_versao := str_GetFileVer(Arquivo);
  stl_Split(s_versao ,'.',Lista);

  VMaior := Lista[0];
  VMenor := Lista[1];
  VRelease:= Lista[2];
  VCompilacao := Lista[3];
  FreeAndNil(Lista);
end;

{ VERIFICA SE O COMPUTADOR ESTÁ CONECTADO A INTERNET                              }
{ Autor : Rafael                                                                  }
{ Data : 02/08/2012                                                               }
function xPing(IP: String): boolean;
  var
  Retorno : String;
begin
  Retorno := str_GetDosOutput('ping www.google.com');

  xPing := False;
  if Pos('TTL=', Retorno) > 0 then
    xPing := True;


end;


function xPingICMP( host: String): Boolean;
var
  IdICMPClient: TIdICMPClient;
begin

  try

    IdICMPClient := TIdICMPClient.Create( nil );

    IdICMPClient.Host := host;
    IdICMPClient.ReceiveTimeout := 500;
    IdICMPClient.Ping;

    result := ( IdICMPClient.ReplyStatus.BytesReceived > 0 );
    IdICMPClient.Free;

  except
    Result := False;
  end

end;

{ FUNCAO QUE VERIFICA SE UM PRECESSO ESTA SENDO EXECUTADO                         }
{ Autor : Leonardo                                                                }
{ Data : 21/08/2012                                                               }


function xProcessoExiste(ExeFileName: string): boolean;
const
PROCESS_TERMINATE=$0001;
var
ContinueLoop: BOOL;
FSnapshotHandle: THandle;
FProcessEntry32: TProcessEntry32{declarar Uses Tlhelp32};
begin
  result := false;

  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle,FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExeFileName))
    or (UpperCase(FProcessEntry32.szExeFile) = UpperCase(ExeFileName))) then
    begin
      Result := true;
      exit;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

//function xValidaEAN(CodBar: string): Boolean;
//var
//sl_CodIni : String;
//il_SeqSom, il_Soma, il_TodNum : Integer;
//begin
//            // se codigo eh invalido
//  IF (Length(Trim(CodBar)) = 0) THEN
//    Result := FALSE
//
//  ELSE
//  IF ( (Length(Trim(CodBar)) <> 13) AND (Length(Trim(CodBar)) <> 8) ) THEN
//    Result := FALSE
//
//          // se codigo eh certo, entao continua
//  ELSE
//  BEGIN
//    sl_CodIni := Copy(Trim(CodBar),1,Length(CodBar)-1);
//    il_Soma := 0;
//    il_SeqSom := 1;
//
//    // se for ean8 entao completa numeros
//    IF Length(Trim(CodBar)) = 8 THEN
//    sl_CodIni := '00000'+sl_CodIni;
//
//         // continua
//
//    FOR il_TodNum := 1 TO Length(sl_CodIni) DO
//    BEGIN
//      IF il_SeqSom = 1 THEN
//      BEGIN
//        il_Soma := il_Soma +
//        (StrToInt(Copy(sl_CodIni, il_TodNum, 1)) * il_SeqSom);
//        il_SeqSom := 3;
//      END
//      ELSE
//      BEGIN
//        il_Soma := il_Soma +
//        (StrToInt(Copy(sl_CodIni, il_TodNum, 1)) * il_SeqSom);
//        il_SeqSom := 1;
//      END
//
//    END;
//
//  // calcula o restante
//
//  // para numeros que nao sao zero no final
//    IF Copy(IntToStr(il_Soma),Length(IntToStr(il_Soma)),1 ) <> '0' THEN
//    BEGIN
//      IF Copy(Trim(CodBar), Length(CodBar), 1) =
//        IntToStr(10-StrToInt( Copy(IntToStr(il_Soma),
//        Length(IntToStr(il_Soma)),1) )) THEN
//        Result := TRUE
//      ELSE
//        Result := FALSE
//    END
//    ELSE
//    IF StrToInt(Copy(Trim(CodBar), Length(CodBar), 1)) = 0 THEN
//      Result := TRUE
//    ELSE
//      Result := FALSE;
//  END;
//enD;
function xValidaEAN(CodBar: string): Boolean;
var
sl_CodIni : String;
il_SeqSom, il_Soma, il_TodNum : Integer;
cod:Int64;
begin

  try
    cod:=StrToInt64('0'+CodBar);
  except
    on ex:exception do
    begin
      Result:=False;
      exit;
    end;
  end;

  if (Length(IntToStr(cod))<7) then
  Begin
    Result:=False;
    exit;
  End;


  // se codigo eh invalido
  IF (Length(Trim(CodBar)) = 0) THEN
  begin
    Result := False;
    exit;
  end;


  IF ( (Length(Trim(CodBar)) <> 13) AND (Length(Trim(CodBar)) <> 8) ) THEN
  begin
    Result := FALSE
  end

          // se codigo eh certo, entao continua
  ELSE
  BEGIN
    sl_CodIni := Copy(Trim(CodBar),1,Length(CodBar)-1);
    il_Soma := 0;
    il_SeqSom := 1;

    // se for ean8 entao completa numeros
    IF Length(Trim(CodBar)) = 8 THEN
    sl_CodIni := '00000'+sl_CodIni;

         // continua

    FOR il_TodNum := 1 TO Length(sl_CodIni) DO
    BEGIN
      IF il_SeqSom = 1 THEN
      BEGIN
        il_Soma := il_Soma +
        (StrToInt(Copy(sl_CodIni, il_TodNum, 1)) * il_SeqSom);
        il_SeqSom := 3;
      END
      ELSE
      BEGIN
        il_Soma := il_Soma +
        (StrToInt(Copy(sl_CodIni, il_TodNum, 1)) * il_SeqSom);
        il_SeqSom := 1;
      END

    END;

  // calcula o restante

  // para numeros que nao sao zero no final
    IF Copy(IntToStr(il_Soma),Length(IntToStr(il_Soma)),1 ) <> '0' THEN
    BEGIN
      IF Copy(Trim(CodBar), Length(CodBar), 1) =
        IntToStr(10-StrToInt( Copy(IntToStr(il_Soma),Length(IntToStr(il_Soma)),1) )) THEN
        Result := TRUE
      ELSE
        Result := FALSE
    END
    ELSE
    IF StrToInt(Copy(Trim(CodBar), Length(CodBar), 1)) = 0 THEN
      Result := TRUE
    ELSE
      Result := FALSE;
  END;
enD;

{ FUNCAO QUE VALIDA O E-MAIL                                                      }
{ Autor : Rafael                                                                  }
{ Data : 05/06/2015                                                               }
function xValidaEmail(aEmail: string): Boolean;
var
  vRegex: TRegEx;
  s_expression: string;
begin

  s_expression := '^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}' +
                       '\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\' +
                       '.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$';

  Result := False;
  vRegex := TRegEx.Create(s_expression);
  try
    if vRegex.IsMatch(aEmail, s_expression) then Result := True;
  finally

  end;
end;


{ FUNCAO QUE NÚMERO DO CNPJ                                                       }
{ Autor : Rafael                                                                  }
{ Data : 19/12/2012                                                               }

function xValidaCNPJ(num: string): boolean;
var
   n1,n2,n3,n4,n5,n6,n7,n8,n9,n10,n11,n12: integer;
   d1,d2: integer;
   digitado, calculado: string;
begin

  n1:=StrToInt(num[1]);
  n2:=StrToInt(num[2]);
  n3:=StrToInt(num[3]);
  n4:=StrToInt(num[4]);
  n5:=StrToInt(num[5]);
  n6:=StrToInt(num[6]);
  n7:=StrToInt(num[7]);
  n8:=StrToInt(num[8]);
  n9:=StrToInt(num[9]);
  n10:=StrToInt(num[10]);
  n11:=StrToInt(num[11]);
  n12:=StrToInt(num[12]);

  d1:=n12*2+n11*3+n10*4+n9*5+n8*6+n7*7+n6*8+n5*9+n4*2+n3*3+n2*4+n1*5;
  d1:=11-(d1 mod 11);
  if d1>=10 then d1:=0;

  d2:=d1*2+n12*3+n11*4+n10*5+n9*6+n8*7+n7*8+n6*9+n5*2+n4*3+n3*4+n2*5+n1*6;
  d2:=11-(d2 mod 11);

  if d2>=10 then d2:=0;

  calculado:=inttostr(d1)+inttostr(d2);
  digitado:=num[13]+num[14];

  if calculado=digitado then
    Result := True
  else
    Result := False;

end;

{ FUNCAO QUE VALIDA NÚMERO DO CPF                                                 }
{ Autor : Leonardo                                                                }
{ Data : 22/10/2013                                                               }

function xValidaCPF(CPF: string): boolean;
var
  dig10, dig11: string;
  s, i, r, peso: integer;
begin
  // length - retorna o tamanho da string (CPF é um número formado por 11 dígitos)
  if ((CPF = '00000000000') or (CPF = '11111111111') or (CPF = '22222222222') or
    (CPF = '33333333333') or (CPF = '44444444444') or (CPF = '55555555555') or
    (CPF = '66666666666') or (CPF = '77777777777') or (CPF = '88888888888') or
    (CPF = '99999999999') or (length(CPF) <> 11)) then
  begin

    Result:= false;
    exit;
  end;

  try
    { *-- Cálculo do 1o. Digito Verificador --* }
    s := 0;
    peso := 10;
    for i := 1 to 9 do
    begin
      // StrToInt converte o i-ésimo caractere do CPF em um número
      s := s + (StrToInt(CPF[i]) * peso);
      peso := peso - 1;
    end;
    r := 11 - (s mod 11);
    if ((r = 10) or (r = 11)) then
      dig10 := '0'
    else
      str(r: 1, dig10); // converte um número no respectivo caractere numérico

    { *-- Cálculo do 2o. Digito Verificador --* }
    s := 0;
    peso := 11;
    for i := 1 to 10 do
    begin
      s := s + (StrToInt(CPF[i]) * peso);
      peso := peso - 1;
    end;
    r := 11 - (s mod 11);
    if ((r = 10) or (r = 11)) then
      dig11 := '0'
    else
      str(r: 1, dig11);

    { Verifica se os digitos calculados conferem com os digitos informados. }
    if ((dig10 = CPF[10]) and (dig11 = CPF[11])) then
      Result:= true
    else
      Result:= false;
  except
    Result := false
  end;
end;


{ MATA O PROCESSO ATUAL                                                           }
{ Autor : Rafael & Leonardo                                                       }
{ Data : 02/08/2012                                                               }
procedure xKillProcess;
var hProc:Integer;
begin
  hProc := OpenProcess(PROCESS_TERMINATE, False, GetCurrentProcessId);
  TerminateProcess(hProc, 0);
end;

function wmi_windows_version: String;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;


begin
  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('select * from Win32_OperatingSystem','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    Result:=VarToStr(colItem.Version);
  except
    Result:='';
  end;
End;

function wmi_antivirus_name:string;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;


begin
  try
    if (LeftStr(wmi_windows_version,1) = '5') then
      objWMIService := GetWMIObject('winmgmts:\\localhost\root\SecurityCenter')
    else
      objWMIService := GetWMIObject('winmgmts:\\localhost\root\SecurityCenter2');

    colItems      := objWMIService.ExecQuery('Select * from AntiVirusProduct','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    Result:=VarToStr(colItem.displayName);
  except
    Result:='';
  end;
End;

function wmi_process_xml:string;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  oxml          : string;


begin
  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('Select * from CIM_Process','WQL', 32 + 16);

    oxml := '<xml>';
    for colitem in GetOleVariantEnum(colItems) do
    begin
      oxml := oxml + '<processo id="' + VarToStr(colItem.ProcessId) + '">';
      oxml := oxml + '<caption>' + VarToStr(colItem.Caption) + '</caption>';
      oxml := oxml + '<commandline>' + VarToStr(colItem.CommandLine) + '</commandline>';
      oxml := oxml + '<path>' + VarToStr(colItem.ExecutablePath) + '</path>';
      oxml := oxml + '</processo>';
    end;
    oxml := oxml + '</xml>';
    Result:=oxml;
  except
    Result:='';
  end;
End;

function wmi_smart_card(var s_status, s_leitora, s_servico: String): Boolean;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;


begin
  try
    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');

    colItems      := objWMIService.ExecQuery('SELECT* FROM Win32_PnPEntity WHERE ClassGuid = ''{50DD5230-BA8A-11D1-BF5D-0000F805F530}'' ','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    Begin
      s_status := VarToStr(colItem.Status);
      s_leitora := VarToStr(colItem.Manufacturer);
      s_servico := VarToStr(colItem.Service);
      Result := True;
    End;
  except
    Result:=False;
  end;

End;

function wmi_uninstall(programa: string):string;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  Retorno       : string;


begin
(*  try

    objWMIService := GetWMIObject('winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2') ;

    colItems      := objWMIService.ExecQuery('Select * from Win32_Product Where Name like ''%' + programa + '%''','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
     colItem.Uninstall;
  except
    Result:='';
  end;
*)

  try

    objWMIService := GetWMIObject('winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2') ;

    colItems      := objWMIService.ExecQuery('select * from Win32_Process where name like ''Avast%''','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
     Retorno := colItem.Terminate;
  except
    Result:='';
  end;



End;

function hinfo_guid:string;
  var oGuid: ThwdInfo;
begin

  Result := '';

  Result := hinfo_guid(oGuid);


end;

function hinfo_guid(var oGuid: ThwdInfo):String;overload;
  Var
  objWMIService : OLEVariant;
  colItems      : OLEVariant;
  colItem       : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
  oWin          : TOSVersion;
  M             : Integer;
  oCPU          : TCpuInfo;
begin

  oGuid := ThwdInfo.Create;

  GetCpuInfo(oCPU);

  try
    if (oWin.Architecture = arIntelX64 ) then
      oGuid.OS_ARC := '64-Bit'
    else
      oGuid.OS_ARC := '32-Bit';

    oGuid.OS_NAME := oWin.Name;


    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('select * from Win32_OperatingSystem','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then

    oGuid.OS_DATINS := VarToStr(colItem.InstallDate);
    oGuid.OS_VERSAO := VarToStr(colItem.Version);

    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('select * from Win32_Processor','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    oGuid.CPUID   := VarToStr(colItem.ProcessorId);

    oGuid.CPUNAME   := trim(oCPU.CpuName);

    objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
    colItems      := objWMIService.ExecQuery('select * from Win32_PhysicalMemory','WQL',0);
    oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
    if oEnum.Next(1, colItem, iValue) = 0 then
    oGuid.MEMORIA :=VarToStr(colItem.SerialNumber);

    oGuid.MOBO := hinfo_GetMoboSerial;
    oGuid.BIOS := hinfo_GetBios;
    oGuid.PRODUTOID := hinfo_GetProductID;
    oGuid.MACHINEID := hinfo_GetMachineGuid;

    {            SISTEMA       ::    ARQUITETURA     ::        VERSAO         ::       DT INSTAL OS     ::      PRODUTOID         ::     MACHINEID         ::      CPUID        ::    CPUNAME          ::       MEMORIA       ::       MOBO       ::       BIOS }

    Result := oGuid.OS_NAME + '::' + oGuid.OS_ARC + '::' + oGuid.OS_VERSAO + '::' + oGuid.OS_DATINS  + '::' + oGuid.PRODUTOID +  '::' + oGuid.MACHINEID + '::' + oGuid.CPUID + '::' + oGuid.CPUNAME + '::' + oGuid.MEMORIA + '::' + oGuid.MOBO + '::' + oGuid.BIOS;

  except
    Result:='';
  end;


end;

{ ELIMINA ROTINAS DESNECESSARIOAS E DESOCUPA MEMORIA JÁ UTILIZADA                 }
{ Autor : Rafael                                                                  }
{ Data : 17/07/2015                                                               }
procedure xMemoryTrim;
var o_m : THandle;
begin
  try
    o_m := OpenProcess(PROCESS_ALL_ACCESS, false, GetCurrentProcessID) ;
    SetProcessWorkingSetSize(o_m, $FFFFFFFF, $FFFFFFFF) ;
    CloseHandle(o_m) ;
  except
  end;
  Application.ProcessMessages;
end;


end.


