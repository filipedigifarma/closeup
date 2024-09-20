unit UPrincipal;

interface

uses
  Winapi.Windows, System.SysUtils, System.StrUtils,
  System.Classes, System.inifiles, System.Zip,System.DateUtils,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.Mask,
  JvExMask, JvToolEdit, Vcl.ComCtrls, ibx.IBQuery, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls, IdBaseComponent, ShellAPI,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdFTP,     IdFTPList, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  IdIntercept, IdGlobal, System.ImageList, Vcl.ImgList, Vcl.Imaging.pngimage,
  PngSpeedButton, Vcl.Menus;



type
  TConfigProgram=class

  // Farmácia
  g_s_FTP_HOST:String;
  g_s_FTP_USER:String;
  g_s_FTP_PWD:String;
  s_COD_REDE:String;

  // Rede
  FTP_HOST_REDE:String;
  FTP_USER_REDE:String;
  FTP_PWD_REDE:String;
  COD_REDE_REDE:String;

  s_CLIENTE_ID:String;
  s_CNPJ:String;
  s_RAZAO_SOCIAL:String;
  s_FANTASIA:String;
  s_TELEFONE:String;
  Diretorio:String;
  is_REDE_PARCEIRA:Boolean;
  PARCEIRO_ID:INTEGER;
  PARCEIRO:String;
  Dt_ENVIO:String;
  s_RETROATIVO:String;
  data_retroativa:String;
  data_inicio:String;
  data_fim:String;
  atualizar_closeup:String;
  ENVIADO_NOW:String;


  public
  constructor create;
  end;
var
    oConfigProgram:TConfigProgram;
    function str_preenche_zero(Valor: string; Tamanho: Integer): string;
    function str_preenche_espaco(Valor: string; Tamanho: Integer): string;

var
  DataI, DataF,MonthInc: String;
  enviando_arquivo:Boolean;
type
  TFrmPrincipal = class(TForm)
    Gbox: TGroupBox;
    IdFTP: TIdFTP;
    Intercept: TIdConnectionIntercept;
    Socket: TIdSSLIOHandlerSocketOpenSSL;
    jvDateInc: TJvDateEdit;
    IdFTPDigi: TIdFTP;
    intDigi: TIdConnectionIntercept;
    SockDigi: TIdSSLIOHandlerSocketOpenSSL;
    pFooter: TPanel;
    bGerar: TButton;
    ImageList1: TImageList;
    Image1: TImage;
    Label3: TLabel;
    Label4: TLabel;
    JvFim: TJvDateEdit;
    JvInicio: TJvDateEdit;
    Label2: TLabel;
    Label1: TLabel;
    pConfig: TPanel;
    Lb_FTP: TLabel;
    LB_user: TLabel;
    Lb_senha: TLabel;
    Lb_Cod: TLabel;
    Ed_ftp: TEdit;
    Ed_User: TEdit;
    Ed_Senha: TEdit;
    Ed_Cod: TEdit;
    bt_Save: TPngSpeedButton;
    PngSpeedButton1: TPngSpeedButton;
    Splitter1: TSplitter;
    rb_Auto: TRadioButton;
    rb_Man: TRadioButton;
    Tray: TTrayIcon;
    Timer1: TTimer;
    Popup: TPopupMenu;
    Abrir1: TMenuItem;
    Sair1: TMenuItem;
    trintaem30: TTimer;
    Panel1: TPanel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    edtFTP_Rede: TEdit;
    edtUser_Rede: TEdit;
    edtSenha_Rede: TEdit;
    Ed_Cod_rede: TEdit;
    rb_Auto_Rede: TRadioButton;
    rb_Man_Rede: TRadioButton;
    Panel2: TPanel;
    Panel3: TPanel;
    cbxFarmacia: TCheckBox;
    cbxRede: TCheckBox;
    rArq_P: TCheckBox;
    rArq_D: TCheckBox;

    procedure JvInicioChange(Sender: TObject);
    procedure BgerarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BConfigClick(Sender: TObject);
    procedure p_status(Status: String);
    procedure rArq_PClick(Sender: TObject);
    procedure rArq_DClick(Sender: TObject);
    procedure bt_SaveClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure PngSpeedButton1Click(Sender: TObject);
    procedure rb_AutoClick(Sender: TObject);
    procedure rb_ManClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Abrir1Click(Sender: TObject);
    procedure Sair1Click(Sender: TObject);
    procedure trintaem30Timer(Sender: TObject);
    procedure cbxFarmaciaClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure envia_arquivos_closeup_temp;
    procedure deletearquivoscloseup;
    procedure envia_arquivo_Closeup;
    procedure envia_arquivo_Closeup_Rede;
    procedure envia_arquivo_Closeup_Farmacia;

  end;

var
  FrmPrincipal: TFrmPrincipal;

implementation

{$R *.dfm}

uses
  Udm,UConfig,uArqCli,uArqProd,uArqVenda, uMySql, uFuncoes;



function str_preenche_zero(Valor: string; Tamanho: Integer): string;

begin
  Valor := RightStr(StringOfChar('0', Tamanho) + Valor, Tamanho);
  result := Valor;
end;


function str_SoNumero(Const Texto: String): String;
// Remove caracteres de uma string deixando apenas numeros
var
  I: Integer;
  S: string;
begin
  S := '';
  for I := 1 To Length(Texto) Do
  begin
    if (Texto[I] in ['0' .. '9']) then
    begin
      S := S + Copy(Texto, I, 1);
    end;
  end;
  result := S;

end;

function str_preenche_espaco(Valor: string; Tamanho: Integer): string;
var
  x: Integer;
begin
  Valor := LeftStr(Valor + StringOfChar(' ', Tamanho), Tamanho);
  result := Valor;
end;

function GeraArquivoPrescricao: String;
var
  nome_arquivo, cod_crm, tp_crm, dia, mes, ano: String;
  txt: TextFile;
  Tabela: TIBQuery;

begin

  dm.bd_ibQuery_init(Tabela);
  Tabela.SQL.Clear;
  Tabela.SQL.Add('SELECT DISTINCT * FROM (select(select cnpj from config) AS cnpj, p.cod_barras,' +
    ' v.RECEITUARIO_UF AS UF, v.CONSELHO_NUMERO AS NR_CRM, ' +
    ' v.CONSELHO_TIPO AS TP_CRM, I.VENDA_NOTA_ID AS CUPOM,c.venda_data_hora, ' +
    ' I.ITEMVEND_QUANT AS QTDE, p.PRODUTO AS PROD, f.fornecedor as lab, ' +
    ' I.vendedor_id as atend from vendas_psicotropicos v '+
    ' left join item_vendas_lotes IL  on(v.VENDA_NOTA_ID = IL.VENDA_NOTA_ID and v.item_venda_id = IL.item_venda_id)'+
    ' left join item_vendas I on(v.VENDA_NOTA_ID = I.VENDA_NOTA_ID and v.item_venda_id = I.item_venda_id) ' +
    ' left join produtos p on(I.produto_id = p.produto_id) ' +
    ' left join cab_vendas c on(I.VENDA_NOTA_ID = c.VENDA_NOTA_ID) ' +
    ' left join fornecedores f on(p.fornecedor_id = f.fornecedor_id) ' +
    ' where c.venda_data_hora between ' + QuotedStr(DataI) + ' and ' +
    QuotedStr(DataF)+' and v.CONSELHO_NUMERO is not null and v.CONSELHO_TIPO is not null ' +
    ' union all'+
    ' select(select cnpj from config) AS cnpj, ' +
    ' p.cod_barras, a.uf_crm as UF, a.CRM AS NR_CRM, ''CRM'' AS TP_CRM, ' +
    ' c.VENDA_NOTA_ID as CUPOM,c.venda_data_hora, I.ITEMVEND_QUANT QTDE, p.PRODUTO AS PROD, '
    + ' f.fornecedor, a.vendedor_id from autorizacoes a  join cab_vendas c '
    + '  on (a.nsu=c.nsu)and c.venda_data_hora between ' + QuotedStr(DataI) + ' and ' +
    QuotedStr(DataF) + 'and a.status=''3'''+
    ' inner join item_vendas i on(c.VENDA_NOTA_ID = i.VENDA_NOTA_ID) ' +
    ' inner join produtos p on (i.produto_id = p.produto_id) '+
    ' inner join fornecedores f on(p.fornecedor_id = f.fornecedor_id) ' +

    ' union all '+
    ' select(select cnpj from config) AS cnpj, '+
    ' p.cod_barras, i.uf_crmcro as UF, i.num_crmcro AS NR_CRM, '+
    ' CASE conselho_crmcro WHEN 1 then ''CRM'' '+
                         ' WHEN 2 then ''CRO'' '+
                         ' END as TP_CRM, '+
    ' c.VENDA_NOTA_ID as CUPOM,c.venda_data_hora, I.ITEMVEND_QUANT QTDE, p.PRODUTO AS PROD, '+
    ' f.fornecedor, i.vendedor_id from item_vendas i inner join cab_vendas c '+
    ' on (i.venda_nota_id=c.venda_nota_id)and c.venda_data_hora between ' + QuotedStr(DataI)
     + ' and ' +QuotedStr(DataF)+
    ' inner join produtos p on (i.produto_id = p.produto_id) '+
    ' inner join fornecedores f on(p.fornecedor_id = f.fornecedor_id) '+
    ' where i.uf_crmcro is not null and i.num_crmcro is not null and conselho_crmcro is not null) ORDER BY VENDA_DATA_HORA '+

     ' ;');
   //Tabela.SQL.SaveToFile('c:\install\saida.sql');
   Tabela.Open;

   if (Tabela.Eof) then
   nome_arquivo:=''
   else
   begin
    nome_arquivo := oConfigProgram.s_COD_REDE + FormatDateTime('yymmdd', frmprincipal.jvDateInc.Date) + '.txt';
    AssignFile(txt, IncludeTrailingBackslash(oConfigProgram.Diretorio) + nome_arquivo);
    Rewrite(txt);


    while (Tabela.Eof = false) do
    begin

      cod_crm := Tabela.FieldByName('uf').AsString;

      if (Tabela.FieldByName('uf').AsString = 'RJ') then
      Begin
        cod_crm := '052';
      end
      else
      begin
        cod_crm := '000';
      end;

      { Convertendo o tipo do crm para M= Médico / V= Veterinário / O= Dentista. }
      tp_crm := Tabela.FieldByName('tp_crm').AsString;

      if (Tabela.FieldByName('tp_crm').AsString = 'CRM') then
      Begin
        tp_crm := 'M';
      End;

      if (Tabela.FieldByName('tp_crm').AsString = 'CRMV') then
      Begin
        tp_crm := 'V';
      End;

      if (Tabela.FieldByName('tp_crm').AsString = 'CRO') then
      Begin
        tp_crm := 'O';
      End;

      if (Tabela.FieldByName('tp_crm').AsString = 'RMS') then
      Begin
        tp_crm := 'S';
      End;

      if (Tabela.FieldByName('tp_crm').AsString = 'CRF') then
      Begin
        tp_crm := 'F';
      End;

      dia := FormatDateTime('dd', (Tabela.FieldByName('venda_data_hora').AsDateTime));
      mes := FormatDateTime('mm', (Tabela.FieldByName('venda_data_hora').AsDateTime));
      ano := FormatDateTime('yyyy',(Tabela.FieldByName('venda_data_hora').AsDateTime));

      { Escrevendo o Arquivo }
      Writeln(txt, str_preenche_zero(oConfigProgram.s_COD_REDE, 4) + str_preenche_zero('1', 4) +
        str_preenche_zero(Tabela.FieldByName('cnpj').AsString, 14) +
        str_preenche_zero(Tabela.FieldByName('cod_barras').AsString, 20) +
        Tabela.FieldByName('uf').AsString + cod_crm +
        str_preenche_zero(Tabela.FieldByName('nr_crm').AsString, 7) + tp_crm +
        str_preenche_zero(Tabela.FieldByName('cupom').AsString, 10) +
        str_preenche_zero(dia, 2) + str_preenche_zero(mes, 2) +
        str_preenche_zero(ano, 4) + str_preenche_zero(Tabela.FieldByName('qtde')
        .AsString, 6) + str_preenche_espaco(Tabela.FieldByName('prod').AsString,
        50) + str_preenche_espaco(Tabela.FieldByName('lab').AsString, 50) + 'M'
        + str_preenche_zero(Tabela.FieldByName('atend').AsString, 10));

      Tabela.Next;
    end;

    CloseFile(txt);
   end;
  result := nome_arquivo;
end;


function gerarArquivoDemanda:String;
var
  s_ArqCli,s_ArqProd,s_ArqVenda,Nome_Zip: String;
  txt: TextFile;
  Tabela: TIBQuery;
  Zip:TzipFile;
begin
// s_ArqCli:= gerarArqCli;
// s_ArqProd:= gerarArqProd;
  s_ArqVenda:= gerarArqVenda;
// Nome_Zip:=oConfigProgram.s_COD_REDE+'M'+FormatDateTime('mm', frmprincipal.jvDateInc.Date)+'.D'+FormatDateTime('dd', frmprincipal.jvDateInc.Date)+'.Zip';
// Zipa o Arquivo
//  Zip := TZipFile.Create;
//  Zip.Open(Nome_Zip, zmWrite);
//  Zip.Add(IncludeTrailingBackslash(oConfigProgram.Diretorio) +s_ArqCli, s_ArqCli);
//  Zip.Add(IncludeTrailingBackslash(oConfigProgram.Diretorio) +s_ArqProd, s_ArqProd);
//  Zip.Add(IncludeTrailingBackslash(oConfigProgram.Diretorio) +s_ArqVenda, s_ArqVenda);
//  Zip.Close;

//  DeleteFile(IncludeTrailingBackslash(oConfigProgram.Diretorio) +s_ArqCli);
//  DeleteFile(IncludeTrailingBackslash(oConfigProgram.Diretorio) +s_ArqProd);
//  DeleteFile(IncludeTrailingBackslash(oConfigProgram.Diretorio) +s_ArqVenda);
  Result := s_ArqVenda;
end;



procedure TFrmprincipal.p_status(Status: String);
begin
  FrmPrincipal.Caption := Status;
  Application.ProcessMessages;
end;

procedure TFrmPrincipal.rArq_DClick(Sender: TObject);
VAR
Arquivo:TiniFile;
begin
 Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');

// if (Arquivo.ReadString('CLOSEUP','DEMANDA','')='S') then
// Arquivo.WriteString ('CLOSEUP', 'DEMANDA', 'N')
// else
 Arquivo.WriteString ('CLOSEUP', 'DEMANDA', 'S');

 freeandnil(Arquivo);
end;

procedure TFrmPrincipal.rArq_PClick(Sender: TObject);
var
Arquivo:TIniFile;
begin
 Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');

// if (Arquivo.ReadString('CLOSEUP','PRESCRICAO','')='S') then
// Arquivo.WriteString ('CLOSEUP', 'PRESCRICAO', 'N')
// else
 Arquivo.WriteString ('CLOSEUP', 'PRESCRICAO', 'S');

 freeandnil(Arquivo);
end;

procedure TFrmPrincipal.rb_AutoClick(Sender: TObject);
var
Arquivo:TIniFile;
begin
 Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');
 Arquivo.WriteString ('CLOSEUP', 'CLOSEUP_ENVIO', 'AUTO');
 freeandnil(Arquivo);
end;

procedure TFrmPrincipal.rb_ManClick(Sender: TObject);
var
Arquivo:TIniFile;
begin
 Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');
 Arquivo.WriteString ('CLOSEUP', 'CLOSEUP_ENVIO', 'MANUAL');
 freeandnil(Arquivo);
end;

procedure TFrmPrincipal.Sair1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TFrmPrincipal.Timer1Timer(Sender: TObject);
begin
  if enviando_arquivo then Exit;

  if (dm.CONECTADO) then
  envia_arquivo_Closeup
  else
  exit;
end;

procedure TFrmPrincipal.BgerarClick(Sender: TObject);
var
  I,n_a,d_f: Integer;
  Arquivo_gerado,UltEnvio: String;
  Arquivo: TINIFile;
  oBanco,oBancoAux:TDataXml;
  b_ftp_pasta_ano:boolean;

function conectaFTP:Boolean;
begin
  try
    dm.p_gravaLog('Conectando FTP', -1);
    if (IdFTPDigi.Connected) then
    IdFTPDigi.Disconnect;

    IdFTPDigi.Connect;
    IdFTPDigi.List;
    dm.p_gravaLog('Conectado FTP', -1);
    result:=true;
  except
  on ex:Exception do
    begin
      dm.p_gravaLog('Erro: ' + ex.Message + ' Aguardando 1 minuto',-1);
      Sleep(60000);
      conectaFTP;
    end;
  end;
end;

function anviaFTP:Boolean;
var mFile: string;
begin

  try
    dm.p_gravaLog('Enviando arquivo ao FTP', -1);

    mFile := IncludeTrailingBackslash(oConfigProgram.Diretorio) + Arquivo_gerado;

    if oConfigProgram.is_REDE_PARCEIRA then
    IdFTPDigi.Put(mFile,  Arquivo_gerado)
    else
    IdFTP.Put(mFile, Arquivo_gerado);

    dm.p_gravaLog('Arquivo enviado FTP', -1);
    IdFTPDigi.Disconnect;
    dm.p_gravaLog('Desconectando...', -1);

   {Escreve a ultima data no INI}
    try
      Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');
      Arquivo.WriteString ('CLOSEUP', 'DT_ULT_ENVIO', jvfim.Text);
      Label4.Caption := jvfim.Text;
      Arquivo.Free;
    except
      on e:exception do
      begin
        MessageDlg('Erro ao Gravar no ini, informe este erro ao nosso suporte ' + e.Message, mtError, [mbOK], 0);
      end;
    end;
   {-------------------------------------}

    result := True;
    App_RegisterDigifarma('Closeup');
  except
    on ex:Exception do
    begin
      dm.p_gravaLog('Erro: ' + ex.Message + ' Aguardando 1 minuto',-1);

      if not DirectoryExists(ExtractFileDir(Application.ExeName) + '\temp\' ) then
      ForceDirectories(ExtractFileDir(Application.ExeName) + '\temp\');

      CopyFile(PChar(mFile), PChar(ExtractFileDir(Application.ExeName) + '\temp\' + Arquivo_gerado), false);


      Sleep(60000);
      anviaFTP;
    end;

  end;
end;

procedure GerarScriptFTP(const FileName, ServerName, UserName, Password, FilenameToSend, ScriptFileName: string);
var
  BatFile: TextFile;
begin
  AssignFile(BatFile, FileName);
  Rewrite(BatFile);
  try
    Writeln(BatFile, 'open ' + ServerName);
    Writeln(BatFile, UserName);
    Writeln(BatFile, Password);
    Writeln(BatFile, 'binary');
    Writeln(BatFile, 'prompt');
    Writeln(BatFile, 'mput ' + FilenameToSend);
    Writeln(BatFile, 'quit');
  finally
    CloseFile(BatFile);
  end;
end;


begin
  {$REGION 'FOR QUE PERCORRE OS DIAS DA GERAÇÃO DO ARQUIVO'}
  for n_a := 0 to DaysBetween(JvInicio.Date, JvFim.Date) do
  begin
    enviando_arquivo := True;
    jvDateInc.Date   := IncDay(jvInicio.Date,n_a);
    dm.p_gravaLog('Eviando arquivo do dia:' + jvDateInc.Text, -1);

    if cbxRede.Checked then
    begin
      if (oConfigProgram.s_RETROATIVO = 'S') then
      begin
        if MonthInc <> FormatDateTime('mm',jvDateInc.Date) then
        begin
          if (MonthInc <> '') then
          begin
            if dm.CONECTADO then
            begin
              dm.p_gravaLog('Atualizando dia:'+jvDateInc.Text+' No servidor',-1);
              oBancoAux:=TDataXml.Create;
              oBancoAux.ServidorWS('www.digifarma.com.br');
              oBancoAux.SQL:='update ifarma.closeup_lojas set data_inicio=:data_inicio where cnpj='+QuotedStr(oConfigProgram.s_CNPJ);
              oBancoAux.ParamName['data_inicio']:=FormatDateTime('yyyy-mm-dd',jvDateInc.Date);
              oBancoAux.Execute;
              FreeAndNil(oBancoAux);
            end;
          end;
          MonthInc := FormatDateTime('mm',jvDateInc.Date);
        end;
      end;
    end;

    {$REGION 'Gera arquivo de acordo com o layout flat file de prescrição'}
    if (rArq_P.Checked) then
    begin
      dm.p_gravaLog('Gerando arquivo de Prescrição', -1);
      DataI := FormatDateTime('yyyy-mm-dd', jvDateInc.Date) + ' 00:00:00';
      DataF := FormatDateTime('yyyy-mm-dd', jvDateInc.Date) + ' 23:59:59';

      if (str_SoNumero(JvInicio.Text) = '') then
      begin
        if (rb_Man.Checked) then
          MessageBox(0, 'Favor informar o periodo', '', MB_OK + MB_ICONERROR);

        exit;
      end;

      if (JvFim.Date > now) or (StrToInt(FormatDateTime('yyyy', JvInicio.Date))
        < 2015) then
      begin
        if (rb_Man.Checked) then
          MessageBox(0, 'Período inválido', '', MB_OK + MB_ICONERROR);

        exit;
      end;

      p_status('Aguarde, gerando arquivo...');
      Arquivo_gerado := GeraArquivoPrescricao;
      dm.p_gravaLog('Arquivo Gerado'+Arquivo_gerado ,-1);

      if (trim(Arquivo_gerado)<>'') then
      begin
        dm.p_gravaLog('Transmitindo.... ' + Arquivo_gerado ,-1);
        p_status('Gerado com sucesso!');

        {***** Manda o Arquivo ao FTP *****}
        if (rb_Man.Checked) then
        if (MessageBox(0,'Arquivo gerado com sucesso, deseja transmitir agora?','',MB_YESNOCANCEL+MB_ICONQUESTION)=IDNO)then
        begin
          p_status('Não enviado');
          MessageDlg('Arquivo Não enviado',mtError,[mbOK],0);
          DeleteFile(IncludeTrailingBackslash(oConfigProgram.Diretorio)+Arquivo_gerado);
          exit;
        end;

        {***** Manda o Arquivo ao FTP *****}
        try
          if (dm.CONECTADO) then
          begin
            dm.p_gravaLog('Acessando FTP', -1);
            GerarScriptFTP('script_ftp.bat', 'ftp.close-upinternational.com.br', oConfigProgram.g_s_FTP_USER, oConfigProgram.g_s_FTP_PWD, Arquivo_gerado, 'ftptempscript');
            ShellExecute(0, nil, 'cmd.exe', PChar('/c ' + 'ftp -s:"script_ftp.bat"'), nil, SW_HIDE);


            DeleteFile(oConfigProgram.Diretorio+Arquivo_gerado);
          end
          else
          begin
             dm.p_gravaLog('Sem Internet aguardando 10 minutos', -1);
             Sleep(600000);
             envia_arquivo_Closeup;
          end;
          except
          on ex:exception do
          begin
            dm.p_gravaLog('Erro ao inserir no FTP '+ex.Message, -1);
            Sleep(600000);
            envia_arquivo_Closeup;
          end;

        end;
      end;
      dm.p_gravaLog('Terminei prescrição',-1);
      {---------------------------------------------------------}
    end;{Fim - Arquivo Prescrição}
    {$ENDREGION}

    {$REGION 'Gera o Arquivo de acordo com o manual de Demanda'}
    if (rArq_D.Checked) then
    begin
      dm.p_gravaLog('Gerando arquivo de Demanda',-1);
      DataI := FormatDateTime('yyyy-mm-dd', JvInicio.Date) + ' 00:00:00';
      DataF := FormatDateTime('yyyy-mm-dd', JvFim.Date) + ' 23:59:59';

      if (str_SoNumero(JvInicio.Text) = '') then
      begin
        if (rb_Man.Checked) then
        MessageBox(0, 'Favor informar o periodo', '', MB_OK + MB_ICONERROR);
        Exit;
      end;

      if (JvFim.Date>now) or (StrToInt(FormatDateTime('yyyy',JvInicio.Date))< 2012) then
      begin
        if (rb_Man.Checked) then
        MessageBox(0, 'Período inválido', '', MB_OK + MB_ICONERROR);
        Exit;
      end;

      p_status('Aguarde, gerando arquivo...');
      Arquivo_gerado := gerarArquivoDemanda;
      dm.p_gravaLog('Arquivo Gerado ' + arquivo_gerado, -1);

      if (arquivo_gerado = '') then
      begin
        dm.p_gravaLog('Sem arquivos pra enviar', -1);
        if (rb_Man.Checked) then
        begin
          MessageBox(0,'Sem Arquivos a enviar','',MB_OK);
          exit;
        end;
      end
      else
      begin

        p_status('Gerado com sucesso!');

       {***** Manda o Arquivo ao FTP *****}
        try
          if (dm.CONECTADO) then
          begin
            dm.p_gravaLog('Acessando FTP', -1);

            IdFTPDigi.Disconnect;
            IdFTPDigi.Host     := oConfigProgram.FTP_HOST_REDE;
            IdFTPDigi.UserName := oConfigProgram.FTP_USER_REDE;
            IdFTPDigi.Password := oConfigProgram.FTP_PWD_REDE;

            IdFTPDigi.Passive := true;
            IdFTPDigi.UseTLS := utUseExplicitTLS;
            conectaFTP;

            b_ftp_pasta_ano := false;

            try
              IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
            except
              on ex: Exception do
              begin
                IdFTPDigi.MakeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
                IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
              end;
            end;

            dm.p_gravaLog('Criando pasta inexistente no servidor', -1);
            try
              IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO +
                '/' + FormatDateTime('yyyy', jvDateInc.Date));
            except
              on ex: Exception do
              begin
                IdFTPDigi.MakeDir(FormatDateTime('yyyy', jvDateInc.Date));
                IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO +
                  '/' + FormatDateTime('yyyy', jvDateInc.Date));
              end;
            end;


//            GerarScriptFTP('script_ftp.bat',
//            'digifarmaonline.com.br',
//            oConfigProgram.FTP_USER_REDE,
//            oConfigProgram.FTP_PWD_REDE,
//            Arquivo_gerado, 'ftptempscript');
//            ShellExecute(0, nil, 'cmd.exe', PChar('/c ' + 'ftp -s:"script_ftp.bat"'), nil, SW_HIDE);

            anviaFTP;

            DeleteFile(oConfigProgram.Diretorio + Arquivo_gerado);
          end
          else
          begin
            dm.p_gravaLog('Sem Internet aguardando 10 minutos', -1);
            Sleep(600000);
            envia_arquivo_Closeup;
          end;
        except
          on ex: Exception do
          begin
            dm.p_gravaLog('Erro ao inserir no FTP ' + ex.Message, -1);
            Sleep(600000);
            envia_arquivo_Closeup;
          end;
        end;
        { ---------------------------------------------------------}

       {Escreve a ultima data no INI}
        try
          dm.p_gravaLog('Acessando INI ',-1);
          Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');
          Arquivo.WriteString ('CLOSEUP', 'DT_ULT_ENVIO', jvfim.Text);

          Label4.Caption:=jvfim.Text;
          Arquivo.Free;
        except
          on e:exception do
          begin
            if (rb_Man.Checked) then
            MessageDlg('Erro ao Gravar no ini, informe este erro ao nosso suporte '+e.Message,mtError,[mbOK],0);
          end;
        end;{Fim - Ultimo Envio no INI}

      end;

    end;
    {$ENDREGION}

    if not rArq_P.Checked and not rArq_D.Checked then
    Exit;

  end;{Final do Período da Data escolhida nos JVdates}
  enviando_arquivo := False;
  {$ENDREGION}

  {$REGION 'SETA O RETROATIVO PRA N NA NUVEM E ATUALIZA O STATUS DO ARQUIVO'}

  if (dm.CONECTADO) then
  begin
    try
      if (oConfigProgram.is_REDE_PARCEIRA) then
      begin
        try
          { Muda o Retroativo para N no banco online }
          dm.p_gravaLog('Atualizando Informações no Servidor', -1);
          if (oConfigProgram.s_RETROATIVO = 'S') then
          begin
            oBancoAux := TDataXml.create;
            oBancoAux.ServidorWS('www.digifarma.com.br');
            oBancoAux.SQL :=
              'update ifarma.closeup_lojas set retroativo=''N'',data_inicio=null,data_fim=null where cnpj='
              + QuotedStr(oConfigProgram.s_CNPJ);

            try
              oBancoAux.Execute;
            except
              on ex: Exception do
              begin
                dm.p_gravaLog('Erro ao acessar Servidor' + ex.Message, -1);
                Sleep(300000);
                envia_arquivo_Closeup;
              end;
            end;
            freeandnil(oBancoAux);
          end;

          oBancoAux := TDataXml.create;
          oBancoAux.ServidorWS('www.digifarma.com.br');
          oBancoAux.SQL :=
            'INSERT INTO ifarma.closeup_arquivos (data_envio,enviado,dia_referente,cnpj,versao) VALUES('
            + QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', now)) + ', ' +
            QuotedStr('S') + ', ' +
            QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', jvDateInc.Date)) +
            ', ' + QuotedStr(oConfigProgram.s_CNPJ) + ', ' +
            QuotedStr(str_GetFileVer('c:\digifarma\closeup\closeup.exe')) + ' )'

            + 'on duplicate key update data_envio=sysdate(), enviado=''S'', dia_referente='
            + QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', jvDateInc.Date)) +
            ',' + 'versao=' +
            QuotedStr(str_GetFileVer('c:\digifarma\closeup\closeup.exe'));

          oBancoAux.Execute;
          freeandnil(oBancoAux);
        except
          on ex: Exception do
            MessageDlg('erro ao inserir online ' + ex.Message, mtError,
              [mbOK], 0);
        end;
      end;
    except
      on ex: Exception do
      begin
        dm.p_gravaLog('Erro ao acessar Servidor' + ex.Message, -1);
        Sleep(600000);
        envia_arquivo_Closeup;
      end;

    end;
  end;
  {$ENDREGION}
  enviando_arquivo := false;

  if (rb_Man.Checked = false) then
  begin
    exit;
  end;
end;

procedure TFrmPrincipal.bt_SaveClick(Sender: TObject);
var
  Tabela: Tibquery;
begin
if (Ed_ftp.Text<>'')and(Ed_User.Text<>'') and (Ed_Senha.Text<>'')and(Ed_Cod.Text<>'') then
 begin
  dm.bd_ibQuery_init(Tabela);
  Tabela.SQL.Clear;
  Tabela.SQL.Add('Select count(*)as QTD from close_up');
  Tabela.Open;
  if (Tabela.FieldByName('qtd').AsInteger <= 0) then
  begin
    dm.bd_ibQuery_init(Tabela);
    Tabela.SQL.Clear;
    Tabela.SQL.Add('insert into close_up(ftp,usuario,senha,cod_rede) values (' +
      QuotedStr(Ed_ftp.Text) + ', ' + QuotedStr(Ed_User.Text) + ', ' +
      QuotedStr(Ed_Senha.Text) +', '+QuotedStr(Ed_Cod.Text) +')');
    Tabela.ExecSQL;
    Tabela.Transaction.CommitRetaining;
    freeandnil(Tabela);

    dm.bd_ibQuery_init(Tabela);
    Tabela.SQL.Clear;
    Tabela.SQL.Add('select C.*, COUNT(*)AS QTD from close_up C GROUP BY C.USUARIO,C.FTP,C.SENHA,C.COD_REDE');
    Tabela.Open;

    oConfigProgram.g_s_FTP_HOST := Tabela.FieldByName('ftp').AsString;
    oConfigProgram.g_s_FTP_USER := Tabela.FieldByName('usuario').AsString;
    oConfigProgram.g_s_FTP_PWD := Tabela.FieldByName('senha').AsString;
    oConfigProgram.s_COD_REDE := Tabela.FieldByName('cod_rede').AsString;
    frmprincipal.JvInicio.Enabled:=true;
    frmprincipal.JvFim.Enabled:=true;
    frmprincipal.Bgerar.Enabled:=true;
    freeandnil(Tabela);
  end
  else
  begin
    dm.bd_ibQuery_init(Tabela);
    Tabela.SQL.Clear;
    Tabela.SQL.Add('Update Close_up set ftp=' + QuotedStr(Ed_ftp.Text) +
      ', Usuario=' + QuotedStr(Ed_User.Text) + ', Senha=' +
      QuotedStr(Ed_Senha.Text)+', Cod_Rede='+QuotedStr(Ed_Cod.Text));
    Tabela.ExecSQL;
    Tabela.Transaction.CommitRetaining;
  end;
 end;


end;


procedure TFrmPrincipal.cbxFarmaciaClick(Sender: TObject);
var
Arquivo:TIniFile;
begin
  Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');
  Arquivo.WriteString ('CLOSEUP', 'CLOSEUP_LOCAL', 'FARMACIA');
  freeandnil(Arquivo);
end;

procedure TFrmPrincipal.deletearquivoscloseup;
  var
  i: integer;
  sr: TSearchRec;
  arquivo:String;
begin

  begin
    I := FindFirst('C:\Digifarma\CloseUp\*.*', faAnyFile, SR);
    while I = 0 do
    begin
      Arquivo:='C:\Digifarma\CloseUp\' + SR.Name;


      if (Arquivo<>'C:\Digifarma\CloseUp\libeay32.dll') and (Arquivo<>'C:\Digifarma\CloseUp\libssl32.dll')
      and (Arquivo<>'C:\Digifarma\CloseUp\script_ftp.bat')
      and (Arquivo<>'C:\Digifarma\CloseUp\ssleay32.dll') and (Arquivo<>'C:\Digifarma\CloseUp\MSVCR100.DLL') then
      begin
        DeleteFile('C:\Digifarma\CloseUp\' + SR.Name);
      end;

       I := FindNext(SR);
    end;
  end;

end;


procedure TFrmPrincipal.envia_arquivos_closeup_temp;
var
  i: Integer;
  sr: TSearchRec;
  arquivo: String;
  ftpConectado: Boolean;

function EnviaArquivoFTP(arquivo: string): Boolean;
var
  mFile: string;
begin
  Result := False;
  try
    dm.p_gravaLog('Enviando arquivo pasta temp ao FTP: ' + arquivo, -1);

    mFile := IncludeTrailingBackslash(oConfigProgram.Diretorio) + 'Temp\' + arquivo;

    if oConfigProgram.is_REDE_PARCEIRA then
      IdFTPDigi.Put(mFile, arquivo)
    else
      IdFTP.Put(mFile, arquivo);

    dm.p_gravaLog('Arquivo pasta temp enviado com sucesso ao FTP', -1);
    IdFTPDigi.Disconnect;
    dm.p_gravaLog('Desconectado do FTP', -1);

    App_RegisterDigifarma('Closeup');
    Result := True;
  except
    on E: Exception do
    begin
      dm.p_gravaLog('Erro ao enviar arquivo: ' + E.Message + '. Tentando novamente em 1 minuto', -1);

      var mDir := 'C:\Digifarma\CloseUp\temp\';

      // Cria diretório temporário, se não existir
      if not DirectoryExists(mDir) then ForceDirectories(mDir);

      // Copia o arquivo para o diretório temporário
      CopyFile(PChar(mFile), PChar(mDir + arquivo), False);

      Sleep(60000); // Espera 1 minuto antes de tentar novamente
      EnviaArquivoFTP(arquivo);
    end;
  end;
end;

function ConectaFTP: Boolean;
begin
  Result := False;
  try
    dm.p_gravaLog('Conectando ao FTP', -1);

    if IdFTPDigi.Connected then
      IdFTPDigi.Disconnect;

    IdFTPDigi.Connect;
    IdFTPDigi.List;
    dm.p_gravaLog('Conectado ao FTP', -1);
    Result := True;
  except
    on E: Exception do
    begin
      dm.p_gravaLog('Erro ao conectar: ' + E.Message + '. Tentando novamente em 1 minuto', -1);
      Sleep(60000); // Espera 1 minuto antes de tentar novamente
      ConectaFTP;
    end;
  end;
end;

procedure CriaDiretoriosFTP;
begin
  try
    IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
  except
    on E: Exception do
    begin
      IdFTPDigi.MakeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
      IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
    end;
  end;

  try
    IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO + '/' + FormatDateTime('yyyy', Now));
  except
    on E: Exception do
    begin
      IdFTPDigi.MakeDir(FormatDateTime('yyyy', Now));
      IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO + '/' + FormatDateTime('yyyy', Now));
    end;
  end;
end;

begin
  if FindFirst('C:\Digifarma\CloseUp\Temp\*.*', faAnyFile, sr) = 0 then
  begin
    repeat
      arquivo := sr.Name;

      if (arquivo <> '.') and (arquivo <> '..') then
      begin
        if dm.CONECTADO then
        begin
          dm.p_gravaLog('Acessando FTP para envio de arquivo da pasta Temp: ' + arquivo, -1);

          // Configura o FTP
          IdFTPDigi.Host     := oConfigProgram.FTP_HOST_REDE;
          IdFTPDigi.UserName := oConfigProgram.FTP_USER_REDE;
          IdFTPDigi.Password := oConfigProgram.FTP_PWD_REDE;
          IdFTPDigi.Passive  := True;
          IdFTPDigi.UseTLS   := utUseExplicitTLS;

          // Conecta ao FTP
          ftpConectado := ConectaFTP;
          if ftpConectado then
          begin
            // Cria diretórios no FTP
            CriaDiretoriosFTP;

            // Envia o arquivo
            if EnviaArquivoFTP(arquivo) then
            begin
              DeleteFile(oConfigProgram.Diretorio + '\Temp\' + arquivo);
              dm.p_gravaLog('Arquivo deletado localmente após envio da pasta Temp: ' + arquivo, -1);
            end;
          end;
        end;
      end;

      i := FindNext(sr);
    until i <> 0;
  end;
  FindClose(sr);
end;


procedure TFrmPrincipal.envia_arquivo_Closeup;
begin
   enviando_arquivo := False;

   {$REGION 'ENVIO AUTOMÁTICO CLIENTE DE REDE'}
   if (cbxRede.Checked) and (enviando_arquivo = False) then
   begin
     enviando_arquivo := True;
     envia_arquivo_Closeup_Rede;
     enviando_arquivo := False;
     dm.p_gravaLog('Envio automatico de cliente de rede',-1);
   end;
   {$ENDREGION}

   {$REGION 'ENVIO AUTOMÁTICO CLIENTE COMUM'}
   if (cbxFarmacia.Checked) and (enviando_arquivo = False) then
   begin
     dm.p_gravaLog('Envio automatico de cliente comum farmácia',-1);
     enviando_arquivo := True;
     envia_arquivo_Closeup_Farmacia;
     enviando_arquivo := False;
   end;
   {$ENDREGION}

   if (rb_Auto.Checked) then
   begin
      frmPrincipal.Tray.Visible := True;
      Application.ShowMainForm := False;
   end;

end;

procedure TFrmPrincipal.envia_arquivo_Closeup_Farmacia;
var
  Tabela: TIBQuery;
  txt:TextFile;
  Arquivo: TINIFile;
  oBanco,oBancoAux:TDataXml;
begin
  try
    dm.bd_ibQuery_init(Tabela);
    Tabela.SQL.Clear;
    Tabela.SQL.Add('select C.*, COUNT(*)AS QTD,CO.CNPJ,CO.RAZAO_SOCIAL from close_up C,'
    + ' CONFIG CO GROUP BY CO.CNPJ, CO.RAZAO_SOCIAL, C.USUARIO, C.FTP, C.SENHA, C.COD_REDE');
    Tabela.Open;
    if (Tabela.FieldByName('qtd').AsInteger=0) then
    begin
      Messagebox(0,'Seu CloseUp está desconfigurado, favor clicar em configurações antes de continuar,'+
      ' qualquer dúvida entre em contato em nosso TeleSuporte','',MB_OK+MB_ICONWARNING);
      JvInicio.Enabled := false;
      JvFim.Enabled    := false;
      Bgerar.Enabled   := false;
    end
    else
    begin
      if not Assigned(oConfigProgram) then
      oConfigProgram:=TConfigProgram.Create;
      oConfigProgram.g_s_FTP_HOST := Tabela.FieldByName('ftp').AsString;
      oConfigProgram.g_s_FTP_USER := Tabela.FieldByName('usuario').AsString;
      oConfigProgram.g_s_FTP_PWD  := Tabela.FieldByName('senha').AsString;
      oConfigProgram.s_COD_REDE   := Tabela.FieldByName('cod_rede').AsString;
      oConfigProgram.Diretorio    := IncludeTrailingBackslash('C:\digifarma')+'CloseUp';

      if not DirectoryExists(oConfigProgram.Diretorio) then
      ForceDirectories(oConfigProgram.Diretorio);
    end;
    except
    on E:Exception do
    begin
      JvInicio.Enabled := False;
      JvFim.Enabled    := False;
      bGerar.Enabled   := False;
    end;
  end;

  FreeAndNil(Tabela);

  try
    Arquivo := TINIFile.Create('c:\Digifarma\digifarma.ini');

    if Arquivo.ReadString ('CLOSEUP', 'PRESCRICAO', ' ')    = 'S' then
    rArq_P.Checked  := True ;
    if Arquivo.ReadString ('CLOSEUP', 'DEMANDA', ' ')       = 'S' then
    rArq_D.Checked  := True ;
    if Arquivo.ReadString ('CLOSEUP', 'CLOSEUP_ENVIO', ' ') = 'AUTO' then
    rb_Auto.Checked :=  True;
    if Arquivo.ReadString ('CLOSEUP', 'CLOSEUP_ENVIO', ' ') = 'MANUAL' then
    rb_Man.Checked  := True;

    Arquivo.Free;
  Except
    on ex: Exception do
    begin
      ShowMessage('Erro: '+ex.Message);
    end;
  end;

  Ed_ftp.Text     := oConfigProgram.g_s_FTP_HOST;
  Ed_User.Text    := oConfigProgram.g_s_FTP_USER;
  Ed_Senha.Text   := oConfigProgram.g_s_FTP_PWD;
  Ed_Cod.Text     := oConfigProgram.s_COD_REDE;

  if (rb_Auto.Checked) and  ((strtodate(oConfigProgram.Dt_ENVIO)-1)<StrToDate(FormatDateTime('dd/mm/yyyy',now))) then
  begin
    jvInicio.Date := StrToDate(oConfigProgram.Dt_ENVIO) - 1;
    jvFim.Date    := StrToDate(oConfigProgram.Dt_ENVIO) - 1;
    bGerar.Click;
  end;
end;

procedure TFrmPrincipal.envia_arquivo_Closeup_Rede;
var
  Tabela: TIBQuery;
  txt:TextFile;
  Arquivo: TINIFile;
  oBanco,oBancoAux:TDataXml;
begin
   Edtftp_rede.Text   := oConfigProgram.FTP_HOST_REDE;
   edtUser_Rede.Text  := oConfigProgram.FTP_USER_REDE;
   edtSenha_Rede.Text := oConfigProgram.FTP_PWD_REDE;
   Ed_Cod_rede.Text   := oConfigProgram.COD_REDE_REDE;

   if (rb_Auto.Checked) and ((strtodate(oConfigProgram.Dt_ENVIO) -1 ) < StrToDate(FormatDateTime('dd/mm/yyyy',now))) then
   begin
      if (oConfigProgram.s_retroativo = 'S') then
      begin
        jvInicio.Date := strtodate(oConfigProgram.data_inicio);
        JvFim.Date    := strtodate(oConfigProgram.data_fim);
      end
      else
      begin
        jvInicio.Date := StrToDate(oConfigProgram.Dt_ENVIO) - 1;
        jvFim.Date    := StrToDate(oConfigProgram.Dt_ENVIO) - 1;
      end;

      try
         if (oConfigProgram.ENVIADO_NOW <> '') then
         begin
            if (StrToDate(oConfigProgram.ENVIADO_NOW)=date-1) then
            begin
              dm.p_gravaLog('Já enviou o arquivo de ontem terminando rotina', -1);
              enviando_arquivo := False;
              Exit;
            end;
         end;

      except

      end;

      bGerar.Click;
   end;
end;

procedure TFrmPrincipal.Abrir1Click(Sender: TObject);
begin
  FrmPrincipal.Show;
end;

procedure TFrmPrincipal.BConfigClick(Sender: TObject);
begin
  if (pConfig.Width = 0) then
  begin
    pConfig.Width := 165;
    pConfig.Realign;
    pConfig.Repaint;
    pConfig.Refresh;
  end
  else
  begin
    pConfig.Width := 0;
    pConfig.Realign;
    pConfig.Repaint;
    pConfig.Refresh;
  end;
end;



procedure TFrmPrincipal.FormCreate(Sender: TObject);
begin
  if (dm.CONECTADO) then
  begin
    dm.p_gravaLog('Gerando arquivo Closeup', -1);

    enviando_arquivo := true;
    dm.p_gravaLog('deletando arquivos anteriores Closeup', -1);
    deletearquivoscloseup;

    {Carrega o Objeto de configuração do programa}
    oConfigProgram := TConfigProgram.Create;

    envia_arquivos_closeup_temp;
    envia_arquivo_Closeup;
  end
  else
  begin
    dm.p_gravaLog('Sem internet aguardando 10 minutos', -1);
    Sleep(600000);
    dm.p_gravaLog('Gerando arquivo Closeup', -1);
    envia_arquivo_Closeup;
  end;
end;

procedure TFrmPrincipal.FormShow(Sender: TObject);
begin
  if (rb_Auto.Checked) then
  begin
    Application.ShowMainForm:=false;
  end;
end;

procedure TFrmPrincipal.trintaem30Timer(Sender: TObject);
var
oBanco:TDataxml;
s_retroativo:String;
begin
 if Assigned(oConfigProgram) then
 begin

   if  (enviando_arquivo = false) and (trim(oConfigProgram.s_CNPJ)<>'') then
   begin
    if (dm.CONECTADO) then
    begin
      try
        dm.p_gravaLog('Conferindo rotina de hora em hora',-1);
        oBanco := TDataXml.create;
        oBanco.ServidorWS('www.digifarma.com.br');
        oBanco.SQL := 'select * from ifarma.closeup_lojas where cnpj=' +
        QuotedStr(oConfigProgram.s_CNPJ);
        oBanco.Execute;

        if (oBanco.Eof = false) then
        begin
          s_RETROATIVO := oBanco.FieldByName('retroativo');
          if (s_RETROATIVO = 'S') then
          begin
            envia_arquivo_Closeup;
          end;

        end;

      except
        on ex: Exception do
        begin
          dm.p_gravaLog('Erro '+ex.Message,-1);
          exit;
        end;
      end;

    end;
   end;
 end;

end;

      procedure TFrmPrincipal.JvInicioChange(Sender: TObject);
begin
// JvFim.Date := JvInicio.Date;
end;

procedure TFrmPrincipal.PngSpeedButton1Click(Sender: TObject);
begin
  Ed_ftp.Clear;
  Ed_User.Clear;
  Ed_Senha.Clear;
  Ed_Cod.Clear;
  rb_Auto.Checked:=false;
  rb_Man.Checked:=false;
end;

{ TConfigProgram }

constructor TConfigProgram.create;

Var
  Tabela:TIBquery;
  oBanco,oBancoAux:TDataXml;
  Arq_Split:TStringList;
begin
 s_retroativo:='N';
 Diretorio:=IncludeTrailingBackslash('C:\digifarma') + 'CloseUp';

 {$REGION 'PEGA OS DADOS DA FARMÁCIA NO BANCO LOCAL'}


  try
   dm.bd_ibQuery_init(Tabela);
   Tabela.SQL.Add('select * from config');
   Tabela.Open;

   if (Tabela.Eof=false) then
   begin
     s_CNPJ:=Tabela.FieldByName('cnpj').AsString;
     s_RAZAO_SOCIAL:=Tabela.FieldByName('razao_social').AsString;
     s_FANTASIA:=Tabela.FieldByName('fantasia').AsString;
     s_TELEFONE:=Tabela.FieldByName('telefone').AsString;
   end;
   FreeAndNil(Tabela);
  except
    on ex:exception do
    begin
      dm.p_gravaLog('Erro ao acessar banco de dados'+ex.Message, -1);
      exit;
    end;

  end;
 {$ENDREGION}


 {$REGION 'OLHA SE A FARMÁCIA TEM CLOSEUP'}
  try
   dm.bd_ibQuery_init(Tabela);
   Tabela.SQL.Add('Select count(*)as QTD from close_up');
   Tabela.Open;

   if (Tabela.FieldByName('QTD').AsInteger>0) then
   begin
      FrmPrincipal.cbxFarmacia.Checked:=True;
   end;

   FreeAndNil(Tabela);
  except
    on ex:exception do
    begin
      dm.p_gravaLog('Cliente não tem CloseUp local'+ex.Message, -1);
    end;
  end;
 {$ENDREGION}


 {$REGION 'CONFERE NO BANCO ONLINE SE O CLIENTE É PARCEIRO E JÁ DETERMINA SE O ARQUIVO A SER GERADO É O DE DEMANDA'}
 if (dm.CONECTADO=false) then
 begin
   dm.p_gravaLog('Sem internet aguardando 10 minutos', -1);
   Sleep(600000);
   create;
 end;

 dm.p_gravaLog('Consultando se o cliente é parceiro e a data a ser enviada', -1);

 try
   oBanco:=Tdataxml.Create;
   oBanco.ServidorWS('www.digifarma.com.br');
   oBanco.SQL:= 'SELECT p.*, c.cliente_id, c.parceiro_id, CURRENT_DATE '
   + ' FROM digifarm_digifarma.clientes c left join  '
   + ' digifarm_digifarma.parceiros p on (c.parceiro_id=p.parceiro_id)'
   + ' where c.cnpj=' + QuotedStr(s_CNPJ);
   oBanco.Execute;
   except
    on ex:exception do
    begin
      dm.p_gravaLog('Erro ao acessar banco de dados online'+ex.Message, -1);
      dm.p_gravaLog('Aguardando 5 minutos', -1);
      Sleep(300000);
      create;
    end;

 end;

 while not oBanco.EOF do
 begin
    PARCEIRO_ID:=STRTOINT('0' + oBanco.FieldByName('parceiro_id'));
    s_CLIENTE_ID:=oBanco.FieldByName('cliente_id');

    if ( (oBanco.FieldByName('closeup_host') <> '') and (oBanco.FieldByName('closeup_user') <> '')
    or   (oBanco.FieldByName('closeup_pass') <> '')) or (s_CNPJ='02695980000110') then
    begin

      FTP_USER_REDE  := oBanco.FieldByName('closeup_user');
      FTP_PWD_REDE   := oBanco.FieldByName('closeup_pass');
      FTP_HOST_REDE  := oBanco.FieldByName('closeup_host');
      PARCEIRO       := oBanco.FieldByName('parceiro');

      is_REDE_PARCEIRA := true;

      frmprincipal.rArq_D.Checked := True;
      frmprincipal.cbxRede.Checked := True;

      if (s_CNPJ='02695980000110') then
      begin
        COD_REDE_REDE:='0145';
        FTP_USER_REDE:='closeup';
        FTP_PWD_REDE:='Dig1Clos3up';
        PARCEIRO:='INOVA';
      end;
    end
    else
    begin
      is_REDE_PARCEIRA := false;
      frmprincipal.cbxRede.Checked := false;
    end;

    Arq_Split:=TStringList.Create;
    stl_Split(oBanco.FieldByName('CURRENT_DATE'),'-', Arq_Split);
    Dt_ENVIO:= Arq_Split[2] + '/' + Arq_Split[1] + '/' + Arq_Split[0];

    FreeAndNil(Arq_Split);

    oBanco.Next
 end;
 FreeAndNil(oBanco);
 {$ENDREGION}


 {$REGION 'CARREGA O OBJETO COM OS DADOS DA REDE PARA GERAÇÃO DO ARQUIVO DE DEMANDA'}
  if (is_REDE_PARCEIRA) then
  begin
    frmPrincipal.Tray.Visible    := True;
    Application.ShowMainForm     := False;
    frmPrincipal.rb_Auto.Checked := True;

    try
     dm.p_gravaLog('Carregando arquivo de demanda e atualizando ou inserindo loja online', -1);

       if (dm.CONECTADO=false) then
       begin
         dm.p_gravaLog('Sem internet aguardando 10 minutos', -1);
         dm.p_gravaLog('Aguardando 5 minutos', -1);
         Sleep(300000);
         create;
       end;

      try
        oBanco := TDataXml.Create;
        oBanco.ServidorWS('www.digifarma.com.br');
        oBanco.SQL := 'select * from ifarma.closeup_lojas where cnpj='+QuotedStr(s_CNPJ);
        oBanco.Execute;
        except
        on ex:exception do
        begin
          dm.p_gravaLog('Erro ao acessar banco de dados online'+ex.Message, -1);
          dm.p_gravaLog('Aguardando 5 minutos', -1);
          Sleep(300000);
          create;
        end;
      end;

      if (oBanco.EOF) then
      begin
        try
          oBancoAux:=TDataXml.Create;
          oBancoAux.ServidorWS('www.digifarma.com.br');
          oBancoAux.SQL := 'INSERT INTO ifarma.closeup_lojas (CLIENTE_ID,RAZAO_SOCIAL,FANTASIA,CNPJ,TELEFONE,PARCEIRO_ID) VALUES('+
                          QuotedStr(s_CLIENTE_ID)+', '+
                          QuotedStr(s_RAZAO_SOCIAL)+', '+
                          QuotedStr(s_FANTASIA)+', '+
                          QuotedStr(s_CNPJ)+', '+
                          QuotedStr(leftstr(str_SoNumero(s_TELEFONE),15))+', '+
                          inttostr(PARCEIRO_ID)
                          + ' ) on duplicate key update fantasia=fantasia ';
          oBancoAux.Execute;
          FreeAndNil(oBancoAux);
          except
          on ex:exception do
          begin
            dm.p_gravaLog('Erro ao acessar banco de dados online'+ex.Message, -1);
            dm.p_gravaLog('Aguardando 5 minutos', -1);
            Sleep(300000);
            create;
          end;
        end;
      end
      else
      begin
        s_retroativo:=oBanco.FieldByName('retroativo');
        atualizar_closeup:=oBanco.FieldByName('atualizar_closeup');

        if (atualizar_closeup='S') then
        begin
          if (dm.CONECTADO=false) then
          begin
            dm.p_gravaLog('Sem internet aguardando 10 minutos', -1);
            Sleep(600000);
            create;
          end;

          try
            oBancoAux:=TDataXml.Create;
            oBancoAux.ServidorWS('www.digifarma.com.br');
            oBancoAux.SQL:='update closeup_lojas set atualizar_closeup=''N'' where cnpj=:cnpj';
            oBancoAux.ParamName['cnpj'] := QuotedStr(s_CNPJ);
            oBancoAux.Execute;
            FreeAndNil(oBancoAux);
          except
            on ex:exception do
            begin
              dm.p_gravaLog('Erro ao acessar banco de dados online'+ex.Message, -1);
              dm.p_gravaLog('Aguardando 5 minutos', -1);
              Sleep(300000);
              create;
            end;
          end;
        end;

        if (s_retroativo = 'S') then
        begin
          Arq_Split := TStringList.Create;
          stl_Split(oBanco.FieldByName('data_inicio'), '-', Arq_Split);
          data_inicio := Arq_Split[2] + '/' + Arq_Split[1] + '/' + Arq_Split[0];

          stl_Split(oBanco.FieldByName('data_fim'),'-', Arq_Split);
          data_fim := Arq_Split[2] + '/' + Arq_Split[1] + '/' + Arq_Split[0];
          FreeAndNil(Arq_Split);

          Dt_ENVIO := data_inicio;
        end;
      end;

      FreeAndNil(oBanco);

      try
        oBanco:=TDataXml.Create;
        oBanco.ServidorWS('www.digifarma.com.br');
        oBanco.SQL := 'select CAST(dia_referente as date)dia_referente '
        + ' from ifarma.closeup_arquivos where cnpj=' + QuotedStr(s_CNPJ);
        oBanco.Execute;

        if oBanco.EOF = false then
        begin
          Arq_Split:=TStringList.Create;
          stl_Split(oBanco.FieldByName('dia_referente'),'-',Arq_Split);
          ENVIADO_NOW := Arq_Split[2] + '/' + Arq_Split[1] + '/' + Arq_Split[0];
          FreeAndNil(Arq_Split);
        end;

        except
        on ex:exception do
        begin
          dm.p_gravaLog('Erro ao acessar banco de dados online'+ex.Message, -1);
          dm.p_gravaLog('Aguardando 5 minutos', -1);
          Sleep(300000);
          create;
        end;

      end;

    except
      on ex:exception do
      Begin
        MessageDlg('Erro ao acessar o servidor '+ex.Message,mtError,[mbok],0);
        dm.p_gravaLog('Aguardando 5 minutos', -1);
        Sleep(300000);
        create;
      End;
    end;

    if not DirectoryExists(Diretorio) then
    ForceDirectories(Diretorio);
  end;
  {$ENDREGION}
end;

end.
