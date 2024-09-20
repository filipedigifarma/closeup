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
  FilePath: String;
begin
  Result := '';
  Tabela := TIBQuery.Create(nil);
  try
    try
      // Inicializa a consulta
      dm.bd_ibQuery_init(Tabela);
      Tabela.SQL.Clear;
      Tabela.SQL.Add(
        'SELECT DISTINCT * FROM (' +
        'SELECT (SELECT cnpj FROM config) AS cnpj, p.cod_barras,' +
        ' v.RECEITUARIO_UF AS UF, v.CONSELHO_NUMERO AS NR_CRM,' +
        ' v.CONSELHO_TIPO AS TP_CRM, I.VENDA_NOTA_ID AS CUPOM, c.venda_data_hora,' +
        ' I.ITEMVEND_QUANT AS QTDE, p.PRODUTO AS PROD, f.fornecedor AS lab,' +
        ' I.vendedor_id AS atend ' +
        'FROM vendas_psicotropicos v ' +
        'LEFT JOIN item_vendas_lotes IL ON (v.VENDA_NOTA_ID = IL.VENDA_NOTA_ID AND v.item_venda_id = IL.item_venda_id) ' +
        'LEFT JOIN item_vendas I ON (v.VENDA_NOTA_ID = I.VENDA_NOTA_ID AND v.item_venda_id = I.item_venda_id) ' +
        'LEFT JOIN produtos p ON (I.produto_id = p.produto_id) ' +
        'LEFT JOIN cab_vendas c ON (I.VENDA_NOTA_ID = c.VENDA_NOTA_ID) ' +
        'LEFT JOIN fornecedores f ON (p.fornecedor_id = f.fornecedor_id) ' +
        'WHERE c.venda_data_hora BETWEEN ' + QuotedStr(DataI) + ' AND ' + QuotedStr(DataF) +
        ' AND v.CONSELHO_NUMERO IS NOT NULL AND v.CONSELHO_TIPO IS NOT NULL ' +
        'UNION ALL ' +
        'SELECT (SELECT cnpj FROM config) AS cnpj, p.cod_barras, a.uf_crm AS UF, a.CRM AS NR_CRM, ' +
        '''CRM'' AS TP_CRM, c.VENDA_NOTA_ID AS CUPOM, c.venda_data_hora, I.ITEMVEND_QUANT AS QTDE,' +
        ' p.PRODUTO AS PROD, f.fornecedor, a.vendedor_id ' +
        'FROM autorizacoes a ' +
        'JOIN cab_vendas c ON (a.nsu = c.nsu) AND c.venda_data_hora BETWEEN ' + QuotedStr(DataI) + ' AND ' + QuotedStr(DataF) +
        ' AND a.status = ''3'' ' +
        'INNER JOIN item_vendas I ON (c.VENDA_NOTA_ID = I.VENDA_NOTA_ID) ' +
        'INNER JOIN produtos p ON (I.produto_id = p.produto_id) ' +
        'INNER JOIN fornecedores f ON (p.fornecedor_id = f.fornecedor_id) ' +
        'UNION ALL ' +
        'SELECT (SELECT cnpj FROM config) AS cnpj, p.cod_barras, i.uf_crmcro AS UF, i.num_crmcro AS NR_CRM,' +
        ' CASE conselho_crmcro WHEN 1 THEN ''CRM'' WHEN 2 THEN ''CRO'' END AS TP_CRM,' +
        ' c.VENDA_NOTA_ID AS CUPOM, c.venda_data_hora, I.ITEMVEND_QUANT AS QTDE,' +
        ' p.PRODUTO AS PROD, f.fornecedor, i.vendedor_id ' +
        'FROM item_vendas I ' +
        'INNER JOIN cab_vendas c ON (I.venda_nota_id = c.venda_nota_id) AND c.venda_data_hora BETWEEN ' + QuotedStr(DataI) + ' AND ' + QuotedStr(DataF) +
        ' INNER JOIN produtos p ON (I.produto_id = p.produto_id) ' +
        'INNER JOIN fornecedores f ON (p.fornecedor_id = f.fornecedor_id) ' +
        'WHERE i.uf_crmcro IS NOT NULL AND i.num_crmcro IS NOT NULL AND conselho_crmcro IS NOT NULL' +
        ') ORDER BY venda_data_hora;'
      );

      dm.p_gravaLog('Query montada, abrindo...', -1);
      Tabela.Open;
      dm.p_gravaLog('Query aberta', -1);

      if Tabela.IsEmpty then
      begin
        dm.p_gravaLog('Nenhum registro encontrado para gerar o arquivo de prescrição.', -1);
        Exit(''); // Retorna string vazia se não houver registros
      end;

      nome_arquivo := oConfigProgram.s_COD_REDE + FormatDateTime('yymmdd', frmprincipal.jvDateInc.Date) + '.txt';
      FilePath := IncludeTrailingBackslash(oConfigProgram.Diretorio) + nome_arquivo;

      AssignFile(txt, FilePath);
      Rewrite(txt);

      try
        while not Tabela.Eof do
        begin
          // Tratamento para cod_crm
          if Tabela.FieldByName('uf').AsString = 'RJ' then
            cod_crm := '052'
          else
            cod_crm := '000';

          // Convertendo o tipo do CRM
          tp_crm := Tabela.FieldByName('tp_crm').AsString;
          if SameText(tp_crm, 'CRM') then
            tp_crm := 'M'
          else if SameText(tp_crm, 'CRMV') then
            tp_crm := 'V'
          else if SameText(tp_crm, 'CRO') then
            tp_crm := 'O'
          else if SameText(tp_crm, 'RMS') then
            tp_crm := 'S'
          else if SameText(tp_crm, 'CRF') then
            tp_crm := 'F'
          else
            tp_crm := ' ';

          // Formatação da data
          dia := FormatDateTime('dd', Tabela.FieldByName('venda_data_hora').AsDateTime);
          mes := FormatDateTime('mm', Tabela.FieldByName('venda_data_hora').AsDateTime);
          ano := FormatDateTime('yyyy', Tabela.FieldByName('venda_data_hora').AsDateTime);

          // Escrevendo no arquivo
          Writeln(txt,
            str_preenche_zero(oConfigProgram.s_COD_REDE, 4) +
            str_preenche_zero('1', 4) +
            str_preenche_zero(Tabela.FieldByName('cnpj').AsString, 14) +
            str_preenche_zero(Tabela.FieldByName('cod_barras').AsString, 20) +
            Tabela.FieldByName('uf').AsString +
            cod_crm +
            str_preenche_zero(Tabela.FieldByName('nr_crm').AsString, 7) +
            tp_crm +
            str_preenche_zero(Tabela.FieldByName('cupom').AsString, 10) +
            str_preenche_zero(dia, 2) +
            str_preenche_zero(mes, 2) +
            str_preenche_zero(ano, 4) +
            str_preenche_zero(Tabela.FieldByName('qtde').AsString, 6) +
            str_preenche_espaco(Tabela.FieldByName('prod').AsString, 50) +
            str_preenche_espaco(Tabela.FieldByName('lab').AsString, 50) +
            'M' +
            str_preenche_zero(Tabela.FieldByName('atend').AsString, 10)
          );

          Tabela.Next;
        end;
      finally
        CloseFile(txt);
      end;

      dm.p_gravaLog('Arquivo de prescrição gerado com sucesso: ' + nome_arquivo, -1);
      Result := nome_arquivo;
    except
      on E: Exception do
      begin
        dm.p_gravaLog('Erro ao gerar o arquivo de prescrição: ' + E.Message, -1);
        // Remove o arquivo se houver erro durante a escrita
        if FileExists(FilePath) then
          DeleteFile(FilePath);
        raise; // Relança a exceção para ser tratada em outro lugar, se necessário
      end;
    end;
  finally
    Tabela.Free;
  end;
end;


function gerarArquivoDemanda:String;
var
  s_ArqCli,s_ArqProd,s_ArqVenda,Nome_Zip: String;
  txt: TextFile;
  Tabela: TIBQuery;
  Zip:TzipFile;
begin
  s_ArqVenda:= gerarArqVenda;
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
  I, n_a, d_f: Integer;
  Arquivo_gerado, UltEnvio: String;
  Arquivo: TINIFile;
  oBanco, oBancoAux: TDataXml;
  b_ftp_pasta_ano: Boolean;

function ConectaFTP(const TipoArquivo: String): Boolean;
begin
  try
    dm.p_gravaLog('Conectando ao FTP para ' + TipoArquivo, -1);

    if IdFTPDigi.Connected then IdFTPDigi.Disconnect;

    if TipoArquivo = 'Prescricao' then
    begin
      IdFTPDigi.Host     := oConfigProgram.g_s_FTP_HOST;
      IdFTPDigi.UserName := oConfigProgram.g_s_FTP_USER;
      IdFTPDigi.Password := oConfigProgram.g_s_FTP_PWD;
    end
    else if TipoArquivo = 'Demanda' then
    begin
      IdFTPDigi.Host     := oConfigProgram.FTP_HOST_REDE;
      IdFTPDigi.UserName := oConfigProgram.FTP_USER_REDE;
      IdFTPDigi.Password := oConfigProgram.FTP_PWD_REDE;
    end;

    IdFTPDigi.Passive := True;
    IdFTPDigi.UseTLS  := utUseExplicitTLS;
    IdFTPDigi.Connect;
    IdFTPDigi.List;
    dm.p_gravaLog('Conectado ao FTP para ' + TipoArquivo, -1);
    Result := True;
  except
    on ex: Exception do
    begin
      dm.p_gravaLog('Erro ao conectar FTP: ' + ex.Message + ' - Aguardando 1 minuto', -1);
      Sleep(60000);
      ConectaFTP(TipoArquivo);
    end;
  end;
end;

function EnviaArquivoFTP(const ArquivoGerado, TipoArquivo: String): Boolean;
var
  mFile: String;
begin
  Result := False;
  try
    dm.p_gravaLog('Enviando arquivo ao FTP: ' + ArquivoGerado, -1);
    mFile := IncludeTrailingBackslash(oConfigProgram.Diretorio) + ArquivoGerado;

    if TipoArquivo = 'Prescricao' then
      IdFTPDigi.Put(mFile, ArquivoGerado)
    else if TipoArquivo = 'Demanda' then
      IdFTPDigi.Put(mFile, ArquivoGerado);

    dm.p_gravaLog('Arquivo enviado com sucesso ao FTP: ' + ArquivoGerado, -1);
    IdFTPDigi.Disconnect;
    dm.p_gravaLog('Desconectado do FTP para ' + TipoArquivo, -1);

    // Atualiza data de último envio no INI
    try
      Arquivo := TINIFile.Create('c:\digifarma\digifarma.ini');
      Arquivo.WriteString('CLOSEUP', 'DT_ULT_ENVIO', jvfim.Text);
      Label4.Caption := jvfim.Text;
    finally
      Arquivo.Free;
    end;

    Result := True;
    App_RegisterDigifarma('Closeup');
  except
    on ex: Exception do
    begin
      dm.p_gravaLog('Erro ao enviar arquivo: ' + ex.Message + ' - Aguardando 1 minuto', -1);
      Sleep(60000);
      EnviaArquivoFTP(ArquivoGerado, TipoArquivo); // Tenta novamente
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

    // Atualiza a data do dia sendo processado
    jvDateInc.Date := IncDay(JvInicio.Date, n_a);
    dm.p_gravaLog('Processando arquivo do dia: ' + jvDateInc.Text, -1);

    // Recupera a data inicial e final para o período sendo processado
    DataI := FormatDateTime('yyyy-mm-dd', jvDateInc.Date) + ' 00:00:00';
    DataF := FormatDateTime('yyyy-mm-dd', jvDateInc.Date) + ' 23:59:59';
    dm.p_gravaLog('Data inicial: ' + DataI + ', Data final: ' + DataF, -1);

    // Atualiza data no servidor, se necessário
    if cbxRede.Checked then
    begin
      if oConfigProgram.s_RETROATIVO = 'S' then
      begin
        if MonthInc <> FormatDateTime('mm', jvDateInc.Date) then
        begin
          if MonthInc <> '' then
          begin
            if dm.CONECTADO then
            begin
              dm.p_gravaLog('Atualizando dia no servidor: ' + jvDateInc.Text, -1);
              try
                oBancoAux := TDataXml.Create;
                oBancoAux.ServidorWS('www.digifarma.com.br');
                oBancoAux.SQL :=
                  'UPDATE ifarma.closeup_lojas SET data_inicio = :data_inicio WHERE cnpj = ' + QuotedStr(oConfigProgram.s_CNPJ);
                oBancoAux.ParamName['data_inicio'] := FormatDateTime('yyyy-mm-dd', jvDateInc.Date);
                oBancoAux.Execute;
              except
                on ex: Exception do
                begin
                  dm.p_gravaLog('Erro ao atualizar dia no servidor: ' + ex.Message, -1);
                  Sleep(300000);  // Aguarda 5 minutos antes de tentar novamente
                  envia_arquivo_Closeup;
                end;
              end;
              FreeAndNil(oBancoAux);
            end;
          end;
          MonthInc := FormatDateTime('mm', jvDateInc.Date);
        end;
      end;
    end;

    // Geração de arquivo de prescrição
    if rArq_P.Checked then
    begin
      dm.p_gravaLog('Gerando arquivo de Prescrição', -1);
      Arquivo_gerado := GeraArquivoPrescricao;

      if (Trim(Arquivo_gerado) <> '') then
      begin
        try
          if (dm.CONECTADO) then
          begin
            dm.p_gravaLog('Acessando FTP', -1);
            GerarScriptFTP('script_ftp.bat', 'ftp.close-upinternational.com.br',
              oConfigProgram.g_s_FTP_USER, oConfigProgram.g_s_FTP_PWD,
              Arquivo_gerado, 'ftptempscript');
            ShellExecute(0, nil, 'cmd.exe',
              PChar('/c ' + 'ftp -s:"script_ftp.bat"'), nil, SW_HIDE);

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
      end;
    end;

    // Geração de arquivo de demanda
    if rArq_D.Checked then
    begin
      dm.p_gravaLog('Gerando arquivo de Demanda', -1);
      Arquivo_gerado := GerarArquivoDemanda;

      if Arquivo_gerado <> '' then
      begin
        dm.p_gravaLog('Arquivo de demanda gerado: ' + Arquivo_gerado, -1);

        // Conecta ao FTP e cria diretórios conforme necessário
        if ConectaFTP('Demanda') then
        begin
          // Criação dos diretórios no FTP
          try
            IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
          except
            on ex: Exception do
            begin
              dm.p_gravaLog('Criando pasta para parceiro: ' + oConfigProgram.PARCEIRO, -1);
              IdFTPDigi.MakeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
              IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO);
            end;
          end;

          // Criação do diretório do ano
          try
            IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO + '/' + FormatDateTime('yyyy', jvDateInc.Date));
          except
            on ex: Exception do
            begin
              dm.p_gravaLog('Criando pasta do ano: ' + FormatDateTime('yyyy', jvDateInc.Date), -1);
              IdFTPDigi.MakeDir(FormatDateTime('yyyy', jvDateInc.Date));
              IdFTPDigi.ChangeDir('/closeup/rede/' + oConfigProgram.PARCEIRO + '/' + FormatDateTime('yyyy', jvDateInc.Date));
            end;
          end;

          // Envia o arquivo de demanda
          EnviaArquivoFTP(Arquivo_gerado, 'Demanda');
        end;
      end;
    end;

    // Finaliza se nenhum arquivo foi selecionado
    if not rArq_P.Checked and not rArq_D.Checked then
      Exit;

  end;

  enviando_arquivo := False;

  {$ENDREGION}



  {$REGION 'SETA O RETROATIVO PRA N NA NUVEM E ATUALIZA O STATUS DO ARQUIVO'}
  if dm.CONECTADO then
  begin
    try
      if oConfigProgram.is_REDE_PARCEIRA then
      begin
        dm.p_gravaLog('Atualizando Informações no Servidor', -1);

        // Verifica e atualiza o status de retroativo
        if oConfigProgram.s_RETROATIVO = 'S' then
        begin
          try
            oBancoAux := TDataXml.Create;
            oBancoAux.ServidorWS('www.digifarma.com.br');
            oBancoAux.SQL :=
              'UPDATE ifarma.closeup_lojas ' +
              'SET retroativo = ''N'', data_inicio = NULL, data_fim = NULL ' +
              'WHERE cnpj = ' + QuotedStr(oConfigProgram.s_CNPJ);
            oBancoAux.Execute;
          except
            on ex: Exception do
            begin
              dm.p_gravaLog('Erro ao acessar servidor para atualização de retroativo: ' + ex.Message, -1);
              Sleep(300000);  // Aguardando 5 minutos antes de tentar novamente
              envia_arquivo_Closeup;  // Tenta novamente
            end;
          end;
          FreeAndNil(oBancoAux);  // Libera o objeto após o uso
        end;

        // Insere ou atualiza informações do arquivo na tabela closeup_arquivos
        try
          oBancoAux := TDataXml.Create;
          oBancoAux.ServidorWS('www.digifarma.com.br');
          oBancoAux.SQL :=
            'INSERT INTO ifarma.closeup_arquivos (data_envio, enviado, dia_referente, cnpj, versao) ' +
            'VALUES(' + QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', Now)) + ', ' +
            QuotedStr('S') + ', ' +
            QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', jvDateInc.Date)) + ', ' +
            QuotedStr(oConfigProgram.s_CNPJ) + ', ' +
            QuotedStr(str_GetFileVer('c:\digifarma\closeup\closeup.exe')) + ')' +
            ' ON DUPLICATE KEY UPDATE ' +
            'data_envio = SYSDATE(), enviado = ''S'', ' +
            'dia_referente = ' + QuotedStr(FormatDateTime('yyyy-mm-dd hh:mm:ss', jvDateInc.Date)) + ', ' +
            'versao = ' + QuotedStr(str_GetFileVer('c:\digifarma\closeup\closeup.exe'));
          oBancoAux.Execute;
        except
          on ex: Exception do
          begin
            dm.p_gravaLog('Erro ao inserir ou atualizar no banco de dados: ' + ex.Message, -1);
            MessageDlg('Erro ao inserir online: ' + ex.Message, mtError, [mbOK], 0);
          end;
        end;
        FreeAndNil(oBancoAux);  // Libera o objeto após o uso
      end;
    except
      on ex: Exception do
      begin
        dm.p_gravaLog('Erro ao acessar servidor: ' + ex.Message, -1);
        Sleep(600000);  // Aguardando 10 minutos antes de tentar novamente
        envia_arquivo_Closeup;  // Tenta novamente
      end;
    end;
  end;
  {$ENDREGION}

  enviando_arquivo := false;

  if (rb_Man.Checked = false) then
  begin
    Exit;
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
  oBanco: TDataXml;
  s_RETROATIVO: String;
begin
  if Assigned(oConfigProgram) then
  begin

    if (enviando_arquivo = false) and (trim(oConfigProgram.s_CNPJ) <> '') then
    begin
      if (dm.CONECTADO) then
      begin
        try
          dm.p_gravaLog('Conferindo rotina de hora em hora', -1);
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
            dm.p_gravaLog('Erro ' + ex.Message, -1);
            Exit;
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

constructor TConfigProgram.Create;

var
  Tabela: TIBQuery;
  oBanco, oBancoAux: TDataXml;
  Arq_Split: TStringList;

  procedure LiberarObjeto(var Obj: TObject);
  begin
    FreeAndNil(Obj);
  end;

  procedure Aguardar(Milisegundos: Integer);
  begin
    dm.p_gravaLog('Aguardando ' + IntToStr(Milisegundos div 60000) + ' minutos', -1);
    Sleep(Milisegundos);
  end;

  function ConectarBancoOnline: Boolean;
  begin
    if not dm.CONECTADO then
    begin
      Aguardar(600000);  // Aguardando 10 minutos
      Result := False;
      Exit;
    end;
    Result := True;
  end;

  procedure ConsultarBancoOnline;
  begin
    oBanco := TDataXml.Create;
    try
      oBanco.ServidorWS('www.digifarma.com.br');
      oBanco.SQL := 'SELECT p.*, c.cliente_id, c.parceiro_id, CURRENT_DATE '
                  + ' FROM digifarm_digifarma.clientes c '
                  + ' LEFT JOIN digifarm_digifarma.parceiros p ON (c.parceiro_id = p.parceiro_id) '
                  + ' WHERE c.cnpj = ' + QuotedStr(s_CNPJ);
      oBanco.Execute;
    except
      on ex: Exception do
      begin
        dm.p_gravaLog('Erro ao acessar banco de dados online: ' + ex.Message, -1);
        Aguardar(300000);  // Aguardando 5 minutos
        Exit;
      end;
    end;
  end;

begin
  s_retroativo := 'N';
  Diretorio := IncludeTrailingBackslash('C:\digifarma') + 'CloseUp';

  {$REGION 'PEGA OS DADOS DA FARMÁCIA NO BANCO LOCAL'}
  try
    dm.bd_ibQuery_init(Tabela);
    Tabela.SQL.Text := 'SELECT * FROM config';
    Tabela.Open;

    if not Tabela.Eof then
    begin
      s_CNPJ := Tabela.FieldByName('cnpj').AsString;
      s_RAZAO_SOCIAL := Tabela.FieldByName('razao_social').AsString;
      s_FANTASIA := Tabela.FieldByName('fantasia').AsString;
      s_TELEFONE := Tabela.FieldByName('telefone').AsString;
    end;
  finally
    LiberarObjeto(TObject(Tabela));
  end;
  {$ENDREGION}

  {$REGION 'OLHA SE A FARMÁCIA TEM CLOSEUP'}
  try
    dm.bd_ibQuery_init(Tabela);
    Tabela.SQL.Text := 'SELECT COUNT(*) AS QTD FROM close_up';
    Tabela.Open;

    FrmPrincipal.cbxFarmacia.Checked := Tabela.FieldByName('QTD').AsInteger > 0;
  finally
    LiberarObjeto(TObject(Tabela));
  end;
  {$ENDREGION}

  {$REGION 'CONFERE NO BANCO ONLINE SE O CLIENTE É PARCEIRO'}
  if not ConectarBancoOnline then
    Exit;

  ConsultarBancoOnline;

  while not oBanco.EOF do
  begin
    PARCEIRO_ID := StrToIntDef(oBanco.FieldByName('parceiro_id'), 0);
    s_CLIENTE_ID := oBanco.FieldByName('cliente_id');

    if (oBanco.FieldByName('closeup_host') <> '') and
       (oBanco.FieldByName('closeup_user') <> '') and
       (oBanco.FieldByName('closeup_pass') <> '') or
       (s_CNPJ = '02695980000110') then
    begin
      FTP_USER_REDE := oBanco.FieldByName('closeup_user');
      FTP_PWD_REDE := oBanco.FieldByName('closeup_pass');
      FTP_HOST_REDE := oBanco.FieldByName('closeup_host');
      PARCEIRO := oBanco.FieldByName('parceiro');
      is_REDE_PARCEIRA := True;

      frmPrincipal.rArq_D.Checked := True;
      frmPrincipal.cbxRede.Checked := True;

      if s_CNPJ = '02695980000110' then
      begin
        COD_REDE_REDE := '0145';
        FTP_USER_REDE := 'closeup';
        FTP_PWD_REDE := 'Dig1Clos3up';
        PARCEIRO := 'INOVA';
      end;
    end
    else
    begin
      is_REDE_PARCEIRA := False;
      frmPrincipal.cbxRede.Checked := False;
    end;

    Arq_Split := TStringList.Create;
    try
      stl_Split(oBanco.FieldByName('CURRENT_DATE'), '-', Arq_Split);
      Dt_ENVIO := Format('%s/%s/%s', [Arq_Split[2], Arq_Split[1], Arq_Split[0]]);
    finally
      LiberarObjeto(TObject(Arq_Split));
    end;

    oBanco.Next;
  end;

  LiberarObjeto(TObject(oBanco));
  {$ENDREGION}

  {$REGION 'CARREGA O OBJETO COM OS DADOS DA REDE'}
  if is_REDE_PARCEIRA then
  begin
    frmPrincipal.Tray.Visible := True;
    Application.ShowMainForm := False;
    frmPrincipal.rb_Auto.Checked := True;

    if not ConectarBancoOnline then
      Exit;

    try
      dm.p_gravaLog('Carregando arquivo de demanda e atualizando ou inserindo loja online', -1);

      oBanco := TDataXml.Create;
      oBanco.ServidorWS('www.digifarma.com.br');
      oBanco.SQL := 'SELECT * FROM ifarma.closeup_lojas WHERE cnpj = ' + QuotedStr(s_CNPJ);
      oBanco.Execute;

      if oBanco.EOF then
      begin
        oBancoAux := TDataXml.Create;
        try
          oBancoAux.ServidorWS('www.digifarma.com.br');
          oBancoAux.SQL := 'INSERT INTO ifarma.closeup_lojas (CLIENTE_ID, RAZAO_SOCIAL, FANTASIA, CNPJ, TELEFONE, PARCEIRO_ID) ' +
                           'VALUES(' + QuotedStr(s_CLIENTE_ID) + ', ' + QuotedStr(s_RAZAO_SOCIAL) + ', ' +
                           QuotedStr(s_FANTASIA) + ', ' + QuotedStr(s_CNPJ) + ', ' +
                           QuotedStr(LeftStr(str_SoNumero(s_TELEFONE), 15)) + ', ' +
                           IntToStr(PARCEIRO_ID) + ') ON DUPLICATE KEY UPDATE fantasia = fantasia';
          oBancoAux.Execute;
        finally
          LiberarObjeto(TObject(oBancoAux));
        end;
      end
      else
      begin
        s_retroativo := oBanco.FieldByName('retroativo');
        atualizar_closeup := oBanco.FieldByName('atualizar_closeup');

        if atualizar_closeup = 'S' then
        begin
          if not ConectarBancoOnline then
            Exit;

          oBancoAux := TDataXml.Create;
          try
            oBancoAux.ServidorWS('www.digifarma.com.br');
            oBancoAux.SQL := 'UPDATE closeup_lojas SET atualizar_closeup = ''N'' WHERE cnpj = :cnpj';
            oBancoAux.ParamName['cnpj'] := s_CNPJ;
            oBancoAux.Execute;
          finally
            LiberarObjeto(TObject(oBancoAux));
          end;
        end;

        if s_retroativo = 'S' then
        begin
          Arq_Split := TStringList.Create;
          try
            stl_Split(oBanco.FieldByName('data_inicio'), '-', Arq_Split);
            data_inicio := Format('%s/%s/%s', [Arq_Split[2], Arq_Split[1], Arq_Split[0]]);

            stl_Split(oBanco.FieldByName('data_fim'), '-', Arq_Split);
            data_fim := Format('%s/%s/%s', [Arq_Split[2], Arq_Split[1], Arq_Split[0]]);
            Dt_ENVIO := data_inicio;
          finally
            LiberarObjeto(TObject(Arq_Split));
          end;
        end;
      end;

      LiberarObjeto(TObject(oBanco));

      oBanco := TDataXml.Create;
      oBanco.ServidorWS('www.digifarma.com.br');
      oBanco.SQL := 'SELECT CAST(dia_referente AS DATE) AS dia_referente ' +
                    'FROM ifarma.closeup_arquivos WHERE cnpj = ' + QuotedStr(s_CNPJ);
      oBanco.Execute;

      if not oBanco.EOF then
      begin
        Arq_Split := TStringList.Create;
        try
          stl_Split(oBanco.FieldByName('dia_referente'), '-', Arq_Split);
          ENVIADO_NOW := Format('%s/%s/%s', [Arq_Split[2], Arq_Split[1], Arq_Split[0]]);
        finally
          LiberarObjeto(TObject(Arq_Split));
        end;
      end;
    except
      on ex: Exception do
      begin
        dm.p_gravaLog('Erro ao acessar o servidor: ' + ex.Message, -1);
        Aguardar(300000);  // Aguardando 5 minutos
        Exit;
      end;
    end;

    if not DirectoryExists(Diretorio) then ForceDirectories(Diretorio);
  end;
  {$ENDREGION}
end;

end.
