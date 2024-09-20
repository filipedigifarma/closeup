unit uArqProd;

interface

uses
  Winapi.Windows, System.SysUtils, System.StrUtils,
  System.Classes, System.inifiles, ibx.ibquery;

function gerarArqProd: String;

implementation

uses
  uPrincipal, uDm, uConfig;

function gerarArqProd: String;
var
  nome_ArqProd, Tipo_Categoria: String;
  txt: TextFile;
  Tabela: Tibquery;
  Count: Integer;

begin
  Count := 2;
  nome_ArqProd := 'P' + oConfigProgram.s_COD_REDE + FormatDateTime('mm',
    frmprincipal.jvDateInc.Date) + '.D' + FormatDateTime('dd',
    frmprincipal.jvDateInc.Date);
  AssignFile(txt, IncludeTrailingBackslash(oConfigProgram.Diretorio) +
    nome_ArqProd);

  Rewrite(txt);

  { Escreve no Txt | CLIENTES � Cabe�alho ou Header }
  Writeln(txt, '7',
    { Tipo de registro	01	Caracter	Fixo �7� (Default = Padr�o) }
    '0', { Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o) }
    oConfigProgram.s_COD_REDE,
    { Seu c�digo CLOSE UP	04	Caracter	Seu c�digo na CLOSE UP (c�digo fornecido pela CLOSE UP) }
    str_preenche_zero(oConfigProgram.s_RAZAO_SOCIAL, 4),
    { Raz�o Social	30	Caracter	Raz�o social da sua empresa }
    str_preenche_zero(oConfigProgram.s_CNPJ, 14),
    { CNPJ	14	Caracter	CNPJ da sua empresa }
    FormatDateTime('ddmmyyyy', frmprincipal.JvInicio.Date),
    { Data in�cio	08	Caracter	Data inicial desta informa��o �ddmmaaaa� (dia da venda) }
    FormatDateTime('ddmmyyyy', frmprincipal.JvFim.Date),
    { Data final	08	Caracter	Data final desta informa��o �ddmmaaaa� (informar a data do �ltimo dia informado no arquivo) }
    FormatDateTime('ddmmyyyy', Now),
    { Data arquivo	08	Caracter	Data da gera��o do arquivo �ddmmaaaa� (informar a data do dia da gera��o do arquivo) }
    '1', { Fixo	01	Caracter	Fixo �1� (N�mero um) (Default = Padr�o) }
    str_preenche_espaco('', 118),
    { Filler	118	Caracter	Espa�os (Default = Padr�o) }
    'CUPbrpro7' { 110	Controle interno CLOSE UP	09	Caracter	Fixo �CUPbrpro7� (Default = Padr�o) } );

  { Escreve no Txt | PRODUTOS � Descri��o ou Detalhamento }
  dm.bd_ibQuery_init(Tabela);
  Tabela.SQL.Add
    ('select p.produto_id,p.produto,p.apresentacao,p.cod_barras,(cast(p.valor_ult_compra as numeric(9,2))*100) as valor_ult_compra,f.*,c.* from produtos p,fornecedores f, categoria c where prod_ativo=''S'' '
    + ' and p.fornecedor_id=f.fornecedor_id ' +
    ' and c.categoria_id=p.categoria_id   ');
  Tabela.Open;
  while not(Tabela.Eof) do
  begin
    if Tabela.FieldByName('tipo').AsString = 'M' then
    begin
      Tipo_Categoria := 'MEDICAMENTOS'
    end
    else if Tabela.FieldByName('tipo').AsString = 'P' then
    begin
      Tipo_Categoria := 'PERFUMARIA'
    end
    ELSE
      Tipo_Categoria := 'OUTROS';

    Writeln(txt, '8',
      { 010	Tipo de Registro	01	Caracter	Fixo �8� (Default = Padr�o) }
      ' ', { 020	Embalagem	01	Caracter	�X� para embalagem fragmentada (somente quando abrir a caixa da embalagem comercial para vender a fra��o, ex ampola, blister). }
      '0', { 030	Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o) }
      str_preenche_zero(Tabela.FieldByName('Produto_id').AsString, 13),
      { 040	C�digo do produto	13	Caracter	C�digo do produto no seu cadastro (completar com zeros � esquerda, cada apresenta��o de produto dever� ter o seu c�digo interno e c�digo separado quando for fracionado).Dever� ser o mesmo informado no arquivo de Vendas
        Caso seja um Distribuidor Hospitalar � obrigat�rio informar um c�digo espec�fico para caixa e outro para unidade ou tudo dever� ser convertido em caixa. }
      '0', { 050	Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o) }
      '0', { 050	Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o) }
      str_preenche_zero(Tabela.FieldByName('cod_barras').AsString, 13),
      { 060	C�digo de barras	13	Caracter	C�digo de barras (obrigat�rio) }
      '1', { 070	Flag	01	Caracter
        1 � C�digo EAN (padr�o embalagem)
        2 � c�digo de barras (outros) }
      str_preenche_espaco(Tabela.FieldByName('produto').AsString + ' ' +
      Tabela.FieldByName('apresentacao').AsString, 70),
      { 080	Nome do produto/ apresenta��o	70	Caracter	Descri��o completa do produto (ex: LIPITOR TABL F.COATE 10mgx 30) }
      str_preenche_espaco('', 8),
      { 090	Filler	8	Caracter	Espa�os (Default = Padr�o) }
      str_preenche_espaco(Tabela.FieldByName('fornecedor').AsString, 40),
      { 100	Fabricante	40	Caracter	Nome do laborat�rio fabricante }
      str_preenche_zero((Tabela.FieldByName('valor_ult_compra').AsString),
      9), { 110	Pre�o f�brica	09	Caracter	Excluir v�rgula. Ex: R$1,20 = 00000120 (preencher com zeros � esquerda) }
      str_preenche_espaco(Tipo_Categoria, 20),
      { 120	Tipo de produto	20	Caracter	Caso tenha classifica��o interna colocar ex; �Perfumaria, Farmac�utica� }
      str_preenche_espaco('', 15),
      { 130	Classifica��o fiscal	15	Caracter	C�digo de classifica��o fiscal / Desconto }
      str_preenche_zero('', 8),
      { 140	Data do cadastro	08	Caracter	Data do cadastro deste produto �ddmmaaaa� (informar a data do ingresso do produto em seu sistema) }
      'P' { 150	Controle Interno CLOSE UP	01	Caracter	Fixo �P� (Default = Padr�o) } );
    Inc(Count);
    Tabela.Next;
  end;
  { Escreve no TXt | PRODUTOS � Control ou Trayler }
  Writeln(txt, '9',
    { 010	Tipo de Registro	01	Caracter	Fixo �9� (Default = Padr�o) }
    '0', { 020	Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o) }
    str_preenche_zero(oConfigProgram.s_COD_REDE, 4),
    { 030	Seu c�digo CLOSE UP	04	Caracter	Seu c�digo na CLOSE UP (C�digo fornecido pela CLOSE UP) }
    str_preenche_zero(inttostr(Count), 8),
    { 040	Total de registros	08	caracter	Total de registros incluindo Header e Trailler (completar com zeros � esquerda) }
    str_preenche_espaco('', 179),
    { 050	Filler	179	caracter	Espa�os (Default = Padr�o) }
    'CUPbrpro9' { 060	Controle interno CLOSE UP 	09	caracter	Fixo �CUPbrpro9� (Default = Padr�o) } );

  CloseFile(txt);
  FreeAndNil(Tabela);
  Result := nome_ArqProd;
end;

end.
