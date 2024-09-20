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

  { Escreve no Txt | CLIENTES – Cabeçalho ou Header }
  Writeln(txt, '7',
    { Tipo de registro	01	Caracter	Fixo “7” (Default = Padrão) }
    '0', { Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão) }
    oConfigProgram.s_COD_REDE,
    { Seu código CLOSE UP	04	Caracter	Seu código na CLOSE UP (código fornecido pela CLOSE UP) }
    str_preenche_zero(oConfigProgram.s_RAZAO_SOCIAL, 4),
    { Razão Social	30	Caracter	Razão social da sua empresa }
    str_preenche_zero(oConfigProgram.s_CNPJ, 14),
    { CNPJ	14	Caracter	CNPJ da sua empresa }
    FormatDateTime('ddmmyyyy', frmprincipal.JvInicio.Date),
    { Data início	08	Caracter	Data inicial desta informação “ddmmaaaa” (dia da venda) }
    FormatDateTime('ddmmyyyy', frmprincipal.JvFim.Date),
    { Data final	08	Caracter	Data final desta informação “ddmmaaaa” (informar a data do último dia informado no arquivo) }
    FormatDateTime('ddmmyyyy', Now),
    { Data arquivo	08	Caracter	Data da geração do arquivo “ddmmaaaa” (informar a data do dia da geração do arquivo) }
    '1', { Fixo	01	Caracter	Fixo “1” (Número um) (Default = Padrão) }
    str_preenche_espaco('', 118),
    { Filler	118	Caracter	Espaços (Default = Padrão) }
    'CUPbrpro7' { 110	Controle interno CLOSE UP	09	Caracter	Fixo “CUPbrpro7” (Default = Padrão) } );

  { Escreve no Txt | PRODUTOS – Descrição ou Detalhamento }
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
      { 010	Tipo de Registro	01	Caracter	Fixo “8” (Default = Padrão) }
      ' ', { 020	Embalagem	01	Caracter	“X” para embalagem fragmentada (somente quando abrir a caixa da embalagem comercial para vender a fração, ex ampola, blister). }
      '0', { 030	Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão) }
      str_preenche_zero(Tabela.FieldByName('Produto_id').AsString, 13),
      { 040	Código do produto	13	Caracter	Código do produto no seu cadastro (completar com zeros à esquerda, cada apresentação de produto deverá ter o seu código interno e código separado quando for fracionado).Deverá ser o mesmo informado no arquivo de Vendas
        Caso seja um Distribuidor Hospitalar é obrigatório informar um código específico para caixa e outro para unidade ou tudo deverá ser convertido em caixa. }
      '0', { 050	Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão) }
      '0', { 050	Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão) }
      str_preenche_zero(Tabela.FieldByName('cod_barras').AsString, 13),
      { 060	Código de barras	13	Caracter	Código de barras (obrigatório) }
      '1', { 070	Flag	01	Caracter
        1 – Código EAN (padrão embalagem)
        2 – código de barras (outros) }
      str_preenche_espaco(Tabela.FieldByName('produto').AsString + ' ' +
      Tabela.FieldByName('apresentacao').AsString, 70),
      { 080	Nome do produto/ apresentação	70	Caracter	Descrição completa do produto (ex: LIPITOR TABL F.COATE 10mgx 30) }
      str_preenche_espaco('', 8),
      { 090	Filler	8	Caracter	Espaços (Default = Padrão) }
      str_preenche_espaco(Tabela.FieldByName('fornecedor').AsString, 40),
      { 100	Fabricante	40	Caracter	Nome do laboratório fabricante }
      str_preenche_zero((Tabela.FieldByName('valor_ult_compra').AsString),
      9), { 110	Preço fábrica	09	Caracter	Excluir vírgula. Ex: R$1,20 = 00000120 (preencher com zeros à esquerda) }
      str_preenche_espaco(Tipo_Categoria, 20),
      { 120	Tipo de produto	20	Caracter	Caso tenha classificação interna colocar ex; “Perfumaria, Farmacêutica” }
      str_preenche_espaco('', 15),
      { 130	Classificação fiscal	15	Caracter	Código de classificação fiscal / Desconto }
      str_preenche_zero('', 8),
      { 140	Data do cadastro	08	Caracter	Data do cadastro deste produto “ddmmaaaa” (informar a data do ingresso do produto em seu sistema) }
      'P' { 150	Controle Interno CLOSE UP	01	Caracter	Fixo “P” (Default = Padrão) } );
    Inc(Count);
    Tabela.Next;
  end;
  { Escreve no TXt | PRODUTOS – Control ou Trayler }
  Writeln(txt, '9',
    { 010	Tipo de Registro	01	Caracter	Fixo “9” (Default = Padrão) }
    '0', { 020	Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão) }
    str_preenche_zero(oConfigProgram.s_COD_REDE, 4),
    { 030	Seu código CLOSE UP	04	Caracter	Seu código na CLOSE UP (Código fornecido pela CLOSE UP) }
    str_preenche_zero(inttostr(Count), 8),
    { 040	Total de registros	08	caracter	Total de registros incluindo Header e Trailler (completar com zeros à esquerda) }
    str_preenche_espaco('', 179),
    { 050	Filler	179	caracter	Espaços (Default = Padrão) }
    'CUPbrpro9' { 060	Controle interno CLOSE UP 	09	caracter	Fixo “CUPbrpro9” (Default = Padrão) } );

  CloseFile(txt);
  FreeAndNil(Tabela);
  Result := nome_ArqProd;
end;

end.
