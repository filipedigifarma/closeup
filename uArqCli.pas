
unit uArqCli;

interface

uses
  Winapi.Windows, System.SysUtils, System.StrUtils,
  System.Classes, System.inifiles,ibx.ibquery;

function gerarArqCli:String;

implementation

uses
uPrincipal,uDm,uConfig;


function gerarArqCli:String;
var
nome_ArqCli:String;
txt:TextFile;
Tabela:Tibquery;
Count:Integer;
begin
  Count:=2;
  nome_ArqCli :='C'+ oConfigProgram.s_COD_REDE + FormatDateTime('mm', frmprincipal.jvDateInc.Date) + '.D'+FormatDateTime('dd', frmprincipal.jvDateInc.Date);
  AssignFile(txt,IncludeTrailingBackslash(oConfigProgram.Diretorio) + nome_ArqCli);

  Rewrite(txt);

  {Escreve no Txt | CLIENTES – Cabeçalho ou Header}
  Writeln(txt,
          '1', {Tipo de registro	01	Caracter	Fixo “1” (Default = Padrão)}
          '0', {Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão)}
          oConfigProgram.s_COD_REDE, {Seu código CLOSE UP	04	Caracter	Seu código na CLOSE UP (código fornecido pela CLOSE UP)}
          str_preenche_zero(oConfigProgram.s_RAZAO_SOCIAL,4),{Razão Social	30	Caracter	Razão social da sua empresa}
          str_preenche_zero(oConfigProgram.s_CNPJ,14),{CNPJ	14	Caracter	CNPJ da sua empresa}
          FormatDateTime('ddmmyyyy', FrmPrincipal.JvInicio.Date),{Data início	08	Caracter	Data inicial desta informação “ddmmaaaa” (dia da venda)}
          FormatDateTime('ddmmyyyy', FrmPrincipal.JvFim.Date),{Data final	08	Caracter	Data final desta informação “ddmmaaaa” (informar a data do último dia informado no arquivo)}
          FormatDateTime('ddmmyyyy', Now),{Data arquivo	08	Caracter	Data da geração do arquivo “ddmmaaaa” (informar a data do dia da geração do arquivo)}
          '1',{Fixo	01	Caracter	Fixo “1” (Número um) (Default = Padrão) }
          str_preenche_espaco('',100),{Filler	100	Caracter	Espaços (Default = Padrão)}
          str_preenche_espaco('',171),{Filler	171	Caracter	Espaços (Default = Padrão)}
          'CUPbrcli1'{Controle interno CLOSE UP	09	Caracter	Fixo “CUPbrcli1” (Default = Padrão)});



  {Escreve no TXT | CLIENTES – Descrição ou Detalhamento}
  dm.bd_ibQuery_init(Tabela);
  Tabela.SQL.Add('select * from clientes where cli_inativo=''N'' ');
  Tabela.open;

  while not (Tabela.Eof) do
  begin
  Writeln(txt,
          '2',{Tipo de Registro	01	Caracter	Fixo “2” (Default = Padrão)}
          str_preenche_zero(Tabela.FieldByName('cliente_id').AsString,14),{Código do cliente	14	Caracter	Código interno do cliente (completar com zeros à esquerda)Deverá ser o mesmo informado no arquivo de Vendas}
          str_preenche_zero(Tabela.FieldByName('cli_cpf').AsString,14),{CNPJ/CPF/CRM/RG	14	Caracter	Identificação oficial (completar com zeros à esquerda)}
          '2',{Flag	01	Caracter

                    1 – CNPJ
                    2 – CPF
                    3 - CRM do médico prescritor, identificado na receita do paciente.
                    5 – Quando o médico utilizar o medicamento na Clínica ou consultório, sem identificar o CNPJ.}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cliente').AsString,40),40),{050	Nome fantasia	40	Caracter	Nome fantasia do estabelecimento}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cliente').AsString,40),40),{060	Razão social	40	Caracter	Razão social do estabelecimento}
          '1',{070	 Flag endereço	01	Caracter	Fixo “1” (Default = Padrão)}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_endereco').AsString,70),70),{080	Endereço	70	Caracter	Endereço completo (ex:Avenida Juárez, 255)}
          str_preenche_espaco('',20),{090	Complemento	20	Caracter	Dados complementares do endereço}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_cep').AsString,8),8),{100	CEP	08	Caracter	CEP}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_cidade').AsString,30),30),{110	Cidade	30	Caracter	CIDADE}
          Tabela.FieldByName('cli_uf').AsString,{120	Estado	02	Caracter	UF}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_telefone').AsString,20),20),{130	Telefone	20	Caracter	Telefone}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_telefone').AsString,20),20),{140	Fax	20	Caracter	Fax}
          FormatDateTime('ddmmyyyy',Tabela.FieldByName('cli_data_cadastro').AsDateTime),{150	Data cadastro	08	Caracter	Data do cadastro do cliente “ddmmaaaa” ou alteração. (informar a data do ingresso do seu Cliente em seu sistema)}
          str_preenche_espaco(Tabela.FieldByName('cli_email').AsString,35),{160	e-mail	35	Caracter	Endereço e-mail}
          str_preenche_espaco('',25),{170	URL	25	Caracter	Endereço da web}
          str_preenche_espaco('',5),{180	Filler	05	Caracter	Espaços (Default = Padrão)}
          'C'{190	Controle CLOSE UP 	01	Caracter	Fixo “C” (Default = Padrão)});
  inc(Count);
  Tabela.Next;
  end;
  {Escreve no TXT | 3/3 CLIENTES – Control ou Trayler}
  Writeln(txt,
          '3',{010	Tipo de Registro	01	Caracter	Fixo “3” (Default = Padrão)}
          '0',{Fixo	01	Caracter	Fixo “0” (zero) (Default = Padrão)}
          str_preenche_zero(oConfigProgram.s_COD_REDE,4),{030	Seu código CLOSE UP	04	Caracter	Seu código na CLOSE UP (Código fornecido pela CLOSE UP)}
          str_preenche_zero(inttostr(Count),8),{040	Total de registros	08	Caracter	Total de registros incluindo Header e Trailler (completar com zeros à esquerda)}
          str_preenche_espaco('',200),{050	Filler	200	Caracter	Espaços(Default = Padrão)}
          str_preenche_espaco('',132),{060	Filler	132	Caracter	Espaços(Default = Padrão)}
          'CUPbrcli3'{070	Controle interno CLOSE UP	09	Caracter	Fixo “CUPbrcli3” (Default = Padrão)});



 CloseFile(txt);
 FreeAndNil(Tabela);
 Result:=nome_ArqCli;
end;


end.
