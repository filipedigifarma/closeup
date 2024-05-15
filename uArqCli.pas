
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

  {Escreve no Txt | CLIENTES � Cabe�alho ou Header}
  Writeln(txt,
          '1', {Tipo de registro	01	Caracter	Fixo �1� (Default = Padr�o)}
          '0', {Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o)}
          oConfigProgram.s_COD_REDE, {Seu c�digo CLOSE UP	04	Caracter	Seu c�digo na CLOSE UP (c�digo fornecido pela CLOSE UP)}
          str_preenche_zero(oConfigProgram.s_RAZAO_SOCIAL,4),{Raz�o Social	30	Caracter	Raz�o social da sua empresa}
          str_preenche_zero(oConfigProgram.s_CNPJ,14),{CNPJ	14	Caracter	CNPJ da sua empresa}
          FormatDateTime('ddmmyyyy', FrmPrincipal.JvInicio.Date),{Data in�cio	08	Caracter	Data inicial desta informa��o �ddmmaaaa� (dia da venda)}
          FormatDateTime('ddmmyyyy', FrmPrincipal.JvFim.Date),{Data final	08	Caracter	Data final desta informa��o �ddmmaaaa� (informar a data do �ltimo dia informado no arquivo)}
          FormatDateTime('ddmmyyyy', Now),{Data arquivo	08	Caracter	Data da gera��o do arquivo �ddmmaaaa� (informar a data do dia da gera��o do arquivo)}
          '1',{Fixo	01	Caracter	Fixo �1� (N�mero um) (Default = Padr�o) }
          str_preenche_espaco('',100),{Filler	100	Caracter	Espa�os (Default = Padr�o)}
          str_preenche_espaco('',171),{Filler	171	Caracter	Espa�os (Default = Padr�o)}
          'CUPbrcli1'{Controle interno CLOSE UP	09	Caracter	Fixo �CUPbrcli1� (Default = Padr�o)});



  {Escreve no TXT | CLIENTES � Descri��o ou Detalhamento}
  dm.bd_ibQuery_init(Tabela);
  Tabela.SQL.Add('select * from clientes where cli_inativo=''N'' ');
  Tabela.open;

  while not (Tabela.Eof) do
  begin
  Writeln(txt,
          '2',{Tipo de Registro	01	Caracter	Fixo �2� (Default = Padr�o)}
          str_preenche_zero(Tabela.FieldByName('cliente_id').AsString,14),{C�digo do cliente	14	Caracter	C�digo interno do cliente (completar com zeros � esquerda)Dever� ser o mesmo informado no arquivo de Vendas}
          str_preenche_zero(Tabela.FieldByName('cli_cpf').AsString,14),{CNPJ/CPF/CRM/RG	14	Caracter	Identifica��o oficial (completar com zeros � esquerda)}
          '2',{Flag	01	Caracter

                    1 � CNPJ
                    2 � CPF
                    3 - CRM do m�dico prescritor, identificado na receita do paciente.
                    5 � Quando o m�dico utilizar o medicamento na Cl�nica ou consult�rio, sem identificar o CNPJ.}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cliente').AsString,40),40),{050	Nome fantasia	40	Caracter	Nome fantasia do estabelecimento}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cliente').AsString,40),40),{060	Raz�o social	40	Caracter	Raz�o social do estabelecimento}
          '1',{070	 Flag endere�o	01	Caracter	Fixo �1� (Default = Padr�o)}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_endereco').AsString,70),70),{080	Endere�o	70	Caracter	Endere�o completo (ex:Avenida Ju�rez, 255)}
          str_preenche_espaco('',20),{090	Complemento	20	Caracter	Dados complementares do endere�o}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_cep').AsString,8),8),{100	CEP	08	Caracter	CEP}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_cidade').AsString,30),30),{110	Cidade	30	Caracter	CIDADE}
          Tabela.FieldByName('cli_uf').AsString,{120	Estado	02	Caracter	UF}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_telefone').AsString,20),20),{130	Telefone	20	Caracter	Telefone}
          str_preenche_espaco(leftstr(Tabela.FieldByName('cli_telefone').AsString,20),20),{140	Fax	20	Caracter	Fax}
          FormatDateTime('ddmmyyyy',Tabela.FieldByName('cli_data_cadastro').AsDateTime),{150	Data cadastro	08	Caracter	Data do cadastro do cliente �ddmmaaaa� ou altera��o. (informar a data do ingresso do seu Cliente em seu sistema)}
          str_preenche_espaco(Tabela.FieldByName('cli_email').AsString,35),{160	e-mail	35	Caracter	Endere�o e-mail}
          str_preenche_espaco('',25),{170	URL	25	Caracter	Endere�o da web}
          str_preenche_espaco('',5),{180	Filler	05	Caracter	Espa�os (Default = Padr�o)}
          'C'{190	Controle CLOSE UP 	01	Caracter	Fixo �C� (Default = Padr�o)});
  inc(Count);
  Tabela.Next;
  end;
  {Escreve no TXT | 3/3 CLIENTES � Control ou Trayler}
  Writeln(txt,
          '3',{010	Tipo de Registro	01	Caracter	Fixo �3� (Default = Padr�o)}
          '0',{Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o)}
          str_preenche_zero(oConfigProgram.s_COD_REDE,4),{030	Seu c�digo CLOSE UP	04	Caracter	Seu c�digo na CLOSE UP (C�digo fornecido pela CLOSE UP)}
          str_preenche_zero(inttostr(Count),8),{040	Total de registros	08	Caracter	Total de registros incluindo Header e Trailler (completar com zeros � esquerda)}
          str_preenche_espaco('',200),{050	Filler	200	Caracter	Espa�os(Default = Padr�o)}
          str_preenche_espaco('',132),{060	Filler	132	Caracter	Espa�os(Default = Padr�o)}
          'CUPbrcli3'{070	Controle interno CLOSE UP	09	Caracter	Fixo �CUPbrcli3� (Default = Padr�o)});



 CloseFile(txt);
 FreeAndNil(Tabela);
 Result:=nome_ArqCli;
end;


end.
