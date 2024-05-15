
unit uArqVenda;

interface

uses
  Winapi.Windows, System.SysUtils, System.StrUtils,
  System.Classes, System.inifiles,ibx.ibquery,Vcl.Dialogs;

function gerarArqVenda:String;

implementation

uses
uPrincipal,uDm,uConfig, uwsCloseUp, uFuncoes;


function gerarArqVenda:String;
var
nome_ArqVenda,Tp_Venda,s_hex,codbarra,itemvend_quant:String;
txt:TextFile;
Tabela:Tibquery;

Count,total_unidades,total_unidades_dev,i,x:Integer;
sl_codbarras:TstringList;
begin

  codbarra:='';
  {Escreve no Txt | CLIENTES � Cabe�alho ou Header}
//  Writeln(txt,
//          '4', {Tipo de registro	01	Caracter	Fixo �4� (Default = Padr�o)}
//          '0', {Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o)}
//          oConfigProgram.s_COD_REDE, {Seu c�digo CLOSE UP	04	Caracter	Seu c�digo na CLOSE UP (c�digo fornecido pela CLOSE UP)}
//          FormatDateTime('ddmmyyyy', FrmPrincipal.JvInicio.Date),{Data in�cio	08	Caracter	Data inicial desta informa��o �ddmmaaaa� (dia da venda)}
//          FormatDateTime('ddmmyyyy', FrmPrincipal.JvFim.Date),{Data final	08	Caracter	Data final desta informa��o �ddmmaaaa� (informar a data do �ltimo dia informado no arquivo)}
//          FormatDateTime('ddmmyyyy', Now),{Data arquivo	08	Caracter	Data da gera��o do arquivo �ddmmaaaa� (informar a data do dia da gera��o do arquivo)}
//          'D',{070	FLAG periodicidade	01	caracter	D � Di�rio (Default = Padr�o) }
//          str_preenche_espaco('',3),{Filler	3 Caracter	Espa�os (Default = Padr�o)}
//          'CUPbrven4'{110	Controle interno CLOSE UP	09	Caracter	Fixo �CUPbrven4� (Default = Padr�o)});


  {Escreve no Txt |  2/3 VENDAS � Descri��o ou Detalhamento}
  dm.bd_ibQuery_init(Tabela);
  Tabela.SQL.Add('select i.produto_id,p.cod_barras,i.itemvend_quant,(select cnpj from config),'
  + ' i.cancelado from item_vendas i, '
  + ' produtos p, '
  + ' cab_vendas c '
  + ' where i.venda_nota_id=c.venda_nota_id '
  + ' and   p.produto_id=i.produto_id '
  + ' and   char_length(p.cod_barras) <=13  and  cast(c.venda_data_hora as date)='
  + QuotedStr(FormatDateTime('yyyy-mm-dd',FrmPrincipal.jvDateInc.date)));
//                 +' and   char_length(p.cod_barras) <=13  and  c.venda_data_hora  between '+QuotedStr(DataI) + ' and '+QuotedStr(DataF) );
  Tabela.Open;

  if tabela.Eof then
  nome_ArqVenda:=''
  else
  begin
    sl_codbarras:=TStringList.Create;

    total_unidades:=0;
    total_unidades_dev:=0;

    Count:=2;

    nome_ArqVenda := 'V' + oConfigProgram.s_CNPJ + FormatDateTime('mm', frmprincipal.jvDateInc.Date) + '.D'+FormatDateTime('dd', frmprincipal.jvDateInc.Date);
    AssignFile(txt,IncludeTrailingBackslash(oConfigProgram.Diretorio) + nome_ArqVenda);
    Rewrite(txt);


    while not (Tabela.Eof) do
    begin
      if (xValidaEAN(Tabela.FieldByName('cod_barras').AsString)) then
      codbarra:=codbarra+Tabela.FieldByName('cod_barras').AsString+';';

      Tabela.Next;
    end;

    s_hex:= IntToHex(strtoint(FormatDateTime('ddmmyyyy',now)),2);
    codbarra:=LeftStr(codbarra,(Length(codbarra)-1));

    sl_codbarras.Delimiter:=';';

    {A Fun��o do WS devolve os produtos que n�o est�o no DWH da Inova}
    dm.p_gravaLog('Consultando C�digo de barras', -1);

    if dm.CONECTADO=false then
    begin
      dm.p_gravaLog('Sem internet aguardando 10 minutos', -1);
      Sleep(600000);
      gerarArqVenda;
    end
    else
    begin
      try
       sl_codbarras.DelimitedText := wsConfere.confere_produto(s_hex ,codbarra);
      except
        on ex: EInOutError do
        begin
          dm.p_gravaLog('Erro ao consultar c�digo de barras I/O: '+ex.Message, -1);
          gerarArqVenda
        end;

        on ex: Exception do
        begin
          dm.p_gravaLog('Erro ao consultar c�digo de barras: '+ex.Message, -1);
          gerarArqVenda
        end;
      end;
      dm.p_gravaLog('Consulta do C�digo de barras terminada', -1);
    end;

    Tabela.First;
    while not (Tabela.Eof) do
    begin
      for I := 0 to sl_codbarras.Count-1 do
      begin
        if (tabela.FieldByName('cod_barras').AsString=sl_codbarras[i]) then
        begin
          inc(x);
          break;
        end;
      end;

      if (x=0) then
      begin
        if Tabela.FieldByName('cancelado').AsString='S' then
        begin
          Tp_Venda := 'D';
          inc(total_unidades_dev);
        end
        else
        begin
          Tp_Venda := 'N';
          inc(total_unidades);
        end;

        itemvend_quant:=StringReplace(Tabela.FieldByName('itemvend_quant').AsString,',','',[rfReplaceAll]);


        if (StrToInt64('0'+itemvend_quant)>100000) or (StrToInt64('0'+itemvend_quant)<=0) then
        itemvend_quant:='1';


        Writeln(txt,
                '5',{010	Tipo de Registro	01	caracter	Fixo �5� (Default = Padr�o)}
                FormatDateTime('dd', frmprincipal.jvDateInc.Date),{020	ID Per�odo	02	caracter	Identifica��o do per�odo Di�rio (01 a 31- dia da informa��o)}
                str_preenche_zero(oConfigProgram.s_CNPJ,14),{030	 C�digo cliente	14	Caracter	C�digo do cliente (completar com zeros � esquerda). Esse c�digo � o mesmo que dever� constar no Arquivo de Clientes {Loja Afiliada Da Inova}
                '1', {040	Flag do cliente	01	Caracter 1 � Fixo (Default = Padr�o)}
                '0',{050	Fixo 01	caracter Fixo �0� (zero) (Default = Padr�o)}
                str_preenche_zero(Tabela.FieldByName('cod_barras').AsString, 13),{060	C�digo produto	13	caracter	C�digo do produto (completar com zeros � esquerda). Esse c�digo � o mesmo que dever� constar no arquivo de Produtos}
                '1',{070	Flag produto	01	caracter	1 �Fixo (Default = Padr�o)}
                Tp_Venda,{080	Flag venda	01	caracter	N - Venda normal, Bonifica��o, Brindes e Doa��es N�o informar consigna��o D - Devolu��o}
                str_preenche_zero(itemvend_quant, 8),{090	Quantidade	08	caracter	Quantidade da ocorr�ncia (completar com zeros � esquerda). Unidades vendidas em cada opera��o. }
                'V'); {100	Filler	01	caracter	�V� (Default = Padr�o)}
        inc(count);
      end;

      x:=0;
      Tabela.Next;
    end;

//  { Escreve no TXt | VENDAS � Control ou Trayler}
// Writeln(txt,
//         '6',{010	Tipo de Registro	01	Caracter	Fixo �6� (Default = Padr�o)}
//         '0',{020	Fixo	01	Caracter	Fixo �0� (zero) (Default = Padr�o)}
//         str_preenche_zero(oConfigProgram.s_COD_REDE,4),{030	Seu c�digo CLOSE UP	04	Caracter	Seu c�digo na CLOSE UP (C�digo fornecido pela CLOSE UP)}
//         str_preenche_zero(inttostr(count),8),{040	Total de registros	08	caracter	Total de registros incluindo Header e Trailler (completar com zeros � esquerda)}
//         str_preenche_zero(inttostr(total_unidades),10),{050	Total unidades	10	caracter	Soma do campo quantidades das descri��es do Tipo de Registro Fixo �5� (1� caracter da Descri��o de Vendas � veja p�gina V-2/3). Completar com zeros � esquerda.}
//         str_preenche_zero(inttostr(total_unidades_dev),10),{060	Total de Unidades devolu��es	10	caracter	Soma do campo quantidades das descri��es do Tipo de Registro Fixo �5� com Flag de venda �D�.}
//         'CUPbrven6'{070 Controle interno CLOSE UP 09 caracter Fixo �CUPbrven6� (Default = Padr�o)});
    CloseFile(txt);
  end;
  FreeAndNil(Tabela);
  Result:=nome_ArqVenda;
end;

end.
