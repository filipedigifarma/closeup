unit uMySql;

interface

uses  System.SysUtils, System.Classes, Vcl.Dialogs

      , {$IFDEF VER280}ibx.ibquery{$ELSE}ibx.ibquery{$ENDIF}

      , Xml.XMLDoc, Xml.XMLIntf, XMLDOM

      , EncdDecd, ZLib,

      Datasnap.DBClient, Data.DB,
      strutils;

      Function ComparaTamanho(List: TStringList; Index1, Index2: Integer): Integer;
      Function DecBase64(Value: String): String;
      Function InStr( Inicio : Integer; lsString, Achar : String ) : Integer;
      Function AnsiInStr( Inicio : Integer; lsString, Achar : String ) : Integer;
      Function Mid( What : String; Index, Count : Integer ) : String;


type

  TDataXML = class
  private
    FScript_Count: Integer;
    FScript_Index: Integer;
    FDataset: Integer;
    FParametro: TStringList;
    FScript: Boolean;
    FSQL: String;
    FREDE: String;
    FRecordCount: Integer;
    FParamCount: Integer;
    FPosition: Integer;
    FErro: Integer;
    FLastID: Int64;
    FError: String;
    FXML:String;
    FDatetime:TDateTime;
    FDoc: IXMLDocument;
    FNotFound, FShowErrors: Boolean;
    FSalvarXML: Boolean;
    FPathXML: String;

    function GetRecordCount: Integer;
    function GetDatetime: TDateTime;
    function ExecSQL(Tipo: String): Boolean;

    procedure CriaCampoCDS(Var AuxCDS: TClientDataSet; CampoNomeBD, CampoTipoBD: String; CampoTamBD: Integer =0; CampoDecBD: Integer=0; CampoEscalaBD: Integer=0);
    procedure GravaCamposScript(const pindex: integer);
    procedure GravaParam(NomeParametro: String; Value: String);
    function LeIndex(Index: Integer): String;
    function LeParam(NomeParametro: String): String;


  public
    constructor Create;virtual;
    destructor destroy;override;

    function ClienteDataSet(Var Aux: TClientDataSet; SQL_QUERY, NOME_TABELA, NOME_SERVER: String; Rede: Integer; Compress: Boolean = False): Boolean;
    function Execute: Boolean;
    function ExecuteZ: Boolean;
    function FieldsCount: Integer;
    function FieldByIndex(Indice: Integer): String;
    function FieldByName(Field: String): String;
    function FieldExist(Field: String): Boolean;
    function FieldSize(Field: String): Int64;
    function FieldType(Field: String): String;
    function FieldName(Indice: Integer): String;
    function EOF: Boolean;
    function IsOnLine: Boolean;
    function ServidorWS(Nome: String): Boolean;
    property SYSDate: TDateTime read GetDatetime;
    procedure Close;
    procedure First;
    procedure Next;
    procedure Previous;
    procedure Last;

  //published

    property Erro: Integer read FErro;
    property Error: String read FError;
    property LastID: Int64 read FLastID;
    property NotFound: Boolean read FNotFound write FNotFound;
    Property ParamCount : Integer read FParamCount;
    Property ParamName[NomeParametro: String] : String read LeParam write GravaParam;
    Property ParamIndex[Index: Integer] : String read LeIndex;
    property Position: Integer read FPosition;
    property RecordCount: Integer read GetRecordCount;
    property RedeID: String read FREDE write FREDE;
    property SalvarXML: Boolean read FSalvarXML write FSalvarXML;
    property PathXML: String read FPathXML write FPathXML;
    property Script: Boolean read FScript write FScript;
    property Script_Count: Integer read FScript_Count;
    property Script_Index: Integer read FScript_Index write GravaCamposScript;
    property SQL: String read FSQL write FSQL;
    property ShowErrors: Boolean read FShowErrors write FShowErrors;

  end;

implementation

uses uwsServer;

procedure TDataXML.GravaCamposScript(const pindex: integer);
  var DOMNodeList: IDOMNodeList;
      DOMNodeSelect: IDOMNodeSelect;
      Campo: IXMLNode;
Begin

  if (pIndex <= FScript_Count) then
  begin
    FScript_Index := pindex;

    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[0];
    FRecordCount := StrToInt( Campo.AttributeNodes['rows'].Text );

  end;



End;


constructor TDataXML.Create;
begin
  inherited;
  try
    FSalvarXML := False;
    FPathXML := '';
    FScript_Count := 0;
    FScript_Index := 0;
    FParametro := TStringList.Create;
    FLastID :=0;
    FScript :=False;
    FSQL :='';
    FREDE :='';
    FXML := '';
    if (FDoc <> nil) then
      FDoc := nil;
    FRecordCount := -1;
    FErro := 0;
    FError := '';
    FPosition := 0;
    FShowErrors := True;
    FNotFound :=False;
  except

  end;
end;

destructor TDataXML.Destroy;
begin
  inherited;
  Close;
end;

  Function ComparaTamanho(List: TStringList; Index1, Index2: Integer): Integer;
    var
      tam1, tam2: integer;
  begin
    tam1 := Length(List[Index1]);
    tam2 := Length(List[Index2]);
    if tam1 < tam2 then
      Result := -1
    else if tam1 > tam2 then Result := 1
    else
      Result := 0;

  End;

  Function DecBase64(Value: String): String;
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


  Function Mid( What : String; Index, Count : Integer ) : String;
  Begin
      Result := Copy( What, Index, Count );
  End;

  Function InStr( Inicio : Integer; lsString, Achar : String ) : Integer;
  Var
      tmpStr : String;
      lnI: Integer;
  Begin
      tmpStr := Copy( lsString, Inicio, MaxInt );
      lnI := Pos( Achar, tmpStr );
      if lnI = 0 then
        lnI := 0
      else
        lnI := Inicio + lnI;
      Result := lnI;
  End;


  Function AnsiInStr( Inicio : Integer; lsString, Achar : String ) : Integer;
  Var
      tmpStr : String;
      lnI: Integer;
  Begin
      tmpStr := Copy( lsString, Inicio, MaxInt );
      lnI := AnsiPos( Achar, tmpStr );
      if lnI = 0 then
        lnI := 0
      else
        lnI := Inicio + lnI;
      Result := lnI;
  End;

procedure TDataXML.CriaCampoCDS(Var AuxCDS: TClientDataSet; CampoNomeBD, CampoTipoBD: String; CampoTamBD: Integer =0; CampoDecBD: Integer=0; CampoEscalaBD: Integer=0);
  var TipoCampo: TFieldType;
      Campo: TFieldDef;
begin


        if (CampoTipoBD = '1') then TipoCampo := TFieldType.ftInteger;
        if (CampoTipoBD = '2') then TipoCampo := TFieldType.ftInteger;
        if (CampoTipoBD = '3') then TipoCampo := TFieldType.ftInteger;
        if (CampoTipoBD = '4') then TipoCampo := TFieldType.ftBCD;
        if (CampoTipoBD = '5') then TipoCampo := TFieldType.ftBCD;
        if (CampoTipoBD = '7') then TipoCampo := TFieldType.ftTimeStamp;
        if (CampoTipoBD = '8') then TipoCampo := TFieldType.ftLargeint;
        if (CampoTipoBD = '9') then TipoCampo := TFieldType.ftLargeint;
        if (CampoTipoBD = '10') then TipoCampo := TFieldType.ftDate;
        if (CampoTipoBD = '11') then TipoCampo := TFieldType.ftTime;
        if (CampoTipoBD = '12') then TipoCampo := TFieldType.ftDateTime;
        if (CampoTipoBD = '252') then TipoCampo := TFieldType.ftBlob;
        if (CampoTipoBD = '253') then TipoCampo := TFieldType.ftString;
        if (CampoTipoBD = '254') then TipoCampo := TFieldType.ftString;
        if (CampoTipoBD = '246') then TipoCampo := TFieldType.ftBCD;

        Campo := AuxCDS.FieldDefs.AddFieldDef;
        Campo.Name := CampoNomeBD;
        Campo.DataType := TipoCampo;

        if (TipoCampo = ftString)  then
        Begin
          if (CampoTamBD >0) then
          Begin
            Campo.Size := CampoTamBD;
          End
          else
          Begin
            Campo.Size := 1;
          End;
        End;

        if (TipoCampo = ftBCD)  then
        Begin
          if (CampoDecBD >0) then
          Begin
            Campo.Size := CampoTamBD;
          End;
        End;

end;

function TDataXML.ClienteDataSet(Var Aux: TClientDataSet; SQL_QUERY, NOME_TABELA, NOME_SERVER: String; Rede: Integer; Compress: Boolean = False): Boolean;
  var QueryTB: TDataXML;
//      TipoCampo: TFieldType;
      Precisao, Escala, I: Integer;
      Tamanho: int64;
      NomeCampo, Tipo: String;
begin

  Result := False;

  QueryTB := TDataXML.Create;
  QueryTB.SQL :=  SQL_QUERY;

  if (Rede <> 0) then
  Begin
    QueryTB.RedeID := IntToStr(Rede);
  End
  else
  Begin
    QueryTB.ServidorWS(NOME_SERVER)
  End;


  if (Compress = False) then
    QueryTB.Execute
  else
    QueryTB.ExecuteZ;

(*  CamposTB := TDataXML.Create;
  CamposTB.SQL :=  'SELECT  COLUMN_NAME '+
                '       , DATA_TYPE '+
                '       , CHARACTER_MAXIMUM_LENGTH '+
                '       , NUMERIC_PRECISION '+
                '       , NUMERIC_SCALE '+
                '       , EXTRA '+
                '       , COLUMN_KEY '+
                '       , COLUMN_DEFAULT '+
                '       , IS_NULLABLE '+
                ' FROM    INFORMATION_SCHEMA.COLUMNS '+
                ' WHERE   table_name = ' + QuotedStr(NOME_TABELA) + ' '+
                '         AND TABLE_SCHEMA = ' + QuotedStr(NOME_BANCO) + ' ';
  CamposTB.Execute;
*)

  for I := 0 to (QueryTB.FieldsCount-1) do
  Begin
    NomeCampo := QueryTB.FieldName(I);
    Tipo := QueryTB.FieldType(NomeCampo);
    Tamanho := QueryTB.FieldSize(NomeCampo);
    Precisao := 0;
    Escala := 0;
    CriaCampoCDS(Aux, NomeCampo, Tipo, Tamanho,Precisao, Escala);
  End;

//  CamposTB.Close;
//  FreeAndNil(CamposTB);

  Aux.CreateDataSet;
  Aux.Active := True;
  Aux.LogChanges := False;
  Aux.EmptyDataSet;

  while (QueryTB.EOF = False) do
  Begin

    Aux.Append;
    for I := 0 to (QueryTB.FieldsCount-1)  do
    Begin
      NomeCampo := QueryTB.FieldName(I);
      if (QueryTB.FieldByIndex(I) <> '') then
        if (Aux.FieldByName(QueryTB.FieldName(I)).DataType = TFieldType.ftBCD) then
        begin
          Aux.FieldByName(QueryTB.FieldName(I)).AsBCD :=  StrToCurr( StringReplace(QueryTB.FieldByIndex(I),'.',',',[rfReplaceAll]));
        end
        else
        if (Aux.FieldByName(QueryTB.FieldName(I)).DataType = TFieldType.ftDate) then
        Begin
          if ( (QueryTB.FieldByIndex(I) <> '0000-00-00') and (QueryTB.FieldByIndex(I) <> '') )  then
            Aux.FieldByName(QueryTB.FieldName(I)).Value :=  QueryTB.FieldByIndex(I);
        End
        else
        if (Aux.FieldByName(QueryTB.FieldName(I)).DataType = TFieldType.ftDateTime) then
        Begin
          if ( (QueryTB.FieldByIndex(I) <> '0000-00-00 00:00:00') and (QueryTB.FieldByIndex(I) <> '') )  then
            Aux.FieldByName(QueryTB.FieldName(I)).Value :=  QueryTB.FieldByIndex(I);
        End
        else
        if (Aux.FieldByName(QueryTB.FieldName(I)).DataType = TFieldType.ftBlob) then
        Begin
          if ( QueryTB.FieldByIndex(I) <> '')  then
          Begin
            Aux.FieldByName(QueryTB.FieldName(I)).Value :=  DecodeBase64(QueryTB.FieldByIndex(I));
          End;
        End
        else
        Begin
          Aux.FieldByName(QueryTB.FieldName(I)).Value :=  QueryTB.FieldByIndex(I);
        End;

    End;
    Aux.Post;
    QueryTB.Next
  End;

  QueryTB.Close;
  FreeAndNil(QueryTB);


  Aux.ApplyUpdates(-1);
  Aux.First;
  // Aux.SaveToFile('c:\siga.xml');
  Result := True;

end;

procedure TDataXML.GravaParam(NomeParametro: String; Value: String);
Begin

  if ((Value = EmptyStr) or (Value = '')) then
  Begin
    FParametro.Values[NomeParametro] := ' ';
  End
  else
  Begin
    FParametro.Values[NomeParametro] := Value;
  End;
  FParamCount := FParametro.Count;
  FParametro.CustomSort(ComparaTamanho);
End;



function TDataXML.LeIndex(Index: Integer): String;
Begin
  Result := FParametro.Names[Index];
End;

function TDataXML.LeParam(NomeParametro: String): String;
Begin
  Result := FParametro.Values[NomeParametro];
End;

function TDataXML.Execute: Boolean;
begin
  Result := ExecSQL('');
end;

function TDataXML.ExecuteZ: Boolean;
Begin
  Result := ExecSQL('zip');
end;

function TDataXML.ExecSQL(Tipo: String): Boolean;
  var l_data_hex, Erro, Desc, Qtd, opcoes : String;
      I, P1, P2: Integer;
      servico: server_ValidadorPortType;
//      Lista: TStringList;

      DOMNodeList: IDOMNodeList;
      DOMNodeSelect: IDOMNodeSelect;

      Retorno: string;
      tamanho: Integer;
      Buff: TBytes;
      s_arq: TBytesStream;
      st: TStream;
      Encoding: TEncoding;

      strInput, strOutput: TStringStream;
      Unzipper: TZDecompressionStream;
begin

  Result := False;
  FLastID :=0;

  try

    if (uwsServer.defURL = '') then
    Begin
      uwsServer.defWSDL := 'http://www.digifarma.com.br/webservice/selectxml.php?wsdl';
      uwsServer.defURL  := 'http://www.digifarma.com.br/webservice/selectxml.php';
    End;

    if (FParametro.Count >0) then
    Begin
      for I := 0 to (FParametro.Count-1) do
      Begin
        FSQL := ReplaceText(FSQL,':'+FParametro.Names[I], QuotedStr(Trim(FParametro.Values[FParametro.Names[I]])))
      End;
    End;

    servico := wsfServer();

    opcoes := '1:1';
    if (FREDE <> '') then
      opcoes := opcoes + ';rede_id=' + FREDE;

    if (FScript = False) then
      opcoes := opcoes + ';script=N'
    else
      opcoes := opcoes + ';script=S';

    if (Tipo = 'zip') then
      opcoes := opcoes + ';zip=S';

    l_data_hex := IntToHex(StrToInt(FormatDateTime('ddmmyyyy', Now)),2);

    FXML := servico.selectxmlf(l_data_hex, FSQL, opcoes);

{----------------------------  ZIP  -------------------------------------------------}
    if (Tipo = 'zip') then
    Begin
      Retorno :=   DecBase64(FXML);

      strInput:= TStringStream.Create(Retorno);
      strOutput:= TStringStream.Create;

      try
        Unzipper:= TZDecompressionStream.Create(strInput);
        try
          strOutput.CopyFrom(Unzipper, Unzipper.Size);
        finally
          Unzipper.Free;
        end;
        FXML:= strOutput.DataString;
      finally
        strInput.Free;
        strOutput.Free;
      end;
    End;
{----------------------------  ZIP  -------------------------------------------------}

    servico := nil;

//    Lista := TStringList.Create;
//    Lista.Add(FXML);
//    lista.SaveToFile('c:\select.txt');


    If (Pos('<codigo>000</codigo>',LowerCase(FXML)) > 0) then
    Begin
      Erro := '000';
    End
    else if (Pos('<codigo>003</codigo>',LowerCase(FXML)) > 0) then
    Begin
      Erro := '003';
    End
    Else
    Begin
      P1 := Pos('<codigo>',LowerCase(FXML))+8;
      P2 := Pos('</codigo>',LowerCase(FXML));
      if (P1 = 0) then Exit;
      Erro := copy(LowerCase(FXML), P1,P2-P1);
    End;

    FDoc := TXMLDocument.Create(nil) as IXMLDocument;

//    FDoc.LoadFromXML(FXML); - ALETRADO PARA RESOLVER PROBLEMA DO ACENTO QUE ESTAVA VINDO DENTRO DO ARQUIVO XML.

    FDoc.LoadFromXML( UTF8Decode(FXML));

    if ((FSalvarXML = True) and (ExtractFilePath(FPathXML) <>'')) then
    Begin
      if (ExtractFileName(FPathXML) = '') then
         FPathXML := IncludeTrailingBackslash(FPathXML) + 'DataXML.xml';
      ForceDirectories( ExtractFilePath(FPathXML));
      FDoc.SaveToFile(FPathXML);
    End;


    FNotFound :=False;
    if (Erro = '003') then            //  NENHUM REGISTRO ENCONTRADO
    Begin
      Result := True;
      FRecordCount :=-1;
      FPosition :=0;
      FNotFound :=True;
    End
    else if (Erro <> '000') then      //  ACONTECEU ALGUM ERRO
    Begin
      Result := False;
      FRecordCount :=-1;

      P1 := instr(0,LowerCase(FXML), '<codigo');
      P1 := instr(P1,LowerCase(FXML), '>');
      P2 := instr(P1+1,LowerCase(FXML), '<')-1;
      if (P1 = 0) then Exit;
      Erro := StringReplace(copy(LowerCase(FXML), P1,P2-P1), '"', '',[rfReplaceAll]);;

      P1 := instr(0,LowerCase(FXML), '<mensagem');
      P1 := instr(P1,LowerCase(FXML), '>');
      P2 := instr(P1+1,LowerCase(FXML), '<')-1;
      if (P1 = 0) then Exit;
      Desc := StringReplace(copy(LowerCase(FXML), P1,P2-P1), '"', '',[rfReplaceAll]);;


      FErro := StrToInt(Erro);
      FError := Desc;
    End
    else if (Erro = '000') then
    Begin

      if (FScript =False) then
      Begin

        if (UpperCase( LeftStr(Trim(FSQL),3)) = 'INS') then
        Begin
          P1 := instr(0,LowerCase(FXML), '<last_id');
          P1 := instr(P1,LowerCase(FXML), '>');
          P2 := instr(P1+1,LowerCase(FXML), '<')-1;
          if (P1 = 0) then Exit;
          FLastID := StrToInt64('0'+ StringReplace(copy(LowerCase(FXML), P1,P2-P1), '"', '',[rfReplaceAll]));
        End;

        P1 := instr(0,LowerCase(FXML), 'select rows=');
        P1 := instr(P1,LowerCase(FXML), '"');
        P2 := instr(P1+1,LowerCase(FXML), '"');

        if (P1 = 0) then Exit;

        Qtd := StringReplace(copy(LowerCase(FXML), P1,P2-P1), '"', '',[rfReplaceAll]);;

        FRecordCount := StrToInt(qtd);
        FPosition := 1;
        Result := True;
      End
      else
      begin

        if ( Assigned(FDoc) and Supports(FDoc.DocumentElement.DOMNode, IDOMNodeSelect, DOMNodeSelect)  )then
        Begin
//          DOMNodeList := DOMNodeSelect.selectNodes('/retorno/script[@num="2"]');
          DOMNodeList := DOMNodeSelect.selectNodes('/retorno/script');
          FScript_Count := DOMNodeList.length;
          if (FScript_Count>0) then
          Begin
            GravaCamposScript(1);
            FPosition := 1;
            Result := True;
          End;

        End;

      end;
    End;
  except on E : Exception do
  Begin
    FErro := 1;
    FError := 'Erro na execução do comando';
    if (ShowErrors = True) then raise Exception.Create('Erro na execução do comando' + #13#10 + E.Message);
  End;

  end;
End;

function TDataXML.FieldByIndex(Indice: Integer): String;
  Var Retorno: String;
      Campo: IXMLNode;
Begin

  if (FScript = False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[1]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[0];


  if (FPosition <= 0) then
  Begin
    FErro := 6;
    FError := 'DataXML não foi criado ou SQL não retornou registros.';
    if (ShowErrors = True) then raise Exception.Create('DataXML não foi criado ou SQL não retornou registros.');
//    Campo._Release;
    Exit;
  End;

  if (Indice < 0) then
  Begin
    FErro := 11;
    FError := 'Valor inválido para ocorrência de valor';
    if (ShowErrors = True) then raise Exception.Create('Valor inválido para ocorrência de valor');
//    Campo._Release;
    Exit;
  End;

  try
    if (FieldType( FieldName(Indice) ) = '252') then
//      Result := DecBase64(FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes[Indice].Text)
      Result := DecBase64(Campo.ChildNodes[(FPosition-1)].ChildNodes[Indice].Text)
    else
//      Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes[Indice].Text;
      Result := Campo.ChildNodes[(FPosition-1)].ChildNodes[Indice].Text;

  except
    FErro := 12;
    FError := 'Erro inesperado ao retornar valor.';
//    Campo._Release;
    if (ShowErrors = True) then raise exception.create('Erro inesperado ao retornar valor');
  end;

//  Campo._Release;

End;

function TDataXML.FieldByName(Field: String): String;
  Var Retorno: String;
      tamanho: Integer;
      Buff: TBytes;
      s_arq: TBytesStream;
      Encoding: TEncoding;
      Campo: IXMLNode;
      Campos: TXMLNodeList;
      F: TXMLDocument;
Begin

  Field := UpperCase(Field);

  if (FScript = False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[1]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[0];

  if (FPosition <= 0) then
  Begin
    FErro := 6;
    FError := 'DataXML não foi criado ou SQL não retornou registros.';
    if (ShowErrors = True) then raise Exception.Create('DataXML não foi criado ou SQL não retornou registros.');
//    Campo._Release;
    Exit;
  End;


//  if (FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes.FindNode(Field) = nil) then
  if (Campo.ChildNodes[(FPosition-1)].ChildNodes.FindNode(Field) = nil) then
  begin
    FErro := 7;
    FError := 'Nome do campo não foi encontrado.';
    if (ShowErrors = True) then raise Exception.Create('Nome do campo não foi encontrado.');
//    Campo._Release;
    Exit;
  end;

  try

    {if (FieldType(Field) = '252') then
      Result := DecBase64( FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes[Field].Text)
    else
      Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes[Field].Text;}


    if (FieldType(Field) = '252') then
    Begin
      Result := DecBase64( Campo.ChildNodes[(FPosition-1)].ChildNodes[Field].Text);
    End
    else
    Begin
      Campo := Campo.ChildNodes[(FPosition-1)].ChildNodes[Field];
      Result := Campo.Text;
    End;

  except
    FErro := 8;
    FError := 'Erro inesperado ao retornar valor.';
//    Campo._Release;
    if (ShowErrors = True) then raise exception.create('Erro inesperado ao retornar valor');
  end;

End;

function TDataXML.FieldsCount: Integer;
  Var Retorno: Integer;
      oNode: IXMLNode;
      Campo: IXMLNode;
Begin

  if (FScript = False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[1]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[1];

  if (FPosition <= 0) then
  Begin
    FErro := 6;
    FError := 'DataXML não foi criado ou SQL não retornou registros.';
    if (ShowErrors = True) then raise Exception.Create('DataXML não foi criado ou SQL não retornou registros.');
//    Campo._Release;
    Exit;
  End;


  try
    Result := Campo.ChildNodes.Count;
//    Campo._Release;

//    Result := FDoc.DocumentElement.ChildNodes.FindNode('campos').ChildNodes.Count;

//    Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes.Count;
  except
    FErro := 13;
    FError := 'Erro inesperado ao retornar quantidade de campos.';
    if (ShowErrors = True) then raise exception.create('Erro inesperado ao retornar quantidade de campos');
//    Campo._Release;
  end;

End;

function TDataXML.FieldExist(Field: String): Boolean;
  var Valor: String;
      Campo: IXMLNode;
Begin

  if (FScript = False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[1]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[0];

  Field := UpperCase(Field);
  Result := False;

  if (FPosition <= 0) then
  Begin
//    Campo._Release;
    Exit;
  End;

//  if (FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes.FindNode(Field) = nil) then
  if (Campo.ChildNodes[(FPosition-1)].ChildNodes.FindNode(Field) = nil) then
  begin
//    Campo._Release;
    Exit;
  end;

  try
//    Valor := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes[Field].Text;
    Valor := Campo.ChildNodes[(FPosition-1)].ChildNodes[Field].Text;
//    Campo._Release;
  except
//    Campo._Release;
    Exit;
  end;
  Result := True;
End;


function TDataXML.FieldName(Indice: Integer): String;
  Var Retorno: Integer;
      Campo: IXMLNode;
Begin

  if (FScript = False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[1]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[1];

  if (FPosition <= 0) then
  Begin
    FErro := 6;
    FError := 'DataXML não foi criado ou SQL não retornou registros.';
    if (ShowErrors = True) then raise Exception.Create('DataXML não foi criado ou SQL não retornou registros.');
//    Campo._Release;
    Exit;
  End;

  if (Indice < 0) then
  Begin
    FErro := 11;
    FError := 'Valor inválido para ocorrência de valor';
    if (ShowErrors = True) then raise Exception.Create('Valor inválido para ocorrência de valor');
//    Campo._Release;
    Exit;
  End;

  try
//    Result := FDoc.DocumentElement.ChildNodes.FindNode('campos').ChildNodes[Indice].LocalName;

    Result := Campo.ChildNodes[Indice].LocalName;
//    Campo._Release;

//    Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes.Count;
//    Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes[Indice].LocalName;
  except
    FErro := 14;
    FError := 'Erro inesperado ao retornar nome do campo.';
    if (ShowErrors = True) then raise exception.create('Erro inesperado ao retornar nome do campo');
//    Campo._Release;
  end;

End;

function TDataXML.FieldSize(Field: String): Int64;
  var Valor: String;
      Campo: IXMLNode;
Begin

  if (FScript = False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[1]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[1];

  Field := UpperCase(Field);
  Result := 0;

  if (FPosition <= 0) then
  Begin
//    Campo._Release;
    Exit;
  End;

//  if (FDoc.DocumentElement.ChildNodes.Nodes[2].ChildNodes.FindNode(Field) = nil) then
  if (Campo.ChildNodes.FindNode(Field) = nil) then
  begin
//    Campo._Release;
    Exit;
  end;

  try
//    Valor := FDoc.DocumentElement.ChildNodes.FindNode('campos').ChildNodes[Field].AttributeNodes['tamanho'].Text;
    Valor := Campo.ChildNodes[Field].AttributeNodes['tamanho'].Text;
//    Campo._Release;
//    Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes.Count;
//    Valor := FDoc.DocumentElement.ChildNodes.Nodes[2].ChildNodes[Field].AttributeNodes['tamanho'].Text; - ORIGINAL
  except
    Exit;
  end;
  Result := StrToInt64(Valor);

End;

function TDataXML.FieldType(Field: String): String;
  var Valor: String;
      Campo: IXMLNode;

Begin

  if (FScript=False) then
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[2]
  else
    Campo := FDoc.DocumentElement.ChildNodes.Nodes[FScript_Index].ChildNodes.Nodes[1];


  Field := UpperCase(Field);
  Result := '';

  if (FPosition <= 0) then
  Begin
//    Campo._Release;
    Exit;
  End;

//  if (FDoc.DocumentElement.ChildNodes.Nodes[2].ChildNodes.FindNode(Field) = nil) then
  if (Campo.ChildNodes.FindNode(Field) = nil) then
  begin
//    Campo._Release;
    Exit;
  end;

  try
    Valor := Campo.ChildNodes[Field].AttributeNodes['tipo'].Text;

//    Valor := FDoc.DocumentElement.ChildNodes.FindNode('campos').ChildNodes[Field].AttributeNodes['tipo'].Text;

//    Result := FDoc.DocumentElement.ChildNodes.Nodes[1].ChildNodes[(FPosition-1)].ChildNodes.Count;
//    Valor := FDoc.DocumentElement.ChildNodes.Nodes[2].ChildNodes[Field].AttributeNodes['tipo'].Text;
  except
//    Campo._Release;
    Exit;
  end;

//    Campo._Release;
  Result := Valor;

End;

function TDataXML.EOF: Boolean;
begin
  If (FPosition > FRecordCount) then
    Result := True
  else
    Result := False;
end;

function TDataXML.IsOnLine: Boolean;
  var P1,P2: Integer;
      l_xml, Data: string;
      R: boolean;
Begin

  if (uwsServer.defURL = '') then
  Begin
    uwsServer.defWSDL := 'http://www.digifarma.com.br/webservice/selectxml.php?wsdl';
    uwsServer.defURL  := 'http://www.digifarma.com.br/webservice/selectxml.php';
  End;


  R := True;

  try
    l_xml := wsfServer.selectxmlf(IntToHex(StrToInt(FormatDateTime('ddmmyyyy', Now)),2), 'select NOW() ''AGORA''', '');
  except
    R := False;
  end;

  if (R=False) then
  Begin
    try
      l_xml := wsfServer.selectxmlf(IntToHex(StrToInt(FormatDateTime('ddmmyyyy', Now)),2), 'select NOW() ''AGORA''', '');
    except
      R := False;
      Exit;
    end;
  End;

  P1 := Pos('<AGORA>',l_xml)+7;
  P2 := Pos('</AGORA>',l_xml);

  if (P1 = 0) then Exit;

  Data := copy(l_xml, P1,p2-p1);

  if ((Data <> '')  and (Length(Data) = 19)) then
  Begin
    Data := Copy(Data,9,2) + '/' + Copy(Data,6,2) + '/' + Copy(Data,1,4) + ' ' + Copy(Data,11,19);
    FDatetime := StrToDateTime(Data);
    Result := True;
  End
  else
    Exit;

End;

function TDataXML.ServidorWS(Nome: String): Boolean;
Begin

  uwsServer.defWSDL := 'http://' + Nome + '/webservice/selectxml.php?wsdl';
  uwsServer.defURL  := 'http://' + Nome + '/webservice/selectxml.php';
End;

procedure TDataXML.First;
begin
  FPosition := 1;
end;
procedure TDataXML.Previous;
begin
  if (FPosition >= 1) then
    FPosition := (FPosition -1)
  else
  Begin
    FErro := 4;
    FError := 'Início do Registro';
    if (ShowErrors = True) then raise exception.create('Início do Registro');
  End;
end;

procedure TDataXML.Next;
begin
  if (FPosition <= fRecordCount) then
    FPosition := (FPosition +1)

  else
  Begin
    FErro := 5;
    FError := 'Fim do Registro';
    if (ShowErrors = True) then raise exception.create('Fim do Registro');
  End;
end;

procedure TDataXML.Last;
begin
  FPosition := FRecordCount;
end;

function TDataXML.GetDatetime: TDateTime;
begin
  if (IsOnLine = False) then
  Begin
    FErro := 9;
    FError := 'Erro de conexão com o servidor';
    if (ShowErrors = True) then raise exception.create('Erro de conexão com o servidor');
    Exit;
  End;

  result := FDatetime;
end;

function TDataXML.GetRecordCount: Integer;
begin
  Result := FRecordCount;
end;

procedure TDataXML.Close;
begin
  Try

    FScript_Count :=0;
    FScript_Index :=0;

    FParametro.Clear;
    FLastID :=0;
    FScript:= False;
    FSQL :='';
    FREDE :='';
    FXML := '';
    if (FDoc <> nil) then
      FDoc := Nil;
    FRecordCount := -1;
    FErro := 0;
    FError := '';
    FPosition := 0;
    FNotFound :=False;
  Except
    Exit;
  End;
end;


end.


