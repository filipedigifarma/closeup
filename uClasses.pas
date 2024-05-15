unit uClasses;

interface

uses System.Classes, Vcl.ComCtrls, System.SysUtils, System.Generics.Collections;

type
  TListColumnTurbo = class helper for TListColumn
  Private
    function GetColumnName: String;
    procedure SetColumnName(Value: String);
  published
    class var
        MetaData: TDictionary<TListColumn, string>;
    class constructor SetUp;
    class destructor TearDown;
    property Name: string read GetColumnName write SetColumnName;
  published
  end;



type
  TListItemTurbo = class helper for TListItem
  Private
    function GetObject: TObject;
    procedure SetObject(Value: TObject);
    function LeCampo(Nome: String): String;
    procedure GravaCampo(Nome: String; Value: String);
  public
    class var
        MetaDataObject: TDictionary<TListItem, TObject>;
    class constructor SetUp;
    class destructor TearDown;
    procedure Ins(Nome: String; Valor: String);
    property FieldByName[FieldName: String]: string read LeCampo write GravaCampo;
    property Objeto: TObject read GetObject write SetObject;
  published
  end;

implementation


function TListColumnTurbo.GetColumnName: String;
begin
  if MetaData.ContainsKey(Self) then
    Result:=MetaData.Items[Self]
  else
    Result:=''; //or whatever
end;

procedure TListColumnTurbo.SetColumnName(Value: String);
begin
  MetaData.AddOrSetValue(Self, Value);
end;

class constructor TListColumnTurbo.SetUp;
begin
  MetaData:=TDictionary<TListColumn, string>.Create;
end;

class destructor TListColumnTurbo.TearDown;
begin
  MetaData.Free;
end;

function TListItemTurbo.GetObject: TObject;
begin
  if MetaDataObject.ContainsKey(Self) then
    Result:=MetaDataObject.Items[Self]
  else
    Result:=nil; //or whatever
end;

procedure TListItemTurbo.SetObject(Value: TObject);
begin
  MetaDataObject.AddOrSetValue(Self, Value);
end;

class constructor TListItemTurbo.SetUp;
begin
  MetaDataObject:=TDictionary<TListItem, TObject>.Create;
end;

class destructor TListItemTurbo.TearDown;
begin
  MetaDataObject.Free;
end;

function TListItemTurbo.LeCampo(Nome: String): String;
  var I: Integer;
      lw: TListView;
Begin

  lw := TListView(Self.ListView);


  for I := 0 to (lw.Columns.Count-1) do
    if ( AnsiLowerCase(lw.Columns[I].Caption) = AnsiLowerCase(Nome) )  then
    Begin
      if (I=0) then
        Result := Self.Caption
      else
        Result := Self.SubItems[I-1];
    End;
End;

procedure TListItemTurbo.GravaCampo(Nome: String; Value: String);
  var I: Integer;
      lw: TListView;
Begin

  lw := TListView(Self.ListView);

  for I := 0 to (lw.Columns.Count-1) do
    if ( AnsiLowerCase(lw.Columns[I].Caption) = AnsiLowerCase(Nome) )  then
    Begin
      if (I=0) then
        Self.Caption := Value
      else
        Self.SubItems[I-1] := Value;
    End;
End;

procedure TListItemTurbo.Ins(Nome: String; Valor: String);
  var I: Integer;
      lw: TListView;
Begin

  lw := TListView(Self.ListView);

  if (Self.SubItems.Count=0) then
    for I := 0 to (lw.Columns.Count-1) do
    Begin
        Self.SubItems.Insert(I , '');
    End;

  for I := 0 to (lw.Columns.Count-1) do
    if ( AnsiLowerCase(lw.Columns[I].Caption) = AnsiLowerCase(Nome) )  then
    Begin
      if (I=0) then
        Self.Caption := Valor
      else
        Self.SubItems[(I-1)] := Valor;
    End;
End;



end.
