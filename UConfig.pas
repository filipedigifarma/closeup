unit UConfig;

interface

uses
  Winapi.Windows, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, ibx.IBQuery;

type
  TFrmConfig = class(TForm)
    Lb_FTP: TLabel;
    Ed_ftp: TEdit;
    LB_user: TLabel;
    Ed_User: TEdit;
    Lb_senha: TLabel;
    Ed_Senha: TEdit;
    Bt_Save: TBitBtn;
    Lb_Cod: TLabel;
    Ed_Cod: TEdit;
    procedure Bt_SaveClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmConfig: TFrmConfig;

implementation

{$R *.dfm}

uses
  Udm,UPrincipal;

procedure TFrmConfig.Bt_SaveClick(Sender: TObject);
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

    oConfigProgram.g_s_FTP_HOST :=Tabela.FieldByName('ftp').AsString;
    oConfigProgram.g_s_FTP_USER:=Tabela.FieldByName('usuario').AsString;
    oConfigProgram.g_s_FTP_PWD:=Tabela.FieldByName('senha').AsString;
    oConfigProgram.s_COD_REDE:=Tabela.FieldByName('cod_rede').AsString;
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
  close;

end;

end.
