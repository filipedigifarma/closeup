program CloseUp;

{$IF CompilerVersion >= 21.0}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$IFEND}



{$R 'Win7UAC.res' 'Win32\Debug\Win7UAC.rc'}
{$R 'Win7UAC.res' 'Win32\Release\Win7UAC.rc'}

uses
  Vcl.Forms,
  Windows,
  UPrincipal in 'UPrincipal.pas' {FrmPrincipal},
  UDm in 'UDm.pas' {DM: TDataModule},
  uArqCli in 'uArqCli.pas',
  uArqProd in 'uArqProd.pas',
  uArqVenda in 'uArqVenda.pas',
  uwsCloseUp in 'uwsCloseUp.pas',
  uClasses in 'uClasses.pas',
  uFuncoes in 'uFuncoes.pas',
  uAutentica in 'C:\Users\Filipe\Dropbox\Programacao\delphi\Lib\uAutentica.pas',
  uMySql in 'C:\Users\Filipe\Dropbox\Programacao\delphi\Lib\uMySql.pas',
  uOleVariantEnum in 'C:\Users\Filipe\Dropbox\Programacao\delphi\Lib\uOleVariantEnum.pas',
  uwsselectxml in 'C:\Users\Filipe\Dropbox\Programacao\delphi\Lib\uwsselectxml.pas';

{$R *.res}
var
HprevHist: HWND;
hMutex : integer;

begin
  Application.Initialize;
  hMutex := CreateMutex(0, TRUE, 'CloseUp');
  if GetLastError = ERROR_ALREADY_EXISTS then
  begin
    Application.Terminate;
    Exit;
  end;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TFrmPrincipal, FrmPrincipal);
  Application.Run;
end.
