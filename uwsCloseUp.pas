// ************************************************************************ //
// The types declared in this file were generated from data read from the
// WSDL File described below:
// WSDL     : http://www.digifarma.com.br/webservice/closeup/servico.php?wsdl
//  >Import : http://www.digifarma.com.br/webservice/closeup/servico.php?wsdl>0
// Encoding : ISO-8859-1
// Version  : 1.0
// (02/03/2016 14:47:46 - - $Rev: 76228 $)
// ************************************************************************ //

unit uwsCloseUp;

interface

uses Soap.InvokeRegistry, Soap.SOAPHTTPClient, System.Types, Soap.XSBuiltIns;

type

  // ************************************************************************ //
  // The following types, referred to in the WSDL document are not being represented
  // in this file. They are either aliases[@] of other types represented or were referred
  // to but never[!] declared in the document. The types from the latter category
  // typically map to predefined/known XML or Embarcadero types; however, they could also 
  // indicate incorrect WSDL documents that failed to declare or import a schema type.
  // ************************************************************************ //
  // !:string          - "http://www.w3.org/2001/XMLSchema"[]


  // ************************************************************************ //
  // Name      : server.ValidadorPortType
  // Namespace : uri:server.Validador
  // soapAction: uri:server.Validador/Validador
  // transport : http://schemas.xmlsoap.org/soap/http
  // style     : rpc
  // use       : encoded
  // binding   : server.ValidadorBinding
  // service   : server.Validador
  // port      : server.ValidadorPort
  // URL       : http://www.digifarma.com.br/webservice/closeup/servico.php
  // ************************************************************************ //
  server_ValidadorPortType = interface(IInvokable)
  ['{435726A2-98D3-2910-1673-433C30D17056}']
    function  confere_produto(const senha: string; const lista: string): string; stdcall;
  end;

function wsConfere(UseWSDL: Boolean=System.False; Addr: string=''; HTTPRIO: THTTPRIO = nil): server_ValidadorPortType;


implementation
  uses System.SysUtils;

function wsConfere(UseWSDL: Boolean; Addr: string; HTTPRIO: THTTPRIO): server_ValidadorPortType;
const
  defWSDL = 'http://www.digifarma.com.br/webservice/closeup/servico.php?wsdl';
  defURL  = 'http://www.digifarma.com.br/webservice/closeup/servico.php';
  defSvc  = 'server.Validador';
  defPrt  = 'server.ValidadorPort';
var
  RIO: THTTPRIO;
begin
  Result := nil;
  if (Addr = '') then
  begin
    if UseWSDL then
      Addr := defWSDL
    else
      Addr := defURL;
  end;
  if HTTPRIO = nil then
    RIO := THTTPRIO.Create(nil)
  else
    RIO := HTTPRIO;
  try
    Result := (RIO as server_ValidadorPortType);
    if UseWSDL then
    begin
      RIO.WSDLLocation := Addr;
      RIO.Service := defSvc;
      RIO.Port := defPrt;
    end else
      RIO.URL := Addr;
  finally
    if (Result = nil) and (HTTPRIO = nil) then
      RIO.Free;
  end;
end;


initialization
  { server.ValidadorPortType }
  InvRegistry.RegisterInterface(TypeInfo(server_ValidadorPortType), 'uri:server.Validador', 'ISO-8859-1', '', 'server.ValidadorPortType');
  InvRegistry.RegisterDefaultSOAPAction(TypeInfo(server_ValidadorPortType), 'uri:server.Validador/Validador');

end.