unit SG_Mail;

interface

uses
  NMsmtp,
  stdctrls;

procedure sendMail(puerto:integer;servidor,idUsuario,remitente,destinatarios,
  cuerpo,asunto:string;EncodeTypeUuCode:boolean=true);

function getRemitente:string;

function getMailByNroUsuario(NroUsuario:string):string;

function getMailsByNroUsuarioFromListBox(listBox:TListBox):string;

procedure sendGMail;

implementation

uses
  sysutils,
  dbtables,
  unitconexion,
  SG_ListBox,

  IdSMTP, IdSSLOpenSSL, IdMessage;

procedure sendMail(puerto:integer;servidor,idUsuario,remitente,destinatarios,
  cuerpo,asunto:string;EncodeTypeUuCode:boolean=true);
var
  NMSMTP:TNMSMTP;
begin

  NMSMTP:=TNMSMTP.Create(nil);

  NMSMTP.Host:=servidor;
  NMSMTP.Port:=puerto;
  NMSMTP.UserID:=idUsuario;

  NMSMTP.PostMessage.FromAddress:=remitente;
  NMSMTP.PostMessage.ToAddress.Text:=destinatarios;
  NMSMTP.PostMessage.Body.Text:=cuerpo;
  NMSMTP.PostMessage.Subject:=asunto;

  if EncodeTypeUuCode then
    NMSMTP.EncodeType:=uuCode
  else
    NMSMTP.EncodeType:=uuMime;

  try

    NMSMTP.Connect;
    NMSMTP.SendMail;
    NMSMTP.Disconnect;

  except

    on Exception do
    begin
      NMSMTP.Disconnect;
      NMSMTP.Free;
      raise;
    end;

  end;

  NMSMTP.Free;

end;

function getRemitente:string;
var
  consulta:string;
  query:TQuery;
begin

  consulta:=
    'select EMail '+
    'from MaestroUsuarios '+
    'where IdUsuario = suser_sname()';

  query:=AbreQuery(consulta);

  Result:=query['EMail'];

end;

function getMailByNroUsuario(NroUsuario:string):string;
var
  consulta:string;
  query:TQuery;
begin

  consulta:=
    'select EMail '+
    'from MaestroUsuarios '+
    'where NroUsuario = '+NroUsuario;

  query:=AbreQuery(consulta);

  Result:=query['EMail'];

end;

function getMailsByNroUsuarioFromListBox(listBox:TListBox):string;
var
  i:integer;
begin

  for i:=0 to listBox.Items.Count-1 do
  begin

    if i>0 then
      Result:=Result+'; ';

    Result:=Result+getMailByNroUsuario(getCodigoListBox(listBox,i));

  end;

end;

procedure sendGMail;
var
  IdSMTP1: TIdSMTP;
  IdSSLIOHandlerSocketOpenSSL1:TIdSSLIOHandlerSocketOpenSSL;
  IdMessage1:TIdMessage;
  //IdSSLIOHandlerSocketOpenSSL1StatusInfo:TIdSSLIOHandlerSocketOpenSSL1StatusInfo;
begin

  IdSSLIOHandlerSocketOpenSSL1:=TIdSSLIOHandlerSocketOpenSSL.Create;
  IdSMTP1:=TIdSMTP.Create;
  IdMessage1:=TIdMessage.create;

  with IdSSLIOHandlerSocketOpenSSL1 do
  begin
    Destination := 'smtp.gmail.com:587';
    Host := 'smtp.gmail.com';
    //MaxLineAction := maException;
    Port := 587;
    SSLOptions.Method := sslvTLSv1;
    SSLOptions.Mode := sslmUnassigned;
    SSLOptions.VerifyMode := [];
    SSLOptions.VerifyDepth := 0;
    //OnStatusInfo := IdSSLIOHandlerSocketOpenSSL1StatusInfo;
  end;

  with IdSMTP1 do
  begin
    //OnStatus := IdSMTP1Status;
    IOHandler := IdSSLIOHandlerSocketOpenSSL1;
    Host := 'smtp.gmail.com';
    Password := 'c3r3br1n';
    Port := 587;
    //SASLMechanisms := <>;
    //seTLS := utUseExplicitTLS;
    Username := 'sebas.goldberg';
  end;

  IdSMTP1.Connect;
  IdSMTP1.Send(IdMessage1);
  IdSMTP1.Disconnect;


end;

end.
