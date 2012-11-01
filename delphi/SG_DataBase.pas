unit SG_DataBase;

interface

uses
  dbtables;

type
  TSG_DataBase=class
  public
    class function getInstance:TSG_DataBase;
    procedure conectar(dataBaseName:string;login:boolean=true;usuario:string='';
      clave:string='');
    function openQuery(consulta:string):TQuery; overload;
    procedure openQuery(consulta:string;query:TQuery); overload;
    procedure execQuery(consulta:string);
    function getClaveInsertada:Variant;
  private
    Database:TDataBase;
  end;

implementation

var
  GB_dataBase:TSG_DataBase;

class function TSG_DataBase.getInstance:TSG_DataBase;
begin

  if GB_dataBase=nil then
    GB_dataBase:=TSG_DataBase.Create;

  Result:=GB_dataBase;

end;

procedure TSG_DataBase.conectar(dataBaseName:string;login:boolean=true;
  usuario:string='';clave:string='');
begin

  if Self.Database<>nil then
    exit;
    
  Self.Database:=TDataBase.create(nil);
  Self.Database.LoginPrompt:=login;
  Self.Database.Params.Add('USER NAME='+usuario);
  Self.Database.Params.Add('PASSWORD=' +clave);
  Self.Database.Params.Add('BLOB SIZE=-1');
  Self.Database.Params.Add('BLOBS TO CACHE=-1');
  Self.Database.AliasName:=dataBaseName;
  Self.Database.DatabaseName:=dataBaseName;
  Self.Database.Connected:=True;
end;

function TSG_DataBase.openQuery(consulta:string):TQuery;
begin

  Result:=TQuery.Create(nil);

  try

    Result.DatabaseName:=Self.Database.DatabaseName;
    Result.SQL.Text:=consulta;
    Result.Open;

  except

    Result.Free;
    raise;

  end;

end;

procedure TSG_DataBase.openQuery(consulta:string;query:TQuery);
begin
  query.DatabaseName:=Self.Database.DatabaseName;
  query.SQL.Text:=consulta;
  query.Open;
end;

procedure TSG_DataBase.execQuery(consulta:string);
var
  query:TQuery;
begin
  query:=TQuery.Create(nil);
  query.DatabaseName:=Self.Database.DatabaseName;
  query.SQL.Text:=consulta;
  query.ExecSQL;
end;

function TSG_DataBase.getClaveInsertada:Variant;
var
  consulta:string;
  query:TQuery;
begin

  consulta:=
    'select scope_identity() as id';

  query:=Self.openQuery(consulta);

  query.First;

  Result:=query['id'];

end;

end.
