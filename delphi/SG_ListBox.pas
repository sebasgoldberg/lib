unit SG_ListBox;

interface

uses
  stdctrls;

procedure addCodigoListBox(listBox:TCustomListBox;codigo,descripcion:string);
function getIndexCodigoListBox(listBox:TCustomListBox;codigo:string):integer;
function existeCodigoListBox(listBox:TListBox;codigo:string):boolean;
procedure addComboBoxObjectToListBox(comboBox:TComboBox;listBox:TListBox);
procedure quitarCodigoListBox(listBox:TListBox);
function getCodigoListBox(listBox:TCustomListBox;indice:integer):string;
procedure listBoxToTable(listBox:TListBox;nombreTabla:string;
  nombreCampoClave:string;valorCampoClave:string;
  nombreCampoValoresListBox:string);
procedure cargarListBox(nombreTabla:string;listBox:TCustomListBox;
  nombreCampoClave:string='id';nombreCampoDescripcion:string='descripcion');

implementation

uses
  dbtables,
  SG_DataBase,
  SG_ComboBox;

type
  TCodigoListBox=class
    public
      valor:string;
  end;

procedure addCodigoListBox(listBox:TCustomListBox;codigo,descripcion:string);
var
  _codigo:TCodigoListBox;
begin
  _codigo:=TCodigoListBox.Create;
  _codigo.valor:=codigo;
  listBox.Items.AddObject(descripcion,_codigo);
end;

function getIndexCodigoListBox(listBox:TCustomListBox;codigo:string):integer;
var
  i:integer;
  encontrado,
  seguir:boolean;
  _codigo:TCodigoListBox;
begin

  encontrado:=false;
  seguir:=true;
  i:=0;

  if listBox.items.count = 0 then
  begin
    Result:=-1;
    exit;
  end;

  while seguir do
  begin
    _codigo:=TCodigoListBox(listBox.items.objects[i]);
    encontrado:=(_codigo.valor=codigo);
    i:=i+1;
    seguir:=not ( encontrado or (i=listBox.items.count) );
  end;

  if encontrado then
    Result:=i-1
  else
    Result:=-1;

end;

function existeCodigoListBox(listBox:TListBox;codigo:string):boolean;
var
  i:integer;
  encontrado,
  seguir:boolean;
  _codigo:TCodigoListBox;
begin

  encontrado:=false;
  seguir:=true;
  i:=0;

  if listBox.items.count = 0 then
  begin
    Result:=false;
    exit;
  end;

  while seguir do
  begin
    _codigo:=TCodigoListBox(listBox.items.objects[i]);
    encontrado:=(_codigo.valor=codigo);
    i:=i+1;
    seguir:=not ( encontrado or (i=listBox.items.count) );
  end;

  Result:=encontrado;

end;

procedure addComboBoxObjectToListBox(comboBox:TComboBox;listBox:TListBox);
var
  codigo:string;
begin

  codigo:=getCodigoComboBox(comboBox);

  if codigo='' then
    exit;

  if not existeCodigoListBox(listBox,codigo) then
    addCodigoListBox(listBox,codigo,comboBox.Text);

  listBox.ItemIndex:=listBox.Items.Count-1;

end;

procedure quitarCodigoListBox(listBox:TListBox);
begin
  if listBox.ItemIndex>=0 then
    listBox.Items.Delete(listBox.ItemIndex);
end;

function getCodigoListBox(listBox:TCustomListBox;indice:integer):string;
var
  codigo:TCodigoListBox;
begin
  codigo:=TCodigoListBox(listBox.Items.Objects[indice]);
  Result:=codigo.valor;
end;

procedure listBoxToTable(listBox:TListBox;nombreTabla:string;
  nombreCampoClave:string;valorCampoClave:string;
  nombreCampoValoresListBox:string);
var
  i:integer;
  consulta:string;
begin

  consulta:=
    'delete from '+nombreTabla+' '+
    'where '+nombreCampoClave+' = '+valorCampoClave;

  TSG_DataBase.getInstance.execQuery(consulta);

  for i:=0 to listBox.Items.Count-1 do
  begin

    consulta:=
      'insert into '+nombreTabla+' '+
      '('+
        nombreCampoClave+', '+
        nombreCampoValoresListBox+
      ') '+
      'values('+
      valorCampoClave+', '+
      getCodigoListBox(listBox,i)+
      ')';

    TSG_DataBase.getInstance.execQuery(consulta);

  end;

end;

procedure cargarListBox(nombreTabla:string;listBox:TCustomListBox;
  nombreCampoClave:string='id';nombreCampoDescripcion:string='descripcion');
var
  consulta:string;
  query:TQuery;
begin

  consulta:=
    'select '+
      nombreCampoClave+' as id, '+
      nombreCampoDescripcion+' as descripcion '+
    'from '+nombreTabla+' '+
    'order by descripcion';

  query:=TSG_DataBase.getInstance.openQuery(consulta);

  query.First;

  while not query.Eof do
  begin

    addCodigoListBox(listBox,query['id'],query['descripcion']);

    query.Next;

  end;


end;

end.
