unit SG_ComboBox;

interface

uses
  stdctrls,
  classes,
  db;

type
  TSG_ComboBox = class(TComboBox)
  private
    { Private declarations }
    _esPosibleActualizarValores:boolean;
    _nombreTabla:string;
    _nombreCampoClave:string;
    _nombreCampoDescripcion:string;

    procedure crearValor(Sender: TObject);
    procedure modificarValor(Sender: TObject);
    procedure borrarValor(Sender: TObject);

    procedure agregarContextPopupActualizacion;

  protected
    { Protected declarations }

  public
    { Public declarations }
    procedure cargarValores;
    procedure addCodigo(codigo,descripcion:string);
    function getCodigo:string;
    procedure validarCodigo(descripcionParametro:string='';
      obligatorio:boolean=false);
    constructor Create(AOwner: TComponent); override;
    procedure seleccionar(descripcion:string);
    procedure seleccionarConNull(dataset:TDataset;nombreCampo:string);

  published
    { Published declarations }
    property esPosibleActualizarValores:boolean read
      _esPosibleActualizarValores write _esPosibleActualizarValores;
    property nombreTabla:string read _nombreTabla write _nombreTabla;
    property nombreCampoClave:string read _nombreCampoClave
       write _nombreCampoClave;
    property nombreCampoDescripcion:string read _nombreCampoDescripcion
      write _nombreCampoDescripcion;

  end;

  TSG_ComboBoxBit = class(TComboBox)
  private

  protected

  public
    procedure cargar;
    function toBit:string;
    procedure setValorSeleccionado(valor:string);
    class function getBitDescription(valor:string):string;

  published

  end;

  TSG_ComboBoxSexo = class(TComboBox)
  private

  protected

  public
    procedure cargar;
    function toChar:string;
    procedure setValorSeleccionado(valor:string);
    class function getCharDescription(valor:string):string;

  published

  end;

procedure Register;

procedure cargarComboBox(nombreTabla:string;comboBox:TComboBox;
  nombreCampoClave:string='id';nombreCampoDescripcion:string='descripcion');
procedure addCodigoComboBox(var combo:TComboBox;descripcion,codigo:string);
function getCodigoComboBox(combo:TComboBox):string;

implementation

uses
  SG_DataBase,
  dbtables,
  Sysutils,
  SG_Valor,
  Forms,
  menus;

type
  TCodigo=class(TObject)
    valor:string;
  end;

procedure Register;
begin
  RegisterComponents('SG', [TSG_ComboBox]);
  RegisterComponents('SG', [TSG_ComboBoxBit]);
  RegisterComponents('SG', [TSG_ComboBoxSexo]);
end;

procedure TSG_ComboBox.crearValor(Sender: TObject);
var
  valor:TFormValor;
begin

  Application.CreateForm(TFormValor,valor);

  try
  
    valor.Caption:='Crear valor';
    valor.nombreTabla:=Self.nombreTabla;
    valor.nombreCampoDescripcion:=Self.nombreCampoDescripcion;
    valor.BitBtnAceptar.OnClick:=valor.BitBtnAceptarClickCrear;

    valor.ShowModal;

    if valor.tareaRealizada then
    begin
      Self.cargarValores;
      Self.seleccionar(valor.EditValor.Text);
    end;

  finally

    valor.Free;

  end;

end;

procedure TSG_ComboBox.modificarValor(Sender: TObject);
var
  valor:TFormValor;
begin

  if Self.ItemIndex<0 then
    raise Exception.Create('Debe seleccionar un valor a ser modificado');

  Application.CreateForm(TFormValor,valor);

  try

    valor.valorAnterior:=Self.Text;
    valor.EditValor.Text:=Self.Text;
    valor.Caption:='Modificar valor '+valor.valorAnterior;
    valor.nombreTabla:=Self.nombreTabla;
    valor.nombreCampoDescripcion:=Self.nombreCampoDescripcion;
    valor.BitBtnAceptar.OnClick:=valor.BitBtnAceptarClickModificar;

    valor.ShowModal;

    if valor.tareaRealizada then
    begin
      Self.cargarValores;
      Self.seleccionar(valor.EditValor.Text);
    end;

  finally

    valor.Free;

  end;

end;

procedure TSG_ComboBox.borrarValor(Sender: TObject);
var
  valor:TFormValor;
begin

  if Self.ItemIndex<0 then
    raise Exception.Create('Debe seleccionar un valor a ser borrado');

  Application.CreateForm(TFormValor,valor);

  try
  
    valor.Caption:='Borrar valor '+Self.Text;
    valor.EditValor.Text:=Self.Text;
    valor.EditValor.Enabled:=false;
    valor.nombreTabla:=Self.nombreTabla;
    valor.nombreCampoDescripcion:=Self.nombreCampoDescripcion;
    valor.BitBtnAceptar.OnClick:=valor.BitBtnAceptarClickBorrar;

    valor.ShowModal;

    if valor.tareaRealizada then
      Self.cargarValores;

  finally

    valor.Free;

  end;

end;

procedure TSG_ComboBox.addCodigo(codigo,descripcion:string);
begin
  addCodigoComboBox(TComboBox(Self),descripcion,codigo);
end;

function TSG_ComboBox.getCodigo:string;
begin
  if Self.Text='' then
    Result:='null'
  else
    Result:=QuotedStr(getCodigoComboBox(TComboBox(Self)));
end;

procedure TSG_ComboBox.validarCodigo(descripcionParametro:string='';
  obligatorio:boolean=false);
begin

  if not obligatorio then
    if Self.Text='' then
      exit;

  if Self.getCodigo='' then
    raise Exception.Create('El valor seleccionado para el parametro '+
      descripcionParametro+' es invalido');

end;

procedure TSG_ComboBox.cargarValores;
begin
  Self.Items.Clear;
  addCodigoComboBox(TComboBox(Self),'','');
  cargarComboBox(Self.nombreTabla,Self,Self.nombreCampoClave,
    nombreCampoDescripcion);
end;

procedure TSG_ComboBox.agregarContextPopupActualizacion;
var
  menuItem:TMenuItem;
begin
  Self.PopupMenu:=TPopupMenu.Create(Self);

  menuItem:=TMenuItem.Create(Self.PopupMenu);
  menuItem.Caption:='Crear';
  menuItem.OnClick:=Self.crearValor;
  Self.PopupMenu.Items.Add(menuItem);

  menuItem:=TMenuItem.Create(Self.PopupMenu);
  menuItem.Caption:='Modificar';
  menuItem.OnClick:=Self.modificarValor;
  Self.PopupMenu.Items.Add(menuItem);

  menuItem:=TMenuItem.Create(Self.PopupMenu);
  menuItem.Caption:='Borrar';
  menuItem.OnClick:=Self.borrarValor;
  Self.PopupMenu.Items.Add(menuItem);

end;

constructor TSG_ComboBox.Create(AOwner: TComponent);
begin

  inherited Create(AOwner);

  //if Self.esPosibleActualizarValores then
    Self.agregarContextPopupActualizacion;

end;

procedure TSG_ComboBox.seleccionarConNull(dataset:TDataset;nombreCampo:string);
begin

  if dataset.FieldByName(nombreCampo).IsNull then
    Self.Text:=''
  else
    Self.seleccionar(dataset[nombreCampo]);

end;

procedure TSG_ComboBox.seleccionar(descripcion:string);
var
  i:integer;
begin

  for i:=0 to Self.Items.Count-1 do
    if Self.Items[i]=descripcion then
    begin
      Self.ItemIndex:=i;
      exit;
    end;

end;

procedure cargarComboBox(nombreTabla:string;comboBox:TComboBox;
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

    SG_ComboBox.addCodigoComboBox(comboBox,query['descripcion'],query['id']);

    query.Next;

  end;


end;

procedure AddCodigoComboBox(var combo:TComboBox;descripcion,
  codigo:string);
var
  _codigo:TCodigo;
begin
  _codigo:=TCodigo.Create;
  _codigo.valor:=codigo;
  combo.Items.AddObject(descripcion,_codigo);
end;

function buscarIndiceComboBox(combo:TComboBox):integer;
var
  i,
  longitudTexto,
  resultado:integer;
  encontrado,
  terminar:boolean;
  principioValor:string;
begin

  // Se obtiene la longitud del texto del edit del combo.
  longitudTexto:=Length(combo.Text);

  i:=0;
  encontrado:=false;
  terminar:=false;

  // Se itera sobre los items del combo.
  while not terminar do
  begin

    // Se obtiene el principio del valor del item i del combo.
    principioValor:=copy(combo.Items.Strings[i],1,longitudTexto);

    // En caso de que coincida con el texto del edit del combo...
    if uppercase(principioValor) = uppercase(combo.Text) then
      // ...entonces se encontro el elemento buscado.
      encontrado:=true
    else
      // Sino se sigue con el siguiente elemento.
      inc(i);

    // En caso que el indice supere la cantidad de elementos del combo o se
    // haya encontrado el elemento buscado se finaliza la iteracion.
    terminar:=((i = combo.Items.Count) or encontrado );
  end;

  if encontrado then
    resultado:=i
  else
    resultado:=-1;

  buscarIndiceComboBox:=resultado;

end;


function getCodigoComboBox(combo:TComboBox):string;
var
  resultado:string;
  indice:integer;
begin

  // En caso de estar seleccionado un item del combo...
//  if combo.ItemIndex >= 0 then
    // ...se obtiene el codigo de dicho item.
//    resultado:=TCodigo(combo.Items.Objects[combo.ItemIndex]).valor
//  else
  begin
    // Sino, se intenta encontrar un item en el combo que comience con la
    // descripcion ingresada en el edit del mismo.
    indice:=buscarIndiceComboBox(combo);

    // En caso de no haber tenido exito...
    if indice < 0 then
      // Se devuelve un string vacio.
      resultado:=''
    else
    begin
      // Sino se asigna el item encontrado al combo.
      combo.ItemIndex:=indice;
      // Se devuelve el codigo de dicho item.
      resultado:=TCodigo(combo.Items.Objects[combo.ItemIndex]).valor;
    end;
  end;

  getCodigoComboBox:=resultado;
end;

procedure TSG_ComboBoxBit.cargar;
begin
  addCodigoComboBox(TComboBox(Self),'','');
  addCodigoComboBox(TComboBox(Self),TSG_ComboBoxBit.getBitDescription('1'),'1');
  addCodigoComboBox(TComboBox(Self),TSG_ComboBoxBit.getBitDescription('0'),'0');
end;

function TSG_ComboBoxBit.toBit:string;
begin
  Result:=SG_ComboBox.getCodigoComboBox(Self);
end;

procedure TSG_ComboBoxBit.setValorSeleccionado(valor:string);
begin
  Self.Text:=TSG_ComboBoxBit.getBitDescription(valor);
end;

class function TSG_ComboBoxBit.getBitDescription(valor:string):string;
begin
  if valor='1' then
    Result:='Si'
  else if valor='0' then
    Result:='No';
end;

procedure TSG_ComboBoxSexo.cargar;
begin
  addCodigoComboBox(TComboBox(Self),'','');
  addCodigoComboBox(TComboBox(Self),TSG_ComboBoxSexo.getCharDescription('M'),'M');
  addCodigoComboBox(TComboBox(Self),TSG_ComboBoxSexo.getCharDescription('F'),'F');
end;

function TSG_ComboBoxSexo.toChar:string;
begin
  Result:=SG_ComboBox.getCodigoComboBox(Self);
end;

procedure TSG_ComboBoxSexo.setValorSeleccionado(valor:string);
var
  descripcion:string;
  i:integer;
begin

  descripcion:=TSG_ComboBoxSexo.getCharDescription(valor);

  for i:=0 to Self.Items.Count-1 do
  begin
    Self.ItemIndex:=i;
    if Self.Items[i]=descripcion then
      exit;
  end;

  Self.ItemIndex:=0;

end;

class function TSG_ComboBoxSexo.getCharDescription(valor:string):string;
begin
  if valor='M' then
    Result:='Masculino'
  else if valor='F' then
    Result:='Femenino';
end;

end.

