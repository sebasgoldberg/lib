unit SG_CheckListBox;

interface

uses
  stdctrls,
  classes,
  checklst;

type
  TSG_CheckListBox = class(TCheckListBox)
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
    constructor Create(AOwner: TComponent); override;
    function getCodigosSeleccionados:TStrings;
    procedure seleccionar(codigos:TStrings);
    function getWhereCondition(fieldName:string):string;

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

  TSG_CheckListBoxClaveExterna=class(TSG_CheckListBox)
  private
    _nombreTablaPrincipal:string;
    _nombreCampoClaveTablaPrincipal:string;
    _nombreCampoValorTablaPrincipal:string;
    cargadaTablaPrincipal:boolean;

  public
    constructor Create(AOwner: TComponent); override;
    procedure cargarValoresClave(clave:string);
    function getCodigosTablaPrincipal(clave:string):TStrings;
    procedure saveCodigosTablaPrincipal(clave:string);

  published
    property nombreTablaPrincipal:string read _nombreTablaPrincipal
      write _nombreTablaPrincipal;
    property nombreCampoClaveTablaPrincipal:string read
      _nombreCampoClaveTablaPrincipal write _nombreCampoClaveTablaPrincipal;
    property nombreCampoValorTablaPrincipal:string read
      _nombreCampoValorTablaPrincipal write _nombreCampoValorTablaPrincipal;
  end;

procedure Register;

implementation

uses
  SG_DataBase,
  dbtables,
  Sysutils,
  SG_Valor,
  Forms,
  menus,
  SG_ListBox;

type
  TCodigo=class(TObject)
    valor:string;
  end;

procedure Register;
begin
  RegisterComponents('SG', [TSG_CheckListBox]);
  RegisterComponents('SG', [TSG_CheckListBoxClaveExterna]);
end;

procedure TSG_CheckListBox.crearValor(Sender: TObject);
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
      Self.cargarValores;

  finally

    valor.Free;

  end;

end;

procedure TSG_CheckListBox.modificarValor(Sender: TObject);
var
  valor:TFormValor;
begin

  if Self.ItemIndex<0 then
    raise Exception.Create('Debe seleccionar un valor a ser modificado');

  Application.CreateForm(TFormValor,valor);

  try

    valor.valorAnterior:=Self.Items[Self.itemindex];
    valor.EditValor.Text:=Self.Items[Self.itemindex];
    valor.Caption:='Modificar valor '+valor.valorAnterior;
    valor.nombreTabla:=Self.nombreTabla;
    valor.nombreCampoDescripcion:=Self.nombreCampoDescripcion;
    valor.BitBtnAceptar.OnClick:=valor.BitBtnAceptarClickModificar;

    valor.ShowModal;

    if valor.tareaRealizada then
      Self.cargarValores;

  finally

    valor.Free;

  end;

end;

procedure TSG_CheckListBox.borrarValor(Sender: TObject);
var
  valor:TFormValor;
begin

  if Self.ItemIndex<0 then
    raise Exception.Create('Debe seleccionar un valor a ser borrado');

  Application.CreateForm(TFormValor,valor);

  try

    valor.Caption:='Borrar valor '+Self.Items[Self.itemindex];
    valor.EditValor.Text:=Self.Items[Self.itemindex];
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

procedure TSG_CheckListBox.addCodigo(codigo,descripcion:string);
begin
  addCodigoListBox(TCustomListBox(Self),descripcion,codigo);
end;

function TSG_CheckListBox.getCodigosSeleccionados:TStrings;
var
  i:integer;
begin

  Result:=TStringList.Create;

  for i:=0 to Self.Items.Count-1 do
    if Self.Checked[i] then
      Result.Add(getCodigoListBox(Self,i));

end;

procedure TSG_CheckListBox.seleccionar(codigos:TStrings);
var
  i:integer;
  indiceCodigoEncontrado:integer;
begin

  for i:=0 to codigos.Count-1 do
  begin

    indiceCodigoEncontrado:=getIndexCodigoListBox(Self,codigos[i]);

    if indiceCodigoEncontrado>=0 then
      Self.Checked[indiceCodigoEncontrado]:=true;

  end;

end;

procedure TSG_CheckListBox.cargarValores;
var
  codigos:TStrings;
begin
  codigos:=Self.getCodigosSeleccionados;
  Self.Items.Clear;
  cargarListBox(Self.nombreTabla,Self,Self.nombreCampoClave,
    nombreCampoDescripcion);
  Self.seleccionar(codigos);
end;

procedure TSG_CheckListBox.agregarContextPopupActualizacion;
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

constructor TSG_CheckListBox.Create(AOwner: TComponent);
begin

  inherited Create(AOwner);

  //if Self.esPosibleActualizarValores then
    Self.agregarContextPopupActualizacion;

end;

constructor TSG_CheckListBoxClaveExterna.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Self.cargadaTablaPrincipal:=false;
end;

function TSG_CheckListBoxClaveExterna.getCodigosTablaPrincipal(clave:string):
  TStrings;
var
  consulta:String;
  query:TQuery;
begin

  Result:=TStringList.Create;

  try

    if clave='' then
      exit;
  
    consulta:=
      'select '+Self.nombreCampoValorTablaPrincipal+' as valor '+
      'from '+Self.nombreTablaPrincipal+' '+
      'where '+Self.nombreCampoClaveTablaPrincipal+' = '+clave;
  
    query:=TSG_DataBase.getInstance.openQuery(consulta);
  
    query.First;
  
    while not query.Eof do
    begin
      Result.Add(IntToStr(query['valor']));
      query.Next;
    end;

  except

    Result.Free;
    raise;

  end;

end;

procedure TSG_CheckListBoxClaveExterna.saveCodigosTablaPrincipal(clave:string);
var
  consulta:string;
  codigosSeleccionados:TStrings;
  i:integer;
begin

  codigosSeleccionados:=Self.getCodigosSeleccionados;

  try

    consulta:=
      'delete '+
      'from '+Self.nombreTablaPrincipal+' '+
      'where '+Self.nombreCampoClaveTablaPrincipal+' = '+clave;

    TSG_DataBase.getInstance.execQuery(consulta);

    for i:=0 to codigosSeleccionados.Count-1 do
    begin

      consulta:=
        'INSERT INTO '+Self.nombreTablaPrincipal+' '+
          '('+Self.nombreCampoClaveTablaPrincipal+', '+
          Self.nombreCampoValorTablaPrincipal+') '+
        'VALUES '+
          '('+clave+', '+
          codigosSeleccionados[i]+')';

      TSG_DataBase.getInstance.execQuery(consulta);

    end;

  finally

    codigosSeleccionados.Free;

  end;

end;

function TSG_CheckListBox.getWhereCondition(fieldName:string):string;
var
  valores:TStrings;
  i:integer;
begin

  valores:=Self.getCodigosSeleccionados;

  try

    if valores.Count=0 then
      Result:=''
    else
    begin

      Result:='and '+fieldName+' in ('+valores[0];

      for i:=1 to valores.Count-1 do
        Result:=Result+', '+valores[i];

      Result:=Result+') ';

    end;

  finally

    valores.Free;

  end;

end;

procedure TSG_CheckListBoxClaveExterna.cargarValoresClave(clave:string);
begin
  Self.cargarValores;

  if not Self.cargadaTablaPrincipal then
  begin
    Self.seleccionar(Self.getCodigosTablaPrincipal(clave));
    Self.cargadaTablaPrincipal:=True;
  end;

end;


end.

