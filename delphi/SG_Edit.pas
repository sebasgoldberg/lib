unit SG_Edit;

interface

uses
  stdctrls;

  procedure validarDatoIngresado(edit:TEdit;descripcionEdit:string);
  function esDigito(caracter:char):boolean;
  procedure validarDatoNumerico(edit:TEdit;descripcionEdit:string);
  procedure validarDatoNumericoEntero(edit:TEdit;descripcionEdit:string);
  procedure validarDatoNumericoConDecimales(edit:TEdit;descripcionEdit:string);

implementation

uses
  sysutils;

procedure validarDatoIngresado(edit:TEdit;descripcionEdit:string);
begin
  if edit.Text='' then
    raise Exception.Create('El parametro '+descripcionEdit+' es obligatorio');
end;

function esDigito(caracter:char):boolean;
begin

  Result:=
    (caracter = '0') or
    (caracter = '1') or
    (caracter = '2') or
    (caracter = '3') or
    (caracter = '4') or
    (caracter = '5') or
    (caracter = '6') or
    (caracter = '7') or
    (caracter = '8') or
    (caracter = '9');

end;

procedure validarDatoNumerico(edit:TEdit;descripcionEdit:string);
var
  longitud,
  i:integer;
begin

  if edit.Text = '' then
    exit;

  longitud:=length(edit.Text);

  for i:=1 to longitud do
    if not esDigito(edit.Text[i]) then
      raise Exception.Create('El parametro '+descripcionEdit+' solo acepta '+
        'caracteres numericos');

end;

procedure validarDatoNumericoEntero(edit:TEdit;descripcionEdit:string);
begin

  if edit.Text = '' then
    exit;

  try

    StrToInt(edit.Text);

  except

    on EConvertError do
      raise Exception.Create('El parametro '+descripcionEdit+' solo acepta '+
        'valores numericos');

  end;

end;

procedure validarDatoNumericoConDecimales(edit:TEdit;descripcionEdit:string);
begin

  if edit.Text = '' then
    exit;

  try

    StrToFloat(edit.Text);

  except

    on EConvertError do
      raise Exception.Create('El parametro '+descripcionEdit+' solo acepta '+
        'valores numéricos decimales');

  end;

end;

end.
