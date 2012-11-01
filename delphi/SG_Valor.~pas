unit SG_Valor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TFormValor = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    EditValor: TEdit;
    BitBtnAceptar: TBitBtn;
    BitBtnCancelar: TBitBtn;
    procedure BitBtnCancelarClick(Sender: TObject);
    procedure BitBtnAceptarClick(Sender: TObject);
    procedure BitBtnAceptarClickCrear(Sender: TObject);
    procedure BitBtnAceptarClickModificar(Sender: TObject);
    procedure BitBtnAceptarClickBorrar(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    nombreTabla:string;
    nombreCampoDescripcion:string;
    tareaRealizada:boolean;
    valorAnterior:string;

  end;

var
  FormValor: TFormValor;

implementation

{$R *.DFM}

uses
  SG_DataBase,
  dbtables,
  SG_Messages;

procedure TFormValor.BitBtnCancelarClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TFormValor.BitBtnAceptarClick(Sender: TObject);
begin
  // Se llama al correspondiente método.
end;

procedure TFormValor.BitBtnAceptarClickCrear(Sender: TObject);
var
  consulta:string;
  query:TQuery;
begin

  if Self.EditValor.Text='' then
    raise Exception.Create('Ingrese un valor');

  consulta:=
    'select '+Self.nombreCampoDescripcion+' '+
    'from '+Self.nombreTabla+' '+
    'where '+Self.nombreCampoDescripcion+' = '+QuotedStr(Self.EditValor.Text);

  query:=TSG_DataBase.getInstance.openQuery(consulta);

  query.First;

  if not query.Eof then
    raise Exception.Create('El valor ingresado ya existe');

  consulta:=
    'insert into '+Self.nombreTabla+' '+
      '('+Self.nombreCampoDescripcion+') '+
    'values '+
      '('+QuotedStr(Self.EditValor.Text)+')';

  TSG_DataBase.getInstance.execQuery(consulta);

  information('Se ha agregado el valor '+Self.EditValor.Text);

  Self.tareaRealizada:=true;

  Self.Close;

end;

procedure TFormValor.BitBtnAceptarClickModificar(Sender: TObject);
var
  consulta:string;
  query:TQuery;
begin

  if Self.EditValor.Text='' then
    raise Exception.Create('Ingrese un valor');

  consulta:=
    'select '+Self.nombreCampoDescripcion+' '+
    'from '+Self.nombreTabla+' '+
    'where '+Self.nombreCampoDescripcion+' = '+QuotedStr(Self.EditValor.Text);

  query:=TSG_DataBase.getInstance.openQuery(consulta);

  query.First;

  if not query.Eof then
    raise Exception.Create('El valor ingresado ya existe');

  consulta:=
    'update '+Self.nombreTabla+' '+
    'set '+
      Self.nombreCampoDescripcion+' = '+QuotedStr(Self.EditValor.Text)+' '+
    'where '+
      Self.nombreCampoDescripcion+' = '+QuotedStr(Self.valorAnterior);

  TSG_DataBase.getInstance.execQuery(consulta);

  information('Se ha modificado el valor '+Self.valorAnterior+' por el nuevo '+
    'valor '+Self.EditValor.Text);

  Self.tareaRealizada:=true;

  Self.Close;

end;

procedure TFormValor.BitBtnAceptarClickBorrar(Sender: TObject);
var
  consulta:string;
begin

  if Self.EditValor.Text='' then
    raise Exception.Create('Ingrese un valor');

  consulta:=
    'delete '+
    'from '+Self.nombreTabla+' '+
    'where '+Self.nombreCampoDescripcion+' = '+QuotedStr(Self.EditValor.Text);

  TSG_DataBase.getInstance.execQuery(consulta);

  information('Se ha borrado el valor '+Self.EditValor.Text);

  Self.tareaRealizada:=true;

  Self.Close;

end;

procedure TFormValor.FormActivate(Sender: TObject);
begin
  Self.tareaRealizada:=false;
end;

end.
