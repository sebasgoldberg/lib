{
--------------------------------------------------------------------------------
                    L O G   D E   M O D I F I C A C I O N E S
--------------------------------------------------------------------------------
ID: SG006
Fecha: 13/02/2008
Programador: Juan Sebasti�n Goldberg

Descripci�n: Se agregan dos rutinas:

1) function Excel2007(grid: TStringGrid; nombreHoja: string):boolean; overload;
Esta rutina muestra un TSaveDialog para que el usuario seleccione el archivo
que desea guardar y el formato. Los formatos soportados son:
  a) Libro de Microsoft Office Excel (*.xls): El formato de la version instalada
  en el equipo (ej: si la version es excel 2007 se guarda con dicho formato).
  b) Libro de Microsoft Excel 97 - Excel 2003 y 5.0/95 (*.xls): El formato que
  hace compatible las distintas versiones de excel.

2) function Excel2007(grid: TStringGrid; nombreHoja, archivo: string;
  formato:integer): boolean; overload;
Esta rutina guarda los datos pasados en grid en un archivo con el formato Excel
especificado (ver documentacion).

--------------------------------------------------------------------------------
ID: SG016
Fecha: 18/04/2008
Programador: Juan Sebasti�n Goldberg

Descripci�n: Se agregan rutinas para guardar en excel 2007 a partir de un data
set.
--------------------------------------------------------------------------------
ID: SG020
Fecha: 27/04/2008
Programador: Juan Sebasti�n Goldberg

Descripci�n: Se quita la funci�n getMaxRegistrosExcel.
--------------------------------------------------------------------------------
ID: SG039
Fecha: 17/03/2009
Programador: Juan Sebasti�n Goldberg

Descripci�n: Se agrega la clase TExcel, la misma permite crear archivos excel
con m�s de una hoja por libro.
--------------------------------------------------------------------------------
}
unit SG_Excel;

interface

uses
   Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
   Grids, StdCtrls, Buttons, ExtCtrls, Db,ComObj, DBTables, DBGrids;

const
// Constantes utilizadas para Excel 2003
  xlWorkbookNormal=-4143; // Libro de Microsoft Office Excel (*.xls)

// Constantes utilizadas para Excel 2007
  xlOpenXMLWorkbook=51; // Libro de Excel (*.xlsx)
  xlExcel8=56; //Libro de Excel 97-2003 (*.xls)

type
  TExcel=class
  public
    class function crearArchivoExcel(archivo:string):TExcel;
    procedure agregarHoja(datos:TDataSet;nombreHoja:string);
    procedure guardar;
    procedure Free;
    function getArchivo:string;

    class function createFromFile(filename:string):TExcel;
    function getActiveSheet:Variant;
    function getMaxRows:integer;
    class function getNumeroColumna(columna:string):integer;
    procedure save;
    procedure quit;
    function getStringFromCell(row,column:integer):string;
    procedure setValueToCell(row,column:integer;value:Variant);
    class function getLetraColumna(numeroColumna:integer):string;

  private
    aplicacion:Variant;
    cantidadHojasAgregadas:integer;
    archivo:string;
    formato:integer;

    constructor Create(archivo:string;formato:integer); overload;
    constructor Create; overload;
  end;

{
Se muestra un dialogo para que el usuario seleccione una ruta y nombre de
archivo para guardar los datos del grid. Tambi�n en dicho di�logo se puede
seleccionar el formato del archivo a guardar. En caso de que el usuario haya
seleccionado el archivo en el que se guardar�n los datos y en funci�n del
formato seleccionado se guarda el archivo.
- grid: Grid que conteniendo los datos a guardar en el archivo.
- nombreHoja: Nombre a asignar a la hoja donde se guardar�n los datos.
- Result: True si se guardo el archivo. Sino false.
}
function Excel2007(grid: TStringGrid; nombreHoja: string;
                   MostrarColumnasOcultas:Boolean = false):boolean; overload;

{
Se muestra un dialogo para que el usuario seleccione una ruta y nombre de
archivo para guardar los datos del grid. Tambi�n en dicho di�logo se puede
seleccionar el formato del archivo a guardar. En caso de que el usuario haya
seleccionado el archivo en el que se guardar�n los datos y en funci�n del
formato seleccionado se guarda el archivo.
- grid: Grid que conteniendo los datos a guardar en el archivo.
- nombreHoja: Nombre a asignar a la hoja donde se guardar�n los datos.
- Result: True si se guardo el archivo. Sino false.
}
function Excel2007(DBGrid:TDBGrid;nombreHoja:string;
                   MostrarColumnasOcultas:Boolean = False):boolean; overload;

{
Se muestra un dialogo para que el usuario seleccione una ruta y nombre de
archivo para guardar los datos del data set. Tambi�n en dicho di�logo se puede
seleccionar el formato del archivo a guardar. En caso de que el usuario haya
seleccionado el archivo en el que se guardar�n los datos y en funci�n del
formato seleccionado se guarda el archivo.
- dataSet: Data set conteniendo los datos a guardar en el archivo.
- nombreHoja: Nombre a asignar a la hoja donde se guardar�n los datos.
- Result: True si se guardo el archivo. Sino false.
}
function Excel2007(dataSet:TDataSet;nombreHoja:string):boolean; overload;

{
Guarda los datos de grid en el archivo pasado con el formato especificado.
- grid: Grid que conteniendo los datos a guardar en el archivo.
- nombreHoja: Nombre a asignar a la hoja donde se guardar�n los datos.
- archivo: Ruta y nombre del archivo donde se guardar�n los datos.
- formato: Formato con el que se generar� el archivo. Los valores aceptados son
los definidos por Microsoft en XlFileFormat Enumeration:
http://msdn2.microsoft.com/en-us/library/bb241279.aspx
- Result: True si se guardo el archivo. Sino false.
}
function Excel2007(grid: TStringGrid; nombreHoja, archivo:string; formato:integer;
                    MostrarColumnasOcultas:Boolean = False): boolean; overload;

{
Guarda los datos de grid en el archivo pasado con el formato especificado.
- grid: Grid que conteniendo los datos a guardar en el archivo.
- nombreHoja: Nombre a asignar a la hoja donde se guardar�n los datos.
- archivo: Ruta y nombre del archivo donde se guardar�n los datos.
- formato: Formato con el que se generar� el archivo. Los valores aceptados son
los definidos por Microsoft en XlFileFormat Enumeration:
http://msdn2.microsoft.com/en-us/library/bb241279.aspx
- Result: True si se guardo el archivo. Sino false.
}
function Excel2007(DBGrid:TDBGrid; nombreHoja, archivo: string;
  formato:integer; MostrarColumnasOcultas:Boolean = False): boolean; overload;

{
Se muestra un dialogo para que el usuario seleccione una ruta y nombre de
archivo para guardar los datos del data set. Tambi�n en dicho di�logo se puede
seleccionar el formato del archivo a guardar. En caso de que el usuario haya
seleccionado el archivo en el que se guardar�n los datos y en funci�n del
formato seleccionado se guarda el archivo.
- dataSet: Data set conteniendo los datos a guardar en el archivo.
- nombreHoja: Nombre a asignar a la hoja donde se guardar�n los datos.
- archivo: Ruta y nombre del archivo donde se guardar�n los datos.
- formato: Formato con el que se generar� el archivo. Los valores aceptados son
los definidos por Microsoft en XlFileFormat Enumeration:
http://msdn2.microsoft.com/en-us/library/bb241279.aspx
- Result: True si se guardo el archivo. Sino false.
}
function Excel2007(dataSet:TDataSet; nombreHoja, archivo: string;
  formato:integer): boolean; overload;

{
Se obtiene la version de Excel instalada en el equipo.
}
function getExcelVersion:String; overload;

   function SaveAsExcelFile(AGrid: TStringGrid; ASheetName, AFileName: string):
   Boolean; overload;

   function SaveAsExcelFile(AGrid: TStringGrid; ASheetName, AFileName: string;
   Cols, Filas:Boolean): Boolean; overload;
   //function SaveAsExcelFile_3(StringGrid: TStringGrid; ASheetName, FileName: string): boolean;
   procedure GuardarArchivoFormatoExcel(g:TstringGrid;fp,fn,tit:String;s:TSaveDialog;
   cols,filas:boolean);

   function GridToExcel(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean;


   Function DatasetToExcel(qrConsulta:TQuery; strNombreHoja:String; btitulos:Boolean): Boolean; overload;
   Function DatasetToExcel(qrConsulta:TQuery; strNombreHoja:String; btitulos:Boolean; var cArchivo:String): Boolean; overload;
   Function DatasetToTxt(qrConsulta:TQuery; strNombreHoja:String; btitulos:Boolean): Boolean;

   Function GrabaFormatoCSV(qrConsulta:TQuery; strArchivo:String;
                  btitulos:Boolean; FrmAux:TForm; lblAux:TLabel; Separador:String):Boolean; overload;

   function GrabaFormatoCSV(AGrid: TStringGrid; AFileName: string):Boolean;
    overload;

implementation

const
  C_MAX_REGISTROS_EXCEL=65000;

{
Se finaliza la aplicacion.
@param Excel Aplicacion a finalizar.
}
procedure finalizarExcel(Excel:OLEVariant);
begin

  //En caso de haberse instanciado la aplicaci�n...
  if not VarIsEmpty(Excel) then
    // Se sale de la misma
    Excel.Quit;

end;

{
Se obtiene la version de Excel instalada en el equipo.
}
function getExcelVersion(var Excel:OLEVariant):String; overload;
begin

  Result:=Excel.Version;

end;

{
Se obtiene la version de Excel instalada en el equipo.
}
function getExcelVersion:String; overload;
var
  Excel:OLEVariant;
begin

  // Se instancia la aplicaci�n.
  Excel:=CreateOleObject('Excel.Application');

  // Se hace invisible la aplicaci�n.
  Excel.Visible:=False;

  // Se obtiene la version del Excel.
  Result:=getExcelVersion(Excel);

  // Se finaliza la aplicacion.
  finalizarExcel(Excel);

end;

{
Muestra un dialogo para guardar un archivo.
-FilePath: Ruta indicada por el usuario.
-FileName: Nombre del archivo.
-formato: Formato utilizado para la versi�n 2007.
-Result: true en caso de tener que guardar el archivo, sino false.
}
function ExcelSaveDialog(var FilePath:String;var FileName:String;
  var formato:integer):boolean;
var
   saveDialog: TSaveDialog;
   versionExcel:string;
begin

  Result:=False;

  // Crea un di�logo para seleccionar el archivo a guardar.
  SaveDialog :=  TSaveDialog.Create(Application);

  // Se inicializan la ruta y el nombre de archivo.
  FilePath:='';
  saveDialog.FileName := FileName;//FileName:='';

  // Se asigna el directorio de la aplicaci�n al di�logo.
  saveDialog.InitialDir := Application.GetNamePath;

  // Se asigna el t�tulo al di�logo.
  saveDialog.Title := 'Exportar a excel';

  versionExcel:=getExcelVersion;

  // Se asignan los formatos con los que se podr� guardar.

  // En caso de que la version de excel sea mayor a la 11 (version 2003)...
  if versionExcel>'11.0' then
    // ...se da como opcion poder guardar en el nuevo formato *.xlsx.
    saveDialog.Filter := 'Libro de Excel (*.xlsx)|*.xlsx|';

  // Se agregan el resto de los formatos.
  saveDialog.Filter:=saveDialog.Filter+
    'Libro de Excel 97-2003 (*.xls)|*.xls|'+
    'Todos los archivos (*.*)|*.*';

  // Se asigna la extenci�n a agregar por default.
  saveDialog.DefaultExt := 'xls';

  // En caso de que el usuario haya seleccionado un archivo al ejecutar el
  // di�logo...
  if saveDialog.Execute then begin

    // Se obtiene el nombre del archivo seleccionado.
    FileName := ExtractFileName(saveDialog.FileName);

    // Se obtiene la ruta del archivo seleccionado.
    FilePath := ExtractFilePath(saveDialog.FileName);

    // Se obtiene el formato con el que se desea guardar la aplicaci�n.
    if (saveDialog.FilterIndex = 1) and (versionExcel>'11.0') then
      formato:=xlOpenXMLWorkbook
    else
      formato:=xlExcel8;

    // En caso de que exista el archivo...
    if FileExists(saveDialog.FileName) then
    begin

      // ...se pregunta si se desea sobreescribir.
      if MessageDlg(Format('Sobrescribir %s', [saveDialog.FileName]),
        mtConfirmation, mbYesNoCancel, 0) = mrYes then
        Result := true

    end
    else
      Result := true;

  end;

  // Destruye el tSaveDialog
  saveDialog.Free;

end;

{
Se inicia el excel, se agrega un libro y se obtiene la primer hoja de dicho
libro a la cual se le setea el nombre pasado.
@param Excel Referencia a la aplicacion iniciada.
@param Sheet Referencia a la hoja obtenida.
@param nombreHoja Nombre a asignarle a la primer hoja. 
}
procedure iniciarExcel(var Excel,Sheet:OLEVariant;nombreHoja:string);
begin

  // Se instancia la aplicaci�n
  Excel := CreateOleObject('Excel.Application');

  // Se hace invisible la aplicaci�n.
  Excel.Visible := False;

  // Se pide no se muestren las alertas.
  Excel.DisplayAlerts:=False;

  // Se agrega un libro al documento.
  Excel.Workbooks.Add();

  // Se obtiene la primer hoja del libro.
  Sheet := Excel.Workbooks[1].WorkSheets[1];

  // Se asigna un nombre a la hoja.
  Sheet.Name := copy(nombreHoja,1,30);

end;

{
Se crea un formulario de progreso.
}
procedure crearFormularioProgreso(var FrmAux:TForm;var lblAux:Tlabel);
var
   FraAux:TGroupBox;
begin
   // Crea un Form
   FrmAux := TForm.create(Application);
   //FrmAux.Name := 'frmAux';
   FrmAux.BorderIcons := frmAux.BorderIcons - [biSystemMenu, biMinimize, biMaximize, biHelp];
   FrmAux.Position := poScreenCenter;
   FrmAux.BorderStyle := bsToolWindow;
   FrmAux.Width := 448;
   FrmAux.Height := 94;

   // Crea Frame
   FraAux := TGroupBox.create(FrmAux);
   FraAux.parent := FrmAux;
   FraAux.Name := 'FraAux';
   FraAux.Caption := '';
   FraAux.Top := 8;
   FraAux.Left := 8;
   FraAux.Width := 425;
   FraAux.Height := 41;

   // Crea label
   lblAux := TLabel.Create(FraAux);
   lblaux.Font.Size := 12;
   lblaux.Font.Style := [fsBold];
   lblAux.parent := FraAux;
   lblAux.Left := 8;
   lblAux.Top := 16;
   lblAux.Name := 'lblAux';
   lblAux.Caption := '';
   lblAux.Visible := true;

   FrmAux.Caption := 'Exportando a formato excel...';
   FrmAux.Show;

end;

{
Se asignan los datos del grid a la hoja pasada.
@param grid Datos a ser asignados a la hoja pasada.
@param Sheet Hoja a la que se le asignaran los datos del grid.
}
procedure asignarDatosHoja(grid:TStringGrid;Sheet:OLEVariant;
                           MostrarColumnasOcultas:Boolean = False); overload;
var
  fila,k,
  columna,
  cantidadFilas:integer;
  frmProgreso:TForm;
  labelProgreso:TLabel;
begin

  // Se crea el formulario de progreso.
  crearFormularioProgreso(frmProgreso,labelProgreso);

  // Se obtiene la cantidad de filas.
  cantidadFilas:=grid.RowCount;

  // Se pasan los datos del grid a la hoja.
  for fila := 0 to grid.RowCount - 1 do begin
    k := 0;
    for columna := 0 to grid.ColCount - 1 do
    begin
      if (MostrarColumnaSOcultas ) OR (NOT MostrarColumnaSOcultas AND (grid.ColWidths[Columna] > 0) ) then begin
          Sheet.Cells[fila + 1,k + 1] := grid.Cells[columna, fila];
          inc(k);
      end;

      // Se setea la informaci�n de progreso.
      labelProgreso.Caption := 'Exportando fila '+IntTostr(fila+1)+' de '+
        IntToStr(cantidadFilas);

      frmProgreso.Refresh;
      labelProgreso.Refresh;
    end;
  end;
  labelProgreso.Free;
  frmProgreso.Release;

end;

{
Se asignan los datos del dataset pasado a la hoja pasada.
@param dataset Dataset de donde se obtendran los datos.
@param Sheet Hoja a la que se asignaran los datos.
}
procedure asignarDatosHoja(DBGrid:TDBGrid;var Sheet:OLEVariant;
                           MostrarColumnasOcultas:Boolean = False); overload;
var
  i,j,k,
  cantidadFilas:integer;
  dataset:TDataSet;
  frmProgreso:TForm;
  labelProgreso:TLabel;
  field:TField;
begin

  // Se crea el formulario de progreso.
  crearFormularioProgreso(frmProgreso,labelProgreso);

  // Se setean los titulos de las columnas del DBGrid.
  k := 0;
  for i:=0 to DBGrid.FieldCount-1 do
    if (MostrarColumnaSOcultas ) OR (NOT MostrarColumnaSOcultas
       AND DBGrid.Columns[i].Visible) then begin
          Sheet.Cells[1,k+1]:=DBGrid.Columns[i].Title.Caption;
          inc(k);
     end;

  // Se obtiene el dataset del DBGrid.
  dataset:=DBGrid.DataSource.DataSet;

  dataset.DisableControls;

  // Se obtiene la cantidad de filas del dataset
  dataset.Last;
  cantidadFilas:=dataset.RecordCount;

  // Se iteran las filas del dataset.
  dataset.First;
  j:=0;

  while not dataset.Eof do
  begin

    // Se setean los valores de cada celda de la fila.
    k := 0;
    for i:=0 to DBGrid.FieldCount-1 do
    begin
       if (MostrarColumnaSOcultas ) OR (NOT MostrarColumnaSOcultas AND DBGrid.Columns[i].Visible) then Begin
          // Se obtiene el campo
          field:=dataset.Fields.FieldByName(DBGrid.Columns[i].FieldName);

          // En caso de tratarce de una fecha...
          if field is TDateTimeField and not field.isNull then
            // ...se le da el formato correcto.
            Sheet.Cells[j+2,k+1]:=FormatDateTime('yyyy-mm-dd',field.AsDateTime)
          else
            // Sino se interpreta como un string.
            Sheet.Cells[j+2,k+1]:=field.AsString;

          inc(k);
       end;
    end;

    // Se setea la informaci�n de progreso.
    labelProgreso.Caption := 'Exportando fila '+IntTostr(j+1)+' de '+
      IntToStr(cantidadFilas);
    frmProgreso.Refresh;
    labelProgreso.Refresh;

    dataset.Next;
    inc(j);
  end;

  dataset.EnableControls;

  labelProgreso.Free;
  frmProgreso.Release;

end;

{
Se asignan los datos del dataset pasado a la hoja pasada.
@param dataset Dataset de donde se obtendran los datos.
@param Sheet Hoja a la que se asignaran los datos.
}
procedure asignarDatosHoja(dataset:TDataSet;var Sheet:OLEVariant); overload;
var
  i,
  j,
  cantidadFilas:integer;
  frmProgreso:TForm;
  labelProgreso:TLabel;
  field:TField;
begin

  // Se crea el formulario de progreso.
  crearFormularioProgreso(frmProgreso,labelProgreso);

  // Se setean los titulos de las columnas del dataset.
  for i:=0 to dataset.FieldCount-1 do
    Sheet.Cells[1,i+1]:=dataset.Fields.Fields[i].FieldName;

  dataset.DisableControls;

  // Se obtiene la cantidad de filas del dataset
  dataset.Last;
  cantidadFilas:=dataset.RecordCount;

  // Se iteran las filas del dataset.
  dataset.First;
  j:=0;
  while not dataset.Eof do
  begin

    // Se setean los valores de cada celda de la fila.
    for i:=0 to dataset.FieldCount-1 do
    begin

      // Se obtiene el campo
      field:=dataset.Fields.FieldByNumber(i+1);

      // En caso de tratarce de una fecha...
      if field is TDateTimeField and not field.isNull then
        // ...se le da el formato correcto.
        Sheet.Cells[j+2,i+1]:=FormatDateTime('yyyy-mm-dd',field.AsDateTime)
      else
        // Sino se interpreta como un string.
        Sheet.Cells[j+2,i+1]:=field.AsString;

    end;

    // Se setea la informaci�n de progreso.
    labelProgreso.Caption := 'Exportando fila '+IntTostr(j+1)+' de '+
      IntToStr(cantidadFilas);
    frmProgreso.Refresh;
    labelProgreso.Refresh;


    dataset.Next;
    inc(j);
  end;

  dataset.EnableControls;

  labelProgreso.Free;
  frmProgreso.Release;

end;

{
Se guarda el libro para aplicacion pasada.
@param Excel Aplicacion en la cual se guardara el libro tratado.
@return true si pudo guardar el libro con exito, sino false.
}
function guardarLibro(Excel:OLEVariant;archivo:string;formato:integer):boolean;
begin

  Result:=False;

  try

    // Se guarda el documento dependiendo del formato solicitado.
    Excel.Workbooks[1].SaveAs(archivo,FileFormat:=formato);

    // En caso de no haber ocurrido una excepci�n se indica se ha tenido
    // exito.
    Result:=True;

  except

    // En caso de no haber podido guardar con el formato especificado y que
    // dicho formato sea distinto del de Excel 2007...
    if (formato<>xlOpenXMLWorkbook) then
    begin

      // ...se intenta guardar con el formato normal de la aplicaci�n (esto se
      // hace para las PCs que no tienen Excel 2007)
      Excel.Workbooks[1].SaveAs(archivo,FileFormat:=xlWorkbookNormal);

      // En caso de no haber ocurrido una excepci�n se indica se ha tenido
      // exito.
      Result:=True;

    end
    else

      // Sino se indica que no se soporta el formato solicitado.
      MessageDlg('El formato solicitado no est� soportado por la versi�n de '+
        'Excel instalada en su equipo. Por favor seleccione otro formato',
        mtError,[mbOk],0);

  end;

end;

function Excel2007(grid: TStringGrid; nombreHoja, archivo: string; formato:integer;
                   MostrarColumnasOcultas:Boolean = False): boolean; overload;
var
   Excel,
   Sheet:OLEVariant;
begin

  Result := false;

  try

    // Se inicia el Excel.
    iniciarExcel(Excel,Sheet,nombreHoja);

    // Se asignan los datos a la hoja.
    //asignarDatosHoja(grid,Sheet);
    asignarDatosHoja(grid,Sheet, MostrarColumnasOcultas);

    // Se guarda el libro.
    Result:=guardarLibro(Excel,archivo,formato);

  finally

    // Se finaliza la aplicacion.
    finalizarExcel(Excel);

  end;

  if Result then
    MessageDlg('Archivo Excel generado correctamente.', mtInformation,
      [mbOk],0);

end;

{
Se graba a formato TXT el grid pasado.
DBGrid: Grid a ser grabado.
archivo: Archivo donde se guardar� la informaci�n del grid.
}
function GrabaFormatoTXT(DBGrid:TDBGrid;archivo: string):boolean; overload;
var
  frmProgreso:TForm;
  labelProgreso:TLabel;
begin

  // Se crea el formulario de progreso.
  crearFormularioProgreso(frmProgreso,labelProgreso);

  DBGrid.DataSource.DataSet.First;

  DBGrid.DataSource.DataSet.DisableControls;

  Result:=GrabaFormatoCSV(TQuery(DBGrid.DataSource.DataSet),archivo,true,frmProgreso,labelProgreso,#9);

  DBGrid.DataSource.DataSet.EnableControls;

  labelProgreso.Free;
  frmProgreso.Release;

end;

{
Se graba a formato TXT el data set pasado.
dataSet: Data set a ser grabado.
archivo: Archivo donde se guardar� la informaci�n del grid.
}
function GrabaFormatoTXT(dataSet:TDataSet;archivo: string):boolean; overload;
var
  frmProgreso:TForm;
  labelProgreso:TLabel;
begin

  // Se crea el formulario de progreso.
  crearFormularioProgreso(frmProgreso,labelProgreso);

  dataSet.First;

  dataSet.DisableControls;

  Result:=GrabaFormatoCSV(TQuery(dataSet),archivo,true,frmProgreso,
    labelProgreso,#9);

  dataSet.EnableControls;

  labelProgreso.Free;
  frmProgreso.Release;

end;

function Excel2007(DBGrid:TDBGrid; nombreHoja, archivo: string; formato:integer;
                   MostrarColumnasOcultas:Boolean = False): boolean; overload;
var
   Excel,
   Sheet:OLEVariant;
begin

  Result := false;

  try

    // Se inicia el Excel.
    iniciarExcel(Excel,Sheet,nombreHoja);

    // Se asignan los datos a la hoja.
    //asignarDatosHoja(DBGrid,Sheet);
    asignarDatosHoja(DBGrid,Sheet, MostrarColumnasOcultas);

    // Se guarda el libro.
    Result:=guardarLibro(Excel,archivo,formato);

  finally

    // Se finaliza la aplicacion.
    finalizarExcel(Excel);

  end;

  if Result then
    MessageDlg('Archivo generado correctamente.', mtInformation,
      [mbOk],0);

end;

function Excel2007(dataSet:TDataSet; nombreHoja, archivo: string;
                   formato:integer): boolean; overload;
var
   Excel,
   Sheet:OLEVariant;
begin

  Result := false;

  try

    // Se inicia el Excel.
    iniciarExcel(Excel,Sheet,nombreHoja);

    // Se asignan los datos a la hoja.
    asignarDatosHoja(dataSet,Sheet);

    // Se guarda el libro.
    Result:=guardarLibro(Excel,archivo,formato);

  finally

    // Se finaliza la aplicacion.
    finalizarExcel(Excel);

  end;

  if Result then
    MessageDlg('Archivo generado correctamente.', mtInformation,
      [mbOk],0);

end;

{
Valida si existe la ruta pasada. La misma debe terminar con '\'.
}
function existeRuta(ruta:string):boolean;
var
  search_rec:TSearchRec;
begin
  // Se verifica si existe la ruta ingresada.
  Result:=(FindFirst(ruta+'*',faAnyFile,search_rec)=0);
  FindClose(search_rec);
end;

function Excel2007(DBGrid:TDBGrid;nombreHoja:string;
                   MostrarColumnasOcultas:Boolean = False):boolean; overload;
var
  ruta,
  archivo:string;
  exportar:Boolean;
  formato:integer;
  st:String;
begin

  Result:=False;

  exportar:=true;
  ruta:='';
  archivo:='';

  // En caso de tener que exportar y no exista la ruta o no se haya ingresado el
  // archivo.
  while exportar and ( (not existeRuta(ruta)) or (archivo = '')) do
    // Se muestra el dialogo para exportar.
    exportar:=ExcelSaveDialog(ruta,archivo,formato);

  // En caso de tener que exportar...
  if exportar then
  begin

    DBGrid.DataSource.DataSet.Last;

    // En caso que la cantidad de registros sea muy grande...
    If DBGrid.DataSource.DataSet.RecordCount > C_MAX_REGISTROS_EXCEL then
    begin

      // ...se verifica si se desea e�xportar a CSV
      st := 'La cantidad de filas excede el l�mite permitido.';
      st := st + 'Se exportar� solo en formato de texto.' + #13;
      st := st + '�Desea continuar? ';
      if MessageDlg(st,mtConfirmation,[mbYes, mbNo],0)=mrYes then

        //if MessageDlg('Esta operaci�n puede durar varios minutos. �Desea '+
        //  'guardar en formato TXT para disminuir el tiempo de grabaci�n ?',
        //  mtConfirmation,[mbYes, mbNo, mbCancel],0)=mrYes then
          Result:=GrabaFormatoTXT(DBGrid,archivo)
      else
        //Result:=Excel2007(DBGrid,nombreHoja,ruta+archivo,formato,MostrarColumnasOcultas);

    end
    else
      Result:=Excel2007(DBGrid,nombreHoja,ruta+archivo,formato,MostrarColumnasOcultas);
  end;
end;

function Excel2007(dataSet:TDataSet;nombreHoja:string):boolean; overload;
var
  ruta,
  archivo:string;
  exportar:Boolean;
  formato:integer;
  st:String;
begin

  Result:=False;

  exportar:=true;
  ruta:='xxx';
  archivo:=nombreHoja;

  // En caso de tener que exportar y no exista la ruta o no se haya ingresado el
  // archivo.
  while exportar and (not existeRuta(ruta)) and (archivo <> '') do
    // Se muestra el dialogo para exportar.
    exportar:=ExcelSaveDialog(ruta,archivo,formato);

  // En caso de tener que exportar...
  if exportar then
  begin

    dataSet.Last;

    // En caso que la cantidad de registros sea muy grande...
    If dataSet.RecordCount > C_MAX_REGISTROS_EXCEL then
    begin

      // ...se verifica si se desea e�xportar a CSV
      st := 'La cantidad de filas excede el l�mite permitido.';
      st := st + 'Se exportar� solo en formato de texto.' + #13;
      st := st + '�Desea continuar? ';
      if MessageDlg(st,mtConfirmation,[mbYes, mbNo],0)=mrYes then

        //if MessageDlg('Esta operaci�n puede durar varios minutos. �Desea '+
        //  'guardar en formato TXT para disminuir el tiempo de grabaci�n ?',
        //  mtConfirmation,[mbYes, mbNo, mbCancel],0)=mrYes then
          Result:=GrabaFormatoTXT(dataset,archivo)
      else
        //Result:=Excel2007(dataSet,nombreHoja,ruta+archivo,formato);

    end
    else
      Result:=Excel2007(dataSet,nombreHoja,ruta+archivo,formato);

  end;

end;

function Excel2007(grid: TStringGrid; nombreHoja: string;
                    MostrarColumnasOcultas:Boolean = false):boolean; overload;
var
  ruta,
  archivo:string;
  exportar:Boolean;
  formato:integer;
  st:String;
begin

  Result:=False;

  exportar:=true;
  ruta:='xxx';
  archivo:=nombreHoja;

  // En caso de tener que exportar y no exista la ruta o no se haya ingresado el
  // archivo.
  while exportar and (not existeRuta(ruta)) and (archivo <> '') do
    // Se muestra el dialogo para exportar.
    exportar:=ExcelSaveDialog(ruta,archivo,formato);

  // En caso de tener que exportar...
  if exportar then
  begin

    // En caso que la cantidad de registros sea muy grande...
    if grid.RowCount > C_MAX_REGISTROS_EXCEL then
    begin

      // ...se verifica si se desea exportar a CSV
      // ...se verifica si se desea e�xportar a CSV
      st := 'La cantidad de filas excede el l�mite permitido.';
      st := st + 'Se exportar� solo en formato de texto.' + #13;
      st := st + '�Desea continuar? ';
      if MessageDlg(st,mtConfirmation,[mbYes, mbNo],0)=mrYes then

        //if MessageDlg('Esta operaci�n puede durar varios minutos. �Desea '+
        //  'guardar en formato TXT para disminuir el tiempo de grabaci�n ?',
        //  mtConfirmation,[mbYes, mbNo, mbCancel],0)=mrYes then
          Result:=GrabaFormatoCSV(grid,archivo)
      else
        //Result:=Excel2007(grid,nombreHoja,ruta+archivo,formato,MostrarColumnasOcultas);
    end
    else
      Result:=Excel2007(grid,nombreHoja,ruta+archivo,formato,MostrarColumnasOcultas);
  end;

end;

function eliminarTabuladoresYSaltosDeLinea(s:string):string;
var
  partes:TStrings;
  i:integer;
begin

  partes:=TStringList.create;

  try

    if ExtractStrings([#13,#10,#9],[' '],PChar(s),partes) = 0 then
      Result:=s
    else
    begin

      Result:='';

      for i:=0 to partes.Count-1 do
        Result:=Result+' '+partes[i];

    end;

  finally

    partes.Free;

  end;

end;

Function GrabaFormatoCSV(qrConsulta:TQuery; strArchivo:String;
                  btitulos:Boolean; FrmAux:TForm; lblAux:TLabel; Separador:String):Boolean;

const
   // Constantes LOCALES
   ENTER = #13 + #10;
var
   Row, Col :Integer;
   f:textFile;
   StrAux:String;
   i,CantRow:Integer;
   Sigue :boolean;
begin

   //FrmAux.Caption := 'Exportando a formato CSV...';
   //FrmAux.Show;

   FrmAux.Caption := 'Exportando a formato ' + strAux;
   FrmAux.Show;


   sigue := True;
   Screen.Cursor := crHourglass;

   // Verifico si existe el archivo
   If FileExists(strArchivo) then begin
      if MessageDlg('El archivo ' + strArchivo + ' ya existe.' + #13#10 +
         'Desea reemplazarlo ?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
         deletefile(strArchivo);
      end
      else begin
         Sigue := False;
      end;
   end;

   // crea el archivo CSV
   If sigue then begin
      AssignFile(f,strArchivo);
      Rewrite(f);

      strAux := '';

      sigue := False;
      CantRow := qrConsulta.RecordCount;
      Row  := 0;

      //Pone los titulos
      if btitulos then begin
         Row := Row + 1;
         For i := 0 to qrConsulta.FieldCount - 1 do begin
            strAux := strAux + qrConsulta.Fields[i].FieldName;
            if i = qrConsulta.FieldCount - 1 then
               strAux := strAux + ENTER
            else
               //strAux := strAux + ',';
               strAux := strAux + Separador;
         end;
         write(f,StrAux);
      end;

      While not qrConsulta.Eof do begin
         Sigue := True;

         Row := Row + 1;
         if btitulos then
            lblaux.Caption := 'Exportando fila ' + IntTostr(Row) + ' de ' + IntToStr(CantRow + 1)
         else
            lblaux.Caption := 'Exportando fila ' + IntTostr(Row) + ' de ' + IntToStr(CantRow);
         lblaux.Refresh;
         strAux := '';

         for col := 0 to qrConsulta.FieldCount - 1 do
         begin

            strAux := strAux +
              eliminarTabuladoresYSaltosDeLinea(
                Trim(qrConsulta.Fields[Col].AsString));

            if col = qrConsulta.FieldCount - 1 then
               strAux := strAux + ENTER
            else
               //strAux := strAux + ',';
               strAux := strAux + Separador;
         end;

         qrConsulta.Next;
         write(f,StrAux);

      end;

      // Cierra archivo CSV
      CloseFile(f);

      MessageDlg('Archivo generado correctamente.', mtInformation,[mbOk], 0);
   end;

   Screen.Cursor := crDefault;

   Result := Sigue;

end;

Function GrabaFormatoExcel(qrConsulta:TQuery; strNombreHoja, strArchivo:String;
                              btitulos:Boolean; FrmAux:TForm; lblAux:Tlabel): Boolean;
const
   xlWBATWorksheet = -4167;
var
   Row, Col, CantRow: integer;
   XLApp, Sheet: OLEVariant;
begin

   FrmAux.Caption := 'Exportando a formato excel...';
   FrmAux.Show;

   Result := False;
   Screen.Cursor := crHourglass;
   XLApp := CreateOleObject('Excel.Application');
   try
      XLApp.Visible := False;
      XLApp.Workbooks.Add(xlWBatWorkSheet);
      // Marcos
      XLApp.DisplayAlerts:=False;
      Sheet := XLApp.Workbooks[1].WorkSheets[1];
      //Sheet.Name := 'My Sheet Name';
      Sheet.Name := copy(strNombreHoja,1,30);

      CantRow := qrConsulta.RecordCount;

      Row  := 0;

      //Pone los titulos
      if btitulos then begin
         Row := Row + 1;
         For col := 0 to qrConsulta.FieldCount - 1 do
            Sheet.Cells[Row, Col + 1] := qrConsulta.Fields[Col].FieldName
      end;

      While not qrConsulta.Eof do begin
         Row := Row + 1;
         if btitulos then
            lblaux.Caption := 'Exportando fila ' + IntTostr(Row) + ' de ' + IntToStr(CantRow + 1)
         else
            lblaux.Caption := 'Exportando fila ' + IntTostr(Row) + ' de ' + IntToStr(CantRow);

         FrmAux.Refresh;
         lblaux.Refresh;

         for col := 0 to qrConsulta.FieldCount - 1 do begin
            if qrConsulta.Fields[col] is TDateTimeField
            and not qrConsulta.Fields[col].isNull then
               Sheet.Cells[row, col + 1] :=
                  FormatDateTime('yyyy-mm-dd', qrConsulta.Fields[Col].AsDateTime)
            else
               Sheet.Cells[row, col + 1] :=
                  Trim(qrConsulta.Fields[Col].AsString);
         end;
         qrConsulta.Next;
      end;

      try
         XLApp.Workbooks[1].SaveAs(strArchivo);
         Result := true;
      except
         // Error ?
      end;

   finally
      if not VarIsEmpty(XLApp) then begin
         XLApp.DisplayAlerts := False;
         XLApp.Quit;
         XLAPP := Unassigned;
         Sheet := Unassigned;
      end;
   end;

   Screen.Cursor := crDefault;

   MessageDlg('Archivo Excel ' + strArchivo + ' generado correctamente.', mtInformation,[mbOk], 0);
end;

Procedure MostrarDialogoExportar(var FilePath:String; var FileName:String; var Respuesta:boolean);
var
   SaveDialog1: TSaveDialog;
begin

   // Crea un TsaveDialog en tiempos de ejecuci�n
   SaveDialog1 :=  TSaveDialog.Create(Application);

   Respuesta := False;

   // Inicializo
   If trim(FilePath) = '' then
      FilePath := 'C:\';
   if trim(FileName) = '' then
      FileName := ' ';

   // Configura el SaveDialog
   SaveDialog1.InitialDir := Application.GetNamePath;
   SaveDialog1.Title := 'Exportar a excel';
   SaveDialog1.Filter := 'Archivos Excel (*.xls)|*.XLS|Todos los archivos (*.*) | *.*';
   SaveDialog1.DefaultExt := 'xls';
   if SaveDialog1.Execute then begin
      FileName := ExtractFileName(SaveDialog1.FileName);
      FilePath := ExtractFilePath(SaveDialog1.FileName);
      if FileExists(SaveDialog1.FileName) then begin
         if MessageDlg(Format('Sobrescribir %s', [SaveDialog1.FileName]),mtConfirmation, mbYesNoCancel, 0) = mrYes then begin
            Respuesta := true;
         end
         else begin
            Exit;
         end;
      end
      else begin
         Respuesta := true;
      end;
   end;

   // Destruye el tSaveDialog
   SaveDialog1.Free;
   //SaveDialog1 := Nil;

end;

Procedure MostrarDialogoExportartxt(var FilePath:String; var FileName:String; var Respuesta:boolean);
var
   SaveDialog1: TSaveDialog;
begin

   // Crea un TsaveDialog en tiempos de ejecuci�n
   SaveDialog1 :=  TSaveDialog.Create(Application);

   Respuesta := False;

   // Inicializo
   If trim(FilePath) = '' then
      FilePath := 'C:\';
   if trim(FileName) = '' then
      FileName := ' ';

   // Configura el SaveDialog
   SaveDialog1.InitialDir := Application.GetNamePath;
   SaveDialog1.Title := 'Exportar a TXT';
   SaveDialog1.Filter := 'Archivos Txt (*.txt)|*.TXT|Todos los archivos (*.*) | *.*';
   SaveDialog1.DefaultExt := 'xls';
   if SaveDialog1.Execute then begin
      FileName := ExtractFileName(SaveDialog1.FileName);
      FilePath := ExtractFilePath(SaveDialog1.FileName);
      if FileExists(SaveDialog1.FileName) then begin
//         if MessageDlg(Format('Sobrescribir %s', [SaveDialog1.FileName]),mtConfirmation, mbYesNoCancel, 0) = mrYes then begin
            Respuesta := true;
//         end
//         else begin
//            Exit;
//         end;
      end
      else begin
         Respuesta := true;
      end;
   end;

   // Destruye el tSaveDialog
   SaveDialog1.Free;
   //SaveDialog1 := Nil;

end;


Function DatasetToExcel(qrConsulta:TQuery; strNombreHoja:String; btitulos:Boolean): Boolean; overload;
   //-----------------------------------------------------------------------------

var
   st:String;
   Seleccion:Integer;

   FrmAux:TForm;
   lblAux:Tlabel;
   FraAux:TGroupBox;
   FilePath,FileName:String;
   strArchivo:String;
   UsuarioConfirmaExportar:Boolean;
   CantidadDeRegistros:integer;

begin

   Result := False;

   // Crea un Form
   FrmAux := TForm.create(Application);
   FrmAux.Name := 'frmAux';
   FrmAux.BorderIcons := frmAux.BorderIcons - [biSystemMenu, biMinimize, biMaximize, biHelp];
   FrmAux.Position := poScreenCenter;
   FrmAux.BorderStyle := bsToolWindow;
   FrmAux.Width := 448;
   FrmAux.Height := 94;

   // Crea Frame
   FraAux := TGroupBox.create(FrmAux);
   FraAux.parent := FrmAux;
   FraAux.Name := 'FraAux';
   FraAux.Caption := '';
   FraAux.Top := 8;
   FraAux.Left := 8;
   FraAux.Width := 425;
   FraAux.Height := 41;

   // Crea label
   lblAux := TLabel.Create(FraAux);
   lblaux.Font.Size := 12;
   lblaux.Font.Style := [fsBold];
   lblAux.parent := FraAux;
   lblAux.Left := 8;
   lblAux.Top := 16;
   lblAux.Name := 'lblAux';
   lblAux.Caption := '';
   lblAux.Visible := true;

   // se inicializan variables locales
   FilePath := '';
   FileName := '';
   UsuarioConfirmaExportar := false;

   // Se llama a la funcion que muestra el TsaveDialog.
   MostrarDialogoExportar(FilePath,FileName,UsuarioConfirmaExportar);

   if UsuarioConfirmaExportar then begin
      strArchivo := FilePath + FileName;

      try
        qrConsulta.DisableControls;

        // Obtengo la cantidad de registros del Tquery
        qrConsulta.last;
        CantidadDeRegistros := qrConsulta.RecordCount;
        qrConsulta.first;

        If (CantidadDeRegistros > C_MAX_REGISTROS_EXCEL) and (CantidadDeRegistros < 65536) then begin
           //st := 'Esta operaci�n puede durar varios minutos. �Desea guardar en formato TXT, para disminuir el tiempo de grabaci�n ?';
           //Seleccion := MessageDlg(st,mtConfirmation, [mbYes, mbNo, mbCancel], 0);
           // ...se verifica si se desea e�xportar a CSV
           st := 'La cantidad de filas excede el l�mite permitido.';
           st := st + 'Se exportar� solo en formato de texto.' + #13;
           st := st + '�Desea continuar? ';
           Seleccion := MessageDlg(st,mtConfirmation,[mbYes, mbNo],0);

           case seleccion of
              mrYes: begin
                        // Le cambio la extension al archivo
                        strArchivo := stringReplace(strArchivo,'.XLS','.TXT',[rfReplaceAll, rfIgnoreCase]);

                        If GrabaFormatoCSV(qrConsulta, strArchivo, btitulos, FrmAux, lblAux, #9) then begin
                           Result := true;
                        end
                        else begin
                           MessageDlg('Error al generar el archivo CSV.', mtError,[mbOk], 0);
                        end;
                     end;

              mrNo: begin
                       //If GrabaFormatoExcel(qrConsulta, strNombreHoja, strArchivo, btitulos, FrmAux, lblAux) then begin
                       //   Result := true;
                       //end
                       //else begin
                       //   MessageDlg('Error al grabar el archivo excel.', mtInformation,[mbOk], 0);
                       //end;
                    end;
              //mrCancel:MessageDlg('Exportaci�n a excel cancelada por el usuario.', mtInformation,[mbOk], 0);
           end;
        end
        else If CantidadDeRegistros < C_MAX_REGISTROS_EXCEL then begin
           If GrabaFormatoExcel(qrConsulta, strNombreHoja, strArchivo, btitulos, FrmAux, lblAux) then begin
              Result := true;
           end
           else begin
              MessageDlg('Error al grabar el archivo excel.', mtInformation,[mbOk], 0);
           end;
        end
        else If CantidadDeRegistros > 65536 then begin
           st := '�Desea guardar en formato CSV debido a que excede el tama�o de una planilla Excel?';
           Seleccion := MessageDlg(st,mtConfirmation, [mbYes, mbCancel], 0);
           case seleccion of
              mrYes: begin
                        If GrabaFormatoCSV(qrConsulta, strArchivo, btitulos, FrmAux, lblAux, #9) then begin
                           Result := true;
                        end
                        else begin
                           MessageDlg('Error al generar el archivo CSV.', mtError,[mbOk], 0);
                        end;
                     end;
              mrCancel:MessageDlg('Exportaci�n a excel cancelada por el usuario.', mtInformation,[mbOk], 0);
           end;
        end;

      Finally
         qrConsulta.EnableControls;

      end;
   end;

   // Libera controles
   lblaux.Free;
   fraAux.free;
   frmAux.Release;

end;

Function DatasetToExcel(qrConsulta:TQuery; strNombreHoja:String; btitulos:Boolean; var cArchivo:String): Boolean; overload;
   //-----------------------------------------------------------------------------

var
   st:String;
   Seleccion:Integer;

   FrmAux:TForm;
   lblAux:Tlabel;
   FraAux:TGroupBox;
   FilePath,FileName:String;
   strArchivo:String;
   UsuarioConfirmaExportar:Boolean;
   CantidadDeRegistros:integer;

begin

   Result := False;

   // Crea un Form
   FrmAux := TForm.create(Application);
   FrmAux.Name := 'frmAux';
   FrmAux.BorderIcons := frmAux.BorderIcons - [biSystemMenu, biMinimize, biMaximize, biHelp];
   FrmAux.Position := poScreenCenter;
   FrmAux.BorderStyle := bsToolWindow;
   FrmAux.Width := 448;
   FrmAux.Height := 94;

   // Crea Frame
   FraAux := TGroupBox.create(FrmAux);
   FraAux.parent := FrmAux;
   FraAux.Name := 'FraAux';
   FraAux.Caption := '';
   FraAux.Top := 8;
   FraAux.Left := 8;
   FraAux.Width := 425;
   FraAux.Height := 41;

   // Crea label
   lblAux := TLabel.Create(FraAux);
   lblaux.Font.Size := 12;
   lblaux.Font.Style := [fsBold];
   lblAux.parent := FraAux;
   lblAux.Left := 8;
   lblAux.Top := 16;
   lblAux.Name := 'lblAux';
   lblAux.Caption := '';
   lblAux.Visible := true;

   // se inicializan variables locales
   FilePath := '';
   FileName := '';
   UsuarioConfirmaExportar := false;

   // Se llama a la funcion que muestra el TsaveDialog.
   mostrarDialogoExportar(FilePath,FileName,UsuarioConfirmaExportar);

   if UsuarioConfirmaExportar then begin
      strArchivo := FilePath + FileName;
      cArchivo   := FileName;

      // Obtengo la cantidad de registros del Tquery
      qrConsulta.last;
      CantidadDeRegistros := qrConsulta.RecordCount;
      qrConsulta.first;

      If (CantidadDeRegistros > C_MAX_REGISTROS_EXCEL) and (CantidadDeRegistros < 65536) then begin
         //st := 'Esta operaci�n puede durar varios minutos. �Desea guardar en formato TXT, para disminuir el tiempo de grabaci�n ?';
         //Seleccion := MessageDlg(st,mtConfirmation, [mbYes, mbNo, mbCancel], 0);
         st := 'La cantidad de filas excede el l�mite permitido.';
         st := st + 'Se exportar� solo en formato de texto.' + #13;
         st := st + '�Desea continuar? ';
         Seleccion := MessageDlg(st,mtConfirmation,[mbYes, mbNo],0);
         case seleccion of
            mrYes: begin
                      // Le cambio la extension al archivo
                      strArchivo := stringReplace(strArchivo,'.XLS','.TXT',[rfReplaceAll, rfIgnoreCase]);

                      If GrabaFormatoCSV(qrConsulta, strArchivo, btitulos, FrmAux, lblAux, #9) then begin
                         Result := true;

                         cArchivo := stringReplace(cArchivo,'.XLS','.TXT',[rfReplaceAll, rfIgnoreCase]);

                      end
                      else begin
                         MessageDlg('Error al generar el archivo CSV.', mtError,[mbOk], 0);
                      end;
                   end;

            mrNo: begin
                     //If GrabaFormatoExcel(qrConsulta, strNombreHoja, strArchivo, btitulos, FrmAux, lblAux) then begin
                     //   Result := true;
                     //
                     //   cArchivo := stringReplace(cArchivo,'.TXT','.XLS',[rfReplaceAll, rfIgnoreCase]);
                     //
                     //   if uppercase(trim(right(cArchivo,4))) <> '.XLS' then
                     //      cArchivo := trim(cArchivo) + '.xls';
                     //end
                     //else begin
                     //   MessageDlg('Error al grabar el archivo excel.', mtInformation,[mbOk], 0);
                     //end;
                  end;
            //mrCancel:MessageDlg('Exportaci�n a excel cancelada por el usuario.', mtInformation,[mbOk], 0);
         end;
      end
      else If CantidadDeRegistros < C_MAX_REGISTROS_EXCEL then begin
         If GrabaFormatoExcel(qrConsulta, strNombreHoja, strArchivo, btitulos, FrmAux, lblAux) then begin
            Result := true;

            cArchivo := stringReplace(cArchivo,'.TXT','.XLS',[rfReplaceAll, rfIgnoreCase]);

         end
         else begin
            MessageDlg('Error al grabar el archivo excel.', mtInformation,[mbOk], 0);
         end;
      end
      else If CantidadDeRegistros > 65536 then begin
         st := '�Desea guardar en formato CSV debido a que excede el tama�o de una planilla Excel?';
         Seleccion := MessageDlg(st,mtConfirmation, [mbYes, mbCancel], 0);
         case seleccion of
            mrYes: begin
                      // Le cambio la extension al archivo
                      strArchivo := stringReplace(strArchivo,'.XLS','.TXT',[rfReplaceAll, rfIgnoreCase]);

                      If GrabaFormatoCSV(qrConsulta, strArchivo, btitulos, FrmAux, lblAux, #9) then begin
                         Result := true;

                         cArchivo := stringReplace(cArchivo,'.XLS','.TXT',[rfReplaceAll, rfIgnoreCase]);

                      end
                      else begin
                         MessageDlg('Error al generar el archivo CSV.', mtError,[mbOk], 0);
                      end;
                   end;
            mrCancel:MessageDlg('Exportaci�n a excel cancelada por el usuario.', mtInformation,[mbOk], 0);
         end;
      end;
   end;

   // Libera controles
   lblaux.Free;
   fraAux.free;
   frmAux.Release;

end;

function RefToCell(ARow, ACol: Integer): string;
var
   str:String;
begin
  str := '';
  str := Chr(Ord('A') + ACol - 1) + IntToStr(ARow);
  //str := 'A' + IntToStr(ACol - 1) + IntToStr(ARow);

  RefToCell := str;
end;

function SaveAsExcelFile_OLD(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean; overload;
const
  xlWBATWorksheet = -4167;
var
  //Row, Col: Integer;
  //GridPrevFile: string;
  XLApp, Sheet, Data: OLEVariant;
  i, j: Integer;
begin
  // Prepare Data
  Data := VarArrayCreate([1, AGrid.RowCount, 1, AGrid.ColCount], varVariant);
  for i := 0 to AGrid.ColCount - 1 do
    for j := 0 to AGrid.RowCount - 1 do
      Data[j + 1, i + 1] := AGrid.Cells[i, j];

  // Create Excel-OLE Object
  Result := False;
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Marcos
    XLApp.DisplayAlerts:=False;

    // Add new Workbook
    XLApp.Workbooks.Add(xlWBatWorkSheet);
    Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet.Name := copy(ASheetName,1,30);

    // Fill up the sheet
    Sheet.Range[RefToCell(1, 1), RefToCell(AGrid.RowCount,AGrid.ColCount)].Value := Data;

    // Save Excel Worksheet
    try
      XLApp.Workbooks[1].SaveAs(AFileName);
      Result := True;
    except
      // Error ?
    end;
  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then begin
       XLApp.DisplayAlerts := False;
       XLApp.Quit;
       XLAPP := Unassigned;
       Sheet := Unassigned;
    end;
  end;
end;

function SaveAsExcelFile(AGrid: TStringGrid; ASheetName, AFileName: string;
Cols, Filas:Boolean): Boolean; overload;
const
  xlWBATWorksheet = -4167;
var
  //Row, Col: Integer;
  //GridPrevFile: string;
  XLApp, Sheet, Data: OLEVariant;
  i, j, i0, j0: Integer;
begin
  // Prepare Data
  Data := VarArrayCreate([1, AGrid.RowCount, 1, AGrid.ColCount], varVariant);
  i0 := 0;
  for i := 0 to AGrid.ColCount - 1 do
    if Cols or (not Cols and (AGrid.ColWidths[i] > 0)) then
    begin
       j0 := 0;
       for j := 0 to AGrid.RowCount - 1 do
          if Filas or (not Filas and (AGrid.RowHeights[j] > 0)) then
          begin
             Data[j0 + 1, i0 + 1] := AGrid.Cells[i, j];
             j0 := j0 + 1;
          end;
       i0 := i0 + 1;
    end;
       // Create Excel-OLE Object
  Result := False;
  XLApp := CreateOleObject('Excel.Application');
  try
     // Hide Excel
     XLApp.Visible := False;
     // Marcos
     XLApp.DisplayAlerts:=False;
     // Add new Workbook
     XLApp.Workbooks.Add(xlWBatWorkSheet);
     Sheet := XLApp.Workbooks[1].WorkSheets[1];
     Sheet.Name := ASheetName;
     // Fill up the sheet
     Sheet.Range[RefToCell(1, 1), RefToCell(AGrid.RowCount,AGrid.ColCount)].Value := Data;

     // Save Excel Worksheet
     try
        XLApp.Workbooks[1].SaveAs(AFileName);
        Result := True;
     except
     // Error ?
     end;
  finally
  // Quit Excel
     if not VarIsEmpty(XLApp) then
     begin
        XLApp.DisplayAlerts := False;
        XLApp.Quit;
        XLAPP := Unassigned;
        Sheet := Unassigned;
     end;
  end;
end;

function SaveAsExcelFile(AGrid: TStringGrid; ASheetName, AFileName: string): boolean; overload;
const
   xlWBATWorksheet = -4167;
var
   Row, Col: integer;
   XLApp, Sheet: OLEVariant;
begin
   Result := false;
   XLApp := CreateOleObject('Excel.Application');
   try
      XLApp.Visible := False;
      XLApp.Workbooks.Add(xlWBatWorkSheet);
      // Marcos
      XLApp.DisplayAlerts:=False;
      Sheet := XLApp.Workbooks[1].WorkSheets[1];
      //Sheet.Name := 'My Sheet Name';
      Sheet.Name := copy(ASheetName,1,30);


      for col := 0 to AGrid.ColCount - 1 do
         for row := 0 to AGrid.RowCount - 1 do
            Sheet.Cells[row + 1,col + 1] := AGrid.Cells[col, row];
            try
               XLApp.Workbooks[1].SaveAs(AFileName);
               Result := True;
            except
               // Error ?
            end;
   finally
      if not VarIsEmpty(XLApp) then begin
         XLApp.DisplayAlerts := False;
         XLApp.Quit;
         XLAPP := Unassigned;
         Sheet := Unassigned;
      end;
   end;
end;


function GrabaFormatoCSV(AGrid: TStringGrid; AFileName: string):Boolean;
const
   // Constantes LOCALES
   ENTER = #13 + #10;
var
   f:textFile;
   StrAux:String;
   i,j,CantRow:Integer;
   Sigue :boolean;

   FrmAux:TForm;
   lblAux:Tlabel;
   FraAux:TGroupBox;

begin

   // Crea un Form
   FrmAux := TForm.create(Application);
   FrmAux.Name := 'frmAux';
   FrmAux.Caption := 'Exportando a formato ';
   FrmAux.BorderIcons := frmAux.BorderIcons - [biSystemMenu, biMinimize, biMaximize, biHelp];
   FrmAux.Position := poScreenCenter;
   FrmAux.BorderStyle := bsToolWindow;
   FrmAux.Width := 448;
   FrmAux.Height := 94;

   // Crea Frame
   FraAux := TGroupBox.create(FrmAux);
   FraAux.parent := FrmAux;
   FraAux.Name := 'FraAux';
   FraAux.Caption := '';
   FraAux.Top := 8;
   FraAux.Left := 8;
   FraAux.Width := 425;
   FraAux.Height := 41;

   // Crea label
   lblAux := TLabel.Create(FraAux);
   lblaux.Font.Size := 12;
   lblaux.Font.Style := [fsBold];
   lblAux.parent := FraAux;
   lblAux.Left := 8;
   lblAux.Top := 16;
   lblAux.Name := 'lblAux';
   lblAux.Caption := '';
   lblAux.Visible := true;

   frmaux.Show;

   sigue := True;
   Screen.Cursor := crHourglass;

   // Verifico si existe el archivo
   If FileExists(Afilename) then begin
      if MessageDlg('El archivo ' + AfileName + ' ya existe.' + #13#10 +
         'Desea reemplazarlo ?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
         deletefile(AfileName);
      end
      else begin
         Sigue := False;
      end;
   end;

   // crea el archivo CSV
   If sigue then begin
      AssignFile(f,AFileName);
      Rewrite(f);

      sigue := False;
      CantRow := Agrid.RowCount - 1;
      for i := 0 to Agrid.Rowcount - 1 do begin
         lblaux.Caption := 'Exportando fila ' + IntTostr(i) + ' de ' + IntToStr(CantRow);
         lblaux.Refresh;
         strAux := '';
         for j := Agrid.FixedCols to Agrid.Colcount - 1 do begin
            If j = Agrid.Colcount - 1 then begin
               strAux := strAux + Agrid.Cells[j,i] + ENTER;
            end
            else begin
               strAux := strAux + Agrid.Cells[j,i] + #9;
            end;
         end;
         // Imprime
         write(f,StrAux);
         Sigue := True;
      end;

      // Cierra archivo CSV
      CloseFile(f);
   end;

   Result := Sigue;

   MessageDlg('Archivo CSV generado correctamente.', mtInformation,[mbOk], 0);

   // Libera controles
   lblaux.Free;
   fraAux.free;
   frmAux.Release;

   Screen.Cursor := crDefault;

end;  // GrabaFormatoCSV

function GridToExcel(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean;

   // -------------------------------------------------------------------------
   function GrabaFormatoExcel(AGrid: TStringGrid; ASheetName,
             AFileName: string):Boolean;
   const
      xlWBATWorksheet = -4167;
   var
      Row, Col, CantRow: integer;
      XLApp, Sheet: OLEVariant;
      FrmAux:TForm;
      lblAux:Tlabel;
      FraAux:TGroupBox;
   begin
      Result := False;
      Screen.Cursor := crHourglass;
         XLApp := CreateOleObject('Excel.Application');
      try
         XLApp.Visible := False;
         XLApp.Workbooks.Add(xlWBatWorkSheet);
         // Marcos
         XLApp.DisplayAlerts:=False;
         Sheet := XLApp.Workbooks[1].WorkSheets[1];
         //Sheet.Name := 'My Sheet Name';
         Sheet.Name := copy(ASheetName,1,30);

         // Crea un Form
         FrmAux := TForm.create(Application);
         FrmAux.Name := 'frmAux';
         FrmAux.Caption := 'Exportando a formato excel...';
         FrmAux.BorderIcons := frmAux.BorderIcons - [biSystemMenu, biMinimize, biMaximize, biHelp];
         FrmAux.Position := poScreenCenter;
         FrmAux.BorderStyle := bsToolWindow;
         FrmAux.Width := 448;
         FrmAux.Height := 94;

         // Crea Frame
         FraAux := TGroupBox.create(FrmAux);
         FraAux.parent := FrmAux;
         FraAux.Name := 'FraAux';
         FraAux.Caption := '';
         FraAux.Top := 8;
         FraAux.Left := 8;
         FraAux.Width := 425;
         FraAux.Height := 41;

         // Crea label
         lblAux := TLabel.Create(FraAux);
         lblaux.Font.Size := 12;
         lblaux.Font.Style := [fsBold];
         lblAux.parent := FraAux;
         lblAux.Left := 8;
         lblAux.Top := 16;
         lblAux.Name := 'lblAux';
         lblAux.Caption := '';
         lblAux.Visible := true;

         frmaux.Show;

         CantRow := Agrid.RowCount - 1;
         for row := 0 to AGrid.RowCount - 1 do begin
            lblaux.Caption := 'Exportando fila ' + IntTostr(Row) + ' de ' + IntToStr(CantRow);
            FraAux.Refresh;

            for col := 0 to AGrid.ColCount - 1 do begin
               if (AGrid.ColWidths[col] > 0) then
                 Sheet.Cells[row + 1,col + 1] := AGrid.Cells[col, row];
            end;
         end;

         try
            XLApp.Workbooks[1].SaveAs(AFileName);
            Result := true;
         except
            // Error ?
         end;

      finally
         if not VarIsEmpty(XLApp) then begin
            XLApp.DisplayAlerts := False;
            XLApp.Quit;
            XLAPP := Unassigned;
            Sheet := Unassigned;
         end;
      end;

      Screen.Cursor := crDefault;

      MessageDlg('Archivo Excel ' + aFileName + ' generado correctamente.', mtInformation,[mbOk], 0);

      // Libera controles
      lblaux.Free;
      fraAux.free;
      frmAux.Release;
   end; // GrabaFormatoExcel
  // -------------------------------------------------------------------------

var
   st:String;
   Seleccion:Integer;

begin
   Result := False;

   If (AGrid.RowCount > C_MAX_REGISTROS_EXCEL) and (AGrid.RowCount < 65536) then begin
      //st := 'Esta operaci�n puede durar varios minutos. �Desea guardar en formato TXT, para disminuir el tiempo de grabaci�n ?';
      //Seleccion := MessageDlg(st,mtConfirmation, [mbYes, mbNo, mbCancel], 0);
      st := 'La cantidad de filas excede el l�mite permitido.';
      st := st + 'Se exportar� solo en formato de texto.' + #13;
      st := st + '�Desea continuar? ';
      Seleccion := MessageDlg(st,mtConfirmation,[mbYes, mbNo],0);
      case seleccion of
         mrYes: begin
                   If GrabaFormatoCSV(AGrid, AFileName) then begin
                      Result := true;
                   end
                   else begin
                      MessageDlg('Error al generar el archivo CSV.', mtError,[mbOk], 0);
                   end;
                end;
         mrNo:  begin
                   //If GrabaFormatoExcel(AGrid, ASheetName, AFileName) then begin
                   //   Result := true;
                   //end
                   //else begin
                   //   MessageDlg('Error al grabar el archivo excel.', mtInformation,[mbOk], 0);
                   //end;
                end;
         //mrCancel:MessageDlg('Exportaci�n a excel cancelada por el usuario.', mtInformation,[mbOk], 0);
      end;
   end
   else If (AGrid.RowCount < C_MAX_REGISTROS_EXCEL) then begin
      If GrabaFormatoExcel(AGrid, ASheetName, AFileName) then begin
         Result := true;
      end
      else begin
         MessageDlg('Error al grabar el archivo excel.', mtInformation,[mbOk], 0);
      end;
   end
   else If (AGrid.RowCount > 65536) then begin
      st := '�Desea guardar en formato CSV debido a que excede el tama�o de una planilla Excel?';
      Seleccion := MessageDlg(st,mtConfirmation, [mbYes, mbCancel], 0);
      case seleccion of
         mrYes: begin
                   If GrabaFormatoCSV(AGrid, AFileName) then begin
                      Result := true;
                   end
                   else begin
                      MessageDlg('Error al generar el archivo CSV.', mtError,[mbOk], 0);
                   end;
                end;
         mrCancel:MessageDlg('Exportaci�n a excel cancelada por el usuario.', mtInformation,[mbOk], 0);
      end;
   end;
end;

//g: Grilla que contiene los datos
//s: Caja de di�logo
//fp: Directorio. Por defecto, c:\
//fn: Nombre del archivo. Por defecto, 'Reporte'
//tit:T�tulo de la caja de di�logo. Por defecto, 'Guardar en formato Excel'
//Cols: True: Todas las columnas. False: Solo las visibles.
//Filas: True: Todas las filas; False: S�lo las visibles.
procedure GuardarArchivoFormatoExcel(g:TstringGrid;fp,fn,tit:String;s:TSaveDialog;
cols,filas:boolean);
var
   sn:String;
begin
   if fp = '' then fp := 'C:\';
   if fn = '' then fn := 'Reporte';
   if tit = '' then tit := 'Guardar en formato Excel';
   // Configura el SaveDialog
   s.InitialDir := Application.GetNamePath;
   s.Title := tit;
   s.Filter := 'Archivos Excel (*.xls)|*.XLS|Todos los archivos (*.*) | *.*';
   s.DefaultExt := 'xls';
   s.FileName := fn;
   if s.Execute then begin
      fn := ExtractFileName(s.FileName);
      fp := ExtractFilePath(s.FileName);
      if FileExists(s.FileName) then begin
         if MessageDlg(Format('Sobrescribir %s', [s.FileName]),mtConfirmation, mbYesNoCancel, 0) = mrYes then begin
            if SaveAsExcelFile(g, sn, s.FileName,cols,filas) then
               MessageDlg('Grabaci�n exitosa.', mtInformation,[mbOk], 0)
            else
               MessageDlg('Grabaci�n cancelada.', mtInformation,[mbOk], 0)
         end
         else begin
            Exit;
         end;
      end
      else begin
         if SaveAsExcelFile(g, sn, s.FileName,cols,filas) then
            MessageDlg('Grabaci�n exitosa.', mtInformation,[mbOk], 0)
         else
            MessageDlg('Grabaci�n cancelada.', mtInformation,[mbOk], 0)
      end;
   end;
end;

procedure XlsWriteCellLabel(XlsStream: TStream; const ACol, ARow: Word;
  const AValue: string);
var
  L: Word;
const
  {$J+}
  CXlsLabel: array[0..5] of Word = ($204, 0, 0, 0, 0, 0);
  {$J-}
begin
  L := Length(AValue);
  CXlsLabel[1] := 8 + L;
  CXlsLabel[2] := ARow;
  CXlsLabel[3] := ACol;
  CXlsLabel[5] := L;
  XlsStream.WriteBuffer(CXlsLabel, SizeOf(CXlsLabel));
  XlsStream.WriteBuffer(Pointer(AValue)^, L);
end;

function SaveAsExcelFile_2(AGrid: TStringGrid; AFileName: string): Boolean;
const
  {$J+} CXlsBof: array[0..5] of Word = ($809, 8, 00, $10, 0, 0); {$J-}
  CXlsEof: array[0..1] of Word = ($0A, 00);
var
  FStream: TFileStream;
  I, J: Integer;
begin
  FStream := TFileStream.Create(PChar(AFileName), fmCreate or fmOpenWrite);
  try
    CXlsBof[4] := 0;
    FStream.WriteBuffer(CXlsBof, SizeOf(CXlsBof));
    for i := 0 to AGrid.ColCount - 1 do
      for j := 0 to AGrid.RowCount - 1 do
        XlsWriteCellLabel(FStream, I, J, AGrid.cells[i, j]);
    FStream.WriteBuffer(CXlsEof, SizeOf(CXlsEof));
    Result := True;
  finally
    FStream.Free;
  end;
end;

//----------------------------------------------------------------------------------------------




Function DatasetToTxt(qrConsulta:TQuery; strNombreHoja:String; btitulos:Boolean): Boolean;
   //-----------------------------------------------------------------------------

var
   //st:String;
   //Seleccion:Integer;

   FrmAux:TForm;
   lblAux:Tlabel;
   FraAux:TGroupBox;
   FilePath,FileName:String;
   strArchivo:String;
   UsuarioConfirmaExportar:Boolean;
   //CantidadDeRegistros:integer;

begin

   Result := False;

   // Crea un Form
   FrmAux := TForm.create(Application);
   FrmAux.Name := 'frmAux';
   FrmAux.BorderIcons := frmAux.BorderIcons - [biSystemMenu, biMinimize, biMaximize, biHelp];
   FrmAux.Position := poScreenCenter;
   FrmAux.BorderStyle := bsToolWindow;
   FrmAux.Width := 448;
   FrmAux.Height := 94;

   // Crea Frame
   FraAux := TGroupBox.create(FrmAux);
   FraAux.parent := FrmAux;
   FraAux.Name := 'FraAux';
   FraAux.Caption := '';
   FraAux.Top := 8;
   FraAux.Left := 8;
   FraAux.Width := 425;
   FraAux.Height := 41;

   // Crea label
   lblAux := TLabel.Create(FraAux);
   lblaux.Font.Size := 12;
   lblaux.Font.Style := [fsBold];
   lblAux.parent := FraAux;
   lblAux.Left := 8;
   lblAux.Top := 16;
   lblAux.Name := 'lblAux';
   lblAux.Caption := '';
   lblAux.Visible := true;

   // se inicializan variables locales
   FilePath := '';
   FileName := '';
   UsuarioConfirmaExportar := false;

   // Se llama a la funcion que muestra el TsaveDialog.
   mostrarDialogoExportartxt(FilePath,FileName,UsuarioConfirmaExportar);

   if UsuarioConfirmaExportar then begin
      strArchivo := FilePath + FileName;

      try
        qrConsulta.DisableControls;

         If GrabaFormatoCSV(qrConsulta, strArchivo, btitulos, FrmAux, lblAux, #9) then begin
            Result := true;
         end
         else begin
            MessageDlg('Error al generar el archivo TXT.', mtError,[mbOk], 0);
         end;
      Finally
         qrConsulta.EnableControls;
      end;
   end;

   // Libera controles
   lblaux.Free;
   fraAux.free;
   frmAux.Release;

end;
//------------------------------------------------------------------------------

class function TExcel.crearArchivoExcel(archivo:string):TExcel;
var
   saveDialog: TSaveDialog;
   versionExcel:string;
   formato:integer;
begin

  Result:=nil;

  // Crea un di�logo para seleccionar el archivo a guardar.
  SaveDialog :=  TSaveDialog.Create(Application);

  saveDialog.InitialDir := ExtractFilePath(archivo);
  saveDialog.FileName:=ExtractFileName(archivo);

  // Se asigna el t�tulo al di�logo.
  saveDialog.Title := 'Exportar a excel';

  versionExcel:=getExcelVersion;

  // Se asignan los formatos con los que se podr� guardar.

  // En caso de que la version de excel sea mayor a la 11 (version 2003)...
  if versionExcel>'11.0' then
    // ...se da como opcion poder guardar en el nuevo formato *.xlsx.
    saveDialog.Filter := 'Libro de Excel (*.xlsx)|*.xlsx|';

  // Se agregan el resto de los formatos.
  saveDialog.Filter:=saveDialog.Filter+
    'Libro de Excel 97-2003 (*.xls)|*.xls|'+
    'Todos los archivos (*.*)|*.*';

  // Se asigna la extenci�n a agregar por default.
  saveDialog.DefaultExt := 'xls';

  // En caso de que el usuario haya seleccionado un archivo al ejecutar el
  // di�logo...
  if saveDialog.Execute then begin

    // Se obtiene el formato con el que se desea guardar la aplicaci�n.
    if (saveDialog.FilterIndex = 1) and (versionExcel>'11.0') then
      formato:=xlOpenXMLWorkbook
    else
      formato:=xlExcel8;

    // En caso de que exista el archivo...
    if FileExists(saveDialog.FileName) then
    begin

      // ...se pregunta si se desea sobreescribir.
      if MessageDlg(Format('Sobrescribir %s', [saveDialog.FileName]),
        mtConfirmation, mbYesNoCancel, 0) = mrYes then
        // Se obtiene el nombre del archivo seleccionado.
        Result:=TExcel.create(saveDialog.FileName,formato);

    end
    else
      // Se obtiene el nombre del archivo seleccionado.
      Result:=TExcel.create(saveDialog.FileName,formato);

  end
  else
    Result:=nil;

  // Destruye el tSaveDialog
  saveDialog.Free;

end;

constructor TExcel.create(archivo:string;formato:integer);
begin

  // Se instancia la aplicaci�n
  Self.aplicacion := CreateOleObject('Excel.Application');

  // Se hace invisible la aplicaci�n.
  Self.aplicacion.Visible := False;

  // Se pide no se muestren las alertas.
  Self.aplicacion.DisplayAlerts:=False;

  // Se agrega un libro al documento.
  Self.aplicacion.Workbooks.Add();

  Self.cantidadHojasAgregadas:=0;

  Self.archivo:=archivo;

  Self.formato:=formato;

end;

constructor TExcel.Create;
begin
end;

class function TExcel.createFromFile(filename:string):TExcel;
begin

  Result:=TExcel.Create;

  // Se instancia la aplicaci�n
  Result.aplicacion:=CreateOleObject('Excel.Application');

  // Se hace invisible la aplicaci�n.
  Result.aplicacion.Visible := False;

  // Se pide no se muestren las alertas.
  Result.aplicacion.DisplayAlerts:=False;

  Result.aplicacion.Workbooks.Open(filename);

end;

function TExcel.getActiveSheet:Variant;
begin
  Result:=Self.aplicacion.ActiveSheet;
end;

function TExcel.getMaxRows:Integer;
begin
   Result:=Self.aplicacion.rows.currentregion.rows.count;
end;

class function TExcel.getNumeroColumna(columna:string):integer;
var
  primeraLetra,
  segundaLetra:char;
  numeroPrimeraLetra,
  numeroSegundaLetra:integer;
  tieneSegundaLetra:boolean;
begin

  if columna='' then
    raise Exception.Create('La columna debe estar entre la A y la IV');

  primeraLetra:=columna[1];

  tieneSegundaLetra:=(Length(columna)=2);

  if (not tieneSegundaLetra) then
  begin
    if (primeraLetra<'A')or(primeraLetra>'Z') then
      raise Exception.Create('La columna debe estar entre la A y la IV')
  end
  else
  begin

    if (primeraLetra<'A')or(primeraLetra>'I') then
      raise Exception.Create('La columna debe estar entre la A y la IV');

    segundaLetra:=columna[2];

    if primeraLetra<'I' then
    begin

      if (segundaLetra<'A')or(segundaLetra>'V') then
        raise Exception.Create('La columna debe estar entre la A y la IV')

    end
    else if (segundaLetra<'A')or(segundaLetra>'Z') then
      raise Exception.Create('La columna debe estar entre la A y la IV')

  end;


  numeroPrimeraLetra:=ord(primeraLetra)-65;
  if (numeroprimeraletra  = 0) and (tieneSegundaLetra) then
     numeroPrimeraLetra:=numeroPrimeraLetra+1;

  if not tieneSegundaLetra then
    Result:=numeroPrimeraLetra+1
  else
  begin
    numeroSegundaLetra:=ord(segundaLetra)-65;
    Result:=(numeroPrimeraLetra*26)+numeroSegundaLetra+1;
  end;

end;

class function TExcel.getLetraColumna(numeroColumna:integer):string;
begin
  if numerocolumna >= 26 then
     Result:='A' + chr(numeroColumna-26+64)
  else
     Result:=chr(numeroColumna+64);

end;

procedure TExcel.save;
begin
  Self.aplicacion.ActiveWorkbook.save;
end;

procedure TExcel.quit;
begin
  Self.aplicacion.quit;
end;

function TExcel.getStringFromCell(row,column:integer):string;
begin

  Result:=Self.getActiveSheet.cells[row,column].value;

end;

procedure TExcel.setValueToCell(row,column:integer;value:Variant);
begin
  Self.getActiveSheet.cells[row,column]:=value;
end;

procedure TExcel.agregarHoja(datos:TDataSet;nombreHoja:string);
var
  hoja:OleVariant;
begin

  inc(Self.cantidadHojasAgregadas);

  // Se obtiene la primer hoja del libro.

  if Self.aplicacion.Workbooks[1].WorkSheets.Count<Self.cantidadHojasAgregadas
    then
    Self.aplicacion.Workbooks[1].WorkSheets.Add(NULL,
      Self.aplicacion.Workbooks[1].WorkSheets[
        Self.aplicacion.Workbooks[1].WorkSheets.Count]);

  hoja:=Self.aplicacion.Workbooks[1].WorkSheets[Self.cantidadHojasAgregadas];

  // Se asigna un nombre a la hoja.
  hoja.Name := copy(nombreHoja,1,30);

  asignarDatosHoja(datos,hoja);

end;

procedure TExcel.guardar;
begin

  guardarLibro(Self.aplicacion,Self.archivo,Self.formato);

end;

procedure TExcel.Free;
begin
  finalizarExcel(Self.aplicacion);
  inherited;
end;

function TExcel.getArchivo:string;
begin
  Result:=Self.archivo;
end;

{
initialization

  // Se recupera la cantidad de l�neas a partir de la cual se sugiere guardar
  // los archivos Excel en formato texto.
  C_MAX_REGISTROS_EXCEL:=getMaxRegistrosExcel;
}

end.
