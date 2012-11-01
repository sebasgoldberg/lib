unit SG_BlobFieldsManager;

interface

{
--------------------------------------------------------------------------------
example of usage (TMyBlobFieldsManager inherit from TSG_BlobFieldsManager and
implements the methods readFields and saveFields):

  // save files to fields.
  myBlobFieldsManager:=TMyBlobFieldsManager.Create(table);
  myBlobFieldsManager.setKey('id1,id2',[1,'value 2']);
  myBlobFieldsManager.saveFieldsTemplate;

  // or save fields to files.
  myBlobFieldsManager:=TMyBlobFieldsManager.Create(table);
  myBlobFieldsManager.setKey('id1,id2',[1,'value 2']);
  myBlobFieldsManager.readFieldsTemplate;

--------------------------------------------------------------------------------
}


uses
  dbtables;

type
  TSG_BlobFieldsManager=class
    private

      keyFields:string;
      keyValues:Variant;

    protected
      table:TTable;

      function goToKey:boolean;
      procedure fileToField(fieldname,fileName:string);
      procedure fieldToFile(fieldname,archivo:string);
      procedure setKey(keyFields:string;keyValues:Variant);

      procedure readFields; virtual;
      procedure saveFields; virtual;

    public

      constructor Create(tableName:string;keyFields:string;keyValues:Variant);
      procedure readFieldsTemplate;
      procedure saveFieldsTemplate;

    end;

implementation

uses
  db,
  unitConexion,
  Sysutils;

constructor TSG_BlobFieldsManager.Create(tableName:string;keyFields:string;
  keyValues:Variant);
begin
  Self.table:=TTable.Create(nil);
  Self.table.DatabaseName:=GBL_DataBase.DatabaseName;
  Self.table.TableName:=tableName;
  Self.setKey(keyFields,keyValues);
end;

procedure TSG_BlobFieldsManager.setKey(keyFields:string;keyValues:Variant);
begin
  Self.keyFields:=keyFields;
  Self.keyValues:=keyValues;
end;

function TSG_BlobFieldsManager.goToKey:boolean;
begin
  Self.table.open;
  Result:=Self.Table.Locate(Self.keyFields,Self.keyValues,[]);
end;

procedure TSG_BlobFieldsManager.fileToField(fieldName,fileName:string);
var
  blob:TBlobField;
begin

  if not FileExists(fileName) then
    exit;

  blob:=TBlobField(self.table.FieldByName(fieldName));

  blob.LoadFromFile(fileName);

end;

procedure TSG_BlobFieldsManager.fieldToFile(fieldname,archivo:string);
var
  blob:TBlobField;
begin

  blob:=TBlobField(Self.Table.FieldByName(fieldName));

  if not blob.IsNull then
    blob.SaveToFile(archivo);

end;

procedure TSG_BlobFieldsManager.saveFields;
begin

  // Here you have to put this for each blob field:
  //Self.fileToField([name of the field],[file name]);

  // if you want to delete the file:
  //DeleteFile(fileName);

end;

procedure TSG_BlobFieldsManager.readFields;
begin

  // Here you have to put this for each blob field:
  //Self.fieldToFile([name of the field],[file name]);

end;

procedure TSG_BlobFieldsManager.saveFieldsTemplate;
begin

  if not Self.goToKey then
  begin
    Self.Table.Insert;
    Self.Table['id']:=1;
  end
  else
    Self.Table.edit;

  Self.saveFields;

  Self.Table.Post;

  Self.Table.Close;

end;

procedure TSG_BlobFieldsManager.readFieldsTemplate;
begin

  if Self.goToKey then
  begin

    Self.readFields;

  end;

  Self.Table.Close;

end;

end.
