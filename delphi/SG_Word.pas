unit SG_Word;

interface

{
example:

  Self.word:=TSG_Word.Create(Self.OleContainer,Self.archivo);

  // to open...
  Self.word.open;

  // to save...
  Self.word.save;

}

uses
  olectnrs,
  controls;

type
  TSG_Word=class

      procedure OleContainerResize(Sender: TObject);

    private
      fileName:string;
      OleContainer:TOleContainer;
      winControl:TWinControl;

      procedure show;

    public
      constructor Create(OleContainer:TOleContainer;fileName:string;
        winControl:TWinControl);
      procedure open;
      procedure save;
  end;

implementation

uses
  Sysutils,
  stdctrls;

constructor TSG_Word.Create(OleContainer:TOleContainer;fileName:string;
  winControl:TWinControl);
begin
  Self.fileName:=fileName;
  Self.OleContainer:=OleContainer;
  Self.OleContainer.OnResize:=Self.OleContainerResize;
  Self.winControl:=winControl;
end;

procedure TSG_Word.open;
begin

  if not FileExists(Self.fileName) then
    OLEContainer.CreateObject('Word.Document.8',false)
  else
  begin
    OLEContainer.CreateObjectFromFile(Self.fileName,false);
  end;

  Self.show;

end;

procedure TSG_Word.save;
begin

  OLEContainer.OleObject.application.ActiveDocument.SaveAs(Self.fileName);
  OLEContainer.Close;

end;

procedure TSG_Word.OleContainerResize(Sender: TObject);
begin

  if (Self.OleContainer.State=osEmpty) or (Self.OleContainer.State=osUIActive)
    then
    exit;

  Self.show;

end;

procedure TSG_Word.show;
begin
  Self.OleContainer.doVerb(ovShow);
  Self.winControl.visible:=true;
  Self.winControl.SetFocus;
  Self.winControl.visible:=false;
end;

end.
