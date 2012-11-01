unit SG_Messages;

interface

function yesNoQuestion(text:string):boolean;
procedure information(text:string);

implementation

uses
  Dialogs,
  Controls;
  
function yesNoQuestion(text:string):boolean;
begin
  Result:=MessageDlg(text,mtConfirmation,[mbYes, mbNo],0)=mrYes;
end;

procedure information(text:string);
begin
  MessageDlg(text,mtInformation,[mbOK],0);
end;

end.
