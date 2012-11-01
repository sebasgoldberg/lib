unit SG_CheckBox;

interface

uses
  stdctrls;

function getCheckBoxSQLValue(checkBox:TCheckBox):string;

implementation

function getCheckBoxSQLValue(checkBox:TCheckBox):string;
begin
  if checkBox.Checked then
    Result:='1'
  else
    Result:='0';
end;

end.
