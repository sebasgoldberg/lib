unit SG_Date;

interface

uses
  Controls;

function delphiDateToSqlDate(date:TDate):String;
function addYears(date:TDate;years:integer):TDate;

implementation

uses
  Sysutils;

function delphiDateToSqlDate(date:TDate):String;
begin
  Result:=QuotedStr(FormatDateTime('yyyymmdd',date));
end;

function addYears(date:TDate;years:integer):TDate;
begin
  Result:=IncMonth(Date,12*years);
end;

end.
