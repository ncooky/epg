unit frm_dm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, DCPcrypt2, DCPmd5, INIFiles;
type
  Tdm = class(TForm)
    EPG_DB: TADOConnection;
    DDL: TADOQuery;
    dml: TADOQuery;
    EPG_Access_DB: TADOConnection;
    AccDDL: TADOQuery;
    Accdml: TADOQuery;
    EPG_DB_2: TADOConnection;
    DDL2: TADOQuery;
    dml2: TADOQuery;
    dmlTanggal: TADOQuery;
    DDLTanggal: TADOQuery;
    DDLIDTable: TADOQuery;
    DDLDateTime: TADOQuery;
    DDLPush: TADOQuery;

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dm: Tdm;
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
  strSQL, strIMGChoice :String;

  procedure RecSetPush(strSQL : String);
  procedure RecSet(strSQL : String);
  procedure RecExc(strSQL : String);
  procedure RecSetAcc(strSQL : String);
  procedure RecExcAcc(strSQL : String);
  procedure JulianToSolar(var Y,M,D:Word);
  procedure RecExc2(strSQL : String);
  procedure RecSet2(strSQL : String);
  Function fncMD5(text:string):string;

  procedure RecSetIDTable(strSQL : String);
  procedure RecSetDateTime(strSQL : String);


implementation


{$R *.dfm}
const
Codes64 = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+/';

Var
EPG_DB : TAdoConnection;
EPG_DB_2 : TAdoConnection;

function Encode64(S: string): string;
var
  i: Integer;
  a: Integer;
  x: Integer;
  b: Integer;
begin
  Result := '';
  a := 0;
  b := 0;
  for i := 1 to Length(s) do
  begin
    x := Ord(s[i]);
    b := b * 256 + x;
    a := a + 8;
    while a >= 6 do
    begin
      a := a - 6;
      x := b div (1 shl a);
      b := b mod (1 shl a);
      Result := Result + Codes64[x + 1];
    end;
  end;
  if a > 0 then
  begin
    x := b shl (6 - a);
    Result := Result + Codes64[x + 1];
  end;
end;

function Decode64(S: string): string;
var
  i: Integer;
  a: Integer;
  x: Integer;
  b: Integer;
begin
  Result := '';
  a := 0;
  b := 0;
  for i := 1 to Length(s) do
  begin
    x := Pos(s[i], codes64) - 1;
    if x >= 0 then
    begin
      b := b * 64 + x;
      a := a + 6;
      if a >= 8 then
      begin
        a := a - 8;
        x := b shr a;
        b := b mod (1 shl a);
        x := x mod 256;
        Result := Result + chr(x);
      end;
    end
    else
      Exit;
  end;
end;

procedure RecSetPush(strSQL : String);
var
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');

 dm.DDLPush.Close;
 dm.DDLPush.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;
      end;
    dm.DDLPush.Connection := dm.EPG_DB;

 dm.DDLPush.SQL.Add(strSQL);
 dm.DDLPush.Open;
end;

procedure RecSet(strSQL : String);
var
  database, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');
 dm.DDL.Close;
 dm.DDL.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;

      end;
 dm.DDL.Connection := dm.EPG_DB;
 dm.DDL.SQL.Add(strSQL);
 dm.DDL.Open;
end;

procedure RecExc(strSQL : String);
var
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');

 dm.dml.Close;
 dm.dml.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;

      end;

 dm.dml.Connection := dm.EPG_DB;
 dm.dml.SQL.Add(strSQL);
 dm.dml.ExecSQL;
end;

procedure RecSetAcc(strSQL : String);
var
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');

 dm.AccDDL.Close;
 dm.AccDDL.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;

      end;

 dm.AccDDL.Connection := dm.EPG_DB;
 dm.AccDDL.SQL.Add(strSQL);
 dm.AccDDL.Open;
end;

procedure RecExcAcc(strSQL : String);
var
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');

 dm.Accdml.Close;
 dm.Accdml.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;

      end;

 dm.Accdml.Connection := dm.EPG_DB;
 dm.Accdml.SQL.Add(strSQL);
 dm.Accdml.ExecSQL;
end;

procedure RecSet2(strSQL : String);
var
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');

 dm.DDL2.Close;
 dm.DDL2.SQL.Clear;

    if (not dm.EPG_DB_2.Connected) then
      begin
        dm.EPG_DB_2 := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB_2.LoginPrompt:=False;
        dm.EPG_DB_2.ConnectionString := conStr;
        dm.EPG_DB_2.Connected := True;

      end;

 dm.DDL2.Connection := dm.EPG_DB_2;
 dm.DDL2.SQL.Add(strSQL);
 dm.DDL2.Open;
end;

procedure RecExc2(strSQL : String);
var
  dataBase, user, pwd: string;
  Ini: TIniFile;
  conStr : string;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');

 dm.dml2.Close;
 dm.dml2.SQL.Clear;

    if (not dm.EPG_DB_2.Connected) then
      begin
        dm.EPG_DB_2 := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB_2.LoginPrompt:=False;
        dm.EPG_DB_2.ConnectionString := conStr;
        dm.EPG_DB_2.Connected := True;

      end;

 dm.dml2.Connection := dm.EPG_DB_2;      
 dm.dml2.SQL.Add(strSQL);
 dm.dml2.ExecSQL;
end;

procedure JulianToSolar(var Y,M,D:Word);
const
  SolarDays : array[1..12] of Word = (31,31,31,31,31,31,30,30,30,30,30,29) ;
  JulianDays : array[1..12] of Word = (31,28,31,30,31,30,31,31,30,31,30,31) ;
var
  R: Real;
  Kabiseh,Kably: Boolean;
  DaysSum,Days: Word;
  SolarYear,SolarMonth,SolarDay: Word;
  I: Integer;
begin
DaysSum := 0 ;
if Y=0 then
Y:=2000
else if Y<1900 then
Y:=1900+Y;
R:=Abs(y-1996)/4;
SolarDay:=0;
SolarMonth:=0;
Days:=0;
if(R=Trunc(R))then
Kabiseh:=True else
Kabiseh:=False;
R:=Abs(y-1-1996)/4;
if(R=Trunc(R))then
Kably:=True else
Kably:=False;
if(m>1)then
for I := m downto 2 do
DaysSum:=DaysSum+JulianDays[I-1];
DaysSum:=DaysSum+d;
if(kabiseh)and(DaysSum>=59)then
Inc(DaysSum);
if(Kabiseh and(DaysSum<80))
or(not Kabiseh and (DaysSum<79))
or(not Kabiseh and Kably and (DaysSum<80))then
SolarYear:=y-622 else
SolarYear:=y-621;
if Kabiseh and (DaysSum>79) then days:=DaysSum-79;
if Kabiseh and (DaysSum<80) then Days:=DaysSum+286 ;
if not Kabiseh and not Kably and (DaysSum>79) then Days:=DaysSum-79;
if not Kabiseh and not Kably and (DaysSum<80) then Days:=DaysSum+286;
if not Kabiseh and Kably and (DaysSum>79) then Days:=DaysSum-79;
if not Kabiseh and Kably and (DaysSum<80) then Days:=DaysSum+287;
DaysSum:=Days;
if(daysSum<=186)then
begin
 SolarMonth:=(DaysSum div 31)+1;
 SolarDay:=DaysSum mod 31;
 If SolarDay=0 then SolarDay:=31;
 If SolarDay=31 then SolarMonth:=SolarMonth-1;
end;
if DaysSum > 186 then
begin
 DaysSum:=DaysSum-186;
 if DaysSum mod 30 = 0 then
 SolarMonth:=(DaysSum div 30)+6 else
 SolarMonth:=(DaysSum div 30)+7;
 SolarDay:=DaysSum mod 30;
 If SolarDay=0 then SolarDay:=30;
end;
Y:=SolarYear;
D:=SolarDay;
M:=SolarMonth;
end;

Function fncMD5(text:string):string;
var
  Hash: TDCP_md5;
  Digest: array[0..15] of byte;
  Source: String;
  i: integer;
  s: string;
 begin
  //Edit2.Clear;
  Source := text;
  Hash:= TDCP_md5.Create(Hash);
  Hash.Init;
  Hash.UpdateStr(Source);
  Hash.Final(Digest);
  s:= '';
  for i:= 0 to 15 do
    s:= s + IntToHex(Digest[i],2);
  //Edit2.Text:= s;
  fncMD5:= s;
end;

procedure RecSetIDTable(strSQL : String);
 begin
   dm.DDLIDTable.Close;
   dm.DDLIDTable.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;

      end;

   dm.DDLIDTable.Connection := dm.EPG_DB;
   dm.DDLIDTable.SQL.Add(strSQL);
   dm.DDLIDTable.Open;
 end;

procedure RecSetDateTime(strSQL : String);
  begin
    dm.DDLDateTime.Close;
    dm.DDLDateTime.SQL.Clear;

    if (not dm.EPG_DB.Connected) then
      begin
        dm.EPG_DB := TADOConnection.Create(nil);
        database := Decode64(Ini.ReadString('Config', 'database', 'Default'));
        user := Decode64(Ini.ReadString('Config', 'user', 'Default'));
        pwd := Decode64(Ini.ReadString('Config', 'pwd', 'Default'));
        conStr :='Provider=MSDAORA.1;User ID='+ user +';Password='+ pwd +';Data Source='+ database +';Persist Security Info=False';
        dm.EPG_DB.LoginPrompt:=False;
        dm.EPG_DB.ConnectionString := conStr;
        dm.EPG_DB.Connected := True;

      end;

    dm.DDLDateTime.Connection := dm.EPG_DB;
    dm.DDLDateTime.SQL.Add(strSQL);
    dm.DDLDateTime.Open;
  end;


end.



