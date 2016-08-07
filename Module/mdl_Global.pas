unit mdl_Global;

interface

uses
  SysUtils, Classes, Messages;

type
  TmdlGlobal = class(TDataModule)
  private
    { Private declarations }
  public
    { Public declarations }
  end;

{Const
  strClPink = $FFCCFF;}
var
  mdlGlobal     : TmdlGlobal;
  strIP         : String;
  strCompName   : String;
  dtLogin       : TDateTime;
  intUserLevel  : Integer;
  strFullName   : String;
  strProgramVersion : String;
  blnNewMO      : Boolean;
  blnNewFrmRevenue : boolean;


  function fncDateTime(): TDateTime;
  function fncTimeToMinute(strTime:String): Integer;
  function fncTimeToSecond(strTime:String): Integer;
  function fncSecondToHour(intSecond:Integer) : String;
  function fncstrFileVersion(): string;

  function fncstrIDTable(strTableName:String; strPK:String) : String;

implementation

uses Winsock, Variants, DateUtils, StdCtrls, Windows, frm_dm;
{$R *.dfm}

function fncstrFileVersion(): string;
var
  N, Len: DWORD;
  Buf: PChar;
  Value: PChar;
  Filename: string;
begin
  Result := '';
  Filename := 'EPG.exe';
  N := GetFileVersionInfoSize(PChar(Filename), N);
  if N > 0 then
  begin
     Buf := AllocMem(N);
     GetFileVersionInfo(PChar(Filename), 0, N, Buf);
     if VerQueryValue(Buf,
                      PChar('StringFileInfo\040904E4\FileVersion'),
                      Pointer(Value), Len) then
        Result := Value;
     FreeMem(Buf, N);
  end;
end;

function fncSecondToHour(intSecond:Integer) : String;
var
  strHour24 : string;
Begin
  fncSecondToHour := '';

  If Length(IntToStr(intSecond div 3600)) = 1
    Then strHour24 := '0' + IntToStr(intSecond div 3600)
      Else strHour24 := IntToStr(intSecond div 3600);
  strHour24 := strHour24 + ':';

  intSecond := intSecond mod 3600;
  If Length(IntToStr(intSecond div 60)) = 1
    Then strHour24 := strHour24 + '0' + IntToStr(intSecond div 60)
      Else strHour24 := strHour24 + IntToStr(intSecond div 60);
  strHour24 := strHour24 + ':';

  If Length(IntToStr(intSecond mod 60)) = 1
    Then strHour24 := strHour24 + '0' + IntToStr(intSecond mod 60)
      Else strHour24 := strHour24 + IntToStr(intSecond mod 60);

  fncSecondToHour := strHour24;
End;

function fncstrIDTable(strTableName:String; strPK:string) : String;
var
  strSQLID : String;
  intID    : Integer;
begin
  strSQLID := 'SELECT MAX(' + strPK + ') as ' + strPK + ' FROM ' + strTableName ;
  RecSetIDTable(strSQLID);
  fncstrIDTable := '0001';
  if not dm.DDLIDTable.Eof Then
    Begin
      if not VarIsNull(dm.DDLIDTable.FieldValues[strPK]) Then
        Begin
          intID := StrToInt(dm.DDLIDTable.FieldValues[strPK]) + 1;
          fncstrIDTable := IntToStr(intID);
        End;
    End;
end;

function fncTimeToMinute(strTime:String): Integer;
var
  intMinute: Integer;
Begin
//  00:00:00
  intMinute := StrToInt(copy(strTime, 1, 2)) * 60;
  intMinute := intMinute + StrToInt(copy(strTime, 4, 2));
  fncTimeToMinute := intMinute;
end;

function fncTimeToSecond(strTime:String): Integer;
var
  intSecond : integer;
begin
  intSecond := StrToInt(copy(strTime, 1, 2)) * 3600;
  intSecond := intSecond + (StrToInt(copy(strTime, 4, 2)) * 60);
  intSecond := intSecond + (StrToInt(copy(strTime, 7, 2)));
  fncTimeToSecond := intSecond;
end;

function fncDateTime(): TDateTime;
var
  strSQLTanggalServer :String;
begin
  strSQLTanggalServer := 'SELECT SysDate FROM DUAL';
  RecSetDateTime(strSQLTanggalServer);
  fncDateTime := dm.DDLDateTime.FieldValues['SysDate'];
end;

end.
