unit frm_Export;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, StrUtils,
  Dialogs, StdCtrls, ExtCtrls, jpeg;

type
  TfrmExport = class(TForm)
    GroupBox1: TGroupBox;
    cbDateEnd: TComboBox;
    cbDateStart: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Button2: TButton;
    ComboBox1: TComboBox;
    Image1: TImage;
    cbVOD: TCheckBox;
    lblDate: TLabel;
    procedure FormShow(Sender: TObject);
    procedure cbDateStartSelect(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure cbDateEndSelect(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmExport: TfrmExport;
  date1, date2 : String;

implementation

{$R *.dfm}

uses
  ComObj, frm_dm, DateUtils, frm_Read, DB, frm_Login, frm_EPG;

const
Codes64 = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+/';

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



procedure TfrmExport.FormShow(Sender: TObject);
var
 item : TStrings;
begin
 cbDateStart.Items.Clear;
 strSQL := 'SELECT to_date(Date_Schedule,''mm/dd/yyyy'') AS Dates FROM ( SELECT distinct to_char(rscheduledate,''mm/dd/yyyy'') AS Date_Schedule ';
 strSQL := strSQL + 'FROM m_readxl ';
 strSQL := strSQL + 'WHERE rchannel = ''' + frmRead.ngReadXL.Cells[1,1] + ''' ';
 strSQL := strSQL + 'AND to_date(to_char(rscheduledate,''mm/dd/yyyy''),''mm/dd/yyyy'') >= to_date(to_char(sysdate,''mm/dd/yyyy''),''mm/dd/yyyy'') ';
 strSQL := strSQL + 'ORDER by Date_Schedule ) ORDER by Dates';
 RecSet(strSQL);

 Item:=cbDateStart.Items.Create;
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['Dates']);
  dm.DDL.Next;
 end;

 strSQL := 'Select sysdate from dual';
 RecSet(strSQL);
 lblDate.Caption := copy(dm.DDL.FieldValues['sysdate'],1,10);
 SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NoMove or SWP_NoSize);
end;

procedure TfrmExport.cbDateStartSelect(Sender: TObject);
var
 item : TStrings;
begin
 cbDateEnd.Items.Clear;
 strSQL := 'SELECT to_date(Date_Schedule,''mm/dd/yyyy'') AS Dates FROM ( SELECT distinct to_char(rscheduledate,''mm/dd/yyyy'') AS Date_Schedule ';
 strSQL := strSQL + 'FROM m_readxl ';
 strSQL := strSQL + 'WHERE rchannel = ''' + frmRead.ngReadXL.Cells[1,1] + ''' ';
 strSQL := strSQL + 'AND rscheduledate >= to_date(''' + cbDateStart.Text +  ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
 strSQL := strSQL + 'ORDER by Date_Schedule ) ORDER by Dates';
 RecSet(strSQL);

 Item:=cbDateEnd.Items.Create;
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['Dates']);
  dm.DDL.Next;
 end;
 dm.DDL.First;
 date1 := FormatDateTime('mmdd',dm.DDL.FieldValues['DATES']);
end;

procedure TfrmExport.Button2Click(Sender: TObject);
begin
 frmExport.Close;
end;


function MidStr
    (Const Str: String; From, Size: Word): String;
begin
  MidStr := Copy(Str, From, Size)
end;

function RightStr
    (Const Str:

String; Size: Word): String;
begin
  if Size > Length(Str) then Size := Length(Str) ;
  RightStr := Copy(Str, Length(Str)-Size+1, Size)
end;

function str_replace(const oldChars, newChars: array of Char; const str: string): string;
  var
    i, j: Integer;
  begin
    Assert(Length(oldChars)=Length(newChars));
    Result := str;
    for i := 1 to Length(Result) do
      for j := 0 to high(oldChars) do
        if Result[i]=oldChars[j] then
        begin
          Result[i] := newChars[j];
          break;
        end;
  end;

function StripNonAlpha(aInput : String) : String;
const
Char = ['0'..'9','A'..'Z','a'..'z','?','.','>','<','+','-','~','!','@','#','$','%','&','*','(',')','_','=','{','}','[',']','|','\','/',':',';',' ', '''', ',', '"'];
var I : Integer;
begin
  result := aInput;
  for I := 1 to length(result)
  do
    begin
      if not (result[I] in Char) then
        result[I] := ' ';
    end;
end;

function ascii
  (Const Str: AnsiString): AnsiString;
begin
  if AnsiContainsText(Str, ' & ')
  then ascii := StringReplace(Str, ' & ', ' &#38; ', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '&')
  then ascii := StringReplace(Str, '&', '&#38;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#x00BF;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#225;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '!')
  then ascii := StringReplace(Str, '!', '&#33;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '"')
  then ascii := StringReplace(Str, '"', '&#34;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '#')
  then ascii := StringReplace(Str, '#', '&#35;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '$')
  then ascii := StringReplace(Str, '$', '&#36;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '%')
  then ascii := StringReplace(Str, '%', '&#37;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '(')
  then ascii := StringReplace(Str, '(', '&#40;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, ')')
  then ascii := StringReplace(Str, ')', '&#41;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '*')
  then ascii := StringReplace(Str, '*', '&#42;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '+')
  then ascii := StringReplace(Str, '+', '&#43;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '-')
  then ascii := StringReplace(Str, '-', '&#45;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '/')
  then ascii := StringReplace(Str, '/', '&#47;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '<')
  then ascii := StringReplace(Str, '<', '&#60;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '=')
  then ascii := StringReplace(Str, '=', '&#61;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '>')
  then ascii := StringReplace(Str, '>', '&#62;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '[')
  then ascii := StringReplace(Str, '[', '&#91;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '\')
  then ascii := StringReplace(Str, '\', '&#92;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, ']')
  then ascii := StringReplace(Str, ']', '&#93;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '^')
  then ascii := StringReplace(Str, '^', '&#94;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '_')
  then ascii := StringReplace(Str, '_', '&#95;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '`')
  then ascii := StringReplace(Str, '`', '&#96;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#160;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#161;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#162;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#163;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#164;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#165;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#166;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#167;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#168;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#169;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#170;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#171;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#172;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '� ')
  then ascii := StringReplace(Str, '� ', '&#174; ', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#175;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '� ')
  then ascii := StringReplace(Str, '� ', '&#176; ', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#177;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#178;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#179;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#180;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#181;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#182;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#183;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#184;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#185;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#186;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#187;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#188;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#189;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#190;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#191;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#192;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#193;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#194;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#195;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#196;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#197;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#198;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#199;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#200;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#201;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#202;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#203;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#204;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#205;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#206;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#207;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#208;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#209;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#210;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#211;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#212;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#213;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#214;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#215;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#216;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#217;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#218;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#219;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#220;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#221;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#222;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#223;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#224;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#225;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#226;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#227;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#228;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#229;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#230;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#231;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#232;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#233;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#234;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#235;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#236;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#237;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#238;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#239;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#240;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#241;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#242;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#243;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#244;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#245;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#246;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#247;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#248;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#249;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#250;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#251;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#252;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#253;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#254;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#255;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8211;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8212;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8216;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8217;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8218;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8220;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8221;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8222;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8224;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8225;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8226;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8230;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8240;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8364;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#8482;', [rfReplaceAll, rfIgnoreCase])

  else ascii := StripNonAlpha(Str);

end;
  
procedure TfrmExport.Button1Click(Sender: TObject);
var
	XML, SGI, BB, Sindo, XTI : TextFile;
	catxt, catxtvis, catxtxti, StrSQLtemp, strEPG, strSynInd, strSynEng, strContent, strChnlNum : String;
	strAmtPackage : String;
	strRating, beforeXML, afterXML, TrimTitle, beforeChnl, afterChnl, bfchname, aftchname, tca, tcadate, catcaxti: String;
	i, ii, x, xtca: Integer;
	strCATemplate, strDatestart, strDateEnd : String;
  AsciiTab : Char;
  PosEp, PosSes, PosKoma, ResSes, ResEp : Integer;
  strEp, strSes, tPosSes : String;
  NotEp, NotSes : Variant;
  sesChar : Char;
  AnsiSynEng, AnsiSynInd, ansiChannel: AnsiString;
  jDate, endDate: TDateTime;
  mjdfloat, bfloat: extended;  
  
const
	sLineBreak = {$IFDEF LINUX} AnsiChar(#10) {$ENDIF}
		{$IFDEF MSWINDOWS} AnsiString(#13#10) {$ENDIF};

const
    Numbers = '0123456789';


begin

try
	{if frmExport.cbVOD.Checked = True then
		AssignFile(SGI, 'C:\SGI\REV_VOD_' + frmRead.ngReadXL.Cells[1,1]+'_'+ date1+ '-' +date2 +'.sgi')
	Else}

	if not DirectoryExists('C:\SGI') then
		begin
			CreateDir('C:\SGI');
		end;
	if not DirectoryExists('C:\SGI\SGI_NDS') then
		begin
			CreateDir('C:\SGI\SGI_NDS');
		end;
	if not DirectoryExists('C:\SGI\SGI_BB') then
		begin
			CreateDir('C:\SGI\SGI_BB');
		end;
	if not DirectoryExists('C:\SGI\SGI_SINDO') then
		begin
			CreateDir('C:\SGI\SGI_SINDO');
		end;
	if not DirectoryExists('C:\SGI\SGI_XML') then
		begin
			CreateDir('C:\SGI\SGI_XML');
		end;
	if not DirectoryExists('C:\SGI\SGI_XTI') then
		begin
			CreateDir('C:\SGI\SGI_XTI');
		end;
	if not DirectoryExists('C:\SGI\LOG_ERROR') then
		begin
			CreateDir('C:\SGI\LOG_ERROR');
		end;

//  AssignFile(actLOG, '\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
    //AssignFile(actLOG, 'C:\SGI\SGI_LOG\' + strUserName +'.log');
    
	AssignFile(SGI, 'C:\SGI\SGI_NDS\' + frmRead.ngReadXL.Cells[1,1]+'_'+ date1+ '-' +date2 +'.sgi');
	AssignFile(BB , 'C:\SGI\SGI_BB\' + frmRead.ngReadXL.Cells[1,1]+'_'+ date1+ '-' +date2 +'-BlackBerry.sgi');
	AssignFile(Sindo, 'C:\SGI\SGI_SINDO\' + frmRead.ngReadXL.Cells[1,1]+'_'+ date1+ '-' +date2 +'-Sindo.csv');
	AssignFile(XML, 'C:\SGI\SGI_XML\' + frmRead.ngReadXL.Cells[1,1]+'_'+ date1+ '-' + date2 +'.xml');

  bfchname := frmRead.ngReadXL.Cells[1,1];
  aftchname := StringReplace(bfchname, ' ', '_', [rfReplaceAll, rfIgnoreCase]);

	AssignFile(XTI , 'C:\SGI\SGI_XTI\' + aftchname + '_'+ date1+ '-' + date2 +'.xml');

try
  AssignFile(actLOGLocal, 'C:\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
  AssignFile(actLOG, '\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
  if fileexists('C:\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
    then append(actLOGLocal)
    else Rewrite(actLOGLocal);

    if fileexists('\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
      then append(actLOG)
      else Rewrite(actLOG);
except
     on E : Exception do
     begin
       AssignFile(ERR , 'C:\SGI\LOG_ERROR\FormExport_'+strUser+'.log');
       if fileexists('C:\SGI\LOG_ERROR\FormExport_'+strUser+'.log')
       then append(ERR)
       else Rewrite(ERR);

       Writeln(ERR , encode64('Koneksi ke data_traffic gagal -> Err.Class: '+ E.ClassName+ ', pesan errornya gini: '+ E.Message) );
       CloseFile(ERR);
       showmessage('Maaf, terjadi kesalahan koneksi, silahkan periksa jaringan ke direktori data_traffic'+sLineBreak+''+sLineBreak+'Terima Kasih' );
       //Writeln(actLOG,'[', FormatDateTime('c',today),'] ', ' FrmRead : koneksi data_traffic gagal: ',E.Message );
       Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' FrmRead : koneksi data_traffic gagal: ',E.Message );
       CloseFile(actLOGLocal);
       frmRead.Close;
     end;
end;

  //Rewrite(actLOG);

	Rewrite(SGI);
	Rewrite(BB);
	Rewrite(Sindo);
	Rewrite(XML);
  Rewrite(XTI);

 ////////////////////// XPush Channel ///////////////////////////////////
	for ii:=0 to ComboBox1.Items.Count-1 do
		begin
			strSQL:='SELECT DISTINCT VODCAPRODUCTID, to_char(VODCAPSTARTDATE,''ddmmyyyy'') AS VODCAPSTARTDATE, to_char(VODCAPENDDATE,''ddmmyyyy'') AS VODCAPENDDATE, VODCASERVICEID, VODEPGTITLE, VODPROGRAMID, VODTRAFFICKEY, ';
			strSQL:=strSQL + 'VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, VODUSERCREATEDATE, ca, mcsiserviceid, mchannel, msginame, mplayout_source ';
			strSQL:=strSQL + 'FROM (  SELECT * ';
			strSQL:=strSQL + 'FROM ( ';
			strSQL:=strSQL + 'SELECT mc.mcsiserviceid, mc.mchannel, mc.mplayout_source, mr.rscheduledate, mr.REPG_TITLE, mr.RDURATION, mr.RRATING, ';
			strSQL:=strSQL + 'mr.RGENRE, mr.RSUBGENRE, mr.RCONTENT, to_char(mr.rscheduledate,''ddmmyyyy'') AS EventStartDate, ';
			strSQL:=strSQL + 'to_char(mr.rscheduledate,''hh24miss'') AS EventStartTime, to_char(mr.rscheduledategmt,''hh24miss'') AS EventStartTimegmt, ';
			strSQL:=strSQL + 'to_char(mr.rscheduledategmt,''ddmmyyyy'') AS EventStartDategmt, msginame, MUSERNIBBLE1, mr.RCATEMPLATE, REPG_TITLE_ORI, mSYNOPSIS_STATUS ';
			strSQL:=strSQL + 'FROM m_channel mc, m_readxl mr ';
			strSQL:=strSQL + 'WHERE mc.mchannel = ''' + frmRead.ngReadXL.Cells[1,1] + ''' ';
			strSQL:=strSQL + 'AND mr.rchannel = mc.mchannel ';
			strSQL:=strSQL + 'AND mr.rscheduledate >= to_date(''' + frmExport.cbDateStart.Items.Strings[ii] +  ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
			strSQL:=strSQL + 'AND mr.rscheduledate <= to_date(''' + frmExport.cbDateStart.Items.Strings[ii] +  ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
			strSQL:=strSQL + ')aaa, ';
			strSQL:=strSQL + '(SELECT count(mca.capackage)+2 as ca FROM m_ca_package mca WHERE mca.cachannel=''' + frmRead.ngReadXL.Cells[1,1] + ''' ) bbb  ) XXX, ';
			strSQL:=strSQL + '( SELECT syEPG_TITLE, SYSynopsis_Ind, SYSynopsis_Eng, sycategory ';
			strSQL:=strSQL + 'FROM M_Synopsis ) YYY, (SELECT * from m_image ) ZZZ, (select * from M_VOD) WWW ';
			strSQL:=strSQL + 'WHERE REPG_TITLE_ORI = syEPG_TITLE(+) AND RGENRE = sycategory(+) AND REPG_TITLE_ORI = iepg_ori(+) AND mchannel=ichannel (+) AND REPG_TITLE_ORI = VODEPGTITLE(+) ORDER BY MChannel';
			RecSetPush(strSQL);

    {if VarIsNull(dm.DDL.FieldValues['VODCAPRODUCTID']) then 
		strCAKosong := ''
	else 
		strCAKosong := dm.DDL.FieldValues['VODCAPRODUCTID'];}

    While not dm.DDLPush.Eof do
		begin
			if not VarIsNull(dm.DDLPush.FieldValues['VODPROGRAMKEY']) and (copy(dm.DDLPush.FieldValues['VODUSERCREATEDATE'],1,10) = lblDate.Caption) and (cbVOD.Checked = False) then
				Begin
					if (dm.DDLPush.FieldValues['VODGROUPKEY']= '12346') or (VarIsNull(dm.DDLPush.FieldValues['VODGROUPKEY'])) then
						Writeln(SGI,'8~',dm.DDLPush.FieldValues['VODCAPRODUCTID'],'~2497~1~B~3~',dm.DDLPush.FieldValues['VODCAPSTARTDATE'],'00000000~',dm.DDLPush.FieldValues['VODCAPENDDATE'],'00000000~',dm.DDLPush.FieldValues['VODCASERVICEID'],'~');
				End;
			dm.DDLPush.Next;
		end;
	End;
  ////////////////////// XPush Channel ///////////////////////////////////

	Writeln(SGI,'5~0700~~~');
	Writeln(BB ,'5~0700~~~');
	Writeln(Sindo,'CHANNEL''S NAME',',','START DATE',',','START TIME',',','DURATION',',','TITLE',',','SYNOPSIS INDONESIA',',','SYNOPSIS ENGLISH');
	Writeln(XML ,'<?xml version="1.0" encoding="ISO-8859-1"?>'+sLineBreak+'<data-set xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');  
 	Writeln(XTI ,'<?xml version="1.0" encoding="UTF-8"?>'+sLineBreak+'<BasicImport xmlns="http://www.uk.nds.com/SSR/XTI/Traffic/0010" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.uk.nds.com/SSR/XTI/Traffic/0010 0010.xsd" utcOffset="+07:00" frameRate="25">');
	i:=2;

	strSQL := 'select mcsiserviceid from m_channel where mchannel = ''' + frmRead.ngReadXL.Cells[1,1] + ''' ';
	RecSet(strSQL);
	strCATemplate := dm.DDL.FieldValues['mcsiserviceid'];
  ansiChannel := stringreplace(frmRead.ngReadXL.Cells[1,1], '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);


  strSQL := 'Select rscheduledate From m_readxl where ';
  strSQL := strSQL + 'RSCHEDULEDATE BETWEEN to_date('''+FormatDateTime('mm/dd/yyyy',StrToDateTime(cbDateStart.Text))+' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
  strSQL := strSQL + 'AND to_date('''+FormatDateTime('mm/dd/yyyy',StrToDateTime(cbDateEnd.Text))+' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
  strSQL := strSQL + 'AND RCHANNEL = ''' +trim(frmRead.ngReadXL.Cells[1,1]) +''' ';
  strSQL := strSQL + 'AND ROWNUM = 1 ';
  RecSet(strSQL);

  strDatestart := FormatDateTime('hh:nn:ss',dm.DDL.FieldValues['rscheduledate']);

  strSQL := 'SELECT tt.rscheduledate, rduration FROM m_readxl tt INNER JOIN ';
  strSQL := strSQL + '(SELECT rchannel, MAX(rscheduledate) AS MaxDateTime FROM m_readxl where ';
  strSQL := strSQL + 'RSCHEDULEDATE BETWEEN to_date('''+FormatDateTime('mm/dd/yyyy',StrToDateTime(cbDateStart.Text))+' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
  strSQL := strSQL + 'AND to_date('''+FormatDateTime('mm/dd/yyyy',StrToDateTime(cbDateEnd.Text))+' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
  strSQL := strSQL + 'AND RCHANNEL = ''' +trim(frmRead.ngReadXL.Cells[1,1]) +''' ';
  strSQL := strSQL + 'GROUP BY rchannel) groupedtt ';
  strSQL := strSQL + 'ON tt.rchannel = groupedtt.rchannel ';
  strSQL := strSQL + 'AND tt.rscheduledate = groupedtt.MaxDateTime ';
  RecSet(strSQL);

  strDateEnd:= copy(dm.DDL.FieldValues['rduration'],1,2)+':'+copy(dm.DDL.FieldValues['rduration'],3,2)+':'+copy(dm.DDL.FieldValues['rduration'],5,2);
  endDate := StrToDateTime(dm.DDL.FieldValues['rscheduledate']);
  endDate := endDate + StrToTime(strDateEnd) - StrToTime ('00:00:01');

  Writeln(XTI ,'<SiEventSchedule deleteStart="'+ FormatDateTime('yyyy/mm/dd',StrToDateTime(cbDateStart.Text)) +  ' '+strDatestart+'" deleteEnd="' + FormatDateTime('yyyy/mm/dd hh:nn:ss',endDate) +  '">');
  Writeln(XTI ,'<siService>',ansiChannel,'</siService>');
  Writeln(XTI ,'<playoutSource>',dm.DDLPush.FieldValues['mplayout_source'],'</playoutSource>');
  Writeln(XTI ,'<activationSource>CHRONOLOGICAL</activationSource>');
  Writeln(XTI ,'<CaSchedule deleteStart="'+ FormatDateTime('yyyy/mm/dd',StrToDateTime(cbDateStart.Text)) +  ' '+strDatestart+'" deleteEnd="' + FormatDateTime('yyyy/mm/dd hh:nn:ss',endDate) +  '" />');
  AsciiTab := Char(09);

	if strCATemplate = '1001' then
		begin
			catxt:='';
		end
	else
		begin
			strSQL := 'SELECT CCADescription ';
			strSQL := strSQL + 'FROM M_CA_PACKAGE, M_CASERVICEID ';
			strSQL := strSQL + 'WHERE CCAID = capackage ';
			strSQL := strSQL + '      AND cachannel = ''' + frmRead.ngReadXL.Cells[1,1] + '''  ';
			RecSet(strSQL);
			catxt:='';
      catxtxti:='';
      //if not VarIsNull(dm.DDLPush.FieldValues['VODPROGRAMKEY']) then
      //  begin
      //    x:=2;
      //  end
      //else x:=0;
      x:=2;
			While not dm.DDL.Eof do
				begin
					catxt:=catxt + IntToStr(i) + '~' + dm.DDL.FieldValues['CCADescription'] + '~' ;
          catxtxti:=catxtxti + AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>'+IntToStr(x)+'</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL.FieldValues['CCADescription']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak;
					i:=i+1;
          x:=x+1;
					dm.DDL.Next;
				end;
			end;
			
			for ii:=0 to ComboBox1.Items.Count-1 do
				begin
					if strCATemplate = '1001' then
						begin
							strSQL:='SELECT 2 as ca, mcsiserviceid, mchannel, mplayout_source, mstb_pairing, rscheduledate, REPG_TITLE,  RKEY_HEX, CHNUM, RDURATION, RRATING, RGENRE, RSUBGENRE, RCONTENT, EventStartDate, EventStartTime, ';
						end
          else if (strCATemplate='2002') or (strCATemplate='2202') or (strCATemplate='100') or (strCATemplate='200') or (strCATemplate='2005') then
            begin
              strSQL:='SELECT ca, mcsiserviceid, mchannel, mplayout_source, mstb_pairing, rscheduledate,  CBNUMBER,  REPG_TITLE,  RKEY_HEX, CHNUM, RDURATION, RRATING, RGENRE, RSUBGENRE, RCONTENT, EventStartDate, EventStartTime, ';
            end
					else
						begin
							strSQL:='SELECT ca, mcsiserviceid, mchannel, mplayout_source, mstb_pairing, rscheduledate,  REPG_TITLE,  RKEY_HEX, CHNUM, RDURATION, RRATING, RGENRE, RSUBGENRE, RCONTENT, EventStartDate, EventStartTime, ';
						end;
					strSQL:=strSQL + 'EventStartTimegmt, EventStartDategmt, SYSynopsis_Ind, SYSynopsis_Eng, VODEPGTITLE, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, ';
					strSQL:=strSQL + 'VODPROGRAMID, VODTRAFFICKEY, VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, msginame, MUSERNIBBLE1, RCATEMPLATE, ';
          strSQL:=strSQL + 'SREPGTITLE, SRGROUPKEY, SRPROGRAMKEY, ';
					strSQL:=strSQL + 'mSYNOPSIS_STATUS, IIMAGEID, to_char(rscheduledate,''dd/mm/yyyy'') as stDate, to_char(rscheduledate,''hh24:mi'') as stTime, to_char(rscheduledate,''hh24:mi:ss'') AS stTimeXML, rduration ';
					strSQL:=strSQL + 'FROM (  SELECT * ';
					strSQL:=strSQL + 'FROM ( ';
					strSQL:=strSQL + 'SELECT mc.mcsiserviceid, mc.mchannel, mc.mchannel_number as CHNUM, mc.mplayout_source, mc.mstb_pairing, mr.rscheduledate, mr.REPG_TITLE, mr.RDURATION, mr.RRATING, mr.RKEY_HEX, ';
					strSQL:=strSQL + 'mr.RGENRE, mr.RSUBGENRE, mr.RCONTENT, to_char(mr.rscheduledate,''ddmmyyyy'') AS EventStartDate, ';
					strSQL:=strSQL + 'to_char(mr.rscheduledate,''hh24miss'') AS EventStartTime, to_char(mr.rscheduledategmt,''hh24miss'') AS EventStartTimegmt, ';
					strSQL:=strSQL + 'to_char(mr.rscheduledategmt,''ddmmyyyy'') AS EventStartDategmt, msginame, MUSERNIBBLE1, mr.RCATEMPLATE, REPG_TITLE_ORI, mSYNOPSIS_STATUS ';
					strSQL:=strSQL + 'FROM m_channel mc, m_readxl mr ';
					strSQL:=strSQL + 'WHERE mc.mchannel = ''' + frmRead.ngReadXL.Cells[1,1] + ''' ';
					strSQL:=strSQL + 'AND mr.rchannel = mc.mchannel ';
					strSQL:=strSQL + 'AND mr.rscheduledate >= to_date(''' + frmExport.ComboBox1.Items.Strings[ii] +  ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
					strSQL:=strSQL + 'AND mr.rscheduledate <= to_date(''' + frmExport.ComboBox1.Items.Strings[ii] +  ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
					strSQL:=strSQL + ')aaa, ';
					strSQL:=strSQL + '(SELECT count(mca.capackage)+2 as ca FROM m_ca_package mca WHERE mca.cachannel=''' + frmRead.ngReadXL.Cells[1,1] + ''' ) bbb  ) UUU, ';
					strSQL:=strSQL + '( SELECT syEPG_TITLE, SYSynopsis_Ind, SYSynopsis_Eng, sycategory ';
					strSQL:=strSQL + 'FROM M_Synopsis ) VVV, (SELECT * from m_image ) WWW, (select * from M_VOD) XXX, (select * from M_SERIES) YYY ';
          if (strCATemplate = '2002') or (strCATemplate = '2202') or (strCATemplate='100') or (strCATemplate='200') or (strCATemplate='2005') then
          begin
            strSQL:=strSQL + ', (select * from M_CHANNEL_BITS where CBCHANNEL =''' + frmRead.ngReadXL.Cells[1,1] + ''') ZZZ ';
          end;

					strSQL:=strSQL + 'WHERE REPG_TITLE_ORI = syEPG_TITLE(+) AND RGENRE = sycategory(+) AND REPG_TITLE_ORI = iepg_ori(+) AND mchannel=ichannel (+) ';
          strSQL:=strSQL + 'AND REPG_TITLE_ORI = VODEPGTITLE (+) AND REPG_TITLE = SREPGTITLE(+) ORDER BY MChannel, RScheduleDate ';
					RecSet(strSQL);

					Writeln(SGI,'1~',dm.DDL.FieldValues['MSGINAME'],'~',dm.DDL.FieldValues['EventStartDate'],'~00000000~24000000~ind~0~0~');
					Writeln(BB ,'1~',dm.DDL.FieldValues['mchannel'],'~',dm.DDL.FieldValues['EventStartDate'],'~00000000~24000000~ind~0~0~');

					While not dm.DDL.Eof do
						begin
							strRating := dm.DDL.FieldValues['RRating'];
							strEPG:=Replace(trim(dm.DDL.FieldValues['REPG_TITLE']), ',',';');
							strContent := dm.DDL.FieldValues['RCONTENT'];
              
          beforeXML := trim(dm.DDL.FieldValues['REPG_TITLE']);
          {if AnsiContainsText(beforeXML, ' & ') then
            begin
              afterXML := StringReplace(beforeXML, ' & ', ' &amp; ', [rfReplaceAll, rfIgnoreCase]);
            end
          else if AnsiContainsText(beforeXML, '&') then
            begin
              afterXML := StringReplace(beforeXML, '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);
            end
           else if AnsiContainsText(beforeXML, '�') then
            begin
              afterXML := StringReplace(beforeXML, '�', '&#x00BF;', [rfReplaceAll, rfIgnoreCase]);
            end
          else
            begin
              afterXML := str_replace(
                ['�','�','�','�','�','�','�','�','�','�','�','�'],
                ['a','e','i','o','u','A','E','I','O','U','n','N'],
                beforeXML
              );
            end;}

          afterXML := ascii(beforeXML);

          PosSes := LastDelimiter('S', beforeXML);
          if AnsiContainsText(beforeXML, ',') then
            begin
              if (AnsiContainsText(beforeXML, ':') and AnsiContainsText(beforeXML, ',')) then
                begin
                  PosKoma := LastDelimiter(',', beforeXML);
                end
              else if AnsiContainsText(beforeXML, ':') then
                begin
                  PosKoma := LastDelimiter(':', beforeXML);
                end
              else
                begin
                  PosKoma := LastDelimiter(',', beforeXML);
                end;
            end
          else if AnsiContainsText(beforeXML, ':') then
            begin
              PosKoma := LastDelimiter(':', beforeXML);
            end
          else
            begin
              PosKoma := 0;
            end;

          if PosSes > PosKoma then
            begin
              tPosSes := AnsiLeftStr( beforeXML, PosKoma);
              PosSes := LastDelimiter('S', tPosSes);
            end;

          if AnsiContainsText(afterXML, 'Ep ') then
            begin
              PosEp := LastDelimiter('Ep', afterXML);
            end
          else if AnsiContainsText(afterXML, ':') then
            begin
              PosEp := LastDelimiter(':', afterXML);
            end
          else
            begin
              PosEp := 0;
            end;

          if PosKoma <> 0 then
            begin
              if PosSes <> 0 then
                begin
                  ResSes := PosKoma  - PosSes - 1;
                  if ResSes <> 0 then
                    begin
                      strSes := MidStr(beforeXML,PosSes+1,ResSes );
                      sesChar := strSes[1];

                      if StrScan(Numbers, sesChar) <> nil then
                        begin
                          if AnsiContainsText(afterXML, '&') then
                            begin
                              trimtitle := AnsiLeftStr(afterXML, PosSes + 2);
                            end
                          else trimtitle := AnsiLeftStr(afterXML, PosSes - 2);
                          NotSes := strSes;
                        end
                      else
                        begin
                          if AnsiContainsText(afterXML, '&') then
                            begin
                              trimtitle := AnsiLeftStr(afterXML, PosKoma + 3);
                            end
                          else trimtitle := AnsiLeftStr(afterXML, PosKoma - 1);
                          NotSes := Null;
                        end;
                    end
                  else
                    begin
                      if AnsiContainsText(afterXML, '&') then
                        begin
                          trimtitle := AnsiLeftStr(afterXML, PosKoma + 3);
                        end
                      else trimtitle := AnsiLeftStr(afterXML, PosKoma - 1);                    
                      NotSes := Null;
                    end;
                end
              else
                begin
                  if AnsiContainsText(AnsiLeftStr(afterXML, PosKoma), '&') then
                    begin
                      trimtitle := AnsiLeftStr(afterXML, PosKoma + 3);
                    end
                  else trimtitle := AnsiLeftStr(afterXML, PosKoma - 1);
                  NotSes := Null;
                end;
            end
          else
            begin
              trimtitle := afterXML;
              NotSes := Null;
            end;

          if PosEp <> 0 then
            begin
              ResEp := length(afterXML) - PosEp - 1;
              strEp := RightStr(afterXML, ResEp);
              NotEp :=  strEp;
            end
          else
            begin
              NotEp := Null;
            end;

          {if ansicontainstext(dm.DDL.FieldValues['mchannel'], '&') then
            begin
              ansiChannel := stringreplace(dm.DDL.FieldValues['mchannel'], '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);
            end
           else if ansicontainstext(dm.DDL.FieldValues['mchannel'], '�') then
              begin
                ansiChannel := stringreplace(dm.DDL.FieldValues['mchannel'], '�', '&#x00BF;', [rfReplaceAll, rfIgnoreCase]);
              end
          else
            begin
              ansiChannel := str_replace(
                ['�','�','�','�','�','�','�','�','�','�','�','�'],
                ['a','e','i','o','u','A','E','I','O','U','n','N'],
              dm.DDL.FieldValues['mchannel']
              );
            end; }

          ansiChannel := ascii(dm.DDL.FieldValues['mchannel']);

			if not VarIsNull(dm.DDL.FieldValues['CHNUM']) then
				begin
					strChnlNum := dm.DDL.FieldValues['CHNUM'];
				end
			else
				begin
					strChnlNum := '0';
				end;
			
			if not VarIsNull(dm.DDL.FieldValues['sysynopsis_ind']) then
				begin
					if dm.DDL.FieldValues['MSYNOPSIS_STATUS'] = 'Y' then
						begin
							if not VarIsNull(dm.DDL.FieldValues['IIMAGEID']) then
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',dm.DDL.FieldValues['IIMAGEID']);
								end
							// else if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
              else if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
								begin
                if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                  begin
  									if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
  										Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~',dm.DDL.FieldValues['VODTRAFFICKEY'],'~','8','~',dm.DDL.FieldValues['RCONTENT'],'~~~~~~~~~~~~~0~',dm.DDL.FieldValues['VODPROGRAMID'],'~',trim(dm.DDL.FieldValues['VODPROGRAMKEY']),'~~')
  									else
  										Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~',dm.DDL.FieldValues['VODTRAFFICKEY'],'~','0','~',dm.DDL.FieldValues['RCONTENT'],'~~~~~~~~~~~~~~~',trim(dm.DDL.FieldValues['VODPROGRAMKEY']),'~~');
  										Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
                  end
								end
							else if not VarIsNull(dm.DDL.FieldValues['SRPROGRAMKEY']) then
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~','~~~~~~~~~~~~~~00',trim(dm.DDL.FieldValues['SRPROGRAMKEY']),'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
								end
							else
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_Ind']),'~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
								end;
								strSynInd:=Replace(trim(dm.DDL.FieldValues['SYSynopsis_Ind']), ',',';');
								strSynEng:=Replace(Trim(dm.DDL.FieldValues['SYSynopsis_Eng']),',',';');

                  AnsiSynEng:= ascii(trim(dm.DDL.FieldValues['SYSynopsis_Eng']));
                  //AnsiSynEng:= StripNonAlpha(AnsiSynEng);

                  AnsiSynInd:= ascii(trim(dm.DDL.FieldValues['SYSynopsis_Ind']));
                //AnsiSynInd:= StripNonAlpha(AnsiSynInd);


                {if AnsiContainsText(dm.DDL.FieldValues['SYSynopsis_Eng'], '&') then
                  begin
                    AnsiSynEng := StringReplace(dm.DDL.FieldValues['SYSynopsis_Eng'], '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);
                    AnsiSynInd := StringReplace(dm.DDL.FieldValues['SYSynopsis_Ind'], '&', '&amp;', [rfReplaceAll, rfIgnoreCase]);
                  end
                else if AnsiContainsText(dm.DDL.FieldValues['SYSynopsis_Eng'], '�') then
                  begin
                    AnsiSynEng := StringReplace(dm.DDL.FieldValues['SYSynopsis_Eng'], '�', '&#x00BF;', [rfReplaceAll, rfIgnoreCase]);
                    AnsiSynInd := StringReplace(dm.DDL.FieldValues['SYSynopsis_Ind'], '�', '&#x00BF;', [rfReplaceAll, rfIgnoreCase]);
                  end
                else if AnsiContainsText(dm.DDL.FieldValues['SYSynopsis_Eng'], '� ') then
                  begin
                    AnsiSynEng := StringReplace(dm.DDL.FieldValues['SYSynopsis_Eng'], '� ', '&#x00BA; ', [rfReplaceAll, rfIgnoreCase]);
                    AnsiSynInd := StringReplace(dm.DDL.FieldValues['SYSynopsis_Ind'], '� ', '&#x00BA; ', [rfReplaceAll, rfIgnoreCase]);
                  end                                   
                else
                  begin
                    AnsiSynEng := str_replace(
                      ['�','�','�','�','�','�','�','�','�','�','�','�'],
                      ['a','e','i','o','u','A','E','I','O','U','n','N'],
                      dm.DDL.FieldValues['SYSynopsis_Eng']
                    );
                    AnsiSynInd := str_replace(
                      ['�','�','�','�','�','�','�','�','�','�','�','�'],
                      ['a','e','i','o','u','A','E','I','O','U','n','N'],
                      dm.DDL.FieldValues['SYSynopsis_Ind']
                    );
                  end;  }

								Writeln(Sindo,dm.DDL.FieldValues['mchannel'],',',dm.DDL.FieldValues['stDate'],',',dm.DDL.FieldValues['stTime'],',',copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),',',strepg,',',strsynind,',',strsyneng);
                if NotEp = Null then
                  begin
									  Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<SynopsisEnglish>'+AnsiSynEng+'</SynopsisEnglish>',sLineBreak,AsciiTab,AsciiTab,'<SynopsisIndo>'+AnsiSynInd+'</SynopsisIndo>',sLineBreak,AsciiTab,AsciiTab,
                    '<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end
                else if NotSes = Null then
                  begin
									  Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<Episode>'+strEp+'</Episode>',sLineBreak,AsciiTab,AsciiTab,'<SynopsisEnglish>'+AnsiSynEng+'</SynopsisEnglish>',sLineBreak,AsciiTab,AsciiTab,'<SynopsisIndo>'+AnsiSynInd+'</SynopsisIndo>',sLineBreak,AsciiTab,AsciiTab,
                    '<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end
                else if not VarisNull(NotSes) then
                  begin
									  Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<Season>'+strSes+'</Season>',sLineBreak,AsciiTab,AsciiTab,'<Episode>'+strEp+'</Episode>',sLineBreak,AsciiTab,AsciiTab,'<SynopsisEnglish>'+AnsiSynEng+'</SynopsisEnglish>',sLineBreak,AsciiTab,AsciiTab,'<SynopsisIndo>'+AnsiSynInd+'</SynopsisIndo>',sLineBreak,AsciiTab,AsciiTab,
                    '<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end;
						end
					else
						begin
							if not VarIsNull(dm.DDL.FieldValues['IIMAGEID']) then
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',dm.DDL.FieldValues['IIMAGEID']);
								end
							// else if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
              else if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
								begin
									if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') or (VarIsNull(dm.DDL.FieldValues['VODGROUPKEY'])) then
										Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~',dm.DDL.FieldValues['VODTRAFFICKEY'],'~','8','~',dm.DDL.FieldValues['RCONTENT'],'~~~~~~~~~~~~~0~',dm.DDL.FieldValues['VODPROGRAMID'],'~',trim(dm.DDL.FieldValues['VODPROGRAMKEY']),'~~')
									else
										Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~',dm.DDL.FieldValues['VODTRAFFICKEY'],'~','0','~',dm.DDL.FieldValues['RCONTENT'],'~~~~~~~~~~~~~~~',trim(dm.DDL.FieldValues['VODPROGRAMKEY']),'~~');
										Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
								end
								else if not VarIsNull(dm.DDL.FieldValues['SRPROGRAMKEY']) then
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~','~~~~~~~~~~~~~~00',trim(dm.DDL.FieldValues['SRPROGRAMKEY']),'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
								end
							else
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
								end;
								Writeln(Sindo,dm.DDL.FieldValues['mchannel'],',',dm.DDL.FieldValues['stDate'],',',dm.DDL.FieldValues['stTime'],',',copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),',',strepg);
                if NotEp = Null then
                  begin
  									Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,
                    AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end
                else if NotSes = Null then
                  begin
  									Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<Episode>'+strEp+'</Episode>',sLineBreak,AsciiTab,AsciiTab,'<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,
                    AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end
                else if not VarisNull(NotSes) then
                  begin
  									Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,
                    '<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<Season>'+strSes+'</Season>',sLineBreak,AsciiTab,AsciiTab,'<Episode>'+strEp+'</Episode>',sLineBreak,AsciiTab,AsciiTab,'<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,
                    AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end;
						end;
			  //////////////////////////////////////////////Penambahan untuk dapat mengadopsi perubahan CA pada channel Vision 1
					if VarIsNull(dm.DDL.FieldValues['RCATEMPLATE']) then
						begin
					 /////////////////////// Xpush Channel /////////////////////////////////
							// if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
              if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
								begin
                if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                  begin
  									if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
  										Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~8~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  									else if dm.DDL.FieldValues['VODGROUPKEY'] = '12345' then
  										Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  									else if dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER' then
  										Writeln(SGI,'4~','1001','~','2','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing']); //~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~PlaceHolder~8~PlaceHolder~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00');
  								end
                end

							/////////////////////// Xpush Channel /////////////////////////////////
							else

							 if dm.DDL.FieldValues['mcsiserviceid'] = '2002' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'20~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2202' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'22~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '100' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'6~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '200' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'12~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2005' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'17~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~0~',catxt);
							{if not VarIsNull(dm.DDL.FieldValues['SRGROUPID']) then
								begin
									if dm.DDL.FieldValues['mcsiserviceid'] = '3002' then
										begin
											Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~2~');
										end
									else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~1~');
								end; }
                if not VarIsNull(dm.DDL.FieldValues['SRGROUPKEY']) then
                  begin
                    Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPKEY']),'~1~');
                  end;
								// if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
									begin
                    if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                      begin
  										  if (dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER') or (dm.DDL.FieldValues['MCSISERVICEID'] = '48') then
  										  	Write(SGI)
  										  else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['VODGROUPKEY']),'~2~');
                      end
									end;
									//Writeln(BB ,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~0~',catxt);
						end
					else
						begin
							strSQL := 'SELECT * FROM ';
							strSQL := strSQL + '(SELECT (Count(ccadescription) + 2) AS CountCA FROM (SELECT CCAdescription FROM m_caserviceid, t_catemplate WHERE TCA_CODE = '''+ dm.DDL.FieldValues['RCATEMPLATE'] + ''' AND TCA_NUMBER = ccaid)), ';
							strSQL := strSQL + '(SELECT ccadescription FROM m_caserviceid, t_catemplate WHERE TCA_CODE = '''+ dm.DDL.FieldValues['RCATEMPLATE'] + ''' AND TCA_NUMBER = ccaid) ';
							RecSet2(strSQL);
							catxtvis:='';
							i:=2;
							While not dm.DDL2.Eof do
								begin
									catxtvis:=catxtvis + IntToStr(i) + '~' + dm.DDL2.FieldValues['CCADescription'] + '~' ;
									i:=i+1;
									dm.DDL2.Next;
								end;
						/////////////////////// Xpush Channel /////////////////////////////////
							// if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
              if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
								begin
                if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                  begin
  									if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
  										Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~8~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  									else if dm.DDL.FieldValues['VODGROUPKEY'] = '12345' then
  										Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  									else if dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER' then
  										Writeln(SGI,'4~','1001','~','2','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing']); //~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~PlaceHolder~8~PlaceHolder~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00');
  								end
                end


						 /////////////////////// Xpush Channel /////////////////////////////////
							else
                if dm.DDL.FieldValues['mcsiserviceid'] = '2002' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'20~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2202' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'22~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2005' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'17~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '100' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'6~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '200' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'12~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis);
								{if not VarIsNull(dm.DDL.FieldValues['SRGROUPID']) then
									begin
										if dm.DDL.FieldValues['mcsiserviceid'] = '3002' then
											begin
												Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~2~');
											end
										else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~1~');
									end; }
                if not VarIsNull(dm.DDL.FieldValues['SRGROUPKEY']) then
                  begin
                    Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPKEY']),'~1~');
                  end;
								if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
									begin
										if (dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER') or (dm.DDL.FieldValues['MCSISERVICEID'] = '48') then
											Write(SGI)
										else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['VODGROUPKEY']),'~2~');
									end;
								//Writeln(BB ,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~0~',catxtvis);
						end;
			 //////////////////////////////////////////////// Akhir dari penambahan
				end
			else
				begin
					if not VarIsNull(dm.DDL.FieldValues['iimageid']) then
						begin
							Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
							Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',dm.DDL.FieldValues['IIMAGEID']);
						end
					else if not VarIsNull(dm.DDL.FieldValues['SRPROGRAMKEY']) then
						begin
							Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~','~~~~~~~~~~~~~~00',trim(dm.DDL.FieldValues['SRPROGRAMKEY']),'~');
							Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
						end
					// else if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
          else if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
						begin
							if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') or (VarIsNull(dm.DDL.FieldValues['VODGROUPKEY'])) then
								Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~',dm.DDL.FieldValues['VODTRAFFICKEY'],'~','8','~',dm.DDL.FieldValues['RCONTENT'],'~~~~~~~~~~~~~0~',dm.DDL.FieldValues['VODPROGRAMID'],'~',trim(dm.DDL.FieldValues['VODPROGRAMKEY']),'~~')
							else
								Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~',dm.DDL.FieldValues['VODTRAFFICKEY'],'~','0','~',dm.DDL.FieldValues['RCONTENT'],'~~~~~~~~~~~~~~~',trim(dm.DDL.FieldValues['VODPROGRAMKEY']),'~~');
								Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
						end
					else
						begin
							Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
							Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~ind~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~',' ~');
						end;
			 //////////////////////////////////////////////Penambahan untuk dapat mengadopsi perubahan CA pada channel Vision 1
						if VarIsNull(dm.DDL.FieldValues['RCATEMPLATE']) then
							begin
			/////////////////////// Start Xpush Channel /////////////////////////////////
								// if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
                if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
									begin
                  if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                    begin
  										if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
  											Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~8~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  										else if dm.DDL.FieldValues['VODGROUPKEY'] = '12345' then
  											Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  										else if dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER' then
  											Writeln(SGI,'4~','1001','~','2','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing']);  //~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~PlaceHolder~8~PlaceHolder~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00');
  									end
                  end
			/////////////////////// End Xpush Channel /////////////////////////////////
								else
                if dm.DDL.FieldValues['mcsiserviceid'] = '2002' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'20~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2202' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'22~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2005' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'17~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '100' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'6~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '200' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt,'12~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxt);

								{if not VarIsNull(dm.DDL.FieldValues['SRGROUPID']) then
									begin
										if dm.DDL.FieldValues['mcsiserviceid'] = '3002' then
											begin
												Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~2~');
											end
										else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~1~');
									end; }
                if not VarIsNull(dm.DDL.FieldValues['SRGROUPKEY']) then
                  begin
                    Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPKEY']),'~1~');
                  end;
								if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
									begin
										if (dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER') or (dm.DDL.FieldValues['MCSISERVICEID'] = '48') then
											Write(SGI)
										else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['VODGROUPKEY']),'~2~');
									end;

								//Writeln(BB ,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL.FieldValues['ca'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~0~',catxt);
							end
						else
							begin
								strSQL := 'SELECT * FROM ';
								strSQL := strSQL + '(SELECT (Count(ccadescription) + 2) AS CountCA FROM (SELECT CCAdescription FROM m_caserviceid, t_catemplate WHERE TCA_CODE = '''+ dm.DDL.FieldValues['RCATEMPLATE'] + ''' AND TCA_NUMBER = ccaid)), ';
								strSQL := strSQL + '(SELECT ccadescription FROM m_caserviceid, t_catemplate WHERE TCA_CODE = '''+ dm.DDL.FieldValues['RCATEMPLATE'] + ''' AND TCA_NUMBER = ccaid) ';
								RecSet2(StrSQL);
								catxtvis:='';
								i:=2;
								While not dm.DDL2.Eof do
									begin
										catxtvis:=catxtvis + IntToStr(i) + '~' + dm.DDL2.FieldValues['CCADescription'] + '~' ;
										i:=i+1;
										dm.DDL2.Next;
									end;
								/////////////////////// Xpush Channel /////////////////////////////////
								// if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
                if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
									begin
                  if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                    begin
  										if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
  											Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~8~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  										else if dm.DDL.FieldValues['VODGROUPKEY'] = '12345' then
  											Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~','12','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00')
  										else if dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER' then
  											Writeln(SGI,'4~','1001','~','2','~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing']); //~','2~PlaceHolder~3~PlaceHolder~4~PlaceHolder~5~PlaceHolder~6~PlaceHolder~','7~PlaceHolder~8~PlaceHolder~9~',dm.DDL.FieldValues['VODFED'],'~10~0~11~','00',dm.DDL.FieldValues['VODTIMEOFFSET'],'00');
  									end
                  end
								/////////////////////// Xpush Channel /////////////////////////////////
								else
                if dm.DDL.FieldValues['mcsiserviceid'] = '2002' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'20~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2202' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'22~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2005' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'17~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '100' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'6~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '200' then
                  begin
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'12~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis);

								{if not VarIsNull(dm.DDL.FieldValues['SRGROUPID']) then
									begin
										if dm.DDL.FieldValues['mcsiserviceid'] = '3002' then
											begin
												Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~2~');
											end
										else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPID']),'~1~');
									end; }
              if not VarIsNull(dm.DDL.FieldValues['SRGROUPKEY']) then
                begin
                  Writeln(SGI,'11~',trim(dm.DDL.FieldValues['SRGROUPKEY']),'~1~');
                end;
							if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
								begin
									if (dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER') or (dm.DDL.FieldValues['MCSISERVICEID'] = '48') then
										Write(SGI)
									else Writeln(SGI,'11~',trim(dm.DDL.FieldValues['VODGROUPKEY']),'~2~');
								end;
							//Writeln(BB ,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA'],'~0~',dm.DDL.FieldValues['RRATING'],'~1~0~',catxtvis);
					end;
					Writeln(Sindo,dm.DDL.FieldValues['mchannel'],',',dm.DDL.FieldValues['stDate'],',',dm.DDL.FieldValues['stTime'],',',copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),',',strepg);

                if NotEp = Null then
                  begin
  									Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,
                    AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end
                else if NotSes = Null then
                  begin
  									Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,'<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<Episode>'+strEp+'</Episode>',sLineBreak,AsciiTab,AsciiTab,'<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,
                    AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end
                else if not VarisNull(NotSes) then
                  begin
  									Writeln(XML,AsciiTab,'<record>',sLineBreak,AsciiTab,AsciiTab,'<ChannelNumber>'+strChnlNum+'</ChannelNumber>',sLineBreak,AsciiTab,AsciiTab,'<Channel>'+ansiChannel+'</Channel>',sLineBreak,AsciiTab,AsciiTab,'<TitleOri>'+afterXML+'</TitleOri>',sLineBreak,AsciiTab,AsciiTab,
                    '<Title>'+trimtitle+'</Title>',sLineBreak,AsciiTab,AsciiTab,'<Season>'+strSes+'</Season>',sLineBreak,AsciiTab,AsciiTab,'<Episode>'+strEp+'</Episode>',sLineBreak,AsciiTab,AsciiTab,'<StartDate_ddmmyyyy>'+dm.DDL.FieldValues['stDate']+'</StartDate_ddmmyyyy>',sLineBreak,AsciiTab,AsciiTab,'<StartTime>'+dm.DDL.FieldValues['stTimeXML']+'</StartTime>',sLineBreak,AsciiTab,AsciiTab,'<Duration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+'</Duration>',sLineBreak,AsciiTab,AsciiTab,'<Rating>'+strRating+'</Rating>',sLineBreak,AsciiTab,AsciiTab,'<Genre>'+dm.DDL.FieldValues['RGENRE']+'</Genre>',sLineBreak,AsciiTab,AsciiTab,'<SubGenre>'+dm.DDL.FieldValues['RSUBGENRE']+'</SubGenre>',sLineBreak,
                    AsciiTab,AsciiTab,'<Content>'+strContent+'</Content>',sLineBreak,AsciiTab,'</record>');
                  end;
				end;

      /////// Start XTI
			Writeln(XTI ,'<SiEvent>');
      Writeln(XTI , AsciiTab , '<displayDateTime>'+ FormatDateTime('yyyy/mm/dd',dm.DDL.FieldValues['rscheduledate']) + ' ' + FormatDateTime('hh:mm:ss',StrToTime(dm.DDL.FieldValues['stTimeXML']))  +':00</displayDateTime>');
      Writeln(XTI , AsciiTab , '<activationDateTime>'+ FormatDateTime('yyyy/mm/dd',dm.DDL.FieldValues['rscheduledate']) + ' ' + FormatDateTime('hh:mm:ss',StrToTime(dm.DDL.FieldValues['stTimeXML']))  +':00</activationDateTime>');
      Writeln(XTI , AsciiTab , '<siTrafficKey>'+ dm.DDL.FieldValues['RKEY_HEX'] +'</siTrafficKey>');
      Writeln(XTI , AsciiTab , '<detailKey>'+ dm.DDL.FieldValues['RKEY_HEX'] +'</detailKey>');
      Writeln(XTI , AsciiTab , '<displayDuration>'+copy(dm.DDL.FieldValues['rduration'],1,2),':',copy(dm.DDL.FieldValues['rduration'],3,2),':',copy(dm.DDL.FieldValues['rduration'],5,2)+':00</displayDuration>');
      if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then Writeln(XTI , AsciiTab , '<oppvPurchaseCode>'+AnsiRightStr(dm.DDL.FieldValues['VODCAPRODUCTID'], 3) +'</oppvPurchaseCode>');
      Writeln(XTI , AsciiTab , '<SiEventDetail>');
            Writeln(XTI,AsciiTab,AsciiTab,'<parentalRatingId>'+strRating+'</parentalRatingId>');
            Writeln(XTI,AsciiTab,AsciiTab,'<genreId>'+dm.DDL.FieldValues['RGENRE']+'</genreId>');
            Writeln(XTI,AsciiTab,AsciiTab,'<subGenreId>'+dm.DDL.FieldValues['RSUBGENRE']+'</subGenreId>');
            if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
              Writeln(XTI,AsciiTab,AsciiTab,'<broadcasterDetail-1>8</broadcasterDetail-1>')
            else Writeln(XTI,AsciiTab,AsciiTab,'<broadcasterDetail-1>'+IntToStr(dm.DDL.FieldValues['MUSERNIBBLE1'])+'</broadcasterDetail-1>');
            Writeln(XTI,AsciiTab,AsciiTab,'<broadcasterDetail-2>'+IntToStr(dm.DDL.FieldValues['RCONTENT'])+'</broadcasterDetail-2>');

            if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
              begin
                if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
                  begin
                     Writeln(XTI,AsciiTab,AsciiTab,'<programKey>'+IntToStr(dm.DDL.FieldValues['VODPROGRAMKEY'])+'</programKey>');
                  end;
              end;
            Writeln(XTI,AsciiTab,AsciiTab,'<SiEventDescription>');
            Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<displayLanguage>ind</displayLanguage>');
	          Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<eventName>'+afterXML+'</eventName>');
            if not VarIsNull(dm.DDL.FieldValues['sysynopsis_ind']) then
              begin
                Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<eventDescription>'+AnsiSynInd+'</eventDescription>');
              end
            else Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<eventDescription> </eventDescription>');
            Writeln(XTI,AsciiTab,AsciiTab,'</SiEventDescription>');
            Writeln(XTI,AsciiTab,AsciiTab,'<SiEventDescription>');
            Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<displayLanguage>eng</displayLanguage>');
	          Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<eventName>'+afterXML+'</eventName>');
            if not VarIsNull(dm.DDL.FieldValues['sysynopsis_eng']) then
              begin
                Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<eventDescription>'+AnsiSynEng+'</eventDescription>');
              end
            else Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<eventDescription> </eventDescription>');
            Writeln(XTI,AsciiTab,AsciiTab,'</SiEventDescription>');
            if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
              begin
                if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
                  begin
                    Writeln(XTI,AsciiTab,AsciiTab,'<SiProgramGroupLink> ');
                    Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<detailKey>'+dm.DDL.FieldValues['RKEY_HEX']+'</detailKey>');
                    Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<groupKey>'+IntToStr(dm.DDL.FieldValues['VODGROUPKEY'])+'</groupKey>');
                    Writeln(XTI,AsciiTab,AsciiTab,AsciiTab,'<groupType>Push</groupType>');
                    Writeln(XTI,AsciiTab,AsciiTab,'</SiProgramGroupLink> ');
                  end;
              end;
      Writeln(XTI , AsciiTab , '</SiEventDetail>');
      Writeln(XTI , AsciiTab , '<CaRequest>');
      Writeln(XTI,AsciiTab,AsciiTab,'<caRequestKey>'+dm.DDL.FieldValues['RKEY_HEX']+'</caRequestKey>');
          if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
            begin
            if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
              begin
                    if dm.DDL.FieldValues['VODEPGTITLE'] = 'FILLER' then
                      begin
                         Writeln(XTI,AsciiTab,AsciiTab,'<caTemplateId>1001</caTemplateId>');
                         Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>0</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>4</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                         Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>1</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>0</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                      end
                    else if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') then
                      begin
                        Writeln(XTI,AsciiTab,AsciiTab,'<caTemplateId>'+IntToStr(dm.DDL.FieldValues['mcsiserviceid'])+'</caTemplateId>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>0</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>'+inttostr(dm.DDL.FieldValues['RRATING'])+'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>1</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>0</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>2</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>PlaceHolder</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>3</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>PlaceHolder</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>4</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>PlaceHolder</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>5</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>PlaceHolder</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>6</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>PlaceHolder</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>7</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>',dm.DDL.FieldValues['VODCAPRODUCTID'],'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>8</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>',dm.DDL.FieldValues['VODCAPRODUCTID'],'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>9</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>',dm.DDL.FieldValues['VODFED'],'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>10</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>0</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');

                        mjdfloat := datetimetojuliandate(strtodatetime('5/24/1968 12:'+copy(dm.DDL.FieldValues['VODTIMEOFFSET'],3,2)+':'+copy(dm.DDL.FieldValues['VODTIMEOFFSET'],5,2)+' AM'));
                        bfloat := mjdfloat - 2440000.5;
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>11</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>',formatfloat('0.######0', bfloat),'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                      end
                    else if (dm.DDL.FieldValues['VODGROUPKEY'] = '12345') then
                      begin
                        Writeln(XTI,AsciiTab,AsciiTab,'<caTemplateId>'+IntToStr(dm.DDL.FieldValues['mcsiserviceid'])+'</caTemplateId>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>0</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>'+inttostr(dm.DDL.FieldValues['RRATING'])+'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>1</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>0</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,catxtxti);
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>9</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>'+inttostr(dm.DDL.FieldValues['VODFED'])+'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>10</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>0</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');

                        mjdfloat := datetimetojuliandate(strtodatetime('5/24/1968 12:'+copy(dm.DDL.FieldValues['VODTIMEOFFSET'],3,2)+':'+copy(dm.DDL.FieldValues['VODTIMEOFFSET'],5,2)+' AM'));
                        bfloat := mjdfloat - 2440000.5;
                        Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>11</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>',formatfloat('0.######0', bfloat),'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
                      end;
                  end
            end
          else
            begin
              Writeln(XTI,AsciiTab,AsciiTab,'<caTemplateId>'+IntToStr(dm.DDL.FieldValues['mcsiserviceid'])+'</caTemplateId>');
              Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>0</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>'+inttostr(dm.DDL.FieldValues['RRATING'])+'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');
              Writeln(XTI,AsciiTab,AsciiTab,'<CaRequestParameter>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterNumber>1</parameterNumber>',sLineBreak,AsciiTab,AsciiTab,AsciiTab,'<parameterValue>'+dm.DDL.FieldValues['mstb_pairing']+'</parameterValue>',sLineBreak,AsciiTab,AsciiTab,'</CaRequestParameter>');

          tcadate := FormatDateTime('mm/dd/yyyy',dm.DDL.FieldValues['rscheduledate']) + ' ' + FormatDateTime('hh:mm:ss',StrToTime(dm.DDL.FieldValues['stTimeXML']));

          strSQL := 'Select TCA_CODE From T_CATEMPLATE ';
          strSQL := strSQL + 'WHERE TCA_CODE =(';
          strSQL := strSQL + 'select RCATEMPLATE from m_readxl ';
          strSQL := strSQL + 'where rchannel = '''+ frmRead.ngReadXL.Cells[1,1] +''' ' ;
          strSQL := strSQL + 'and rscheduledate = to_date(''' + tcadate +  ''',''mm/dd/yyyy hh24:mi:ss''))';
          RecSet2(strSQL);

          if varisnull(dm.DDL2.FieldValues['TCA_CODE']) then
            tca := ''
          else tca := dm.DDL2.FieldValues['TCA_CODE'];

          if tca <> '' then
            begin
              strSQL := 'SELECT CCADescription FROM T_CATEMPLATE, M_CASERVICEID WHERE CCAID = tca_number AND tca_code = '''+ tca +''' ';
              RecSet2(strSQL);
        			catcaxti:='';
              xtca:=2;
        			While not dm.DDL2.Eof do
        				begin
                  catcaxti:=catcaxti + AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>'+IntToStr(xtca)+'</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL2.FieldValues['CCADescription']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak;
                  xtca:=xtca+1;
        					dm.DDL2.Next;
        				end;
              Writeln(XTI,catcaxti);
            end
          else Writeln(XTI,catxtxti);
                if dm.DDL.FieldValues['mcsiserviceid'] = '2002' then
                  begin
                     Writeln(XTI,AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>20</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL.FieldValues['CBNUMBER']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2202' then
                  begin
                     Writeln(XTI,AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>22</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL.FieldValues['CBNUMBER']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '2005' then
                  begin
                     Writeln(XTI,AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>17</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL.FieldValues['CBNUMBER']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak);
                     Writeln(SGI,'4~',dm.DDL.FieldValues['mcsiserviceid'],'~',dm.DDL2.FieldValues['CountCA']+1,'~0~',dm.DDL.FieldValues['RRATING'],'~1~',dm.DDL.FieldValues['mstb_pairing'],'~',catxtvis,'17~',dm.DDL.FieldValues['CBNUMBER']);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '100' then
                  begin
                     Writeln(XTI,AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>6</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL.FieldValues['CBNUMBER']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak);
                  end
                else if dm.DDL.FieldValues['mcsiserviceid'] = '200' then
                  begin
                     Writeln(XTI,AsciiTab+ AsciiTab+'<CaRequestParameter>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterNumber>12</parameterNumber>'+sLineBreak+AsciiTab+AsciiTab+AsciiTab+'<parameterValue>'+dm.DDL.FieldValues['CBNUMBER']+'</parameterValue>'+sLineBreak+AsciiTab+AsciiTab+'</CaRequestParameter>'+sLineBreak);
                  end;

            end;

      Writeln(XTI , AsciiTab , '</CaRequest>');
      Writeln(XTI ,'</SiEvent>');

      //// stop XTI
				dm.DDL.Next;
		end;
     //////////////////////////////////////////////// Akhir dari penambahan
		Writeln(SGI,'1~',dm.DDL.FieldValues['MSGINAME'],'~',dm.DDL.FieldValues['EventStartDate'],'~00000000~24000000~eng~1~0~');
		Writeln(BB ,'1~',dm.DDL.FieldValues['mchannel'],'~',dm.DDL.FieldValues['EventStartDate'],'~00000000~24000000~eng~1~0~');

		if strCATemplate = '1001' then
			begin
				strSQL:='SELECT 2 as ca, mcsiserviceid, mchannel, rscheduledate, REPG_TITLE, RDURATION, RRATING, RGENRE, RSUBGENRE, RCONTENT, EventStartDate, EventStartTime, ';
			end
		else
			begin
				strSQL:='SELECT ca, mcsiserviceid, mchannel, rscheduledate, REPG_TITLE, RDURATION, RRATING, RGENRE, RSUBGENRE, RCONTENT, EventStartDate, EventStartTime, ';
			end;
      strSQL:=strSQL + 'EventStartTimegmt, EventStartDategmt, SYSynopsis_Ind, SYSynopsis_Eng, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, ';
      strSQL:=strSQL + 'VODPROGRAMID, VODTRAFFICKEY, VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, msginame, MUSERNIBBLE1, RCATEMPLATE, mSYNOPSIS_STATUS, IIMAGEID ';
			strSQL:=strSQL + '	FROM ( SELECT * ';
			strSQL:=strSQL + '		FROM ( ';
			strSQL:=strSQL + '			SELECT mc.mcsiserviceid, mc.mchannel, mr.rscheduledate, mr.REPG_TITLE, mr.RDURATION, mr.RRATING, ';
			strSQL:=strSQL + '			mr.RGENRE, mr.RSUBGENRE, mr.RCONTENT, to_char(mr.rscheduledate,''ddmmyyyy'') AS EventStartDate, ';
			strSQL:=strSQL + '			to_char(mr.rscheduledate,''hh24miss'') AS EventStartTime, to_char(mr.rscheduledategmt,''hh24miss'') AS EventStartTimegmt, ';
			strSQL:=strSQL + '			to_char(mr.rscheduledategmt,''ddmmyyyy'') AS EventStartDategmt, msginame, MUSERNIBBLE1, mr.RCATEMPLATE, REPG_TITLE_ORI, mSYNOPSIS_STATUS ';
			strSQL:=strSQL + '			FROM m_channel mc, m_readxl mr ';
			strSQL:=strSQL + '			WHERE mc.mchannel = ''' + frmRead.ngReadXL.Cells[1,1] + ''' ';
			strSQL:=strSQL + '			AND mr.rchannel = mc.mchannel ';
			strSQL:=strSQL + '			AND mr.rscheduledate >= to_date(''' + frmExport.ComboBox1.Items.Strings[ii] +  ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
			strSQL:=strSQL + '			AND mr.rscheduledate <= to_date(''' + frmExport.ComboBox1.Items.Strings[ii] +  ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
			strSQL:=strSQL + '			)aaa, ';
			strSQL:=strSQL + '			(SELECT count(mca.capackage)+2 as ca FROM m_ca_package mca WHERE mca.cachannel=''' + frmRead.ngReadXL.Cells[1,1] + ''' ) bbb  ) XXX, ';
			strSQL:=strSQL + '			( SELECT syEPG_TITLE, SYSynopsis_Ind, SYSynopsis_Eng, sycategory ';
			strSQL:=strSQL + '			FROM M_Synopsis ) YYY, (SELECT * from m_image ) ZZZ, (select * from M_VOD) WWW ';
			strSQL:=strSQL + '			WHERE REPG_TITLE_ORI = syEPG_TITLE(+) AND RGENRE = sycategory(+) AND REPG_TITLE_ORI = iepg_ori(+) AND mchannel=ichannel (+) AND REPG_TITLE_ORI=VODEPGTITLE (+) ORDER BY MChannel, RScheduleDate ';
			RecSet(strSQL);
			While not dm.DDL.Eof do
				begin
					strRating := dm.DDL.FieldValues['RRating'];
					if not VarIsNull(dm.DDL.FieldValues['sysynopsis_eng']) then
						begin
							if dm.DDL.FieldValues['MSYNOPSIS_STATUS'] = 'Y' then
								begin
									if not VarIsNull(dm.DDL.FieldValues['IIMAGEID']) then
										begin
											Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
											Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,dm.DDL.FieldValues['IIMAGEID']);
										end
									// else if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
                  else if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
										begin
											if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') or (VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']))  then
												Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~','8','~',dm.DDL.FieldValues['RCONTENT'],'~')
											else
												Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~','0','~',dm.DDL.FieldValues['RCONTENT'],'~');
												Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,' ~');
										end										
									else
										begin
											Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
											Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,' ~');
										end;
								end
							else
								begin
									if not VarIsNull(dm.DDL.FieldValues['IIMAGEID']) then
										begin
											Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
											Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,dm.DDL.FieldValues['IIMAGEID']);
										end
									// else if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
                  else if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
										begin
											if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') or (VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']))  then
												Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~','8','~',dm.DDL.FieldValues['RCONTENT'],'~')
											else
												Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~','0','~',dm.DDL.FieldValues['RCONTENT'],'~');
												Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~',trim(dm.DDL.FieldValues['SYSynopsis_ENG']),'~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,' ~');
										end										
									else
										begin
											Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
											Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',strRating,'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,' ~');
										end;
								end;
						end
					else
						begin
							if not VarIsNull(dm.DDL.FieldValues['IIMAGEID']) then
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',dm.DDL.FieldValues['RRATING'],'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',dm.DDL.FieldValues['RRATING'],'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,dm.DDL.FieldValues['IIMAGEID']);
								end
							// else if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
              else if (dm.DDL.FieldValues['mcsiserviceid'] = '39') or (dm.DDL.FieldValues['mcsiserviceid'] = '48') then
								begin
									if (dm.DDL.FieldValues['VODGROUPKEY'] = '12346') or (VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']))  then
										Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',dm.DDL.FieldValues['RRATING'],'~~~~','8','~',dm.DDL.FieldValues['RCONTENT'],'~')
									else
										Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',dm.DDL.FieldValues['RRATING'],'~~~~','0','~',dm.DDL.FieldValues['RCONTENT'],'~');
										Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~0~',dm.DDL.FieldValues['RRATING'],'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,' ~');
								end								
							else
								begin
									Writeln(SGI,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',dm.DDL.FieldValues['RRATING'],'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');
									Writeln(BB ,'2~',dm.DDL.FieldValues['EventStartDate'],'~',dm.DDL.FieldValues['EventStartTime'],'00~',trim(dm.DDL.FieldValues['RDURATION']),'00~',trim(dm.DDL.FieldValues['REPG_TITLE']),'~ ~eng~0~~',dm.DDL.FieldValues['RGENRE'],'~',dm.DDL.FieldValues['RSUBGENRE'],'~',dm.DDL.FieldValues['RRATING'],'~~~~',dm.DDL.FieldValues['MUSERNIBBLE1'],'~',dm.DDL.FieldValues['RCONTENT'],'~');//,dm.DDL.FieldValues['IIMAGEID']);
								end;
						end;
						dm.DDL.Next;
				end;
		end;
  Writeln(XTI, '</SiEventSchedule>');
  Writeln(XTI, '</BasicImport>');
	Writeln(XML, '</data-set>');
  Writeln(actLOG,'[', FormatDateTime('c',today),'] ', ' Export ', frmRead.ngReadXL.Cells[1,1] ,' ', date1 ,'-', date2, ' berhasil');
  Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' Export ', frmRead.ngReadXL.Cells[1,1] ,' ', date1 ,'-', date2, ' berhasil');

	CloseFile(SGI);
	CloseFile(BB);
	CloseFile(Sindo);
	CloseFile(XML);
  CloseFile(XTI);
  CloseFile(actLOG);
  CloseFile(actLOGLocal);
	strSQL := 'DELETE FROM TEMP_READXL ';
	strSQL := strSQL + ' WHERE TRCHANNEL = ''' + dm.DDL.FieldValues['mchannel'] + ''' ';
	RecExc(strSQL);
  MessageDlg('Create File Finished!'+sLineBreak+'-> SGI file at C:\SGI\SGI_NDS'+sLineBreak+'-> BB file at C:\SGI\SGI_BB'+sLineBreak+'-> SINDO file at C:\SGI\SGI_SINDO'+sLineBreak+'-> XML file at C:\SGI\SGI_XML'+sLineBreak+'-> NDS XTI file at C:\SGI\SGI_XTI', mtInformation, [mbOK], 0);
	//ShowMessage('Create File Finished!'+sLineBreak+'-> SGI file at C:\SGI\SGI_NDS'+sLineBreak+'-> BB file at C:\SGI\SGI_BB'+sLineBreak+'-> SINDO file at C:\SGI\SGI_SINDO'+sLineBreak+'-> XML file at C:\SGI\SGI_XML'+sLineBreak+'-> NDS XTI file at C:\SGI\SGI_XTI');
	frmExport.Close;

except
     on E : Exception do
     begin
       AssignFile(ERR , 'C:\SGI\LOG_ERROR\FormExport_' + frmRead.ngReadXL.Cells[1,1]+'_'+strUser+'.log');
       Rewrite(ERR);

       Writeln(ERR , encode64('coba kita liat apa yang dilakuin sama si '+ strUserName +', dia lagi coba export channel '+frmRead.ngReadXL.Cells[1,1] +' trus keluar error, detailnya -> Err.Class: '+ E.ClassName+ ', pesan errornya gini: '+ E.Message) );
       CloseFile(ERR);
       showmessage('Maaf, terdapat kesalahan dalam aplikasi, mohon kirimkan Error Log di :'+sLineBreak+''+sLineBreak+'C:\SGI\LOG_ERROR\FormExport_' + frmRead.ngReadXL.Cells[1,1]+'_'+strUser+'.log'+sLineBreak+''+sLineBreak+'kirimkan sebagai attachment ke mailto:broadcastapp@indovision.tv '+sLineBreak+'agar dapat diteliti dan diperbaiki apabila terjadi kesalahan pada aplikasi ini.'+sLineBreak+''+sLineBreak+'Terima Kasih' );
       //ShowMessage('Exception class name = '+E.ClassName);
       //ShowMessage('Exception message = '+ E.Message );
       Writeln(actLOG,'[', FormatDateTime('c',today),'] ', ' Export ', frmRead.ngReadXL.Cells[1,1] , date1 ,'-', date2, 'gagal: ',E.Message );
       CloseFile(actLOG);
       Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' Export ', frmRead.ngReadXL.Cells[1,1] , date1 ,'-', date2, 'gagal: ',E.Message );
       CloseFile(actLOGLocal);
       frmExport.Close;
     end;
end;
end;

procedure TfrmExport.cbDateEndSelect(Sender: TObject);
var
  item : TStrings;
begin
 ComboBox1.Items.Clear;
 strSQL := 'SELECT to_date(Date_Schedule,''mm/dd/yyyy'') AS Dates FROM ( SELECT distinct to_char(rscheduledate,''mm/dd/yyyy'') AS Date_Schedule ';
 strSQL := strSQL + 'FROM m_readxl ';
 strSQL := strSQL + 'WHERE rchannel = ''' + frmRead.StringGrid1.Cells[3,1] + ''' ';
 strSQL := strSQL + 'AND rscheduledate >= to_date(''' + cbDateStart.Text +  ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
 strSQL := strSQL + 'AND rscheduledate <= to_date(''' + cbDateEnd.Text +  ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
 strSQL := strSQL + 'ORDER by Date_Schedule ) ORDER by Dates';
 RecSet(strSQL);
 Item:=ComboBox1.Items.Create;
 While not dm.DDL.Eof do
  begin
   item.Add(dm.DDL.FieldValues['Dates']);
   dm.DDL.Next;
  end;
 dm.DDL.Last;
 Date2 := FormatDateTime('mmddyy',dm.DDL.FieldValues['DATES']);
end;

end.
