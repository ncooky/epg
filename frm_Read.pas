unit frm_Read;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, Menus, NxColumns, NxColumnClasses, ExtCtrls, jpeg,
  ComCtrls, shellapi, StrUtils, INIFiles;

const
WM_ICONTRAY = WM_USER + 1;  

type TDateTime = type Double;

type
  TfrmRead = class(TForm)

    StringGrid1: TStringGrid;
    OpenDialog1: TOpenDialog;
    MainMenu1: TMainMenu;
    Master1: TMenuItem;
    Synopsis1: TMenuItem;
    Exit1: TMenuItem;
    EPG1: TMenuItem;
    INSERT1: TMenuItem;
    CAPackage1: TMenuItem;
    SynopsisManual: TMenuItem;
    Menu1: TMenuItem;
    LogOut1: TMenuItem;
    User1: TMenuItem;
    Synopsis2: TMenuItem;
    Fromfile1: TMenuItem;
    CaServiceID1: TMenuItem;
    Channel1: TMenuItem;
    Vision11: TMenuItem;
    Clear1: TMenuItem;
    ScrollBox1: TScrollBox;
    ngReadXL: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn14: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    NxTextColumn9: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    NxTextColumn11: TNxTextColumn;
    NxTextColumn12: TNxTextColumn;
    NxTextColumn13: TNxTextColumn;
    ScrollBox2: TScrollBox;
    Image1: TImage;
    Button1x: TButton;
    Button2x: TButton;
    pbRead: TProgressBar;
    Check1: TMenuItem;
    Image2: TMenuItem;
    FormFile1: TMenuItem;
    View1: TMenuItem;
    NxTextColumn15: TNxTextColumn;
    NxTextColumn16: TNxTextColumn;
    SeriesLink1: TMenuItem;
    SeriesLink2: TMenuItem;
    NxTextColumn17: TNxTextColumn;
    NxTextColumn18: TNxTextColumn;
    NxTextColumn19: TNxTextColumn;
    NxTextColumn20: TNxTextColumn;
    NxTextColumn21: TNxTextColumn;
    Button1: TLabel;
    Button2: TLabel;
    Help1: TMenuItem;
    About1: TMenuItem;
    DecryptLog1: TMenuItem;
    MJDConversion1: TMenuItem;
    StatusBar1: TStatusBar;
    DatabaseMaintainer1: TMenuItem;
    setup: TMenuItem;
    DBChooser: TMenuItem;
    procedure Button1xClick(Sender: TObject);
    procedure Button2xClick(Sender: TObject);
    procedure Synopsis1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure EPG1Click(Sender: TObject);
    procedure CAPackage1Click(Sender: TObject);
    procedure SynopsisManualClick(Sender: TObject);
    procedure LogOut1Click(Sender: TObject);
    procedure User1Click(Sender: TObject);
    procedure Fromfile1Click(Sender: TObject);
    procedure CaServiceID1Click(Sender: TObject);
    procedure Channel1Click(Sender: TObject);
    procedure Vision11Click(Sender: TObject);
    procedure Clear1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Check1Click(Sender: TObject);
    procedure FormFile1Click(Sender: TObject);
    procedure View1Click(Sender: TObject);
    procedure SeriesLink1Click(Sender: TObject);
    procedure SeriesLink2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DecryptLog1Click(Sender: TObject);
    procedure MJDConversion1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure DatabaseMaintainer1Click(Sender: TObject);
    procedure DBChooserClick(Sender: TObject);
  private
    TrayIconData:TNotifyIconData;
    { Private declarations }
  public
    { Public declarations }
  end;



var
  frmRead: TfrmRead;
  function fncStrToDateTime(strTanggal:string; strTime:string):TDateTime;
  function Replace(Dest, SubStr, Str: string): string;
  function valstrtodatetime(valstr : double; strDate:string) : string;
  function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;

implementation

{$R *.dfm}

uses
  ComObj, frm_dm, DateUtils, frm_export, frm_Synopsis, frm_EPG, frm_ExEPG,
  frm_InsertCA, frm_SynopsisManual, frm_Login, frm_User, frm_SynopsisXL,
  frm_CAServiceID, frm_Channel, frm_Vis1, DB, frm_Check, frm_Image,
  frm_VOD, frm_SeriesLink, frm_About, frm_Decode, MJDConverter,
  ProgressBar, frm_maintaindb;

  const
Codes64 = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+/';

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

function ascii
  (Const Str: AnsiString): String;
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

  else if AnsiContainsText(Str, ',')
  then ascii := StringReplace(Str, ',', '&#44;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '-')
  then ascii := StringReplace(Str, '-', '&#45;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '.')
  then ascii := StringReplace(Str, '.', '&#46;', [rfReplaceAll, rfIgnoreCase])

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

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#174;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�-')
  then ascii := StringReplace(Str, '�-', '&#174;&#45;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#175;', [rfReplaceAll, rfIgnoreCase])

  else if AnsiContainsText(Str, '�')
  then ascii := StringReplace(Str, '�', '&#176;', [rfReplaceAll, rfIgnoreCase])

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

  else ascii := Str;

end;

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT SYID FROM M_SYNOPSIS ORDER BY SYID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   strAngka := dm.DDL.FieldValues['SYID'] + 1;
  end;
 fncangka:=IntToStr(strAngka);
end;

function fncangka2():string;
var
 strAngka1 : integer;
begin
 strSQL := 'SELECT max(RID) as RID FROM M_READXL ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka1 := 1;
  end
 else
  begin
   strAngka1 := dm.DDL.FieldValues['RID'] + 1;
  end;
 fncangka2:=IntToStr(strAngka1);
end;

function  GetAppVersion:string;
   var
    Size, Size2: DWord;
    Pt, Pt2: Pointer;
   begin
     Size := GetFileVersionInfoSize(PChar (ParamStr (0)), Size2);
     if Size > 0 then
     begin
       GetMem (Pt, Size);
       try
          GetFileVersionInfo (PChar (ParamStr (0)), 0, Size, Pt);
          VerQueryValue (Pt, '\', Pt2, Size2);
          with TVSFixedFileInfo (Pt2^) do
          begin
            Result:= ' Version '+
                     IntToStr (HiWord (dwFileVersionMS)) + '.' +
                     IntToStr (LoWord (dwFileVersionMS)) + '.' +
                     IntToStr (HiWord (dwFileVersionLS)) + '.' +
                     IntToStr (LoWord (dwFileVersionLS)) + ' @Mid 2016';
         end;
       finally
         FreeMem (Pt);
       end;
     end;
   end;
   
function fncStrToDateTime(strTanggal:string; strTime:string):TDateTime;
var
  dtSetting : TFormatSettings;
  dtTanggal : TDateTime;
Begin
  GetLocaleFormatSettings(GetUserDefaultLCID, dtSetting);
  Screen.Cursor := crHourGlass;
  dtSetting.DateSeparator := '/';
  dtSetting.TimeSeparator := ':';
  dtSetting.ShortDateFormat := 'mm/dd/yyyy';
  dtSetting.ShortTimeFormat := 'HH:mm:ss';
End;

function Replace(Dest, SubStr, Str: string): string;
var
  Position: Integer;
begin
 Position:=Pos(SubStr, Dest);
  while Position<>0 do
  begin
   Delete(Dest, Position, Length(SubStr));
   Insert(Str, Dest, Position);
   Result:=Dest;
   Position:=Pos(SubStr, Dest);
  end;
 Result:=Dest;
end;

function valstrtodatetime(valstr : double; strDate:string) : string;
var
  dt : TdateTime;
  strTanggal : String;
begin
  dt := valstr;
  strTanggal := copy(strDate, 5, 2) + '/' + copy(strDate, 7, 2) + '/' + copy(strDate, 1, 4);
  strTanggal := strTanggal + ' ' + Formatdatetime('HH:mm:ss',dt);
  result := strTanggal;
end;

function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Open the Workbook
    XLApp.Workbooks.Open(AXLSFile);

    // Sheet := XLApp.Workbooks[1].WorkSheets[1];
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];

    // In order to know the dimension of the WorkSheet, i.e the number of rows
    // and the number of columns, we activate the last non-empty cell of it

    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;

    // Set Stringgrid's row &col dimensions.

    AGrid.RowCount := x;
    AGrid.ColCount := y;
    frmRead.pbRead.Max:=x;
    frmRead.pbRead.Min:=0;
    frmRead.pbRead.Visible:=True;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant

    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].value;
    //  Define the loop for filling in the TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
       //begin
      //AGrid.Cells[(r - 1), (k - 1)] := '';
      AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
     //end;
      Inc(k, 1);
      AGrid.RowCount := k + 1;
      frmRead.pbRead.Position:=k;
    until k > x;
    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;

  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
    end;
  end;
end;

procedure TfrmRead.Button1xClick(Sender: TObject);
var
 i, rating, angka, climax, xxx, xx, y : integer;
 DateFloat, DateFloat1, climaxDateFloat, awalDateFloat : Double;
 tanggalawal, tanggalclimax, Durstring, DateString, awaldatestring, climaxdatestring, ReplaceString, GroupKeyKosong, ProgramKeyKosong, kosong, ind_kosong, eng_kosong, rid, tca, repascii :String;
 VODCaProductKosong, VodStartDateKosong, VodEndDateKosong, VODStatusKosong : String;



begin
  if OpenDialog1.Execute then
    begin
     //i := 1;
     Screen.Cursor:=crHourGlass;
     for I := 0 to StringGrid1.RowCount - 1 do StringGrid1.Rows[I].Clear();
     if Xls_To_StringGrid(StringGrid1, OpenDialog1.FileName) then
        begin
         ngReadXL.ClearRows;
         strSQL := 'SELECT mchannel FROM M_CHANNEL WHERE mchannel = ''' + StringGrid1.Cells[3,1] + ''' ';
         RecSet(strSQL);
         if not dm.DDL.Eof then
          begin
           strSQL := 'SELECT distinct TRCHANNEL FROM TEMP_READXL WHERE TRCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
           RecSet(strSQL);
           if dm.DDL.Eof then
            begin
             i := 1;
             pbRead.Max:= StringGrid1.RowCount-2;
             pbRead.Min:= 0;
             pbRead.Visible := True;

             climax:= StringGrid1.RowCount-4;
             climaxDateFloat:=strtofloat(StringGrid1.Cells[5,climax]);
             awalDateFloat:= strtofloat(StringGrid1.Cells[5,1]);
             climaxdatestring:=valstrtodatetime(climaxDateFloat, StringGrid1.Cells[4,climax]);
             awaldatestring:=valstrtodatetime(awalDateFloat, StringGrid1.Cells[4,1]);

             tanggalawal := LeftStr(awaldatestring, 10);
             tanggalclimax := LeftStr(climaxdatestring, 10);

             strSQL := 'DELETE FROM M_READXL ';
             strSQL := strSQL + ' WHERE RCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
             strSQL := strSQL + ' AND RSCHEDULEDATE >= TO_Date(''' + tanggalawal + ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
             strSQL := strSQL + ' AND RSCHEDULEDATE <= TO_Date(''' + tanggalclimax + ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
             strSQL := strSQL + ' AND RSCHEDULEDATEGMT >= TO_Date(''' + tanggalawal + ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'')-0.29167 ';
             strSQL := strSQL + ' AND RSCHEDULEDATEGMT <= TO_Date(''' + tanggalclimax + ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'')-0.29167 ';
             RecExc2(strSQL);



             //angka := StrToInt(fncangka2);
             while StringGrid1.Cells[0,i]<>'<eof>' do
               Begin
                if StringGrid1.Cells[10,i] <> '' then
                  begin
                   strSQL := 'Select TCA_CODE From T_CATEMPLATE WHERE TCA_CODE =''' +  StringGrid1.Cells[10,i] + ''' ';
                   RecSet(strSQL);

                   if varisnull(dm.DDL.FieldValues['TCA_CODE']) then
                      tca := ''
                   else tca := dm.DDL.FieldValues['TCA_CODE'];
                  end
                else tca := '';


                DateFloat:=strtofloat(StringGrid1.Cells[5,i]);
                DateFloat1:=strtofloat(StringGrid1.Cells[6,i]);
                DateString:=valstrtodatetime(DateFloat, StringGrid1.Cells[4,i]);
                Durstring:=valstrtodatetime(DateFloat1, StringGrid1.Cells[4,i]);
                ReplaceString:=Replace(StringGrid1.Cells[2,i], '''', '`');
                ReplaceString:=Replace(ReplaceString, '`', '''''');


                strSQL := 'INSERT INTO SGI.TEMP_READXL ( ';
                strSQL := strSQL + 'tRID, tRCHANNEL, tRSCHEDULEDATE, tRSCHEDULEDATEGMT, ';
                strSQL := strSQL + 'tREPG_TITLE, tRDURATION, tRRATING, ';
                if tca <> '' then
                  begin
                   strSQL := strSQL + 'tRGENRE, tRSUBGENRE, tRCONTENT, tRCATEMPLATE, tRUSER_INSERT, tREPG_TITLE_ORI, tRKEY_HEX) ';
                  end
                else
                  begin
                   strSQL := strSQL + 'tRGENRE, tRSUBGENRE, tRCONTENT, tRUSER_INSERT, tREPG_TITLE_ORI, tRKEY_HEX) ';
                  end;
                strSQL := strSQL + 'VALUES ( ';
                strSQL := strSQL + '''' + inttostr(i) + ''', ';   // tRID
                strSQL := strSQL + '''' + StringGrid1.Cells[3,i] + ''', '; // tRCHANNEL
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss''), '; // tRSCHEDULEDATE
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'')-0.29167, '; // tRSCHEDULEDATEGMT
                strSQL := strSQL + '''' + ReplaceString + ''', '; // tREPG_TITLE
                strSQL := strSQL + '''' + copy(Durstring, 12, 2) + copy(Durstring, 15, 2) + copy(Durstring, 18, 2) + ''', '; // tRDURATION
                if StringGrid1.Cells[7,i]='0' then rating := 0    // tRRATING
                else if StringGrid1.Cells[7,i]='7' then rating := 2
                else if StringGrid1.Cells[7,i]='8' then rating := 4
                else if StringGrid1.Cells[7,i]='9' then rating := 6
                else if StringGrid1.Cells[7,i]='10' then rating := 8
                else if StringGrid1.Cells[7,i]='11' then rating := 10
                else if StringGrid1.Cells[7,i]='12' then rating := 12
                else if StringGrid1.Cells[7,i]='13' then rating := 15;
                strSQL := strSQL + '''' + IntToStr(rating) + ''', ';
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],1,2) + ''', ';  // tRGENRE
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],3,2) + ''', ';  // tRSUBGENRE
                strSQL := strSQL + '''' + StringGrid1.Cells[9,i] + ''', '; // tRCONTENT

                if tca <> '' then
                  begin
                    strSQL := strSQL + '''' + StringGrid1.Cells[10,i] + ''', '; // tRCATEMPLATE
                  end;

                strSQL := strSQL + '''' + strUser + ''', ';
                ReplaceString:=Replace(StringGrid1.Cells[1,i], '''', '');
                strSQL := strSQL + 'upper(''' + trim(ReplaceString) + '''), ';
                strSQL := strSQL + '''' + 'NDSXTI-' + inttohex(i, 13) + ''') ';   // tRID
                //strSQL := strSQL + '''' + StringGrid1.Cells[11,i] + ''', '; // TRGROUPID
                //strSQL := strSQL + '''' + StringGrid1.Cells[12,i] + ''' )'; // TRPROGRAMID
                RecExc(strSQL);

                strSQL := 'INSERT INTO SGI.M_READXL ( ';
                strSQL := strSQL + 'RCHANNEL, RSCHEDULEDATE, RSCHEDULEDATEGMT, ';
                //strSQL := strSQL + 'RID, RCHANNEL, RSCHEDULEDATE, RSCHEDULEDATEGMT, ';
                strSQL := strSQL + 'REPG_TITLE, RDURATION, RRATING, ';
                if tca <> '' then
                  begin
                    //strSQL := strSQL + 'RGENRE, RSUBGENRE, RCONTENT, RCATEMPLATE, RUSER_INSERT, REPG_TITLE_ORI, RDATE_CREATE, RKEY_HEX) ';
                    strSQL := strSQL + 'RGENRE, RSUBGENRE, RCONTENT, RCATEMPLATE, RUSER_INSERT, REPG_TITLE_ORI, RDATE_CREATE) ';
                  end
                else
                  begin
                    //strSQL := strSQL + 'RGENRE, RSUBGENRE, RCONTENT, RUSER_INSERT, REPG_TITLE_ORI, RDATE_CREATE, RKEY_HEX) ';
                    strSQL := strSQL + 'RGENRE, RSUBGENRE, RCONTENT, RUSER_INSERT, REPG_TITLE_ORI, RDATE_CREATE) ';
                  end;
                strSQL := strSQL + 'VALUES ( ';
                //strSQL := strSQL + '''' + inttostr(angka) + ''', ';
                strSQL := strSQL + '''' + StringGrid1.Cells[3,i] + ''', ';
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss''), ';
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'')-0.29167, ';
                ReplaceString:=Replace(StringGrid1.Cells[2,i], '''', '`');
                ReplaceString:=Replace(ReplaceString, '`', '''''');
                strSQL := strSQL + '''' + ReplaceString + ''', ';
                Durstring:=valstrtodatetime(DateFloat1, StringGrid1.Cells[4,i]);
                strSQL := strSQL + '''' + copy(Durstring, 12, 2) + copy(Durstring, 15, 2) + copy(Durstring, 18, 2) + ''', ';
                if StringGrid1.Cells[7,i]='0' then rating := 0
                else if StringGrid1.Cells[7,i]='7' then rating := 2
                else if StringGrid1.Cells[7,i]='8' then rating := 4
                else if StringGrid1.Cells[7,i]='9' then rating := 6
                else if StringGrid1.Cells[7,i]='10' then rating := 8
                else if StringGrid1.Cells[7,i]='11' then rating := 10
                else if StringGrid1.Cells[7,i]='12' then rating := 12
                else if StringGrid1.Cells[7,i]='13' then rating := 15;
                strSQL := strSQL + '''' + IntToStr(rating) + ''', ';
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],1,2) + ''', ';
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],3,2) + ''', ';
                strSQL := strSQL + '''' + StringGrid1.Cells[9,i] + ''', ';
                if tca <> '' then
                  begin
                    strSQL := strSQL + '''' + StringGrid1.Cells[10,i] + ''', ';
                  end;
                strSQL := strSQL + '''' + strUser + ''', ';
                ReplaceString:=Replace(StringGrid1.Cells[1,i], '''', '');
                strSQL := strSQL + 'upper(''' + trim(ReplaceString) + '''), ';
                strSQL := strSQL + 'sysdate ) ';
                //strSQL := strSQL + 'sysdate , ';

                //strSQL := strSQL + '''' + 'NDSXTI-' + inttohex(angka, 13) + ''') ';
                //strSQL := strSQL + '''' + StringGrid1.Cells[11,i] + ''', ';
                //strSQL := strSQL + '''' + StringGrid1.Cells[12,i] + ''' )';
                RecExc(strSQL);


                strSQL := 'SELECT RID FROM SGI.M_READXL ';
                strSQL := strSQL + 'WHERE RCHANNEL=''' + StringGrid1.Cells[3,i] + ''' ';
                strSQL := strSQL + 'AND RSCHEDULEDATE=TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'')';
                RecSet(strSQL);


                angka :=  StrToInt(dm.DDL.FieldValues['RID']);

                strSQL := 'UPDATE SGI.M_READXL ';
                strSQL := strSQL + 'SET RKEY_HEX=''NDSXTI-' + inttohex(angka, 13) + ''' ';
                strSQL := strSQL + 'WHERE RCHANNEL=''' + StringGrid1.Cells[3,i] + ''' ';
                strSQL := strSQL + 'AND RSCHEDULEDATE=TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'') ';
                RecExc(strSQL);


                i:=i+1;
                pbRead.Position:=i;
                angka:=angka+1;
               End;
             ngReadXL.ClearRows;
             strSQL := ' SELECT * FROM ';
             strSQL := strSQL + ' (select * FROM m_readxl WHERE RCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
             DateFloat:=strtofloat(StringGrid1.Cells[5,1]);
             DateString:=valstrtodatetime(DateFloat, StringGrid1.Cells[4,1]);
             strSQL := strSQL + ' AND RSCHEDULEDATE >= TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'') ';
             DateFloat:=strtofloat(StringGrid1.Cells[5,i-1]);
             DateString:=valstrtodatetime(DateFloat, StringGrid1.Cells[4,i-1]);
             strSQL := strSQL + ' AND RSCHEDULEDATE <= TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'') ORDER by RID) xxx, ';
             strSQL := strSQL + ' (select * FROM m_synopsis) yyy, ';
            // strSQL := strSQL + ' (select * FROM M_Series) aaa, ';
             strSQL := strSQL + ' (select * FROM M_VOD) bbb ';
             strSQL := strSQL + ' WHERE repg_title_ori = syepg_title(+) ';
             strSQL := strSQL + ' AND rgenre = sycategory(+) ';
             //strSQL := strSQL + ' AND REPG_TITLE_ORI = SREpgTitle(+) ';
             strSQL := strSQL + ' AND REPG_TITLE_ORI = VODEPGTITLE(+) ';
             strSQL := strSQL + ' AND REPG_TITLE_ORI <> ''FILLER'' ';
             RecSet(strSQL);
             if not dm.DDL.Eof then
              begin
                while not dm.DDL.Eof do
                  Begin
                   if not VarIsNull(dm.DDL.FieldValues['RCATEMPLATE']) then
                    begin
                     kosong:= dm.DDL.FieldValues['RCATEMPLATE'];
                    end
                   else
                    begin
                     kosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['sysynopsis_ind']) then
                    begin
                     ind_kosong:= dm.DDL.FieldValues['sysynopsis_ind'];
                    end
                   else
                    begin
                     ind_kosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['sysynopsis_ind']) then
                    begin
                     eng_kosong:= dm.DDL.FieldValues['sysynopsis_eng'];
                    end
                   else
                    begin
                     eng_kosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODCAPRODUCTID']) then
                    begin
                     VODCaProductKosong:= dm.DDL.FieldValues['VODCAPRODUCTID'];
                    end
                   else
                    begin
                     VODCaProductKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODCAPSTARTDATE']) then
                    begin
                     VodStartDateKosong:= dm.DDL.FieldValues['VODCAPSTARTDATE'];
                    end
                   else
                    begin
                     VodStartDateKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODCAPENDDATE']) then
                    begin
                     VodEndDateKosong:= dm.DDL.FieldValues['VODCAPENDDATE'];
                    end
                   else
                    begin
                     VodEndDateKosong:='';
                    end;
                   
                   if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                    begin
                     GroupKeyKosong:= dm.DDL.FieldValues['VODGROUPKEY'];
                    end
                   else
                    begin
                     GroupKeyKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
                    begin
                     ProgramKeyKosong:= dm.DDL.FieldValues['VODPROGRAMKEY'];
                    end
                   else
                    begin
                     ProgramKeyKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODSTATUS']) then
                    begin
                     VODStatusKosong:= dm.DDL.FieldValues['VODSTATUS'];
                    end
                   else
                    begin
                     VODStatusKosong:='';
                    end;

                     ngReadXL.AddCells([dm.DDL.FieldValues['RID'],
                                        dm.DDL.FieldValues['RCHANNEL'],
                                        FormatDateTime('MMM/dd/yyyy HH:mm:ss', dm.DDL.FieldValues['RSCHEDULEDATE']),
                                        FormatDateTime('MMM/dd/yyyy HH:mm:ss', dm.DDL.FieldValues['RSCHEDULEDATEGMT']),
                                        dm.DDL.FieldValues['REPG_TITLE'],
                                        dm.DDL.FieldValues['repg_title_ori'],
                                        dm.DDL.FieldValues['RDURATION'],
                                        dm.DDL.FieldValues['RRATING'],
                                        dm.DDL.FieldValues['RGENRE'],
                                        dm.DDL.FieldValues['RSUBGENRE'],
                                        dm.DDL.FieldValues['RCONTENT'],
                                        dm.DDL.FieldValues['RKEY_HEX'],
                                        kosong,
                                        ind_kosong,
                                        eng_kosong,
                                        VODCaProductKosong,
                                        VodStartDateKosong,
                                        VodEndDateKosong,
                                        GroupKeyKosong,
                                        ProgramKeyKosong,
                                        VODStatusKosong
                                      ]);
                    dm.DDL.Next;
                 end;
                  i:=i+1;
                  pbRead.Position:=i-1;
              end
            else
              begin
               ngReadXL.ClearRows;
              end;
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
                   AssignFile(ERR , 'C:\SGI\LOG_ERROR\FormRead_'+strUser+'.log');
                     if fileexists('C:\SGI\LOG_ERROR\FormRead_'+strUser+'.log')
                   then append(ERR)
                     else Rewrite(ERR);

                     Writeln(ERR , encode64('Koneksi ke data_traffic gagal -> Err.Class: '+ E.ClassName+ ', pesan errornya gini: '+ E.Message) );
                     CloseFile(ERR);
                       showmessage('Maaf, terjadi kesalahan koneksi, silahkan periksa jaringan ke direktori data_traffic'+sLineBreak+''+sLineBreak+'Terima Kasih' );
       //ShowMessage('Exception class name = '+E.ClassName);
       //ShowMessage('Exception message = '+ E.Message );
                     Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' FrmRead : koneksi data_traffic gagal: ',E.Message );
                     CloseFile(actLOGLocal);
                     frmRead.Close;
                   end;
              end;
              pbRead.Visible:=False;
              Screen.Cursor:=crDefault;
              Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', strUserName, ' berhasil import channel ', dm.DDL.FieldValues['RCHANNEL']);
              Writeln(actLOG,'[', FormatDateTime('c',today),'] ', strUserName, ' berhasil import channel ',dm.DDL.FieldValues['RCHANNEL']);
              CloseFile(actLOG);
              CloseFile(actLOGLocal);
              MessageDlg('Table has been successfully imported!', mtInformation, [mbOK], 0);
              //ShowMessage('Table has been successfully imported!');
              frmExport.Show;
            end
           else
            begin
             screen.Cursor:=crDefault;
             pbRead.Visible:=False;
             MessageDlg('Channel '+StringGrid1.Cells[3,i]+' in use, wait a moment', mtInformation, [mbOK], 0);
             //ShowMessage('Channel '+StringGrid1.Cells[3,i]+' in use, wait a moment');
            end;
          end
         else
          begin
           pbRead.Visible:=False;
           Screen.Cursor:=crDefault;
           MessageDlg('Unknown channel name !', mtInformation, [mbOK], 0);
           //ShowMessage('Unknown channel name !');
          end;
        end;
    end;
end;

procedure TfrmRead.Button2xClick(Sender: TObject);
begin
 if ngReadXL.RowCount=0 then
 begin
  //frmExEPG.Show;
 end
 else
 begin
 frmExport.Show;
 end;
end;

procedure TfrmRead.Synopsis1Click(Sender: TObject);
var
 angka, ReplaceString, epgReplace, synInd, synEng  :String;
 i: Integer;
begin
 if strUserACC = 'Admin' then
  begin
   Screen.Cursor:=crHourGlass;
   frmSynopsis.ngSipnosis.ClearRows;

   i:=1;
   strSQL := 'Select title_ori, synopsis_i, synopsis_e, Category_code ';
   strSQL := strSQL + 'from synopsis ';
   strSQL := strSQL + 'order by title_ori ';
   RecSetAcc(strSQL);

   while not dm.AccDDL.Eof do
   begin
        angka:=fncangka;
         epgReplace:=Replace(trim(dm.AccDDL.FieldValues['title_ori']), '''','');
         epgReplace:=Replace(trim(epgReplace), '"','');

         if not VarIsNull(dm.AccDDL.FieldValues['synopsis_i'])then
          begin
           synInd:=Replace(trim(dm.AccDDL.FieldValues['synopsis_i']), '''','`');
           synInd:=Replace(trim(synInd), '`','''''');
           synInd:=Replace(trim(synInd), '"','');
          end
         else synInd:='';

         if not VarIsNull(dm.AccDDL.FieldValues['synopsis_e'])then
          begin
           synEng:=Replace(trim(dm.AccDDL.FieldValues['synopsis_e']), '''','`');
           synEng:=Replace(trim(synEng), '`','''''');
           synEng:=Replace(trim(synEng), '"','');
          end
         else synEng:='';

         ReplaceString:= copy(dm.AccDDL.FieldValues['Category_code'],1,1);
         strSQL := 'INSERT INTO SGI.M_SYNOPSIS ( ';
         strSQL := strSQL + 'SYID, SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, ';
         strSQL := strSQL + 'SYCATEGORY, SYUSER_CREATE, SYUSER_CREATEDATE, ';
         strSQL := strSQL + 'SYUSER_UPDATE, SYUSER_UPDATEDATE) ';
         strSQL := strSQL + 'VALUES ( ';
         strSQL := strSQL + '''' + angka + ''', ';
         strSQL := strSQL + 'upper(''' + epgReplace + '''), ';
         strSQL := strSQL + '''' + synInd + ''', ';
         strSQL := strSQL + '''' + synEng + ''', ';
         strSQL := strSQL + '''' + ReplaceString + ''', ';
         strSQL := strSQL + '''' + strUser + ''', ';
         strSQL := strSQL + 'sysdate, ';
         strSQL := strSQL + '''' + strUser + ''', ';
         strSQL := strSQL + 'sysdate) ';
         RecExc(strSQL);
    i:=i+1;
    dm.AccDDL.Next;
   end ;
   frmSynopsis.Show;
   Screen.Cursor:=crDefault;
   ShowMessage('Table has been exported!');
  end
 else
  begin
   ShowMessage('You Are Not Authorized');
  end;
end;

procedure TfrmRead.Exit1Click(Sender: TObject);
begin
 Application.Terminate;
end;


procedure TfrmRead.EPG1Click(Sender: TObject);
begin
 frmSchEditor.ShowModal;
end;

procedure TfrmRead.CAPackage1Click(Sender: TObject);
begin
 if strUserACC = 'Admin' then
  begin
   frmCA.ShowModal;
  end
 else
  begin
   ShowMessage('You Are Not Authorized');
  end;
end;

procedure TfrmRead.SynopsisManualClick(Sender: TObject);
begin
 frmSynopsisManual.ShowModal;
end;

procedure TfrmRead.LogOut1Click(Sender: TObject);
begin
 frmRead.Hide;
 frmChannel.close;
 frmUser.Close;
 frmLogin.Show;
end;

procedure TfrmRead.User1Click(Sender: TObject);
begin
 frmUser.ShowModal;
end;

procedure TfrmRead.Fromfile1Click(Sender: TObject);
begin
 if OpenDialog1.Execute then
  begin
   frmSynopsisXL.pbSynopsis.Max := StringGrid1.RowCount - 2;
   frmSynopsisXL.pbSynopsis.Min := 0;
   frmSynopsisXL.Show;
  end;
end;

procedure TfrmRead.CaServiceID1Click(Sender: TObject);
begin
 if strUserACC = 'Admin' then
  begin
   frmCAServiceID.ShowModal;
  end
 else
  begin
   ShowMessage('You Are Not Authorized');
  end;
end;

procedure TfrmRead.Channel1Click(Sender: TObject);
begin
 if strUserACC = 'Admin' then
  begin
   frmChannel.ShowModal;
  end
 else
  begin
   ShowMessage('You Are Not Authorized');
  end;
end;

procedure TfrmRead.Vision11Click(Sender: TObject);
begin
 if strUserACC = 'Admin' then
  begin
   frmVis1.ShowModal;
  end
 else
  begin
   ShowMessage('You Are Not Authorized');
  end;
end;

procedure TfrmRead.Clear1Click(Sender: TObject);
begin
 strSQL := 'DELETE FROM TEMP_READXL ';
 strSQL := strSQL + ' WHERE TRCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
 RecExc(strSQL);
 pbRead.Visible:=false;
end;

procedure TfrmRead.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Application.Terminate;
end;

procedure TfrmRead.FormCreate(Sender: TObject);
var
  ExecuteResult: integer;
  //Path: string;
  em_subject, em_body, em_mail: string;
begin
 with TrayIconData do
  begin
    cbSize := SizeOf(TrayIconData);
    Wnd := Handle;
    uID := 0;
    uFlags := NIF_MESSAGE + NIF_ICON + NIF_TIP;
    uCallbackMessage := WM_ICONTRAY;
    hIcon := Application.Icon.Handle;
    StrPCopy(szTip, Application.Title);

  end;
  //ShellExecute(Handle, nil, 'cmd.exe', '/C tnsping 192.168.110.81', nil, SW_HIDE);
  //Path := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName));
  //ExecuteResult := ShellExecute(handle, nil, 'cmd.exe', '/C tnsping 192.168.110.82', nil, SW_SHOWNORMAL);
  //if ExecuteResult <= 32 then ShowMessage('Error: ' + IntToStr(ExecuteResult));
  Shell_NotifyIcon(NIM_ADD, @TrayIconData);

  {em_subject := 'This is the subject line';
  em_body := 'Message body text goes here';
  em_mail := 'mailto:nusaputra@indovision.tv?subject=' +
    em_subject + '&body=' + em_body ;
  ShellExecute(
    Application.Handle,
    'open',
    PChar(em_mail),
    nil,
    nil,
    SW_SHOWNORMAL
  );}
end;

procedure TfrmRead.FormDestroy(Sender: TObject);
begin
 Shell_NotifyIcon(NIM_DELETE, @TrayIconData);
end;

procedure TfrmRead.Check1Click(Sender: TObject);
begin
if OpenDialog1.Execute then
  begin
   frmCheck.ShowModal;
  end;
end;

procedure TfrmRead.FormFile1Click(Sender: TObject);
begin
 if OpenDialog1.Execute then
  begin
   strIMGChoice := 'Upload';
   frmImage.Show;
  end;
end;

procedure TfrmRead.View1Click(Sender: TObject);
begin
  strIMGChoice := 'View';
  frmImage.ShowModal;
end;

procedure TfrmRead.SeriesLink1Click(Sender: TObject);
begin
  frmVOD.ShowModal;
  frmVOD.grpJustCA.Visible := False;
end;

procedure TfrmRead.SeriesLink2Click(Sender: TObject);
begin
  frmSeriesLink.ShowModal;
end;



procedure TfrmRead.Button1Click(Sender: TObject);
var
 i, rating, angka, climax, xxx, xx, y : integer;
 DateFloat, DateFloat1, climaxDateFloat, awalDateFloat : Double;
 tanggalawal, tanggalclimax, Durstring, DateString, awaldatestring, climaxdatestring, ReplaceString, GroupKeyKosong, ProgramKeyKosong, kosong, ind_kosong, eng_kosong, rid :String;
 VODCaProductKosong, VodStartDateKosong, VodEndDateKosong, VODStatusKosong : String;



begin
  if OpenDialog1.Execute then
    begin
     //i := 1;
     Screen.Cursor:=crHourGlass;
     for I := 0 to StringGrid1.RowCount - 1 do StringGrid1.Rows[I].Clear();
     if Xls_To_StringGrid(StringGrid1, OpenDialog1.FileName) then
        begin
         ngReadXL.ClearRows;
         strSQL := 'SELECT mchannel FROM M_CHANNEL WHERE mchannel = ''' + StringGrid1.Cells[3,1] + ''' ';
         RecSet(strSQL);
         if not dm.DDL.Eof then
          begin
           strSQL := 'SELECT distinct TRCHANNEL FROM TEMP_READXL WHERE TRCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
           RecSet(strSQL);
           if dm.DDL.Eof then
            begin
             i := 1;
             pbRead.Max:= StringGrid1.RowCount-2;
             pbRead.Min:= 0;
             pbRead.Visible := True;

             climax:= StringGrid1.RowCount-4;
             climaxDateFloat:=strtofloat(StringGrid1.Cells[5,climax]);
             awalDateFloat:= strtofloat(StringGrid1.Cells[5,1]);
             climaxdatestring:=valstrtodatetime(climaxDateFloat, StringGrid1.Cells[4,climax]);
             awaldatestring:=valstrtodatetime(awalDateFloat, StringGrid1.Cells[4,1]);

             tanggalawal := LeftStr(awaldatestring, 10);
             tanggalclimax := LeftStr(climaxdatestring, 10);

             strSQL := 'DELETE FROM M_READXL ';
             strSQL := strSQL + ' WHERE RCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
             strSQL := strSQL + ' AND RSCHEDULEDATE >= TO_Date(''' + tanggalawal + ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'') ';
             strSQL := strSQL + ' AND RSCHEDULEDATE <= TO_Date(''' + tanggalclimax + ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'') ';
             strSQL := strSQL + ' AND RSCHEDULEDATEGMT >= TO_Date(''' + tanggalawal + ' 00:00:00'',''mm/dd/yyyy hh24:mi:ss'')-0.29167 ';
             strSQL := strSQL + ' AND RSCHEDULEDATEGMT <= TO_Date(''' + tanggalclimax + ' 23:59:59'',''mm/dd/yyyy hh24:mi:ss'')-0.29167 ';
             RecExc2(strSQL);


             angka := StrToInt(fncangka2);
             while StringGrid1.Cells[0,i]<>'<eof>' do
               Begin
                DateFloat:=strtofloat(StringGrid1.Cells[5,i]);
                DateFloat1:=strtofloat(StringGrid1.Cells[6,i]);
                DateString:=valstrtodatetime(DateFloat, StringGrid1.Cells[4,i]);
                Durstring:=valstrtodatetime(DateFloat1, StringGrid1.Cells[4,i]);
                ReplaceString:=Replace(StringGrid1.Cells[2,i], '''', '`');
                ReplaceString:=Replace(ReplaceString, '`', '''''');

                strSQL := 'INSERT INTO SGI.TEMP_READXL ( ';
                strSQL := strSQL + 'tRID, tRCHANNEL, tRSCHEDULEDATE, tRSCHEDULEDATEGMT, ';
                strSQL := strSQL + 'tREPG_TITLE, tRDURATION, tRRATING, ';
                strSQL := strSQL + 'tRGENRE, tRSUBGENRE, tRCONTENT, tRCATEMPLATE, tRUSER_INSERT, tREPG_TITLE_ORI, tRKEY_HEX) ';
                strSQL := strSQL + 'VALUES ( ';
                strSQL := strSQL + '''' + inttostr(i) + ''', ';   // tRID
                strSQL := strSQL + '''' + StringGrid1.Cells[3,i] + ''', '; // tRCHANNEL
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss''), '; // tRSCHEDULEDATE
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'')-0.29167, '; // tRSCHEDULEDATEGMT
                strSQL := strSQL + '''' + ReplaceString + ''', '; // tREPG_TITLE
                strSQL := strSQL + '''' + copy(Durstring, 12, 2) + copy(Durstring, 15, 2) + copy(Durstring, 18, 2) + ''', '; // tRDURATION
                if StringGrid1.Cells[7,i]='0' then rating := 0    // tRRATING
                else if StringGrid1.Cells[7,i]='7' then rating := 2
                else if StringGrid1.Cells[7,i]='8' then rating := 4
                else if StringGrid1.Cells[7,i]='9' then rating := 6
                else if StringGrid1.Cells[7,i]='10' then rating := 8
                else if StringGrid1.Cells[7,i]='11' then rating := 10
                else if StringGrid1.Cells[7,i]='12' then rating := 12
                else if StringGrid1.Cells[7,i]='13' then rating := 15;
                strSQL := strSQL + '''' + IntToStr(rating) + ''', ';
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],1,2) + ''', ';  // tRGENRE
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],3,2) + ''', ';  // tRSUBGENRE
                strSQL := strSQL + '''' + StringGrid1.Cells[9,i] + ''', '; // tRCONTENT
                strSQL := strSQL + '''' + StringGrid1.Cells[10,i] + ''', '; // tRCATEMPLATE
                strSQL := strSQL + '''' + strUser + ''', ';
                ReplaceString:=Replace(StringGrid1.Cells[1,i], '''', '');
                strSQL := strSQL + 'upper(''' + trim(ReplaceString) + '''), ';
                strSQL := strSQL + '''' + 'NDSXTI-' + inttohex(i, 13) + ''') ';   // tRID
                //strSQL := strSQL + '''' + StringGrid1.Cells[11,i] + ''', '; // TRGROUPID
                //strSQL := strSQL + '''' + StringGrid1.Cells[12,i] + ''' )'; // TRPROGRAMID
                RecExc(strSQL);

                strSQL := 'INSERT INTO SGI.M_READXL ( ';
                strSQL := strSQL + 'RID, RCHANNEL, RSCHEDULEDATE, RSCHEDULEDATEGMT, ';
                strSQL := strSQL + 'REPG_TITLE, RDURATION, RRATING, ';
                strSQL := strSQL + 'RGENRE, RSUBGENRE, RCONTENT, RCATEMPLATE, RUSER_INSERT, REPG_TITLE_ORI, RDATE_CREATE, RKEY_HEX) ';
                strSQL := strSQL + 'VALUES ( ';
                strSQL := strSQL + '''' + inttostr(angka) + ''', ';
                strSQL := strSQL + '''' + StringGrid1.Cells[3,i] + ''', ';
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss''), ';
                strSQL := strSQL + 'TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'')-0.29167, ';
                ReplaceString:=Replace(StringGrid1.Cells[2,i], '''', '`');
                ReplaceString:=Replace(ReplaceString, '`', '''''');
                strSQL := strSQL + '''' + ReplaceString + ''', ';
                DateString:=valstrtodatetime(DateFloat1, StringGrid1.Cells[4,i]);
                strSQL := strSQL + '''' + copy(DateString, 12, 2) + copy(DateString, 15, 2) + copy(DateString, 18, 2) + ''', ';
                if StringGrid1.Cells[7,i]='0' then rating := 0
                else if StringGrid1.Cells[7,i]='7' then rating := 2
                else if StringGrid1.Cells[7,i]='8' then rating := 4
                else if StringGrid1.Cells[7,i]='9' then rating := 6
                else if StringGrid1.Cells[7,i]='10' then rating := 8
                else if StringGrid1.Cells[7,i]='11' then rating := 10
                else if StringGrid1.Cells[7,i]='12' then rating := 12
                else if StringGrid1.Cells[7,i]='13' then rating := 15;
                strSQL := strSQL + '''' + IntToStr(rating) + ''', ';
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],1,2) + ''', ';
                strSQL := strSQL + '''' + copy(StringGrid1.Cells[8,i],3,2) + ''', ';
                strSQL := strSQL + '''' + StringGrid1.Cells[9,i] + ''', ';
                strSQL := strSQL + '''' + StringGrid1.Cells[10,i] + ''', ';
                strSQL := strSQL + '''' + strUser + ''', ';
                ReplaceString:=Replace(StringGrid1.Cells[1,i], '''', '');
                strSQL := strSQL + 'upper(''' + trim(ReplaceString) + '''), ';
                strSQL := strSQL + 'sysdate , ';
                strSQL := strSQL + '''' + 'NDSXTI-' + inttohex(angka, 13) + ''') ';
                //strSQL := strSQL + '''' + StringGrid1.Cells[11,i] + ''', ';
                //strSQL := strSQL + '''' + StringGrid1.Cells[12,i] + ''' )';
                RecExc(strSQL);
                i:=i+1;
                pbRead.Position:=i;
                angka:=angka+1;
               End;
             ngReadXL.ClearRows;
             strSQL := ' SELECT * FROM ';
             strSQL := strSQL + ' (select * FROM m_readxl WHERE RCHANNEL = ''' + StringGrid1.Cells[3,1] + ''' ';
             DateFloat:=strtofloat(StringGrid1.Cells[5,1]);
             DateString:=valstrtodatetime(DateFloat, StringGrid1.Cells[4,1]);
             strSQL := strSQL + ' AND RSCHEDULEDATE >= TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'') ';
             DateFloat:=strtofloat(StringGrid1.Cells[5,i-1]);
             DateString:=valstrtodatetime(DateFloat, StringGrid1.Cells[4,i-1]);
             strSQL := strSQL + ' AND RSCHEDULEDATE <= TO_Date(''' + DateString + ''',''mm/dd/yyyy hh24:mi:ss'') ORDER by RID) xxx, ';
             strSQL := strSQL + ' (select * FROM m_synopsis) yyy, ';
            // strSQL := strSQL + ' (select * FROM M_Series) aaa, ';
             strSQL := strSQL + ' (select * FROM M_VOD) bbb ';
             strSQL := strSQL + ' WHERE repg_title_ori = syepg_title(+) ';
             strSQL := strSQL + ' AND rgenre = sycategory(+) ';
             //strSQL := strSQL + ' AND REPG_TITLE_ORI = SREpgTitle(+) ';
             strSQL := strSQL + ' AND REPG_TITLE_ORI = VODEPGTITLE(+) ';
             strSQL := strSQL + ' AND REPG_TITLE_ORI <> ''FILLER'' ';
             RecSet(strSQL);
             if not dm.DDL.Eof then
              begin
                while not dm.DDL.Eof do
                  Begin
                   if not VarIsNull(dm.DDL.FieldValues['RCATEMPLATE']) then
                    begin
                     kosong:= dm.DDL.FieldValues['RCATEMPLATE'];
                    end
                   else
                    begin
                     kosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['sysynopsis_ind']) then
                    begin
                     ind_kosong:= dm.DDL.FieldValues['sysynopsis_ind'];
                    end
                   else
                    begin
                     ind_kosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['sysynopsis_ind']) then
                    begin
                     eng_kosong:= dm.DDL.FieldValues['sysynopsis_eng'];
                    end
                   else
                    begin
                     eng_kosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODCAPRODUCTID']) then
                    begin
                     VODCaProductKosong:= dm.DDL.FieldValues['VODCAPRODUCTID'];
                    end
                   else
                    begin
                     VODCaProductKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODCAPSTARTDATE']) then
                    begin
                     VodStartDateKosong:= dm.DDL.FieldValues['VODCAPSTARTDATE'];
                    end
                   else
                    begin
                     VodStartDateKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODCAPENDDATE']) then
                    begin
                     VodEndDateKosong:= dm.DDL.FieldValues['VODCAPENDDATE'];
                    end
                   else
                    begin
                     VodEndDateKosong:='';
                    end;
                   
                   if not VarIsNull(dm.DDL.FieldValues['VODGROUPKEY']) then
                    begin
                     GroupKeyKosong:= dm.DDL.FieldValues['VODGROUPKEY'];
                    end
                   else
                    begin
                     GroupKeyKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) then
                    begin
                     ProgramKeyKosong:= dm.DDL.FieldValues['VODPROGRAMKEY'];
                    end
                   else
                    begin
                     ProgramKeyKosong:='';
                    end;

                   if not VarIsNull(dm.DDL.FieldValues['VODSTATUS']) then
                    begin
                     VODStatusKosong:= dm.DDL.FieldValues['VODSTATUS'];
                    end
                   else
                    begin
                     VODStatusKosong:='';
                    end;

                     ngReadXL.AddCells([dm.DDL.FieldValues['RID'],
                                        dm.DDL.FieldValues['RCHANNEL'],
                                        FormatDateTime('MMM/dd/yyyy HH:mm:ss', dm.DDL.FieldValues['RSCHEDULEDATE']),
                                        FormatDateTime('MMM/dd/yyyy HH:mm:ss', dm.DDL.FieldValues['RSCHEDULEDATEGMT']),
                                        dm.DDL.FieldValues['REPG_TITLE'],
                                        dm.DDL.FieldValues['repg_title_ori'],
                                        dm.DDL.FieldValues['RDURATION'],
                                        dm.DDL.FieldValues['RRATING'],
                                        dm.DDL.FieldValues['RGENRE'],
                                        dm.DDL.FieldValues['RSUBGENRE'],
                                        dm.DDL.FieldValues['RCONTENT'],
                                        dm.DDL.FieldValues['RKEY_HEX'],
                                        kosong,
                                        ind_kosong,
                                        eng_kosong,
                                        VODCaProductKosong,
                                        VodStartDateKosong,
                                        VodEndDateKosong,
                                        GroupKeyKosong,
                                        ProgramKeyKosong,
                                        VODStatusKosong
                                      ]);
                    dm.DDL.Next;
                 end;
                  i:=i+1;
                  pbRead.Position:=i-1;
              end
            else
              begin
               ngReadXL.ClearRows;
              end;
              pbRead.Visible:=False;
              Screen.Cursor:=crDefault;
              ShowMessage('Table has been exported!');
              frmExport.ShowModal;
            end
           else
            begin
             screen.Cursor:=crDefault;
             pbRead.Visible:=False;
             ShowMessage('Channel in use, wait a moment');
            end;
          end
         else
          begin
           pbRead.Visible:=False;
           Screen.Cursor:=crDefault;
           ShowMessage('Unknown channel name !');
          end;
        end;
    end;
end;

procedure TfrmRead.Button2Click(Sender: TObject);
begin
 if ngReadXL.RowCount=0 then
 begin
  //frmExEPG.Show;
 end
 else
 begin
 frmExport.Show;
 end;
end;

procedure TfrmRead.About1Click(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TfrmRead.FormShow(Sender: TObject);
var
Ini: TIniFile;
database: string;
begin
   Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');
   database := Decode64(Ini.ReadString('Config', 'database', 'Default'));


   strSQL := 'SELECT UUSR_ACC, UUSR_DESCRIPTION FROM M_USER WHERE UUSR_NAME = ''' + frmlogin.edtUsrName.Text + ''' ';
   RecSet(strSQL);

   if not (dm.DDL.FieldValues['UUSR_ACC'] = 'Admin') then
    begin
      frmread.CAPackage1.Visible := false;
      frmread.CaServiceID1.Visible := false;
      frmread.Vision11.Visible := false;
      frmread.Channel1.Visible := false;
      frmread.User1.Caption := 'Change Password';
      frmread.DecryptLog1.Visible := false;
      frmread.DatabaseMaintainer1.Visible := false;
    end
   else
    begin
      frmread.CAPackage1.Visible := true;
      frmread.CaServiceID1.Visible := true;
      frmread.Vision11.Visible := true;
      frmread.Channel1.Visible := true;
      frmread.User1.Caption := 'Manage User';
    end;

    ngReadXL.ClearRows;

    frmread.Caption := 'Schedule Converter '+ getappversion +' - Created By Dimas S, Ade I, Nugraha S & BA Team �2007-2016';
    frmread.StatusBar1.Panels[0].Text := 'Selamat Datang ' + dm.DDL.FieldValues['UUSR_DESCRIPTION'];
    frmread.StatusBar1.Panels[1].Text := 'Database : ' + database;

end;

procedure TfrmRead.DecryptLog1Click(Sender: TObject);
begin
frmdecode64.show;
end;

procedure TfrmRead.MJDConversion1Click(Sender: TObject);
begin
frmMJD.show;
end;

procedure TfrmRead.Timer1Timer(Sender: TObject);
begin
//Label1.Left:=Label1.left-1;
//if (Label1.left+Label1.width) <= 0 then
//Label1.left:=frmread.Width-(label1.Width+2);
end;

procedure TfrmRead.DatabaseMaintainer1Click(Sender: TObject);
begin
MaintainBox.show;
end;


procedure TfrmRead.DBChooserClick(Sender: TObject);
var
DestPath, DestFile : string;
begin
  DestPath := ExtractFilePath(Application.EXEName);
  DestFile := DestPath + 'EPGdbSetting.exe';

            Application.Terminate;
            ShellExecute(Application.Handle, PChar('open'), PChar(DestFile),
            PChar(''), nil, SW_NORMAL)
end;

end.
