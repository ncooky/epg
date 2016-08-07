unit frm_Login;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, XPMan, ToolWin, ActnMan, ActnCtrls,
  ActnColorMaps, ComCtrls, OleCtrls, DCPcrypt2, DCPmd5, URLMon, ShellApi, INIFiles;


type
  TfrmLogin = class(TForm)
    edtUsrName: TEdit;
    edtUsrPass: TEdit;
    btnLogin: TButton;
    Button2: TButton;
    XPManifest1: TXPManifest;
    Image1: TImage;
    Label1: TLabel;
    procedure btnLoginClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLogin: TfrmLogin;
  strUser, strUserName, strUserACC : String;
  Hash: TDCP_md5;
  actLOG, ERR, actLOGLocal : TextFile;
  today : TDateTime;


implementation

uses frm_dm, frm_Read;

{$R *.dfm}
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

function DownloadFile(SourceFile, DestFile: string): Boolean; 
begin 
  try
    Result := UrlDownloadToFile(nil, PChar(SourceFile), PChar(DestFile), 0, nil) = 0;
  except
    Result := False;
  end; 
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
            Result:= IntToStr (HiWord (dwFileVersionMS)) + '.' +
                     IntToStr (LoWord (dwFileVersionMS)) + '.' +
                     IntToStr (HiWord (dwFileVersionLS)) + '.' +
                     IntToStr (LoWord (dwFileVersionLS)) ; // ' @Mid 2016';
         end;
       finally
         FreeMem (Pt);
       end;
     end;
   end;

function  AppVersion:string;
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
            Result:= IntToStr (HiWord (dwFileVersionMS)) +
                     IntToStr (LoWord (dwFileVersionMS)) +
                     IntToStr (HiWord (dwFileVersionLS)) + 
                     IntToStr (LoWord (dwFileVersionLS)) ; // ' @Mid 2016';
         end;
       finally
         FreeMem (Pt);
       end;
     end;
   end;

procedure TfrmLogin.btnLoginClick(Sender: TObject);
var strPass : String;
begin

try
  AssignFile(actLOG, '\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
  if fileexists('\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
    then append(actLOG)
    else Rewrite(actLOG);
       Writeln(actLOG,'[', FormatDateTime('c',today),'] ', 'Login Window berhasil');
       CloseFile(actLOG);
except
     on E : Exception do
     begin
       showmessage('Maaf, terdapat kesalahan sambungan ke jaringan "data_traffic" , mohon periksa kondisi jaringan anda!' +sLineBreak+''+sLineBreak+'Terima Kasih' );
       frmLogin.Close ;
     end;
end;

 strPass := fncMD5(edtUsrPass.Text);
 strSQL := 'SELECT * FROM M_USER WHERE UUSR_NAME = ''' + edtUsrName.Text + ''' and UUSR_PASSWORD = ''' + strPass + ''' ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   ShowMessage('Wrong User Name And Password!!!');
   edtUsrName.Text:='';
   edtUsrPass.Text:='';
   edtUsrName.SetFocus;
  end
 else
  begin
   frmRead.Show;
   strUser:='';
   strUser:=edtUsrName.Text;
   strUserACC:= dm.DDL.FieldValues['UUSR_ACC'];
   strUserName:= dm.DDL.FieldValues['UUSR_DESCRIPTION'];
   edtUsrName.Text:='';
   edtUsrPass.Text:='';
   frmLogin.Hide;
//    AssignFile(actLOG, '\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
//    if fileexists('\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
    try
        AssignFile(actLOGLocal, 'C:\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
        if fileexists('C:\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
        then append(actLOGLocal)
        else Rewrite(actLOGLocal);
        Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', strUserName, ' berhasil masuk');
        CloseFile(actLOGLocal);
    except
       on E : Exception do
       begin
         AssignFile(ERR , 'C:\SGI\LOG_ERROR\FormLogin_'+strUser+'.log');
         if fileexists('C:\SGI\LOG_ERROR\FormLogin_'+strUser+'.log')
         then append(ERR)
         else Rewrite(ERR);

         Writeln(ERR , encode64('Penyimpanan Log Gagal -> Err.Class: '+ E.ClassName+ ', pesan errornya gini: '+ E.Message) );
         CloseFile(ERR);
         showmessage('Maaf, terjadi kesalahan Penyimpanan Log, silahkan Tutup dahulu aplikasi ini, dan jalankan kembali'+sLineBreak+''+sLineBreak+'Terima Kasih' );
         //ShowMessage('Exception class name = '+E.ClassName);
           //ShowMessage('Exception message = '+ E.Message );
         Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ',strUserName, ' frmLogin : penyimpanan log gagal: ',E.Message );
         CloseFile(actLOGLocal);
       end;
    end;
  end;
end;

procedure TfrmLogin.Button2Click(Sender: TObject);
begin
 Application.Terminate;
end;

procedure TfrmLogin.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key=#13 then
  begin
    btnLoginClick(Sender);
  end;
end;

procedure TfrmLogin.FormShow(Sender: TObject);

var
width, LUpdate, LoldDate : integer;
oldfile : TextFile;
Ini: TIniFile;
DestPath, DestFile, SourceFile, SrcFile : string;
//const
//  SourceFile = '\\192.168.180.180\data_traffic\SGI\SGI\UPDATE\EPGLoader.exe';

begin


width := round(frmLogin.Width / 2) - round(Label1.Width/2) - 22;
Label1.Caption := 'Version ' + GetAppVersion + ' @mid 2016';
Label1.Left := width ;
today := Now;

strSQL := 'Select * from M_VERSION';
RecSet(strSQL);
      Ini := TIniFile.Create(ExtractFilePath(Application.EXEName) + 'epg.ini');
      try
        SrcFile := Ini.ReadString('Config', 'path','Default');
        SourceFile := SrcFile+'EPGLoader.exe';



       if SrcFile = 'Default' then
        begin
          SourceFile := '\\192.168.180.180\data_traffic\SGI\SGI\UPDATE\EPGLoader.exe';
          Ini.WriteString('Config', 'path',ExtractFilePath(SourceFile));
        end;

        Ini.WriteString('Versions', 'ver', AppVersion);
        Ini.WriteString('Versions', 'verstr', GetAppVersion);

      finally
        Ini.Free;
      end;


      DestPath := ExtractFilePath(Application.EXEName);
      DestFile := DestPath + 'EPGLoader.exe';

      LUpdate := FileAge(SourceFile);
      LoldDate := FileAge(DestFile);

      if LoldDate < LUpdate then
       begin
        if DownloadFile(SourceFile, DestFile) then
          begin
            //ShowMessage('Aplikasi EPG - XTI Versi '+ GetAppVersion +' berhasil diperbaharui dengan versi '+ dm.DDL.FieldValues['VERSTR']);
            Application.Terminate;
            ShellExecute(Application.Handle, PChar('open'), PChar(DestFile),
            PChar(''), nil, SW_NORMAL)
          end
        else
        ShowMessage('Error while downloading ' + SourceFile)
       end;


      //SourceFile := SrcFile+'\EPG.exe';
    {if StrToInt(AppVersion) < StrToInt(dm.DDL.FieldValues['VERSION']) then
      begin
        showmessage('Aplikasi akan update automatis, mohon tunggu.... ');
        if DownloadFile(SourceFile, DestFile) then
          begin
            ShowMessage('Aplikasi EPG - XTI Versi '+ GetAppVersion +' berhasil diperbaharui dengan versi '+ dm.DDL.FieldValues['VERSTR']);
            Application.Terminate;
            ShellExecute(Application.Handle, PChar('open'), PChar(DestFile),
            PChar(''), nil, SW_NORMAL)
          end
        else
        ShowMessage('Error while downloading ' + SourceFile)

      end;}
      
		//if not DirectoryExists('\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG') then
		//begin
		//	CreateDir('\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG');
		//end;
    if not DirectoryExists('C:\SGI\SGI_LOG') then
		begin
			CreateDir('C:\SGI\SGI_LOG');
		end;

end;

end.
