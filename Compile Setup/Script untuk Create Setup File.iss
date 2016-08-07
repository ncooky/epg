; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

#define MyAppName "Excel to NDS Scheduller Converter"
#define MyAppDir "SGI"
#define MyAppVersion "6.0.4.3"
#define MyAppPublisher "Nugraha Saputra"
#define MyAppURL "http://www.anekajual.com"
#define MyAppExeName "EPG.exe"

[Setup]
AppId={{136AB635-DA8F-460D-98A0-ED51566E29C4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={sd}\{#MyAppDir}
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\EPG.exe
SetupIconFile=E:\Project\Project XL to DB new - Test XTI\images\xti2.ico
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
;OutputDir=userdocs:Inno Setup Examples Output

[CustomMessages]
OptionsFormCaption=Setup options...
RepairButtonCaption=Repair
UninstallButtonCaption=Uninstall

[Code]
const
  mrRepair = 100;
  mrUninstall = 101;

function ShowOptionsForm: TModalResult;
var
  OptionsForm: TSetupForm;
  RepairButton: TNewButton;
  UninstallButton: TNewButton;
begin
  Result := mrNone;
  OptionsForm := CreateCustomForm;
  try
    OptionsForm.Width := 220;
    OptionsForm.Caption := ExpandConstant('{cm:OptionsFormCaption}');
    OptionsForm.Position := poScreenCenter;

    RepairButton := TNewButton.Create(OptionsForm);
    RepairButton.Parent := OptionsForm;
    RepairButton.Left := 8;
    RepairButton.Top := 8;
    RepairButton.Width := OptionsForm.ClientWidth - 16;
    RepairButton.Caption := ExpandConstant('{cm:RepairButtonCaption}');
    RepairButton.ModalResult := mrRepair;

    UninstallButton := TNewButton.Create(OptionsForm);
    UninstallButton.Parent := OptionsForm;
    UninstallButton.Left := 8;
    UninstallButton.Top := RepairButton.Top + RepairButton.Height + 8;
    UninstallButton.Width := OptionsForm.ClientWidth - 16;
    UninstallButton.Caption := ExpandConstant('{cm:UninstallButtonCaption}');
    UninstallButton.ModalResult := mrUninstall;

    OptionsForm.ClientHeight := RepairButton.Height + UninstallButton.Height + 24;
    Result := OptionsForm.ShowModal;
  finally
    OptionsForm.Free;
  end;
end;

function GetUninstallerPath: string;
var
  RegKey: string;
begin
  Result := '';
  RegKey := Format('%s\%s_is1', ['Software\Microsoft\Windows\CurrentVersion\Uninstall', 
    '{#emit SetupSetting("AppId")}']);
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE, RegKey, 'UninstallString', Result) then
    RegQueryStringValue(HKEY_CURRENT_USER, RegKey, 'UninstallString', Result);
end;

function InitializeSetup: Boolean;
var
  UninstPath: string;
  ResultCode: Integer;  
begin
  Result := True;
  UninstPath := RemoveQuotes(GetUninstallerPath);
  if UninstPath <> '' then
  begin
    case ShowOptionsForm of
      mrRepair: Result := True;
      mrUninstall: 
      begin
        Result := False;
        if not Exec(UninstPath, '', '', SW_SHOW, ewNoWait, ResultCode) then
          MsgBox(FmtMessage(SetupMessage(msgUninstallOpenError), [UninstPath]), mbError, MB_OK);
      end;
    else
      Result := False;
    end;
  end;
end;

[Files]
Source: "EPG.exe"; DestDir: "{app}"
Source: "EPGLoader.exe"; DestDir: "{app}"
Source: "EPGdbSetting.exe"; DestDir: "{app}"
Source: "epg.ini"; DestDir: "{app}"
;Source: "MyProg.chm"; DestDir: "{app}"
;Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme

[Icons]
Name: "{group}\Excel to NDS Scheduller Converter"; Filename: "{app}\EPGLoader.exe"
Name: "{group}\{cm:UninstallProgram,Excel to NDS Scheduller Converter}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Excel to NDS Scheduller Converter"; Filename: "{app}\EPGLoader.exe"

