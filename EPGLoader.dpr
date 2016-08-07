program EPGLoader;

uses
  Forms,
  frm_updater in 'frm_updater.pas' {frm_update};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'EPG Updater';
  Application.CreateForm(Tfrm_update, frm_update);
  Application.Run;
end.
