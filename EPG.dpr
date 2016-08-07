program EPG;



uses
  Forms,
  frm_Read in 'frm_Read.pas' {frmRead},
  frm_dm in 'frm_dm.pas' {dm},
  frm_Export in 'frm_Export.pas' {frmExport},
  frm_Image in 'frm_Image.pas' {frmImage},
  frm_EPG in 'frm_EPG.pas' {frmSchEditor},
  frm_ExEPG in 'frm_ExEPG.pas' {frmExEPG},
  frm_InsertCA in 'frm_InsertCA.pas' {frmCA},
  frm_SynopsisManual in 'frm_SynopsisManual.pas' {frmSynopsisManual},
  frm_Login in 'frm_Login.pas' {frmLogin},
  frm_User in 'frm_User.pas' {frmUser},
  frm_SynopsisXL in 'frm_SynopsisXL.pas' {frmSynopsisXL},
  frm_CAServiceID in 'frm_CAServiceID.pas' {frmCAServiceID},
  frm_Channel in 'frm_Channel.pas' {frmChannel},
  frm_Vis1 in 'frm_Vis1.pas' {frmVis1},
  ProgressBar in 'ProgressBar.pas' {Form1},
  frm_Edit1CA in 'frm_Edit1CA.pas' {frmEditCaEvent},
  frm_Check in 'frm_Check.pas' {frmCheck},
  frm_VOD in 'frm_VOD.pas' {frmVOD},
  mdl_Global in 'Module\mdl_Global.pas' {mdlGlobal: TDataModule},
  frm_SeriesLink in 'frm_SeriesLink.pas' {frmSeriesLink},
  frm_CCI_bit in 'frm_CCI_bit.pas' {frmCCIBits},
  frm_About in 'frm_About.pas' {AboutBox},
  frm_Decode in 'frm_Decode.pas' {frmdecode64},
  MJDConverter in 'MJDConverter.pas' {frmMJD},
  frm_maintaindb in 'frm_maintaindb.pas' {MaintainBox};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'NDS Schedule Converter';
  Application.CreateForm(TfrmLogin, frmLogin);
  Application.CreateForm(TfrmRead, frmRead);
  Application.CreateForm(Tdm, dm);
  Application.CreateForm(TfrmExport, frmExport);
  Application.CreateForm(TfrmImage, frmImage);
  Application.CreateForm(TfrmSchEditor, frmSchEditor);
  Application.CreateForm(TfrmExEPG, frmExEPG);
  Application.CreateForm(TfrmCA, frmCA);
  Application.CreateForm(TfrmSynopsisManual, frmSynopsisManual);
  Application.CreateForm(TfrmUser, frmUser);
  Application.CreateForm(TfrmSynopsisXL, frmSynopsisXL);
  Application.CreateForm(TfrmCAServiceID, frmCAServiceID);
  Application.CreateForm(TfrmChannel, frmChannel);
  Application.CreateForm(TfrmVis1, frmVis1);
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TfrmEditCaEvent, frmEditCaEvent);
  Application.CreateForm(TfrmCheck, frmCheck);
  Application.CreateForm(TfrmVOD, frmVOD);
  Application.CreateForm(TmdlGlobal, mdlGlobal);
  Application.CreateForm(TfrmSeriesLink, frmSeriesLink);
  Application.CreateForm(TfrmCCIBits, frmCCIBits);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(Tfrmdecode64, frmdecode64);
  Application.CreateForm(TfrmMJD, frmMJD);
  Application.CreateForm(TMaintainBox, MaintainBox);
  Application.Run;
end.
