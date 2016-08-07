unit frm_Channel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, NxColumns, NxColumnClasses, Menus, StdCtrls, Commctrl;

type
  TfrmChannel = class(TForm)
    PopupMenu1: TPopupMenu;
    AddRow1: TMenuItem;                                                                                                              
    AddNew1: TMenuItem;
    Update1: TMenuItem;
    Exit1: TMenuItem;
    ScrollBox1: TScrollBox;
    ScrollBox2: TScrollBox;
    ngChannel: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxComboBoxColumn1: TNxComboBoxColumn;
    NxTextColumn7: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    NxTextColumn9: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    NxTextColumn11: TNxTextColumn;
    NxTextColumn13: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxComboBoxColumn2: TNxComboBoxColumn;
    NxComboBoxColumn3: TNxComboBoxColumn;
    Delete1: TMenuItem;
    NxTextColumn6: TNxTextColumn;
    Panel1: TPanel;
    cbCATemplate: TComboBox;
    Label2: TLabel;
    Label1: TLabel;
    cbChannel: TComboBox;
    Button1: TButton;
    Panel2: TPanel;
    Image1: TImage;
    CCINumber1: TMenuItem;
    NxTextColumn12: TNxComboBoxColumn;
    NxComboBoxColumn4: TNxCheckBoxColumn;
    chkHDCh: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure cbChannelSelect(Sender: TObject);
    procedure AddRow1Click(Sender: TObject);
    procedure AddNew1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Update1Click(Sender: TObject);
    procedure ngChannelMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Button1Click(Sender: TObject);
    procedure ngChannelSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure Delete1Click(Sender: TObject);
    procedure CCINumber1Click(Sender: TObject);
    procedure chkHDChClick(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmChannel: TfrmChannel;
  X, Y :integer;

implementation

uses frm_dm, frm_Login, Types, frm_CCI_bit;

{$R *.dfm}

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
                     IntToStr (LoWord (dwFileVersionLS)) + ' @Early 2016';
         end;
       finally
         FreeMem (Pt);
       end;
     end;
   end;

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT MCID FROM M_CHANNEL ORDER BY MCID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   strAngka := dm.DDL.FieldValues['MCID'] + 1;
  end;
 fncangka:=IntToStr(strAngka);
end;

function fncnmr():string;
var
 strNMR : integer;
begin
 strSQL := 'SELECT CBID FROM M_CHANNEL_BITS ORDER BY CBID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strNMR := 1;
  end
 else
  begin
   dm.ddl.Last;
   strNMR := dm.DDL.FieldValues['CBID'] + 1;
  end;
 fncnmr:=IntToStr(strNMR);
end;

procedure prcShow(channel:string);

var
 i:integer;

function IfNull( const Value, Default : OleVariant ) : OleVariant;
begin
  if Value = NULL then
    Result := Default
  else
    Result := Value;
end;                                

begin
frmChannel.Update1.Enabled := False;






 strSQL := 'SELECT * FROM M_CHANNEL, m_catemplate where mcsiserviceid = catcode';
 if frmChannel.cbChannel.Text <> trim('All Channel') then
  begin
   strSQL := strSQL + ' AND MCHANNEL = ''' + channel + ''' ';
   frmChannel.Update1.Enabled := True;
   frmChannel.chkHDCh.Visible := False;
   //frmChannel.chkHDCh.Checked := False;
  end
 else if (frmChannel.cbChannel.Text = trim('All Channel')) and (frmChannel.chkHDCh.checked = True) then
  begin
   strSQL := strSQL + ' AND MCHANNEL Like ''%HD'' ';
   frmChannel.chkHDCh.Visible  := True;
  end
 else if frmChannel.cbChannel.Text = trim('All Channel') then
    begin
        frmChannel.chkHDCh.Enabled := True;
    end;

 strSQL := strSQL + ' ORDER BY MCHANNEL';
 RecSet(strSQL);

 //frmChannel.chkHDCh.Enabled := True;

 i:=1;
 if (dm.DDL.FieldValues['MCSISERVICEID']='2002') or (dm.DDL.FieldValues['MCSISERVICEID']='2202') or (dm.DDL.FieldValues['MCSISERVICEID']='2005') or (dm.DDL.FieldValues['MCSISERVICEID']='100') or (dm.DDL.FieldValues['MCSISERVICEID']='200') then
  begin
    frmChannel.CCINumber1.Visible := True;
  end
 else
  begin
    frmChannel.CCINumber1.Visible := False;
  end;

//  if frmChannel.cbChannel.Text <> trim('All Channel') then
//      frmChannel.chkHDCh.Checked := false;
//       frmChannel.CCINumber1.Visible := false;

  if (frmChannel.cbChannel.Text = trim('All Channel')) and (frmChannel.chkHDCh.checked = True) then
    frmChannel.CCINumber1.Visible := false;

 frmChannel.ngChannel.ClearRows;
 while not dm.DDL.Eof do
 begin
  frmChannel.ngChannel.AddCells([inttostr(i),
                                dm.DDL.FieldValues['MCHANNEL'],
                                dm.DDL.FieldValues['MCHANNEL_DESCRIPTION'],
                                dm.DDL.FieldValues['CATdescription'],
                                dm.DDL.FieldValues['MSGINAME'],
                                dm.DDL.FieldValues['MUSERNIBBLE1'],
                                dm.DDL.FieldValues['MSYNOPSIS_STATUS'],
                                dm.DDL.FieldValues['MCHANNEL'],
                                dm.DDL.FieldValues['MCHANNEL_DESCRIPTION'],
                                dm.DDL.FieldValues['CATdescription'],
                                dm.DDL.FieldValues['MSGINAME'],
                                dm.DDL.FieldValues['MUSERNIBBLE1'],
                                dm.DDL.FieldValues['MSYNOPSIS_STATUS'],
                                IfNull(dm.DDL.FieldValues['MPLAYOUT_SOURCE'],''),
                                IfNull(dm.DDL.FieldValues['MCHANNEL_NUMBER'],0),
                                IfNull(dm.DDL.FieldValues['MSTB_PAIRING'],'0'),
                                IfNull(dm.DDL.FieldValues['MCH_ACTIVE'],'0')
                                  ]);
   i:=i+1;
   dm.DDL.Next;
 end;
end;

procedure ShowBalloonTip(Control: TWinControl; Icon: integer; Title: pchar; Text: PWideChar;
BackCL, TextCL: TColor);
const 
  TOOLTIPS_CLASS = 'tooltips_class32'; 
  TTS_ALWAYSTIP = $01; 
  TTS_NOPREFIX = $02; 
  TTS_BALLOON = $40; 
  TTF_SUBCLASS = $0010; 
  TTF_TRANSPARENT = $0100; 
  TTF_CENTERTIP = $0002; 
  TTM_ADDTOOL = $0400 + 50; 
  TTM_SETTITLE = (WM_USER + 32); 
  ICC_WIN95_CLASSES = $000000FF; 
type 
  TOOLINFO = packed record 
    cbSize: Integer; 
    uFlags: Integer; 
    hwnd: THandle; 
    uId: Integer; 
    rect: TRect; 
    hinst: THandle; 
    lpszText: PWideChar; 
    lParam: Integer; 
  end; 
var 
  hWndTip: THandle; 
  ti: TOOLINFO; 
  hWnd: THandle; 
begin 
  hWnd    := Control.Handle; 
  hWndTip := CreateWindow(TOOLTIPS_CLASS, nil, 
    WS_POPUP or TTS_NOPREFIX or TTS_BALLOON or TTS_ALWAYSTIP, 
    0, 0, 0, 0, hWnd, 0, HInstance, nil); 
  if hWndTip <> 0 then 
  begin 
    SetWindowPos(hWndTip, HWND_TOPMOST, 0, 0, 0, 0, 
      SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);
    ti.cbSize := SizeOf(ti);
    ti.uFlags := TTF_CENTERTIP or TTF_TRANSPARENT or TTF_SUBCLASS;
    ti.hwnd := hWnd; 
    ti.lpszText := Text; 
    Windows.GetClientRect(hWnd, ti.rect); 
    SendMessage(hWndTip, TTM_SETTIPBKCOLOR, BackCL, 0); 
    SendMessage(hWndTip, TTM_SETTIPTEXTCOLOR, TextCL, 0); 
    SendMessage(hWndTip, TTM_ADDTOOL, 1, Integer(@ti)); 
    SendMessage(hWndTip, TTM_SETTITLE, Icon mod 4, Integer(Title)); 
  end; 
end;

procedure TfrmChannel.FormShow(Sender: TObject);
var
 i:integer;
 item : TStrings;
begin
 cbChannel.Clear;
 ngChannel.ClearRows;
 frmChannel.CCINumber1.Visible := False;
 frmChannel.AddRow1.Visible := false;
 frmChannel.Caption := 'Form Channel Manager ' + GetAppVersion;
 strSQL := 'SELECT MCHANNEL FROM M_CHANNEL ORDER BY MCHANNEL';
 RecSet(strSQL);

 Item:=cbChannel.Items.Create;
 item.Add('All Channel');
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['MCHANNEL']);
  dm.DDL.Next;
 end;
 cbChannel.ItemIndex:=0;

 strSQL := ' SELECT catdescription from m_catemplate order by catdescription';
 RecSet(strSQL);

 item:=cbCATemplate.Items.Create;
 item.Add('-SELECT ONE-');
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['catdescription']);
  dm.DDL.Next;
 end;
 cbCATemplate.ItemIndex:=0;

 strSQL := ' SELECT catdescription from m_catemplate order by catdescription';
 RecSet(strSQL);

 item:=NxComboBoxColumn2.Items.Create;
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['catdescription']);
  dm.DDL.Next;
 end;

end;

procedure TfrmChannel.cbChannelSelect(Sender: TObject);
begin
 Screen.Cursor:=crHourGlass;
 ngChannel.ClearRows;
 prcShow(cbChannel.Text);
 Screen.Cursor:=crDefault;
end;

procedure TfrmChannel.AddRow1Click(Sender: TObject);
begin
 ngChannel.AddRow(1);
end;

procedure TfrmChannel.AddNew1Click(Sender: TObject);
begin
 ngChannel.ClearRows;
 ngChannel.AddRow(1);
 frmchannel.Update1.Enabled := true;
end;

procedure TfrmChannel.Exit1Click(Sender: TObject);
begin
 frmChannel.Close;
end;

procedure TfrmChannel.Update1Click(Sender: TObject);
var
 i:integer;
 angka, nmr, OldCh, txterror:String;
 stbparing:integer;
label
 errortext;
 	const
		sLineBreak = {$IFDEF LINUX} AnsiChar(#10) {$ENDIF}
			{$IFDEF MSWINDOWS} AnsiString(#13#10) {$ENDIF};


begin
Screen.Cursor:=crHourGlass;
   if ngChannel.Cells[1,y] = '' then
    begin
      txterror := 'Please Input Channel Name';
      ngChannel.SelectCell(1,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[2,y] = '' then
    begin
      txterror := 'Please Input Channel Description';
      ngChannel.SelectCell(2,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[3,y] = '' then
    begin
      txterror := 'Please Select CA Template';
      ngChannel.SelectCell(3,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[4,y] = '' then
    begin
      txterror := 'Please Input Name';
      ngChannel.SelectCell(4,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[5,y] = '' then
    begin
      txterror := 'Please Select User Nible';
      ngChannel.SelectCell(5,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[6,y] = '' then
    begin
      txterror := 'Please Select Synopsys mode';
      ngChannel.SelectCell(6,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[13,y] = '' then
    begin
      txterror := 'Please Input Encoder Number';
      ngChannel.SelectCell(13,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[14,y] = '' then
    begin
      txterror := 'Please Input Viewing Number';
      ngChannel.SelectCell(14,y) ;
      Goto errortext;
    end
   else if  ngChannel.Cells[15,y] = '' then
    begin
      txterror := 'Please select STB Pairing';
      ngChannel.SelectCell(15,y) ;
      Goto errortext;
    end
   else

 for i := 0 to ngChannel.RowCount-1 do
 begin

  strSQL := 'SELECT CATCODE FROM M_CATEMPLATE ';
  strSQL := strSQL + 'WHERE catdescription = ''' + ngChannel.Cells[3,i] + ''' ';
  RecSet2(strSQL);
    if (dm.DDL2.FieldValues['CATCODE'] = '2002') or (dm.DDL2.FieldValues['CATCODE'] = '2202')  or (dm.DDL2.FieldValues['CATCODE'] = '2005')  or (dm.DDL2.FieldValues['CATCODE'] = '100') or (dm.DDL2.FieldValues['CATCODE'] = '200')  then
      begin
        strSQL := 'SELECT * FROM M_CHANNEL_BITS ';
        strSQL := strSQL + 'WHERE CBCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
        RecSet2(strSQL);
        if dm.DDL2.Eof then
          begin
            ShowMessage('it seems you chose CGMS-A mode, You Must Input 2 Digits HEX number before it saved!');
            frmCCIBits.ShowModal;
          end;
      end
    else
      begin
        strSQL := 'SELECT * FROM M_CHANNEL_BITS ';
        strSQL := strSQL + 'WHERE CBCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
        RecSet2(strSQL);
        if not dm.DDL2.Eof then
          begin
            strSQL := 'DELETE FROM M_CHANNEL_BITS ';
            strSQL := strSQL + 'WHERE CBCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
            RecExc2(strSQL);
          end

      end;

  angka:=fncangka;
  nmr:=fncnmr;
  strSQL := 'SELECT * FROM M_CHANNEL ';
  strSQL := strSQL + 'WHERE MCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
  strSQL := strSQL + 'AND MCSISERVICEID = (select catcode from m_catemplate  ';
  strSQL := strSQL + 'where catdescription = ''' + ngChannel.Cells[9,i] + ''') ';
  RecSet(strSQL);


try
  AssignFile(actLOGLocal, 'C:\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
  if fileexists('C:\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
    then append(actLOGLocal)
    else Rewrite(actLOGLocal);
except
     on E : Exception do
     begin
       showmessage('Maaf, terdapat kesalahan dalam penyimpanan LOG, mohon periksa akses level pada PC Anda!'+sLineBreak+sLineBreak+ 'Terima Kasih' );
       frmChannel.Close ;
     end;
end;
try
  AssignFile(actLOG, '\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log');
  if fileexists('\\192.168.180.180\data_traffic\SGI\SGI\SGI_LOG\' + strUserName +'_'+ FormatDateTime('ddmmyyyy',today)+'.log')
    then append(actLOG)
    else Rewrite(actLOG);
except
     on E : Exception do
     begin
       showmessage('Maaf, terdapat kesalahan dalam penyimpanan LOG, mohon periksa kondisi jaringan anda!' +sLineBreak+''+sLineBreak+'Terima Kasih' );
       frmChannel.Close ;
     end;
end;     
if ngChannel.Cells[15,y] = '1' then
begin
  stbparing:=MessageDlg('Are you sure want to pairing the Set-Top Box?'+sLineBreak+'if you choose YES then the channel must be pairing with linked STB'+sLineBreak+'if you choose NO, the channel can view on any STB'+sLineBreak+'0 = UnPairing'+sLineBreak+'1 = Pairing', mtConfirmation, [mbYes, mbNo], 0);
end
else stbparing := mrNo;

  	if dm.DDL.Eof then
		begin
			strSQL := 'INSERT INTO SGI.M_CHANNEL ( ';
			strSQL := strSQL + 'MCID, MCHANNEL, MCHANNEL_DESCRIPTION, MCSISERVICEID, ';
			strSQL := strSQL + 'MSGINAME, MUSERNIBBLE1, MUSER_CREATE, MUSER_CREATEDATE, ';
			strSQL := strSQL + 'MUSER_UPDATE, MUSER_UPDATEDATE, MSYNOPSIS_STATUS, MCHANNEL_NUMBER, ';
      {if stbparing = mrYes  then }strSQL := strSQL + 'MPLAYOUT_SOURCE, MSTB_PAIRING, MCH_ACTIVE )';
     // if stbparing = mrNo  then strSQL := strSQL + 'MPLAYOUT_SOURCE )';
			strSQL := strSQL + 'VALUES ( ';
			strSQL := strSQL + '''' + angka + ''', ';
			strSQL := strSQL + '''' + ngChannel.Cells[1,i] + ''', ';
			strSQL := strSQL + '''' + ngChannel.Cells[2,i] + ''', ';
			strSQL := strSQL + '(select catcode from m_catemplate ';
			strSQL := strSQL + 'where catdescription =''' + ngChannel.Cells[3,i] + '''), ';
			strSQL := strSQL + '''' + ngChannel.Cells[4,i] + ''', ';
			strSQL := strSQL + '''' + ngChannel.Cells[5,i] + ''', ';
			strSQL := strSQL + '''' + strUser +  ''', ';
			strSQL := strSQL + 'sysdate, ';
			strSQL := strSQL + '''' + strUser +  ''', ';
			strSQL := strSQL + 'sysdate, ';
			strSQL := strSQL + '''' + ngChannel.Cells[6,i] + ''', ';
			strSQL := strSQL + '''' + ngChannel.Cells[14,i] + ''', ';
			{if stbparing = mrYes  then }strSQL := strSQL + '''' + ngChannel.Cells[13,i] + ''', ';
 			{if stbparing = mrNo  then }strSQL := strSQL + '''' + ngChannel.Cells[15,i] + ''', ';
			if ngChannel.Cells[16,i] = 'True' then strSQL := strSQL + '''1'' ) '
      else strSQL := strSQL + '''0'' ) ';
			RecExc(strSQL);

      //catat log
      Writeln(actLOG,'[', FormatDateTime('c',today),'] ', ' Insert New Channel : ', ngChannel.Cells[1,i] , 'berhasil');
      Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' Insert New Channel : ', ngChannel.Cells[1,i] , 'berhasil');
      ShowMessage('New Channel Has been Inserted!');
		end
	else
		begin
      OldCh:= dm.DDL.FieldValues['MCHANNEL'];
			strSQL := ' UPDATE M_CHANNEL SET MCHANNEL = ''' + ngChannel.Cells[1,i] + ''', MCHANNEL_DESCRIPTION = ''' + ngChannel.Cells[2,i] + ''', ';
			strSQL := strSQL + ' MCHANNEL_NUMBER = ''' + ngChannel.Cells[14,i] + ''', ';
			strSQL := strSQL + ' MPLAYOUT_SOURCE = ''' + ngChannel.Cells[13,i] + ''', ';
      if ngChannel.Cells[16,i] = 'True' then strSQL := strSQL + ' MCH_ACTIVE = ''1'', '
      else strSQL := strSQL + ' MCH_ACTIVE = ''0'', ';
 			//strSQL := strSQL + ' MCH_ACTIVE = ''' + ngChannel.Cells[16,i] + ''', ';
			if stbparing = mrYes then strSQL := strSQL + ' MSTB_PAIRING = ''' + ngChannel.Cells[15,i] + ''', '
      else if stbparing = mrNo then strSQL := strSQL + ' MSTB_PAIRING = ''0'', '
      else strSQL := strSQL + ' MSTB_PAIRING = ''' + ngChannel.Cells[15,i] + ''', ';
			strSQL := strSQL + ' MCSISERVICEID = (select catcode from m_catemplate ';
			strSQL := strSQL + 'where catdescription = ''' + ngChannel.Cells[3,i] + '''), ';
			strSQL := strSQL + ' MSGINAME = ''' + ngChannel.Cells[4,i] + ''', ';
			strSQL := strSQL + ' MUSERNIBBLE1 = ''' + ngChannel.Cells[5,i] + ''', ';
			strSQL := strSQL + ' MUSER_UPDATE = ''' + strUser + ''', ';
			strSQL := strSQL + ' MUSER_UPDATEDATE = sysdate, ';
			strSQL := strSQL + 'MSYNOPSIS_STATUS = ''' + ngChannel.Cells[6,i] + ''' ';
			strSQL := strSQL + ' WHERE MCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
			strSQL := strSQL + 'AND MCSISERVICEID = (select catcode from m_catemplate ';
			strSQL := strSQL + 'where catdescription = ''' + ngChannel.Cells[9,i] + ''') ';
			RecExc(strSQL);

      if OldCh <> ngChannel.Cells[1,i] then
        begin
      //catat log
          Writeln(actLOG,'[', FormatDateTime('c',today),'] ', ' Update Channel : ', OldCh , ' menjadi ', ngChannel.Cells[1,i] , ' berhasil' );
          Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' Update Channel : ', OldCh, ' menjadi ', ngChannel.Cells[1,i] , ' berhasil');
          ShowMessage('Channel '+OldCh+ ' has been changed to '+ ngChannel.Cells[1,i]+' !');
        end
      else
        begin
          Writeln(actLOG,'[', FormatDateTime('c',today),'] ', ' Update Channel : ', ngChannel.Cells[1,i] , ' berhasil' );
          Writeln(actLOGLocal,'[', FormatDateTime('c',today),'] ', ' Update Channel : ', ngChannel.Cells[1,i] , ' berhasil');
          ShowMessage('Channel '+ngChannel.Cells[1,i]+' has been updated!');
        end;
		end;




  strSQL := 'SELECT * FROM M_READXL ';
  strSQL := strSQL + 'WHERE RCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
  RecSet(strSQL);
  if not dm.DDL.Eof then
  begin
   strSQL := ' UPDATE M_READXL SET RCHANNEL = ''' + ngChannel.Cells[1,i] + ''' ';
   strSQL := strSQL + ' WHERE RCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
   RecExc(strSQL);
  end;

  strSQL := 'SELECT * FROM M_CA_PACKAGE ';
  strSQL := strSQL + 'WHERE CACHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
  RecSet(strSQL);
  if not dm.DDL.Eof then
  begin
   strSQL := ' UPDATE M_CA_PACKAGE SET CACHANNEL = ''' + ngChannel.Cells[1,i] + ''' ';
   strSQL := strSQL + ' WHERE CACHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
   RecExc(strSQL);
  end;


 end;

 ngChannel.ClearRows;
 CloseFile(actLOG);
 CloseFile(actLOGLocal);
 FormShow(Sender);
 prcShow(cbChannel.Text);
 Screen.Cursor:=crDefault;
 exit;


errortext:
  ShowMessage(txterror);
  exit;
     
end;

procedure TfrmChannel.ngChannelMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if Button = mbRight Then
    Begin
      PopupMenu1.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmChannel.Button1Click(Sender: TObject);
begin
 frmChannel.Close;
end;

procedure TfrmChannel.ngChannelSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
  X := ACol;
  Y := ARow;
end;

procedure TfrmChannel.Delete1Click(Sender: TObject);
Var
  strDelete : String;
  buttonSelected : Integer;
begin
  buttonSelected := messagedlg('Apakah anda ingin menghapus channel '+ngChannel.Cells[1,y]+' ?',mtWarning , [mbYes, mbNo] , 0);
  if buttonSelected = mrYes     then
    begin
    buttonSelected := messagedlg('Apakah anda benar-benar ingin menghapus channel '+ngChannel.Cells[1,y]+' ?',mtError , [mbYes, mbNo] , 0);
    if buttonSelected = mrYes     then
      begin
       buttonSelected := messagedlg('Sekali lagi, Apakah anda benar-benar sangat ingin menghapus channel '+ngChannel.Cells[1,y]+' ?',mtConfirmation , [mbYes, mbNo] , 0);
        if buttonSelected = mrYes     then
          begin
           strDelete := ngChannel.Cells[1,y];
           strSQL := 'DELETE FROM M_CHANNEL WHERE MCHANNEL = ''' + ngChannel.Cells[1,y] + '''  ';
           RecExc(strSQL);
           strSQL := 'DELETE FROM M_CA_PACKAGE WHERE CACHANNEL = ''' + ngChannel.Cells[1,y] + '''  ';
           RecExc(strSQL);
           FormShow(Sender);
           prcShow(cbChannel.Text);
           //messagedlg('Channel '+ngChannel.Cells[1,y]+' sudah dihapus',mtConfirmation , mbOKCancel, 0);
           //   if buttonSelected = mrOK     then
           //     begin
           ShowMessage('CHANNEL '+strDelete+' Has Been Removed');
           //     end;
          end;
      end;
    end;
end;

procedure TfrmChannel.CCINumber1Click(Sender: TObject);
begin
frmCCIBits.Show;
end;


//procedure TfrmChannel.ngChannelMouseLeave(Sender: TObject);
//  var
//  i:integer;
//  strballoon:string;
//  myWideString:WideString;
//  myWideCharPtr:PWideChar;
// 	const
//		sLineBreak = {$IFDEF LINUX} AnsiChar(#10) {$ENDIF}
//			{$IFDEF MSWINDOWS} AnsiString(#13#10) {$ENDIF};
//begin
//for i := 0 to ngChannel.RowCount-1 do
// begin
//  strSQL := 'SELECT CATCODE FROM M_CATEMPLATE ';
//  strSQL := strSQL + 'WHERE catdescription = ''' + ngChannel.Cells[3,i] + ''' ';
//  RecSet2(strSQL);
//    if (dm.DDL2.FieldValues['CATCODE'] = '2002') or (dm.DDL2.FieldValues['CATCODE'] = '2202')  or (dm.DDL2.FieldValues['CATCODE'] = '2005')  or (dm.DDL2.FieldValues['CATCODE'] = '100') or (dm.DDL2.FieldValues['CATCODE'] = '200')  then
//      begin
//        strSQL := 'SELECT * FROM M_CHANNEL_BITS ';
//        strSQL := strSQL + 'WHERE CBCHANNEL = ''' + ngChannel.Cells[7,i] + ''' ';
//        RecSet2(strSQL);
//
//            strballoon := 'CCI Value : ' + dm.DDL2.FieldValues['CBNUMBER'];
//            myWideString  := strballoon;
//            myWideCharPtr := Addr(myWideString[1]);
//
//        if not dm.DDL2.Eof then
//          begin
//           ShowBalloonTip(ngchannel, 0, '', myWideCharPtr , clBlue, clNavy);
//          end;
//      end
//    else ShowBalloonTip(ngchannel,0, '', ' ' , clBlue, clNavy);
// end;
//end;

procedure TfrmChannel.chkHDChClick(Sender: TObject);
begin
 Screen.Cursor:=crHourGlass;
 ngChannel.ClearRows;
 prcShow(cbChannel.Text);
 Screen.Cursor:=crDefault;
end;

end.
