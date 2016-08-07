unit frm_VOD;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid, Menus, NxColumnClasses, NxColumns, StdCtrls, ComCtrls, jpeg,
  Grids, Buttons;

type
  TfrmVOD = class(TForm)
    ppmVOD: TPopupMenu;
    AddRow1: TMenuItem;
    AddNew1: TMenuItem;
    Update1: TMenuItem;
    Exit1: TMenuItem;
    ScrollBox1: TScrollBox;
    Image1: TImage;
    ScrollBox2: TScrollBox;
    ngVOD: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    edtEpg: TEdit;
    Label1: TLabel;
    pbVOD: TProgressBar;
    Delete1: TMenuItem;
    StringGrid1: TStringGrid;
    btnImport: TBitBtn;
    OpenDialog1: TOpenDialog;
    BitBtn1: TBitBtn;
    NxTextColumn9: TNxTextColumn;
    cboFilterVOD: TComboBox;
    NxTextColumn13: TNxTextColumn;
    btnJustCA: TBitBtn;
    grpJustCA: TGroupBox;
    Shape2: TShape;
    ngJustCA: TNextGrid;
    NxTextColumn15: TNxTextColumn;
    NxTextColumn16: TNxTextColumn;
    NxTextColumn17: TNxTextColumn;
    NxTextColumn18: TNxTextColumn;
    NxTextColumn19: TNxTextColumn;
    BitBtn2: TBitBtn;
    Label2: TLabel;
    cboCAPStart: TComboBox;
    cboCAPEnd: TComboBox;
    Shape3: TShape;
    Label3: TLabel;
    Label4: TLabel;
    NxTextColumn20: TNxTextColumn;
    NxTextColumn21: TNxTextColumn;
    NxTextColumn22: TNxTextColumn;
    NxTextColumn23: TNxTextColumn;
    NxTextColumn24: TNxTextColumn;
    NxTextColumn25: TNxTextColumn;
    NxTextColumn26: TNxTextColumn;
    NxTextColumn27: TNxTextColumn;
    NxTextColumn28: TNxTextColumn;
    NxTextColumn29: TNxTextColumn;
    Panel1: TPanel;
    Panel2: TPanel;
    NxTextColumn5: TNxDateColumn;
    NxTextColumn8: TNxDateColumn;
    NxTextColumn3: TNxComboBoxColumn;
    NxTextColumn12: TNxNumberColumn;
    NxTextColumn7: TNxComboBoxColumn;
    NxTextColumn11: TNxComboBoxColumn;
    NxTextColumn14: TNxNumberColumn;
    procedure ngVODMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure AddRow1Click(Sender: TObject);
//    procedure FormResize(Sender: TObject);
    procedure Update1Click(Sender: TObject);
    procedure AddNew1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure edtEpgKeyPress(Sender: TObject; var Key: Char);
    procedure ngVODSelectCell(Sender: TObject; ACol,
      ARow: Integer);
    procedure Delete1Click(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure btnJustCAClick(Sender: TObject);
    procedure edtEpgClick(Sender: TObject);
    procedure cboCAPEndSelect(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure cboFilterVODClick(Sender: TObject);
    procedure ngVODClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
   // procedure edtEpgChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  X, Y : Integer;
  frmVOD: TfrmVOD;
  function fncAutoID():string;
  Procedure prcShowVOD(epg:string);
  function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
  function Replace(Dest, SubStr, Str: string): string;

  function fncVODCAProductID():String;
  function fncAutoCAServiceID():string;
  function fncAutoProgramID():string;
  function fncAutoProgramKey():string;
  
implementation

uses ComObj, frm_dm, frm_Login, mdl_Global, frm_EPG;

{$R *.dfm}

function SetCueBanner(const Edit: TEdit;
const Placeholder: String): Boolean;
const
  EM_SETCUEBANNER = $1501;
var
  UniStr: WideString;
begin
  UniStr := Placeholder;
  SendMessage(Edit.Handle, EM_SETCUEBANNER, WParam(True),LParam(UniStr));
  Result := GetLastError() = ERROR_SUCCESS;
end;

function fncAutoID():string;
var
 strID : integer;
begin
 strSQL := 'SELECT VODID FROM M_VOD ORDER BY VODID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strID := 1;
  end
 else
  begin
   dm.ddl.Last;
   strID := dm.DDL.FieldValues['VODID'] + 1;
  end;
 fncAutoID:=IntToStr(strID);
end;

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

Procedure prcShowVOD(epg:string);
var
 strSQL : String;
 grpKey, prgKey : String;
 i,ii : integer;
 xyz : byte;
 dblVodDateStart, dblVodDateEnd  : Double;
 VodDateStart, VodDateEnd  : TDateTime;

begin

 frmVOD.ngVOD.ClearRows;
 strSQL := 'select Jumlah_VOD, VODEPGTITLE, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, VODPROGRAMID, VODTRAFFICKEY, ';
 strSQL := strSQL + 'VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, VODAMOUNT ';
 strSQL := strSQL + 'FROM (SELECT COUNT(0) AS Jumlah_VOD FROM M_VOD WHERE ';
 //strSQL := strSQL + ' VODEPGTITLE LIKE ''%' + epg + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Tittle'
    Then strSQL := strSQL + '      VODEPGTITLE LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'CA Product ID'
    Then strSQL := strSQL + '      VODCAPRODUCTID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'CA Start Date'
    Then strSQL := strSQL + '      VODCAPSTARTDATE = To_Date(''' + AnsiUpperCase(epg) + ''',''MM/dd/yyyy'') ';
 if frmVOD.cboFilterVOD.Text = 'CA End Date'
    Then strSQL := strSQL + '      VODCAPENDDATE = To_Date(''' + AnsiUpperCase(epg) + ''',''MM/dd/yyyy'') ';
 if frmVOD.cboFilterVOD.Text = 'CA Service ID'
    Then strSQL := strSQL + '      VODCASERVICEID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Program ID'
    Then strSQL := strSQL + '      VODPROGRAMID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Traffic Key'
    Then strSQL := strSQL + '      VODTRAFFICKEY LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Group Key'
    Then strSQL := strSQL + '      VODGROUPKEY LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Program Key'
    Then strSQL := strSQL + '      VODPROGRAMKEY LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'FED'
    Then strSQL := strSQL + '      VODFED LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Time Offset'
    Then strSQL := strSQL + '      VODTIMEOFFSET LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Status'
    Then strSQL := strSQL + '      VODSTATUS LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Amount'
    Then strSQL := strSQL + '      VODAMOUNT LIKE ''%' + AnsiUpperCase(epg) + '%'' ';

 strSQL := strSQL + ' )xxx, ';
 strSQL := strSQL + '(SELECT VODEPGTITLE, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, VODPROGRAMID, ';
 strSQL := strSQL + 'VODTRAFFICKEY, VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, VODAMOUNT FROM M_VOD WHERE ';
 //strSQL := strSQL + ' VODEPGTITLE LIKE ''%' + epg + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Tittle'
    Then strSQL := strSQL + '      VODEPGTITLE LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'CA Product ID'
    Then strSQL := strSQL + '      VODCAPRODUCTID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'CA Start Date'
    Then strSQL := strSQL + '      VODCAPSTARTDATE = To_Date(''' + AnsiUpperCase(epg) + ''',''MM/dd/yyyy'') ';
 if frmVOD.cboFilterVOD.Text = 'CA End Date'
    Then strSQL := strSQL + '      VODCAPENDDATE = To_Date(''' + AnsiUpperCase(epg) + ''',''MM/dd/yyyy'') ';
 if frmVOD.cboFilterVOD.Text = 'CA Service ID'
    Then strSQL := strSQL + '      VODCASERVICEID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Program ID'
    Then strSQL := strSQL + '      VODPROGRAMID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Traffic Key'
    Then strSQL := strSQL + '      VODTRAFFICKEY LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Group Key'
    Then strSQL := strSQL + '      VODGROUPKEY LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Program Key'
    Then strSQL := strSQL + '      VODPROGRAMKEY LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'FED'
    Then strSQL := strSQL + '      VODFED LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Time Offset'
    Then strSQL := strSQL + '      VODTIMEOFFSET LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmVOD.cboFilterVOD.Text = 'Status'
    Then strSQL := strSQL + '      VODSTATUS LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
  if frmVOD.cboFilterVOD.Text = 'Amount'
    Then strSQL := strSQL + '      VODAMOUNT LIKE ''%' + AnsiUpperCase(epg) + '%'' ';

 strSQL := strSQL + ' )yyy ';
 strSQL := strSQL + 'ORDER BY VODEPGTITLE ';
 RecSet2(strSQL);

 xyz := dm.DDL2.FieldValues['Jumlah_VOD'];

 i:=1;
 if not VarIsNull(dm.DDL2.FieldValues['Jumlah_VOD']) then
  begin
   ii:=dm.DDL2.FieldValues['Jumlah_VOD'];
  end
 else
  begin
   ii:=0;
  end;
 frmVOD.ngVOD.AddRow(ii);
 frmVOD.pbVOD.Min:=0;
 frmVOD.ngVOD.BeginUpdate;
 frmVOD.pbVOD.Visible:=True;
 frmVOD.pbVOD.Max:=ii;
 while not dm.DDL2.Eof do
  begin

   if not VarIsNull(dm.DDL2.FieldValues['VODGROUPKEY'])then
    begin
     grpKey:=dm.DDL2.FieldValues['VODGROUPKEY'];
    end
   else grpKey:='';

   if not VarIsNull(dm.DDL2.FieldValues['VODPROGRAMKEY'])then
    begin
     prgKey:=dm.DDL2.FieldValues['VODPROGRAMKEY'];
    end
   else prgKey:='';

   frmVOD.ngVOD.Cell[0, i-1].AsString := inttostr(i);
   frmVOD.ngVOD.Cell[1, i-1].AsString := dm.DDL2.FieldValues['VODEPGTITLE'];
   //frmVOD.ngVOD.Cell[2, i-1].AsString := dm.DDL2.FieldValues['VODCAPRODUCTID'];
   if VarIsNull(dm.DDL2.FieldValues['VODCAPRODUCTID'])
    then frmVOD.ngVOD.Cell[2, i-1].AsString := ''
     Else frmVOD.ngVOD.Cell[2, i-1].AsString := dm.DDL2.FieldValues['VODCAPRODUCTID'];;
   frmVOD.ngVOD.Cell[3, i-1].AsString := dm.DDL2.FieldValues['VODCAPSTARTDATE'];
   frmVOD.ngVOD.Cell[4, i-1].AsString := dm.DDL2.FieldValues['VODCAPENDDATE'];
   //frmVOD.ngVOD.Cell[5, i-1].AsString := dm.DDL2.FieldValues['VODCASERVICEID'];
   if VarIsNull(dm.DDL2.FieldValues['VODCASERVICEID'])
    then frmVOD.ngVOD.Cell[5, i-1].AsString := ''
     Else frmVOD.ngVOD.Cell[5, i-1].AsString := dm.DDL2.FieldValues['VODCASERVICEID'];;
   //frmVOD.ngVOD.Cell[6, i-1].AsString := dm.DDL2.FieldValues['VODPROGRAMID'];
   if VarIsNull(dm.DDL2.FieldValues['VODPROGRAMID'])
    then frmVOD.ngVOD.Cell[6, i-1].AsString := ''
     Else frmVOD.ngVOD.Cell[6, i-1].AsString := dm.DDL2.FieldValues['VODPROGRAMID'];;
   frmVOD.ngVOD.Cell[7, i-1].AsString := prgKey;
   frmVOD.ngVOD.Cell[8, i-1].AsString := grpKey;
   frmVOD.ngVOD.Cell[9, i-1].AsString := prgKey;
   frmVOD.ngVOD.Cell[10, i-1].AsString := dm.DDL2.FieldValues['VODFED'];
   frmVOD.ngVOD.Cell[11, i-1].AsString := dm.DDL2.FieldValues['VODTIMEOFFSET'];
   //frmVOD.ngVOD.Cell[12, i-1].AsString := dm.DDL2.FieldValues['VODSTATUS'];

   if VarIsNull(dm.DDL2.FieldValues['VODSTATUS'])
    then frmVOD.ngVOD.Cell[12, i-1].AsString := ''
     Else frmVOD.ngVOD.Cell[12, i-1].AsString := dm.DDL2.FieldValues['VODSTATUS'];

   frmVOD.ngVOD.Cell[13, i-1].AsString := dm.DDL2.FieldValues['VODAMOUNT'];;

   i:=i+1;
   frmVOD.pbVOD.Position:=i;
   dm.DDL2.Next;
  end;
 frmVOD.ngVOD.EndUpdate;
 frmVOD.pbVOD.Visible:=False;
end;

procedure TfrmVOD.FormShow(Sender: TObject);
begin
SetCueBanner(edtEPG, 'Input search keyword Here, then press enter');
end;
procedure TfrmVOD.ngVODMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if Button = mbRight Then
    Begin
      ppmVOD.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmVOD.AddRow1Click(Sender: TObject);
begin      
 ngVOD.AddRow(1);
end;

procedure TfrmVOD.Update1Click(Sender: TObject);
var
 i : integer;
 epgReplace, VODCAServiceID, VODProgramID, grpKey, prgKey, angka, strCAProductID, strCode : String;
 VODCaProduct, VODFED, VODStatus, VODAmount, strTimeOffset : String;
 VodDateStart, VodDateEnd : TDateTime;
begin
 for i := 0 to ngVOD.RowCount-1 do
 begin
  angka:=fncAutoID;
  strSQL := 'SELECT * FROM M_VOD ';
  strSQL := strSQL + 'WHERE VODEPGTITLE = ''' + ngVOD.Cells[1,i] + ''' ';
  RecSet(strSQL);

   epgReplace:=Replace(trim(ngVOD.Cells[1,i]), '''','');
   epgReplace:=Replace(trim(epgReplace), '"','');

   VODCaProduct := ngVOD.Cells[2,i];
   VodDateStart := StrToDateTime(ngVOD.Cells[3,i]);
   VodDateEnd := StrToDateTime(ngVOD.Cells[4,i]);
   VODCAServiceID := ngVOD.Cells[5,i];
   VODProgramID := ngVOD.Cells[6,i];

   {if not VarIsNull(ngVOD.Cells[8,i])then
    begin
     grpKey:=Replace(trim(ngVOD.Cells[8,i]), '''','`');
     grpKey:=Replace(trim(grpKey), '`','''''');
     grpKey:=Replace(trim(grpKey), '"','``');
     grpKey:=Replace(trim(grpKey),#13,'');
    end
   else grpKey:='';}

   grpKey := ngVOD.Cells[8,i];

   if not VarIsNull(ngVOD.Cells[9,i])then
    begin
     prgKey:=Replace(trim(ngVOD.Cells[9,i]), '''','`');
     prgKey:=Replace(trim(prgKey), '`','''''');
     prgKey:=Replace(trim(prgKey), '"','``');
    end
   else prgKey:='';

   VODFED := ngVOD.Cells[10,i];
                                       //
   strTimeOffset := ngVOD.Cells[11,i];
   strTimeOffset := Format('%.*d',[6, strtoint(strTimeOffset)]);

   VODStatus := ngVOD.Cells[12,i];

   VODAmount := ngVOD.Cells[13,i];

  if dm.DDL.Eof then
  begin

   If (grpKey = '12346') or (grpKey = '') then
    Begin
     strSQL := 'SELECT MAX(SubSTR(VODCAPRODUCTID, 5, 6)) + 1 AS VODCAPRODUCTID FROM M_VOD ';
     RecSet(strSQL);
     strCAProductID := dm.DDL.FieldValues['VODCAPRODUCTID'];

     if Length(strCAProductID) = 1 then strCode := 'OPPV00000' + strCAProductID;
     if Length(strCAProductID) = 2 then strCode := 'OPPV0000' + strCAProductID;
     if Length(strCAProductID) = 3 then strCode := 'OPPV000' + strCAProductID;
     if Length(strCAProductID) = 4 then strCode := 'OPPV00' + strCAProductID;
     if Length(strCAProductID) = 5 then strCode := 'OPPV0' + strCAProductID;
     if Length(strCAProductID) = 6 then strCode := 'OPPV' + strCAProductID;
    End
   Else strCode := '';

   //IDCAServiceID:= fncAutoCAServiceID;
   If (grpKey = '12346') or (grpKey='') then
    Begin
      VODCAServiceID := fncAutoCAServiceID;
    End
   Else VODCAServiceID := '';

   //IDProgramID:= fncAutoProgramID;
   If (grpKey = '12346') or (grpKey='') then
    Begin
      VODProgramID := fncAutoProgramID;
    End
   Else VODProgramID := '';

   prgKey:= fncAutoProgramKey;
   strSQL := 'INSERT INTO SGI.M_VOD ( ';
   strSQL := strSQL + 'VODID, VODEPGTITLE, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, VODPROGRAMID, VODTRAFFICKEY, ';
   strSQL := strSQL + 'VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, VODAMOUNT, VODUSERCREATE, VODUSERCREATEDATE, ';
   strSQL := strSQL + 'VODUSERUPDATE, VODUSERUPDATEDATE) ';
   strSQL := strSQL + 'VALUES ( ';
   strSQL := strSQL + '''' + angka + ''', ';
   strSQL := strSQL + '''' + epgReplace + ''', ';
   strSQL := strSQL + '''' + strCode +  ''', ';
   strSQL := strSQL + 'TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateStart) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
   strSQL := strSQL + 'TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateEnd) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
   strSQL := strSQL + '''' + VODCAServiceID + ''', ';
   strSQL := strSQL + '''' + VODProgramID + ''', ';
   strSQL := strSQL + '''' + prgKey + ''', ';
   strSQL := strSQL + '''' + grpKey + ''', ';
   strSQL := strSQL + '''' + prgKey + ''', ';
   strSQL := strSQL + '''' + VODFED + ''', ';
   strSQL := strSQL + '''' + strTimeOffset + ''', ';
   strSQL := strSQL + '''' + VODStatus + ''', ';
   strSQL := strSQL + '''' + VODAmount + ''', ';
   strSQL := strSQL + '''' + strUser + ''', ';
   strSQL := strSQL + 'sysdate, ';
   strSQL := strSQL + '''' + strUser + ''', ';
   strSQL := strSQL + 'sysdate) ';
   RecExc(strSQL);

  end
  else
  begin
   strSQL := ' UPDATE M_VOD SET VODEPGTITLE = ''' + ngVOD.Cells[1,i] + ''', ';
   //strSQL := strSQL + ' VODEPGTITLE = ''' + epgReplace +  ''', ';
   strSQL := strSQL + ' VODCAPRODUCTID = ''' + VODCaProduct +  ''', ';
   strSQL := strSQL + ' VODCAPSTARTDATE = TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateStart) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
   strSQL := strSQL + ' VODCAPENDDATE = TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateEnd) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
   strSQL := strSQL + ' VODCASERVICEID = ''' + VODCAServiceID + ''', ';
   strSQL := strSQL + ' VODPROGRAMID = ''' + VODProgramID + ''', ';
   strSQL := strSQL + ' VODTRAFFICKEY = ''' + prgKey + ''', ';
   strSQL := strSQL + ' VODGROUPKEY = ''' + grpKey + ''', ';
   strSQL := strSQL + ' VODPROGRAMKEY = ''' + prgKey + ''', ';
   strSQL := strSQL + ' VODFED = ''' + VODFED + ''', ';
   strSQL := strSQL + ' VODTIMEOFFSET = ''' + strTimeOffset + ''', ';
   strSQL := strSQL + ' VODSTATUS = ''' + VODStatus + ''', ';
   strSQL := strSQL + ' VODAMOUNT = ''' + VODAmount + ''', ';
   strSQL := strSQL + ' VODUSERUPDATE = ''' + strUser + ''', ';
   strSQL := strSQL + ' VODUSERUPDATEDATE = sysdate ';
   strSQL := strSQL + ' WHERE VODEPGTITLE = ''' + epgReplace + ''' ';

   RecExc(strSQL);
  end;

 end;
 ngVOD.ClearRows;
 prcShowVOD(edtEpg.Text);
 ShowMessage('Data Has Been Saved!');
end;

procedure TfrmVOD.AddNew1Click(Sender: TObject);
begin
 ngVOD.ClearRows;
 ngVOD.AddRow(1);
end;      

procedure TfrmVOD.Exit1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmVOD.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
 frmVOD.Close;
end;

procedure TfrmVOD.edtEpgKeyPress(Sender: TObject;
  var Key: Char);
begin
if key=#13 then
  begin
    Screen.Cursor:=crHourGlass;
    ngVOD.ClearRows;
    if trim(edtEpg.text)<>'' then
    begin
      prcShowVOD(edtEpg.Text);
    end;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmVOD.ngVODSelectCell(Sender: TObject;
  ACol, ARow: Integer);
begin
  X := ACol;
  Y := ARow;
end;

procedure TfrmVOD.Delete1Click(Sender: TObject);
var
  strDelSys : String;
begin
   Screen.Cursor := crHourGlass;
   strDelSys := ngVOD.Cells[1,y];
   strSQL := 'DELETE FROM M_VOD WHERE VODEPGTITLE = '''+trim(ngVOD.Cells[1,y])+''' ';
   RecExc(strSQL);
   prcShowVOD(edtEpg.Text);
   ShowMessage('VOD Component '+strDelSys+' Has Been Removed');
   Screen.Cursor := crDefault;
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

    // Assign the Variant associated with the WorkSheet to the Delphi Variant

    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].value;
    //  Define the loop for filling in the TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
      Inc(k, 1);
      AGrid.RowCount := k + 1;
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

function fncVODCAProductID():String;
  var
    strNomor : string;
    intNomor : integer;

  begin
    strSQL := 'SELECT VODCAPRODUCTID ';
    strSQL := strSQL + 'FROM M_VOD ';
    strSQL := strSQL + 'WHERE VODCAPRODUCTID LIKE ''OPPV%''';
    strSQL := strSQL + 'ORDER BY VODCAPRODUCTID DESC';
    RecSet(strSQL);
    if dm.DDL.Eof then
      begin
        strNomor := '0001';
      end
    else
      begin
        dm.ddl.First;
        intNomor := StrToInt(copy(dm.DDL.FieldValues['VODCAPRODUCTID'], 5, 4)) + 1;
        if Length(IntToStr(intNomor)) = 1 then strNomor := '000' + IntToStr(intNomor);
        if Length(IntToStr(intNomor)) = 2 then strNomor := '00' + IntToStr(intNomor);
        if Length(IntToStr(intNomor)) = 3 then strNomor := '0' + IntToStr(intNomor);
        if Length(IntToStr(intNomor)) = 4 then strNomor := IntToStr(intNomor);
      end;
    fncVODCAProductID := 'OPPV'+strNomor;
end;


function fncAutoCAServiceID() : String;
var
  strSQLID : String;
  intID    : Integer;
begin
  strSQLID := 'SELECT MAX(VODCASERVICEID) AS VODCASERVICEID FROM M_VOD' ;
  RecSetIDTable(strSQLID);

  fncAutoCAServiceID := '1';

  if not dm.DDLIDTable.Eof Then
    Begin
      if not VarIsNull(dm.DDLIDTable.FieldValues['VODCASERVICEID']) Then
        Begin
          intID := StrToInt(dm.DDLIDTable.FieldValues['VODCASERVICEID']) + 1;
          fncAutoCAServiceID := IntToStr(intID);
        End;
    End;
end;

function fncAutoProgramID():string;
var
  strSQLID : String;
  intID    : Integer;
begin
  strSQLID := 'SELECT MAX(VODPROGRAMID) AS VODPROGRAMID FROM M_VOD' ;
  RecSetIDTable(strSQLID);

  fncAutoProgramID := '1';

  if not dm.DDLIDTable.Eof Then
    Begin
      if not VarIsNull(dm.DDLIDTable.FieldValues['VODPROGRAMID']) Then
        Begin
          intID := StrToInt(dm.DDLIDTable.FieldValues['VODPROGRAMID']) + 1;
          fncAutoProgramID := IntToStr(intID);
        End;
    End;
end;

function fncAutoProgramKey():string;
var
 strID : integer;
begin
 strSQL := 'SELECT VODPROGRAMKEY FROM M_VOD ORDER BY VODPROGRAMKEY ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strID := 1;
  end
 else
  begin
   dm.ddl.Last;
   strID := dm.DDL.FieldValues['VODPROGRAMKEY'] + 1;
  end;
 fncAutoProgramKey:=IntToStr(strID);
end;

procedure TfrmVOD.btnImportClick(Sender: TObject);
var
 i : integer;
 IDTable, IDCAServiceID, IDProgramID, strCAProductID, strCode, epgReplace : String;
 VODCAServiceID, VODProgramID, grpKey, prgKey, VODFED, VODStatus, strTimeOffset, VODAmount : String;
 dblVodDateStart, dblVodDateEnd, dblTimeOffset  : Double;
 VodDateStart, VodDateEnd, VodTimeOffset  : TDateTime; 
 //SGI : TextFile;

begin

 if OpenDialog1.Execute then
 begin
 i := 1;
 Screen.Cursor:=crHourGlass;


  if Xls_To_StringGrid(StringGrid1, OpenDialog1.FileName) then
    begin
     pbVOD.Max := StringGrid1.RowCount - 2;
     pbVOD.Min := 0;
     pbVOD.Visible := True;

      while StringGrid1.Cells[0, i]<>'<eof>' do
        begin
          IDTable:=fncAutoID;

          prgKey:= fncAutoProgramKey;

          epgReplace:=Replace(trim(StringGrid1.Cells[1,i]), '''','');
          epgReplace:=Replace(trim(epgReplace), '"','');

          VodDateStart := StrToDateTime(StringGrid1.Cells[2,i]);
          VodDateEnd := StrToDateTime(StringGrid1.Cells[3,i]);
          //VODCAServiceID := StringGrid1.Cells[4,i];
          //VODProgramID := StringGrid1.Cells[5,i];

          {if not VarIsNull(StringGrid1.Cells[4,i])then
            begin
              grpKey:=Replace(trim(StringGrid1.Cells[4,i]), '''','`');
              grpKey:=Replace(trim(grpKey), '`','''''');
              grpKey:=Replace(trim(grpKey), '�','"');
              grpKey:=Replace(trim(grpKey), '�','"');
            end
          else grpKey:='';}

          grpKey := StringGrid1.Cells[4,i];

          //IDCAServiceID:= fncAutoCAServiceID;
          If (grpKey = '12346') or (grpKey='') then
           Begin
            IDCAServiceID := fncAutoCAServiceID;
           End
          Else IDCAServiceID := '';

          //IDProgramID:= fncAutoProgramID;
          If (grpKey = '12346') or (grpKey='') then
           Begin
            IDProgramID := fncAutoProgramID;
           End
          Else IDProgramID := '';

          if not VarIsNull(prgKey)then
            begin
             prgKey:=Replace(trim(prgKey), '''','`');
             prgKey:=Replace(trim(prgKey), '`','''''');
             prgKey:=Replace(trim(prgKey), '�','"');
             prgKey:=Replace(trim(prgKey), '�','"');
            end
          else prgKey:='';

          VODFED := StringGrid1.Cells[5,i];

          //VodTimeOffset := StrToDateTime(StringGrid1.Cells[6,i]);
          dblTimeOffset := StrToFloat(StringGrid1.Cells[6,i]);
          strTimeOffset := FormatDateTime('HHMMss', dblTimeOffset);

          VODStatus := StringGrid1.Cells[7,i];
          VODAmount := StringGrid1.Cells[8,i];

          strSQL := 'SELECT VODEPGTITLE FROM M_VOD WHERE VODEPGTITLE = upper(''' + epgReplace + ''') ';
          //strSQL := strSQL + 'AND VODCASERVICEID = ''' + IDCAServiceID + ''' ';
          RecSet(strSQL);
          if not dm.DDL.Eof Then
              Begin
                MessageDlg('VOD CA Product already exist!', mtWarning, [mbok], 0);
                pbVOD.Visible := False;
                Screen.Cursor:=crDefault;
                Exit;
              End;

          if not dm.DDL.Eof then
            begin

             strSQL := 'delete FROM M_VOD WHERE VODEPGTITLE = upper(''' + epgReplace + ''') ';
             //strSQL := strSQL + 'AND VODCASERVICEID = ''' + IDCAServiceID + ''' ';
             RecExc(strSQL);

             If (grpKey = '12346') or (grpKey='') then
              Begin
               strSQL := 'SELECT MAX(SubSTR(VODCAPRODUCTID, 5, 6)) + 1 AS VODCAPRODUCTID FROM M_VOD ';
               RecSet(strSQL);
               strCAProductID := dm.DDL.FieldValues['VODCAPRODUCTID'];

               if Length(strCAProductID) = 1 then strCode := 'OPPV00000' + strCAProductID;
               if Length(strCAProductID) = 2 then strCode := 'OPPV0000' + strCAProductID;
               if Length(strCAProductID) = 3 then strCode := 'OPPV000' + strCAProductID;
               if Length(strCAProductID) = 4 then strCode := 'OPPV00' + strCAProductID;
               if Length(strCAProductID) = 5 then strCode := 'OPPV0' + strCAProductID;
               if Length(strCAProductID) = 6 then strCode := 'OPPV' + strCAProductID;
              End
             Else strCode := '';

              strSQL := 'INSERT INTO SGI.M_VOD ( ';
              strSQL := strSQL + 'VODID, VODEPGTITLE, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, VODPROGRAMID, VODTRAFFICKEY, ';
              strSQL := strSQL + 'VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, VODAmount, VODUSERCREATE, VODUSERCREATEDATE, ';
              strSQL := strSQL + 'VODUSERUPDATE, VODUSERUPDATEDATE) ';
              strSQL := strSQL + 'VALUES ( ';
              strSQL := strSQL + '''' + IDTable + ''', ';
              //strSQL := strSQL + '''' + epgReplace + ''', ';
              strSQL := strSQL + 'Upper(''' + epgReplace + '''), ';
              strSQL := strSQL + '''' + strCode +  ''', ';
              strSQL := strSQL + 'TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateStart) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
              strSQL := strSQL + 'TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateEnd) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
              strSQL := strSQL + '''' + IDCAServiceID + ''', ';
              strSQL := strSQL + '''' + IDProgramID + ''', ';
              strSQL := strSQL + '''' + prgKey + ''', ';
              strSQL := strSQL + '''' + grpKey + ''', ';
              strSQL := strSQL + '''' + prgKey + ''', ';
              strSQL := strSQL + '''' + VODFED + ''', ';
              strSQL := strSQL + '''' + strTimeOffset + ''', ';
              strSQL := strSQL + '''' + VODStatus + ''', ';
              strSQL := strSQL + '''' + VODAmount + ''', ';
              strSQL := strSQL + '''' + strUser + ''', ';
              strSQL := strSQL + 'sysdate, ';
              strSQL := strSQL + '''' + strUser + ''', ';
              strSQL := strSQL + 'sysdate) ';
              RecExc(strSQL);
            end;

           if dm.DDL.Eof then
            begin

             If (grpKey = '12346') or (grpKey ='') then
              Begin
               strSQL := 'SELECT MAX(SubSTR(VODCAPRODUCTID, 5, 6)) + 1 AS VODCAPRODUCTID FROM M_VOD ';
               RecSet(strSQL);
               strCAProductID := dm.DDL.FieldValues['VODCAPRODUCTID'];

               if Length(strCAProductID) = 1 then strCode := 'OPPV00000' + strCAProductID;
               if Length(strCAProductID) = 2 then strCode := 'OPPV0000' + strCAProductID;
               if Length(strCAProductID) = 3 then strCode := 'OPPV000' + strCAProductID;
               if Length(strCAProductID) = 4 then strCode := 'OPPV00' + strCAProductID;
               if Length(strCAProductID) = 5 then strCode := 'OPPV0' + strCAProductID;
               if Length(strCAProductID) = 6 then strCode := 'OPPV' + strCAProductID;
              End
             Else strCode := '';

             strSQL := 'INSERT INTO SGI.M_VOD ( ';
             strSQL := strSQL + 'VODID, VODEPGTITLE, VODCAPRODUCTID, VODCAPSTARTDATE, VODCAPENDDATE, VODCASERVICEID, VODPROGRAMID, VODTRAFFICKEY, ';
             strSQL := strSQL + 'VODGROUPKEY, VODPROGRAMKEY, VODFED, VODTIMEOFFSET, VODSTATUS, VODAMOUNT, VODUSERCREATE, VODUSERCREATEDATE, ';
             strSQL := strSQL + 'VODUSERUPDATE, VODUSERUPDATEDATE) ';
             strSQL := strSQL + 'VALUES ( ';
             strSQL := strSQL + '''' + IDTable + ''', ';
             //strSQL := strSQL + '''' + epgReplace + ''', ';
             strSQL := strSQL + 'Upper(''' + epgReplace + '''), ';
             strSQL := strSQL + '''' + strCode +  ''', ';
             strSQL := strSQL + 'TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateStart) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
             strSQL := strSQL + 'TO_DATE(''' + FormatDateTime('MM/dd/yyyy HH:MM:ss', VodDateEnd) + ''',''mm/dd/yyyy hh24:mi:ss''), ';
             strSQL := strSQL + '''' + IDCAServiceID + ''', ';
             strSQL := strSQL + '''' + IDProgramID + ''', ';
             strSQL := strSQL + '''' + prgKey + ''', ';
             strSQL := strSQL + '''' + grpKey + ''', ';
             strSQL := strSQL + '''' + prgKey + ''', ';
             strSQL := strSQL + '''' + VODFED + ''', ';
             strSQL := strSQL + '''' + strTimeOffset + ''', ';
             strSQL := strSQL + '''' + VODStatus + ''', ';
             strSQL := strSQL + '''' + VODAmount + ''', ';
             strSQL := strSQL + '''' + strUser + ''', ';
             strSQL := strSQL + 'sysdate, ';
             strSQL := strSQL + '''' + strUser + ''', ';
             strSQL := strSQL + 'sysdate) ';
             RecExc(strSQL);
            end;

          strSQL := 'SELECT * FROM M_VOD WHERE VODEPGTITLE = upper(''' + epgReplace + ''') ';
          RecSet(strSQL);
          while not dm.DDL.Eof do
            begin
             frmVOD.ngVOD.AddRow(1);
             frmVOD.ngVOD.BeginUpdate;
             frmVOD.ngVOD.Cell[0, i-1].AsString := inttostr(i);
             frmVOD.ngVOD.Cell[1, i-1].AsString := dm.DDL.FieldValues['VODEPGTITLE'];
             //frmVOD.ngVOD.Cell[2, i-1].AsString := dm.DDL.FieldValues['VODCAPRODUCTID'];
             if VarIsNull(dm.DDL.FieldValues['VODCAPRODUCTID'])
              then frmVOD.ngVOD.Cell[2, i-1].AsString := ''
               Else frmVOD.ngVOD.Cell[2, i-1].AsString := dm.DDL.FieldValues['VODCAPRODUCTID'];;
             frmVOD.ngVOD.Cell[3, i-1].AsString := dm.DDL.FieldValues['VODCAPSTARTDATE'];
             frmVOD.ngVOD.Cell[4, i-1].AsString := dm.DDL.FieldValues['VODCAPENDDATE'];
             //frmVOD.ngVOD.Cell[5, i-1].AsString := dm.DDL.FieldValues['VODCASERVICEID'];
             if VarIsNull(dm.DDL.FieldValues['VODCASERVICEID'])
              then frmVOD.ngVOD.Cell[5, i-1].AsString := ''
               Else frmVOD.ngVOD.Cell[5, i-1].AsString := dm.DDL.FieldValues['VODCASERVICEID'];;
             //frmVOD.ngVOD.Cell[6, i-1].AsString := dm.DDL.FieldValues['VODPROGRAMID'];
             if VarIsNull(dm.DDL.FieldValues['VODPROGRAMID'])
              then frmVOD.ngVOD.Cell[6, i-1].AsString := ''
               Else frmVOD.ngVOD.Cell[6, i-1].AsString := dm.DDL.FieldValues['VODPROGRAMID'];;
             frmVOD.ngVOD.Cell[7, i-1].AsString := dm.DDL.FieldValues['VODTRAFFICKEY'];
             //frmVOD.ngVOD.Cell[8, i-1].AsString := dm.DDL.FieldValues['VODGROUPKEY'];
             if VarIsNull(dm.DDL.FieldValues['VODGROUPKEY'])
              then frmVOD.ngVOD.Cell[8, i-1].AsString := ''
               Else frmVOD.ngVOD.Cell[8, i-1].AsString := dm.DDL.FieldValues['VODGROUPKEY'];;
             frmVOD.ngVOD.Cell[9, i-1].AsString := dm.DDL.FieldValues['VODPROGRAMKEY'];
             frmVOD.ngVOD.Cell[10, i-1].AsString := dm.DDL.FieldValues['VODFED'];
             frmVOD.ngVOD.Cell[11, i-1].AsString := dm.DDL.FieldValues['VODTIMEOFFSET'];;
             //frmVOD.ngVOD.Cell[12, i-1].AsString := dm.DDL.FieldValues['VODSTATUS'];;
             if VarIsNull(dm.DDL.FieldValues['VODSTATUS'])
              then frmVOD.ngVOD.Cell[12, i-1].AsString := '0'
               Else frmVOD.ngVOD.Cell[12, i-1].AsString := dm.DDL.FieldValues['VODSTATUS'];;
             if VarIsNull(dm.DDL.FieldValues['VODAMOUNT'])
              then frmVOD.ngVOD.Cell[13, i-1].AsString := '0'
               Else frmVOD.ngVOD.Cell[13, i-1].AsString := dm.DDL.FieldValues['VODAMOUNT'];;

             frmVOD.ngVOD.EndUpdate;
             dm.DDL.Next;
            end;

          i:=i+1;
          pbVOD.Position:=i-1;
        end;

     //AssignFile(SGI, 'C:\SGI\' + trim(frmSchEditor.cbChannelSch.Text)+'_'+ FormatDateTime('mmdd',frmSchEditor.dtpStart.Date)+ '-' + FormatDateTime('mmddyyy',frmSchEditor.dtpEnd.Date) +'.sgi');

     {if not VarIsNull(dm.DDL.FieldValues['VODPROGRAMKEY']) and (cbVOD.Checked = False) then
       Begin
        if (dm.DDL.FieldValues['VODGROUPKEY']= '12346') or (VarIsNull(dm.DDL.FieldValues['VODGROUPKEY'])) then
         Writeln(SGI,'8~',dm.DDL.FieldValues['VODCAPRODUCTID'],'~2497~1~B~3~',dm.DDL.FieldValues['VODCAPSTARTDATE'],'00000000~',dm.DDL.FieldValues['VODCAPENDDATE'],'00000000~',dm.DDL.FieldValues['VODCASERVICEID'],'~');
       End; }


     pbVOD.Visible := False;
     ShowMessage('Table has been exported!');
    end;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmVOD.BitBtn1Click(Sender: TObject);
begin
  frmVOD.Close;
end;

procedure TfrmVOD.btnJustCAClick(Sender: TObject);
begin

  ngJustCA.ClearRows;
  grpJustCA.Visible := True; 

  strSQL := 'SELECT VODCAProductId FROM M_VOD ORDER BY VODCAProductId Desc ';
  RecSet(strSQL);
  cboCAPStart.Items.Clear;
  cboCAPEnd.Items.Clear;
  while not dm.DDL.Eof do
    Begin
     if not VarIsNull(dm.DDL.FieldValues['VODCAProductId']) then
      begin
       cboCAPStart.Items.Add(dm.DDL.FieldValues['VODCAProductId']);
       cboCAPEnd.Items.Add(dm.DDL.FieldValues['VODCAProductId']);
      end;
     dm.DDL.Next;
    End;

end;

procedure TfrmVOD.edtEpgClick(Sender: TObject);
begin
  grpJustCA.Visible := False;
end;   

procedure TfrmVOD.cboCAPEndSelect(Sender: TObject);
var
  I, I2 : Integer;
begin
  ngJustCA.ClearRows;

  strSQL := 'SELECT VODEpgTitle, VODCAProductId, VODCAPStartDate, VODCAPEndDate, VODCAServiceID, VODUserCreateDate, VODProgramID, VODStatus, VODAmount ';
  strSQL := strSQL + 'FROM M_VOD ';
  strSQL := strSQL + 'WHERE VODCAProductId BETWEEN ''' + cboCAPStart.Text + ''' ';
  strSQL := strSQL + '      AND ''' + cboCAPEnd.Text + ''' ';
  //strSQL := strSQL + '      AND VODUserCreateDate >= sysdate ' ;
  strSQL := strSQL + '      AND to_date(to_char(VODUserCreateDate,''mm/dd/yyyy''),''mm/dd/yyyy'') >= to_date(to_char(sysdate,''mm/dd/yyyy''),''mm/dd/yyyy'') ';
  strSQL := strSQL + 'Order By VODCAProductId ';
  RecSet(strSQL);

  if dm.DDL.Eof then Exit;
  ngJustCA.AddRow(dm.DDL.RecordCount);
  ngJustCA.BeginUpdate;
  I := 0;
  While Not dm.DDL.Eof do
    Begin
      ngJustCA.Cell[0, I].AsString := IntToStr(I);
      ngJustCA.Cell[1, I].AsString := dm.DDL.FieldValues['VODEpgTitle'];
      ngJustCA.Cell[2, I].AsString := dm.DDL.FieldValues['VODCAProductId'];
      //ngJustCA.Cell[3, I].AsString := dm.DDL.FieldValues['VODCAPStartDate'];
      //ngJustCA.Cell[4, I].AsString := dm.DDL.FieldValues['VODCAPEndDate'];
      ngJustCA.Cell[3, I].AsString := FormatDateTime('ddMMyyyy', dm.DDL.FieldValues['VODCAPStartDate']);
      ngJustCA.Cell[4, I].AsString := FormatDateTime('ddMMyyyy', dm.DDL.FieldValues['VODCAPEndDate']);
      ngJustCA.Cell[5, I].AsString := dm.DDL.FieldValues['VODCAServiceID'];
      ngJustCA.Cell[6, I].AsString := FormatDateTime('MM/dd/yyyy', dm.DDL.FieldValues['VODCAPStartDate']);
      ngJustCA.Cell[7, I].AsString := FormatDateTime('MM/dd/yyyy', dm.DDL.FieldValues['VODCAPEndDate']);
      ngJustCA.Cell[8, I].AsString := FormatDateTime('MM/dd/yyyy', dm.DDL.FieldValues['VODUserCreateDate']);
      ngJustCA.Cell[9, I].AsString := dm.DDL.FieldValues['VODProgramID'];
      ngJustCA.Cell[10, I].AsString := dm.DDL.FieldValues['VODStatus'];
      ngJustCA.Cell[11, I].AsString := dm.DDL.FieldValues['VODAmount'];
      I := I + 1;
      dm.DDL.Next;
    End;
  ngJustCA.EndUpdate;    

  for I := 0 to ngJustCA.RowCount - 1 do
    Begin
      if odd(I) then
        Begin
          for I2 := 0 to 11 do
            Begin
              ngJustCA.Cell[I2, I].Color := clGrayText;
            End;
        End;
    End;

end;

procedure TfrmVOD.FormActivate(Sender: TObject);
begin
  SetCueBanner(edtEPG, 'Input search keyword Here, then press enter');
  strSQL := 'SELECT VODCAProductId FROM M_VOD ORDER BY VODCAProductId Desc ';
  RecSet(strSQL);
  cboCAPStart.Items.Clear;
  cboCAPEnd.Items.Clear;
  while not dm.DDL.Eof do
    Begin
     if not VarIsNull(dm.DDL.FieldValues['VODCAProductId']) then
      begin
       cboCAPStart.Items.Add(dm.DDL.FieldValues['VODCAProductId']);
       cboCAPEnd.Items.Add(dm.DDL.FieldValues['VODCAProductId']);
      end;
     dm.DDL.Next;
    End;

end;

procedure TfrmVOD.BitBtn2Click(Sender: TObject);
var
 JustCANDS : TextFile;
 JustCASMS : TextFile;
 i : Integer;
begin

  if not DirectoryExists('C:\SGI\Upload NDS') Then CreateDir('C:\SGI\Upload NDS');
  if not DirectoryExists('C:\SGI\Upload SMS') Then CreateDir('C:\SGI\Upload SMS');

  AssignFile(JustCANDS, 'C:\SGI\Upload NDS\NDS_CAProduct_From ' + cboCAPStart.Text + ' To ' + cboCAPEnd.Text +'.sgi');
  Rewrite(JustCANDS);
  AssignFile(JustCASMS, 'C:\SGI\Upload SMS\SMS_CAProduct_From ' + cboCAPStart.Text + ' To ' + cboCAPEnd.Text +'.sgi');
  Rewrite(JustCASMS);

  for i:= 0 to ngJustCA.RowCount - 1 do
   Begin
   // 8~OPPV000047~2497~1~B~3~1502201300000000~3112201300000000~9392~
   // Writeln(SGI,'8~',dm.DDLPush.FieldValues['VODCAPRODUCTID'],'~2497~1~B~3~',dm.DDLPush.FieldValues['VODCAPSTARTDATE'],'00000000~',dm.DDLPush.FieldValues['VODCAPENDDATE'],'00000000~',dm.DDLPush.FieldValues['VODCASERVICEID'],'~');

    Writeln(JustCANDS, '8~' + frmVOD.ngJustCA.Cells[2, i] + '~2497~1~B~3~' + frmVOD.ngJustCA.Cells[3, i] + '00000000~' + frmVOD.ngJustCA.Cells[4, i] + '00000000~' + frmVOD.ngJustCA.Cells[5, i] + '~');
    Writeln(JustCASMS, '8~' + frmVOD.ngJustCA.Cells[9, i] + '~' + frmVOD.ngJustCA.Cells[2, i] + '~2497~1~B~3~' + frmVOD.ngJustCA.Cells[3, i] + '00000000~' + frmVOD.ngJustCA.Cells[4, i] + '00000000~' + frmVOD.ngJustCA.Cells[5, i] + '~' + frmVOD.ngJustCA.Cells[1, i] + '~' + frmVOD.ngJustCA.Cells[11, i] + '~' + frmVOD.ngJustCA.Cells[10, i] + '~');

   End;

  CloseFile(JustCANDS);
  CloseFile(JustCASMS);
  ShowMessage('Just CA Product SGI For NDS And SMS Created !');

end;

procedure TfrmVOD.cboFilterVODClick(Sender: TObject);
begin
  grpJustCA.Visible := False;
end;

procedure TfrmVOD.ngVODClick(Sender: TObject);
begin
  grpJustCA.Visible := False;
end;





end.
