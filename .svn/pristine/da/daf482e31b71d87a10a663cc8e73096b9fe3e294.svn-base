unit frm_SynopsisManual;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid, Menus, NxColumnClasses, NxColumns, StdCtrls, ComCtrls, jpeg;

type
  TfrmSynopsisManual = class(TForm)
    ppmSysMan: TPopupMenu;
    AddRow1: TMenuItem;
    AddNew1: TMenuItem;
    Update1: TMenuItem;
    Exit1: TMenuItem;
    ScrollBox1: TScrollBox;
    Image1: TImage;
    ScrollBox2: TScrollBox;
    ngSynopsisManual: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    NxTextColumn11: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    button1: TButton;
    cbCategory: TComboBox;
    Search: TButton;
    Delete1: TMenuItem;
    pbSysManual: TProgressBar;
    Panel1: TPanel;
    Label1: TLabel;
    cboFilterSyn: TComboBox;
    edtEpg: TEdit;
    Label2: TLabel;
    Panel2: TPanel;
    procedure ngSynopsisManualMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure AddRow1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbEpgSelect(Sender: TObject);
    procedure dtpAwalChange(Sender: TObject);
//    procedure FormResize(Sender: TObject);
    procedure Update1Click(Sender: TObject);
    procedure AddNew1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure cbCategorySelect(Sender: TObject);
    procedure edtEpgChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SearchClick(Sender: TObject);
    procedure edtEpgKeyPress(Sender: TObject; var Key: Char);
    procedure ngSynopsisManualSelectCell(Sender: TObject; ACol,
      ARow: Integer);
    procedure Delete1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  X, Y : Integer;
  frmSynopsisManual: TfrmSynopsisManual;
  Procedure prcShow(epg:string; category:string);

implementation

uses frm_dm, frm_Login;

{$R *.dfm}

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

Procedure prcShow(epg:string; category:string);
var
 strSQL : String;
 synInd, synEng : String;
 i,ii : integer;
 xyz : byte;
begin
 frmSynopsisManual.ngSynopsisManual.ClearRows; 
 strSQL := 'select JUMLAH_SYPNO, SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, SYCATEGORY ';
 strSQL := strSQL + 'FROM (SELECT COUNT(0) AS JUMLAH_SYPNO FROM M_SYNOPSIS WHERE ';
 if frmSynopsisManual.cboFilterSyn.Text = 'EPG Tittle'
    Then strSQL := strSQL + '      SYEPG_TITLE LIKE ''%' + epg + '%'' ';
 if frmSynopsisManual.cboFilterSyn.Text = 'Synopsis Ind'
    Then strSQL := strSQL + '      SYSYNOPSIS_IND LIKE ''%' + epg + '%'' ';
 if frmSynopsisManual.cboFilterSyn.Text = 'Synopsis Eng'
    Then strSQL := strSQL + '      SYSYNOPSIS_ENG LIKE ''%' + epg + '%'' ';

 if frmSynopsisManual.cbCategory.Text <> trim('All CATEGORY') then
  begin
   strSQL := strSQL + ' AND SYCATEGORY = ''' + category + ''' )xxx,';
  end
 else
  begin
   strSQL := strSQL + ' )xxx, ';
  end;

 strSQL := strSQL + '(SELECT SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, SYCATEGORY FROM M_SYNOPSIS WHERE ';
 if frmSynopsisManual.cboFilterSyn.Text = 'EPG Tittle'
    Then strSQL := strSQL + '      SYEPG_TITLE LIKE ''%' + epg + '%'' ';
 if frmSynopsisManual.cboFilterSyn.Text = 'Synopsis Ind'
    Then strSQL := strSQL + '      SYSYNOPSIS_IND LIKE ''%' + epg + '%'' ';
 if frmSynopsisManual.cboFilterSyn.Text = 'Synopsis Eng'
    Then strSQL := strSQL + '      SYSYNOPSIS_ENG LIKE ''%' + epg + '%'' ';

 if frmSynopsisManual.cbCategory.Text <> trim('All CATEGORY') then
  begin
   strSQL := strSQL + ' AND SYCATEGORY = ''' + category + ''' ) yyy ';
  end
 else
  begin
   strSQL := strSQL + ' )yyy ';
  end;

 strSQL := strSQL + 'ORDER BY SYCATEGORY, SYEPG_TITLE ';
 RecSet2(strSQL);

 xyz := dm.DDL2.FieldValues['JUMLAH_SYPNO'];
 //frmSynopsisManual.Caption := inttostr(xyz);
 i:=1;
 if not VarIsNull(dm.DDL2.FieldValues['jumlah_SYPNO']) then
  begin
   ii:=dm.DDL2.FieldValues['jumlah_SYPNO'];
  end
 else
  begin
   ii:=0;
  end;
 frmSynopsisManual.ngSynopsisManual.AddRow(ii);
 frmSynopsisManual.pbSysManual.Min:=0;
 frmSynopsisManual.ngSynopsisManual.BeginUpdate;
 frmSynopsisManual.pbSysManual.Visible:=True;
 frmSynopsisManual.pbSysManual.Max:=ii;
 while not dm.DDL2.Eof do
  begin

   if not VarIsNull(dm.DDL2.FieldValues['SYSYNOPSIS_IND'])then
    begin
     synInd:=dm.DDL2.FieldValues['SYSYNOPSIS_IND'];
    end
   else synInd:='';

   if not VarIsNull(dm.DDL2.FieldValues['SYSYNOPSIS_ENG'])then
    begin
     synEng:=dm.DDL2.FieldValues['SYSYNOPSIS_ENG'];
    end
   else synEng:='';

   frmSynopsisManual.ngSynopsisManual.Cell[0, i-1].AsString := inttostr(i);
   frmSynopsisManual.ngSynopsisManual.Cell[1, i-1].AsString := dm.DDL2.FieldValues['SYEPG_TITLE'];
   frmSynopsisManual.ngSynopsisManual.Cell[2, i-1].AsString := dm.DDL2.FieldValues['SYCATEGORY'];
   frmSynopsisManual.ngSynopsisManual.Cell[3, i-1].AsString := synInd;
   frmSynopsisManual.ngSynopsisManual.Cell[4, i-1].AsString := synEng;
   frmSynopsisManual.ngSynopsisManual.Cell[5, i-1].AsString := dm.DDL2.FieldValues['SYEPG_TITLE'];
   frmSynopsisManual.ngSynopsisManual.Cell[6, i-1].AsString := dm.DDL2.FieldValues['SYCATEGORY'];
   frmSynopsisManual.ngSynopsisManual.Cell[7, i-1].AsString := synInd;
   frmSynopsisManual.ngSynopsisManual.Cell[8, i-1].AsString := synEng;

   {frmSynopsisManual.ngSynopsisManual.AddCells([inttostr(i),
                                              dm.DDL.FieldValues['SYEPG_TITLE'],
                                              dm.DDL.FieldValues['SYCATEGORY'],
                                              synInd,
                                              synEng,
                                              dm.DDL.FieldValues['SYEPG_TITLE'],
                                              dm.DDL.FieldValues['SYCATEGORY'],
                                              synInd,
                                              synEng
                                              ]);  }
   i:=i+1;
   frmSynopsisManual.pbSysManual.Position:=i;
   dm.DDL2.Next;
  end;
 frmSynopsisManual.ngSynopsisManual.EndUpdate;
 frmSynopsisManual.pbSysManual.Visible:=False;
end;

procedure TfrmSynopsisManual.ngSynopsisManualMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if Button = mbRight Then
    Begin
      ppmSysMan.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmSynopsisManual.AddRow1Click(Sender: TObject);
begin
 ngSynopsisManual.AddRow(1);
end;

procedure TfrmSynopsisManual.FormShow(Sender: TObject);
var
 item : TStrings;
begin
 {cbChannel.Clear;
 strSQL := 'SELECT SYEPG_TITLE FROM M_CHANNEL ORDER BY MCHANNEL';
 RecSet(strSQL);

 Item:=cbEpg.Items.Create;
 item.Add('All EPG');
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['SYEPG_TITLE']);
  dm.DDL.Next;
 end;
 cbChannel.ItemIndex:=0;}
 ngSynopsisManual.ClearRows;
 cbCategory.Clear;
 edtEpg.Clear;
 strSQL := 'SELECT DISTINCT SYCATEGORY FROM M_SYNOPSIS ORDER BY SYCATEGORY';
 RecSet(strSQL);

 Item:=cbCategory.Items.Create;
 item.Add('All CATEGORY');
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['SYCATEGORY']);
  dm.DDL.Next;
 end;
 cbCategory.ItemIndex:=0;
 //prcShow(edtEpg.Text,cbCategory.Text);
end;

procedure TfrmSynopsisManual.cbEpgSelect(Sender: TObject);
begin
 {Screen.Cursor:=crHourGlass;
 ngSynopsisManual.ClearRows;
 prcShow(cbChannel.Text,dtpAwal.Date);
 Screen.Cursor:=crDefault; }
end;

procedure TfrmSynopsisManual.dtpAwalChange(Sender: TObject);
begin
 {Screen.Cursor:=crHourGlass;
 ngSynopsisManual.ClearRows;
 prcShow(cbChannel.Text, dtpAwal.Date);
 Screen.Cursor:=crDefault;  }
end;

{procedure TfrmSynopsisManual.FormResize(Sender: TObject);
begin
 pbSysManual.Width:=frmSynopsisManual.Width-9;
 ngSynopsisManual.Height:=frmSynopsisManual.Height-202;
 ngSynopsisManual.Width :=frmSynopsisManual.Width-34;
// Shape1.Width := frmSynopsisManual.Width-10;
 Button1.Top := ngSynopsisManual.Top + ngSynopsisManual.Height + 5;
 Image2.Width:= frmSynopsisManual.Width-9;
 Image2.Height:=frmSynopsisManual.Height-34;
 Image1.Left:=Image2.Width-181;
 Image3.Top := ngSynopsisManual.Top + ngSynopsisManual.Height + 5;
 image3.Left:= Image2.Width-117;
end; }

procedure TfrmSynopsisManual.Update1Click(Sender: TObject);
var
 i : integer;
 epgReplace, synEng, synInd, angka : string;
begin
 for i := 0 to ngSynopsisManual.RowCount-1 do
 begin
  angka:=fncangka;
  strSQL := 'SELECT * FROM M_SYNOPSIS ';
  strSQL := strSQL + 'WHERE SYEPG_TITLE = ''' + ngSynopsisManual.Cells[5,i] + ''' ';
  strSQL := strSQL + 'AND SYCATEGORY = ''' + ngSynopsisManual.Cells[6,i] + ''' ';
  RecSet(strSQL);

   epgReplace:=Replace(trim(ngSynopsisManual.Cells[1,i]), '''','');
   epgReplace:=Replace(trim(epgReplace), '"','');

   if not VarIsNull(ngSynopsisManual.Cells[3,i])then
    begin
     synInd:=Replace(trim(ngSynopsisManual.Cells[3,i]), '''','`');
     synInd:=Replace(trim(synInd), '`','''''');
     synInd:=Replace(trim(synInd), '"','``');
     synInd:=Replace(trim(synInd),#13,'');
    end
   else synInd:='';

   if not VarIsNull(ngSynopsisManual.Cells[4,i])then
    begin
     synEng:=Replace(trim(ngSynopsisManual.Cells[4,i]), '''','`');
     synEng:=Replace(trim(synEng), '`','''''');
     synEng:=Replace(trim(synEng), '"','``');
    end
   else synEng:='';

  if dm.DDL.Eof then
  begin
   strSQL := 'INSERT INTO SGI.M_SYNOPSIS ( ';
   strSQL := strSQL + 'SYID, SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, ';
   strSQL := strSQL + 'SYCATEGORY, SYUSER_CREATE, SYUSER_CREATEDATE, ';
   strSQL := strSQL + 'SYUSER_UPDATE, SYUSER_UPDATEDATE) ';
   strSQL := strSQL + 'VALUES ( ';
   strSQL := strSQL + '''' + angka + ''', ';
   strSQL := strSQL + '''' + epgReplace + ''', ';
   strSQL := strSQL + '''' + synInd + ''', ';
   strSQL := strSQL + '''' + synEng +  ''', ';
   strSQL := strSQL + '''' + ngSynopsisManual.Cells[2,i] + ''', ';
   strSQL := strSQL + '''' + strUser +  ''', ';
   strSQL := strSQL + 'sysdate, ';
   strSQL := strSQL + '''' + strUser +  ''', ';
   strSQL := strSQL + 'sysdate ) ';
   RecExc(strSQL);
  end
  else
  begin
   strSQL := ' UPDATE M_SYNOPSIS SET SYEPG_TITLE = ''' + epgReplace + ''', SYCATEGORY = ''' + ngSynopsisManual.Cells[2,i] + ''', ';
   strSQL := strSQL + ' SYSYNOPSIS_IND = ''' + synInd + ''', ';
   strSQL := strSQL + ' SYSYNOPSIS_ENG = ''' + synEng + ''', ';
   strSQL := strSQL + ' SYUSER_UPDATE = ''' + strUser + ''', ';
   strSQL := strSQL + ' SYUSER_UPDATEDATE = sysdate ';
   strSQL := strSQL + ' WHERE SYEPG_TITLE = ''' + ngSynopsisManual.Cells[5,i] + ''' ';
   strSQL := strSQL + 'AND SYCATEGORY = ''' + ngSynopsisManual.Cells[6,i] + ''' ';
   RecExc(strSQL);
  end;

 end;
 ngSynopsisManual.ClearRows;
 prcShow(edtEpg.Text,cbCategory.TEXT);
 ShowMessage('Data Has Been Saved!');
end;

procedure TfrmSynopsisManual.AddNew1Click(Sender: TObject);
begin
 ngSynopsisManual.ClearRows;
 ngSynopsisManual.AddRow(1);
end;

procedure TfrmSynopsisManual.Exit1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmSynopsisManual.Button1Click(Sender: TObject);
begin
 frmSynopsisManual.Close;
end;

procedure TfrmSynopsisManual.cbCategorySelect(Sender: TObject);
begin
 Screen.Cursor:=crHourGlass;
 ngSynopsisManual.ClearRows;
 prcShow(edtEpg.Text, cbCategory.text);
 Screen.Cursor:=crDefault;
end;

procedure TfrmSynopsisManual.edtEpgChange(Sender: TObject);
begin
{ Screen.Cursor:=crHourGlass;
 ngSynopsisManual.ClearRows;
 if trim(edtEpg.text)<>'' then
  begin
   prcShow(edtEpg.Text,cbCategory.text);
  end;
 Screen.Cursor:=crDefault; }
end;

procedure TfrmSynopsisManual.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
 frmSynopsisManual.Close;
end;

procedure TfrmSynopsisManual.SearchClick(Sender: TObject);
begin
 Screen.Cursor:=crHourGlass;
 ngSynopsisManual.ClearRows;
 if trim(edtEpg.text)<>'' then
  begin
   prcShow(edtEpg.Text,cbCategory.text);
  end;
 Screen.Cursor:=crDefault;
end;

procedure TfrmSynopsisManual.edtEpgKeyPress(Sender: TObject;
  var Key: Char);
begin
if key=#13 then
  begin
    Screen.Cursor:=crHourGlass;
    ngSynopsisManual.ClearRows;
    if trim(edtEpg.text)<> '' then
    begin
      prcShow(edtEpg.Text,cbCategory.text);
    end;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmSynopsisManual.ngSynopsisManualSelectCell(Sender: TObject;
  ACol, ARow: Integer);
begin
  X := ACol;
  Y := ARow;
end;

procedure TfrmSynopsisManual.Delete1Click(Sender: TObject);
var
  strDelSys : String;
begin
   Screen.Cursor := crHourGlass;
   strDelSys := ngSynopsisManual.Cells[1,y];
   strSQL := 'DELETE FROM M_SYNOPSIS WHERE SYEPG_TITLE = '''+trim(ngSynopsisManual.Cells[1,y])+''' and SYCATEGORY = '''+trim(ngSynopsisManual.Cells[2,y])+''' ';
   RecExc(strSQL);
   prcShow(edtEpg.Text,cbCategory.text);
   ShowMessage('SYNOPSIS '+strDelSys+' Has Been Removed');
   Screen.Cursor := crDefault;
end;

end.
