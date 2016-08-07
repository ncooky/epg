unit frm_Synopsis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, jpeg, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, NxColumns, NxColumnClasses, StdCtrls;

type
  TfrmSynopsis = class(TForm)
    ngSipnosis: TNextGrid;
    Image1: TImage;
    Shape1: TShape;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    NxTextColumn7: TNxTextColumn;
    Shape2: TShape;
    Button1: TButton;
    Button2: TButton;
    NxTextColumn3: TNxTextColumn;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSynopsis: TfrmSynopsis;

implementation

uses frm_dm, frm_Read, frm_Login;

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

procedure TfrmSynopsis.Button1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmSynopsis.Button2Click(Sender: TObject);
var
 ii : integer;
 angka : String;
begin
 Screen.Cursor:=crHourGlass;
 //strSQL := 'DELETE FROM SGI.M_Synopsis';
 //RecExc(strSQL);

 for ii := 0 to ngSipnosis.RowCount-1 do
 begin
  angka:=fncangka;
  strSQL := ' select * FROM M_SYNOPSIS';
  strSQL := strSQL + ' WHERE SYEPG_TITLE = ''' + ngSipnosis.Cells[1,ii] + ''' ';
  strSQL := strSQL + ' AND SYCATEGORY = ''' + copy(ngSipnosis.Cells[2,ii],1,1) + ''' ';
  RecSet(StrSQL);
  if dm.DDL.Eof then
   begin
     strSQL := 'INSERT INTO SGI.M_SYNOPSIS ( ';
     strSQL := strSQL + 'SYID, SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, ';
     strSQL := strSQL + 'SYCATEGORY, SYUSER_CREATE, SYUSER_CREATEDATE, ';
     strSQL := strSQL + 'SYUSER_UPDATE, SYUSER_UPDATEDATE) ';
     strSQL := strSQL + 'VALUES ( ';
     strSQL := strSQL + '''' + angka + ''', ';
     strSQL := strSQL + '''' + ngSipnosis.Cells[1,ii] + ''', ';
     strSQL := strSQL + '''' + ngSipnosis.Cells[3,ii] + ''', ';
     strSQL := strSQL + '''' + ngSipnosis.Cells[4,ii] + ''', ';
     strSQL := strSQL + '''' + copy(ngSipnosis.Cells[2,ii],1,1) + ''', ';
     strSQL := strSQL + '''' + strUser + ''', ';
     strSQL := strSQL + 'sysdate, ';
     strSQL := strSQL + '''' + strUser + ''', ';
     strSQL := strSQL + 'sysdate) ';
     RecExc(strSQL);
  end;
 end;
 Screen.Cursor:=crDefault;
 ShowMessage('Data has been exported!');
end;
end.
