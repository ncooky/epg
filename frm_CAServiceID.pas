unit frm_CAServiceID;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NxColumns, NxColumnClasses, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, jpeg, ExtCtrls, Menus,
  StdCtrls;

type
  TfrmCAServiceID = class(TForm)
    PopupMenu1: TPopupMenu;
    AddNewRow1: TMenuItem;
    Update1: TMenuItem;
    Exit1: TMenuItem;
    ScrollBox1: TScrollBox;
    Image1: TImage;
    ScrollBox2: TScrollBox;
    ngServiceID: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    Button1: TButton;
    ScrollBox3: TScrollBox;
    Label1: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    cboSearch: TComboBox;
    lblSearch: TLabel;
    txtSearch: TEdit;
    Button3: TButton;
    procedure FormShow(Sender: TObject);
    procedure ngServiceIDMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure AddNewRow1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Update1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure txtSearchKeyPress(Sender: TObject; var Key: Char);
    procedure Button3Click(Sender: TObject);
//    procedure txtSearchKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCAServiceID: TfrmCAServiceID;
  Procedure prcShow(epg:string);

implementation

uses frm_dm, frm_Login, frm_VOD;

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

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT CID FROM M_CASERVICEID ORDER BY CID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   strAngka := dm.DDL.FieldValues['CID'] + 1;
  end;
 fncangka:=IntToStr(strAngka); 
end;

procedure TfrmCAServiceID.FormShow(Sender: TObject);
var
 i:integer;
begin
 SetCueBanner(txtSearch, 'Input CA ID or CA Description Here, then press enter');
 Label1.Caption := GetAppVersion;
 frmCAServiceID.Caption := 'Form CA Service ID '+ GetAppVersion;
 i:=1;
 strSQL := 'SELECT * FROM M_CASERVICEID ORDER BY CCADESCRIPTION ';
 RecSet(strSQL);
 ngServiceID.ClearRows;
 while not dm.DDL.Eof do
  begin
   frmCAServiceID.ngServiceID.AddCells([inttostr(i),
                                        dm.DDL.FieldValues['CCAID'],
                                        dm.DDL.FieldValues['CCADESCRIPTION'],
                                        dm.DDL.FieldValues['CCAID'],
                                        dm.DDL.FieldValues['CCADESCRIPTION']
                                        ]);
   i:=i+1;
   dm.ddl.next;
  end;
end;

procedure TfrmCAServiceID.ngServiceIDMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if Button = mbRight Then
    Begin
      PopupMenu1.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmCAServiceID.AddNewRow1Click(Sender: TObject);
begin
 ngServiceID.AddRow(1);
end;

procedure TfrmCAServiceID.Exit1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmCAServiceID.Update1Click(Sender: TObject);
var
 i:integer;
 angka : String;
begin
if strUserACC = 'Admin' then
begin
 Screen.Cursor:=crHourGlass;
 for i := 0 to ngServiceID.RowCount-1 do
  begin
   angka:=fncangka;
   strSQL := 'SELECT * FROM M_CASERVICEID ';
   strSQL := strSQL + 'WHERE CCAID = ''' + ngServiceID.Cells[3,i] + ''' ';
   RecSet(strSQL);

   if dm.DDL.Eof then
    begin
     strSQL := 'INSERT INTO SGI.M_CASERVICEID ( ';
     strSQL := strSQL + 'CID, CCAID, CCADESCRIPTION, CAUSERCREATE, ';
     strSQL := strSQL + 'CAUSERCREATEDATE, CAUSERUPDATE, CAUSERUPDATEDATE) ';
     strSQL := strSQL + 'VALUES ( ';
     strSQL := strSQL + '''' + angka + ''', ';
     strSQL := strSQL + '''' + ngServiceID.Cells[1,i] + ''', ';
     strSQL := strSQL + '''' + ngServiceID.Cells[2,i] + ''', ';
     strSQL := strSQL + '''' + strUser +  ''', ';
     strSQL := strSQL + 'sysdate, ';
     strSQL := strSQL + '''' + strUser +  ''', ';
     strSQL := strSQL + 'sysdate) ';
     RecExc(strSQL);
    end
   else
    begin
     strSQL := ' UPDATE M_CASERVICEID SET CCAID = ''' + ngServiceID.Cells[1,i] + ''', CCADESCRIPTION = ''' + ngServiceID.Cells[2,i] + ''', ';
     strSQL := strSQL + ' CAUSERUPDATE = ''' + strUser + ''', ';
     strSQL := strSQL + ' CAUSERUPDATEDATE = sysdate ';
     strSQL := strSQL + ' WHERE CCAID = ''' + ngServiceID.Cells[3,i] + ''' ';
     RecExc(strSQL);
    end;
  end;
 ngServiceID.ClearRows;
 FormShow(Sender);
 Screen.Cursor:=crDefault;
 ShowMessage('Data Has Been Saved!');
 end
else
 begin
  ShowMessage('You Are Not Authorized');
 end;
end;

procedure TfrmCAServiceID.Button1Click(Sender: TObject);
begin
 frmCAServiceID.Close;
end;


Procedure prcShow (epg:string);
var
strSQL : String;
i: integer;

begin
 i:=1;
 strSQL := 'SELECT * FROM M_CASERVICEID WHERE ';
 if frmCAServiceID.cboSearch.Text = 'CA ID'
    Then strSQL := strSQL + '      CCAID LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 if frmCAServiceID.cboSearch.Text = 'CA Description'
    Then strSQL := strSQL + '      CCADESCRIPTION LIKE ''%' + AnsiUpperCase(epg) + '%'' ';
 strSQL := strSQL + 'ORDER BY CCADESCRIPTION ';
 RecSet(strSQL);
 frmCAServiceID.ngServiceID.ClearRows;
 while not dm.DDL.Eof do
  begin
   frmCAServiceID.ngServiceID.AddCells([inttostr(i),
                                        dm.DDL.FieldValues['CCAID'],
                                        dm.DDL.FieldValues['CCADESCRIPTION'],
                                        dm.DDL.FieldValues['CCAID'],
                                        dm.DDL.FieldValues['CCADESCRIPTION']
                                        ]);
   i:=i+1;
   dm.ddl.next;
  end;

  end;

procedure TfrmCAServiceID.txtSearchKeyPress(Sender: TObject;
  var Key: Char);
begin
if key=#13 then
  begin
    Screen.Cursor:=crHourGlass;
    frmCAServiceID.ngServiceID.ClearRows;
    if trim(txtSearch.text)<>'' then
    begin
      prcShow(txtSearch.Text);
    end
    else FormShow(Sender);
    Screen.Cursor:=crDefault;
  end;
end;



procedure TfrmCAServiceID.Button3Click(Sender: TObject);
begin
    Screen.Cursor:=crHourGlass;
    frmCAServiceID.ngServiceID.ClearRows;
    if trim(txtSearch.text)<>'' then
    begin
      prcShow(txtSearch.Text);
    end
    else FormShow(Sender);
    Screen.Cursor:=crDefault;
end;

end.
