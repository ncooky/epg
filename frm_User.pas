unit frm_User;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, jpeg, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, NxColumns, NxColumnClasses, Menus, DCPcrypt2,
  DCPmd4, DCPmd5;

type
  TfrmUser = class(TForm)
    edtName: TEdit;
    Label1: TLabel;
    edtUserName: TEdit;
    Label3: TLabel;
    edtPass: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    edtRePass: TEdit;
    Label6: TLabel;
    Image1: TImage;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Label7: TLabel;
    cbUserAcc: TComboBox;
    btnSave: TButton;
    Button2: TButton;
    ngUserInfo: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxComboBoxColumn1: TNxComboBoxColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    PopupMenu1: TPopupMenu;
    AddRow1: TMenuItem;
    Save1: TMenuItem;
    Panel1: TPanel;
    Image3: TImage;
    DeleteUser1: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure AddRow1Click(Sender: TObject);
    procedure Save1Click(Sender: TObject);
    procedure ngUserInfoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ngUserInfoSelectCell(Sender: TObject; ACol, ARow: Integer);
    procedure ngUserInfoKeyPress(Sender: TObject; var Key: Char);
    procedure DeleteUser1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmUser: TfrmUser;
  tempname, tempusername, temppass : String;
  x, y : integer;

implementation

uses frm_dm, frm_Login;

{$R *.dfm}

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT USRID FROM M_USER ORDER BY USRID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   strAngka := dm.DDL.FieldValues['USRID'] + 1;
  end;
 fncangka:=IntToStr(strAngka);
end;

procedure TfrmUser.FormShow(Sender: TObject);
begin
 strSQL := 'SELECT UUSR_NAME, UUSR_ACC FROM M_USER WHERE UUSR_NAME = ''' + strUser + '''';
 RecSet(strSQL);

 //if not VarIsNull(dm.DDL.FieldValues['UUSR_NAME']) then
  //begin
   if dm.DDL.FieldValues['UUSR_ACC']='Admin' then
    begin
     strSQL := 'SELECT * FROM M_USER ORDER BY UUSR_NAME ';
     RecSet2(strSQL);
     ngUserInfo.ClearRows;
     frmuser.DeleteUser1.Visible := true;
     frmuser.AddRow1.Visible := true;
     while not dm.DDL2.Eof do
      begin
       ngUserInfo.AddCells([dm.DDL2.FieldValues['USRID'],
                    dm.DDL2.FieldValues['UUSR_NAME'],
                    dm.DDL2.FieldValues['UUSR_DESCRIPTION'],
                    dm.DDL2.FieldValues['UUSR_PASSWORD'],
                    dm.DDL2.FieldValues['UUSR_ACC'],
                    dm.DDL2.FieldValues['UUSR_NAME']
                  ]);
       dm.DDL2.Next;
      end;
     ngUserInfo.Visible:=True;
    end
   else
    begin
     strSQL := 'SELECT * FROM M_USER WHERE UUSR_NAME = ''' + strUser + ''' ORDER BY UUSR_NAME ';
     RecSet2(strSQL);
     ngUserInfo.ClearRows;
     frmuser.DeleteUser1.Visible := false;
     frmuser.AddRow1.Visible := False;
     while not dm.DDL2.Eof do
      begin
       ngUserInfo.AddCells([dm.DDL2.FieldValues['USRID'],
                    dm.DDL2.FieldValues['UUSR_NAME'],
                    dm.DDL2.FieldValues['UUSR_DESCRIPTION'],
                    dm.DDL2.FieldValues['UUSR_PASSWORD'],
                    dm.DDL2.FieldValues['UUSR_ACC'],
                    dm.DDL2.FieldValues['UUSR_NAME']
                  ]);
       dm.DDL2.Next;
      end;
      NxTextColumn2.Enabled := False;
      NxTextColumn4.Enabled := False;
      NxComboBoxColumn1.Enabled := False;
    { cbUserAcc.Visible:=False;
     Label7.Visible:=False;
     ngUserInfo.Visible:=False;
     edtName.Text:=dm.DDL2.FieldValues['UUSR_DESCRIPTION'];
     edtUserName.Text:=dm.DDL2.FieldValues['UUSR_NAME'];
     tempname:=dm.DDL2.FieldValues['UUSR_DESCRIPTION'];
     tempusername:=dm.DDL2.FieldValues['UUSR_NAME'];
     temppass:=dm.DDL2.FieldValues['UUSR_PASSWORD'];    }
    end;
  //end;

end;

procedure TfrmUser.btnSaveClick(Sender: TObject);
var
// i : integer;
 angka, tmpusr, txterror, txtsuccess: String;
 label
  errortext, successtext;
begin
 Screen.Cursor:=crHourGlass;
 {strSQL:= 'SELECT UUSR_ACC FROM M_USER WHERE UUSR_NAME = ''' + strUser + ''' ';
 RecSet(strSQL);
 if dm.DDL.FieldValues['UUSR_ACC'] <> 'Admin' then
  begin
   if trim(edtPass.Text) = '' then
    begin
     strSQL := ' UPDATE M_USER SET UUSR_NAME = ''' + edtUserName.Text + ''', UUSR_DESCRIPTION = ''' + edtName.Text + ''' ';
     strSQL := strSQL + ' WHERE UUSR_NAME = ''' + tempusername + ''' ';
     RecExc(strSQL);
     ShowMessage('Data Has Been Saved!');
     FormShow(Sender);
    end
   else
    begin
     if Trim(edtPass.Text) = trim(edtRePass.Text) then
      begin
       strSQL := ' UPDATE M_USER SET UUSR_NAME = ''' + edtUserName.Text + ''', UUSR_DESCRIPTION = ''' + edtName.Text + ''', UUSR_PASSWORD = ''' + fncMD5(edtPass.Text) + ''' ';
       strSQL := strSQL + ' WHERE UUSR_NAME = ''' + tempusername + ''' ';
       RecExc(strSQL);
       ShowMessage('Data Has Been Saved!');
       FormShow(Sender);
      end
     else
      begin
       ShowMessage('Please verify your password again !!!');
       edtRePass.SetFocus;
      end;
    end;
  end
 else
  begin    }
   {for i := 0 to ngUserInfo.RowCount-1 do
    begin
     strSQL:= 'SELECT UUSR_NAME FROM M_USER WHERE UUSR_NAME = ''' + ngUserInfo.Cells[5,i] + ''' ';
     RecSet2(strSQL);
     angka:=fncangka;
     if dm.DDL2.Eof then
      begin
       strSQL := 'INSERT INTO SGI.M_USER ( ';
       strSQL := strSQL + 'USRID, UUSR_NAME, UUSR_DESCRIPTION, UUSR_PASSWORD, UUSR_ACC) ';
       strSQL := strSQL + 'VALUES ( ';
       strSQL := strSQL + '''' + angka + ''', ';
       strSQL := strSQL + '''' + ngUserInfo.Cells[1,i] + ''', ';
       strSQL := strSQL + '''' + ngUserInfo.Cells[2,i] + ''', ';
       strSQL := strSQL + '''' + fncMD5(ngUserInfo.Cells[3,i]) + ''', ';
       strSQL := strSQL + '''' + ngUserInfo.Cells[4,i] +  ''') ';
       RecExc2(strSQL);
      end
     else
      begin
       strSQL := ' UPDATE M_USER SET UUSR_NAME = ''' + ngUserInfo.Cells[1,i] + ''', UUSR_DESCRIPTION = ''' + ngUserInfo.Cells[2,i] + ''', UUSR_PASSWORD = ''' + fncMD5(ngUserInfo.Cells[3,i]) + ''', ';
       strSQL := strSQL + ' UUSR_ACC = ''' + ngUserInfo.Cells[4,i] + ''' WHERE UUSR_NAME = ''' + ngUserInfo.Cells[5,i] + ''' ';
       RecExc2(strSQL);
      end;
    end;    }
   tmpusr := ngUserInfo.Cells[1,y];

   if ngUserInfo.Cells[1,y] = '' then
    begin
      txterror := 'Please Input The UserName';
      ngUserInfo.SelectCell(1,y) ;
      Goto errortext;
    end
   else if  ngUserInfo.Cells[2,y] = '' then
    begin
      txterror := 'Please Input User Description';
      ngUserInfo.SelectCell(2,y) ;
      Goto errortext;
    end
   else if  ngUserInfo.Cells[3,y] = '' then
    begin
      txterror := 'Please Input The Password';
      ngUserInfo.SelectCell(3,y) ;
      Goto errortext;
    end
   else if  ngUserInfo.Cells[4,y] = '' then
    begin
      txterror := 'Please Choose The User Access Rights' ;
      ngUserInfo.SelectCell(4,y) ;
      Goto errortext;
    end;
   strSQL:= 'SELECT UUSR_NAME FROM M_USER WHERE UUSR_NAME = ''' + tmpusr + ''' ';
   RecSet2(strSQL);
   angka:=fncangka;
   if dm.DDL2.Eof then
    begin
      strSQL := 'INSERT INTO SGI.M_USER ( ';
      strSQL := strSQL + 'USRID, UUSR_NAME, UUSR_DESCRIPTION, UUSR_PASSWORD, UUSR_ACC) ';
      strSQL := strSQL + 'VALUES ( ';
      strSQL := strSQL + '''' + angka + ''', ';
      strSQL := strSQL + '''' + ngUserInfo.Cells[1,y] + ''', ';
      strSQL := strSQL + '''' + ngUserInfo.Cells[2,y] + ''', ';
      strSQL := strSQL + '''' + fncMD5(ngUserInfo.Cells[3,y]) + ''', ';
      strSQL := strSQL + '''' + ngUserInfo.Cells[4,y] +  ''') ';
      RecExc2(strSQL);
      txtsuccess:= 'User has been created!';
      Goto successtext;
    end
   else
    begin
     strSQL := ' UPDATE M_USER SET UUSR_NAME = ''' + ngUserInfo.Cells[1,y] + ''', UUSR_DESCRIPTION = ''' + ngUserInfo.Cells[2,y] + ''', UUSR_PASSWORD = ''' + fncMD5(ngUserInfo.Cells[3,y]) + ''', ';
     strSQL := strSQL + ' UUSR_ACC = ''' + ngUserInfo.Cells[4,y] + ''' WHERE UUSR_NAME = ''' + tmpusr + ''' ';
     RecExc2(strSQL);
     txtsuccess:= 'User has been updated!';
     Goto successtext;
    end;



errortext:
  ShowMessage(txterror);
  exit;

successtext:
  ShowMessage(txtsuccess);

   FormShow(Sender);
  //end;
 Screen.Cursor:=crDefault;
end;

procedure TfrmUser.Button2Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmUser.AddRow1Click(Sender: TObject);
begin
 ngUserInfo.AddRow(1);
end;

procedure TfrmUser.Save1Click(Sender: TObject);
begin
 btnSaveClick(Sender);
end;

procedure TfrmUser.ngUserInfoMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
    if Button = mbRight Then
    Begin
      PopupMenu1.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmUser.ngUserInfoSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
  x := ACol;
  y := ARow;
end;

procedure TfrmUser.ngUserInfoKeyPress(Sender: TObject; var Key: Char);
begin
if key=#13 then
  begin
    Screen.Cursor:=crHourGlass;
    frmUser.btnSaveClick(Sender);
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmUser.DeleteUser1Click(Sender: TObject);
Var
  strDelete : String;
//  choose: Integer;
begin
    strSQL := 'Select UUSR_ACC from M_USER Where UUSR_NAME = ''' + ngUserInfo.Cells[1,y] + '''  ';
    RecSet(strSQL);
    if not (dm.ddl.FieldValues['UUSR_ACC']='Admin') then
      begin
       strDelete := ngUserInfo.Cells[1,y];
       strSQL := 'DELETE FROM M_USER WHERE UUSR_NAME = ''' + ngUserInfo.Cells[1,y] + '''  ';
       RecExc(strSQL);
   //prcShow(cbChannel.Text);
       ShowMessage('User '+strDelete+' Has Been Removed');
       FormShow(Sender);
      end
    else ShowMessage('User '+strDelete+' is Admin');
end;

end.
