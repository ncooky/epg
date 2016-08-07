unit frm_Vis1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, NxColumns, NxColumnClasses, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, ExtCtrls, jpeg;

type
  TfrmVis1 = class(TForm)
    ppmCAPkgV1: TPopupMenu;
    AddRow1: TMenuItem;
    Save1: TMenuItem;
    AddNew1: TMenuItem;
    Exit1: TMenuItem;
    Delete1: TMenuItem;
    Shape2: TShape;
    ScrollBox1: TScrollBox;
    Image1: TImage;
    ScrollBox2: TScrollBox;
    ngVis1: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    cbTemplate: TComboBox;
    Button1: TButton;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
    procedure AddRow1Click(Sender: TObject);
    procedure AddNew1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Save1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure cbTemplateSelect(Sender: TObject);
    procedure ngVis1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmVis1: TfrmVis1;

implementation

uses frm_dm, frm_Login;

{$R *.dfm}

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT TCAID FROM T_CATEMPLATE ORDER BY TCAID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   strAngka := dm.DDL.FieldValues['TCAID'] + 1;
  end;
 fncangka:=IntToStr(strAngka); 
end;

procedure prcShow(channel:string);
var
 i : integer;
begin
 strSQL := ' SELECT tca_code, tca_number from t_catemplate  ';
 if channel <> trim('All Templates') then
  begin
   strSQL := strSQL + 'WHERE tca_code = ''' + channel + ''' ';
  end;
 strSQL := strSQL + ' ORDER by tca_code';
 RecSet(strSQL);
 i:=1;
 while not dm.DDL.Eof do
  begin
     frmVis1.ngVis1.AddCells([inttostr(i),
                        dm.DDL.FieldValues['tca_code'],
                        dm.DDL.FieldValues['tca_number'],
                        dm.DDL.FieldValues['tca_code'],
                        dm.DDL.FieldValues['tca_number']
                        ]);
     i:=i+1;
     dm.DDL.Next;
  end;
end;

procedure TfrmVis1.FormShow(Sender: TObject);
var
 item : TStrings;
begin
 cbTemplate.Clear;
 strSQL := 'SELECT distinct tca_code FROM t_catemplate ORDER BY tca_code';
 RecSet(strSQL);

 item:=cbTemplate.Items.Create;
 item.Add('All Templates');
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['tca_code']);
  dm.DDL.Next;
 end;
 cbTemplate.ItemIndex:=0;
end;

procedure TfrmVis1.AddRow1Click(Sender: TObject);
begin
 ngVis1.AddRow(1);
end;

procedure TfrmVis1.AddNew1Click(Sender: TObject);
begin
 ngVis1.ClearRows;
 ngVis1.AddRow(1);
end;

procedure TfrmVis1.Exit1Click(Sender: TObject);
begin
 frmVis1.Close;
end;

procedure TfrmVis1.Save1Click(Sender: TObject);
var
 i : integer;
 epgReplace, synEng, synInd, angka : string;
begin
 if strUserACC = 'Admin' then
  begin
   for i := 0 to ngVis1.RowCount-1 do
   begin
    angka:=fncangka;
    strSQL := 'SELECT tca_code, tca_number FROM t_catemplate ';
    strSQL := strSQL + 'WHERE tca_code = ''' + ngVis1.Cells[3,i] + ''' ';
    strSQL := strSQL + 'AND tca_number = ''' + ngVis1.Cells[4,i] + ''' ';
    RecSet(strSQL);

    if dm.DDL.Eof then
    begin
     strSQL := 'INSERT INTO SGI.t_catemplate ( ';
     strSQL := strSQL + 'TCAID, tca_code, tca_number, TCA_USERCREATE, TCA_USERCREATEDATE, TCA_USERUPDATE, TCA_USERUPDATEDATE) ';
     strSQL := strSQL + 'VALUES ( ';
     strSQL := strSQL + '''' + angka + ''', ';
     if frmVis1.cbTemplate.Text='All Templates' then
      begin
       strSQL := strSQL + '''' + ngVis1.Cells[1,i] + ''', ';
      end
     else
      begin
       strSQL := strSQL + '''' + trim(cbTemplate.Text) + ''', ';
      end;
     strSQL := strSQL + '''' + ngVis1.Cells[2,i] +  ''', ';
     strSQL := strSQL + '''' + strUser +  ''', ';
     strSQL := strSQL + 'sysdate, ';
     strSQL := strSQL + '''' + strUser +  ''', ';
     strSQL := strSQL + 'sysdate ) ';
     RecExc(strSQL);
    end
   else
    begin
     strSQL := ' UPDATE t_catemplate SET tca_code = ''' + ngVis1.Cells[1,i] + ''', tca_number = ''' + ngVis1.Cells[2,i] + ''', TCA_USERUPDATE = ''' + strUser + ''', TCA_USERUPDATEDATE = sysdate ';
     strSQL := strSQL + ' WHERE tca_code = ''' + ngVis1.Cells[3,i] + ''' ';
     strSQL := strSQL + 'AND tca_number = ''' + ngVis1.Cells[4,i] + ''' ';
     RecExc(strSQL);
    end;

   end;
   ngVis1.ClearRows;
   prcShow(cbTemplate.Text);
   ShowMessage('Data Has Been Saved!');
  end
 else
  begin
   ShowMessage('You Are Not Authorized');
  end;
end;


procedure TfrmVis1.Button1Click(Sender: TObject);
begin
 frmVis1.Close;
end;

procedure TfrmVis1.cbTemplateSelect(Sender: TObject);
begin
 ngVis1.ClearRows;
 prcShow(cbTemplate.Text);
end;

procedure TfrmVis1.ngVis1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
   if Button = mbRight Then
    Begin
      ppmCAPkgV1.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

end.
