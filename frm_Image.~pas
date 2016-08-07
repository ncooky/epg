unit frm_Image;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, jpeg, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, NxColumns, NxColumnClasses, StdCtrls, Grids,
  ComCtrls, Menus;

type
  TfrmImage = class(TForm)
    ngImage: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    Image1: TImage;
    StringGrid1: TStringGrid;
    pbImage: TProgressBar;
    Button2: TButton;
    edtImage: TEdit;
    ppmImage: TPopupMenu;
    Delete1: TMenuItem;
    Label1: TLabel;
    AddRow1: TMenuItem;
    AddRow2: TMenuItem;
    Update1: TMenuItem;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    NxTextColumn7: TNxTextColumn;
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure edtImageKeyPress(Sender: TObject; var Key: Char);
    procedure ngImageMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Delete1Click(Sender: TObject);
    procedure AddRow1Click(Sender: TObject);
    procedure AddRow2Click(Sender: TObject);
    procedure Update1Click(Sender: TObject);
    procedure ngImageKeyPress(Sender: TObject; var Key: Char);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmImage: TfrmImage;
  strChannel, strTitle, strImageid : String;
  x, y : Integer;

implementation

uses frm_dm , ComObj, frm_Read, frm_Login;

{$R *.dfm}

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT max(IID) as iid FROM M_IMAGE ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   if VarIsNull(dm.DDL.FieldValues['IID'])
    Then strAngka := 1
      Else strAngka := dm.DDL.FieldValues['IID'] + 1;
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

function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
begin
  Result := False;
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.Workbooks.Open(AXLSFile);
    Sheet := XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[1];
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
    x := XLApp.ActiveCell.Row;
    y := XLApp.ActiveCell.Column;
    AGrid.RowCount := x;
    AGrid.ColCount := y;
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].value;
    k := 1;
    repeat
      for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := RangeMatrix[K, R];
      Inc(k, 1);
      AGrid.RowCount := k + 1;
    until k > x;
    RangeMatrix := Unassigned;

  finally
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

procedure TfrmImage.FormShow(Sender: TObject);
Var
 i, ii, ctrA, ctrB : integer;
 angka, epgReplace, synInd, synEng : String;
begin
ngImage.ClearRows;
if strIMGChoice='Upload' then
  begin
   i := 1;
   Screen.Cursor:=crHourGlass;
    if Xls_To_StringGrid(StringGrid1, frmRead.OpenDialog1.FileName) then
      begin
       pbImage.Max := StringGrid1.RowCount - 2;
       pbImage.Min := 0;
        while StringGrid1.Cells[0, i]<>'<eof>' do
          begin
            angka:=fncangka;
            ctrA:=StrToInt(angka);
            epgReplace:=Replace(trim(StringGrid1.Cells[1,i]), '''','');
            epgReplace:=Replace(trim(epgReplace), '"','');

            strSQL := 'SELECT * FROM M_IMAGE WHERE IEPG_ORI = upper(''' + epgReplace + ''') ';
            strSQL := strSQL + 'AND ICHANNEL = ''' + TRIM(StringGrid1.Cells[0,i]) + ''' ';
            strSQL := strSQL + 'AND IIMAGEID = ''' + TRIM(StringGrid1.Cells[2,i]) + ''' ';
            RecSet(strSQL);
            if not dm.DDL.Eof then
              begin
                strSQL := 'delete FROM M_IMAGE WHERE IEPG_ORI = upper(''' + epgReplace + ''') ';
                strSQL := strSQL + 'AND ICHANNEL = ''' + TRIM(StringGrid1.Cells[0,i])+ ''' ';
                strSQL := strSQL + 'AND IIMAGEID = ''' + TRIM(StringGrid1.Cells[2,i]) + ''' ';
                RecExc(strSQL);

                strSQL := 'INSERT INTO SGI.M_IMAGE ( ';
                strSQL := strSQL + 'IID, ICHANNEL, IEPG_ORI, IIMAGEID, ';
                strSQL := strSQL + 'IUSER_CREATE, ICREATE_DATE, IUSER_UPDATE, ';
                strSQL := strSQL + 'IUPDATE_DATE ) ';
                strSQL := strSQL + 'VALUES ( ';
                strSQL := strSQL + '''' + angka + ''', ';//IID
                strSQL := strSQL + '''' + TRIM(StringGrid1.Cells[0,i]) + ''', ';//ICHANNEL
                strSQL := strSQL + 'Upper(''' + epgReplace + '''), ';//IEPG_ORI
                strSQL := strSQL + '''' + TRIM(StringGrid1.Cells[2,i]) + ''', ';//IIMAGEID
                strSQL := strSQL + '''' + strUser + ''', ';//IUSER_CREATE
                strSQL := strSQL + 'sysdate, ';//ICREATE_DATE
                strSQL := strSQL + '''' + strUser + ''', ';//USER_UPDATE
                strSQL := strSQL + 'sysdate) ';//IUPDATE_DATE
                RecExc(strSQL);
                ctrB:=StrToInt(angka);
              end;

             if dm.DDL.Eof then
              begin
                strSQL := 'INSERT INTO SGI.M_IMAGE ( ';
                strSQL := strSQL + 'IID, ICHANNEL, IEPG_ORI, IIMAGEID, ';
                strSQL := strSQL + 'IUSER_CREATE, ICREATE_DATE, IUSER_UPDATE, ';
                strSQL := strSQL + 'IUPDATE_DATE ) ';
                strSQL := strSQL + 'VALUES ( ';
                strSQL := strSQL + '''' + angka + ''', ';//IID
                strSQL := strSQL + '''' + TRIM(StringGrid1.Cells[0,i]) + ''', ';//ICHANNEL
                strSQL := strSQL + 'Upper(''' + epgReplace + '''), ';//IEPG_ORI
                strSQL := strSQL + '''' + TRIM(StringGrid1.Cells[2,i]) + ''', ';//IIMAGEID
                strSQL := strSQL + '''' + strUser + ''', ';//IUSER_CREATE
                strSQL := strSQL + 'sysdate, ';//ICREATE_DATE
                strSQL := strSQL + '''' + strUser + ''', ';//USER_UPDATE
                strSQL := strSQL + 'sysdate) ';//IUPDATE_DATE
                RecExc(strSQL);
                ctrB:=StrToInt(angka);
              end;
              
            strSQL := 'SELECT * FROM M_IMAGE WHERE IEPG_ORI = upper(''' + epgReplace + ''') ';
            strSQL := strSQL + 'AND ICHANNEL = ''' + TRIM(StringGrid1.Cells[0,i]) + ''' ';
            strSQL := strSQL + 'AND IIMAGEID = ''' + TRIM(StringGrid1.Cells[2,i]) + ''' ';
            strSQL := strSQL + 'AND IID >= ''' + IntToStr(ctrA)+ ''' ';
            strSQL := strSQL + 'AND IID <= ''' + IntToStr(ctrB) + ''' ';
            RecSet(strSQL);
            while not dm.DDL.Eof do
              begin
               frmImage.ngImage.AddRow(1);
               frmImage.ngImage.BeginUpdate;
               frmImage.ngImage.Cell[0, i-1].AsString := inttostr(i);
               frmImage.ngImage.Cell[1, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
               frmImage.ngImage.Cell[2, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
               frmImage.ngImage.Cell[3, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
               frmImage.ngImage.Cell[4, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
               frmImage.ngImage.Cell[5, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
               frmImage.ngImage.Cell[6, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
               frmImage.ngImage.EndUpdate;
               dm.DDL.Next;
              end;
            i:=i+1;
            pbImage.Position:=i-1;
          end;
       Screen.Cursor:=crDefault;
       ShowMessage('Table has been exported!');
      end;
  end
  else
  if strIMGChoice = 'View' then
  begin
    Screen.Cursor:=crHourGlass;
    i:=1;
    strSQL:='SELECT * FROM M_IMAGE ORDER BY ICHANNEL, IEPG_ORI';
    RecSet(strSQL);
    while not dm.DDL.Eof do
    begin
      frmImage.ngImage.AddRow(1);
      frmImage.ngImage.BeginUpdate;
      frmImage.ngImage.Cell[0, i-1].AsString := inttostr(i);
      frmImage.ngImage.Cell[1, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
      frmImage.ngImage.Cell[2, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
      frmImage.ngImage.Cell[3, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
      frmImage.ngImage.Cell[4, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
      frmImage.ngImage.Cell[5, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
      frmImage.ngImage.Cell[6, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
      frmImage.ngImage.EndUpdate;
      dm.DDL.Next;
      i:=i+1;
    end;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmImage.Button2Click(Sender: TObject);
begin
 frmImage.Close;
end;

procedure TfrmImage.edtImageKeyPress(Sender: TObject; var Key: Char);
var
i : integer;
begin
  if key=#13 then
  begin
    Screen.Cursor:=crHourGlass;
    ngImage.ClearRows;
    if trim(edtImage.text)<>'' then
    begin
      i:=1;
      strSQL:='SELECT * FROM M_IMAGE WHERE IEPG_ORI LIKE '''+trim(edtImage.Text)+'%'' ORDER BY ICHANNEL, IEPG_ORI';
      RecSet(strSQL);
      while not dm.DDL.Eof do
      begin
        frmImage.ngImage.AddRow(1);
        frmImage.ngImage.BeginUpdate;
        frmImage.ngImage.Cell[0, i-1].AsString := inttostr(i);
        frmImage.ngImage.Cell[1, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
        frmImage.ngImage.Cell[2, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
        frmImage.ngImage.Cell[3, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
        frmImage.ngImage.Cell[4, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
        frmImage.ngImage.Cell[5, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
        frmImage.ngImage.Cell[6, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
        frmImage.ngImage.EndUpdate;
        dm.DDL.Next;
        i:=i+1;
      end;
    Screen.Cursor:=crDefault;
    end;
  end
end;

procedure TfrmImage.ngImageMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbRight Then
    Begin
      ppmImage.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmImage.Delete1Click(Sender: TObject);
var
  i : integer;
begin
  Screen.Cursor:=crHourGlass;
  strSQL := 'DELETE FROM M_IMAGE WHERE ICHANNEL = ''' + ngImage.Cells[1,y] + ''' and IEPG_ORI = ''' + ngImage.Cells[2,y] + ''' and IIMAGEID = ''' + ngImage.Cells[3,y] + '''   ';
  RecExc(strSQL);
  i:=1;
  strSQL:='SELECT * FROM M_IMAGE WHERE IEPG_ORI LIKE '''+trim(edtImage.Text)+'%'' ORDER BY ICHANNEL, IEPG_ORI';
  RecSet(strSQL);
  while not dm.DDL.Eof do
  begin
    frmImage.ngImage.AddRow(1);
    frmImage.ngImage.BeginUpdate;
    frmImage.ngImage.Cell[0, i-1].AsString := inttostr(i);
    frmImage.ngImage.Cell[1, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
    frmImage.ngImage.Cell[2, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
    frmImage.ngImage.Cell[3, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
    frmImage.ngImage.Cell[4, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
    frmImage.ngImage.Cell[5, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
    frmImage.ngImage.Cell[6, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
    frmImage.ngImage.EndUpdate;
    dm.DDL.Next;
    i:=i+1;
  end;
  Screen.Cursor:=crDefault;
  ShowMessage('IMAGE '+ ngImage.Cells[3,y] +' Has Been Removed');
end;

procedure TfrmImage.AddRow1Click(Sender: TObject);
begin
 ngImage.ClearRows;
 ngImage.AddRow(1);
end;

procedure TfrmImage.AddRow2Click(Sender: TObject);
begin
 ngImage.AddRow(1);
end;

procedure TfrmImage.Update1Click(Sender: TObject);
var
 i:integer;
 angka:String;
begin
  Screen.Cursor:=crHourGlass;
  for i := 0 to ngImage.RowCount-1 do
  begin
    angka:=fncangka;
    strSQL := 'SELECT * FROM M_IMAGE ';
    strSQL := strSQL + 'WHERE ICHANNEL = ''' + ngImage.Cells[4,i] + ''' ';
    strSQL := strSQL + 'AND IEPG_ORI = ''' + ngImage.Cells[5,i] + ''' ';
    strSQL := strSQL + 'AND IIMAGEID = ''' + ngImage.Cells[6,i] + ''' ';
    RecSet(strSQL);

    if dm.DDL.Eof then
    begin
      strSQL := 'INSERT INTO SGI.M_IMAGE ( ';
      strSQL := strSQL + 'IID, ICHANNEL, IEPG_ORI, IIMAGEID, ';
      strSQL := strSQL + 'IUSER_CREATE, ICREATE_DATE, IUSER_UPDATE, ';
      strSQL := strSQL + 'IUPDATE_DATE ) ';
      strSQL := strSQL + 'VALUES ( ';
      strSQL := strSQL + '''' + angka + ''', ';//IID
      strSQL := strSQL + '''' + TRIM(ngImage.Cells[1,i]) + ''', ';//ICHANNEL
      strSQL := strSQL + 'Upper(''' + UpperCase(trim(ngImage.Cells[2,i])) + '''), ';//IEPG_ORI
      strSQL := strSQL + '''' + TRIM(ngImage.Cells[3,i]) + ''', ';//IIMAGEID
      strSQL := strSQL + '''' + strUser + ''', ';//IUSER_CREATE
      strSQL := strSQL + 'sysdate, ';//ICREATE_DATE
      strSQL := strSQL + '''' + strUser + ''', ';//USER_UPDATE
      strSQL := strSQL + 'sysdate) ';//IUPDATE_DATE
      RecExc(strSQL);
    end
  else
    begin
      strSQL := ' UPDATE M_IMAGE SET ICHANNEL = ''' + ngImage.Cells[1,i] + ''', IEPG_ORI = ''' + UpperCase(ngImage.Cells[2,i]) + ''', IIMAGEID = ''' + ngImage.Cells[3,i] + ''' ';
      strSQL := strSQL + 'WHERE ICHANNEL = ''' + ngImage.Cells[4,i] + ''' ';
      strSQL := strSQL + ' and IEPG_ORI = ''' + ngImage.Cells[5,i] + ''' ';
      strSQL := strSQL + ' and IIMAGEID = ''' + ngImage.Cells[6,i] +''' ';
      RecExc(strSQL);
    end;
  end;
  ngImage.ClearRows;
  strSQL:='SELECT * FROM M_IMAGE ORDER BY ICHANNEL, IEPG_ORI';
  RecSet(strSQL);
  i:=1;
  while not dm.DDL.Eof do
  begin
    frmImage.ngImage.AddRow(1);
    frmImage.ngImage.BeginUpdate;
    frmImage.ngImage.Cell[0, i-1].AsString := inttostr(i);
    frmImage.ngImage.Cell[1, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
    frmImage.ngImage.Cell[2, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
    frmImage.ngImage.Cell[3, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
    frmImage.ngImage.Cell[4, i-1].AsString := dm.DDL.FieldValues['ICHANNEL'];
    frmImage.ngImage.Cell[5, i-1].AsString := dm.DDL.FieldValues['IEPG_ORI'];
    frmImage.ngImage.Cell[6, i-1].AsString := dm.DDL.FieldValues['IIMAGEID'];
    frmImage.ngImage.EndUpdate;
    dm.DDL.Next;
    i:=i+1;
  end;
  Screen.Cursor:=crDefault;
  ShowMessage('Data Has Been Saved!');
end;

procedure TfrmImage.ngImageKeyPress(Sender: TObject; var Key: Char);
begin
  if key=#13 then
  begin
    Update1Click(Sender);
  end;
end;

end.
