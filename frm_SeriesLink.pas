unit frm_SeriesLink;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid, Menus, NxColumnClasses, NxColumns, StdCtrls, ComCtrls, jpeg,
  Grids, Buttons;

type
  TfrmSeriesLink = class(TForm)
    ppmSeriesLink: TPopupMenu;
    AddRow1: TMenuItem;
    AddNew1: TMenuItem;
    Update1: TMenuItem;
    Exit1: TMenuItem;
    ScrollBox1: TScrollBox;
    Image1: TImage;
    ScrollBox2: TScrollBox;
    ngSeries: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn8: TNxTextColumn;
    NxTextColumn10: TNxTextColumn;
    NxTextColumn11: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    edtEpg: TEdit;
    pbSeriesLink: TProgressBar;
    Delete1: TMenuItem;
    StringGrid1: TStringGrid;
    btnImport: TBitBtn;
    OpenDialog1: TOpenDialog;
    BitBtn1: TBitBtn;
    cboFilterSeries: TComboBox;
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    procedure ngSeriesMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure AddRow1Click(Sender: TObject);
//    procedure FormResize(Sender: TObject);
    procedure Update1Click(Sender: TObject);
    procedure AddNew1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure edtEpgKeyPress(Sender: TObject; var Key: Char);
    procedure ngSeriesSelectCell(Sender: TObject; ACol,
      ARow: Integer);
    procedure Delete1Click(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  X, Y : Integer;
  frmSeriesLink: TfrmSeriesLink;
  Procedure prcShowSeries(epg:string);
  function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string): Boolean;
  function Replace(Dest, SubStr, Str: string): string;

implementation

uses ComObj, frm_dm, frm_Login;

{$R *.dfm}

function fncAutoID():string;
var
 strID : integer;
begin
 strSQL := 'SELECT SRID FROM M_SERIES ORDER BY SRID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strID := 1;
  end
 else
  begin
   dm.ddl.Last;
   strID := dm.DDL.FieldValues['SRID'] + 1;
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

Procedure prcShowSeries(epg:string);
var
 strSQL : String;
 grpKey, prgKey : String;
 i,ii : integer;
 xyz : byte;
begin
 frmSeriesLink.ngSeries.ClearRows;
 strSQL := 'select Jumlah_Series, SREPGTITLE, SRGROUPKEY, SRPROGRAMKEY ';
 strSQL := strSQL + 'FROM (SELECT COUNT(0) AS Jumlah_Series FROM M_Series WHERE ';
 //strSQL := strSQL + ' SREPG_TITLE LIKE ''%' + epg + '%'' ';
 if frmSeriesLink.cboFilterSeries.Text = 'Tittle'
    Then strSQL := strSQL + '      SREPGTITLE LIKE ''%' + epg + '%'' ';
 if frmSeriesLink.cboFilterSeries.Text = 'Group Key'
    Then strSQL := strSQL + '      SRGROUPKEY LIKE ''%' + epg + '%'' ';
 if frmSeriesLink.cboFilterSeries.Text = 'Program Key'
    Then strSQL := strSQL + '      SRPROGRAMKEY LIKE ''%' + epg + '%'' ';

 strSQL := strSQL + ' )xxx, ';
 strSQL := strSQL + '(SELECT SREPGTITLE, SRGROUPKEY, SRPROGRAMKEY FROM M_Series WHERE ';
 //strSQL := strSQL + ' SREPG_TITLE LIKE ''%' + epg + '%'' ';
  if frmSeriesLink.cboFilterSeries.Text = 'Tittle'
    Then strSQL := strSQL + '      SREPGTITLE LIKE ''%' + epg + '%'' ';
 if frmSeriesLink.cboFilterSeries.Text = 'Group Key'
    Then strSQL := strSQL + '      SRGROUPKEY LIKE ''%' + epg + '%'' ';
 if frmSeriesLink.cboFilterSeries.Text = 'Program Key'
    Then strSQL := strSQL + '      SRPROGRAMKEY LIKE ''%' + epg + '%'' ';

 strSQL := strSQL + ' )yyy ';
 strSQL := strSQL + 'ORDER BY SREPGTITLE ';
 RecSet2(strSQL);

 xyz := dm.DDL2.FieldValues['Jumlah_Series'];

 i:=1;
 if not VarIsNull(dm.DDL2.FieldValues['Jumlah_Series']) then
  begin
   ii:=dm.DDL2.FieldValues['Jumlah_Series'];
  end
 else
  begin
   ii:=0;
  end;
 frmSeriesLink.ngSeries.AddRow(ii);
 frmSeriesLink.pbSeriesLink.Min:=0;
 frmSeriesLink.ngSeries.BeginUpdate;
 frmSeriesLink.pbSeriesLink.Visible:=True;
 frmSeriesLink.pbSeriesLink.Max:=ii;
 while not dm.DDL2.Eof do
  begin

   if not VarIsNull(dm.DDL2.FieldValues['SRGROUPKEY'])then
    begin
     grpKey:=dm.DDL2.FieldValues['SRGROUPKEY'];
    end
   else grpKey:='';

   if not VarIsNull(dm.DDL2.FieldValues['SRPROGRAMKEY'])then
    begin
     prgKey:=dm.DDL2.FieldValues['SRPROGRAMKEY'];
    end
   else prgKey:='';

   frmSeriesLink.ngSeries.Cell[0, i-1].AsString := inttostr(i);
   frmSeriesLink.ngSeries.Cell[1, i-1].AsString := dm.DDL2.FieldValues['SREPGTITLE'];
   frmSeriesLink.ngSeries.Cell[2, i-1].AsString := grpKey;
   frmSeriesLink.ngSeries.Cell[3, i-1].AsString := prgKey;

   i:=i+1;
   frmSeriesLink.pbSeriesLink.Position:=i;
   dm.DDL2.Next;
  end;
 frmSeriesLink.ngSeries.EndUpdate;
 frmSeriesLink.pbSeriesLink.Visible:=False;
end;

procedure TfrmSeriesLink.ngSeriesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if Button = mbRight Then
    Begin
      ppmSeriesLink.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
    End;
end;

procedure TfrmSeriesLink.AddRow1Click(Sender: TObject);
begin
 ngSeries.AddRow(1);
end;

procedure TfrmSeriesLink.Update1Click(Sender: TObject);
var
 i : integer;
 epgReplace, grpKey, prgKey, angka : string;
begin
 for i := 0 to ngSeries.RowCount-1 do
 begin
  angka:=fncAutoID;
  strSQL := 'SELECT * FROM M_Series ';
  strSQL := strSQL + 'WHERE SREPGTITLE = ''' + ngSeries.Cells[1,i] + ''' ';
  RecSet(strSQL);

   epgReplace:=Replace(trim(ngSeries.Cells[1,i]), '''','');
   epgReplace:=Replace(trim(epgReplace), '"','');

   if not VarIsNull(ngSeries.Cells[2,i])then
    begin
     grpKey:=Replace(trim(ngSeries.Cells[2,i]), '''','`');
     grpKey:=Replace(trim(grpKey), '`','''''');
     grpKey:=Replace(trim(grpKey), '"','``');
     grpKey:=Replace(trim(grpKey),#13,'');
    end
   else grpKey:='';

   if not VarIsNull(ngSeries.Cells[3,i])then
    begin
     prgKey:=Replace(trim(ngSeries.Cells[3,i]), '''','`');
     prgKey:=Replace(trim(prgKey), '`','''''');
     prgKey:=Replace(trim(prgKey), '"','``');
    end
   else prgKey:='';

  if dm.DDL.Eof then
  begin
   strSQL := 'INSERT INTO SGI.M_Series ( ';
   strSQL := strSQL + 'SRID, SREPGTITLE, SRGROUPKEY, SRPROGRAMKEY, ';
   strSQL := strSQL + 'SRUSERCREATE, SRUSERCREATEDATE, ';
   strSQL := strSQL + 'SRUSERUPDATE, SRUSERUPDATEDATE) ';
   strSQL := strSQL + 'VALUES ( ';
   strSQL := strSQL + '''' + angka + ''', ';
   strSQL := strSQL + '''' + epgReplace + ''', ';
   strSQL := strSQL + '''' + grpKey + ''', ';
   strSQL := strSQL + '''' + prgKey +  ''', ';
   strSQL := strSQL + '''' + strUser +  ''', ';
   strSQL := strSQL + 'sysdate, ';
   strSQL := strSQL + '''' + strUser +  ''', ';
   strSQL := strSQL + 'sysdate ) ';
   RecExc(strSQL);
  end
  else
  begin
   strSQL := ' UPDATE M_Series SET SREPGTITLE = ''' + epgReplace + ''', ';
   strSQL := strSQL + ' SRGROUPKEY = ''' + grpKey + ''', ';
   strSQL := strSQL + ' SRPROGRAMKEY = ''' + prgKey + ''', ';
   strSQL := strSQL + ' SRUSERUPDATE = ''' + strUser + ''', ';
   strSQL := strSQL + ' SRUSERUPDATEDATE = sysdate ';
   strSQL := strSQL + ' WHERE SREPGTITLE = ''' + ngSeries.Cells[1,i] + ''' ';
   RecExc(strSQL);
  end;

 end;
 ngSeries.ClearRows;
 prcShowSeries(edtEpg.Text);
 ShowMessage('Data Has Been Saved!');
end;

procedure TfrmSeriesLink.AddNew1Click(Sender: TObject);
begin
 ngSeries.ClearRows;
 ngSeries.AddRow(1);
end;


procedure TfrmSeriesLink.Exit1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmSeriesLink.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
 frmSeriesLink.Close;
end;

procedure TfrmSeriesLink.edtEpgKeyPress(Sender: TObject;
  var Key: Char);
begin
if key=#13 then
  begin
    Screen.Cursor:=crHourGlass;
    ngSeries.ClearRows;
    if trim(edtEpg.text)<>'' then
    begin
      prcShowSeries(edtEpg.Text);
    end;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmSeriesLink.ngSeriesSelectCell(Sender: TObject;
  ACol, ARow: Integer);
begin
  X := ACol;
  Y := ARow;
end;

procedure TfrmSeriesLink.Delete1Click(Sender: TObject);
var
  strDelSys : String;
begin
   Screen.Cursor := crHourGlass;
   strDelSys := ngSeries.Cells[1,y];
   strSQL := 'DELETE FROM M_Series WHERE SREPGTITLE = '''+trim(ngSeries.Cells[1,y])+''' ';
   RecExc(strSQL);
   prcShowSeries(edtEpg.Text);
   ShowMessage('Series '+strDelSys+' Has Been Removed');
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

procedure TfrmSeriesLink.btnImportClick(Sender: TObject);
var
 i : integer;
 IDTable, epgReplace, grpKey, prgKey : String;
begin

 if OpenDialog1.Execute then
 begin
 i := 1;
 Screen.Cursor:=crHourGlass;


  if Xls_To_StringGrid(StringGrid1, OpenDialog1.FileName) then
    begin
     pbSeriesLink.Max := StringGrid1.RowCount - 2;
     pbSeriesLink.Min := 0;
      while StringGrid1.Cells[0, i]<>'<eof>' do
        begin
          IDTable:=fncAutoID;
          epgReplace:=Replace(trim(StringGrid1.Cells[1,i]), '''','');
          epgReplace:=Replace(trim(epgReplace), '"','');

          if not VarIsNull(StringGrid1.Cells[2,i])then
            begin
              grpKey:=Replace(trim(StringGrid1.Cells[2,i]), '''','`');
              grpKey:=Replace(trim(grpKey), '`','''''');
              grpKey:=Replace(trim(grpKey), '�','"');
              grpKey:=Replace(trim(grpKey), '�','"');
            end
          else grpKey:='';

           if not VarIsNull(StringGrid1.Cells[3,i])then
              begin
               prgKey:=Replace(trim(StringGrid1.Cells[3,i]), '''','`');
               prgKey:=Replace(trim(prgKey), '`','''''');
               prgKey:=Replace(trim(prgKey), '�','"');
               prgKey:=Replace(trim(prgKey), '�','"');
              end
           else prgKey:='';

          strSQL := 'SELECT SREPGTITLE FROM M_SERIES WHERE SREPGTITLE = ''' + epgReplace + ''' ';
          RecSet(strSQL);
          if not dm.DDL.Eof then
            begin
              strSQL := 'delete FROM M_SERIES WHERE SREPGTITLE = ''' + epgReplace + ''' ';
              RecExc(strSQL);

              strSQL := 'INSERT INTO SGI.M_SERIES ( ';
              strSQL := strSQL + 'SRID, SREPGTITLE, SRGROUPKEY, SRPROGRAMKEY, ';
              strSQL := strSQL + 'SRUSERCREATE, SRUSERCREATEDATE, ';
              strSQL := strSQL + 'SRUSERUPDATE, SRUSERUPDATEDATE) ';
              strSQL := strSQL + 'VALUES ( ';
              strSQL := strSQL + '''' + IDTable + ''', ';
              strSQL := strSQL + '''' + epgReplace + ''', ';
              strSQL := strSQL + '''' + grpKey + ''', ';
              strSQL := strSQL + '''' + prgKey + ''', ';
              strSQL := strSQL + '''' + strUser + ''', ';
              strSQL := strSQL + 'sysdate, ';
              strSQL := strSQL + '''' + strUser + ''', ';
              strSQL := strSQL + 'sysdate) ';
              RecExc(strSQL);
            end;

           if dm.DDL.Eof then
            begin
             strSQL := 'INSERT INTO SGI.M_SERIES ( ';
             strSQL := strSQL + 'SRID, SREPGTITLE, SRGROUPKEY, SRPROGRAMKEY, ';
             strSQL := strSQL + 'SRUSERCREATE, SRUSERCREATEDATE, ';
             strSQL := strSQL + 'SRUSERUPDATE, SRUSERUPDATEDATE) ';
             strSQL := strSQL + 'VALUES ( ';
             strSQL := strSQL + '''' + IDTable + ''', ';
             strSQL := strSQL + '''' + epgReplace + ''', ';
             strSQL := strSQL + '''' + grpKey + ''', ';
             strSQL := strSQL + '''' + prgKey + ''', ';
             strSQL := strSQL + '''' + strUser + ''', ';
             strSQL := strSQL + 'sysdate, ';
             strSQL := strSQL + '''' + strUser + ''', ';
             strSQL := strSQL + 'sysdate) ';
             RecExc(strSQL);
            end;

          strSQL := 'SELECT * FROM M_SERIES WHERE SREPGTITLE = ''' + epgReplace + ''' ';
          RecSet(strSQL);
          while not dm.DDL.Eof do
            begin
             frmSeriesLink.ngSeries.AddRow(1);
             frmSeriesLink.ngSeries.BeginUpdate;
             frmSeriesLink.ngSeries.Cell[0, i-1].AsString := inttostr(i);
             frmSeriesLink.ngSeries.Cell[1, i-1].AsString := dm.DDL.FieldValues['SREPGTITLE'];
             frmSeriesLink.ngSeries.Cell[2, i-1].AsString := dm.DDL.FieldValues['SRGROUPKEY'];
             frmSeriesLink.ngSeries.Cell[3, i-1].AsString := dm.DDL.FieldValues['SRPROGRAMKEY'];;
             frmSeriesLink.ngSeries.EndUpdate;
             dm.DDL.Next;
            end;

          i:=i+1;
          pbSeriesLink.Position:=i-1;
        end;

     ShowMessage('Table has been exported!');
    end;
    Screen.Cursor:=crDefault;
  end;
end;

procedure TfrmSeriesLink.BitBtn1Click(Sender: TObject);
begin
  frmSeriesLink.Close;
end;

end.
