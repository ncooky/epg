unit frm_SynopsisXL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, jpeg, ExtCtrls, StdCtrls, ComCtrls, NxColumns,
  NxColumnClasses, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid;

type
  TfrmSynopsisXL = class(TForm)
    StringGrid1: TStringGrid;
    ScrollBox1: TScrollBox;
    Button1: TButton;
    pbSynopsis: TProgressBar;
    ngSynopsis: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    Image2: TImage;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSynopsisXL: TfrmSynopsisXL;

implementation

uses ComObj, frm_Read, frm_dm, frm_Login, DB;

{$R *.dfm}

function fncangka():string;
var
 strAngka : integer;
begin
 strSQL := 'SELECT max(SYID) as syid FROM M_SYNOPSIS ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strAngka := 1;
  end
 else
  begin
   dm.ddl.Last;
   if VarIsNull(dm.DDL.FieldValues['syid'])
    Then strAngka := 1
      Else strAngka := dm.DDL.FieldValues['SYID'] + 1;
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

procedure TfrmSynopsisXL.FormShow(Sender: TObject);
var
 i, ii : integer;
 angka, epgReplace, synInd, synEng : String;
begin
 i := 1;
 Screen.Cursor:=crHourGlass;
  if Xls_To_StringGrid(StringGrid1, frmRead.OpenDialog1.FileName) then
    begin
     pbSynopsis.Max := StringGrid1.RowCount - 2;
     pbSynopsis.Min := 0;
      while StringGrid1.Cells[0, i]<>'<eof>' do
        begin
          angka:=fncangka;
          epgReplace:=Replace(trim(StringGrid1.Cells[1,i]), '''','');
          epgReplace:=Replace(trim(epgReplace), '"','');

          if not VarIsNull(StringGrid1.Cells[2,i])then
            begin
              synInd:=Replace(trim(StringGrid1.Cells[2,i]), '''','`');
              synInd:=Replace(trim(synInd), '`','''''');
              //synInd:=Replace(trim(synInd), '"','`');
              synInd:=Replace(trim(synInd), '�','"');
              synInd:=Replace(trim(synInd), '�','"');
            end
          else synInd:='';

           if not VarIsNull(StringGrid1.Cells[3,i])then
              begin
               synEng:=Replace(trim(StringGrid1.Cells[3,i]), '''','`');
               synEng:=Replace(trim(synEng), '`','''''');
               //synEng:=Replace(trim(synEng), '"','');
               synEng:=Replace(trim(synEng), '�','"');
               synEng:=Replace(trim(synEng), '�','"');
              end
           else synEng:='';

          strSQL := 'SELECT SYEPG_TITLE FROM M_SYNOPSIS WHERE SYEPG_TITLE = upper(''' + epgReplace + ''') ';
          strSQL := strSQL + 'AND SYCATEGORY = ''' + copy(StringGrid1.Cells[7,i],1,2) + ''' ';
          RecSet(strSQL);
          if not dm.DDL.Eof then
            begin
              strSQL := 'delete FROM M_SYNOPSIS WHERE SYEPG_TITLE = upper(''' + epgReplace + ''') ';
              strSQL := strSQL + 'AND SYCATEGORY = ''' + copy(StringGrid1.Cells[7,i],1,2) + ''' ';
              RecExc(strSQL);

              strSQL := 'INSERT INTO SGI.M_SYNOPSIS ( ';
              strSQL := strSQL + 'SYID, SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, ';
              strSQL := strSQL + 'SYCATEGORY, SYUSER_CREATE, SYUSER_CREATEDATE, ';
              strSQL := strSQL + 'SYUSER_UPDATE, SYUSER_UPDATEDATE) ';
              strSQL := strSQL + 'VALUES ( ';
              strSQL := strSQL + '''' + angka + ''', ';
              strSQL := strSQL + 'Upper(''' + epgReplace + '''), ';
              strSQL := strSQL + '''' + synInd + ''', ';
              strSQL := strSQL + '''' + synEng + ''', ';
              strSQL := strSQL + '''' + copy(StringGrid1.Cells[7,i],1,2) + ''', ';
              strSQL := strSQL + '''' + strUser + ''', ';
              strSQL := strSQL + 'sysdate, ';
              strSQL := strSQL + '''' + strUser + ''', ';
              strSQL := strSQL + 'sysdate) ';
              RecExc(strSQL);
            end;

           if dm.DDL.Eof then
            begin
             strSQL := 'INSERT INTO SGI.M_SYNOPSIS ( ';
             strSQL := strSQL + 'SYID, SYEPG_TITLE, SYSYNOPSIS_IND, SYSYNOPSIS_ENG, ';
             strSQL := strSQL + 'SYCATEGORY, SYUSER_CREATE, SYUSER_CREATEDATE, ';
             strSQL := strSQL + 'SYUSER_UPDATE, SYUSER_UPDATEDATE) ';
             strSQL := strSQL + 'VALUES ( ';
             strSQL := strSQL + '''' + angka + ''', ';
             strSQL := strSQL + 'Upper(''' + epgReplace + '''), ';
             strSQL := strSQL + '''' + synInd + ''', ';
             strSQL := strSQL + '''' + synEng + ''', ';
             strSQL := strSQL + '''' + copy(StringGrid1.Cells[7,i],1,2) + ''', ';
             strSQL := strSQL + '''' + strUser + ''', ';
             strSQL := strSQL + 'sysdate, ';
             strSQL := strSQL + '''' + strUser + ''', ';
             strSQL := strSQL + 'sysdate) ';
             RecExc(strSQL);
            end;

          strSQL := 'SELECT * FROM M_SYNOPSIS WHERE SYEPG_TITLE = upper(''' + epgReplace + ''') ';
          strSQL := strSQL + 'AND SYCATEGORY = ''' + copy(StringGrid1.Cells[7,i],1,2) + ''' ';
          RecSet(strSQL);
          while not dm.DDL.Eof do
            begin
             frmSynopsisXL.ngSynopsis.AddRow(1);
             frmSynopsisXL.ngSynopsis.BeginUpdate;
             frmSynopsisXL.ngSynopsis.Cell[0, i-1].AsString := inttostr(i);
             frmSynopsisXL.ngSynopsis.Cell[1, i-1].AsString := dm.DDL.FieldValues['SYEPG_TITLE'];
             frmSynopsisXL.ngSynopsis.Cell[2, i-1].AsString := dm.DDL.FieldValues['SYCATEGORY'];
             frmSynopsisXL.ngSynopsis.Cell[3, i-1].AsString := dm.DDL.FieldValues['SYSYNOPSIS_IND'];;
             frmSynopsisXL.ngSynopsis.Cell[4, i-1].AsString := dm.DDL.FieldValues['SYSYNOPSIS_ENG'];;
             frmSynopsisXL.ngSynopsis.EndUpdate;
             dm.DDL.Next;
            end;

          i:=i+1;
          pbSynopsis.Position:=i-1;
        end;
     Screen.Cursor:=crDefault;
     ShowMessage('Table has been exported!');
    end;
end;

procedure TfrmSynopsisXL.Button1Click(Sender: TObject);
begin
 frmSynopsisXL.Close;
end;

end.
