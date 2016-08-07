unit frm_Check;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NxColumns, NxColumnClasses, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, Grids, StdCtrls, jpeg,
  ExtCtrls;

type
  TfrmCheck = class(TForm)
    ngCheck: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    StringGrid1: TStringGrid;
    Image1: TImage;
    Button1: TButton;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCheck: TfrmCheck;

implementation

uses ComObj, frm_Read, frm_dm, frm_Login;

{$R *.dfm}

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

procedure TfrmCheck.FormShow(Sender: TObject);
var
 i : integer;
 angka, epgReplace, synInd, synEng : String;
begin
 i := 1;
 Screen.Cursor:=crHourGlass;
  if Xls_To_StringGrid(StringGrid1, frmRead.OpenDialog1.FileName) then
    begin
     //pbSynopsis.Max := StringGrid1.RowCount - 2;
     //pbSynopsis.Min := 0;
      while StringGrid1.Cells[0, i]<>'<eof>' do
        begin
          //angka:=fncangka;
          epgReplace:=Replace(trim(StringGrid1.Cells[1,i]), '''','');
          epgReplace:=Replace(trim(epgReplace), '"','');

          strSQL := 'SELECT * FROM M_SYNOPSIS WHERE SYEPG_TITLE LIKE upper(''' + epgReplace + '%' + ''') ';
          strSQL := strSQL + 'AND SYCATEGORY = ''' + copy(StringGrid1.Cells[2,i],1,1) + ''' ';
          RecSet(strSQL);
          while not dm.DDL.Eof do
            begin
            frmCheck.ngCheck.AddCells([dm.DDL.FieldValues['SYEPG_TITLE'],
                                               dm.DDL.FieldValues['SYCATEGORY'],
                                               dm.DDL.FieldValues['sysynopsis_ind'],
                                               dm.DDL.FieldValues['sysynopsis_eng']
                                              ]);
            dm.DDL.next
            end;
         i:=i+1;
         //pbSynopsis.Position:=i-1;
        end;

   Screen.Cursor:=crDefault;
   ShowMessage('Table has been exported!');
  end;
end;

procedure TfrmCheck.Button1Click(Sender: TObject);
begin
  frmCheck.Close;
end;

end.
