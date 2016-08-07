unit frm_Edit1CA;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NxColumns, NxColumnClasses, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, StdCtrls, Menus;

type
  TfrmEditCaEvent = class(TForm)
    ngEditCAEvent: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    NxTextColumn6: TNxTextColumn;
    Button1: TButton;
    Button2: TButton;
    ppmEditSch: TPopupMenu;
    InsertRow1: TMenuItem;
    Delete1: TMenuItem;
    Delete2: TMenuItem;
    procedure Button1Click(Sender: TObject);
    procedure ngEditCAEventMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure InsertRow1Click(Sender: TObject);
    procedure Delete1Click(Sender: TObject);
    procedure ngEditCAEventSelectCell(Sender: TObject; ACol,
      ARow: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEditCaEvent: TfrmEditCaEvent;
  x ,y : integer;

implementation

uses frm_dm, frm_EPG;

{$R *.dfm}

procedure TfrmEditCaEvent.Button1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmEditCaEvent.ngEditCAEventMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
 if Button = mbRight Then
  Begin
    ppmEditSch.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
  End;
end;

procedure TfrmEditCaEvent.InsertRow1Click(Sender: TObject);
begin
 ngEditCAEvent.InsertRow(y);
end;

procedure TfrmEditCaEvent.Delete1Click(Sender: TObject);
begin
 ngEditCAEvent.DeleteRow(y);
end;

procedure TfrmEditCaEvent.ngEditCAEventSelectCell(Sender: TObject; ACol,
  ARow: Integer);
begin
 x:=ACol;
 y:=ARow;
end;

end.
