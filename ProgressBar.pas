unit ProgressBar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls;

type
  TForm1 = class(TForm)
    ProgressBar: TProgressBar;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses frm_SynopsisManual;

{$R *.dfm}

procedure TForm1.FormActivate(Sender: TObject);
var
 i, ii : integer;
begin

end;

end.
