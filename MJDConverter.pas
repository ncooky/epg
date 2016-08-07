unit MJDConverter;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DateUtils;

type
  TfrmMJD = class(TForm)
    EdtYear: TEdit;
    Label1: TLabel;
    EdtMonth: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    EdtDate: TEdit;
    EdtHour: TEdit;
    Label4: TLabel;
    EdtMin: TEdit;
    Label5: TLabel;
    EdtSec: TEdit;
    Label6: TLabel;
    Button1: TButton;
    LblRes: TLabel;
    LblRes2: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMJD: TfrmMJD;

implementation

{$R *.dfm}

procedure TfrmMJD.Button1Click(Sender: TObject);
var
jDate: TDateTime;
mjdfloat: extended;
bfloat: extended;

begin
LblRes.Caption  := '';
LblRes2.Caption  := '';
LblRes.Caption := 'Result Julian Date = ';
LblRes2.Caption := 'Result Date = ';


jdate := strtodatetime(EdtDate.Text+'/'+EdtMonth.Text+'/'+EdtYear.Text+' '+edthour.Text+':'+edtmin.Text+':'+edtsec.Text+' AM');
//jdate := strtodatetime(edthour.Text+':'+edtmin.Text+':'+edtsec.Text+' AM');
LblRes2.Caption:= LblRes2.Caption + datetimetostr(jdate);

mjdfloat := DateTimeToJulianDate(jdate);
bfloat := mjdfloat - 2440000.5;

//LblRes.Caption:= LblRes.Caption + FloatToStr(DateTimeToJulianDate(jdate));
LblRes.Caption:= LblRes.Caption + formatfloat('0.######0', bfloat);



end;

end.
