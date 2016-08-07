unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, Buttons;

type
  Tfrm_Splash = class(TForm)
    Image1: TImage;
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    procedure SpeedButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frm_Splash: Tfrm_Splash;

implementation

uses TesADO;

{$R *.dfm}

procedure Tfrm_Splash.SpeedButton1Click(Sender: TObject);
begin
Form1.Show;
end;

end.
