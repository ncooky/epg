unit frm_About;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, jpeg;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    OKButton: TButton;
    LogoAJ: TImage;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutBox: TAboutBox;

implementation

{$R *.dfm}

function  GetAppVersion:string;
   var
    Size, Size2: DWord;
    Pt, Pt2: Pointer;
   begin
     Size := GetFileVersionInfoSize(PChar (ParamStr (0)), Size2);
     if Size > 0 then
     begin
       GetMem (Pt, Size);
       try
          GetFileVersionInfo (PChar (ParamStr (0)), 0, Size, Pt);
          VerQueryValue (Pt, '\', Pt2, Size2);
          with TVSFixedFileInfo (Pt2^) do
          begin
            // Result:= ' Ver '+
            Result:= IntToStr (HiWord (dwFileVersionMS)) + '.' +
                     IntToStr (LoWord (dwFileVersionMS)) + '.' +
                     IntToStr (HiWord (dwFileVersionLS)) + '.' +
                     IntToStr (LoWord (dwFileVersionLS));
         end;
       finally
         FreeMem (Pt);
       end;
     end;
   end;

procedure TAboutBox.FormShow(Sender: TObject);
begin
Version.Caption := 'Version';
Version.Caption := version.caption+' : ' + GetAppVersion;
end;

end.
 
