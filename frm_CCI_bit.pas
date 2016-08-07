unit frm_CCI_bit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TfrmCCIBits = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    lblChannel: TLabel;
    edtCCI: TEdit;
    Button1: TButton;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure edtCCIKeyPress(Sender: TObject; var Key: Char);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCCIBits: TfrmCCIBits;

implementation

uses frm_dm, frm_Login, frm_Channel;

{$R *.dfm}

function fncnmr():string;
var
 strNMR : integer;
begin
 strSQL := 'SELECT CBID FROM M_CHANNEL_BITS ORDER BY CBID ';
 RecSet(strSQL);

 if dm.DDL.Eof then
  begin
   strNMR := 1;
  end
 else
  begin
   dm.ddl.Last;
   strNMR := dm.DDL.FieldValues['CBID'] + 1;
  end;
 fncnmr:=IntToStr(strNMR);
end;

procedure prcSave (chn:String; cci:string);
var
 i:integer;
 nmr:String;
begin
  strSQL := 'SELECT * FROM M_CHANNEL_BITS WHERE CBCHANNEL = ''' + frmChannel.ngChannel.Cells[1,0] + ''' ';
  RecSet(strSQL);

nmr:=fncnmr;
  strSQL := 'SELECT * FROM M_CHANNEL_BITS ';
  strSQL := strSQL + 'WHERE CBCHANNEL = ''' + frmCCIBits.lblChannel.Caption + ''' ';
  RecSet(strSQL);

  if dm.DDL.Eof then
    begin
      strSQL := 'INSERT INTO SGI.M_CHANNEL_BITS ( ';
      strSQL := strSQL + 'CBID, CBCHANNEL, CBNUMBER ) ';
      strSQL := strSQL + 'VALUES ( ';
      strSQL := strSQL + '''' + nmr + ''', ';
      strSQL := strSQL + '''' + frmCCIBits.lblChannel.Caption + ''', ';
      strSQL := strSQL + '''' + frmCCIBits.edtCCI.Text  + ''') ';
      RecExc(strSQL);
      ShowMessage('CCI Bits Inserted!');
      frmCCIBits.Close;
    end
  else
    begin
      strSQL := ' UPDATE M_CHANNEL_BITS SET CBNUMBER = ''' + frmCCIBits.edtCCI.Text + ''' ';
      strSQL := strSQL + ' WHERE CBCHANNEL = ''' + frmCCIBits.lblChannel.Caption + ''' ';
      RecExc(strSQL);
      ShowMessage('CCI Bits Updated!');
      frmCCIBits.Close;
    end
end;

procedure TfrmCCIBits.Button1Click(Sender: TObject);
begin
if edtCCI.Text='' then
  begin
    ShowMessage('You Must Input 2 Digits HEX Bits Number!');
  end
else
  begin
    Screen.Cursor:=crHourGlass;
      prcSave(frmCCIBits.edtCCI.Text,frmCCIBits.lblChannel.Caption);
    Screen.Cursor:=crDefault;
  end;
end;



procedure TfrmCCIBits.FormShow(Sender: TObject);
var
 item : TStrings;
begin
  strSQL := 'SELECT * FROM M_CHANNEL_BITS WHERE CBCHANNEL = ''' + frmChannel.ngChannel.Cells[1,0] + ''' ';
  RecSet(strSQL);

  if dm.DDL.Eof then
    begin
      lblChannel.Caption := frmChannel.ngChannel.Cells[1,0];
    end
  else
    begin
      lblChannel.Caption := dm.DDL.FieldValues['CBCHANNEL'];
      edtCCI.Text := dm.DDL.FieldValues['CBNUMBER'];
    end;

end;

procedure TfrmCCIBits.edtCCIKeyPress(Sender: TObject; var Key: Char);
begin
if key=#13 then
  begin
    if edtCCI.Text='' then
      begin
        ShowMessage('You Must Input 2 Digits HEX Bits Number!');
      end
    else
      begin
        Screen.Cursor:=crHourGlass;
          prcSave(frmCCIBits.edtCCI.Text,frmCCIBits.lblChannel.Caption);
        Screen.Cursor:=crDefault;
      end;
  end;
end;

procedure TfrmCCIBits.Button2Click(Sender: TObject);
begin
frmCCIBits.Close; 
end;

end.
