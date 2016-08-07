unit frm_ExEPG;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls;

type
  TfrmExEPG = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    dtpAwal: TDateTimePicker;
    dtpAkhir: TDateTimePicker;
    cbExEPGChannel: TComboBox;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmExEPG: TfrmExEPG;

implementation

uses frm_dm;

{$R *.dfm}

procedure TfrmExEPG.FormShow(Sender: TObject);
var
 strSQ : String;
 item : TStrings;
 i:integer;
begin
 i:=1;
 strSQL := 'SELECT distinct echannel ';
 strSQL := strSQL + 'FROM m_epg ';
 strSQL := strSQL + 'ORDER by echannel ';
 RecSet(strSQL);

 while not dm.DDL.Eof do
 begin
  strSQl := 'INSERT INTO M_channel ( ';
  strSQl := strSQl + 'MCID, MCHANNEL, MCHANNEL_DESCRIPTION, MCSISERVICEID ) ';
  strSQl := strSQl + 'VALUES ( ';
  strSQl := strSQl + '''' + IntToStr(i) + ''', ';
  strSQl := strSQl + '''' + dm.DDL.FieldValues['echannel'] + ''', ';
  strSQl := strSQl + '''' + dm.DDL.FieldValues['echannel'] + ''', ';
  strSQl := strSQl + '''1002'') ';
  RecExc(strSQl);
  dm.DDL.Next;


  i:=i+1;
 end;
//RecExc(strSQL);
 {Item:=cbDateStart.Items.Create;
 while not dm.DDL.Eof do
 begin
  item.Add(dm.DDL.FieldValues['Date_Schedule']);
  dm.DDL.Next;
 end;   }

end;

end.
