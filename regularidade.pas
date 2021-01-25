unit regularidade;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, SHDocVw, StdCtrls, ExtCtrls, jpeg;

type
  TForm1 = class(TForm)
    WebBrowser1: TWebBrowser;
    Edit1: TEdit;
    Button1: TButton;
    Edit2: TEdit;
    Edit3: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    EditHd: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Shape1: TShape;
    Shape2: TShape;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure WebBrowser1DocumentComplete(Sender: TObject;
      const pDisp: IDispatch; var URL: OleVariant);
    procedure Edit2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit3KeyPress(Sender: TObject; var Key: Char);
    procedure EditHdKeyPress(Sender: TObject; var Key: Char);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure Edit1Change(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure Edit3Change(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  prefixo, hora, minuto: Integer;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin

webbrowser1.Navigate('http://webcptm/operacao/Regularidade/ARQTXT/regImporta.asp');

prefixo := 0;
hora := 0;
minuto := 0;

end;

procedure TForm1.Button1Click(Sender: TObject);
begin

if ((edithd.Text = '') or (edit1.text = '') or (edit2.text = '') or (edit3.text = '')) then

Messagebox(0,'Não pode haver campos em branco, Verifique!','Regularidade Automárica',mb_ok)

else

begin
try
WebBrowser1.OleObject.Document.all.Item('hrparh', 0).value := edit1.text;
WebBrowser1.OleObject.Document.all.Item('hrparm', 0).value := edit2.text;
WebBrowser1.OleObject.Document.all.Item('hrpars', 0).value := '00';
WebBrowser1.OleObject.Document.all.Item('txprf', 0).value := 'ud' + edit3.text;
WebBrowser1.OleObject.Document.all.Item('btact1', 0).click;

minuto := strtoint(edit2.text);
minuto := minuto + strtoint(editHd.text);

if minuto >= 60 then
begin
hora := strtoint(edit1.text);
hora := hora + 1;

if hora = 24 then hora := 00;

edit1.text := inttostr(hora);

minuto := minuto - 60;
end;


edit2.text := inttostr(minuto);



prefixo := strtoint(edit3.Text);
prefixo := prefixo + 2;
edit3.Text := inttostr(prefixo);

except
Messagebox(0,'Erro ao carregar a página!','Regularidade Automática',mb_ok);
end;

end;

end;

procedure TForm1.WebBrowser1DocumentComplete(Sender: TObject;
  const pDisp: IDispatch; var URL: OleVariant);
begin
edit2.setfocus;
end;

procedure TForm1.Edit2KeyPress(Sender: TObject; var Key: Char);
begin
if not (key in ['0'..'9',#8,#13,#9]) = true then key := #0;
if key = #13 then
begin
  button1.SetFocus;
  button1.click;
  end;
end;

procedure TForm1.Edit3KeyPress(Sender: TObject; var Key: Char);
begin
if not (key in ['0'..'9',#8,#13,#9]) = true then key := #0;
if key = #13 then
begin
button1.setfocus;
button1.click;
end;
end;

procedure TForm1.EditHdKeyPress(Sender: TObject; var Key: Char);
begin
if not (key in ['0'..'9',#8,#13,#9]) = true then key := #0;

if key = #13 then perform(wm_nextdlgctl,0,0,);

end;

procedure TForm1.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
if not (key in ['0'..'9',#8,#13,#9]) = true then key := #0;
if key = #13 then perform(wm_nextdlgctl,0,0,);
end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
if length(tedit(sender).text) = 2 then edit2.setfocus;

end;

procedure TForm1.Edit2Change(Sender: TObject);
begin
if length(tedit(sender).text) = 2 then edit3.setfocus;
end;

procedure TForm1.Edit3Change(Sender: TObject);
begin
if length(tedit(sender).text) = 3 then button1.setfocus;
end;

end.
