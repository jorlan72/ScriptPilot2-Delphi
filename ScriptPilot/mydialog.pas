unit mydialog;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.pngimage,
  Vcl.ExtCtrls;

type
  TfrmMyDialog = class(TForm)
    LabelMyQuestion: TLabel;
    ButtonYes: TButton;
    ButtonCancel: TButton;
    ButtonNo: TButton;
    ImageMessage: TImage;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMyDialog: TfrmMyDialog;

implementation

{$R *.dfm}

end.
