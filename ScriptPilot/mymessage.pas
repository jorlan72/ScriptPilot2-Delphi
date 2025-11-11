unit mymessage;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.pngimage,
  Vcl.ExtCtrls, System.ImageList, Vcl.ImgList, Vcl.BaseImageCollection,
  Vcl.ImageCollection, Vcl.VirtualImage, Vcl.VirtualImageList;

type
  TfrmMessage = class(TForm)
    LabelMessageTitle: TLabel;
    LabelMessageText: TLabel;
    Button1: TButton;
    ImageMessage: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMessage: TfrmMessage;
  TempBitmap: TBitmap;

implementation

{$R *.dfm}

procedure TfrmMessage.Button1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmMessage.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmMessage.FormCreate(Sender: TObject);
begin
  TempBitmap := TBitmap.Create;
end;

end.
