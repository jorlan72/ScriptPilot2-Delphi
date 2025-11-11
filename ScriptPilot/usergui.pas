unit usergui;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, intercode, Vcl.ExtCtrls, Vcl.WinXPanels,
  Vcl.StdCtrls, System.UITypes, myMessage, myDialog, Vcl.ComCtrls;

type
  TfrmUserGUI = class(TForm)
    PanelGUITop: TPanel;
    PanelGUIBottom: TPanel;
    PanelGUIClient: TPanel;
    CardPanelGUIClient: TCardPanel;
    CardPanelGUITop: TCardPanel;
    CardPanelGUIBottom: TCardPanel;
    CardGUITopText: TCard;
    CardGUIBottomMessage: TCard;
    LabelTopText: TLabel;
    LabelMessageOKButtonBottom: TLabel;
    ButtonMessageBottom: TButton;
    PanelInterrupt: TPanel;
    ButtonScriptAbort: TButton;
    CardGUIBottomBlank: TCard;
    CardGUIClientBlank: TCard;
    CardGUITopBlank: TCard;
    CardGUITopProgressBar: TCard;
    GUITopProgressBar: TProgressBar;
    CardGUIBottomProgressBar: TCard;
    GUIBottomProgressBar: TProgressBar;
    CardGUIBottomText: TCard;
    LabelBottomText: TLabel;
    CardGUIClientText: TCard;
    LabelClientText: TLabel;
    procedure ButtonMessageBottomClick(Sender: TObject);
    procedure ButtonScriptAbortClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmUserGUI: TfrmUserGUI;
  MyMessageTitle, MyMessageBody : String;

procedure ShowViewMessage(Title, Msg: String; Modal: Boolean; Index: Integer);
function ShowViewDialog(Title: String; Index: Integer): TModalResult;

implementation

{$R *.dfm}

procedure TfrmUserGUI.ButtonScriptAbortClick(Sender: TObject);
var
  UserResponse : Integer;
begin
  Debug('User clicked the Abort button. Asking if user wants to stop script execution');
  UserResponse := ShowViewDialog('Do you want to stop script execution?', 0);
  if UserResponse = mrYes then
    begin
      Debug('User selected to abort the script');
      AbortButtonClicked := True;
      frmUserGUI.BorderStyle := bsSizeable;
      frmUserGUI.WindowState := wsNormal;
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowTop', IntToStr(Integer(frmUserGUI.Top)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowLeft', IntToStr(Integer(frmUserGUI.Left)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowWidth', IntToStr(Integer(frmUserGUI.Width)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowHeight', IntToStr(Integer(frmUserGUI.Height)));
    end else
    begin
      Debug('User selected NOT to abort the script. Script will continue');
      exit;
    end;
end;

procedure TfrmUserGUI.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ScriptRunning = True then
    begin
      ShowViewMessage('Script is currently running!', 'Abort script first or wait for the script to end', False, 0);
      CanClose := False;
      exit;
    end;
end;

procedure TfrmUserGUI.ButtonMessageBottomClick(Sender: TObject);
begin
  OKButtonPressed := True;
end;

procedure ShowViewMessage(Title, Msg: String; Modal: Boolean; Index: Integer);
// MY PERSONAL SHOWMESSAGE
var
  Frm: TfrmMessage;
  PrevAlpha : Boolean;
  PrevAlphaValue : Integer;
begin
  PrevAlpha := frmUserGUI.AlphaBlend;
  PrevAlphaValue := frmUserGUI.AlphaBlendValue;
  frmUserGUI.AlphaBlend := true;
  frmUserGUI.AlphaBlendValue := 200;
  Frm := TfrmMessage.Create(frmUserGUI);
  Frm.LabelMessageTitle.Caption := Title;
  Frm.LabelMessageText.Caption := Msg;
  if Modal then Frm.ShowModal else Frm.Show;
  frmUserGUI.AlphaBlend := PrevAlpha;
  frmUserGUI.AlphaBlendValue := PrevAlphaValue;
end;

function ShowViewDialog(Title: String; Index: Integer): TModalResult;
// MY PERSONAL SHOWMESSAGEDLG
var
  Frm: TfrmMyDialog;
  PrevAlpha : Boolean;
  PrevAlphaValue : Integer;
begin
  PrevAlpha := frmUserGUI.AlphaBlend;
  PrevAlphaValue := frmUserGUI.AlphaBlendValue;
  frmUserGUI.AlphaBlend := true;
  frmUserGUI.AlphaBlendValue := 200;
  Frm := TfrmMyDialog.Create(frmUserGUI);
  try
    Frm.LabelMyQuestion.Caption := Title;
    Result := Frm.ShowModal;
  finally
    Frm.Free;
  end;
  frmUserGUI.AlphaBlend := PrevAlpha;
  frmUserGUI.AlphaBlendValue := PrevAlphaValue;
end;

end.
