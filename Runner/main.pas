unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.WinXPanels, Vcl.ExtCtrls, Vcl.Grids,
  Vcl.ValEdit, Vcl.StdCtrls, Intercode;

type
  TfrmRunner = class(TForm)
    CardPanelRunner: TCardPanel;
    CardVariables: TCard;
    GroupBoxVariables: TGroupBox;
    ValueListEditorVariables: TValueListEditor;
    CardIncoming: TCard;
    GroupBoxIncoming: TGroupBox;
    MemoIncomingMessages: TMemo;
    TimerStartupShit: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure TimerStartupShitTimer(Sender: TObject);
  private
    // INCOMING WINDOWS MESSAGES FROM SCRIPTPILOT RECEIVED VIA VMCOPYDATA
    procedure WMCopyData(var Msg: TWMCopyData); message WM_COPYDATA;
  public
    { Public declarations }
  end;

var
  frmRunner: TfrmRunner;

implementation

{$R *.dfm}

procedure TFrmRunner.WMCopyData(var Msg: TWMCopyData);
// INCOMING WINDOWS MESSAGES FROM SCRIPTPILOT RECEIVED VIA VMCOPYDATA
// RECEIVEDSTRING CONTAINS STRINGDATA AND STRINGCODE (MESSAGEID)
// STRINGCODE USED TO IDENTIFY WHAT KIND OF STRING IS BEING RECEIVED
var
  ReceivedString: string;
  MessageId: Cardinal;
begin
  MessageId := Msg.CopyDataStruct^.dwData;
  SetLength(ReceivedString, Msg.CopyDataStruct^.cbData div SizeOf(Char));
  Move(Msg.CopyDataStruct^.lpData^, PChar(ReceivedString)^, Msg.CopyDataStruct^.cbData);
    case MessageId of
      0: begin
           frmRunner.MemoIncomingMessages.Lines.Add('Received echo message from the boss. Replying back');
           Echo('ScriptPilot 2', ParamStr(1));
         end;
      1: begin
           frmRunner.MemoIncomingMessages.Lines.Add(ReceivedString);
         end;
      90: begin
           frmRunner.MemoIncomingMessages.Lines.Add(ReceivedString);
           Application.Terminate;
         end;
    end;
end;

procedure TfrmRunner.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // save runner form size, position and state - we do this in any case, even if the option us unchecked in settings
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowState', IntToStr(Integer(frmRunner.WindowState)));
  // only save form position and size if the form is <> maximized or minimized
  if frmRunner.WindowState = wsNormal then
    begin
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowTop', IntToStr(Integer(frmRunner.Top)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowLeft', IntToStr(Integer(frmRunner.Left)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowWidth', IntToStr(Integer(frmRunner.Width)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowHeight', IntToStr(Integer(frmRunner.Height)));
    end;
end;

procedure TfrmRunner.FormCreate(Sender: TObject);
begin
  if ParamStr(1) <> '' then
  begin // if started from ScriptPilot
    frmRunner.MemoIncomingMessages.Lines.Add('');
    frmRunner.MemoIncomingMessages.Lines.Add('ScriptPilot Runner Version ' + GetApplicationVersion);
    frmRunner.MemoIncomingMessages.Lines.Add('Runner name      : ' + ParamStr(1));
    frmRunner.MemoIncomingMessages.Lines.Add('Kill On Idle     : ' + ParamStr(2));
    frmRunner.MemoIncomingMessages.Lines.Add('');
    frmRunner.MemoIncomingMessages.Lines.Add('Ready.');
    frmRunner.MemoIncomingMessages.Lines.Add('');
    frmRunner.Caption := ParamStr(1);
    SendWinMessage('ScriptPilot 2', 'Runner ' + ParamStr(1) + ': Ready', 3);
    SendWinMessage('ScriptPilot 2', 'Runner ' + ParamStr(1) + ': Ready', 4);
  end else
  begin // if started from explorer or command line
    frmRunner.MemoIncomingMessages.Lines.Add('');
    frmRunner.MemoIncomingMessages.Lines.Add('ScriptPilot Runner Version ' + GetApplicationVersion);
    frmRunner.MemoIncomingMessages.Lines.Add('Runner name : Lost Runner');
    frmRunner.MemoIncomingMessages.Lines.Add('Must be called from ScriptPilot');
    frmRunner.MemoIncomingMessages.Lines.Add('');
    frmRunner.MemoIncomingMessages.Lines.Add('Not Ready.');
    frmRunner.MemoIncomingMessages.Lines.Add('');
    frmRunner.MemoIncomingMessages.Lines.Add('Please close this window and use the "RunnerActive" command from ScriptPilot');
    frmRunner.Caption := 'Lost Runner';
  end;
end;

procedure TfrmRunner.FormShow(Sender: TObject);
begin
// load runner form state and attributes
  if ParamStr(3) <> 'HideRunner' then
    begin
      frmRunner.AlphaBlend := true;
      frmRunner.AlphaBlendValue := 255;
      if StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveFormPositionAndState', 'false')) = true then
        begin
          frmRunner.Top := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowTop', '100'));
          frmRunner.Left := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowLeft', '320'));
          frmRunner.Width := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowWidth', '1244'));
          frmRunner.Height := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowHeight', '761'));
          if StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunnerWindowState', '0')) <> 2 then frmRunner.WindowState := wsNormal else frmRunner.WindowState := wsMaximized;
        end;
      frmRunner.AlphaBlend := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendScriptRunner', 'false'));
      frmRunner.AlphaBlendValue := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendScriptRunnerLevel', '255'));
    end else
    begin
      // enable timer to set additional settings that can not be called from onCreate or onShow
      frmRunner.TimerStartupShit.Enabled := true;
    end;
end;

procedure TfrmRunner.TimerStartupShitTimer(Sender: TObject);
begin
  frmRunner.TimerStartupShit.Enabled := false;
  frmRunner.Visible := false;
end;

end.
