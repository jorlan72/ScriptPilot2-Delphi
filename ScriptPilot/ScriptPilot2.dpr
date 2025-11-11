program ScriptPilot2;

uses
  Vcl.Forms,
  Windows,
  main in 'main.pas' {frmMain},
  Vcl.Themes,
  Vcl.Styles,
  commands in 'commands.pas',
  intercode in 'intercode.pas',
  runner in 'runner.pas',
  usergui in 'usergui.pas' {frmUserGUI},
  mymessage in 'mymessage.pas' {frmMessage},
  mydialog in 'mydialog.pas' {frmMyDialog};

{$R *.res}

var
  // These are globally available for all units
  hWndScriptPilot: HWND;

begin
  // Check if a window with the title "ScriptPilot 2" already exists - if no parameters are found for executing runners.
  hWndScriptPilot := FindWindow(nil, 'ScriptPilot 2');
  if ParamCount = 0 then if hWndScriptPilot <> 0 then
    begin
    // Window found, bring it to the front
    SetForegroundWindow(hWndScriptPilot);
    // If the window is minimized, restore it
    ShowWindow(hWndScriptPilot, SW_RESTORE);
    Halt; // Terminate the application
    end;
  if ParamCount = 0 then // normal startup
    begin
      Application.Initialize;
      Application.MainFormOnTaskbar := True;
      TStyleManager.TrySetStyle('Windows10 BlackPearl');
      Application.Title := 'ScriptPilot 2.0';
      Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmUserGUI, frmUserGUI);
  Application.Run;
    end else
    begin // startup with parameters - right now, same as normal startup
      Application.Initialize;
      Application.MainFormOnTaskbar := True;
      TStyleManager.TrySetStyle('Windows10 BlackPearl');
      Application.Title := 'ScriptPilot 2.0';
      Application.CreateForm(TfrmMain, frmMain);
      Application.Run;
    end;
end.
