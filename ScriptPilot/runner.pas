unit runner;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Dialogs, System.IOUtils, System.UITypes, ShellAPI, StrUtils, Vcl.Forms,
  intercode, UserGUI;

var
  CommandToExecute : String;
  Parameter1, Parameter2, Parameter3, Parameter4, Parameter5 : Variant;

// ALL PROCEDURES BELOW CORRESPONDS TO A COMMAND - EVERY COMMAND RECORD WILL HAVE A PROCEDURE POINTER THAT POINTS TO A PROCEDURE WITH THE SAME NAME AS THE COMMAND ITSELF
Procedure Remark;
Procedure SetTag;
Procedure GotoTag;
Procedure SetRunnerActive;
Procedure WaitForButtonClick;
Procedure WaitForSeconds;
Procedure IncludeScript;
Procedure SetVariable;
Procedure UserViewTopText;
Procedure UserViewTopProgressBar;
Procedure UserViewBottomText;
Procedure UserViewBottomProgressBar;
Procedure UserViewClientText;
Procedure UserViewTopClear;
Procedure UserViewBottomClear;
Procedure UserViewClientClear;
Procedure IfVariable;

implementation

Procedure Remark;
begin
//
end;

Procedure SetTag;
begin
//
end;

Procedure WaitForButtonClick;
begin
  frmUserGUI.LabelMessageOKButtonBottom.Caption := Parameter1;
  frmUserGUI.ButtonMessageBottom.Caption := Parameter2;
  frmUserGUI.CardPanelGUIBottom.ActiveCard := frmUserGUI.CardGUIBottomMessage;
  Debug('Script paused. Waiting for end user to click OK button to continue');
  OKButtonPressed := false;
  while not OKButtonPressed do
    begin
      sleep(100);
      Application.ProcessMessages;
      if AbortButtonClicked then OKButtonPressed := true;
    end;
  frmUserGUI.CardPanelGUIBottom.ActiveCard := frmUserGUI.CardGUIBottomBlank;
  if not AbortButtonClicked then Debug('OK button clicked. Script will now continue');
end;

Procedure WaitForSeconds;
var
 interval : integer;
 counter : integer;
begin
  Debug('Pausing script for ' + Parameter1 + ' seconds');
  interval := StrToInt(Parameter1) * 10;
  for counter := interval downto 0 do
    begin
      sleep(100);
      application.ProcessMessages;
      if AbortButtonClicked then break;
    end;
  Debug(Parameter1 + ' second pause ended. Script will now continue');
end;

Procedure SetRunnerActive;
var
  MaxWait : Integer;
  ActiveRunner : TDataInfo;
begin
  Debug('Activating runner ' + Parameter1);
  if not Runners.KeyExists(Parameter1) then
    begin
      CurrentNumberOfRunners := Runners.Count;
      if (CurrentNumberOfRunners + 1) > MaxNumberOfRunners then
        begin
          Warning('The maximum numbers of runners has been reached');
          Debug('If we are not stopping on errors, then the current active runner will remain as active');
          Debug('Current active runner is : ' + CurrentRunner);
        end else
        begin
          AreRunnerReady := false;
          Debug('Creating runner and waiting for it to report back');
          if FileExists(FileRunnerExe) then
            begin
              // paramstr(0) = runner.exe (FileRunnerExe)
              // paramstr(1) = runner name
              // paramstr(2) = kill runner on idle - true or false
              // paramstr(3) = hide the runner - true or false
              // paramstr(4) = wait for the runner to finish (block main thread) - true or false
              if VarAsType(Parameter3, varBoolean) then Parameter2 := Parameter2 + ' HideRunner';
              LaunchProcess(FileRunnerExe, '"' + VarToStr(Parameter1) + '" ' + VarToStr(Parameter2), VarAsType(Parameter3, varBoolean), false);
            end else
            begin
              Error('Runner.exe not found');
              Exit;
            end;
          for MaxWait := 0 to 120 do
            begin
              Sleep(50);
              Application.ProcessMessages;
              if AreRunnerReady then break;
            end;
          if AreRunnerReady then
            begin
              Debug('Runner operative. Sending commands to ' + Parameter1);
              ActiveRunner.DataName := Parameter1;
              Runners.AddOrUpdate(Parameter1, ActiveRunner);
              CurrentRunner := Parameter1;
            end else
            begin
              Error('Timeout - Runner ' + Parameter1 + ' did not report back');
              Exit;
            end;
        end;
    end else
    begin
      Debug('Checking if ' + Parameter1 + ' is alive. Sending echo message');
      EchoReply := '';
      Echo(Parameter1, 'ScriptPilot 2');
      for MaxWait := 0 to 120 do
        begin
          Sleep(50);
          Application.ProcessMessages;
          if EchoReply <> '' then break;
        end;
      if EchoReply <> '' then
        begin
          Debug('Setting active and sending new commands to ' + Parameter1);
          CurrentRunner := Parameter1;
        end else
        begin
          Error('Timeout - Runner ' + Parameter1 + ' did not report back');
          Exit;
        end;
    end;
end;

Procedure IncludeScript;
begin
  SendWinMessage('ScriptPilot 2', 'IncludeScript,' + Parameter1, 1);
end;

Procedure SetVariable;
begin
  Debug('Updating variables dictionary with variable name "' + Parameter1 + '", and setting value to "' + Parameter2 + '"');
  Variables.AddOrUpdate(Parameter1, Parameter2);
end;

Procedure UserViewTopText;
begin
  frmUserGUI.CardPanelGUITop.ActiveCard := frmUserGUI.CardGUITopText;
  frmUserGUI.LabelTopText.Caption := Parameter1;
  frmUserGUI.LabelTopText.Visible := true;
  Debug('Setting User View top panel to textmode and showing text "' + Parameter1 + '"');
end;

Procedure UserViewTopProgressBar;
begin
  frmUserGUI.CardPanelGUITop.ActiveCard := frmUserGUI.CardGUITopProgressBar;
  frmUserGUI.GUITopProgressBar.Position := Parameter1;
  Debug('Setting User View top panel to progressbar and setting position "' + Parameter1 + '"');
end;

Procedure UserViewBottomText;
begin
  frmUserGUI.CardPanelGUIBottom.ActiveCard := frmUserGUI.CardGUIBottomText;
  frmUserGUI.LabelBottomText.Caption := Parameter1;
  frmUserGUI.LabelBottomText.Visible := true;
  Debug('Setting User View bottom panel to textmode and showing text "' + Parameter1 + '"');
end;

Procedure UserViewBottomProgressBar;
begin
  frmUserGUI.CardPanelGUIBottom.ActiveCard := frmUserGUI.CardGUIBottomProgressBar;
  frmUserGUI.GUIBottomProgressBar.Position := Parameter1;
  Debug('Setting User View bottom panel to progressbar and setting position "' + Parameter1 + '"');
end;

Procedure UserViewClientText;
begin
  frmUserGUI.CardPanelGUIClient.ActiveCard := frmUserGUI.CardGUIClientText;
  frmUserGUI.LabelClientText.Caption := Parameter1 + #10#13 + Parameter2 + #10#13 + Parameter3 + #10#13 + Parameter4 + #10#13 + Parameter5;
  Debug('Setting User View client panel to textmode and showing text "' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5 +  '"');
end;

Procedure UserViewTopClear;
begin
  frmUserGUI.CardPanelGUITop.ActiveCard := frmUserGUI.CardGUITopBlank;
  Debug('Clearing User View top area');
end;

Procedure UserViewBottomClear;
begin
  frmUserGUI.CardPanelGUIBottom.ActiveCard := frmUserGUI.CardGUIBottomBlank;
  Debug('Clearing User View bottom area');
end;

Procedure UserViewClientClear;
begin
  frmUserGUI.CardPanelGUIClient.ActiveCard := frmUserGUI.CardGUIClientBlank;
  Debug('Clearing User View client area');
end;

Procedure GotoTag;
begin
  CurrentCommandCounter := StrToInt(Parameter1);
  Debug('GotoTag - Jumping to row ' + Parameter1);
end;

Procedure IfVariable;
var
  Num1, Num2: Double;
  IsNumeric: Boolean;
  Result : Boolean;
begin
  Result := false;
  Debug('IfVariable - Comparing values');
  IsNumeric := (LowerCase(Parameter2) = 'less') or (LowerCase(Parameter2) = 'more');
  if IsNumeric then
  begin
    if not TryStrToFloat(Parameter1, Num1) or not TryStrToFloat(Parameter3, Num2) then
      begin
        Warning('Comparing strings with "Less" or "More" is not possible. Parameters must be numbers');
      end;
    if LowerCase(Parameter2) = 'less' then Result := Num1 < Num2 else Result := Num1 > Num2;
  end else if LowerCase(Parameter2) = 'equal' then
  begin
    if TryStrToFloat(Parameter1, Num1) and TryStrToFloat(Parameter3, Num2) then Result := Num1 = Num2
    else Result := LowerCase(Parameter1) = LowerCase(Parameter3);
  end;
  if Result then Debug('Result true and script will jump to row ' + Parameter4) else Debug('Result false and script will continue with next row');
  // NEED TO ADD JUMP ROUTINE
end;

end.
