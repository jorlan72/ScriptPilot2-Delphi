program Runner;

uses
  Vcl.Forms, Windows,
  main in 'main.pas' {frmRunner},
  Vcl.Themes,
  Vcl.Styles,
  intercode in '..\intercode.pas',
  mymessage in '..\mymessage.pas' {frmMessage};

{$R *.res}

begin
  Application.Initialize;
  if ParamStr(3) = 'HideRunner' then
    begin
      Application.ShowMainForm := false;
      Application.MainFormOnTaskbar := false;
    end else
    begin
      Application.ShowMainForm := true;
      Application.MainFormOnTaskbar := true;
    end;
  Application.Title := ParamStr(1);
  if ParamStr(1) = '' then Application.Title := 'Lost Runner' else Application.Title := ParamStr(1);
  TStyleManager.TrySetStyle('Windows10 BlackPearl');
  Application.CreateForm(TfrmRunner, frmRunner);
  if ParamStr(3) = 'HideRunner' then ShowWindow(Application.Handle, SW_HIDE);
  Application.Run;
end.
