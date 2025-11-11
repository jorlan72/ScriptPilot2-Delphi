
// SCRIPTPILOT 2
// CodeRed 2024 - Jorgen Lanesskog
// ScriptPilot enables visual presentation of automation routines and adds the capability for end-user interaction during script execution.
// "Developed over time with a generous pour of red wine"
//
// CODE FOR THE MAIN FORM AND APPLICATION LOGIC (MAIN.PAS)
// COMMAND DICTIONARY IN COMMANDS.PAS
// CODE SHARED BETWEEN EDITOR AND RUNNER IN INTERCODE.PAS
// USER GUI FOR SCRIPT PRESENTATION IN USERGUI.PAS
// RUNNER.PAS FOR COMMAND EXECUTION
// MYMESSAGE.PAS AND MYDIALOG.PAS FOR APPLICATION MESSAGES AND DIALOGS


unit main;

interface

// DELPHI BUILDT IN CLASSES
uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.WinXPanels, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Imaging.pngimage, Vcl.ComCtrls, Vcl.AppEvnts, System.Generics.Collections, Vcl.Menus, System.Hash,
  Vcl.Mask, System.JSON, System.IOUtils, Registry, System.UITypes, System.Actions, Vcl.ActnList, System.ImageList, Vcl.ImgList, Vcl.Buttons, ShellAPI,
  Commands, Runner, intercode, Vcl.Grids, Vcl.ValEdit, UserGUI, Vcl.Samples.Spin, Clipbrd, StrUtils, Vcl.WinXCtrls, MyMessage, MyDialog, System.Math,
  System.Diagnostics;

// MAIN FORM TYPES
type
  TfrmMain = class(TForm)
    PanelBottom: TPanel;
    PanelMenu: TPanel;
    PanelLog: TPanel;
    PanelScript: TPanel;
    PanelSettings: TPanel;
    PanelCodeRed: TPanel;
    TimerStartUP: TTimer;
    ApplicationEvents: TApplicationEvents;
    PopupMenuScriptEditor: TPopupMenu;
    MenuClone: TMenuItem;
    MenuDelete: TMenuItem;
    MenuDivider02: TMenuItem;
    ActionListShortcuts: TActionList;
    ActionActiavateSettings: TAction;
    ActionActivateLog: TAction;
    ActionActivateEditor: TAction;
    MenuDivider03: TMenuItem;
    MenuUndo: TMenuItem;
    ImageList48: TImageList;
    MenuRedo: TMenuItem;
    PopupMenuLog: TPopupMenu;
    MenuClearApplicationLog: TMenuItem;
    MenuClearOutputLog: TMenuItem;
    MenuDividerLog01: TMenuItem;
    MenuClearApplicationLoganddeletehistory: TMenuItem;
    MenuClearOutputLoganddeletehistory: TMenuItem;
    PopupMenuOpenFolders: TPopupMenu;
    MenuOpenScriptFolder: TMenuItem;
    MenuOpenDataFolder: TMenuItem;
    MenuOpenMediaFolder: TMenuItem;
    ActionOpenScriptFolder: TAction;
    ActionOpenDataFolder: TAction;
    ActionOpenMediaFolder: TAction;
    MenuDivider01: TMenuItem;
    CardPanelClient: TCardPanel;
    CardLog: TCard;
    SplitterLogs: TSplitter;
    GroupBoxMemoLog: TGroupBox;
    MemoLog: TMemo;
    GroupBoxOutputLog: TGroupBox;
    MemoOutputLog: TMemo;
    CardEditor: TCard;
    ListCommands: TListView;
    PanelScriptEdit: TPanel;
    PanelCommands: TPanel;
    GroupBoxCommandIndex: TGroupBox;
    PanelCommandIndexLeft: TPanel;
    ComboBoxCommand: TComboBox;
    ComboBoxP1: TComboBox;
    ComboBoxP2: TComboBox;
    ComboBoxP3: TComboBox;
    ComboBoxP4: TComboBox;
    ComboBoxP5: TComboBox;
    PanelAddButtons: TPanel;
    ButtonInsertCommand: TButton;
    ButtonUpdateCommand: TButton;
    ScrollBoxDescriptions: TScrollBox;
    LabelCommandLongDescription: TLabel;
    LabelCommandName: TLabel;
    LabelCommandShortDescription: TLabel;
    PanelScriptMenu: TPanel;
    GroupBoxFileOptions: TGroupBox;
    ButtonScriptNew: TButton;
    ButtonScriptOpen: TButton;
    ButtonScriptSave: TButton;
    ButtonScriptSaveAs: TButton;
    GroupBoxExecuteScript: TGroupBox;
    LabelScriptname: TLabel;
    ButtonScriptRun: TButton;
    CardSettings: TCard;
    ButtonSettingsSave: TButton;
    PanelSettings01: TPanel;
    GroupBoxApplicationOptions: TGroupBox;
    CheckBoxGUIFullScreen: TCheckBox;
    CheckBoxApplicationState: TCheckBox;
    CheckBoxPersistantApplicationLog: TCheckBox;
    CheckBoxPersistantOutputLog: TCheckBox;
    CheckBoxMenuMouseHover: TCheckBox;
    GroupBoxRunOptions: TGroupBox;
    LabelMaximumRunners: TLabel;
    CheckBoxStopExecutionOnWarnings: TCheckBox;
    CheckBoxResolveHostnames: TCheckBox;
    SpinEditMaxRunners: TSpinEdit;
    PanelSettings02: TPanel;
    GroupBoxFolderOptions: TGroupBox;
    LabeledEditScriptFolder: TLabeledEdit;
    ButtonBrowseScriptFolder: TButton;
    ButtonBrowseDatafolder: TButton;
    LabeledEditDataFolder: TLabeledEdit;
    ButtonBrowseMediaFolder: TButton;
    LabeledEditMediaFolder: TLabeledEdit;
    GroupBoxAlphablend: TGroupBox;
    CheckBoxMainFormAlphablend: TCheckBox;
    CheckBoxScriptRunnerAlphablend: TCheckBox;
    TrackBarMainForm: TTrackBar;
    TrackBarScriptRunner: TTrackBar;
    TrackBarUserGUI: TTrackBar;
    CheckBoxUserGUIAlphablend: TCheckBox;
    CardWelcome: TCard;
    ImageSplash: TImage;
    PanelLogo: TPanel;
    LabelScripPilot: TLabel;
    LabelVersion: TLabel;
    LabelCopyRight: TLabel;
    LabelDescriptionSPPurpose: TLabel;
    LabelDescriptionSPWine: TLabel;
    GroupBoxEncryptPasswords: TGroupBox;
    LabeledEditSeed: TLabeledEdit;
    LabeledEditPlainPassword: TLabeledEdit;
    LabeledEditEncryptedPassword: TLabeledEdit;
    ButtonEncrypt: TButton;
    ButtonCopy: TButton;
    MenuOpenRegistry: TMenuItem;
    CheckBoxDebugMode: TCheckBox;
    CheckBoxPauseAtEnd: TCheckBox;
    CheckBoxKillRunners: TCheckBox;
    PanelEditorTop: TPanel;
    CheckBoxFloatingHints: TCheckBox;
    PanelSettingsBottom: TPanel;
    CheckBoxLeaveScriptExpanded: TCheckBox;
    CheckBoxHideEditor: TCheckBox;
    ButtonMainAbort: TButton;
    CheckBoxScriptSanity: TCheckBox;
    TimerEditComboBox: TTimer;
    ButtonHint: TButton;
    CheckBoxLoopDetection: TCheckBox;
    MenuDividerLog02: TMenuItem;
    MenuLogCopy: TMenuItem;
    MenuLogSelectAll: TMenuItem;
    CardHelp: TCard;
    LabelLinkHome: TLabel;
    ActionActivateWelcome: TAction;

// MAIN FORM EVENT CALLS
    procedure PanelLogMouseEnter(Sender: TObject);
    procedure PanelScriptMouseEnter(Sender: TObject);
    procedure PanelSettingsMouseEnter(Sender: TObject);
    procedure PanelCodeRedMouseEnter(Sender: TObject);
    procedure TimerStartUPTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ApplicationEventsHint(Sender: TObject);
    procedure ComboBoxCommandChange(Sender: TObject);
    procedure ListCommandsDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure ListCommandsMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ListCommandsDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure ButtonInsertCommandClick(Sender: TObject);
    procedure PopupMenuScriptEditorPopup(Sender: TObject);
    procedure MenuCloneClick(Sender: TObject);
    procedure MenuDeleteClick(Sender: TObject);
    procedure ButtonBrowseScriptFolderClick(Sender: TObject);
    procedure ButtonBrowseDatafolderClick(Sender: TObject);
    procedure ButtonBrowseMediaFolderClick(Sender: TObject);
    procedure ComboBoxP1Enter(Sender: TObject);
    procedure ComboBoxP2Enter(Sender: TObject);
    procedure ComboBoxP3Enter(Sender: TObject);
    procedure ComboBoxP4Enter(Sender: TObject);
    procedure ComboBoxP5Enter(Sender: TObject);
    procedure ButtonScriptRunClick(Sender: TObject);
    procedure ButtonScriptNewClick(Sender: TObject);
    procedure ButtonScriptSaveClick(Sender: TObject);
    procedure ButtonScriptSaveAsClick(Sender: TObject);
    procedure ScrollBoxDescriptionsMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure ButtonSettingsSaveClick(Sender: TObject);
    procedure CheckBoxMainFormAlphablendClick(Sender: TObject);
    procedure CheckBoxScriptRunnerAlphablendClick(Sender: TObject);
    procedure TrackBarMainFormChange(Sender: TObject);
    procedure ButtonScriptOpenClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure ListCommandsClick(Sender: TObject);
    procedure MenuUndoClick(Sender: TObject);
    procedure MenuRedoClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure MenuClearApplicationLogClick(Sender: TObject);
    procedure MenuClearOutputLogClick(Sender: TObject);
    procedure MenuClearApplicationLoganddeletehistoryClick(Sender: TObject);
    procedure MenuClearOutputLoganddeletehistoryClick(Sender: TObject);
    procedure ComboBoxCommandEnter(Sender: TObject);
    procedure ButtonInsertCommandEnter(Sender: TObject);
    procedure ListCommandsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure ButtonUpdateCommandEnter(Sender: TObject);
    procedure MenuOpenScriptFolderClick(Sender: TObject);
    procedure MenuOpenDataFolderClick(Sender: TObject);
    procedure MenuOpenMediaFolderClick(Sender: TObject);
    procedure ButtonUpdateCommandClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CheckBoxUserGUIAlphablendClick(Sender: TObject);
    procedure CheckBoxMenuMouseHoverClick(Sender: TObject);
    procedure ButtonEncryptClick(Sender: TObject);
    procedure ButtonCopyClick(Sender: TObject);
    procedure MenuOpenRegistryClick(Sender: TObject);
    procedure CheckBoxDebugModeClick(Sender: TObject);
    procedure ScrollBoxDescriptionsResize(Sender: TObject);
    procedure ListCommandsDblClick(Sender: TObject);
    procedure CheckBoxFloatingHintsClick(Sender: TObject);
    procedure ButtonMainAbortClick(Sender: TObject);
    procedure TimerEditComboBoxTimer(Sender: TObject);
    procedure ButtonHintClick(Sender: TObject);
    procedure PopupMenuLogPopup(Sender: TObject);
    procedure MenuLogSelectAllClick(Sender: TObject);
    procedure MenuLogCopyClick(Sender: TObject);
    procedure LabelScripPilotClick(Sender: TObject);

  private
    // INCOMING WINDOWS MESSAGES FROM RUNNER RECEIVED VIA VMCOPYDATA
    procedure WMCopyData(var Msg: TWMCopyData); message WM_COPYDATA;
  public
  end;

// GLOBAL VARIABLES
type
  TListItemState = record // used to save and load the script listview state for undo function
  Caption: string;
  SubItems: TArray<string>;
  end;
  TListViewState = TArray<TListItemState>;

var
  frmMain: TfrmMain;
  StartTime : TdateTime;
  FolderScriptFiles : String;
  FolderDataFiles : String;
  FolderMediaFiles : String;
  ListCommandsInitialChecksum, ListCommandsCurrentChecksum: string; // used to check for changes in the script editor
  LoadedScriptFilenameNoPath : String;
  LoadedScriptFilenameFullPath : String;
  FileOpenSaveError : Boolean;
  ListViewState : TListViewState; // used to save and load the script listview state for undo function
  UndoStack, RedoStack : TArray<TListViewState>; // used for multiple level of undo and redo
  AreWeExitingTheApplication : Boolean; // flag for telling code that the application is about to terminate
  LastSelectedItem : TListItem; // The index of the selected line in ListCommands. It will contain the last index if multiple items selected
  UpdateFromCommandList : Boolean; // Indicating to UpdateCommandInfo that an item is clicked in the ListCommands
  DebugOn : Boolean; // Send additional logs to log memo
  pp1, pp2, pp3, pp4, pp5 : String; // Previous Paramenter Value
  LoopCounter : Integer; // Script Loop Detection
  LoopCommands : Integer;
  LoopTime : Integer;

// GLOBAL PROCEDURES AND FUNCTIONS (NOT EVENT BASED)
procedure LogItStamp(logtext : string; space : integer); // LOG TO MEMOLOG WITH TIMESTAMP
procedure LogIt(logtext : string; space : integer); // LOG TO MEMOLOG
procedure DebugItStamp(logtext : string; space : integer); // LOG TO DEBUGLOG WITH TIMESTAMP
procedure StartTimer; // USED TO CALCULATE TIME USAGE FOR CODE
function StopTimer: string; // USED IN COMBINATION WITH STARTTIMER
Function PathExists(const Path: string): Boolean; // CHECK IF A PATH EXISTS - RETURNS FALSE IF NOT - REQUIRES SYSTEM.IOUTILS
function GenerateListViewChecksum(ListView: TListView): string; // USED TO CHECK IF THERE HAS BEEN CHANGES TO THE SCRIPT - FOR SAVING CHANGES
procedure SaveScriptToFile(ListView: TListView; const FileName: string); // SAVE THE SCRIPT CONFIGURATION TO A JSON FILE AND WATERMARK IT
procedure LoadScriptFromFile(ListView: TListView; const FileName: string); // LOAD THE SCRIPT CONFIGURATION FROM A JSON FILE AND CHECK WATERMARK
procedure InsertScriptFile(ListView: TListView; const FileName: string; InsertIndex: Integer); // INSERT A SCRIPT CONFIGURATION FROM A JSON FILE INTO AN INDEX
procedure UpdateCommandInfo; // USED TO UPDATE COMMAND PARAMETERS IN COMMAND INDEX (IN EDITOR) + COMMAND DESCRIPTIONS
procedure RenumberListCommands; // USED TO RENUMBER THE ITEMS IN LISTCOMMANDS (IN EDITOR) WHEN ORDER OF ITEMS CHANGE
procedure HandleSaveOptions(Sender: TObject); // USED FOR OPEN, SAVE, SAVE AS AND NEW SCRIPT FILES - ONE PROCEDURE FOR HANDLING ALL SCRIPT EDITOR FILE OPERATIONS
procedure SaveListViewState(ListView: TListView; out State: TListViewState; PushOntoStack: Boolean = True); // SAVE COMMAND SCRIPT LISTVIEW STATE FOR UNDO FUNCTION
procedure LoadListViewState(ListView: TListView; const State: TListViewState); // LOAD COMMAND SCRIPT LISTVIEW STATE FOR UNDO FUNCTION
procedure PushListViewState(var Stack: TArray<TListViewState>; const State: TListViewState); // USED FOR MULTIPLE LEVEL UNDO
procedure PopListViewState(var Stack: TArray<TListViewState>; out State: TListViewState); // USED FOR MULTIPLE LEVEL UNDO
procedure PopulateComboBoxWithFiles(ComboBox: TComboBox; const DirectoryPath: string); // USED FOR ADDING FILES FOUND IN A DIRECTORY TO A DROP-DOWN MENU
Function ValidateCommandAndParameters : Boolean; // VALIDATE COMMAND AND ALL PARAMETERS
Procedure GetUserGUIReadyForRun; // CALLED FROM THE RUN BUTTON WHEN SCRIPT STARTS
procedure SearchListViewAndFillComboBox(const SearchWord: string; ListView: TListView; ComboBox: TComboBox); // USED TO FIND LISTVIEW ITEMS AND ADD TO COMBONOXES
procedure VariableStringReplace(var Str: variant); // REPLACES PARAMETERS WITH VALUES FOUND IN THE VARIABLE DATA STRUCTURE
procedure SetShowHintsForFormComponents(AForm: TForm; ShowHints: Boolean); // ENABLE OR DISABLE FLOATING HINTS
procedure ShowMyMessage(Title, Msg: String; Modal: Boolean; Index: Integer); // ALTERNATIVE TO SHOWMESSAGE
function ShowMyDialog(Title: String; Index: Integer): TModalResult; // ALTERNATIVE TO SHOWMESSAGEDLG
procedure ScanScriptVarAndTags; // SCAN EDITOR SCRIPT FOR VARIABLES AND TAGS FOR IN EDIT VALIDATION
Procedure ChangeParamTypesBeforeRun; // CHANGE SOME COMMANDS PARAMETER TYPE DURING SCRIPT RUN
Procedure ChangeParamTypesAfterRun; // CHANGE SOME COMMANDS PARAMETER TYPE BACK TO INITIAL VALUES WHEN SCRIPT FINSIH

implementation

{$R *.dfm}

procedure TFrmMain.WMCopyData(var Msg: TWMCopyData);
// INCOMING WINDOWS MESSAGES FROM RUNNER RECEIVED VIA VMCOPYDATA
// RECEIVEDSTRING CONTAINS STRINGDATA AND STRINGCODE (MESSAGEID)
// STRINGCODE USED TO IDENTIFY WHAT KIND OF STRING IS BEING RECEIVED
// CODE 0 = ECHO
// CODE 1 = COMMAND
// CODE 2 = NORMAL LOG
// CODE 3 = DEBUG LOG
// CODE 4 = RUNNER READY
// CODE 5 = VARIABLE TABLE
// CODE 6 = OUTPUT LOG
// CODE 7 = ERROR
// CODE 8 = WARNING
var
  ReceivedString: string;
  MessageId: Cardinal;
  ParameterText : String;
  TagCounter : Integer;
begin
  MessageId := Msg.CopyDataStruct^.dwData;
  SetLength(ReceivedString, Msg.CopyDataStruct^.cbData div SizeOf(Char));
  Move(Msg.CopyDataStruct^.lpData^, PChar(ReceivedString)^, Msg.CopyDataStruct^.cbData);
  ReceivedString := Trim(ReceivedString);
    case MessageId of
      0: begin // ECHO
           EchoReply := ReceivedString;
           DebugItStamp('Echo reply from ' + ReceivedString, 0);
         end;
      1: begin // COMMAND
           if Pos('IncludeScript', ReceivedString) > 0 then
             begin
               ParameterText := Copy(ReceivedString, Pos(',', ReceivedString) + 1, Length(ReceivedString) - 1);
               DebugItStamp('Inserting script file "' + ParameterText + '" into script at position ' + IntToStr(CurrentCommandCounter), 0);
               InsertScriptFile(frmMain.ListCommands, FolderScriptFiles + ParameterText, CurrentCommandCounter);
               if not FileOpenSaveError then
                 begin
                   RenumberListCommands;
                   DebugItStamp('Scanning for script tags', 0);
                   for TagCounter := 0 to ListCommands.Items.Count - 1 do
                   begin
                     if lowercase(frmMain.ListCommands.Items[TagCounter].SubItems[0]) = 'settag' then
                       begin
                         Variables.AddOrUpdate(frmMain.ListCommands.Items[TagCounter].SubItems[1], IntToStr(TagCounter));
                         DebugItStamp('Updating tag ' + frmMain.ListCommands.Items[TagCounter].SubItems[1] + ' at row ' + IntToStr(TagCounter), 0);
                       end;
                   end;
                 end else
                 begin
                   LogIt('', 2);
                   LogItStamp('********************************  SCRIPT ERROR OCCURED!  ********************************', 1);
                   LogItStamp('Error inserting script file into current script from command at row ' + IntToStr(CurrentCommandCounter) + '.', 0);
                   LogItStamp('Failed command : ' + CommandToExecute + ' ' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5, 0);
                   LogItStamp('File format not supported or file could not be accessed',0);
                   LogItStamp('Stopping script execution', 1);
                   LogItStamp('****************************************************************************************', 2);
                   ShowViewMessage('Error inserting script file into current script from command at row ' + IntToStr(CurrentCommandCounter) + '.', 'Failed command : ' + CommandToExecute + ' ' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5 + #10#13 + #10#13 + 'File format not supported or file could not be accessed', True,0);
                   CommandToExecute := '';
                   DidScriptComplete := false;
                 end;
             end;
         end;
      2: begin // LOG
           LogItStamp(ReceivedString, 0);
         end;
      3: begin // DEBUG LOG
           DebugItStamp(ReceivedString, 0);
         end;
      4: begin // RUNNER READY
           AreRunnerReady := true;
         end;
      7: begin // ERROR - WILL STOP SCRIPT IN ANY CASE
           CommandToExecute := '';
           DidScriptComplete := false;
           LogIt('', 2);
           LogItStamp('********************************  FATAL ERROR OCCURED!  ********************************', 1);
           LogItStamp(ReceivedString, 0);
           LogItStamp('Stopping script execution', 1);
           LogItStamp('****************************************************************************************', 2);
           ShowMyMessage('Fatal Error Occured', ReceivedString + #10#13 + 'Stopping script execution', true, 0);
         end;
      8: begin // WARNING - WILL STOP SCRIPT IF STOP ON WARNING ENABLED
           if frmMain.CheckBoxStopExecutionOnWarnings.Checked then CommandToExecute := '';
           WarningHappened := true;
           LogIt('', 2);
           LogItStamp('********************************  WARNING!  ********************************', 1);
           LogItStamp(ReceivedString, 0);
           if frmMain.CheckBoxStopExecutionOnWarnings.Checked then LogItStamp('Stopping script execution when "Stop Script On Warnings" checked', 1);
           LogItStamp('****************************************************************************************', 2);
           if frmMain.CheckBoxStopExecutionOnWarnings.Checked then ShowMyMessage('WARNING', ReceivedString + #10#13 + 'Stopping script execution when "Stop Script On Warnings" checked', true, 0);
           if frmMain.CheckBoxStopExecutionOnWarnings.Checked then DidScriptComplete := false;
         end;
    end;
end;

procedure PopulateComboBoxWithFiles(ComboBox: TComboBox; const DirectoryPath: string);
// USED FOR ADDING FILES FOUND IN A DIRECTORY TO A DROP-DOWN MENU
var
  SearchRec: TSearchRec;
  FindResult: Integer;
begin
  ComboBox.Items.BeginUpdate;
  try
    ComboBox.Items.Clear;
    FindResult := FindFirst(DirectoryPath + '\*.*', faAnyFile, searchRec);
    while FindResult = 0 do
    begin
      if (searchRec.Attr and faDirectory) = 0 then comboBox.Items.Add(searchRec.Name);
      findResult := FindNext(searchRec);
    end;
    FindClose(searchRec);
  finally
    comboBox.Items.EndUpdate;
  end;
end;

procedure PushListViewState(var Stack: TArray<TListViewState>; const State: TListViewState);
// USED FOR MULTIPLE LEVEL UNDO
begin
  SetLength(Stack, Length(Stack) + 1);
  Stack[High(Stack)] := State;
end;

procedure PopListViewState(var Stack: TArray<TListViewState>; out State: TListViewState);
// USED FOR MULTIPLE LEVEL UNDO
begin
  if Length(Stack) > 0 then
  begin
    State := Stack[High(Stack)];
    SetLength(Stack, Length(Stack) - 1);
  end;
end;

procedure SaveListViewState(ListView: TListView; out State: TListViewState; PushOntoStack: Boolean = True);
// SAVE COMMAND SCRIPT LISTVIEW STATE FOR UNDO FUNCTION
var
  i, j: Integer;
  ItemState: TListItemState;
begin
  SetLength(State, ListView.Items.Count);
  for i := 0 to ListView.Items.Count - 1 do
    begin
      ItemState.Caption := ListView.Items[i].Caption;
      SetLength(ItemState.SubItems, ListView.Items[i].SubItems.Count);
      for j := 0 to ListView.Items[i].SubItems.Count - 1 do ItemState.SubItems[j] := ListView.Items[i].SubItems[j];
      State[i] := ItemState;
    end;
  if PushOntoStack then
    begin
      PushListViewState(UndoStack, State);
      frmMain.MenuUndo.Enabled := Length(UndoStack) > 0;
    end;
end;

procedure LoadListViewState(ListView: TListView; const State: TListViewState);
// LOAD COMMAND SCRIPT LISTVIEW STATE FOR UNDO FUNCTION
var
  i, j: Integer;
  Item: TListItem;
begin
  ListView.Items.BeginUpdate;
  try
    ListView.Items.Clear;
    for i := 0 to High(State) do
    begin
      Item := ListView.Items.Add;
      Item.Caption := State[i].Caption;
      for j := 0 to High(State[i].SubItems) do Item.SubItems.Add(State[i].SubItems[j]);
    end;
  finally
    ListView.Items.EndUpdate;
    frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
  end;
end;

Function PathExists(const Path: string): Boolean;
// CHECK IF A PATH EXISTS - RETURNS FALSE IF NOT - REQUIRES SYSTEM.IOUTILS
begin
  Result := TFile.Exists(Path) or TDirectory.Exists(Path);
end;

procedure TfrmMain.ApplicationEventsHint(Sender: TObject);
// FIRES ON APPLICATION HINTS AND PUTS THE OBJECTS HINT TEXT TO THE STATUSPANEL
begin
  frmMain.PanelBottom.Caption := ' ' + Application.Hint;
end;

procedure TfrmMain.CheckBoxDebugModeClick(Sender: TObject);
// TURN ON OR OFF DEBUG MODE DURING SCRIPT RUN
begin
  DebugOn := frmMain.CheckBoxDebugMode.Checked;
end;

procedure TfrmMain.CheckBoxFloatingHintsClick(Sender: TObject);
// TURN ON OR OFF FLOATING HINTS TO EVERY OBJECT
begin
  SetShowHintsForFormComponents(frmMain, frmMain.CheckBoxFloatingHints.Checked);
end;

procedure TfrmMain.CheckBoxMainFormAlphablendClick(Sender: TObject);
// ENABLE ALPHABLEND ON THE MAIN FORM OR NOT
begin
  frmMain.AlphaBlend := frmMain.CheckBoxMainFormAlphablend.Checked;
  frmMain.TrackBarMainForm.Enabled := frmMain.CheckBoxMainFormAlphablend.Checked;
end;

procedure TfrmMain.CheckBoxMenuMouseHoverClick(Sender: TObject);
// ENABLE OR DISABLE MENU ACTIVATION ON MOUSE HOVER ON THE MAIN MENU
begin
  if not frmMain.CheckBoxMenuMouseHover.Checked then
  begin
    frmMain.PanelScript.OnMouseEnter := nil;
    frmMain.PanelLog.OnMouseEnter := nil;
    frmMain.PanelSettings.OnMouseEnter := nil;
    frmMain.PanelCodeRed.OnMouseEnter := nil;
  end else
  begin
    frmMain.PanelScript.OnMouseEnter := frmMain.PanelScriptMouseEnter;
    frmMain.PanelLog.OnMouseEnter := frmMain.PanelLogMouseEnter;
    frmMain.PanelSettings.OnMouseEnter := frmMain.PanelSettingsMouseEnter;
    frmMain.PanelCodeRed.OnMouseEnter := frmMain.PanelCodeRedMouseEnter;
  end;
end;

procedure TfrmMain.CheckBoxScriptRunnerAlphablendClick(Sender: TObject);
// ENABLE ALPHABLEND ON THE SCRIPT RUNNER FORM OR NOT
begin
  frmMain.TrackBarScriptRunner.Enabled := frmMain.CheckBoxScriptRunnerAlphablend.Checked;
end;

procedure TfrmMain.CheckBoxUserGUIAlphablendClick(Sender: TObject);
// ENABLE ALPHABLEND ON THE USER GUI FORM OR NOT
begin
  frmUserGUI.AlphaBlend := frmMain.CheckBoxUserGUIAlphablend.Checked;
  frmMain.TrackBarUserGUI.Enabled := frmMain.CheckBoxUserGUIAlphablend.Checked;
end;

procedure TfrmMain.ComboBoxCommandChange(Sender: TObject);
// CHECKS IF WRITTEN OR SELECTED COMMAND EXISTS AND THEN UPDATES PARAMETER FIELDS AND COMMAND DESCRIPTIONS FROM COMMAND MAP
begin
  UpdateCommandInfo;
end;

procedure TfrmMain.ComboBoxCommandEnter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ComboBoxCommand.Hint;
end;

procedure TfrmMain.ComboBoxP1Enter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ComboBoxP1.Hint;
  if frmMain.ComboBoxP1.Items.Count >= 1 then frmMain.ComboBoxP1.DroppedDown := true;
  frmMain.TimerEditComboBox.Enabled := true;
end;

procedure TfrmMain.ComboBoxP2Enter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ComboBoxP2.Hint;
  if frmMain.ComboBoxP2.Items.Count >= 1 then frmMain.ComboBoxP2.DroppedDown := true;
end;

procedure TfrmMain.ComboBoxP3Enter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ComboBoxP3.Hint;
  if frmMain.ComboBoxP3.Items.Count >= 1 then frmMain.ComboBoxP3.DroppedDown := true;
end;

procedure TfrmMain.ComboBoxP4Enter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ComboBoxP4.Hint;
  if frmMain.ComboBoxP4.Items.Count >= 1 then frmMain.ComboBoxP4.DroppedDown := true;
end;

procedure TfrmMain.ComboBoxP5Enter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ComboBoxP5.Hint;
  if frmMain.ComboBoxP5.Items.Count >= 1 then frmMain.ComboBoxP5.DroppedDown := true;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
// APPLICATION TERMINATION PROCEDURE - SAVE SETTINGS
begin
  // save application main form size, position and state - we do this in any case, even if the option us unchecked in settings
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowState', IntToStr(Integer(frmMain.WindowState)));
  // only save form position and size if the form is <> maximized or minimized
  if frmMain.WindowState = wsNormal then
    begin
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowTop', IntToStr(Integer(frmMain.Top)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowLeft', IntToStr(Integer(frmMain.Left)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowWidth', IntToStr(Integer(frmMain.Width)));
      WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowHeight', IntToStr(Integer(frmMain.Height)));
    end;
// save application log if checked
  if frmMain.CheckBoxPersistantApplicationLog.Checked then
    begin
      LogItStamp('Application shutdown',0);
      frmMain.MenuClearApplicationLog.Tag := 1;
      frmMain.MenuClearApplicationLog.Click;
    end;
// save output log if checked
  if frmMain.CheckBoxPersistantOutputLog.Checked then
    begin
      frmMain.MenuClearOutputLog.Tag := 1;
      frmMain.MenuClearOutputLog.Click;
    end;
end;

procedure TfrmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
// IF THERE ARE ANY CHANGES TO THE SCRIPT, ASK TO SAVE BEFORE APPLICATION TERMINATE
var
  UserResponse : Integer;
begin
  // do not quit if script is running
  if ScriptRunning = True then
    begin
      ShowMyMessage('Script is currently running!', 'Abort script first or wait for the script to end', False, 0);
      CanClose := False;
      exit;
    end;
  // check if there are unsaved changes to the script
  AreWeExitingTheApplication := true; // tell code that the application is about to exit
  if (GenerateListViewChecksum(ListCommands) <> ListCommandsInitialChecksum) and (LoadedScriptFilenameFullPath <> '') then
    begin
      UserResponse := ShowMyDialog('Do you want to save the script changes?', 0);
      if UserResponse = mrYes then
        begin
        HandleSaveOptions(frmMain.ButtonScriptNew)
        end else
      if UserResponse = mrNo then
        begin
          CanClose := true;
          exit;
        end;
    end else HandleSaveOptions(frmMain.ButtonScriptNew);
  if GenerateListViewChecksum(ListCommands) <> ListCommandsInitialChecksum then CanClose := False else CanClose := True;
  if not CanClose then AreWeExitingTheApplication := false; // if user cancel termination, then reset to false
end;

function GetRunnerPath: string;
// FIND THE RUNNER.EXE TO SEE IF IT IS IN THE SCRIPTPILOT FOLDER OR WE ARE DOING A DELPHI SESSION
var
  AppPath, BasePath: string;
begin
  AppPath := ExtractFilePath(Application.ExeName);
  BasePath := ExcludeTrailingPathDelimiter(AppPath);
  BasePath := ExtractFilePath(ExcludeTrailingPathDelimiter(ExtractFilePath(BasePath)));
  Result := IncludeTrailingPathDelimiter(BasePath) + 'Runner\Win64\Release\Runner.exe';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
// CODE EXECUTED DURING CREATION OF MAIN FORM - BUILD COMMAND MAP - SETS INITIAL PARAMETERS
var
  SingleCommandInMap : TPair<string, TCommandInfo>;
  SingleParameterInMap : String;
  DecryptedSeed : String;
begin
// starting a timer to calculate application startup time
  StartTimer;
// load application log history if present and checked
  frmMain.CheckBoxPersistantApplicationLog.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveApplicationLog', 'false'));
  if frmMain.CheckBoxPersistantApplicationLog.Checked then if FileExists(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log') then frmMain.MemoLog.Lines.LoadFromFile(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log');
// load output log history if present and checked
  frmMain.CheckBoxPersistantOutputLog.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveOutputLog', 'false'));
  if frmMain.CheckBoxPersistantOutputLog.Checked then if FileExists(ExtractFilePath(Application.ExeName) + 'OutputLog.log') then frmMain.MemoOutputLog.Lines.LoadFromFile(ExtractFilePath(Application.ExeName) + 'OutputLog.log');
// setting the welcome card to the startup card
  frmMain.CardPanelClient.ActiveCard := frmMain.CardWelcome;
// adding startup messages to the application log
  LogIt('',0);
  LogIt('Welcome to ScriptPilot 2',0);
  LogIt('Build : ' + GetApplicationVersion,1);
// adding all commands from command.pas to the dictionary
  AddAllCommandsToDictionary;
// parsing each command in the command map and adding them to the command combobox in the command index section
  for SingleCommandInMap in CommandMap do
    begin
      SingleParameterInMap := SingleCommandInMap.Value.ParametersLong[0];
      frmMain.ComboBoxCommand.Items.Add(SingleParameterInMap);
    end;
// setting initial gui settings
  frmMain.LabelVersion.Caption := 'Version : ' + GetApplicationVersion;
// check if runner.exe i present - checks normal position and the the position when running from delphi IDE
  FileRunnerExe := '';
  if fileexists(extractfilepath(application.ExeName + 'runner.exe')) then FileRunnerExe := extractfilepath(application.ExeName + 'runner.exe')
  else if fileexists(GetRunnerPath) then FileRunnerExe := GetRunnerPath;
  if FileRunnerExe = '' then
    begin
      LogIt('Runner.exe NOT found. A very limited set of commands will work. Any commands accessing runners will make the script fail', 0);
      LogIt('Please get a copy of the runner.exe and place it in the same folder as scriptpilot2.exe', 1);
    end;
// read registry settings and set settings page
  frmMain.CheckBoxResolveHostnames.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ResolveHostnamesInScript', 'true'));
  ResolveHostnamesInScript := frmMain.CheckBoxResolveHostnames.Checked;
  frmMain.CheckBoxGUIFullScreen.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunUserGUIFullScreen', 'false'));
  frmMain.CheckBoxMenuMouseHover.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'MenuMouseHover', 'true'));
  frmMain.CheckBoxDebugMode.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'DebugMode', 'false'));
  frmMain.CheckBoxPauseAtEnd.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'PauseScriptAtEnd', 'false'));
  frmMain.CheckBoxKillRunners.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'KillRunners', 'true'));
  frmMain.CheckBoxFloatingHints.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FloatingHints', 'true'));
  frmMain.CheckBoxLeaveScriptExpanded.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'LeaveScriptExpanded', 'false'));
  frmMain.CheckBoxStopExecutionOnWarnings.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'StopExecutionOnWarnings', 'true'));
  frmMain.CheckBoxScriptSanity.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'CheckScriptSanity', 'true'));
  frmMain.CheckBoxHideEditor.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'HideEditor', 'false'));
  frmMain.CheckBoxLoopDetection.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'LoopDetection', 'true'));
  frmMain.SpinEditMaxRunners.Value := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'MaximumRunners', '10'));
  LoopCommands := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'LoopCommands', '25'));
  LoopTime := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'LoopTime', '1000'));
  SetShowHintsForFormComponents(frmMain, frmMain.CheckBoxFloatingHints.Checked);
  if not frmMain.CheckBoxMenuMouseHover.Checked then
    begin
      frmMain.PanelScript.OnMouseEnter := nil;
      frmMain.PanelLog.OnMouseEnter := nil;
      frmMain.PanelSettings.OnMouseEnter := nil;
      frmMain.PanelCodeRed.OnMouseEnter := nil;
    end else
    begin
      frmMain.PanelScript.OnMouseEnter := frmMain.PanelScriptMouseEnter;
      frmMain.PanelLog.OnMouseEnter := frmMain.PanelLogMouseEnter;
      frmMain.PanelSettings.OnMouseEnter := frmMain.PanelSettingsMouseEnter;
      frmMain.PanelCodeRed.OnMouseEnter := frmMain.PanelCodeRedMouseEnter;
    end;
// load encryption settings
  DecryptedSeed := ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'CurrentEncryptionSeed', '');
  if DecryptedSeed <> '' then DecryptedSeed := DecryptString(DecryptedSeed, InternalSeed);
  frmMain.LabeledEditSeed.Text := DecryptedSeed;
// load main form state and attributes
  frmMain.CheckBoxApplicationState.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveFormPositionAndState', 'true'));
// set folders for scripts, data and media
  FolderScriptFiles := ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FolderScriptFiles', ExtractFilePath(Application.ExeName));
  FolderDataFiles := ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FolderDataFiles', ExtractFilePath(Application.ExeName));
  FolderMediaFiles := ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FolderMediaFiles', ExtractFilePath(Application.ExeName));
  frmMain.LabeledEditScriptFolder.Text := FolderScriptFiles;
  frmMain.LabeledEditDataFolder.Text := FolderDataFiles;
  frmMain.LabeledEditMediaFolder.Text := FolderMediaFiles;
// start with a blank script
  LoadedScriptFilenameNoPath := '';
  LoadedScriptFilenameFullPath := '';
  frmMain.LabelScriptName.Caption := 'Untitled';
  ScriptRunning := false;
// set the initial checksum of the script editor
  ListCommandsInitialChecksum := GenerateListViewChecksum(frmMain.ListCommands);
// initialize undo stack variables
  SetLength(ListViewState, 0);
  SetLength(UndoStack, 0);
  SetLength(RedoStack, 0);
  EditVariables := TMyValueList.Create;
  EditTags := TMyValueList.Create;
// "TimerStartUPTimer" will fire after this and be disabled - contains more initialization code - alphablend fade will also not work from oncreate
  frmMain.TimerStartUP.Enabled := true;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
// DYNAMIC MEMORY CLEANUP
begin
  SetLength(ListViewState, 0);
  SetLength(UndoStack, 0);
  SetLength(RedoStack, 0);
  CommandMap.Free;
  EditVariables.Free;
  EditTags.Free;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  if frmMain.CheckBoxApplicationState.Checked then
    begin
      frmMain.Top := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowTop', '100'));
      frmMain.Left := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowLeft', '320'));
      frmMain.Width := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowWidth', '1244'));
      frmMain.Height := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowHeight', '761'));
      if StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ApplicationWindowState', '0')) <> 2 then frmMain.WindowState := wsNormal else frmMain.WindowState := wsMaximized;
    end;
end;

procedure TfrmMain.LabelScripPilotClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar('www.scriptpilot.no'), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.ListCommandsClick(Sender: TObject);
// UPDATE EDIT COMBOBOXES WITH THE CLICKED COMMAND INFO FOR EDITING
begin
  if (GetKeyState(VK_MENU) < 0) then
    begin
      frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
      frmMain.ButtonUpdateCommand.Enabled := frmMain.ListCommands.SelCount = 1;
      frmMain.ButtonInsertCommand.Enabled := frmMain.ListCommands.SelCount < 2;
      frmMain.GroupBoxCommandIndex.Enabled := frmMain.ListCommands.SelCount < 2;
      exit;
    end;
  LastSelectedItem := ListCommands.Selected;
  if Assigned(LastSelectedItem) then
    begin
      ComboBoxCommand.Text := LastSelectedItem.SubItems[0];
      UpdateFromCommandList := true;
      UpdateCommandInfo;
      pp1 := frmMain.ComboBoxP1.Text;
      pp2 := frmMain.ComboBoxP2.Text;
      pp3 := frmMain.ComboBoxP3.Text;
      pp4 := frmMain.ComboBoxP4.Text;
      pp5 := frmMain.ComboBoxP5.Text;
    end else
    begin
      frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
      frmMain.ButtonUpdateCommand.Enabled := frmMain.ListCommands.SelCount = 1;
      frmMain.ButtonInsertCommand.Enabled := frmMain.ListCommands.SelCount < 2;
      frmMain.GroupBoxCommandIndex.Enabled := frmMain.ListCommands.SelCount < 2;
    end;
  if frmMain.ListCommands.SelCount = 1 then
    begin
      frmMain.ButtonInsertCommand.Caption := '&Insert';
      frmMain.ButtonInsertCommand.Hint := 'Insert Command|Insert above any selected row (use ALT + click to select without edits changing)';
    end else
    begin
      frmMain.ButtonInsertCommand.Caption := '&Add';
      frmMain.ButtonInsertCommand.Hint := 'Add Command|Add command to the end of the script';
    end;
end;

procedure TfrmMain.ListCommandsDblClick(Sender: TObject);
begin
  // Adjust sizes based on the scale factor
  if frmMain.PanelEditorTop.Height = Round(20 * ScaleFactor) then frmMain.PanelEditorTop.Height := Round(350 * ScaleFactor)
  else frmMain.PanelEditorTop.Height := Round(20 * ScaleFactor);
end;

procedure TfrmMain.ListCommandsDragDrop(Sender, Source: TObject; X, Y: Integer);
// SCRIPT EDITOR DRAG AND DROP - INSERT DROPPED ITEMS AND REARANGE LIST
var
  TargetItem, NewItem: TListItem;
  TargetIndex, I, J, FirstSelectedIndex: Integer;
  IsDraggingDown: Boolean;
  ItemData: array of record Caption: string; SubItems: TStringList; end;
begin
  SaveListViewState(frmMain.ListCommands, ListViewState, true);
  if (Source = ListCommands) and (ListCommands.SelCount > 0) then
    begin
      TargetItem := ListCommands.GetItemAt(X, Y);
      FirstSelectedIndex := ListCommands.Selected.Index;
      IsDraggingDown := Assigned(TargetItem) and (TargetItem.Index > FirstSelectedIndex);
      if Assigned(TargetItem) then TargetIndex := TargetItem.Index else TargetIndex := ListCommands.Items.Count;
      SetLength(ItemData, ListCommands.SelCount);
      J := 0;
      for I := 0 to ListCommands.Items.Count - 1 do
        begin
          if ListCommands.Items[I].Selected then
            begin
              ItemData[J].Caption := ListCommands.Items[I].Caption;
              ItemData[J].SubItems := TStringList.Create;
              ItemData[J].SubItems.Assign(ListCommands.Items[I].SubItems);
              Inc(J);
            end;
        end;
      for I := ListCommands.Items.Count - 1 downto 0 do if ListCommands.Items[I].Selected then ListCommands.Items[I].Delete;
      if IsDraggingDown then Dec(TargetIndex, Length(ItemData));
      for I := 0 to High(ItemData) do
        begin
          NewItem := ListCommands.Items.Insert(TargetIndex + I);
          NewItem.Caption := ItemData[I].Caption;
          NewItem.SubItems.Assign(ItemData[I].SubItems);
          ItemData[I].SubItems.Free;
        end;
    end;
  RenumberListCommands;
  if frmMain.ListCommands.SelCount = 1 then
    begin
      frmMain.ButtonInsertCommand.Caption := '&Insert';
      frmMain.ButtonInsertCommand.Hint := 'Insert Command|Insert above any selected row (use ALT + click to select without edits changing)';
    end else
    begin
      frmMain.ButtonInsertCommand.Caption := '&Add';
      frmMain.ButtonInsertCommand.Hint := 'Add Command|Add command to the end of the script';
    end;
end;

procedure TfrmMain.ListCommandsDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
// SCRIPT EDITOR DRAG AND DROP - ACCEPT LISTVIEW ITEMS
var
  ScrollHeight: Integer;
begin
  Accept := Source = frmMain.ListCommands;
  ScrollHeight := 20;
  if Y < ScrollHeight then frmMain.ListCommands.Scroll(0, -ScrollHeight) else if Y > frmMain.ListCommands.ClientHeight - ScrollHeight then frmMain.ListCommands.Scroll(0, ScrollHeight);
end;

procedure TfrmMain.ListCommandsMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
// SCRIPT EDITOR DRAG AND DROP - START THE OPERATION WHEN ITEM CLICKED
var
  Item: TListItem;
begin
  if Button = mbLeft then
    begin
      Item := ListCommands.GetItemAt(X, Y);
      if Assigned(Item) then ListCommands.BeginDrag(False);
    end;
end;

procedure TfrmMain.ListCommandsSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  if Selected then frmMain.ListCommandsClick(Sender);
end;

procedure TfrmMain.MenuClearApplicationLoganddeletehistoryClick(Sender: TObject);
// CLEAR THE SESSION LOG AND DELETE APPLICATION.LOG, IF IT IS PRESENT
begin
  frmMain.MemoLog.Clear;
  frmMain.MenuClearApplicationLog.Tag := 0;
  if FileExists(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log') then if DeleteFile(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log') then LogItStamp('Application Log cleared and history deleted',0);
end;

procedure TfrmMain.MenuClearOutputLoganddeletehistoryClick(Sender: TObject);
// CLEAR THE OUTPUT LOG AND DELETE OUTPUT.LOG, IF IT IS PRESENT
begin
  frmMain.MemoOutputLog.Clear;
  frmMain.MenuClearOutputLog.Tag := 1;
  if FileExists(ExtractFilePath(Application.ExeName) + 'OutputLog.log') then if DeleteFile(ExtractFilePath(Application.ExeName) + 'OutputLog.log') then LogItStamp('Output Log cleared and history deleted',0);
end;

procedure TfrmMain.MenuClearApplicationLogClick(Sender: TObject);
// CLEAR THE APPLICATION LOG FOR THIS SESSION
var
  FileContents: TStringList;
begin
// save application log if checked
  if frmMain.CheckBoxPersistantApplicationLog.Checked then
    begin
      if frmMain.MenuClearApplicationLog.Tag = 1 then
        begin
          if FileExists(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log') then
            begin
              FileContents := TStringList.Create;
              try
                FileContents.LoadFromFile(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log');
                frmMain.MemoLog.Text := FileContents.Text + frmMain.MemoLog.Text;
              finally
                FileContents.Free;
              end;
            end;
        end;
      frmMain.MemoLog.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + 'ApplicationLog.log');
    end;
  frmMain.MemoLog.Clear;
  // indicate that the log has been cleared at least once
  frmMain.MenuClearApplicationLog.Tag := 1;
end;

procedure TfrmMain.MenuClearOutputLogClick(Sender: TObject);
// CLEAR THE OUTPUT LOG FOR THIS SESSION
var
  FileContents: TStringList;
begin
// save application log if checked
  if frmMain.CheckBoxPersistantOutputLog.Checked then
    begin
      if frmMain.MenuClearOutputLog.Tag = 1 then
        begin
          if FileExists(ExtractFilePath(Application.ExeName) + 'OutputLog.log') then
            begin
              FileContents := TStringList.Create;
              try
                FileContents.LoadFromFile(ExtractFilePath(Application.ExeName) + 'OutputLog.log');
                frmMain.MemoOutputLog.Text := FileContents.Text + frmMain.MemoOutputLog.Text;
              finally
                FileContents.Free;
              end;
            end;
        end;
      frmMain.MemoOutputLog.Lines.SaveToFile(ExtractFilePath(Application.ExeName) + 'OutputLog.log');
    end;
  frmMain.MemoOutputLog.Clear;
  // indicate that the log has been cleared at least once
  frmMain.MenuClearOutputLog.Tag := 1;
end;

procedure TfrmMain.MenuCloneClick(Sender: TObject);
// CLONE ONE OR MORE SELECTED ITEMS IN THE SCRIPT EDITOR LISTVIEW
// ALL CLONED ITEMS WILL BE PLACED BELOW THE "LOWEST" SELECTED ITEM
var
  LastSelectedIndex, i, j: Integer;
  NewItem, CurrentItem: TListItem;
begin
  SaveListViewState(frmMain.ListCommands, ListViewState, true);
  if ListCommands.SelCount = 0 then Exit;
  LastSelectedIndex := -1;
  for i := 0 to ListCommands.Items.Count - 1 do if ListCommands.Items[i].Selected then LastSelectedIndex := i;
  if LastSelectedIndex = -1 then Exit;
  for i := 0 to ListCommands.Items.Count - 1 do
    begin
      if ListCommands.Items[i].Selected then
        begin
          CurrentItem := ListCommands.Items[i];
          NewItem := ListCommands.Items.Insert(LastSelectedIndex + 1);
          NewItem.Caption := CurrentItem.Caption;
          for j := 0 to CurrentItem.SubItems.Count - 1 do NewItem.SubItems.Add(CurrentItem.SubItems[j]);
          Inc(LastSelectedIndex);
        end;
    end;
  RenumberListCommands;
  UpdateCommandInfo;
end;

procedure TfrmMain.MenuDeleteClick(Sender: TObject);
// DELETE ONE OR MORE SELECTED ITEMS IN THE SCRIPT EDITOR LISTVIEW
var
  i, NextIndex: Integer;
begin
  SaveListViewState(frmMain.ListCommands, ListViewState, true);
  NextIndex := ListCommands.Selected.Index;
  for i := ListCommands.Items.Count - 1 downto 0 do if ListCommands.Items[i].Selected then
    begin
     ListCommands.Items[i].Delete;
    end;
  RenumberListCommands;
  UpdateCommandInfo;
  if NextIndex >= ListCommands.Items.Count then NextIndex := ListCommands.Items.Count - 1;
  if (ListCommands.Items.Count > 0) and (NextIndex >= 0) then
    begin
      ListCommands.Items[NextIndex].Selected := True;
      ListCommands.Items[NextIndex].Focused := True;
    end
end;

procedure TfrmMain.MenuLogCopyClick(Sender: TObject);
var
  Memo: TMemo;
  PopupMenu: TPopupMenu;
begin
  if Sender is TMenuItem then
    begin
      PopupMenu := (Sender as TMenuItem).GetParentMenu as TPopupMenu;
      if Assigned(PopupMenu) and (PopupMenu.PopupComponent is TMemo) then
        begin
          Memo := TMemo(PopupMenu.PopupComponent);
          Memo.CopyToClipboard;
        end;
    end;
end;

procedure TfrmMain.MenuLogSelectAllClick(Sender: TObject);
var
  Memo: TMemo;
  PopupMenu: TPopupMenu;
begin
  if Sender is TMenuItem then
    begin
      PopupMenu := (Sender as TMenuItem).GetParentMenu as TPopupMenu;
      if Assigned(PopupMenu) and (PopupMenu.PopupComponent is TMemo) then
        begin
          Memo := TMemo(PopupMenu.PopupComponent);
          Memo.SelectAll;
        end;
    end;
end;


procedure TfrmMain.MenuOpenDataFolderClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar(FolderDataFiles), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.MenuOpenMediaFolderClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar(FolderMediaFiles), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.MenuOpenRegistryClick(Sender: TObject);
begin
  OpenRegeditToKey('HKEY_CURRENT_USER\SOFTWARE\CodeRed\ScriptPilot');
end;

procedure TfrmMain.MenuOpenScriptFolderClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar(FolderScriptFiles), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.MenuRedoClick(Sender: TObject);
// REDO FUNCTION FOR LAST EXECUTED ACTION ON SCRIPT EDITOR
var
  CurrentState, RedoState: TListViewState;
begin
  if Length(RedoStack) > 0 then
  begin
    SaveListViewState(frmMain.ListCommands, CurrentState, False); // Do not push onto UndoStack
    PopListViewState(RedoStack, RedoState);
    LoadListViewState(frmMain.ListCommands, RedoState);
    PushListViewState(UndoStack, CurrentState); // Push the current state onto the Undo stack
  end;
  frmMain.MenuUndo.Enabled := Length(UndoStack) > 0;
  frmMain.MenuRedo.Enabled := Length(RedoStack) > 0;
end;

procedure TfrmMain.MenuUndoClick(Sender: TObject);
// UNDO FUNCTION FOR LAST EXECUTED ACTION ON SCRIPT EDITOR
var
  CurrentState, UndoState: TListViewState;
begin
  if Length(UndoStack) > 0 then
  begin
    SaveListViewState(frmMain.ListCommands, CurrentState, False); // Do not push onto UndoStack
    PopListViewState(UndoStack, UndoState);
    LoadListViewState(frmMain.ListCommands, UndoState);
    PushListViewState(RedoStack, CurrentState); // Push the current state onto the Redo stack
  end;
  frmMain.MenuUndo.Enabled := Length(UndoStack) > 0;
  frmMain.MenuRedo.Enabled := Length(RedoStack) > 0;
end;

procedure TfrmMain.PanelCodeRedMouseEnter(Sender: TObject);
// HIGHLIGHTS THE WELCOME MENU AND SHOWS THE WELCOME CARD
begin
  frmMain.PanelCodeRed.ParentBackground := false;
  frmMain.PanelScript.ParentBackground := true;
  frmMain.PanelLog.ParentBackground := true;
  frmMain.PanelSettings.ParentBackground := true;
  frmMain.PanelCodeRed.Font.Size := 14;
  frmMain.PanelLog.Font.Size := 10;
  frmMain.PanelScript.Font.Size := 10;
  frmMain.PanelSettings.Font.Size := 10;
  frmMain.CardPanelClient.ActiveCard := frmMain.CardWelcome;
end;

procedure TfrmMain.PanelLogMouseEnter(Sender: TObject);
// HIGHLIGHTS THE LOG MENU AND SHOWS THE LOG CARD
begin
  frmMain.PanelCodeRed.ParentBackground := true;
  frmMain.PanelScript.ParentBackground := true;
  frmMain.PanelLog.ParentBackground := false;
  frmMain.PanelSettings.ParentBackground := true;
  frmMain.PanelCodeRed.Font.Size := 10;
  frmMain.PanelLog.Font.Size := 14;
  frmMain.PanelScript.Font.Size := 10;
  frmMain.PanelSettings.Font.Size := 10;
  frmMain.CardPanelClient.ActiveCard := frmMain.CardLog;
end;

procedure TfrmMain.PanelScriptMouseEnter(Sender: TObject);
// HIGHLIGHTS THE EDITOR MENU AND SHOWS THE EDITOR CARD
begin
  frmMain.PanelCodeRed.ParentBackground := true;
  frmMain.PanelScript.ParentBackground := false;
  frmMain.PanelLog.ParentBackground := true;
  frmMain.PanelSettings.ParentBackground := true;
  frmMain.PanelCodeRed.Font.Size := 10;
  frmMain.PanelLog.Font.Size := 10;
  frmMain.PanelScript.Font.Size := 14;
  frmMain.PanelSettings.Font.Size := 10;
  frmMain.CardPanelClient.ActiveCard := frmMain.CardEditor;
end;

procedure TfrmMain.PanelSettingsMouseEnter(Sender: TObject);
// HIGHLIGHTS THE SETTINGS MENU AND SHOWS THE SETTINGS CARD
begin
  frmMain.PanelCodeRed.ParentBackground := true;
  frmMain.PanelScript.ParentBackground := true;
  frmMain.PanelLog.ParentBackground := true;
  frmMain.PanelSettings.ParentBackground := false;
  frmMain.PanelCodeRed.Font.Size := 10;
  frmMain.PanelLog.Font.Size := 10;
  frmMain.PanelScript.Font.Size := 10;
  frmMain.PanelSettings.Font.Size := 14;
  frmMain.CardPanelClient.ActiveCard := frmMain.CardSettings;
end;

procedure TfrmMain.PopupMenuLogPopup(Sender: TObject);
var
  Memo: TMemo;
  PopupMenu: TPopupMenu;
begin
  if Sender is TPopupMenu then
  begin
    PopupMenu := TPopupMenu(Sender);
    if Assigned(PopupMenu) and (PopupMenu.PopupComponent is TMemo) then
    begin
      Memo := TMemo(PopupMenu.PopupComponent);
      frmMain.MenuLogCopy.Enabled := Memo.SelLength > 0;
    end;
  end;
end;

procedure TfrmMain.PopupMenuScriptEditorPopup(Sender: TObject);
// SCRIPT EDITOR POP-UP MENU FOR COMMAND TOOLS LIKE CLONE, EDIT, DELETE OF COMMANDS IN THE SCRIPT
begin
 frmMain.MenuClone.Enabled := frmMain.ListCommands.SelCount > 0;
 frmMain.MenuDelete.Enabled := frmMain.ListCommands.SelCount > 0;
 frmMain.MenuUndo.Enabled := Length(UndoStack) > 0;
 frmMain.MenuRedo.Enabled := Length(RedoStack) > 0;
end;

procedure TfrmMain.ScrollBoxDescriptionsMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
// MAKE SURE THE MOUSE WHEEL WORKS IN THE SCROLLBOX FOR COMMAND DESCRIPTIONS WHEN TEXT IS MORE THAN SCREEN SIZE
begin
  if WheelDelta > 0 then ScrollBoxDescriptions.VertScrollBar.Position := ScrollBoxDescriptions.VertScrollBar.Position - 20
  else ScrollBoxDescriptions.VertScrollBar.Position := ScrollBoxDescriptions.VertScrollBar.Position + 20;
  Handled := True;
end;

procedure TfrmMain.ScrollBoxDescriptionsResize(Sender: TObject);
// MAKE SURE THE TEXT FITS IN THE SCOLLBOX
begin
  frmMain.LabelCommandName.Width := frmMain.ScrollBoxDescriptions.ClientWidth - 20;
  frmMain.LabelCommandShortDescription.Width := frmMain.ScrollBoxDescriptions.ClientWidth - 20;
  frmMain.LabelCommandLongDescription.Width := frmMain.ScrollBoxDescriptions.ClientWidth - 20;
end;

Procedure LogItStamp(logtext : string; space : integer);
// LOG TO MEMOLOG WITH TIMESTAMP
var
  Counter : Integer;
begin
  frmMain.MemoLog.Lines.Add(datetimetostr(now) + ' : ' + logtext);
  if space > 0 then
    begin
      for Counter := 1 to space do begin
        frmMain.MemoLog.Lines.Add('');
      end;
    end;
  frmMain.MemoLog.SelStart := Length(frmMain.MemoLog.Text);
  SendMessage(frmMain.MemoLog.Handle, WM_VSCROLL, SB_BOTTOM, 0);
end;

Procedure LogIt(logtext : string; space : integer);
// LOG TO MEMOLOG
var
  Counter : Integer;
begin
  frmMain.MemoLog.Lines.Add(logtext);
  if space > 0 then
    begin
      for Counter := 1 to space do begin
        frmMain.MemoLog.Lines.Add('');
      end;
    end;
  frmMain.MemoLog.SelStart := Length(frmMain.MemoLog.Text);
  SendMessage(frmMain.MemoLog.Handle, WM_VSCROLL, SB_BOTTOM, 0);
end;

Procedure DebugItStamp(logtext : string; space : integer);
// LOG TO MEMOLOG WITH TIMESTAMP
var
  Counter : Integer;
begin
  if not frmMain.CheckBoxDebugMode.Checked then exit;
  frmMain.MemoLog.Lines.Add(datetimetostr(now) + ' : ' + logtext);
  if space > 0 then
    begin
      for Counter := 1 to space do begin
        frmMain.MemoLog.Lines.Add('');
      end;
    end;
  frmMain.MemoLog.SelStart := Length(frmMain.MemoLog.Text);
  SendMessage(frmMain.MemoLog.Handle, WM_VSCROLL, SB_BOTTOM, 0);
end;

Procedure StartTimer;
// USED TO CALCULATE TIME USAGE FOR CODE
begin
  StartTime := Now;
end;

Function StopTimer: string;
// USED IN COMBINATION WITH STARTTIMER
var
  ElapsedTime: TDateTime;
begin
  ElapsedTime := Now - StartTime;
  Result := FormatDateTime('hh:nn:ss.zzz', ElapsedTime);
end;

procedure TfrmMain.TimerEditComboBoxTimer(Sender: TObject);
// DELAYED CODE TO WORKAROUND THE AUTOSELECT OF COMBOBOX ITEMS - WANT TO SELECT TEXT BETWEEN HASHTAGS
begin
  frmMain.TimerEditComboBox.Enabled := false;
  if (frmMain.ComboBoxP1.Text = '') and (frmMain.ComboBoxCommand.Text = 'SetTag') then
    begin
      frmMain.ComboBoxP1.Text := '#YourTagNameHere#';
      frmMain.ComboBoxP1.SelStart := 1;
      frmMain.ComboBoxP1.SelLength := Length(frmMain.ComboBoxP1.Text) - 2;
      exit;
    end;
  if (frmMain.ComboBoxP1.Text = '') and (frmMain.ComboBoxCommand.Text = 'SetVariable') then
    begin
      frmMain.ComboBoxP1.Text := '#YourVariableNameHere#';
      frmMain.ComboBoxP1.SelStart := 1;
      frmMain.ComboBoxP1.SelLength := Length(frmMain.ComboBoxP1.Text) - 2;
      exit;
    end;
end;

procedure TfrmMain.TimerStartUPTimer(Sender: TObject);
// CODE THAT FIRES RIGHT AFTER MAIN FORM CREATION TO FADE IN APP WITH ALPHABLEND AND SET OTHER RUNTIME PARAMETERS
var
  alphablend : Integer;
begin
  frmMain.TimerStartUP.Enabled := false;
// fade in main form
  alphablend := 1;
    While alphablend < 255 do
      begin
        frmMain.AlphaBlendValue := alphablend;
        application.ProcessMessages;
        sleep(1);
        inc(alphablend, 5);
        if alphablend > 255 then break;
      end;
// set persistant alphablend values, if set
  frmMain.CheckBoxMainFormAlphablend.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendMainForm', 'true'));
  if not frmMain.CheckBoxMainFormAlphablend.Checked then frmMain.AlphaBlend := false else
    begin
      frmMain.AlphaBlend := true;
      frmMain.TrackBarMainForm.Position := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendMainFormLevel', '255'));
      frmMain.AlphaBlendValue := frmMain.TrackBarMainForm.Position;
    end;
  frmMain.CheckBoxScriptRunnerAlphablend.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendScriptRunner', 'false'));
  frmMain.TrackBarScriptRunner.Position := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendScriptRunnerLevel', '255'));
  frmMain.CheckBoxUserGUIAlphablend.Checked := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendUserGUI', 'false'));
  frmMain.TrackBarUserGUI.Position := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendUserGUILevel', '255'));
// stopping time tracking and notifying that we are ready for work
  Logit('Application startup time : ' +  StopTimer,1);
  LogItStamp('Ready.',1);
end;

procedure TfrmMain.TrackBarMainFormChange(Sender: TObject);
// CHANGE ALPHABLEND VALUE ON MAIN FORM ACCORDING TO TRACK BAR
begin
  frmMain.AlphaBlendValue := frmMain.TrackBarMainForm.Position;
end;

procedure RenumberListCommands;
// USED TO RENUMBER THE ITEMS IN LISTCOMMANDS (IN EDITOR) WHEN ORDER OF ITEMS CHANGE
var
  I: Integer;
begin
  for I := 0 to frmMain.ListCommands.Items.Count - 1 do frmMain.ListCommands.Items[I].Caption := IntToStr(I);
  frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
  frmMain.ButtonUpdateCommand.Enabled := frmMain.ListCommands.SelCount = 1;
  frmMain.ButtonInsertCommand.Enabled := frmMain.ListCommands.SelCount < 2;
  frmMain.GroupBoxCommandIndex.Enabled := frmMain.ListCommands.SelCount < 2;
end;

procedure TfrmMain.ButtonBrowseDatafolderClick(Sender: TObject);
// SET THE FOLDER FOR SEARCHING FOR FILES
begin
with TFileOpenDialog.Create(nil) do
  try
    Options := [fdoPickFolders];
    if Execute then frmMain.LabeledEditDataFolder.Text := FileName + '\';
  finally
    Free;
  end;
end;

procedure TfrmMain.ButtonBrowseMediaFolderClick(Sender: TObject);
// SET THE FOLDER FOR SEARCHING FOR FILES
begin
with TFileOpenDialog.Create(nil) do
  try
    Options := [fdoPickFolders];
    if Execute then frmMain.LabeledEditMediaFolder.Text := FileName + '\';
  finally
    Free;
  end;
end;

procedure TfrmMain.ButtonBrowseScriptFolderClick(Sender: TObject);
// SET THE FOLDER FOR SEARCHING FOR FILES
begin
with TFileOpenDialog.Create(nil) do
  try
    Options := [fdoPickFolders];
    if Execute then frmMain.LabeledEditScriptFolder.Text := FileName + '\';
  finally
    Free;
  end;
end;

procedure TfrmMain.ButtonCopyClick(Sender: TObject);
// COPY ENCRYPTED TEXT TO THE CLIPBOARD FROM THE SETTINGS PAGE
begin
  Clipboard.AsText := frmMain.LabeledEditEncryptedPassword.Text;
  frmMain.LabeledEditEncryptedPassword.Clear;
  frmMain.LabeledEditPlainPassword.Clear;
end;

procedure TfrmMain.ButtonEncryptClick(Sender: TObject);
// ENCRYPT THE PLAIN TEXT WITH BLOWFISH
var
 TextToEncrypt : String;
 EncryptedSeed : String;
begin
  TextToEncrypt := EncryptString(frmMain.LabeledEditPlainPassword.Text, frmMain.LabeledEditSeed.Text);
  frmMain.LabeledEditEncryptedPassword.Text := TextToEncrypt;
  EncryptedSeed := EncryptString(frmMain.LabeledEditSeed.Text, InternalSeed);
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'CurrentEncryptionSeed', EncryptedSeed);
end;

procedure TfrmMain.ButtonHintClick(Sender: TObject);
var
  CurrentCommandInfo: TCommandInfo;
  ComboboxNumber : Integer;
  ComboBox : TComboBox;
  RandomIndex: Integer;
begin
  if CommandMap.TryGetValue(LowerCase(frmMain.ComboBoxCommand.Text), CurrentCommandInfo) then
    begin
      for ComboboxNumber := 1 to 5 do
        begin
          ComboBox := frmMain.FindComponent('ComboBoxP' + IntToStr(ComboBoxNumber)) as TComboBox;
          with ComboBox do
          if ComboBox.Enabled then
            begin
              if (ComboBox.Items.Count >= 2) and not (CurrentCommandInfo.Types[ComboBoxNumber] in [ptText, ptRequiredText, ptRange]) then
                begin
                  RandomIndex := RandomRange(0, ComboBox.Items.Count);
                  ComboBox.ItemIndex := RandomIndex;
                end
              else ComboBox.Text := CurrentCommandInfo.Example[ComboBoxNumber];
            end;
        end;
    end else
      begin
        ShowMyMessage('Command selected for parameter hints not valid', 'Please select a command and try again', true, 0);
      end;
end;

procedure TfrmMain.ButtonScriptNewClick(Sender: TObject);
// CREATE A NEW SCRIPT
begin
  HandleSaveOptions(sender);
  UpdateCommandInfo;
end;

procedure TfrmMain.ButtonScriptSaveAsClick(Sender: TObject);
// SAVE CURRENT SCRIPT AS A COPY WITH A NEW NAME AND SET IT ACTIVE
begin
  HandleSaveOptions(sender);
end;

procedure TfrmMain.ButtonScriptOpenClick(Sender: TObject);
// OPEN A SCRIPTFILE
begin
  HandleSaveOptions(sender);
  UpdateCommandInfo;
end;


// ***************************************************************************************************************************
// *************************   MAIN RUN ENGINE START   ***********************************************************************
// Long procedure, but have not found any good reason to refactor it.. yet..
// ***************************************************************************************************************************

procedure TfrmMain.ButtonScriptRunClick(Sender: TObject);
// RUN THE SCRIPT IN THE EDITOR (LISTCOMMANDS)
var
  CurrentCommandInfo: TCommandInfo;
  RunnerKey: string;
  ErrorInfo : String;
  BeforeRunState: TListViewState;
  ListItem: TListItem;
  TagCounter : Integer;
  ExcludedCommands: TStringList;
  NotVariableReplaced : TStringList;
  ValidationOK : Boolean;
  Stopwatch: TStopwatch;

begin
  if frmMain.ListCommands.Items.Count = 0 then
    begin
      ShowMyMessage('Will not run an empty script!', 'Please enter some commands first.', true, 0);
      exit;
    end;

  // commands not to be executing or reported string replaced during script loop
  ExcludedCommands := TStringList.Create;
  ExcludedCommands.Add('settag');
  ExcludedCommands.Add('setremark');
  NotVariableReplaced := TStringList.Create;
  NotVariableReplaced.Add('setvariable');
  // start the timer and go
  StartTimer;
  frmMain.PanelEditorTop.Height := Round(20 * ScaleFactor);
  frmMain.ListCommands.Selected := nil;
  LogIt('', 0);
  LogItStamp('*** RUN SCRIPT START ***',0);
  DebugItStamp('Starting script environment', 0);
  LogItStamp('Executing script "' + frmMain.LabelScriptname.Caption + '"', 0);
  // initialize script environment variables and structures
  SaveListViewState(frmMain.ListCommands, BeforeRunState, false);
  frmMain.CardEditor.Enabled := false;
  frmMain.CardSettings.Enabled := false;
  frmMain.CardWelcome.Enabled := false;
  frmMain.PanelBottom.Caption := 'ScriptPilot 2 - CodeRed 2024 - J. Lanesskog';
  MaxNumberOfRunners := frmMain.SpinEditMaxRunners.Value;
  CurrentNumberOfRunners := 0;
  CurrentRunner := '';
  CurrentCommandCounter := 0;
  LoopCounter := 0;
  CommandToExecute := 'Start';
  DidScriptComplete := True;
  WarningHappened := false;
  ScriptRunning := True;
  Variables := TMyValueList.Create;
  Runners := TMyDataList.Create;
  Connections := TMyDataList.Create;
  // get usergui ready
  DebugItStamp('Starting User View', 0);
  GetUserGUIReadyForRun;
  frmMain.ButtonMainAbort.Visible := true;
  frmMain.ButtonMainAbort.Cancel := true;
  // get all tags
  DebugItStamp('Scanning for script tags', 0);
  for TagCounter := 0 to ListCommands.Items.Count -1 do
    begin
      if lowercase(frmMain.ListCommands.Items[TagCounter].SubItems[0]) = 'settag' then
        begin
          Variables.AddOrUpdate(frmMain.ListCommands.Items[TagCounter].SubItems[1], IntToStr(TagCounter));
          DebugItStamp('Found tag with name ' + frmMain.ListCommands.Items[TagCounter].SubItems[1] + ' at row ' + IntToStr(TagCounter), 0);
        end;
    end;
  // change some parameter types before script run - ifvariable
  ChangeParamTypesBeforeRun;

  // *** SCRIPT LOOP START *****************************************************
  Stopwatch := TStopwatch.StartNew;
  DebugItStamp('Reading script from row 0', 0);
  while CommandToExecute <> '' do
    begin
      // highlight the current command in the editor
      frmMain.ListCommands.Items[CurrentCommandCounter].Selected := True;
      // read the next command from the editor
      ListItem := frmMain.ListCommands.Items[CurrentCommandCounter];
      // scroll to needed row if the row is not visible
      ListItem.MakeVisible(false);
      // read the current row into command variables
      CommandToExecute := lowercase(frmMain.ListCommands.Items[CurrentCommandCounter].SubItems[0]);
      Parameter1 := frmMain.ListCommands.Items[CurrentCommandCounter].SubItems[1];
      Parameter2 := frmMain.ListCommands.Items[CurrentCommandCounter].SubItems[2];
      Parameter3 := frmMain.ListCommands.Items[CurrentCommandCounter].SubItems[3];
      Parameter4 := frmMain.ListCommands.Items[CurrentCommandCounter].SubItems[4];
      Parameter5 := frmMain.ListCommands.Items[CurrentCommandCounter].SubItems[5];
      // find the command in the command map
      if ExcludedCommands.IndexOf(CommandToExecute) = -1 then
        begin
          if CommandMap.TryGetValue(CommandToExecute, CurrentCommandInfo) then
            begin
              // send command to debug if activ - before stringreplace
              if NotVariableReplaced.IndexOf(CommandToExecute) = -1 then DebugItStamp(Trim('ORIGINAL : ' + CommandToExecute + ' ' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5) , 0);
              // stringreplace parameters to the variable data
              if (CurrentCommandInfo.Types[1] <> ptNotUsed) and (CurrentCommandInfo.StringReplacement[1]) then VariableStringReplace(Parameter1);
              if (CurrentCommandInfo.Types[2] <> ptNotUsed) and (CurrentCommandInfo.StringReplacement[2]) then VariableStringReplace(Parameter2);
              if (CurrentCommandInfo.Types[3] <> ptNotUsed) and (CurrentCommandInfo.StringReplacement[3]) then VariableStringReplace(Parameter3);
              if (CurrentCommandInfo.Types[4] <> ptNotUsed) and (CurrentCommandInfo.StringReplacement[4]) then VariableStringReplace(Parameter4);
              if (CurrentCommandInfo.Types[5] <> ptNotUsed) and (CurrentCommandInfo.StringReplacement[5]) then VariableStringReplace(Parameter5);
              // send command to debug if activ - after stringreplace
              DebugItStamp(Trim('EXECUTING : ' + CommandToExecute + ' ' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5) , 0);
              // validating parameters before executing command
              ValidationOK := true;
              if ValidationOK and (CurrentCommandInfo.Types[1] <> ptNotUsed) then if not ValidateCommand(CommandToExecute, Parameter1, 1, ErrorInfo) then ValidationOK := false;
              if ValidationOK and (CurrentCommandInfo.Types[2] <> ptNotUsed) then if not ValidateCommand(CommandToExecute, Parameter2, 2, ErrorInfo) then ValidationOK := false;
              if ValidationOK and (CurrentCommandInfo.Types[3] <> ptNotUsed) then if not ValidateCommand(CommandToExecute, Parameter3, 3, ErrorInfo) then ValidationOK := false;
              if ValidationOK and (CurrentCommandInfo.Types[4] <> ptNotUsed) then if not ValidateCommand(CommandToExecute, Parameter4, 4, ErrorInfo) then ValidationOK := false;
              if ValidationOK and (CurrentCommandInfo.Types[5] <> ptNotUsed) then if not ValidateCommand(CommandToExecute, Parameter5, 5, ErrorInfo) then ValidationOK := false;
              if not ValidationOK then
                begin
                     LogIt('', 2);
                     LogItStamp('********************************  SCRIPT ERROR OCCURED!  ********************************', 1);
                     LogItStamp('Parameter contains incorrect type of data at row ' + IntToStr(CurrentCommandCounter) + '.', 0);
                     LogItStamp('Failed command : ' + CommandToExecute + ' ' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5, 0);
                     LogItStamp(ErrorInfo,0);
                     LogItStamp('Stopping script execution', 1);
                     LogItStamp('****************************************************************************************', 2);
                     ShowViewMessage('Parameter contains incorrect type of data at row ' + IntToStr(CurrentCommandCounter), 'Script will stop execution because of failed command : ' + #10#13 + #10#13 + CommandToExecute + ' ' + Parameter1 + ' ' + Parameter2 + ' ' + Parameter3 + ' ' + Parameter4 + ' ' + Parameter5 + #10#13 + #10#13 + ErrorInfo, True, 0);
                     CommandToExecute := '';
                     DidScriptComplete := false;
                end else
                begin
                  // execute the current command
                  if CommandToExecute <> '' then CurrentCommandInfo.Proc();
                end;
            end;
        end;
      // remove highlight from the command executed
      ListCommands.Selected := nil;
      // if error or warning made the CommandToExecute = '' then do not increase counter else increase the counter to go to next command
      if CommandToExecute <> '' then Inc(CurrentCommandCounter);
      // if we have reached the end of the script, do not try to read a new command, but set it to ''. This will end the loop.
      if CurrentCommandCounter = ListCommands.Items.Count then CommandToExecute := '';
      // if abort button clicked then end the loop
      if AbortButtonClicked then CommandToExecute := '';
      // script loop detection
      Application.ProcessMessages;
      Inc(LoopCounter);
      if (frmMain.CheckBoxLoopDetection.Checked) and (LoopCounter > LoopCommands) and (Stopwatch.ElapsedMilliseconds < LoopTime) then
        begin
          LogIt('',0);
          LogItStamp('*** Script loop detected ***', 0);
          DebugItStamp(IntToStr(LoopCommands) + ' commands executed in ' + FloatToStr(Stopwatch.ElapsedMilliseconds) + ' milliseconds', 0);
          DebugItStamp('Threshold set to ' + IntToStr(LoopCommands) + ' commands within ' + IntToStr(LoopTime) + ' milliseconds indicates a loop', 0);
          DebugItStamp('These values can be set in the Registry. Add string values "LoopCommands" and "LoopTime"', 1);
          var Userresponse : Integer;
          UserResponse := ShowViewDialog('Possible script loop detected. Abort script?', 0);
          if UserResponse = mrYes then
            begin
              DebugItStamp('Script loop detected and user decided to abort execution', 0);
              AbortButtonClicked := true;
              CommandToExecute := '';
            end else
            begin
              DebugItStamp('Script loop detected but user approved it', 0);
              Inc(LoopCommands, 10);
              LoopCounter := 0;
              Stopwatch.Reset;
              Stopwatch.Start;
            end;
        end;
      if (frmMain.CheckBoxLoopDetection.Checked = false) and (LoopCounter > LoopCommands) and (Stopwatch.ElapsedMilliseconds < LoopTime) then
        begin
          LogIt('',0);
          LogItStamp('*** Script loop detected ***', 0);
          DebugItStamp(IntToStr(LoopCommands) + ' commands executed in ' + FloatToStr(Stopwatch.ElapsedMilliseconds) + ' milliseconds', 0);
          DebugItStamp('Threshold set to ' + IntToStr(LoopCommands) + ' commands within ' + IntToStr(LoopTime) + ' milliseconds indicates a loop', 0);
          DebugItStamp('These values can be set in the Registry. Add string values "LoopCommands" and "LoopTime"', 1);
          DebugItStamp('Loop Detection handling deactivated - Script will continue',1);
          LoopCounter := 0;
          Stopwatch.Reset;
          Stopwatch.Start;
        end;
    end;
  // *** SCRIPT RUN LOOP END *****************************************************

  // if pause script at end
  if not ABortButtonClicked then
    begin
      If (frmMain.CheckBoxPauseAtEnd.Checked) and (DidScriptComplete) then // if all went well
        begin
          LogitStamp('Last line in script executed', 0);
          LogItStamp('Pausing before doing script environment cleanup',0);
          ShowViewMessage('Last line in script executed','Pausing before doing script environment cleanup.' + #10#13 +  'You can turn this dialog on or off at the settings page' + #10#13 + #10#13 + 'Click OK to proceed', true, 0);
        end else If (frmMain.CheckBoxPauseAtEnd.Checked) then // if something went wrong
        begin
          LogitStamp('Script execution stopped due to an error', 0);
          LogItStamp('Pausing before doing script environment cleanup',0);
          ShowViewMessage('Script execution stopped due to an error','Pausing before doing script environment cleanup' + #10#13 +  'You can turn this dialog on or off at the settings page' + #10#13 +  #10#13 + 'Click OK to proceed', True, 0);
        end;
    end;
  // script end cleanup
  if not frmMain.CheckBoxLeaveScriptExpanded.Checked then LoadListViewState(frmMain.ListCommands, BeforeRunState);
  DebugitStamp('CLEANUP START : Cleaning up script environment', 0);
  frmUserGUI.BorderStyle := bsSizeable;
  frmUserGUI.WindowState := wsNormal;
  DebugitStamp('Writing User View position to Registry', 0);
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowTop', IntToStr(Integer(frmUserGUI.Top)));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowLeft', IntToStr(Integer(frmUserGUI.Left)));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowWidth', IntToStr(Integer(frmUserGUI.Width)));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowHeight', IntToStr(Integer(frmUserGUI.Height)));
  if frmMain.CheckBoxKillRunners.Checked then
    begin
      DebugitStamp('Killing runners', 0);
      for RunnerKey in Runners.GetAllKeys do
        begin
          SendWinMessage(RunnerKey, 'I kill you now..', 90);
        end;
    end;
  DebugitStamp('Closing User View', 0);
  // frmUserGUI will not close if script is running!
  ScriptRunning := false;
  frmUserGUI.Close;
  frmMain.ButtonMainAbort.Visible := false;
  frmMain.ButtonMainAbort.Cancel := false;
  if frmMain.CheckBoxHideEditor.Checked then
    begin
    frmMain.Visible := true;
    frmMain.Show;
    frmMain.SetFocus;
    end;
  DebugitStamp('Clearing memory tables', 0);
  Variables.Free;
  Runners.Free;
  NotVariableReplaced.Free;
  ExcludedCommands.Free;
  DebugitStamp('Clearing open connections', 0);
  Connections.Free;
  // change some parameter types after script finish - ifvariable
  ChangeParamTypesAfterRun;
  DebugitStamp('CLEANUP COMPLETE: Cleaning up script environment', 0);
  LogitStamp('Script ended at row : ' + IntToStr(CurrentCommandCounter), 0);
  LogItStamp('Script completed using a total time of ' + StopTimer, 0);
  if not AbortButtonClicked then
    begin
      if DidScriptComplete then
        begin
          if WarningHappened then LogItStamp('*** RUN SCRIPT COMPLETED WITH WARNINGS ***', 1) else LogItStamp('*** RUN SCRIPT COMPLETE ***', 1);
        end else
        begin
          if WarningHappened then LogItStamp('*** RUN SCRIPT STOPPED ON WARNING ***', 1) else LogItStamp('*** RUN SCRIPT FAILED ***', 1);
        end;
    end else
    begin
      LogItStamp('*** RUN SCRIPT ABORTED ***', 1);
    end;
  OKButtonPressed := false;
  AbortButtonClicked := false;
  frmMain.PanelEditorTop.Height := Round(350 * ScaleFactor);
  frmMain.CardEditor.Enabled := true;
  frmMain.CardSettings.Enabled := true;
  frmMain.CardWelcome.Enabled := true;
  frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
  frmMain.ButtonUpdateCommand.Enabled := frmMain.ListCommands.SelCount = 1;
  frmMain.ButtonInsertCommand.Enabled := frmMain.ListCommands.SelCount < 2;
  frmMain.GroupBoxCommandIndex.Enabled := frmMain.ListCommands.SelCount < 2;
end;

// ***************************************************************************************************************************
// *************************   MAIN RUN ENGINE END   *************************************************************************
// ***************************************************************************************************************************


procedure TfrmMain.ButtonScriptSaveClick(Sender: TObject);
// SAVE CURRENT SCRIPT AND SET IT ACTIVE
begin
  HandleSaveOptions(sender);
end;

procedure TfrmMain.ButtonSettingsSaveClick(Sender: TObject);
// SAVE SETTINGS FROM THE SETTINGS PAGE TO REGISTRY
var
  TempFilename : String;
begin
  if (not PathExists(frmMain.LabeledEditScriptFolder.Text)) and (frmMain.LabeledEditScriptFolder.Text <> '') then
    begin
      ShowMyMessage('Please correct the following:', 'The Script folder specified does not exist. Please specify a different folder', true, 0);
      frmMain.LabeledEditScriptFolder.SetFocus;
      exit;
    end else if frmMain.LabeledEditScriptFolder.Text = '' then
      begin
        FolderScriptFiles := ExtractFilePath(Application.ExeName);
      end else
        begin
          TempFilename := frmMain.LabeledEditScriptFolder.Text;
          if not TempFilename.EndsWith('\') then TempFilename := TempFilename + '\';
          frmMain.LabeledEditScriptFolder.Text := TempFilename;
          FolderScriptFiles := frmMain.LabeledEditScriptFolder.Text;
        end;
  if (not PathExists(frmMain.LabeledEditDataFolder.Text)) and (frmMain.LabeledEditDataFolder.Text <> '') then
    begin
      ShowMyMessage('Please correct the following:', 'The Data folder specified does not exist. Please specify a different folder', true, 0);
      frmMain.LabeledEditDataFolder.SetFocus;
      exit;
    end else if frmMain.LabeledEditDataFolder.Text = '' then
      begin
        FolderDataFiles := ExtractFilePath(Application.ExeName);
      end else
        begin
          TempFilename := frmMain.LabeledEditDataFolder.Text;
          if not TempFilename.EndsWith('\') then TempFilename := TempFilename + '\';
          frmMain.LabeledEditDataFolder.Text := TempFilename;
          FolderDataFiles := frmMain.LabeledEditDataFolder.Text;
        end;
  if (not PathExists(frmMain.LabeledEditMediaFolder.Text)) and (frmMain.LabeledEditMediaFolder.Text <> '') then
    begin
      ShowMyMessage('Please correct the following:', 'The Media folder specified does not exist. Please specify a different folder', true, 0);
      frmMain.LabeledEditMediaFolder.SetFocus;
      exit;
    end else if frmMain.LabeledEditMediaFolder.Text = '' then
      begin
        FolderMediaFiles := ExtractFilePath(Application.ExeName);
      end else
        begin
          TempFilename := frmMain.LabeledEditMediaFolder.Text;
          if not TempFilename.EndsWith('\') then TempFilename := TempFilename + '\';
          frmMain.LabeledEditMediaFolder.Text := TempFilename;
          FolderMediaFiles := frmMain.LabeledEditMediaFolder.Text;
        end;
  LogItStamp('Saving settings to registry',1);
// write all setting values to registry
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FolderScriptFiles', FolderScriptFiles);
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FolderDataFiles', FolderDataFiles);
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FolderMediaFiles', FolderMediaFiles);
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveFormPositionAndState', BoolToStr(frmMain.CheckBoxApplicationState.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveApplicationLog', BoolToStr(frmMain.CheckBoxPersistantApplicationLog.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'SaveOutputLog', BoolToStr(frmMain.CheckBoxPersistantOutputLog.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendMainForm', BoolToStr(frmMain.CheckBoxMainFormAlphablend.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendMainFormLevel', IntToStr(frmMain.TrackBarMainForm.Position));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendUserGUI', BoolToStr(frmMain.CheckBoxUserGUIAlphablend.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendUserGUILevel', IntToStr(frmMain.TrackBarUserGUI.Position));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendScriptRunner', BoolToStr(frmMain.CheckBoxScriptRunnerAlphablend.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendScriptRunnerLevel', IntToStr(frmMain.TrackBarScriptRunner.Position));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'StopExecutionOnWarnings', BoolToStr(frmMain.CheckBoxStopExecutionOnWarnings.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'CheckScriptSanity', BoolToStr(frmMain.CheckBoxScriptSanity.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'ResolveHostnamesInScript', BoolToStr(frmMain.CheckBoxResolveHostnames.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'RunUserGUIFullScreen', BoolToStr(frmMain.CheckBoxGUIFullScreen.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'MenuMouseHover', BoolToStr(frmMain.CheckBoxMenuMouseHover.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'MaximumRunners', IntToStr(frmMain.SpinEditMaxRunners.Value));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'DebugMode', BoolToStr(frmMain.CheckBoxDebugMode.Checked, True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'PauseScriptAtEnd', BoolToStr(frmMain.CheckBoxPauseAtEnd.Checked , True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'KillRunners', BoolToStr(frmMain.CheckBoxKillRunners.Checked , True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'FloatingHints', BoolToStr(frmMain.CheckBoxFloatingHints.Checked , True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'LeaveScriptExpanded', BoolToStr(frmMain.CheckBoxLeaveScriptExpanded.Checked , True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'HideEditor', BoolToStr(frmMain.CheckBoxHideEditor.Checked , True));
  WriteRegistryString(HKEY_CURRENT_USER, RegistryKey, 'LoopDetection', BoolToStr(frmMain.CheckBoxLoopDetection.Checked , True));
// log all setting values
  LogIt('Script folder                  : ' +  FolderScriptFiles,0);
  LogIt('Data folder                    : ' +  FolderDataFiles,0);
  LogIt('Media folder                   : ' +  FolderMediaFiles,0);
  LogIt('Run User View in full screen   : ' +  BoolToStr(frmMain.CheckBoxGUIFullScreen.Checked, True),0);
  LogIt('Menu activates on mouse hover  : ' +  BoolToStr(frmMain.CheckBoxMenuMouseHover.Checked, True),0);
  LogIt('Save form position and state   : ' +  BoolToStr(frmMain.CheckBoxApplicationState.Checked, true),0);
  LogIt('Show floating hints            : ' +  BoolToStr(frmMain.CheckBoxFloatingHints.Checked, True),0);
  LogIt('Leave script expanded          : ' +  BoolToStr(frmMain.CheckBoxLeaveScriptExpanded.Checked, True),0);
  LogIt('Save application log           : ' +  BoolToStr(frmMain.CheckBoxPersistantApplicationLog.Checked, true),0);
  LogIt('Save output log                : ' +  BoolToStr(frmMain.CheckBoxPersistantOutputLog.Checked, true),0);
  LogIt('Debug mode                     : ' +  BoolToStr(frmMain.CheckBoxDebugMode.Checked, True),0);
  LogIt('Stop execution on warnings     : ' +  BoolToStr(frmMain.CheckBoxStopExecutionOnWarnings.Checked, True),0);
  LogIt('Check script sanity before run : ' +  BoolToStr(frmMain.CheckBoxScriptSanity.Checked, True),0);
  LogIt('Hide editor on script run      : ' +  BoolToStr(frmMain.CheckBoxHideEditor.Checked, True),0);
  LogIt('Resolve hostnames in script    : ' +  BoolToStr(frmMain.CheckBoxResolveHostnames.Checked, True),0);
  LogIt('Pause script at end            : ' +  BoolToStr(frmMain.CheckBoxPauseAtEnd.Checked, True),0);
  LogIt('Kill runners at script end     : ' +  BoolToStr(frmMain.CheckBoxKillRunners.Checked, True),0);
  LogIt('Loop detection                 : ' +  BoolToStr(frmMain.CheckBoxLoopDetection.Checked, True),0);
  LogIt('Maximum active runners         : ' +  IntToStr(frmMain.SpinEditMaxRunners.Value),0);
  LogIt('Alphablend scriptpilot         : ' +  BoolToStr(frmMain.CheckBoxMainFormAlphablend.Checked, true),0);
  LogIt('Alphablend scriptpilot level   : ' +  IntToStr(frmMain.TrackBarMainForm.Position),0);
  LogIt('Alphablend user view           : ' +  BoolToStr(frmMain.CheckBoxUserGUIAlphablend.Checked, true),0);
  LogIt('Alphablend user view level     : ' +  IntToStr(frmMain.TrackBarUserGUI.Position),0);
  LogIt('Alphablend runner              : ' +  BoolToStr(frmMain.CheckBoxScriptRunnerAlphablend.Checked, true),0);
  LogIt('Alphablend runner level        : ' +  IntToStr(frmMain.TrackBarScriptRunner.Position),0);
  LogIt('',0);
  ShowMyMessage('Settings saved', 'All settings where saved to the registry and will be restored at next startup', true, 0);
end;

procedure TfrmMain.ButtonUpdateCommandClick(Sender: TObject);
// UPDATE CURRENT COMMAND WITH MODIFIED PARAMETERS
begin
  if not ValidateCommandAndParameters then exit;
  SaveListViewState(frmMain.ListCommands, ListViewState, true);
  frmMain.ListCommands.Items.Item[frmMain.ListCommands.ItemIndex].SubItems[0] := frmMain.ComboBoxCommand.Text;
  frmMain.ListCommands.Items.Item[frmMain.ListCommands.ItemIndex].SubItems[1] := frmMain.ComboBoxP1.Text;
  frmMain.ListCommands.Items.Item[frmMain.ListCommands.ItemIndex].SubItems[2] := frmMain.ComboBoxP2.Text;
  frmMain.ListCommands.Items.Item[frmMain.ListCommands.ItemIndex].SubItems[3] := frmMain.ComboBoxP3.Text;
  frmMain.ListCommands.Items.Item[frmMain.ListCommands.ItemIndex].SubItems[4] := frmMain.ComboBoxP4.Text;
  frmMain.ListCommands.Items.Item[frmMain.ListCommands.ItemIndex].SubItems[5] := frmMain.ComboBoxP5.Text;
end;

procedure TfrmMain.ButtonUpdateCommandEnter(Sender: TObject);
// UPDATE HINT WHEN MOUSE OVER BUTTON
begin
  frmMain.PanelBottom.Caption := frmMain.ButtonUpdateCommand.Hint;
end;

function GenerateListViewChecksum(ListView: TListView): string;
// USED TO CHECK IF THERE HAS BEEN CHANGES TO THE SCRIPT - FOR SAVING CHANGES
var
  i, j: Integer;
  HashBuilder: TStringBuilder;
begin
  HashBuilder := TStringBuilder.Create;
  try
    for i := 0 to ListView.Items.Count - 1 do
      begin
        HashBuilder.Append(ListView.Items[i].Caption);
        for j := 0 to ListView.Items[i].SubItems.Count - 1 do HashBuilder.Append(ListView.Items[i].SubItems[j]);
      end;
    Result := THashMD5.GetHashString(HashBuilder.ToString);
  finally
    HashBuilder.Free;
  end;
end;

procedure SaveScriptToFile(ListView: TListView; const FileName: string);
// SAVE THE SCRIPT CONFIGURATION TO A JSON FILE AND WATERMARK IT - REQUIRES SYSTEM.JSON
var
  i, j: Integer;
  JSONArray: TJSONArray;
  ItemObject: TJSONObject;
  Strings: TStringList;
  Watermark: string;
begin
  FileOpenSaveError := false;
  Watermark := 'ScriptPilot 2 Script file - CodeRed - Jorgen Lanesskog';
  JSONArray := TJSONArray.Create;
  try
    for i := 0 to ListView.Items.Count - 1 do
      begin
        ItemObject := TJSONObject.Create;
        ItemObject.AddPair('caption', ListView.Items[i].Caption);
        for j := 0 to ListView.Items[i].SubItems.Count - 1 do ItemObject.AddPair('subitem' + IntToStr(j), ListView.Items[i].SubItems[j]);
        JSONArray.AddElement(ItemObject);
      end;
        Strings := TStringList.Create;
          try
            Strings.Add(Watermark); // Add watermark at the beginning
            Strings.Add(JSONArray.ToString);
              try
                Strings.SaveToFile(FileName);
                  except
                    on E: EFOpenError do
                      begin
                        FileOpenSaveError := true;
                        LogItStamp('ERROR   : Error saving file: ' + E.Message,0);
                        ShowMyMessage('Error saving file', E.Message, True, 0);
                      end;
              end;
              finally
                Strings.Free;
          end;
    finally
      JSONArray.Free;
  end;
end;

procedure LoadScriptFromFile(ListView: TListView; const FileName: string);
// LOAD THE SCRIPT CONFIGURATION FROM A JSON FILE AND CHECK WATERMARK - REQUIRES SSYSTEM.JSON
var
  Strings: TStringList;
  JSONArray: TJSONArray;
  ItemObject: TJSONObject;
  Item: TListItem;
  i, j: Integer;
  Watermark, FileContent: string;
begin
  FileOpenSaveError := false;
  JSONArray := nil; // Initialization
  Watermark := 'ScriptPilot 2 Script file - CodeRed - Jorgen Lanesskog';
  Strings := TStringList.Create;
  try
    try
      Strings.LoadFromFile(FileName);
      if (Strings.Count > 0) and (Strings[0] = Watermark) then
        begin
          Strings.Delete(0); // Remove the watermark line
          FileContent := Strings.Text.Trim;
          JSONArray := TJSONObject.ParseJSONValue(FileContent) as TJSONArray;
          if Assigned(JSONArray) then
            begin
              ListView.Items.BeginUpdate;
                try
                  ListView.Items.Clear;
                  for i := 0 to JSONArray.Count - 1 do
                    begin
                      ItemObject := JSONArray.Items[i] as TJSONObject;
                      Item := ListView.Items.Add;
                      Item.Caption := ItemObject.Values['caption'].Value;
                      for j := 0 to ItemObject.Count - 2 do // -2 to account for the caption and zero-based index
                        begin
                          if ItemObject.Values['subitem' + IntToStr(j)] <> nil then
                          Item.SubItems.Add(ItemObject.Values['subitem' + IntToStr(j)].Value);
                        end;
                    end;
                  finally
                  ListView.Items.EndUpdate;
                end;
            end else
              begin
                FileOpenSaveError := true;
                LogItStamp('ERROR   : Invalid JSON format.',0);
                ShowMyMessage('Error', 'Invalid JSON format', True, 0);
              end;
        end else
          begin
            FileOpenSaveError := true;
            LogItStamp('ERROR   : Invalid file format or watermark not found.',0);
            ShowMyMessage('Error', 'Invalid file format or watermark not found', True, 0);
          end;
      except
        on E: Exception do
          begin
            FileOpenSaveError := true;
            LogItStamp('ERROR   : Error loading file: ' + E.Message,0);
            ShowMyMessage('Error loading file', E.Message, True, 0);
          end;
    end;
    finally
    Strings.Free;
    if Assigned(JSONArray) then JSONArray.Free;
  end;
end;

procedure InsertScriptFile(ListView: TListView; const FileName: string; InsertIndex: Integer);
// LOAD THE SCRIPT CONFIGURATION FROM A JSON FILE, CHECK WATERMARK, AND INSERT AT A SPECIFIED INDEX
var
  Strings: TStringList;
  JSONArray: TJSONArray;
  ItemObject: TJSONObject;
  Item: TListItem;
  i, j: Integer;
  Watermark, FileContent: string;
begin
  FileOpenSaveError := false;
  JSONArray := nil;
  Watermark := 'ScriptPilot 2 Script file - CodeRed - Jorgen Lanesskog';
  Strings := TStringList.Create;
  try
    try
      Strings.LoadFromFile(FileName);
      if (Strings.Count > 0) and (Strings[0] = Watermark) then
      begin
        Strings.Delete(0);
        FileContent := Strings.Text.Trim;
        JSONArray := TJSONObject.ParseJSONValue(FileContent) as TJSONArray;
        if Assigned(JSONArray) then
        begin
          ListView.Items.BeginUpdate;
          ListView.Items.Delete(InsertIndex);
          try
            for i := 0 to JSONArray.Count - 1 do
            begin
              ItemObject := JSONArray.Items[i] as TJSONObject;
              Item := ListView.Items.Insert(InsertIndex);
              Inc(InsertIndex);
              Item.Caption := ItemObject.Values['caption'].Value;
              for j := 0 to ItemObject.Count - 2 do
              begin
                if ItemObject.Values['subitem' + IntToStr(j)] <> nil then
                  Item.SubItems.Add(ItemObject.Values['subitem' + IntToStr(j)].Value);
              end;
            end;
          finally
            ListView.Items.EndUpdate;
          end;
        end else
        begin
          FileOpenSaveError := true;
          LogItStamp('ERROR   : Invalid JSON format.',0);
        end;
      end else
      begin
        FileOpenSaveError := true;
        LogItStamp('ERROR   : Invalid file format or watermark not found.',0);
      end;
    except
      on E: Exception do
      begin
        FileOpenSaveError := true;
        LogItStamp('ERROR   : Error inserting file: ' + E.Message,0);
      end;
    end;
  finally
    Strings.Free;
    if Assigned(JSONArray) then JSONArray.Free;
  end;
end;

procedure TfrmMain.ButtonInsertCommandClick(Sender: TObject);
// BUTTON CLICK IN COMMAND INDEX TO ADD OR INSERT COMMAND INTO SCRIPT
var
  NewItem: TListItem;
begin
  if not ValidateCommandAndParameters then exit;
  SaveListViewState(frmMain.ListCommands, ListViewState, true);
  if ListCommands.ItemIndex <> -1 then NewItem := ListCommands.Items.Insert(frmMain.ListCommands.ItemIndex) else NewItem := ListCommands.Items.Add;
  if ListCommands.ItemIndex <> -1 then NewItem.Caption := IntToStr(frmMain.ListCommands.ItemIndex) else NewItem.Caption := IntToStr(ListCommands.Items.Count - 1);
  NewItem.SubItems.Add(frmMain.ComboBoxCommand.Text);
  NewItem.SubItems.Add(frmMain.ComboBoxP1.Text);
  NewItem.SubItems.Add(frmMain.ComboBoxP2.Text);
  NewItem.SubItems.Add(frmMain.ComboBoxP3.Text);
  NewItem.SubItems.Add(frmMain.ComboBoxP4.Text);
  NewItem.SubItems.Add(frmMain.ComboBoxP5.Text);
  RenumberListCommands;
  frmMain.ComboBoxCommand.SetFocus;
end;

procedure TfrmMain.ButtonInsertCommandEnter(Sender: TObject);
// MAKE SURE THE STATUSBAR HINT REFLECTS THE EDIT FIELD WHEN MOUSE NOT HOVER OVER OBJECT
begin
  frmMain.PanelBottom.Caption := frmMain.ButtonInsertCommand.Hint;
end;

procedure TfrmMain.ButtonMainAbortClick(Sender: TObject);
// WHEN USER WANTS TO ABORT THE SCRIPT THAT IS RUNNING
var
  UserResponse : Integer;
begin
  Debug('User clicked the Abort button. Asking if user wants to stop script execution');
  UserResponse := ShowMyDialog('Do you want to stop script execution?', 0);
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

procedure HandleSaveOptions(Sender: TObject);
// USED FOR OPEN, SAVE, SAVE AS AND NEW SCRIPT FILES - ONE PROCEDURE FOR HANDLING ALL SCRIPT EDITOR FILE OPERATIONS
var
  UserResponse: Integer;
  SaveDialog: TSaveDialog;
  OpenFileDialog: TOpenDialog;
  SenderButton: TButton;
  ForceSaveAs: Boolean;
  OpenButton : Boolean;
begin
  ForceSaveAs := False;
  OpenButton := False;
  SenderButton := nil;
  if Sender is TButton then // look at the sender option and identify what button called the procedure
    begin
      SenderButton := TButton(Sender);
      if SenderButton.Caption = 'Sa&ve As' then ForceSaveAs := True; // to skip asking if you want to save
      if SenderButton.Caption = '&Open' then OpenButton := True; // to identify an open request
    end;
  ListCommandsCurrentChecksum := GenerateListViewChecksum(frmMain.ListCommands); // read changes checksum
  // if the user hits "new", but the script is already blank and has no changes
  if (ListCommandsCurrentChecksum = ListCommandsInitialChecksum) and (SenderButton.Caption = '&New') and (frmMain.LabelScriptname.Caption = 'Untitled') then exit;
  // if the user hits "save", but the script is already blank and has no changes
  if (ListCommandsCurrentChecksum = ListCommandsInitialChecksum) and (SenderButton.Caption = '&Save') and (not OpenButton) and (frmMain.LabelScriptname.Caption = 'Untitled') then
    begin
      ShowMyMessage('ScriptPilot Editor', 'Please enter at least one command before trying to save the script!', true, 0);
      exit;
    end;
  // if the user hits "save as", but the script is already blank and has no changes
  if (ListCommandsCurrentChecksum = ListCommandsInitialChecksum) and (forceSaveAs) and (not OpenButton) and (frmMain.LabelScriptname.Caption = 'Untitled') then
    begin
      ShowMyMessage('ScriptPilot Editor', 'Please enter at least one command before trying to save the script!', true, 0);
      exit;
    end;
  // if the user hits "save", but there has been no changes to the named script
  if (ListCommandsCurrentChecksum = ListCommandsInitialChecksum) and (SenderButton.Caption = '&Save') then exit;
  // if named script is changed and the user hits save. Skip save dialog and just save.
  if (ListCommandsCurrentChecksum <> ListCommandsInitialChecksum) and (LoadedScriptFilenameFullPath <> '') and (not ForceSaveAs) and (not OpenButton) and ((SenderButton.Caption <> '&New') or AreWeExitingTheApplication) then
    begin
      SaveScriptToFile(frmMain.ListCommands, LoadedScriptFilenameFullPath);
      if FileOpenSaveError then
        begin
          FileOpenSaveError := false;
          Exit;
        end;
    end
  else if (ListCommandsCurrentChecksum <> ListCommandsInitialChecksum) or forceSaveAs then // if script changed or save as button clicked
    begin
      if (LoadedScriptFilenameFullPath = '') and (SenderButton.Caption = '&Save') then ForceSaveAs := true; // if first save of a new file, do not ask to save changes, with the save or save as button
      if not ForceSaveAs then UserResponse := ShowMyDialog('Do you want to save the script changes?', 0) // ask if changes should be saved in case of new or open
      else UserResponse := mrYes; // just kick in the save dialog
      if UserResponse = mrYes then // start save dialog
        begin
          SaveDialog := TSaveDialog.Create(nil);
          try
            SaveDialog.Title := 'Save script file';
            SaveDialog.Filter := 'ScriptPilot Files (*.sp2)|*.sp2';
            SaveDialog.DefaultExt := 'sp2';
            SaveDialog.InitialDir := FolderScriptFiles;
            SaveDialog.Options := [ofOverwritePrompt, ofPathMustExist, ofNoReadOnlyReturn, ofHideReadOnly];
            if SaveDialog.Execute then
              begin
                SaveScriptToFile(frmMain.ListCommands, SaveDialog.FileName); // save editor list view to json format
                if FileOpenSaveError then // the SaveScriptToFile procedure sets this boolean during operation if any errors during save
                  begin
                    FileOpenSaveError := false; // reset to false and exit
                    Exit;
                  end;
                LoadedScriptFilenameFullPath := SaveDialog.FileName;
                LoadedScriptFilenameNoPath := ExtractFileName(LoadedScriptFilenameFullPath);
              end else exit;
          finally
            SaveDialog.Free;
          end;
        end
      else if UserResponse = mrCancel then Exit;
    end;
  if SenderButton.Caption = '&New' then // set file and path variables to blank and caption to untitled
    begin
      frmMain.ListCommands.Clear;
      LoadedScriptFilenameNoPath := '';
      LoadedScriptFilenameFullPath := '';
      frmMain.LabelScriptName.Caption := 'Untitled';
      if not AreWeExitingTheApplication then LogItStamp('Started a new script',0);
    end else if SenderButton.Caption = '&Open' then // kick in the open dialog and open the file
    begin
      OpenFileDialog := TOpenDialog.Create(nil);
        try
          OpenFileDialog.Title := 'Open Script File';
          OpenFileDialog.Filter := 'ScriptPilot Files (*.sp2)|*.sp2';
          OpenFileDialog.DefaultExt := 'sp2';
          OpenFileDialog.InitialDir := FolderScriptFiles;
          OpenFileDialog.Options := [ofFileMustExist, ofPathMustExist, ofHideReadOnly];
          if OpenFileDialog.Execute then
            begin
              LoadScriptFromFile(frmMain.ListCommands, OpenFileDialog.FileName);
              if FileOpenSaveError then // the LoadScriptFromFile procedure sets this boolean during operation if any errors during open
                begin
                  FileOpenSaveError := false; // reset to false and exit
                  Exit;
                end;
              LoadedScriptFilenameFullPath := OpenFileDialog.FileName;
              LoadedScriptFilenameNoPath := ExtractFileName(LoadedScriptFilenameFullPath);
              frmMain.LabelScriptName.Caption := LoadedScriptFilenameNoPath;
              LogItStamp('Opened script file "' + LoadedScriptFilenameNoPath + '"',0)
            end;
          finally
            OpenFileDialog.Free;
        end;
    end else if SenderButton.Caption = '&Save' then
    begin
      frmMain.LabelScriptName.Caption := LoadedScriptFilenameNoPath;
      LogItStamp('Updatet script file "' + LoadedScriptFilenameNoPath + '"',0)
    end else if SenderButton.Caption = 'Sa&ve As' then
    begin
      frmMain.LabelScriptName.Caption := LoadedScriptFilenameNoPath;
      LogItStamp('Saved script file as "' + LoadedScriptFilenameNoPath + '"',0)
    end;
  SetLength(ListViewState, 0);
  SetLength(UndoStack, 0);
  SetLength(RedoStack, 0);
  if not AreWeExitingTheApplication then  // only execute if we are not exiting the app
    begin
      frmMain.ListCommands.SetFocus;
      frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
      frmMain.GroupBoxCommandIndex.Enabled := frmMain.ListCommands.SelCount < 2;
    end;
  ListCommandsInitialChecksum := GenerateListViewChecksum(frmMain.ListCommands); // reset changes checksum
end;

Procedure UpdateCommandInfo;
// USED TO UPDATE COMMAND PARAMETERS IN COMMAND INDEX (IN EDITOR) + COMMAND DESCRIPTIONS
var
  CurrentCommandInfo: TCommandInfo;
  ParamType: TParamType;
  ComboboxNumber : Integer;
  ComboBox : TComboBox;
  Prefix : String;
  Key: string;
begin
  if CommandMap.TryGetValue(LowerCase(frmMain.ComboBoxCommand.Text), CurrentCommandInfo) then
    begin
      ScanScriptVarAndTags;
      frmMain.LabelCommandName.Caption := CurrentCommandInfo.ParametersLong[0];
      frmMain.LabelCommandShortDescription.Caption := CurrentCommandInfo.SmallDescription;
      frmMain.LabelCommandLongDescription.Caption := CurrentCommandInfo.LargeDescription;
      frmMain.LabelCommandName.Width := frmMain.ScrollBoxDescriptions.ClientWidth - 20;
      frmMain.LabelCommandShortDescription.Width := frmMain.ScrollBoxDescriptions.ClientWidth - 20;
      frmMain.LabelCommandLongDescription.Width := frmMain.ScrollBoxDescriptions.ClientWidth - 20;
      // update all edit fields according to command selected
      for ComboboxNumber := 1 to 5 do
        begin
          ComboBox := frmMain.FindComponent('ComboBoxP' + IntToStr(ComboBoxNumber)) as TComboBox;
          ParamType := CurrentCommandInfo.Types[ComboBoxNumber];
          Prefix := 'Parameter ' + IntToStr(ComboboxNumber) + ' : ';
          with ComboBox do
            begin
              items.BeginUpdate;
              Clear;
              case ParamType of
                ptNotUsed:
                  begin
                    Enabled := false;
                    Visible := false;
                  end;
                ptBoolean:
                  begin
                    ComboBox.Items.Add('True');
                    ComboBox.Items.Add('False');
                    Style := csDropDownList;
                    TextHint := Prefix + CurrentCommandInfo.Parameters[ComboboxNumber];
                    Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                    if CurrentCommandInfo.CanUseVariables[ComboboxNumber] then
                      begin
                        for key in EditVariables.GetAllKeys do
                          begin
                            ComboBox.Items.Add(Key);
                          end;
                      end;
                    Enabled := true;
                    Visible := true;
                    if UpdateFromCommandList then ComboBox.ItemIndex := ComboBox.Items.IndexOf(LastSelectedItem.SubItems[ComboBoxNumber]);
                  end;
                ptOperator:
                  begin
                    ComboBox.Items.Add('Equal');
                    ComboBox.Items.Add('Less');
                    ComboBox.Items.Add('More');
                    Style := csDropDownList;
                    TextHint := Prefix + CurrentCommandInfo.Parameters[ComboboxNumber];
                    Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                    Enabled := true;
                    Visible := true;
                    if UpdateFromCommandList then ComboBox.ItemIndex := ComboBox.Items.IndexOf(LastSelectedItem.SubItems[ComboBoxNumber]);
                  end;
                ptSSHConnection:
                  begin
                    SearchListViewAndFillComboBox('SSHConnection', frmMain.ListCommands, ComboBox);
                    Style := csDropDownList;
                    TextHint := Prefix + CurrentCommandInfo.Parameters[ComboboxNumber];
                    Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                    Enabled := true;
                    Visible := true;
                    if UpdateFromCommandList then ComboBox.ItemIndex := ComboBox.Items.IndexOf(LastSelectedItem.SubItems[ComboBoxNumber]);
                  end;
                ptExcelFile, ptTextFile:
                  begin
                    PopulateComboBoxWithFiles(ComboBox, FolderDataFiles);
                    Style := csDropDownList;
                    TextHint := Prefix + 'Select a file from the dropdown';
                    Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                    Enabled := true;
                    Visible := true;
                    if UpdateFromCommandList then ComboBox.ItemIndex := ComboBox.Items.IndexOf(LastSelectedItem.SubItems[ComboBoxNumber]);
                  end;
                ptScriptFile:
                  begin
                    PopulateComboBoxWithFiles(ComboBox, FolderScriptFiles);
                    Style := csDropDownList;
                    TextHint := Prefix + 'Select a file from the dropdown';
                    Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                    Enabled := true;
                    Visible := true;
                    if UpdateFromCommandList then ComboBox.ItemIndex := ComboBox.Items.IndexOf(LastSelectedItem.SubItems[ComboBoxNumber]);
                  end;
                ptImageFile, ptVideoFile:
                  begin
                    PopulateComboBoxWithFiles(ComboBox, FolderMediaFiles);
                    Style := csDropDownList;
                    TextHint := Prefix + 'Select a file from the dropdown';
                    Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                    Enabled := true;
                    Visible := true;
                    if UpdateFromCommandList then ComboBox.ItemIndex := ComboBox.Items.IndexOf(LastSelectedItem.SubItems[ComboBoxNumber]);
                  end else
                    begin
                      TextHint := Prefix + CurrentCommandInfo.Parameters[ComboboxNumber];
                      Hint := CurrentCommandInfo.ParametersLong[ComboboxNumber];
                      if CurrentCommandInfo.CanUseVariables[ComboboxNumber] then
                        begin
                          for key in EditVariables.GetAllKeys do
                            begin
                              ComboBox.Items.Add(Key);
                            end;
                        end;
                      if CurrentCommandInfo.CanUseTags[ComboboxNumber] then
                        begin
                          for key in EditTags.GetAllKeys do
                            begin
                              ComboBox.Items.Add(Key);
                            end;
                        end;
                      Style := csDropDown;
                      Enabled := true;
                      Visible := true;
                      if UpdateFromCommandList then ComboBox.Text := LastSelectedItem.SubItems[ComboBoxNumber];
                    end;
              end; {case of}
              items.EndUpdate;
            end;
        end;
    // if command is not found in the dictionary, do this
    end else
    begin
      frmMain.LabelCommandName.Caption := 'No matching command found in the command map';
      frmMain.LabelCommandShortDescription.Caption := 'Please edit your search';
      frmMain.LabelCommandLongDescription.Caption := '';
      for ComboboxNumber := 1 to 5 do
        begin
          ComboBox := frmMain.FindComponent('ComboBoxP' + IntToStr(ComboBoxNumber)) as TComboBox;
          with ComboBox do
            begin
              ComboBox.Clear;
              enabled := false;
              visible := false;
            end;
        end;
    end;
  UpdateFromCommandList := false;
  frmMain.ButtonScriptRun.Enabled := frmMain.ListCommands.Items.Count <> 0;
  frmMain.ButtonUpdateCommand.Enabled := frmMain.ListCommands.SelCount = 1;
  frmMain.ButtonInsertCommand.Enabled := frmMain.ListCommands.SelCount < 2;
  frmMain.GroupBoxCommandIndex.Enabled := frmMain.ListCommands.SelCount < 2;
end;

Function ValidateCommandAndParameters : Boolean;
// VALIDATE COMMAND AND ALL PARAMETERS
var
  ErrorInfo : String;
  ComboboxNumber : Integer;
  ComboBox : TComboBox;
  ErrorMessage : String;
begin
  Result := false;
  for ComboboxNumber := 1 to 5 do
    begin
      ComboBox := frmMain.FindComponent('ComboBoxP' + IntToStr(ComboBoxNumber)) as TComboBox;
      if not ValidateCommand(frmMain.ComboBoxCommand.Text, ComboBox.Text, ComboBoxNumber, ErrorInfo) then
        begin
          if ErrorInfo = 'Not a valid command' then
            begin
              ErrorMessage := 'Not a valid command, or no command selected.' + #10#13 + 'Please try again';
              ShowMyMessage(ErrorInfo, ErrorMessage, True, 0);
              Exit;
            end else
            begin
              ErrorMessage :=
              'Validation failed for parameter' + IntToStr(ComboBoxNumber) + '  -  ' + Combobox.TextHint +  #10#13 + #10#13 +
              '"' + ComboBox.Text + '" is not a valid value  -  ' + ErrorInfo;
              ShowMyMessage(ErrorInfo, ErrorMessage, True, 0);
              // exit with error message and do not update command to the ListCommands
              exit;
            end;
        end;
    end;
  Result := true;
end;

Procedure GetUserGUIReadyForRun;
// CALLED FROM THE RUN BUTTON WHEN SCRIPT STARTS
begin
  OKButtonPressed := false;
  AbortButtonClicked := false;
  if frmMain.CheckBoxHideEditor.Checked then frmMain.Visible := false;
  frmUserGUI.CardPanelGUITop.ActiveCard := frmUserGUI.CardGUITopBlank;
  frmUserGUI.CardPanelGUIClient.ActiveCard := frmUserGUI.CardGUIClientBlank;
  frmUserGUI.CardPanelGUIBottom.ActiveCard := frmUserGUI.CardGUIBottomBlank;
  frmUserGUI.Visible := true;
  if frmMain.CheckBoxApplicationState.Checked then
    begin
      frmUserGUI.Top := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowTop', '100'));
      frmUserGUI.Left := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowLeft', '320'));
      frmUserGUI.Width := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowWidth', '1244'));
      frmUserGUI.Height := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowHeight', '761'));
      if StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'UserGUIWindowState', '0')) <> 2 then frmUserGUI.WindowState := wsNormal else frmUserGUI.WindowState := wsMaximized;
    end;
  frmUserGUI.AlphaBlend := StrToBool(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendUserGUI', 'false'));
  frmUserGUI.AlphaBlendValue := StrToInt(ReadRegistryString(HKEY_CURRENT_USER, RegistryKey, 'AlphablendUserGUILevel', '255'));
  if frmMain.CheckBoxGUIFullScreen.Checked then
    begin
      frmUserGUI.BorderStyle := bsNone;
      frmUserGUI.WindowState := wsMaximized;
    end;
  frmUserGUI.SetFocus;
end;

procedure SearchListViewAndFillComboBox(const SearchWord: string; ListView: TListView; ComboBox: TComboBox);
// SEARCH LISTVIEW FOR ITEMS AND FILL A COMBOBOX WITH RESULT
var
  i: Integer;
  ListItem: TListItem;
begin
  for i := 0 to ListView.Items.Count - 1 do
  begin
    ListItem := ListView.Items[i];
    if ListItem.SubItems.Count >= 2 then
    begin
      if AnsiContainsText(ListItem.SubItems[0], SearchWord) then
      begin
        ComboBox.Items.Add(ListItem.SubItems[1]);
      end;
    end;
  end;
end;

procedure VariableStringReplace(var Str: Variant);
// REPLACE TEXT IN STRINGS WITH VALUES FROM THE VARIABLE DATA
var
  Key: string;
  Keys: TArray<string>;
  PosStart, KeyLen, InsertLen: Integer;
  FindString, ReplaceValue: String;
begin
  FindString := VarToStr(Str);
  Keys := Variables.GetAllKeys;
  for Key in Keys do
  begin
    KeyLen := Length(Key);
    PosStart := Pos(Key, FindString);
    while PosStart > 0 do
    begin
      Delete(FindString, PosStart, KeyLen);
      ReplaceValue := Variables.GetValue(Key);
      Insert(ReplaceValue, FindString, PosStart);
      InsertLen := Length(ReplaceValue);
      PosStart := Pos(Key, FindString, PosStart + InsertLen - 1);
    end;
  end;
  Str := FindString; // Assign the modified string back to the variant
end;


procedure SetShowHintsForFormComponents(AForm: TForm; ShowHints: Boolean);
// ENABLE OR DISABLE THE SHOWHINTS OF COMPONENTS
var
  i: Integer;
  AComponent: TComponent;
begin
  for i := 0 to AForm.ComponentCount - 1 do
  begin
    AComponent := AForm.Components[i];
    if AComponent is TControl then
      TControl(AComponent).ShowHint := ShowHints;
  end;
end;

procedure ShowMyMessage(Title, Msg: String; Modal: Boolean; Index: Integer);
// MY PERSONAL SHOWMESSAGE
var
  Frm: TfrmMessage;
  PrevAlpha : Boolean;
  PrevAlphaValue : Integer;
begin
  PrevAlpha := frmMain.AlphaBlend;
  PrevAlphaValue := frmMain.AlphaBlendValue;
  frmMain.AlphaBlend := true;
  frmMain.AlphaBlendValue := 200;
  Frm := TfrmMessage.Create(frmMain);
  Frm.LabelMessageTitle.Caption := Title;
  Frm.LabelMessageText.Caption := Msg;
  if Modal then Frm.ShowModal else Frm.Show;
  frmMain.AlphaBlend := PrevAlpha;
  frmMain.AlphaBlendValue := PrevAlphaValue;
end;

function ShowMyDialog(Title: String; Index: Integer): TModalResult;
// MY PERSONAL SHOWMESSAGEDLG
var
  Frm: TfrmMyDialog;
  PrevAlpha : Boolean;
  PrevAlphaValue : Integer;
begin
  PrevAlpha := frmMain.AlphaBlend;
  PrevAlphaValue := frmMain.AlphaBlendValue;
  frmMain.AlphaBlend := true;
  frmMain.AlphaBlendValue := 200;
  Frm := TfrmMyDialog.Create(frmMain);
  try
    Frm.LabelMyQuestion.Caption := Title;
    Result := Frm.ShowModal;
  finally
    Frm.Free;
  end;
  frmMain.AlphaBlend := PrevAlpha;
  frmMain.AlphaBlendValue := PrevAlphaValue;
end;

procedure ScanScriptVarAndTags;
// SCAN ACTIVE SCRIPT FOR VARIABLES AND ADD THEM TO THE EDITORVARIABLES DICTIONARY
var
  I: Integer;
begin
  EditVariables.Clear;
  EditTags.Clear;
  for I := 0 to frmMain.ListCommands.Items.Count - 1 do
    begin
     if frmMain.ListCommands.Items[I].SubItems[0] = 'SetVariable' then EditVariables.AddOrUpdate(frmMain.ListCommands.Items[I].SubItems[1],IntToStr(I));
     if frmMain.ListCommands.Items[I].SubItems[0] = 'SetTag' then EditTags.AddOrUpdate(frmMain.ListCommands.Items[I].SubItems[1], IntToStr(I));
    end;
end;

Procedure ChangeParamTypesBeforeRun;
begin
  if CommandMap.TryGetValue('ifvariable', Info) then
    begin
      Info.Types[1] := ptRequiredText;
      Info.Types[4] := ptPositiveNumber;
      CommandMap.AddOrSetValue('ifvariable', Info);
    end;
  if CommandMap.TryGetValue('gototag', Info) then
    begin
      Info.Types[1] := ptPositiveNumber;
      CommandMap.AddOrSetValue('gototag', Info);
    end;
end;

Procedure ChangeParamTypesAfterRun;
begin
  if CommandMap.TryGetValue('ifvariable', Info) then
    begin
      Info.Types[1] := ptVariableOrTag;
      Info.Types[4] := ptVariableOrTag;
      CommandMap.AddOrSetValue('ifvariable', Info);
    end;
  if CommandMap.TryGetValue('gototag', Info) then
    begin
      Info.Types[1] := ptVariableOrTag;
      CommandMap.AddOrSetValue('gototag', Info);
    end;
end;

end.
