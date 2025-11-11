unit commands;

interface

Uses System.Generics.Collections, System.SysUtils, System.Classes, WinSock, WinSock2, WinAPI.Windows,
     runner, intercode;

// PROCEDURE POINTER TO ADD A PROCEDURE TO A COMMAND RECORD
Type
  TProcedurePointer = procedure;

// PARAMETER TYPE TO DO VALIDATION OF THE PARAMETERS FOR A COMMAND
type
  TParamType = (
    ptText,                       // CAN BE ANY TEXT. NO RESTRICTIONS
    ptRequiredText,               // MUST BE SOME TEXT. CANNOT BE BLANK
    ptPositiveNumber,             // MUST BE A POSTITIVE NUMBER
    ptNegativeNumber,             // MUST BE A NEGATIVE NUMBER
    ptRange,                      // MUST BE A NUMBER WITHIN A DEFINED RANGE
    ptBoolean,                    // MUST BE TRUE OR FALSE
    ptIPAddress,                  // MUST BE AN IP ADDRESS OF IPV4 OR IPV6
    ptHostname,                   // MUST BE A HOSTNAME - DNS LOOKUP
    ptIPAddressorHostname,        // MUST BE AN IP ADDRESS OR A HOSTNAME
    ptDate,                       // MUST BE A DATE
    ptExcelSheet,                 // MUST BE AN EXCEL WORKBOOK SHEET
    ptPassword,                   // WILL BE TREATED AS A PASSWORD
    ptExcelFile,                  // MUST BE AN EXCEL FILE THAT EXISTS
    ptTextFile,                   // MUST BE A TEXT FILE THAT EXISTS
    ptVideoFile,                  // MUST BE A VIDEO FILE THAT EXISTS
    ptImageFile,                  // MUST BE AN IMAGE FILE THAT EXISTS
    ptScriptFile,                 // MUST BE A SCRIPTPILOT SCRIPTFILE
    ptSSHConnection,              // MUST BE A DEFINED CONNECTION NAME
    ptVariableOrTag,              // MUST START AND END WITH A HASHTAG #
    ptOperator,                   // MUST BE EQUAL, LESS OR MORE
    ptNotUsed                     // PARMETER IS NOT USED
  );

// FOR RESOLVING HOSTNAMES FOR IPV4 AND IPV6
type
  addrinfo = record
    ai_flags: Integer;
    ai_family: Integer;
    ai_socktype: Integer;
    ai_protocol: Integer;
    ai_addrlen: size_t;
    ai_canonname: PAnsiChar;
    ai_addr: PSockAddr;
    ai_next: Paddrinfo;
  end;

// A RECORD HOLDING ALL INFORMATION FOR A COMMAND
type
  TCommandInfo = record
    Proc: TProcedurePointer;                        // POINTER TO A PROCEDURE - PROCEDURE WILL HAVE SAME NAME AS THE COMMAND
    SmallDescription: string;                       // SMALL DESCRIPTION OF THE COMMAND ITSELF
    LargeDescription: string;                       // LARGER DESCRIPTION OF THE COMMAND ITSELF. CONTAINS COMBINED LINE1-LINE5 STRINGS
    Parameters: array[0..5] of string;              // SMALL DESCRIPTION OF UP TO 5 PARAMETERS - IF CONTAINS "NOTUSED", THEN IT WILL BE DISCARED / OVERLOOKED
    ParametersLong: array[0..5] of string;          // LARGER DESCRIPTION OF UP TO 5 PARAMETERS
    StringReplacement: array[0..5] of Boolean;      // BOOLEAN VALUES FOR SELECTING IF PARAMETERS SHOULD BE STRING REPLACED OR NOT DURING EXECUTION
    CanUseVariables: array[0..5] of Boolean;        // PARAMETERS CAN BE SWAPPED FOR VARIABLES
    CanUseTags: array[0..5] of Boolean;             // PARAMETERS CAN BE SWAPPED FOR TAGS
    Value: array[0..5] of string;                   // SET A DEFAULT VALUE IF NEEDED
    Types: array[0..5] of TParamType;               // SET THE PARAMETER TYPE FOR VALIDATION
    MinValue: array[0..5] of Integer;               // SET A RANGE MINIMUM INTEGER VALUE
    MaxValue: array[0..5] of Integer;               // SET A RANGE MAXIMUM INTEGER VALUE
    Example: array[1..5] of string;                 // EXAMPLE OF USAGE
  end;

Procedure AddAllCommandsToDictionary;   // WILL ADD EVERY COMMAND IN THIS UNIT TO THE COMMAND DICTIONARY
function ValidateCommand(Command, ParameterValue: string; ParameterNumber : Integer; out ErrorInfo: string): Boolean;
function getaddrinfo(node: PAnsiChar; service: PAnsiChar; const hints: Paddrinfo; out res: Paddrinfo): Integer; stdcall; external 'ws2_32.dll' name 'getaddrinfo';
procedure freeaddrinfo(ai: Paddrinfo); stdcall; external 'ws2_32.dll' name 'freeaddrinfo';
function getnameinfo(sa: PSockAddr; salen: Integer; host: PAnsiChar; hostlen: DWORD; serv: PAnsiChar; servlen: DWORD; flags: Integer): Integer; stdcall; external 'ws2_32.dll' name 'getnameinfo';
Procedure ResetInfo;

const
  IPv4_OCTETS = 4;
  IPv6_GROUPS = 8;
  MAX_HEX_GROUP_VALUE = $FFFF;
  HexSet: TSysCharSet = ['0'..'9', 'A'..'F', 'a'..'f'];
  NI_MAXHOST = 1025;
  NI_NUMERICHOST = 1;

var
  CommandMap: TDictionary<string, TCommandInfo>;  // COMMANDMAP IS A DICTIONARY OF ALL THE COMMANDS WITH BELONGING RECORDS
  ResolveHostnamesInScript : Boolean;
  Info: TCommandInfo;
  Line1, Line2, Line3, Line4, Line5 : String;

implementation

Procedure AddAllCommandsToDictionary;
// ADD ALL COMMANDS TO THE COMMANDMAP DICTIONARY
begin
  CommandMap := TDictionary<string, TCommandInfo>.Create;


// START DEFINING ALL THE COMMAND RECORDS FROM HERE

// ************************************************************************
// COMMAND Remark BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'remark';
  Info.ParametersLong[0] := 'Remark';
  Info.SmallDescription := 'Remark sections of the script';
  Line1 := 'Introduce a comment into your script for clarity and documentation purposes.';
  Line2 := 'Leverage any parameter to embed descriptive text, enhancing script readability.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Parameters[1] := 'Description text';
  Info.Parameters[2] := 'Description text';
  Info.Parameters[3] := 'Description text';
  Info.Parameters[4] := 'Description text';
  Info.Parameters[5] := 'Description text';
  Info.ParametersLong[1] := 'Add text in all or any of the parameters to use them as descriptions for whatever purpose';
  Info.ParametersLong[2] := 'Add text in all or any of the parameters to use them as descriptions for whatever purpose';
  Info.ParametersLong[3] := 'Add text in all or any of the parameters to use them as descriptions for whatever purpose';
  Info.ParametersLong[4] := 'Add text in all or any of the parameters to use them as descriptions for whatever purpose';
  Info.ParametersLong[5] := 'Add text in all or any of the parameters to use them as descriptions for whatever purpose';
  Info.Types[1] := ptText;
  Info.Types[2] := ptText;
  Info.Types[3] := ptText;
  Info.Types[4] := ptText;
  Info.Types[5] := ptText;
  Info.StringReplacement[1] := true;
  Info.StringReplacement[2] := true;
  Info.StringReplacement[3] := true;
  Info.StringReplacement[4] := true;
  Info.StringReplacement[5] := true;
  Info.CanUseVariables[1] := true;
  Info.CanUseVariables[2] := true;
  Info.CanUseVariables[3] := true;
  Info.CanUseVariables[4] := true;
  Info.CanUseVariables[5] := true;
  Info.Example[1] := 'This is a descriptive text';
  Info.Example[2] := 'It can be used to descibe';
  Info.Example[3] := 'sections of commands';
  Info.Example[4] := 'Any parameters can have content';
  Info.Example[5] := 'or not';
  Info.Proc := @Remark;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND Remark END
// ************************************************************************

// ************************************************************************
// COMMAND SetTag BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'settag';
  Info.ParametersLong[0] := 'SetTag';
  Info.SmallDescription := 'Create and add a new tag to the script';
  Line1 := 'Establishes a navigational anchor within your script, identified by a unique tag.';
  Line2 := '';
  Line3 := 'Enable dynamic script flow by using commands that seek out tags for conditional execution,';
  Line4 := 'or use commands like "GotoTag" to directly jump to a designated script row.' + #10#13;
  Line5 := 'Tag names MUST start and end with a hashtag (#) - Example #MyTag# - The names are also case sensitive in use.';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Name of tag - Start and end with #';
  Info.ParametersLong[1] := 'Enter the name of the tag. Make sure to use a unique name that does not match other words used in the script. Example: #MySpecialTag#';
  Info.Types[1] := ptVariableOrTag;
  Info.Example[1] := '#MySpecialTag#';
  Info.Proc := @SetTag;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND SetTag END
// ************************************************************************

// ************************************************************************
// COMMAND GotoTag BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'gototag';
  Info.ParametersLong[0] := 'GotoTag';
  Info.SmallDescription := 'Jump script execution to a defined tag in the script';
  Line1 := 'Navigate to a point in the script that has the defined tag set with "SetTag".';
  Line2 := '';
  Line3 := 'Script will jump from the current row, to the row containing the specified tag, and continue execution from there.' + #10#13;
  Line4 := 'Be careful not to create script loops when jumping backwards in the script.';
  Line5 := '';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Name of tag to jump to';
  Info.ParametersLong[1] := 'Enter the name of a tag to jump to';
  Info.Types[1] := ptVariableOrTag;
  Info.StringReplacement[1] := true;
  Info.CanUseTags[1] := true;
  Info.Example[1] := '#MyGotoTag#';
  Info.Proc := @GotoTag;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND GotoTag END
// ************************************************************************

// ************************************************************************
// COMMAND SetRunnerActive BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'setrunneractive';
  Info.ParametersLong[0] := 'SetRunnerActive';
  Info.SmallDescription := 'Activate or create a runner for work';
  Line1 := 'Create and/or activate a specific runner to undertake command execution within its own dedicated runspace.' + #10#13 + 'Meaning, when this runner execute commands, it will not impact other runners.' + #10#13;
  Line2 := 'You can have multiple runners working with commands at the same time, but only one (the active one) can receive commands from the script, at any given moment.' + #10#13;
  Line3 := 'Not being active, in this case, only mean they are not actively receiving new commands.' + #10#13 + 'They could still be working with previously received commands.' + #10#13;
  Line4 := 'The active runner will be able to receive commands until a different runner is selected, or until the script ends.' + #10#13;
  Line5 := 'You can automatically terminate runners by setting the "Kill on idle" to "True". Otherwise, they will keep running, and be ready for work, until the script ends.';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Name of the runner';
  Info.ParametersLong[1] := 'Enter the name of the runner to create and/or activate for the upcoming commands';
  Info.Parameters[2] := 'Kill runner on idle';
  Info.ParametersLong[2] := 'Automatically remove the runner when it is out of commands and running idle';
  Info.Parameters[3] := 'Hide runner';
  Info.ParametersLong[3] := 'Make the runner invisible and hide it from the Windows Taskbar';
  Info.Types[1] := ptRequiredText;
  Info.Types[2] := ptBoolean;
  Info.Types[3] := ptBoolean;
  Info.StringReplacement[1] := true;
  Info.StringReplacement[2] := true;
  Info.StringReplacement[3] := true;
  Info.CanUseVariables[1] := true;
  Info.CanUseVariables[2] := true;
  Info.CanUseVariables[3] := true;
  Info.Example[1] := 'MyBestRunner';
  Info.Example[2] := 'True';
  Info.Example[3] := 'False';
  Info.Proc := @SetRunnerActive;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND SetRunnerActive END
// ************************************************************************

// ************************************************************************
// COMMAND WaitForButtonClick BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'waitforbuttonclick';
  Info.ParametersLong[0] := 'WaitForButtonClick';
  Info.SmallDescription := 'Pause script and wait for the user to click a button';
  Line1 := 'Pauses the script and no more commands will be executed, before the user clicks a button.';
  Line2 := 'Runners are not impacted by this command and will continue working, if they are not idle.';
  Line3 := 'The current active runner will not receive any new commands, before the user clicks the button.';
  Line4 := '';
  Line5 := 'Specify the message to display to the user, and the text to display on the button that must be clicked, before the script continues.';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Message to display';
  Info.Parameters[2] := 'Text on button';
  Info.ParametersLong[1] := 'Enter the message to display to the user. This message will be displayed to the left of the button.';
  Info.ParametersLong[2] := 'Enter the text on the button. Example: "OK" or "Continue"';
  Info.Types[1] := ptRequiredText;
  Info.Types[2] := ptRequiredText;
  Info.StringReplacement[1] := true;
  Info.StringReplacement[2] := true;
  Info.CanUseVariables[1] := true;
  Info.CanUseVariables[2] := true;
  Info.Example[1] := 'Click to continue script';
  Info.Example[2] := 'Continue';
  Info.Proc := @WaitForButtonClick;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND WaitForButtonClick END
// ************************************************************************

// ************************************************************************
// COMMAND IncludeScript BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'includescript';
  Info.ParametersLong[0] := 'IncludeScript';
  Info.SmallDescription := 'Insert the content of a script';
  Line1 := 'Insert and include the content of another script at the current position.' + #10#13;
  Line2 := 'Existing commands after the insert point, will be pushed to be executed after the inserted content.';
  Line3 := '';
  Line4 := 'The scripts will merge into a combined script at runtime.';
  Line5 := '';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Select script to insert';
  Info.ParametersLong[1] := 'Select the script to insert from the list.';
  Info.Types[1] := ptScriptFile;
  Info.Example[1] := 'MyExtraScriptFile.sp2';
  Info.Proc := @IncludeScript;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND IncludeScript END
// ************************************************************************

// ************************************************************************
// COMMAND SetVariable BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'setvariable';
  Info.ParametersLong[0] := 'SetVariable';
  Info.SmallDescription := 'Create and add a new variable';
  Line1 := 'Creates a new variable that can be used throughout the script.';
  Line2 := 'If the variable already exist, it will be updated with the value provided.' + #10#13;
  Line3 := 'Variables are used to store values of any kind and you can swap command parameters for variables, when needed.' + #10#13;
  Line4 := 'An example can be gathering user input, like username and password, grabbed to variables and then used later for making SSH connections.' + #10#13 + #10#13 + 'All variables must be created with the Variable command and given an initial value (can be blank), before they can be used elsewhere in the script' + #10#13;
  Line5 := 'Variable names must start and end with a hashtag (#) - Example #MyVariable# - The names are also case sensitive in use.';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Name of variable - Start and end with #';
  Info.Parameters[2] := 'Value of variable';
  Info.ParametersLong[1] := 'Enter the name of the variable. Make sure to use a unique name that does not match other words used in the script. Example: #MySpecialVariable#';
  Info.ParametersLong[2] := 'Enter the value of the variable. The value can be anything from any text to any numbers, or blank.';
  Info.Types[1] := ptVariableOrTag;
  Info.Types[2] := ptText;
  Info.Example[1] := '#MyGreatVariable#';
  Info.Example[2] := '10.10.1.1';
  Info.Proc := @SetVariable;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND SetVariable END
// ************************************************************************

// ************************************************************************
// COMMAND WaitForSeconds BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'waitforseconds';
  Info.ParametersLong[0] := 'WaitForSeconds';
  Info.SmallDescription := 'Pause script and wait for a number of seconds';
  Line1 := 'Pauses the script and no more commands will be executed, before the number of seconds has passed.' + #10#13;
  Line2 := 'Runners are not impacted by this command and will continue working, if they are not idle.';
  Line3 := 'The current active runner will not receive any new commands, before the seconds has passed.';
  Line4 := '';
  Line5 := '';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'Number of seconds (1-32000)';
  Info.ParametersLong[1] := 'Enter the number of seconds to pause the script. It must be within the range of 1 to 32000 seconds';
  Info.Types[1] := ptRange;
  Info.StringReplacement[1] := true;
  Info.MinValue[1] := 1;
  Info.MaxValue[1] := 32000;
  Info.CanUseVariables[1] := true;
  Info.Example[1] := '10';
  Info.Proc := @WaitForSeconds;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND WaitForSeconds END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewTopText BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewtoptext';
  Info.ParametersLong[0] := 'UserViewTopText';
  Info.SmallDescription := 'Display a message at the top of the User View';
  Line1 := 'Display a text message to the user at the top of the User View.' + #10#13;
  Line2 := 'The will linger until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Parameters[1] := 'Message text';
  Info.ParametersLong[1] := 'Enter the text of the message you want to display to the user.';
  Info.Types[1] := ptText;
  Info.StringReplacement[1] := true;
  Info.CanUseVariables[1] := true;
  Info.Example[1] := 'Text on the top';
  Info.Proc := @UserViewTopText;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewTopText END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewTopProgressBar BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewtopprogressbar';
  Info.ParametersLong[0] := 'UserViewTopProgressBar';
  Info.SmallDescription := 'Display a progress bar at the top of the User View';
  Line1 := 'Display a progress bar to the user at the top of the User View and set the position as a value between 0 and 100.' + #10#13;
  Line2 := 'The will linger until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Parameters[1] := 'Progress position';
  Info.ParametersLong[1] := 'Enter the position for the progress bar. It must be a value between 0 and 100.';
  Info.Types[1] := ptRange;
  Info.MinValue[1] := 0;
  Info.MaxValue[1] := 100;
  Info.StringReplacement[1] := true;
  Info.CanUseVariables[1] := true;
  Info.Example[1] := '42';
  Info.Proc := @UserViewTopProgressBar;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewTopProgressBar END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewBottomText BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewbottomtext';
  Info.ParametersLong[0] := 'UserViewBottomText';
  Info.SmallDescription := 'Display a message at the bottom of the User View';
  Line1 := 'Display a text message to the user at the bottom of the User View.' + #10#13;
  Line2 := 'The will linger until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Parameters[1] := 'Message text';
  Info.ParametersLong[1] := 'Enter the text of the message you want to display to the user.';
  Info.Types[1] := ptText;
  Info.StringReplacement[1] := true;
  Info.CanUseVariables[1] := true;
  Info.Example[1] := 'Text at the bottom';
  Info.Proc := @UserViewBottomText;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewBottomText END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewBottomProgressBar BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewbottomprogressbar';
  Info.ParametersLong[0] := 'UserViewBottomProgressBar';
  Info.SmallDescription := 'Display a progress bar at the bottom of the User View';
  Line1 := 'Display a progress bar to the user at the bottom of the User View and set the position as a value between 0 and 100.' + #10#13;
  Line2 := 'The will linger until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Parameters[1] := 'Progress position';
  Info.ParametersLong[1] := 'Enter the position for the progress bar. It must be a value between 0 and 100.';
  Info.Types[1] := ptRange;
  Info.MinValue[1] := 0;
  Info.MaxValue[1] := 100;
  Info.StringReplacement[1] := true;
  Info.CanUseVariables[1] := true;
  Info.Example[1] := '42';
  Info.Proc := @UserViewBottomProgressBar;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewBottomProgressBar END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewClientText BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewclienttext';
  Info.ParametersLong[0] := 'UserViewClientText';
  Info.SmallDescription := 'Display up to 5 lines of text in the User View View';
  Line1 := 'Display up to 5 lines of text in the client area of the User View.' + #10#13;
  Line2 := 'The will linger until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Parameters[1] := 'Text to display';
  Info.Parameters[2] := 'Text to display';
  Info.Parameters[3] := 'Text to display';
  Info.Parameters[4] := 'Text to display';
  Info.Parameters[5] := 'Text to display';
  Info.ParametersLong[1] := 'Enter any text to display';
  Info.ParametersLong[2] := 'Enter any text to display';
  Info.ParametersLong[3] := 'Enter any text to display';
  Info.ParametersLong[4] := 'Enter any text to display';
  Info.ParametersLong[5] := 'Enter any text to display';
  Info.Types[1] := ptText;
  Info.Types[2] := ptText;
  Info.Types[3] := ptText;
  Info.Types[4] := ptText;
  Info.Types[5] := ptText;
  Info.StringReplacement[1] := true;
  Info.StringReplacement[2] := true;
  Info.StringReplacement[3] := true;
  Info.StringReplacement[4] := true;
  Info.StringReplacement[5] := true;
  Info.CanUseVariables[1] := true;
  Info.CanUseVariables[2] := true;
  Info.CanUseVariables[3] := true;
  Info.CanUseVariables[4] := true;
  Info.CanUseVariables[5] := true;
  Info.Example[1] := 'This is text in the';
  Info.Example[2] := 'client area of the';
  Info.Example[3] := 'User View. It can be used';
  Info.Example[4] := 'to relay any information';
  Info.Example[5] := 'to the user.';
  Info.Proc := @UserViewClientText;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewClientText END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewClientClear BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewclientclear';
  Info.ParametersLong[0] := 'UserViewClientClear';
  Info.SmallDescription := 'Clear the user view client area';
  Line1 := 'No parameters. Clear the user view client area.' + #10#13;
  Line2 := 'The area will remain content free until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Types[1] := ptNotUsed;
  Info.Types[2] := ptNotUsed;
  Info.Types[3] := ptNotUsed;
  Info.Types[4] := ptNotUsed;
  Info.Types[5] := ptNotUsed;
  Info.Proc := @UserViewClientClear;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewClientClear END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewTopClear BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewtopclear';
  Info.ParametersLong[0] := 'UserViewTopClear';
  Info.SmallDescription := 'Clear the user view top area';
  Line1 := 'No parameters. Clear the user view top area.' + #10#13;
  Line2 := 'The area will remain content free until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Types[1] := ptNotUsed;
  Info.Types[2] := ptNotUsed;
  Info.Types[3] := ptNotUsed;
  Info.Types[4] := ptNotUsed;
  Info.Types[5] := ptNotUsed;
  Info.Proc := @UserViewTopClear;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewTopClear END
// ************************************************************************

// ************************************************************************
// COMMAND UserViewBottomClear BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'userviewbottomclear';
  Info.ParametersLong[0] := 'UserViewBottomClear';
  Info.SmallDescription := 'Clear the user view bottom area';
  Line1 := 'No parameters. Clear the user view bottom area.' + #10#13;
  Line2 := 'The area will remain content free until something else is shown.';
  Info.LargeDescription := Line1 + #10#13 + Line2;
  Info.Types[1] := ptNotUsed;
  Info.Types[2] := ptNotUsed;
  Info.Types[3] := ptNotUsed;
  Info.Types[4] := ptNotUsed;
  Info.Types[5] := ptNotUsed;
  Info.Proc := @UserViewBottomClear;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND UserViewBottomClear END
// ************************************************************************

// ************************************************************************
// COMMAND IfVariable BEGIN
// ************************************************************************
  ResetInfo;
  Info.Parameters[0] := 'ifvariable';
  Info.ParametersLong[0] := 'IfVariable';
  Info.SmallDescription := 'Check variable value and conditionally jump to a tag';
  Line1 := 'Check variable value and conditionally jump to a tag if the value is less, more or equal to a compared value.';
  Line2 := 'The compared value can be literal or the value of another variable.' + #10#13;
  Line3 := 'Script will jump to defined tag if command evaluates to true, else script will continue on next row.' + #10#13;
  Line4 := 'Equal can compare any values, but less or more can only compare numbers.';
  Info.LargeDescription := Line1 + #10#13 + Line2+ #10#13 + Line3+ #10#13 + Line4;
  Info.Parameters[1] := 'Variable to check';
  Info.Parameters[2] := 'Operator';
  Info.Parameters[3] := 'Value to compare';
  Info.Parameters[4] := 'Tag to jump to if true';
  Info.ParametersLong[1] := 'The name of the variable to compare to a value';
  Info.ParametersLong[2] := 'The operator of equal, less and more';
  Info.ParametersLong[3] := 'Value to compare to';
  Info.ParametersLong[4] := 'The tag that the script should jump to if the command evaluates to true.';
  Info.Types[1] := ptVariableOrTag;
  Info.Types[2] := ptOperator;
  Info.Types[3] := ptText;
  Info.Types[4] := ptVariableOrTag;
  Info.StringReplacement[1] := true;
  Info.StringReplacement[3] := true;
  Info.StringReplacement[4] := true;
  Info.CanUseVariables[1] := true;
  Info.CanUseVariables[3] := true;
  Info.CanUseTags[4] := true;
  Info.Example[1] := '#MyVariable#';
  Info.Example[2] := 'equal';
  Info.Example[3] := '10';
  Info.Example[4] := '#MyJumpTag#';
  Info.Proc := @IfVariable;
  CommandMap.Add(Info.Parameters[0], Info);
// ************************************************************************
// COMMAND IfVariable END
// ************************************************************************

end;  // END COMMAND MAP - ALL COMMANDS ADDED


function IsValidIPv4(const IPAddress: string): Boolean;
var
  Octets: TStringList;
  I, OctetValue: Integer;
begin
  Result := False;
  Octets := TStringList.Create;
  try
    Octets.Delimiter := '.';
    Octets.DelimitedText := IPAddress;
    if Octets.Count <> IPv4_OCTETS then Exit;
    for I := 0 to IPv4_OCTETS - 1 do
    begin
      if not TryStrToInt(Octets[I], OctetValue) or (OctetValue < 0) or (OctetValue > 255) then Exit;
    end;
    Result := True;
  finally
    Octets.Free;
  end;
end;

function IsValidHexGroup(const Group: string): Boolean;
var
  I: Integer;
begin
  for I := 1 to Length(Group) do
  begin
    if not CharInSet(Group[I], ['0'..'9', 'A'..'F', 'a'..'f']) then
    begin
      Exit(False);
    end;
  end;
  Result := True;
end;

function IsValidIPv6(const IPAddress: string): Boolean;
var
  Groups: TStringList;
  I, DoubleColonCount: Integer;
begin
  Result := False;
  Groups := TStringList.Create;
  try
    Groups.Delimiter := ':';
    Groups.DelimitedText := IPAddress;
    DoubleColonCount := 0;
    for I := 0 to Groups.Count - 1 do
    begin
      if Groups[I] = '' then
      begin
        Inc(DoubleColonCount);
        if DoubleColonCount > 1 then Exit;  // More than one '::' is not valid
        Continue;
      end;
      if (Length(Groups[I]) > 4) or not IsValidHexGroup(Groups[I]) then Exit;
    end;
    if (DoubleColonCount = 0) and (Groups.Count <> IPv6_GROUPS) then Exit; // No '::' and not exactly 8 groups
    if (DoubleColonCount = 1) and (Groups.Count > IPv6_GROUPS) then Exit; // '::' present but too many groups
    Result := True;
  finally
    Groups.Free;
  end;
end;

function IsValidIP(const IPAddress: string): Boolean;
begin
  if IPAddress = '' then Exit(False);
  Result := IsValidIPv4(IPAddress) or IsValidIPv6(IPAddress);
end;

function ResolveHostnameToIP(const Hostname: string; var HostnameInfo: string): string;
var
  WSAData: TWSAData;
  Hints: TAddrInfo;
  AddrInfo: PAddrInfo;
  RetVal: Integer;
begin
  Result := '';
  HostnameInfo := '';
  if Hostname = '' then
  begin
    HostnameInfo := 'Empty Hostname';
    Exit;
  end;
  if IsValidIP(Hostname) then
  begin
    HostnameInfo := 'This is an IP address, not a hostname';
    Exit;
  end;
  if not ResolveHostnamesInScript then
  begin
    HostnameInfo := '';
    Exit;
  end;
  if WSAStartup(MAKEWORD(2, 2), WSAData) <> 0 then
  begin
    HostnameInfo := 'WSAStartup failed';
    Exit;
  end;
  try
    FillChar(Hints, SizeOf(Hints), 0);
    Hints.ai_family := AF_UNSPEC; // Supports both IPv4 and IPv6
    Hints.ai_socktype := SOCK_STREAM;
    RetVal := getaddrinfo(PAnsiChar(AnsiString(Hostname)), nil, @Hints, AddrInfo);
    if RetVal <> 0 then
    begin
      HostnameInfo := 'Error resolving hostname: ' + SysErrorMessage(WSAGetLastError);
      Exit;
    end;
  finally
    if AddrInfo <> nil then
      freeaddrinfo(AddrInfo);
    WSACleanup;
  end;
end;

Function ParamTypeToString(Param: TParamType): string;
begin
  case Param of
    ptText:                  Result := 'Any text or blank';
    ptRequiredText:          Result := 'Required text';
    ptPositiveNumber:        Result := 'Positive number';
    ptNegativeNumber:        Result := 'Negative number';
    ptRange:                 Result := 'Within a defined range';
    ptBoolean:               Result := 'True or False';
    ptIPAddress:             Result := 'IPv4 or IPv6 address';
    ptHostname:              Result := 'Resolvable hostname';
    ptIPAddressorHostname:   Result := 'IPv4, IPv6 or resolvable hostname';
    ptDate:                  Result := 'Date string';
    ptExcelSheet:            Result := 'Excel sheet that exists';
    ptPassword:              Result := 'Password text';
    ptExcelFile:             Result := 'Excel file that exists';
    ptTextFile:              Result := 'Text file that exists';
    ptVideoFile:             Result := 'Video file that exists';
    ptImageFile:             Result := 'Image file that exists';
    ptScriptFile:            Result := 'ScriptPilot script file';
    ptSSHConnection:         Result := 'Name of an existing SSH connection';
    ptVariableOrTag:         Result := 'Variable or Tag name that must start and end with a Hastag';
    ptOperator:              Result := 'Must contain operator of Equal, Less or More';
    ptNotUsed:               Result := 'Parameter not used';
  end;
end;

function ValidateCommand(Command, ParameterValue: string; ParameterNumber : Integer; out ErrorInfo: string): Boolean;
// VALIDATION OF COMMANDS //
var
  ParamType: TParamType;
  CurrentCommandInfo: TCommandInfo;
  IntValue: Integer;
begin
  Result := True;  // return True if validation OK, else False
  ErrorInfo := ''; // return any error text explaining what went wrong
  // check command, value and parameter type
  if not CommandMap.TryGetValue(LowerCase(Command), CurrentCommandInfo) then
    begin
      Result := False;
      ErrorInfo := 'Not a valid command';
    end else
    begin
      // Set the ParamType to the active parameter to check within command used
      ParamType := CurrentCommandInfo.Types[ParameterNumber];
      // Handle the different parameter types
      case paramType of
        ptNotUsed, ptText:
          begin
            Result := True;
          end;
        ptRequiredText:
          begin
            if ParameterValue = '' then
              begin
                ErrorInfo := 'Parameter can not be empty';
                Result := false;
              end;
          end;
        ptVariableOrTag:
          begin
            if ParameterValue = '' then
              begin
                ErrorInfo := 'Parameter can not be empty';
                Result := false;
              end else
              begin
                if Pos('#', ParameterValue) <> 1 then
                  begin
                    ErrorInfo := 'Name must start with a hashtag (#)';
                    Result := false;
                  end else if ParameterValue[Length(ParameterValue)] <> '#' then
                    begin
                      ErrorInfo := 'Name must end with a hashtag (#)';
                      Result := false;
                    end;
              end;
          end;
        ptSSHConnection:
          begin
            if ParameterValue = '' then
              begin
                ErrorInfo := 'You must specify an existing SSH connection. Please use SSHConnect first';
                Result := false;
              end;
          end;
        ptPositiveNumber:
          begin
            if not TryStrToInt(ParameterValue, IntValue) then
              begin
                ErrorInfo := 'Parameter must be a number';
                Result := false;
              end;
            if not TryStrToInt(ParameterValue, IntValue) or (IntValue < 0) then
              begin
                ErrorInfo := 'Parameter must be a positive number';
                Result := false;
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptNegativeNumber:
          begin
            if not TryStrToInt(ParameterValue, IntValue) then
              begin
                ErrorInfo := 'Parameter must be a number';
                Result := false;
              end;
            if not TryStrToInt(ParameterValue, IntValue) or (IntValue >= 0) then
              begin
                ErrorInfo := 'Parameter must be a negative number';
                Result := false;
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptRange:
          begin
            if not TryStrToInt(ParameterValue, IntValue) then
              begin
                ErrorInfo := 'Parameter must be a number';
                Result := false;
              end;
            if TryStrToInt(ParameterValue, IntValue) then if (IntValue < CurrentCommandInfo.MinValue[ParameterNumber]) or (IntValue > CurrentCommandInfo.MaxValue[ParameterNumber]) then
              begin
                ErrorInfo := 'Parameter must be number within the range of ' + IntToStr(CurrentCommandInfo.MinValue[ParameterNumber]) + ' and ' + IntToStr(CurrentCommandInfo.MaxValue[ParameterNumber]);
                Result := false;
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptBoolean:
          begin
            if not ((LowerCase(ParameterValue) = 'true') or (LowerCase(ParameterValue) = 'false')) then
              begin
                ErrorInfo := 'Parameter must be a boolean value of "True" or "False"';
                Result := false;
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptOperator:
          begin
            if LowerCase(ParameterValue) <> 'equal' then if LowerCase(ParameterValue) <> 'less' then if LowerCase(ParameterValue) <> 'more' then
              begin
                ErrorInfo := 'Parameter must be value of "Equal", "Less" or "More"';
                Result := false;
              end;
          end;
        ptIPAddress:
          begin
            if not IsValidIP(ParameterValue) then
              begin
                ErrorInfo := 'Parameter is not a valid IPv4 or IPv6 IP address';
                Result := false;
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptHostname:
          begin
            ResolveHostnameToIP(ParameterValue, ErrorInfo);
            if ErrorInfo <> '' then
              begin
                Result := false;
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptIPAddressorHostname:
          begin
            ResolveHostnameToIP(ParameterValue, ErrorInfo);
            if ErrorInfo <> '' then Result := IsValidIP(ParameterValue) else Result := True;
            if Result = False then
              begin
                ErrorInfo := 'Parameter does not contain a valid hostname, as it could not be resolved to an IP address. Nor does it contain a valid IPv4 or IPv6 address';
              end;
            if not ScriptRunning then if (ParameterValue <> '') and (ParameterValue[1] = '#') and (ParameterValue[Length(ParameterValue)] = '#') then result := true;
          end;
        ptDate:
          begin
            //
          end;
        ptExcelSheet:
          begin
            // if ExcelFileScript.GetSheetIndex(ParameterValue, false) >= 0 then Result := true else Result := false;
          end;
        ptPassword:
          begin
            //
          end;
        ptExcelFile:
          begin
            //
          end;
        ptTextFile:
          begin
            //
          end;
        ptVideoFile:
          begin
            //
          end;
        ptImageFile:
          begin
            //
          end;
        ptScriptFile:
          begin
            if ParameterValue = '' then
              begin
                ErrorInfo := 'Parameter can not be empty. You must select a script from the list';
                Result := false;
              end;
          end;
      end;
    end;
end;

Procedure ResetInfo;
begin
  Info.Parameters[0] := '';
  Info.ParametersLong[0] := '';
  Info.SmallDescription := '';
  Line1 := '';
  Line2 := '';
  Line3 := '';
  Line4 := '';
  Line5 := '';
  Info.LargeDescription := Line1 + #10#13 + Line2 + #10#13 + Line3 + #10#13 + Line4 + #10#13 + Line5;
  Info.Parameters[1] := 'NOTUSED';
  Info.Parameters[2] := 'NOTUSED';
  Info.Parameters[3] := 'NOTUSED';
  Info.Parameters[4] := 'NOTUSED';
  Info.Parameters[5] := 'NOTUSED';
  Info.ParametersLong[1] := '';
  Info.ParametersLong[2] := '';
  Info.ParametersLong[3] := '';
  Info.ParametersLong[4] := '';
  Info.ParametersLong[5] := '';
  Info.Types[1] := ptNotUsed;
  Info.Types[2] := ptNotUsed;
  Info.Types[3] := ptNotUsed;
  Info.Types[4] := ptNotUsed;
  Info.Types[5] := ptNotUsed;
  Info.StringReplacement[1] := false;
  Info.StringReplacement[2] := false;
  Info.StringReplacement[3] := false;
  Info.StringReplacement[4] := false;
  Info.StringReplacement[5] := false;
  Info.MinValue[1] := 0;
  Info.MinValue[2] := 0;
  Info.MinValue[3] := 0;
  Info.MinValue[4] := 0;
  Info.MinValue[5] := 0;
  Info.MaxValue[1] := 0;
  Info.MaxValue[2] := 0;
  Info.MaxValue[3] := 0;
  Info.MaxValue[4] := 0;
  Info.MaxValue[5] := 0;
  Info.CanUseVariables[1] := false;
  Info.CanUseVariables[2] := false;
  Info.CanUseVariables[3] := false;
  Info.CanUseVariables[4] := false;
  Info.CanUseVariables[5] := false;
  Info.CanUseTags[1] := false;
  Info.CanUseTags[2] := false;
  Info.CanUseTags[3] := false;
  Info.CanUseTags[4] := false;
  Info.CanUseTags[5] := false;
  Info.Example[1] := '';
  Info.Example[2] := '';
  Info.Example[3] := '';
  Info.Example[4] := '';
  Info.Example[5] := '';
end;

end.
