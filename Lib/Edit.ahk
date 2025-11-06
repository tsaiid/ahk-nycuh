/*
    v3.0 changes.

    To Do:
    Possible script breakers:
    ### Some/All items already moved to the forum document

    Bug fixes, updates, enhancements, etc.:
    ### Some/All items already moved to the forum document

    New functions (main library):
    ### Some/All items already moved to the forum document

    New Add-on functions:
    ### Some/All items already moved to the forum document
*/
/*
Group: Main Library

    This is the main Edit library which contains standard functions, helper
    functions, and internal functions.  All functions in this library will have
    names that begin with "Edit_".  See the _Edit.Doc.ahk document for more
    information.
*/
;------------------------------
;
; Function: Edit_ActivateParent
;
; Description:
;
;   Activate (makes foremost) the parent window of the Edit control if needed.
;   If the window is minimized, it is automatically restored prior to being
;   activated.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   This function only actives the parent window of the Edit control.  It does
;   not give focus to the Edit control.  If needed, call <Edit_SetFocus> which
;   has the option to activate the parent window before giving focus to the
;   Edit control.
;
;-------------------------------------------------------------------------------
Edit_ActivateParent(hEdit)
    {
    ;-- Get the handle to the parent window
    hParent:=DllCall("GetParent","UPtr",hEdit,"UPtr")

    ;-- Activate if needed
    IfWinNotActive ahk_id %hParent%
        {
        WinActivate ahk_id %hParent%

        ;-- Still not active? (rare)
        IfWinNotActive ahk_id %hParent%
            {
            ;-- Give the window an additional 250 ms to activate
            WinWaitActive ahk_id %hParent%,,0.25
            if ErrorLevel
                Return False
            }
        }

    Return True
    }

;------------------------------
;
; Function: Edit_BuildOO
;
; Description:
;
;   Build an options object.
;
; Type:
;
;   Internal function.  Subject to change.  Do not use.
;
; Parameters:
;
;   p_Options - An associative array of options or a string of zero or more
;       space-delimited options.  The options in this parameter have the highest
;       precedence.  If the parameter contains a string of options, the options
;       are processed from left (lowest precedence) to right (highest
;       precedence).
;
;   p_Default1 - [Optional] An associative array of options or a string of zero
;       or more space-delimited options.  The options in this parameter have the
;       2nd highest precedence (compared to p_Options).  If the parameter
;       contains a string of options, the options are processed from left
;       (lowest precedence) to right (highest precedence).
;
;   p_Default2 - [Optional] An associative array of options or a string of zero
;       or more space-delimited options.  The options in this parameter have the
;       lowest precedence (compared to p_Options and p_Default1).  If the
;       parameter contains a string of options, the options are processed from
;       left (lowest precedence) to right (highest precedence).
;
;   p_NVTable - [Optional] See the *Name and Value* section for more
;       information.
;
; Returns:
;
;   An object that contains an associative array of options.  The key is the
;   option name.
;
; Options:
;
;   There are two types options: Boolean and Assignment.
;
;   *Boolean*
;
;   Boolean options should only be tested as TRUE or FALSE.  Ex: "if o.Print".
;   A boolean option tests as FALSE by default so in most cases, simply not
;   defining a boolean option will give the same results as explicitly disabling
;   the option, i.e., setting the option value to FALSE.  However, if a boolean
;   option is enabled by default (i.e., included in the p_Default1 or p_Default2
;   parameters), the only way to disable it is to disable it explicitly.
;
;   *Assignment*
;
;   Assignment options are null by default so in most cases, simply not defining
;   an assignment option will give the same results as explicitly assigning a
;   null value to an assignment option.  However, if an assignment option is
;   specified by default (i.e., included in the p_Default1 or p_Default2
;   parameters), the only way to remove the value is to set the value to null
;   explicitly.
;
;   *Definition*
;
;   Options are defined as either a space-delimited list of strings or an
;   associative array.
;
;   *String Options*
;
;   String options are zero or more space-delimited list of string values.  For
;   example, "Print Title=aBc".
;
;   To enable a boolean string option, the option name can be specified as-is
;   (Ex: "Print") or it can be preceded with the "+" character.  Ex: "+Print".
;   To disable a boolean string option, precede the option name with the "-"
;   character.  Ex: "-Print".
;
;   Assignment string options are composed of three components: 1) The option
;   name, 2) the "=" character, and then 3) the option value.  Ex: Color=Red.
;   If the option value contains one or more space characters, enclose the
;   entire value in double quote characters.  Ex: Title="My Title".  If one or
;   more double quote characters are to be included as part of the option value,
;   use two double quote characters to specify that a single double quote
;   character is to be used.  Ex: Title="I did it ""My"" Way".
;
;   String options are processed in the order they are defined, i.e. from left
;   to right.
;
;   *Associative Array*
;
;   Options defined in an associative array are done a bit differently.
;
;   For a boolean option, set the key to the option name and set the value to
;   TRUE to enable or set to FALSE to disable.  For example:
;   Options:={AutoScroll:True,Print:False}.  Note, a boolean option is disabled
;   by default so the only reason define a boolean option and set it as disabled
;   is to override a default value.
;
;   For an assignment option, specify the key as the option name and the value
;   as an integer or string value.  For example:
;   Options:={NumberOfCopies:9,RDelim:"`r`n"}.  There is no way to disable an
;   assignment option per se.  It's up each program to determine whether an
;   assignment option should be used or not.  In many cases, setting the value
;   of a string assignment option to null will effectively disable the option.
;   For example, Options:={Title:""}.  For an assignment option that contains
;   number, it depends on the option.  For some options, setting the value to 0
;   will do it (Ex: {NumberOfCopies:0}).  Sometimes, setting the value to null
;   (Ex: {NumberOfCopies:""}) is more correct.  It's up to the developer to
;   determine the correct value.
;
; Name and Value Options:
;
;   AutoHotkey often supports options that combine the option name and option
;   value into a single string value.  For example, "Icon32".  In this example,
;   "Icon" is the option name and 23 is the option value.  So that the options
;   can be backwards compatible with AutoHotkey commands and/or functions, some
;   of the functions in this library also support these types of options.
;
;   The p_NVTable parameter is used to identify the "name and value" options
;   for a particular request.  If specified, the p_NVTable parameter must
;   contain an associative array.  Every key in the array is an option name.
;   The value for each key is an AutoHotkey data type that is used to verify the
;   integrity of the option value.  The data type value can also be set to null
;   (Ex: {OptionName:""}) to indicate that any value, including null can be
;   used.  In addition, the following custom types are supported:
;
;       {Null} - An value, including null, can be used.  Example of use:
;           NVOptions:={Title:""}
;
;       NoNull - An value, except null, can be used.  Example of use:
;           NVOptions:={Title:"NoNull"}
;   .
;
;   *Example of Use*
;
;   The following is an example of how this parameter could be used.  For this
;   example, the developer specifies the following list of options:
;
;       (start code)
;       Icon32 IconRight Col4 TitleMyTitle
;       (end)
;
;   To identify and extract the "name and value" options, the developer assigns
;   the following associative array to the p_NVTable parameter:
;
;       (start code)
;       {Icon:"Integer",Col,"Integer",Title:""}
;       (end)
;
;   The process has three steps.  In the step 1, the options are processed
;   normally, i.e. "name and value" options are not considered.  In the first
;   step, the function will create the following option values:
;
;       (start code)
;       Key             Value
;       ---             -----
;       Icon32          1
;       IconRight       1
;       Col4            1
;       TitleMyTitle    1
;       (end)
;
;   In step 2, the all options are compared against the associative array that
;   was passed in the p_NVTable parameter.  Each option key is compared with
;   each key in the array passed in the p_NVTable parameter.  If the first
;   part of the option key matches a key in p_NVTable parameter, then the
;   rest of the option key is compared against the data type specified.  If
;   there is a match, a new option is created.  In step 2, the following new
;   option/values are created:
;
;       (start code)
;       Key             Value
;       ---             -----
;       Icon            23
;       Col             4
;       Title           MyTitle
;       (end)
;
;   Note that the "IconRight" option is identified as a possible "name and
;   value" option but because the value after "Icon" string is not an integer,
;   it is skipped.
;
;   In step 3, the original "name and value" options are deleted and so the
;   final results are the following:
;
;       (start code)
;       Key             Value
;       ---             -----
;       Icon            23
;       IconRight       1
;       Col             4
;       Title           MyTitle
;       (end)
;
;   *Issues and Considerations*
;
;   This method works in most cases, but the design does not accommodate all
;   situations.  For example, if "Icon32" and "Icon" are both set, then the
;   "Icon" option will be treated the same as an boolean option and assigned a
;   TRUE value.  In this example, this would be same as if a "Icon1" option were
;   specified.  In most cases, this would not be a desired result.
;
;   The "NoNull" data type doesn't stop the developer from entering an option
;   name without a value (Ex: "Title"), it just stops the program from setting
;   the option value to null.  Instead, the option is treated as a boolean
;   option and the value is set to TRUE (1).  In most cases, this is not the
;   desired result.  Hint: The "NoNull" data type should not be used in most
;   cases.  Instead, the null type (i.e. the array value is set to "") should be
;   set and the developer should test to see of the option value is null
;   after the function returns.
;
;   Please note that the associative array's value must contain a valid
;   AutoHotkey data type (Ex: "Integer") or a custom alternative value that is
;   documented in this function.  Any other value may cause the script to fail.
;   Be sure to test thoroughly.
;
;-------------------------------------------------------------------------------
Edit_BuildOO(p_Options,p_Default1:="",p_Default2:="",p_NVTable:="")
    {
    Static s_SpecialChars:="¢¤¥¦§©ª«®µ¶"
                ;-- A list of characters that are unlikely to be in the
                ;   parameter values or in the data files.

    ;[==============]
    ;[  Initialize  ]
    ;[==============]
    o:=Object()
    o.SetCapacity(25)

    ;-- Copy or overwrite options if needed
    if IsObject(p_Default2)
        {
        For Key,Value in p_Default2
            o[Key]:=Value

        p_Default2:=""
        }

    if IsObject(p_Default1)
        {
        For Key,Value in p_Default1
            o[Key]:=Value

        p_Default1:=""
        }

    if IsObject(p_Options)
        {
        For Key,Value in p_Options
            o[Key]:=Value

        p_Options:=""
        }

    ;-- Build Options
    ;   Note: A trailing space is always a trailing delimiter.  This simplifies
    ;   the pattern necessary to extract the values from this parameter.
    Options:=p_Default2 . A_Space . p_Default1 . A_Space . p_Options . A_Space

    ;[========]
    ;[  Prep  ]
    ;[========]
    ;-- Find a single character that can be used to replace two double quote
    ;   characters
    Loop Parse,s_SpecialChars
        {
        if Options Contains %A_LoopField%
            Continue

        DQReplacementChar:=A_LoopField
        Break
        }

    ;-- Replace all occurrences of two double quote characters with a single
    ;   unique character.
    StringReplace Options,Options,"",%DQReplacementChar%,All

    ;[===================]
    ;[  Extract/Process  ]
    ;[===================]
    ;   Note: If undefined, assignment options are assumed to be null and
    ;   boolean options are assumed to be FALSE (not enabled)
    StartPos:=1
    Loop
        {
        AssignmentOption:=False
        OptionValue     :=""

        ;-- Find the next option.  The result is loaded to OptionName.
        ;   Note: If the trailing character in OptionName is a "=" character,
        ;   the option is an assignment option and the option value follows.  If
        ;   the trailing character in OptionName is a space, the option is a
        ;   standard boolean option.
        if not FoundPos:=RegExMatch(Options,"i)[+-]?\S+?[ =]",OptionName,StartPos)
            Break

        ;-- Update start position
        StartPos:=FoundPos+StrLen(OptionName)

        ;-- Enabled?
        OptionEnabled:=True  ;-- The default
        if (SubStr(OptionName,1,1)="-")
            OptionEnabled:=False

        ;-- If needed, remove leading +/- character
        if (SubStr(OptionName,1,1)="+" or SubStr(OptionName,1,1)="-")
            StringTrimLeft OptionName,OptionName,1

        ;---------------------
        ;-- Assignment option
        ;---------------------
        if (SubStr(OptionName,0)="=")
            {
            AssignmentOption:=True

            ;-- Remove trailing assignment (i.e. "=") character
            StringTrimRight OptionName,OptionName,1

            ;-- Double quoted value
            if (SubStr(Options,StartPos,1)="""")
                {
                ;-- Find quoted value.  The result is loaded to OptionValue.
                ;   Note: The result includes the DQ characters.
                if not FoundPos:=RegExMatch(Options,"is)"".*?""",OptionValue,StartPos)
                    Continue

                ;-- Update start position
                StartPos:=FoundPos+StrLen(OptionValue)

                ;-- Remove leading and trailing DQ characters
                OptionValue:=SubStr(OptionValue,2,-1)

                ;-- Convert any double quote replacement characters with a
                ;   single double quote character
                StringReplace
                    ,OptionValue
                    ,OptionValue
                    ,%DQReplacementChar%
                    ,"
                    ,All
                }
             else
                {
                ;-- Find unquoted value.  The result is loaded to OptionValue.
                if not FoundPos:=RegExMatch(Options,"is).*? ",OptionValue,StartPos)
                    Continue

                ;-- Update start position
                StartPos:=FoundPos+StrLen(OptionValue)

                ;-- Remove trailing delimiter (i.e. space) character
                StringTrimRight OptionValue,OptionValue,1

                ;-- If the value is a single DQ replacement character, i.e. the
                ;   replacement for two double quote characters, assume the
                ;   developer intended to assign a null value to the option or
                ;   (the more likely reason), the resulting value of a variable
                ;   used in the Options parameter was null.   Ex:
                ;   Title="%$Title%".  The developer surrounded the variable in
                ;   double quote characters just in case the value included one
                ;   or more space characters.
                if (OptionValue=DQReplacementChar)
                    OptionValue:=""
                }
            }
         else
            ;------------------
            ;-- Boolean option
            ;------------------
            ;-- Remove trailing space character from option name
            StringTrimRight OptionName,OptionName,1

        ;----------------------
        ;-- Valid option name?
        ;----------------------
        ValidOptionName:=True
        Loop Parse,OptionName
            {
            if A_LoopField is AlNum
                Continue

            if A_LoopField in #,_,@,$  ;,?,[,]
                Continue

            ValidOptionName:=False
            Break
            }

        if not ValidOptionName  ;-- Debugging condition only
            {
            outputdebug,
               (ltrim join`s
                Function: %A_ThisFunc% - Invalid option name: %OptionName%
               )

            Continue
            }

        ;--------------------------
        ;-- Create option variable
        ;--    and assign value
        ;--------------------------
        if AssignmentOption
            o[OptionName]:=OptionEnabled ? OptionValue:""
         else
            o[OptionName]:=OptionEnabled ? True:False
        }

    ;[============================]
    ;[  "Name and Value" options  ]
    ;[============================]
    if IsObject(p_NVTable)
        {
        NVOptions:=Object()
        NVOptions.SetCapacity(10)
        KeysToDelete:=Array()
        KeysToDelete.SetCapacity(10)
        For OptionKey,OptionValue in o
            {
            For NVKey,NVType in p_NVTable
                {
                if (InStr(OptionKey,NVKey)=1)
                    {
                    NVValue:=SubStr(OptionKey,StrLen(NVKey)+1)

                    ;-- Test value
                    AddOption:=False
                    if NVType is Space
                        AddOption:=True
                     else if (NVType="NoNull")
                        {
                        if StrLen(NVValue)
                            AddOption:=True
                        }
                     else if NVValue is %NVType%
                        AddOption:=True

                    if AddOption
                        {
                        NVOptions[NVKey]:=NVValue
                        KeysToDelete.Push(OptionKey)
                        }
                    }
                }
            }

        ;-- Delete old keys
        ;   Note: Keys that will be replaced when the new "name and value"
        ;   options are added are not deleted.
        For Index,OptionKey in KeysToDelete
            if not NVOptions.HasKey(OptionKey)
                o.Delete(OptionKey)

        ;-- Add/Replace "name and value" options
        For Key,Value in NVOptions
            o[Key]:=Value
        }

    ;[========]
    ;[  Post  ]
    ;[========]
    ;-- Free any left over capacity and return object
    o.SetCapacity(0)
    Return o
    }

;------------------------------
;
; Function: Edit_CanUndo
;
; Description:
;
;   Return TRUE if there are any actions in the Edit control's undo queue,
;   otherwise FALSE.
;
;-------------------------------------------------------------------------------
Edit_CanUndo(hEdit)
    {
    Static EM_CANUNDO:=0xC6
    SendMessage EM_CANUNDO,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_CharFromPos
;
; Description:
;
;   Get information about the character and the line closest to a specified
;   point in the client area of the Edit control.
;
; Parameters:
;
;   X, Y - The coordinates of a point in the Edit control's client area relative
;       to the upper-left corner of the client area.
;
;   r_CharPos - [Output, Optional] See the *Character Position* section for more
;       information.
;
;   r_LineIndex - [Output, Optional] See the *Line Index* section for more
;       information.
;
; Returns:
;
;   The value of the r_CharPos output variable.  See the description above for
;   more information.
;
; Character Position:
;
;   The r_CharPos parameter is an optional output ByRef parameter.  If
;   specified, it must be set to a variable name.  Ex: CharPos.
;
;   This output variable contains the zero-based index of the character nearest
;   the specified point (X and Y parameters).  This index is relative to the
;   beginning of the control, not the beginning of the line.  If the specified
;   point is beyond the last character in the Edit control, the index indicates
;   the last character in the control or zero (0) if the Edit control is empty.
;   See the *Remarks* section for more information.
;
; Line Index:
;
;   The r_LineIndex parameter is an optional output ByRef parameter.  If
;   specified, it must be set to a variable name.  Ex: LineIndex.
;
;   This output variable contains the zero-based index of the line that contains
;   the character.  For a single-line Edit control, this value is zero (0).
;
;   [Horizontal] If the specified point is beyond the last character in a line,
;   the index indicates that same line of the last character of the line.
;
;   [Vertical] If the specified point is beyond the last character in the Edit
;   control, the index indicates the last line in the control or zero (0) if the
;   Edit control is empty.
;
;   See the *Remarks* section for more information.
;
; Remarks:
;
;   If the specified point is outside the bounds of the Edit control, the
;   return value and all output variables (r_CharPos and r_LineIndex) are set to
;   -1.  Scroll bars (when showing) are considered to be outside the bounds of
;   the Edit control.
;
; Observations:
;
;   When the Y coordinate (Y parameter) is on the very top pixel of the
;   formatting rectangle (usually 1), this function will return the line index
;   to the line immediately before the first visible line instead of the index
;   to the first visible line.  This only occurs if the first visible line is
;   not the first line of text in the Edit control.  If needed, adding 1 to the
;   Y parameter in this case will work around this idiosyncrasy.
;
;   When the Y coordinate (Y parameter) is on the very bottom pixel of the
;   formatting rectangle, this function may return the line index to the line
;   immediately after the last visible line instead of the index to the last
;   visible line.  This idiosyncrasy only occurs if the last visible line ends
;   at the very bottom of the formatting rectangle.  This usually does not occur
;   unless the height of the Edit control was calculated to fit an exact number
;   of lines.  [Note] The GUI "r" option (Ex: "r10") will sometimes generate the
;   height to fit an exact number of lines but the "r" calculation is flawed and
;   it will often generate a inaccurate height. [/Note] If needed, subtracting 1
;   to the Y parameter in this case will work around this idiosyncrasy.
;
;-------------------------------------------------------------------------------
Edit_CharFromPos(hEdit,X,Y,ByRef r_CharPos:="",ByRef r_LineIndex:="")
    {
    Static Dummy69154302
          ,EM_CHARFROMPOS        :=0xD7
          ,EM_GETFIRSTVISIBLELINE:=0xCE
          ,EM_LINEINDEX          :=0xBB

    ;-- Get the character position from the specified coordinates
    SendMessage EM_CHARFROMPOS,0,(Y<<16)|X,,ahk_id %hEdit%

    ;-- Out of bounds?
    if (ErrorLevel<<32>>32=-1)
        {
        r_CharPos  :=-1
        r_LineIndex:=-1
        Return -1
        }

    ;-- Extract values (UShort)
    r_CharPos  :=ErrorLevel&0xFFFF  ;-- LOWORD
    r_LineIndex:=ErrorLevel>>16     ;-- HIWORD

    ;-- Convert from UShort to UInt using known UInt values as reference
    SendMessage EM_GETFIRSTVISIBLELINE,0,0,,ahk_id %hEdit%
    FirstLine:=ErrorLevel-1
    if (FirstLine>r_LineIndex)
        r_LineIndex:=r_LineIndex+(65536*Floor((FirstLine+(65535-r_LineIndex))/65536))

    SendMessage EM_LINEINDEX,FirstLine<0 ? 0:FirstLine,0,,ahk_id %hEdit%
    FirstCharPos:=ErrorLevel
    if (FirstCharPos>r_CharPos)
        r_CharPos:=r_CharPos+(65536*Floor((FirstCharPos+(65535-r_CharPos))/65536))

    Return r_CharPos
    }

;------------------------------
;
; Function: Edit_Clear
;
; Description:
;
;   Clear (delete) the current selection, if any.
;
; Remarks:
;
;   If text was selected, Undo can be used to reverse this action.
;
;-------------------------------------------------------------------------------
Edit_Clear(hEdit)
    {
    Static WM_CLEAR:=0x303
    SendMessage WM_CLEAR,0,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_Convert2DOS
;
; Description:
;
;   Convert the Unix (LF), DOS/Unix mix (CR+LF and LF), and Mac (OS9 and
;   earlier) (CR) end-of-line formats to the DOS/Windows end-of-line format
;   (CR+LF).
;
;-------------------------------------------------------------------------------
Edit_Convert2DOS(p_Text)
    {
    StringReplace p_Text,p_Text,`r`n,`n,All             ;-- Convert DOS to Unix
    StringReplace p_Text,p_Text,`r,`n,All               ;-- Convert Mac to Unix
    StringReplace p_Text,p_Text,`n,`r`n,All             ;-- Convert Unix to DOS
    Return p_Text
    }

;------------------------------
;
; Function: Edit_Convert2Mac
;
; Description:
;
;   Convert the DOS/Windows (CR+LF), DOS/Unix mix (CR+LF and LF), and Unix (LF)
;   end-of-line formats to the Mac (OS 9 and earlier) end-of-line format (CR).
;
;-------------------------------------------------------------------------------
Edit_Convert2Mac(p_Text)
    {
    StringReplace p_Text,p_Text,`r`n,`r,All             ;-- Convert DOS to Mac
    StringReplace p_Text,p_Text,`n,`r,All               ;-- Convert Unix to Mac
    if StrLen(p_Text)
        if (SubStr(p_Text,0)!="`r")
            p_Text.="`r"

    Return p_Text
    }

;------------------------------
;
; Function: Edit_Convert2Unix
;
; Description:
;
;   Convert the DOS/Windows (CR+LF), DOS/Unix mix (CR+LF and LF), and Mac (OS 9
;   and earlier) (CR) end-of-line formats to the Unix end-of-line format (LF).
;
;-------------------------------------------------------------------------------
Edit_Convert2Unix(p_Text)
    {
    StringReplace p_Text,p_Text,`r`n,`n,All             ;-- Convert DOS to Unix
    StringReplace p_Text,p_Text,`r,`n,All               ;-- Convert Mac to Unix
    if StrLen(p_Text)
        if (SubStr(p_Text,0)!="`n")
            p_Text.="`n"

    Return p_Text
    }

;------------------------------
;
; Function: Edit_Copy
;
; Description:
;
;   Copy the current selection (if any) to the clipboard in CF_TEXT format.
;
;-------------------------------------------------------------------------------
Edit_Copy(hEdit)
    {
    Static WM_COPY:=0x301
    SendMessage WM_COPY,0,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_Cut
;
; Description:
;
;   Delete the current selection, if any, and copy the deleted text to the
;   clipboard in CF_TEXT format.
;
;-------------------------------------------------------------------------------
Edit_Cut(hEdit)
    {
    Static WM_CUT:=0x300
    SendMessage WM_CUT,0,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_Disable
;
; Description:
;
;   Disable ("gray out") the Edit control.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   For AutoHotkey GUIs, the *<GUIControlGet at
;   https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command can be
;   used instead.  Ex: GUIControl 24:Disable,%hEdit%
;
;-------------------------------------------------------------------------------
Edit_Disable(hEdit)
    {
    Control Disable,,,ahk_id %hEdit%
    Return ErrorLevel ? False:True
    }

;------------------------------
;
; Function: Edit_EmptyUndoBuffer
;
; Description:
;
;   Reset (clear) the undo flag of the Edit control.  The undo flag is set
;   whenever there is an operation in the Edit control can be undone.
;
;-------------------------------------------------------------------------------
Edit_EmptyUndoBuffer(hEdit)
    {
    Static EM_EMPTYUNDOBUFFER:=0xCD
    SendMessage EM_EMPTYUNDOBUFFER,0,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_Enable
;
; Description:
;
;   Enable the Edit control if it was previously disabled.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   For AutoHotkey GUIs, the *<GUIControlGet at
;   https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command can be
;   used instead.  Ex: GUIControl 12:Enable,%hEdit%
;
;-------------------------------------------------------------------------------
Edit_Enable(hEdit)
    {
    Control Enable,,,ahk_id %hEdit%
    Return ErrorLevel ? False:True
    }

;------------------------------
;
; Function: Edit_EnableZoom
;
; Description:
;
;   Enable or disable zooming for the Edit control.
;
; Parameters:
;
;   p_Enable - [Optional] If set to TRUE (the default) or if not specified,
;       zooming is enabled for the Edit control.  If set to FALSE, zooming is
;       disabled for the Edit control.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Requirements:
;
;   Windows 10+
;
; Calls To Other Functions:
;
; * <Edit_SetExtendedStyle>
;
;-------------------------------------------------------------------------------
Edit_EnableZoom(hEdit,p_Enable:=True)
    {
    Static ES_EX_ZOOMABLE:=0x10
    Return Edit_SetExtendedStyle(hEdit,ES_EX_ZOOMABLE,p_Enable ? ES_EX_ZOOMABLE:0)
    }

;------------------------------
;
; Function: Edit_FindText
;
; Description:
;
;   Find text in the Edit control.
;
; Parameters:
;
;   p_SearchText - The text to search for.  Ex: "cat".
;
;   p_Min, p_Max -  See the *Search Range* section for more information.
;
;   p_Options - [Optional] See the *Options* section for more information.
;
;   r_RegExOutput - [Output, Optional] See the *RegEx Output* section for more
;       information.
;
; Returns:
;
;   The zero-based character index of the first character of the match (this can
;   be 0) or -1 if no match is found.
;
; Calls To Other Functions:
;
; * <Edit_GetText>
; * <Edit_GetTextRange>
;
; Options:
;
;   The p_Options parameter is used to specify zero or more search options.  The
;   following option names can be used:
;
;       MatchCase - The search is case sensitive.  This option is ignored if the
;           "RegEx" option is also specified.
;
;       RegEx - Regular expression search.  If this option is included, the
;           p_SearchText parameter should contain a regular expression.
;
;       Reset - [Advanced feature] Clear the saved text and memory created by
;           the "Static" option so that the next use of the "Static" option will
;           get the text directly from the Edit control.  To clear the saved
;           text and memory without performing a search, use the following
;           syntax:
;
;               Edit_FindText("","",0,0,"Reset")
;
;           A reset can also be performed by calling <Edit_FindTextReset>.
;
;       Static - [Advanced feature] The text collected from the Edit control
;           remains in memory is used to satisfy the search request.  The text
;           remains in memory until the "Reset" option is used or until the
;           "Static" option is not used.
;
;           _Advantages_: Search time is reduced 10 to 80 percent (or more)
;           depending on the amount of the text in the Edit control.  Please
;           note that there is no speed increase on the first use of the
;           "Static" option.
;
;           _Disadvantages_: While this option is in use, any changes in the
;           Edit control are not reflected in the search.
;
;           Hint: Do not use this option unless performing multiple search
;           requests on the Edit control that will not be modified while
;           searching.
;
;       WholeWord - The specified search text must match whole words.  This
;           option is ignored if the "RegEx" option is also specified.
;
;   If more than one option is specified, the options must be delimited by a
;   space character.  Ex: "RegEx Static".
;
; Search Range:
;
;   The p_Min and p_Max parameters are used to specify the search range within
;   the Edit control.
;
;   Set the p_Min parameter to the zero-based character index of the first
;   character in the search range.  The default is 0 indicating the first
;   character in the Edit control.  Set the p_Max parameter to the zero-based
;   character index immediately following the last character in the search
;   range.  To search to the end of the text, set p_Max to -1 (the default).
;
;   A couple of examples... To search the first 20 characters of the text, set
;   p_Min to 0 and set p_Max to 20.  To search all the text in the Edit control,
;   set p_Min to 0 (the default) and p_Max to -1 (the default).
;
;   To search backward, the roles and descriptions of the p_Min and p_Max
;   parameters are reversed.  For example, to search the first 30 characters of
;   the control in reverse, set p_Min to 30 and p_Max to 0.  Please note that
;   when using the "RegEx" option (p_Options parameter), some regular expression
;   patterns work fine when performing a backward search while others do not
;   work as expected. Your results will vary.
;
; RegEx Output:
;
;   The r_RegExOutput parameter is an optional output ByRef parameter.  If
;   specified, it must be set to a variable name.  Ex: MyRegExOutput.
;
;   If the p_Options parameter contains "RegEx" or "WholeWord" and if the search
;   was successful, this variable contains the part of the source text that
;   matched the regular expression pattern.  Otherwise, this variable will be
;   null.
;
;   Please note that for a whole word search, the value in this variable will be
;   the same as p_SearchText parameter in most cases.  The case of the text
;   may be different if the "MatchCase" option is not included.
;
; Remarks:
;
;   Most search requests use a small or an "average" amount of computer
;   resources.  However, if performing a complex search (Ex: regular
;   expression), searching a large amount of text, calling this function in a
;   loop, or repeatedly calling this function via a keyboard shortcut (Ex: Find
;   Next), this function or the use of the function can become resource
;   intensive.  See <Function Performance> for information on how to improve the
;   function performance.
;
; Programming Notes:
;
;   Searching using a regular expression can produce results that have a dynamic
;   number of characters.  For this reason, searching for the "next" pattern
;   (forward or backward) may produce different results from developer to
;   developer depending on how the values of p_Min and p_Max are determined.
;
;-------------------------------------------------------------------------------
Edit_FindText(hEdit,p_SearchText,p_Min:=0,p_Max:=-1,p_Options:="",ByRef r_RegExOut:="")
    {
    Static Dummy62903754
          ,s_Text

          ;-- Messages
          ,WM_GETTEXTLENGTH:=0xE

    ;-- Initialize
    r_RegExOut:=""
    if InStr(A_Space . p_Options . A_Space," Reset ")
        s_Text:=""

    ;-- Bounce and return -1 if there is nothing to search
    if not StrLen(p_SearchText)
        Return -1

    SendMessage WM_GETTEXTLENGTH,0,0,,ahk_id %hEdit%
    MaxLen:=ErrorLevel
    if (MaxLen=0)
        Return -1

    ;-- Parameters
    if (p_Min<0 or p_Max>MaxLen)
        p_Min:=MaxLen

    if (p_Max<0 or p_Max>MaxLen)
        p_Max:=MaxLen

    ;-- Anything to search?
    if (p_Min=p_Max)
        Return -1

    ;-- Get text
    if InStr(A_Space . p_Options . A_Space," Static ")
        {
        if not StrLen(s_Text)
            s_Text:=Edit_GetText(hEdit)

        Text:=SubStr(s_Text,(p_Max>p_Min) ? p_Min+1:p_Max+1,(p_Max>p_Min) ? p_Max:p_Min)
        }
     else
        {
        s_Text:=""
        Text:=Edit_GetTextRange(hEdit,(p_Max>p_Min) ? p_Min:p_Max,(p_Max>p_Min) ? p_Max:p_Min)
        }

    ;-- Look for it
    MatchCase:=InStr(A_Space . p_Options . A_Space," MatchCase ") ? True:False
    RegEx    :=InStr(A_Space . p_Options . A_Space," RegEx ") ? True:False
    WholeWord:=InStr(A_Space . p_Options . A_Space," WholeWord ") ? True:False
    if not (RegEx or WholeWord)
        FoundPos:=InStr(Text,p_SearchText,MatchCase,(p_Max>p_Min) ? 1:0)-1
     else  ;-- RegEx or Whole Word
        {
        if RegEx
            p_SearchText:=RegExReplace(p_SearchText,"^P\)?","",1)   ;-- Remove P or P)
         else  ;-- WholeWord
            p_SearchText:=(MatchCase ? "":"i)") . "\b" . p_SearchText . "\b"

        if (p_Max>p_Min)  ;-- Search forward
            {
            FoundPos:=RegExMatch(Text,p_SearchText,r_RegExOut,1)-1
            if ErrorLevel
                {
                outputdebug,
                   (ltrim join`s
                    Function: %A_ThisFunc% - RegExMatch error.
                    ErrorLevel: %ErrorLevel%
                   )

                FoundPos:=-1
                }
            }
         else  ;-- Search backward
            {
            ;-- Programming notes:
            ;
            ;    *  The first search begins from the user-defined minimum
            ;       position.  This will establish the true minimum position to
            ;       begin search calculations.  If nothing is found, no
            ;       additional searching is necessary.
            ;
            ;    *  The RE_MinPos, RE_MaxPos, and RE_StartPos variables contain
            ;       1-based values.
            ;
            RE_MinPos     :=1
            RE_MaxPos     :=StrLen(Text)
            RE_StartPos   :=RE_MinPos
            Saved_FoundPos:=-1
            Saved_RegExOut:=""
            Loop
                {
                ;-- Positional search.  Last found match (if any) wins
                FoundPos:=RegExMatch(Text,p_SearchText,r_RegExOut,RE_StartPos)-1
                if ErrorLevel
                    {
                    outputdebug,
                       (ltrim join`s
                        Function: %A_ThisFunc% - RegExMatch error.
                        ErrorLevel: %ErrorLevel%
                       )

                    FoundPos:=-1
                    Break
                    }

                ;-- If found, update saved and RE_MinPos, else update RE_MaxPos
                if (FoundPos>-1)
                    {
                    Saved_FoundPos:=FoundPos
                    Saved_RegExOut:=r_RegExOut
                    RE_MinPos     :=FoundPos+2
                    }
                else
                    RE_MaxPos:=RE_StartPos-1

                ;-- Are we done?
                if (RE_MinPos>RE_MaxPos or RE_MinPos>StrLen(Text))
                    {
                    FoundPos  :=Saved_FoundPos
                    r_RegExOut:=Saved_RegExOut
                    Break
                    }

                ;-- Calculate new start position
                RE_StartPos:=RE_MinPos+Floor((RE_MaxPos-RE_MinPos)/2)
                }
            }
        }

    ;-- Adjust FoundPos
    if (FoundPos>-1)
        FoundPos+=(p_Max>p_Min) ? p_Min:p_Max

    Return FoundPos
    }

;------------------------------
;
; Function: Edit_FindTextReset
;
; Description:
;
;   Clear the saved text created by the "Static" flag.
;
; Calls To Other Functions:
;
; * <Edit_FindText>
;
;-------------------------------------------------------------------------------
Edit_FindTextReset()
    {
    Edit_FindText("","",0,0,"Reset")
    }

;------------------------------
;
; Function: Edit_GetActiveHandles
;
; Description:
;
;   Find the handles for the active control and active window.
;
; Type:
;
;   Helper function.
;
; Parameters:
;
;   hEdit - [Output, Optional] If specified, this variable contains the handle
;       of the active Edit control.  The value is zero (0) if the active
;       control is not an Edit control.
;
;   hWindow - [Output, Optional] If specified, this variable contains the handle
;       of the active window.
;
;   p_MsgBox - [Optional] If set to TRUE, an error MsgBox is displayed if the
;       active control is not an Edit control.
;
; Returns:
;
;   The handle of the active Edit control (tests as TRUE) or FALSE (0) if the
;   active control is not an Edit control.
;
;-------------------------------------------------------------------------------
Edit_GetActiveHandles(ByRef hEdit:="",ByRef hWindow:="",p_MsgBox:=False)
    {
    WinGet hWindow,ID,A
    ControlGetFocus ControlID,A
    if (SubStr(ControlID,1,4)="Edit")
        {
        ControlGet hEdit,hWnd,,%ControlID%,A
        Return hEdit
        }

    if p_MsgBox
        MsgBox
            ,0x40010  ;-- 0x0 (OK button) + 0x10 ("Error" icon) + 0x40000 (AOT)
            ,Error
            ,This request cannot be performed on this control.

    Return False
    }

;------------------------------
;
; Function: Edit_GetCaretIndex
;
; Description:
;
;   Get the position of the caret in the Edit control.
;
; Returns:
;
;   The zero-based index value of the position of the caret.  FALSE (0) is
;   also returned if there is an error.
;
; Requirements:
;
;   Windows 10+
;
; Observations:
;
;   The caret index is the same as the character position with exceptions.  If
;   the Edit control is empty, the caret index is 0.  There is no valid
;   character position in this case.  If the caret is immediately after the last
;   character in the Edit control, the caret index is the same as the character
;   position of the last character plus 1.  There is no valid character position
;   in this case.
;
;-------------------------------------------------------------------------------
Edit_GetCaretIndex(hEdit)
    {
    Static EM_GETCARETINDEX:=0x1512  ;-- ECM_FIRST+18
    SendMessage EM_GETCARETINDEX,0,0,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetComboBoxEdit
;
; Description:
;
;   Get the handle to the Edit control attached to a combo box.
;
; Parameters:
;
;   hCombo - The handle to a combo box control.
;
; Returns:
;
;   The handle to the Edit control attached to a combo box (tests as TRUE) if
;   successful, otherwise FALSE.
;
; Credit:
;
;   Code adapted from an example posted by *just me*.  The link to that post
;   has been lost.
;
;-------------------------------------------------------------------------------
Edit_GetComboBoxEdit(hCombo)
    {
    Static sizeofCOMBOBOXINFO:=A_PtrSize=8 ? 64:52

    ;-- Create and initialize COMBOBOXINFO structure
    VarSetCapacity(COMBOBOXINFO,sizeofCOMBOBOXINFO)
    NumPut(sizeofCOMBOBOXINFO,COMBOBOXINFO,0,"UInt")    ;-- cbSize

    ;-- Get ComboBox info
    if DllCall("GetComboBoxInfo","UPtr",hCombo,"UPtr",&COMBOBOXINFO)
        Return NumGet(COMBOBOXINFO,A_PtrSize=8 ? 48:44,"UPtr")
            ;-- hwndItem

    ;-- Error
    outputdebug,
       (ltrim join`s
        Function: %A_ThisFunc% - GetComboBoxInfo error.  Unable to get the Edit
        control handle attached to the following combo box handle: %hCombo%
       )

    Return False
    }

;------------------------------
;
; Function: Edit_GetCueBanner
;
; Description:
;
;   Get the text that is displayed as the textual cue, or tip, in the Edit
;   control.
;
; Parameters:
;
;   p_MaxSize - [Optional] The maximum number of characters including the
;       terminating null character.  The default is 1024.
;
; Returns:
;
;   The cue banner text from the specified Edit control.  The return value will
;   be null if a cue banner was not set or if there was an error.
;
; Remarks:
;
;   Single-line Edit control only.
;
;-------------------------------------------------------------------------------
Edit_GetCueBanner(hEdit,p_MaxSize:=1024)
    {
    Static EM_GETCUEBANNER:=0x1502  ;-- ECM_FIRST+2
    VarSetCapacity(wText,p_MaxSize*2)
    SendMessage EM_GETCUEBANNER,&wText,p_MaxSize,,ahk_id %hEdit%
    if ErrorLevel  ;-- Cue banner text found
        Return A_IsUnicode ? wText:StrGet(&wText,-1,"UTF-16")
    }

;------------------------------
;
; Function: Edit_GetEndOfLine
;
; Description:
;
;   Get the end-of-line character (as a flag) for the Edit control.
;
; Returns:
;
;   A flag that identifies the end-of-line character(s) used by the Edit
;   control.  See the "Edit Control End-Of-Line Flags" section in the
;   Edit_Constants.ahk document for a list of possible flag values.
;
;   Please note that if the end-of-line character was set to
;   EC_ENDOFLINE_DETECTFROMCONTENT, the detected end-of-line character is
;   returned.
;
; Requirements:
;
;   Windows 10+
;
;-------------------------------------------------------------------------------
Edit_GetEndOfLine(hEdit)
    {
    Static EM_GETENDOFLINE:=0x150D  ;-- ECM_FIRST+13
    SendMessage EM_GETENDOFLINE,0,0,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetExtendedStyle
;
; Description:
;
;   Get the extended style for the Edit control.
;
; Returns:
;
;   The extended style for the Edit control (can be 0).  FALSE (0) is also
;   returned if there is an error.
;
; Requirements:
;
;   Windows 10+
;
; Remarks:
;
;   See the "Edit Control Extended Styles" section in the Edit_Constants.ahk
;   document for a list of possible extended styles for the Edit control.
;
;   The extended styles for the Edit control have nothing to do with the
;   extended styles used with the CreateWindowEx or SetWindowLong functions.
;
;-------------------------------------------------------------------------------
Edit_GetExtendedStyle(hEdit)
    {
    Static EM_GETEXTENDEDSTYLE:=0x150B  ;-- ECM_FIRST+11
    SendMessage EM_GETEXTENDEDSTYLE,0,0,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetFirstVisibleLine
;
; Description:
;
;   Return the zero-based index of the uppermost visible line.
;
; Observations:
;
;   The Microsoft documentation states that for a single-line Edit control, the
;   return value is the zero-based index of the first visible character.  Based
;   upon observation and testing, this is not true.  The return value for
;   single-line Edit control is always 0 which is the line index for the first
;   and only line in a single-line Edit control.  This was tested on Windows 10.
;   At this writing, this has not been tested on any other version of Windows.
;
;-------------------------------------------------------------------------------
Edit_GetFirstVisibleLine(hEdit)
    {
    Static EM_GETFIRSTVISIBLELINE:=0xCE
    SendMessage EM_GETFIRSTVISIBLELINE,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetFont
;
; Description:
;
;   Get the font with which the Edit control is currently using to draw text.
;
; Returns:
;
;   The handle to the font used by the Edit control or 0 if the using the system
;   font (very rare).
;
; Remarks:
;
;   This function can be used to get the font of any control.  Just specify the
;   handle to the desired control as the first parameter.  Ex: Edit_GetFont(hLV)
;   where hLV is the handle to a ListView control.
;
;-------------------------------------------------------------------------------
Edit_GetFont(hEdit)
    {
    Static WM_GETFONT:=0x31
    SendMessage WM_GETFONT,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetLastVisibleLine
;
; Description:
;
;   Get the zero-based line index of the last (lowermost) visible line on the
;   Edit control.
;
; Parameters:
;
;   p_Threshold - [Optional] See the *Threshold* section for more information.
;
; Returns:
;
;   The zero-based index of the last visible line on the Edit control.
;
; Calls To Other Functions:
;
; * <Edit_CharFromPos>
;
; Threshold:
;
;   The optional p_Threshold parameter is used to identify the minimum amount of
;   vertical space (in pixels) that the last visible line must occupy before it
;   is identified as visible.  The function default is 4.  The minimum value
;   is 1.
;
;   Question: Should this parameter be changed depending on the font size?
;   Answer: Maybe.  The default value will cover most situations for small and
;   medium-sized fonts.  Setting to a larger value for larger fonts will
;   probably provide value.  However, if set to a value greater than the height
;   of characters in the font, one or more of the last visible lines will be
;   identified as invisible.
;
;   Alternatively, set to this parameter to "Ascent" to set this value to the
;   number of pixels that occupy the ascent portion of the characters in the
;   current font.  Ascent is the portion of the character cell above the
;   baseline.  If the last visible line is a least the height of the font's
;   ascent, many of the characters in the line will be completely visible and
;   the entire line will be readable in most cases.
;
;   Please note that this function works with the Zoom feature that is built
;   into the Edit control starting with Windows 10.  However, there are several
;   considerations.  As you would expect, the size of the font is increased or
;   decreased as the text in the Edit control is zoomed in or out.
;   Subsequently, the height of the characters will increase or decrease as the
;   text is zoomed in or out.  While the default or specified threshold for the
;   last visible line (p_Threshold parameter) may be adequate for small to
;   medium font sizes, it may not be adequate for larger font sizes.  The
;   threshold value will not change unless the developer adds additional code to
;   modify this parameter while zooming.  Also note that font attached to Edit
;   control does not change while the Edit control is zoomed and so the
;   threshold size set when the p_Threshold parameter is set to "Ascent" will
;   also not change while zooming.
;
; Remarks:
;
;   The return value for a single-line Edit control is 0 which indicates the
;   zero-based index of the one and only line on a single-line Edit control.
;   Exception: If the p_Threshold parameter is set to a value larger that the
;   height of characters of the font attached to the Edit control, then -1 is
;   returned.
;
;   Depending on the value of the p_Threshold parameter, partially visible lines
;   are considered to be visible.  See the *Threshold* section (above) and <Last
;   Visible Line> for more information.
;
;   To calculate the total number of visible lines, use the following
;   calculation.
;
;       (start code)
;       Edit_GetLastVisibleLine(hEdit) - Edit_GetFirstVisibleLine(hEdit) + 1
;       (end)
;
;-------------------------------------------------------------------------------
Edit_GetLastVisibleLine(hEdit,p_Threshold:="")
    {
    Static Dummy39278506
          ,HWND_DESKTOP:=0
          ,RECT
          ,TEXTMETRIC
          ,Dummy1:=VarSetCapacity(RECT,16)
          ,Dummy2:=VarSetCapacity(TEXTMETRIC,A_IsUnicode ? 60:56)
          ,s_DefaultThreshold:=4

          ;-- Messages
          ,EM_GETRECT:=0xB2
          ,WM_GETFONT:=0x31

    ;-- Parameters
    if p_Threshold is Space              ;-- Parameter default
        p_Threshold:=s_DefaultThreshold  ;-- Function default
     else if p_Threshold is Integer
        {
        if (p_Threshold<1)
            p_Threshold:=1  ;-- Minimum
        }
     else if InStr(p_Threshold,"Ascent")
        {
        ;-- Get the font attached to the Edit control
        SendMessage WM_GETFONT,0,0,,ahk_id %hEdit%
        hFont:=ErrorLevel

        ;-- Select the font into a desktop device context
        hDC:=DllCall("GetDC","UPtr",HWND_DESKTOP)
        DllCall("SaveDC","UPtr",hDC)
        DllCall("SelectObject","UPtr",hDC,"UPtr",hFont)

        ;-- Get the text metrics for the font
        DllCall("GetTextMetrics","UPtr",hDC,"UPtr",&TEXTMETRIC)

        ;-- Housekeeping
        DllCall("RestoreDC","UPtr",hDC,"Int",-1)
        DllCall("ReleaseDC","UPtr",HWND_DESKTOP,"UPtr",hDC)

        ;-- Extract ascent value
        p_Threshold:=NumGet(TEXTMETRIC,4,"Int")  ;-- tmAscent
        }
     else
        p_Threshold:=s_DefaultThreshold  ;-- Function default

    ;-- Get the Y coordinates of the bottom pixels of the formatting rectangle
    SendMessage EM_GETRECT,0,&RECT,,ahk_id %hEdit%
    LastVisibleY:=NumGet(RECT,12,"Int")-p_Threshold

    ;-- Get the last visible line
    Edit_CharFromPos(hEdit,1,LastVisibleY,Dummy,LastVisibleLine)
    Return LastVisibleLine
    }

;------------------------------
;
; Function: Edit_GetLimitText
;
; Description:
;
;   Return the current text limit for the Edit control.
;
; Remarks:
;
;   The maximum text length is 0x7FFFFFFE (2,147,483,646) characters for a
;   single-line Edit control and 0xFFFFFFFF (4,294,967,295) for a multiline Edit
;   control.  These values are returned if no limit has been set on the Edit
;   control.
;
;-------------------------------------------------------------------------------
Edit_GetLimitText(hEdit)
    {
    Static EM_GETLIMITTEXT:=0xD5
    SendMessage EM_GETLIMITTEXT,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetLine
;
; Description:
;
;   Get the text of a specific line from the Edit control.
;
; Parameters:
;
;   p_LineIndex - [Optional] Set to the zero-based index of the line to
;       retrieve.  Ex: 27.  If set  to -1 (the default) or if not specified, the
;       current line is retrieved.  This parameter is ignored if used on a
;       single-line Edit control.  In this case, the first and only line is
;       always retrieved.
;
;   p_Length - [Optional] Set to the length of the line of the text to be
;       extracted.  Ex: 15.  If set to a value less than the length of the line,
;       only that portion of the line is returned.  If set to the length of the
;       line or larger, the entire line is returned.  If set to -1 (the default)
;       or if not specified, the function will automatically determine the
;       length of the line.
;
; Returns:
;
;   The text of the specified line up to the length (p_Length) specified.  A
;   null string is returned if the line is empty/null.  For a multiline Edit
;   control, a null string is also returned if the specified line index is
;   greater than the number of lines in the Edit control.
;
; Programming Notes:
;
;   The EM_GETLINE message copies a line of text from the Edit control into a
;   buffer.  The size of the buffer is determined by the p_Length parameter.  To
;   avoid an "out of memory" exception that could cause AutoHotkey to crash, the
;   buffer size is limited to 65535 bytes which is the maximum size allowed by
;   the EM_LINELENGTH message.  For ANSI, this will allow for a maximum of
;   65,535 characters.  For Unicode, a maximum of 32,767 characters.  Since most
;   lines are less than 1,024 characters, a requirement for more than 32,767
;   characters (the Unicode maximum) should be very rare although it is
;   theoretically possible.  No, this has not been tested.
;
;-------------------------------------------------------------------------------
Edit_GetLine(hEdit,p_LineIndex:=-1,p_Length:=-1)
    {
    Static Dummy30246819

          ;-- Messages
          ,EM_GETLINE     :=0xC4
          ,EM_LINEFROMCHAR:=0xC9
          ,EM_LINEINDEX   :=0xBB
          ,EM_LINELENGTH  :=0xC1

    ;-- Parameters
    if (p_LineIndex<0)
        {
        ;-- Get the line index of the current line
        SendMessage EM_LINEFROMCHAR,-1,0,,ahk_id %hEdit%
        p_LineIndex:=ErrorLevel
        }

    ;-- If needed, determine the line length
    if (p_Length<0)
        {
        ;-- Convert the line index to character position
        SendMessage EM_LINEINDEX,p_LineIndex,0,,ahk_id %hEdit%
        CharPos:=ErrorLevel<<32>>32  ;-- Convert UInt to Int

        ;-- Return null string if CharPos is -1.  This occurs if p_LineIndex is
        ;   out of range.
        if (CharPos<0)
            Return

        ;-- Get the line length
        SendMessage EM_LINELENGTH,CharPos,0,,ahk_id %hEdit%
        p_Length:=ErrorLevel
        }

    ;-- Return null string if the length is 0
    ;   Note: p_Length can be 0 because the length for the line is 0 or because
    ;   the p_Length parameter was set 0.
    if (p_Length=0)
        Return

    ;------------
    ;
    ;   Note: At this point, p_LineIndex is not negative but it might contain
    ;   a value that is larger than the maximum line index.  The rest of the
    ;   function will respond correctly (return null string) if this occurs.
    ;
    ;---------------------------------------------------------------------------

    ;-- Create and initialize buffer
    ;   Programming note: nSize is the size of the buffer in bytes.  Since the
    ;   first WORD (aka UShort) of the buffer is set to the number TCHARs to be
    ;   retrieved, the minimum buffer size is 2 bytes.  Also, since the largest
    ;   value that fit in a UShort (2 byte) number is 65535, the buffer size is
    ;   automatically truncated to this size if needed.  This limitation will
    ;   keep the script from creating an unusually large buffer and avoid an
    ;   "out of memory" exception.
    ;
    nSize:=A_IsUnicode ? p_Length*2:p_Length=1 ? 2:p_Length
    if (nSize>65535)
        {
        nSize   :=65535
        p_Length:=A_IsUnicode ? 32767:65535
            ;-- The largest values supported by the EM_GETLINE message
        }

    VarSetCapacity(Text,nSize)
    NumPut(p_Length,Text,0,"UShort")

    ;-- Get line
    ;   Programming notes: When the EM_GETLINE message is sent, ErrorLevel
    ;   contains the number of TCHARs copied to the buffer.  ErrorLevel is zero
    ;   (0) if the line is null or if the line number specified is greater than
    ;   the number of lines in the Edit control.  If the p_Length parameter was
    ;   set by the developer, the Text variable may contain up to 2 "extra"
    ;   characters.  Using the SubStr command with ErrorLevel as the length of
    ;   the string ensures that only valid text from the Edit control is
    ;   returned.
    SendMessage EM_GETLINE,p_LineIndex,&Text,,ahk_id %hEdit%
    Return SubStr(Text,1,ErrorLevel)
    }

;------------------------------
;
; Function: Edit_GetLineCount
;
; Description:
;
;   Return the total number of text lines in a multiline Edit control.  If the
;   Edit control is empty, 1 is returned.  The return value will never be less
;   than 1.
;
; Remarks:
;
;   The value returned is for the number of lines in the Edit control.  Very
;   long lines (more than 1,024 characters) and/or the word wrap style may
;   introduce additional lines in the control.
;
;   The EM_GETLINECOUNT message used by this function is designed to be used on
;   a multiline Edit control.  However, if used on a single-line Edit control,
;   1 is returned.
;
;-------------------------------------------------------------------------------
Edit_GetLineCount(hEdit)
    {
    Static EM_GETLINECOUNT:=0xBA
    SendMessage EM_GETLINECOUNT,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetMargins
;
; Description:
;
;   Get the margins for the Edit control as measured in pixels.
;
; Parameters:
;
;   r_LeftMargin, r_RightMargin - [Output, Optional] If specified, these
;       variables contain the margins of the Edit control in pixels.
;
; Returns:
;
;   An AutoHotkey object (tests as TRUE) that contains the following properties:
;
;   LeftMargin, RightMargin - The margins of the Edit control in pixels.
;
; Remarks:
;
;   This function uses the EM_GETMARGINS message without modifications.  This
;   message along with companion EM_SETMARGINS message is _not_ DPI aware.  For
;   a DPI aware version of this function, use <Edit_GetMarginsInPixels> instead.
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_GetMargins(hEdit,ByRef r_LeftMargin:="",ByRef r_RightMargin:="")
    {
    Static EM_GETMARGINS:=0xD4
    SendMessage EM_GETMARGINS,0,0,,ahk_id %hEdit%
    r_LeftMargin :=ErrorLevel&0xFFFF    ;-- LOWORD of result
    r_RightMargin:=ErrorLevel>>16       ;-- HIWORD of result
    Return {LeftMargin:r_LeftMargin,RightMargin:r_RightMargin}
    }

;------------------------------
;
; Function: Edit_GetMarginsInInches
;
; Description:
;
;   Get the margins for the Edit control as measured in inches.
;
; Parameters:
;
;   r_LeftMargin, r_RightMargin - [Output, Optional] If specified, these
;       variables contains the margins of the Edit control as measured in
;       inches.  Ex: 0.333333 (1/3 inch)
;
; Returns:
;
;   An AutoHotkey object (tests as TRUE) that contains the following properties:
;
;   LeftMargin, RightMargin - The margins of the Edit control as measured in
;       inches.  Ex: 0.5 (1/2 inch).
;
; Remarks:
;
;   This function is designed to be used in conjunction with
;   <Edit_SetMarginsInInches>.
;
;-------------------------------------------------------------------------------
Edit_GetMarginsInInches(hEdit,ByRef r_LeftMargin:="",ByRef r_RightMargin:="")
    {
    Static EM_GETMARGINS:=0xD4
    SendMessage EM_GETMARGINS,0,0,,ahk_id %hEdit%
    r_LeftMargin :=(ErrorLevel&0xFFFF)/96               ;-- LOWORD of result
    r_RightMargin:=(ErrorLevel>>16)/96                  ;-- HIWORD of result
    Return {LeftMargin:r_LeftMargin,RightMargin:r_RightMargin}
    }

;------------------------------
;
; Function: Edit_GetMarginsInPixels
;
; Description:
;
;   Get the margins for the Edit control as measured in pixels.
;
; Parameters:
;
;   r_LeftMargin, r_RightMargin - [Output, Optional] If specified, these
;       variables contains the margins of the Edit control as measured in
;       pixels.  Ex: 25.
;
; Returns:
;
;   An AutoHotkey object (tests as TRUE) with the following properties:
;
;   LeftMargin, RightMargin - The margins of the Edit control in pixels.  Ex:
;       25.
;
; Remarks:
;
;   Unlike <Edit_GetMargins>, this function is DPI aware.  The returned margin
;   values are scaled to the computer's current DPI.  If conversion is required,
;   the results are rounded to the nearest pixel.
;
;   This function is designed to be used in conjunction with
;   <Edit_SetMarginsInPixels>.
;
;-------------------------------------------------------------------------------
Edit_GetMarginsInPixels(hEdit,ByRef r_LeftMargin:="",ByRef r_RightMargin:="")
    {
    Static EM_GETMARGINS:=0xD4
    SendMessage EM_GETMARGINS,0,0,,ahk_id %hEdit%
    r_LeftMargin :=Round((ErrorLevel&0xFFFF)*A_ScreenDPI/96)
        ;-- LOWORD of result
    r_RightMargin:=Round((ErrorLevel>>16)*A_ScreenDPI/96)
        ;-- HIWORD of result
    Return {LeftMargin:r_LeftMargin,RightMargin:r_RightMargin}
    }

;------------------------------
;
; Function: Edit_GetModify
;
; Description:
;
;   Get the state of the Edit control's modification flag.  The flag indicates
;   whether the contents of the Edit control have been modified.
;
; Returns:
;
;   TRUE if the Edit control has been modified, otherwise FALSE.
;
;-------------------------------------------------------------------------------
Edit_GetModify(hEdit)
    {
    Static EM_GETMODIFY:=0xB8
    SendMessage EM_GETMODIFY,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetPasswordChar
;
; Description:
;
;   Get the password character that the Edit control displays when the user
;   enters text.
;
; Returns:
;
;   The decimal value of the character that is displayed in place of the
;   characters typed by the user.  If a password character has not been set, 0
;   is returned.
;
; Remarks:
;
;   For most versions of Windows, the default password character decimal value
;   is 9679 (black circle).  If needed, use the built-in AutoHotkey *<Chr at
;   https://www.autohotkey.com/docs/v1/lib/Chr.htm>* function to convert the
;   return value to a character.  For example:
;
;       (start code)
;       PWChar:=Chr(Edit_GetPasswordChar(hEdit))
;       (end code)
;
;   Exception: When using the ANSI version of AutoHotkey, the *Chr* function
;   will only work correctly if the value is between 1 and 255.
;
;-------------------------------------------------------------------------------
Edit_GetPasswordChar(hEdit)
    {
    Static EM_GETPASSWORDCHAR:=0xD2
    Return DllCall("SendMessageW","UPtr",hEdit,"UInt",EM_GETPASSWORDCHAR,"UInt",0,"UInt",0)
    }

;------------------------------
;
; Function: Edit_GetPos
;
; Description:
;
;   Get the position and size of the Edit control.  See the *Remarks* section
;   for more information.
;
; Parameters:
;
;   r_X, r_Y - [Output, Optional] If specified, these variables contain the
;       coordinates of the Edit control relative to the client-area of the
;       parent window.
;
;   r_Width, r_Height - [Output, Optional] If specified, these variables
;       contain the width and height of the Edit control.
;
; Returns:
;
;   An AutoHotkey object (tests as TRUE) with the following properties:
;
;   X, Y - The coordinates of the Edit control relative to the client-area of
;       the parent window.
;
;   Width, Height - The width and height of the Edit control.
;
;   SizeStr - The size of the Edit control as a string of characters.  Ex:
;       "800x600".
;
; Remarks:
;
;   This function returns values similar to the AutoHotkey
;   *GUIControlGet,OutputVar,Pos* command.  The coordinates values (i.e. X and
;   Y) are relative to the parent window's client area.  However, unlike the
;   *GUIControlGet,OutputVar,Pos* command, the return values from this function
;   do not change when the computer's DPI changes.  The returned values are
;   actual values, not calculated values based on the current screen DPI.  In
;   addition, this function works on all Edit controls whereas the
;   *<GUIControlGet at
;    https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command only
;    works on Edit controls created using the AutoHotkey "gui Add" command.
;
;   If the Edit control was created using the AutoHotkey "gui Add" command and
;   the "-DPIScale" option is specified, the *<GUIControlGet at
;   https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command can be
;   used instead.  The *<ControlGetPos at
;   https://www.autohotkey.com/docs/v1/lib/ControlGetPos.htm>* and *<WinGetPos
;   at https://www.autohotkey.com/docs/v1/lib/WinGetPos.htm>* commands are not
;   DPI aware and so if only interested in the width and/or height values, these
;   commands can be used on all Edit controls.  Hint: The native AutoHotkey
;   commands are more efficient and should be used whenever possible.
;
;-------------------------------------------------------------------------------
Edit_GetPos(hEdit,ByRef r_X:="",ByRef r_Y:="",ByRef r_Width:="",ByRef r_Height:="")
    {
    ;-- Get the dimensions of the bounding rectangle of the Edit control
    VarSetCapacity(RECT,16)
    if not DllCall("GetWindowRect","UPtr",hEdit,"UPtr",&RECT)
        {
        outputdebug Function: %A_ThisFunc% - Call to GetWindowRect failed.
        r_X:=r_Y:=r_Width:=r_Height:=0
        Return {X:r_X,Y:r_Y,Width:r_Width,Height:r_Height,SizeStr:r_Width . "x" . r_Height}
        }

    r_Width :=NumGet(RECT,8,"Int")-NumGet(RECT,0,"Int")  ;-- Width=right-left
    r_Height:=NumGet(RECT,12,"Int")-NumGet(RECT,4,"Int") ;-- Height=bottom-top

    ;-- Convert the screen coordinates to client-area coordinates.
    ;   Note: The "ScreenToClient" system function reads and then updates the
    ;   first 8-bytes of the RECT structure.
    DllCall("ScreenToClient","UPtr",DllCall("GetParent","UPtr",hEdit,"UPtr"),"UPtr",&RECT)

    ;-- Update the output variables
    r_X:=NumGet(RECT,0,"Int")                           ;-- left
    r_Y:=NumGet(RECT,4,"Int")                           ;-- top

    ;-- Build and return object
    Return {X:r_X,Y:r_Y,Width:r_Width,Height:r_Height,SizeStr:r_Width . "x" . r_Height}
    }

;------------------------------
;
; Function: Edit_GetRect
;
; Description:
;
;   Get the formatting rectangle of the Edit control.
;
; Parameters:
;
;   r_Left..r_Bottom - [Output, Optional] If specified, these variables contain
;       the values of the formatting rectangle of the Edit control.
;
; Returns:
;
;   An AutoHotkey object (tests as TRUE) with the following properties:
;
;   Left..Bottom - The formatting rectangle of the window as individual values.
;       To get the formatting rectangle as a RECT structure, use the RECT
;       property.
;
;   RECT - A RECT structure that contains the formatting rectangle of the Edit
;       control.
;
;   SizeStr - The size of the formatting rectangle as a string of characters.
;       Ex: "794x594".
;
;   Width, Height - The width and height of the formatting rectangle.
;
;   X, Y - The coordinates of the formatting rectangle.  These are the same
;       values as the Left and Top properties.
;
;-------------------------------------------------------------------------------
Edit_GetRect(hEdit,ByRef r_Left:="",ByRef r_Top:="",ByRef r_Right:="",ByRef r_Bottom:="")
    {
    Static Dummy26084539
          ,RECT
          ,Dummy1:=VarSetCapacity(RECT,16)

          ;-- Message
          ,EM_GETRECT:=0xB2

    ;-- Get formatting rectangle
    SendMessage EM_GETRECT,0,&RECT,,ahk_id %hEdit%

    ;-- Create and populate the return object and set output parameters
    RO:=Object()
    RO.SetCapacity(10)
    RO.X:=RO.Left:=r_Left:=NumGet(RECT,0,"Int")
    RO.Y:=RO.Top:=r_Top  :=NumGet(RECT,4,"Int")
    RO.Right:=r_Right    :=NumGet(RECT,8,"Int")
    RO.Bottom:=r_Bottom  :=NumGet(RECT,12,"Int")
    RO.Width             :=RO.Right-RO.Left
    RO.Height            :=RO.Bottom-RO.Top
    RO.SizeStr           :=RO.Width . "x" . RO.Height

    ;-- RECT structure
    RO.SetCapacity("RECT",16)
    DllCall("CopyRect","UPtr",RO.GetAddress("RECT"),"UPtr",&RECT)

    ;-- Return object
    Return RO
    }

;------------------------------
;
; Function: Edit_GetSel
;
; Description:
;
;   Get the start and end character positions of the current selection.
;
; Parameters:
;
;   r_StartSelPos, r_EndSelPos - [Output, Optional] If specified, these
;       variables contains the start and end character positions of the
;       selection.
;
; Returns:
;
;   The start position of the selection.
;
; Observations:
;
;   Regardless of how the text is selected (i.e. from left to right, right to
;   left, top to bottom, or bottom to top), StartSelPos will always contain the
;   lower character position value.  If no text is selected, EndSelPos will be
;   the same as StartSelPos.  If text is selected, EndSelPos will always contain
;   the higher character position value.
;
;-------------------------------------------------------------------------------
Edit_GetSel(hEdit,ByRef r_StartSelPos:="",ByRef r_EndSelPos:="")
    {
    Static Dummy53041672
          ,s_StartSelPos
          ,s_EndSelPos
          ,Dummy1:=VarSetCapacity(s_StartSelPos,4)
          ,Dummy2:=VarSetCapacity(s_EndSelPos,4)

          ;-- Message
          ,EM_GETSEL:=0xB0

    ;-- Get the current select positions
    SendMessage EM_GETSEL,&s_StartSelPos,&s_EndSelPos,,ahk_id %hEdit%
    r_StartSelPos:=NumGet(s_StartSelPos,0,"UInt")
    r_EndSelPos  :=NumGet(s_EndSelPos,0,"UInt")
    Return r_StartSelPos
    }

;------------------------------
;
; Function: Edit_GetSelText
;
; Description:
;
;   Return the currently selected text (if any).
;
; Calls To Other Functions:
;
; * <Edit_GetLine>
; * <Edit_GetText>
;
; Programming Notes:
;
;   The Edit control does not support the EM_GETSELTEXT message.  If the
;   selection is on one line, the EM_GETLINE messages is used, otherwise the
;   WM_GETTEXT message is used.
;
;-------------------------------------------------------------------------------
Edit_GetSelText(hEdit)
    {
    Static Dummy21374560
          ,s_StartSelPos
          ,s_EndSelPos
          ,Dummy1:=VarSetCapacity(s_StartSelPos,4)
          ,Dummy2:=VarSetCapacity(s_EndSelPos,4)

          ;-- Messages
          ,EM_GETSEL      :=0xB0
          ,EM_LINEFROMCHAR:=0xC9
          ,EM_LINEINDEX   :=0xBB

    ;-- Get the current select positions
    SendMessage EM_GETSEL,&s_StartSelPos,&s_EndSelPos,,ahk_id %hEdit%
    StartSelPos:=NumGet(s_StartSelPos,0,"UInt")
    EndSelPos  :=NumGet(s_EndSelPos,0,"UInt")

    ;-- Return a null string if nothing is selected
    if (StartSelPos=EndSelPos)
        Return

    ;-- Get the index of line with the start select position
    SendMessage EM_LINEFROMCHAR,StartSelPos,0,,ahk_id %hEdit%
    FirstSelectedLine:=ErrorLevel

    ;-- Get the index of line with the end select position
    SendMessage EM_LINEFROMCHAR,EndSelPos,0,,ahk_id %hEdit%
    LastSelectedLine:=ErrorLevel

    ;-- Selection on one line?
    if (FirstSelectedLine=LastSelectedLine)
        {
        ;-- Get the index of the first character of the selected line
        SendMessage EM_LINEINDEX,FirstSelectedLine,0,,ahk_id %hEdit%
        FirstCharPos:=ErrorLevel<<32>>32  ;-- Convert UInt to Int

        ;-- Get the selected text from a line
        ;   Note: If the text was selected using a mouse or with keyboard
        ;   shortcuts, the Edit_GetLine function will work as expected.
        ;   However, if text was programmatically selected (Ex: Edit_SetSel
        ;   function), there is tiny chance that the end selection is after the
        ;   carriage return character but before the line feed character.  The
        ;   Edit_GetLine function automatically stops loading the line when the
        ;   end-of-line (EOL) characters are reached so if the selection
        ;   includes one, but not both, of the end-of-line characters, the
        ;   Edit_GetLine function will not get all of the selected characters.
        ;   Testing the length of the text collected by the Edit_GetLine
        ;   function against the start and end selection values will identify if
        ;   there is a problem.  If the lengths do not match, the function will
        ;   fall through and call the Edit_GetText function to get the selected
        ;   text.
        SelectedText:=SubStr(Edit_GetLine(hEdit,FirstSelectedLine,EndSelPos-FirstCharPos),StartSelPos-FirstCharPos+1)
        if (StrLen(SelectedText)=EndSelPos-StartSelPos)
            Return SelectedText

        ;-- If we get this far, fall through and call Edit_GetText
        }

    ;-- Selection includes multiple lines and/or EOL character(s).  Use
    ;   Edit_GetText.
    Return SubStr(Edit_GetText(hEdit,EndSelPos),StartSelPos+1)
    }

;------------------------------
;
; Function: Edit_GetStyle
;
; Description:
;
;   Return an integer that represents the styles currently set for the Edit
;   control.
;
;-------------------------------------------------------------------------------
Edit_GetStyle(hEdit)
    {
;;;;;    Static GWL_STYLE:=-16
;;;;;    Return DllCall("GetWindowLong","UInt",hEdit,"Int",GWL_STYLE)
    ControlGet EditStyle,Style,,,ahk_id %hEdit%
    if ErrorLevel
        outputdebug,
           (ltrim join`s
            Function: %A_ThisFunc% - Unexpected error from ControlGet command.
           )

    Return EditStyle
    }

;------------------------------
;
; Function: Edit_GetText
;
; Description:
;
;   Return all text from the Edit control up to p_Length length.  If p_Length is
;   set to -1 (the default), all text in the Edit control is returned.
;
; Remarks:
;
;   This function is similar to the AutoHotkey *<GUIControlGet at
;   https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command (for AHK
;   GUIs) and the *<ControlGetText at
;   https://www.autohotkey.com/docs/v1/lib/ControlGetText.htm>* command except
;   that end-of-line (EOL) characters from the retrieved text are not
;   automatically converted (CR+LF to LF).  If needed, use <Edit_Convert2Unix>
;   to convert the text to the AutoHotkey text format.
;
;-------------------------------------------------------------------------------
Edit_GetText(hEdit,p_Length:=-1)
    {
    Static Dummy34127859
          ,WM_GETTEXT:=0xD
          ,WM_GETTEXTLENGTH:=0xE

    ;-- If needed, determine the length of the text
    if (p_Length<0)
        {
        SendMessage WM_GETTEXTLENGTH,0,0,,ahk_id %hEdit%
        p_Length:=ErrorLevel
        }

    ;-- Add 1 to the length for a trailing null character.
    ;   Note: This adjustment is for the WM_GETTEXT message which requires that
    ;   the specified length include a trailing null character.
    p_Length+=1

    ;-- Get text
    VarSetCapacity(Text,p_Length*(A_IsUnicode ? 2:1))
    SendMessage WM_GETTEXT,p_Length,&Text,,ahk_id %hEdit%
    Return Text
    }

;------------------------------
;
; Function: Edit_GetTextLength
;
; Description:
;
;   Return the number of text characters in the Edit control.
;
; Remarks:
;
;   The length includes end-of-line (EOL) characters if any.
;
;-------------------------------------------------------------------------------
Edit_GetTextLength(hEdit)
    {
    Static WM_GETTEXTLENGTH:=0xE
    SendMessage WM_GETTEXTLENGTH,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_GetTextRange
;
; Description:
;
;   Get a range of characters from the Edit control.
;
; Parameters:
;
;   p_StartPos - [Optional] The zero-based character index of the first
;       character in the range.  If set to 0 (the default) or if not specified,
;       the starting character is the first character of the Edit control.
;
;   p_EndPos - [Optional] The zero-based character index of the last character
;       in the range plus 1.  If set to -1 (the default) or if not specified,
;       the end character is the last character in the Edit control.
;
; Calls To Other Functions:
;
; * <Edit_GetText>
;
; Remarks:
;
;   This function does not use the EM_GETTEXTRANGE message because it is only
;   supported by the Rich Edit control.  This function uses <Edit_GetText>
;   (WM_GETTEXT message) to get the requested range of characters from the Edit
;   control.
;
;   Unlike the EM_SETSEL message, the start and end positions are not
;   interchangeable for this function.  The start position must be the lower
;   value and the end position must be the higher value.
;
;-------------------------------------------------------------------------------
Edit_GetTextRange(hEdit,p_StartPos:=0,p_EndPos:=-1)
    {
    ;-- Parameters
    if p_StartPos is not Integer
        p_StartPos:=0
     else if (p_StartPos<0)
        p_StartPos:=0

    if p_EndPos is not Integer
        p_EndPos:=-1

    ;-- Get text range
    Return SubStr(Edit_GetText(hEdit,p_EndPos),p_StartPos+1)
    }

;------------------------------
;
; Function: Edit_GetZoom
;
; Description:
;
;   Get the current zoom ratio for a multiline Edit control.
;
; Parameters:
;
;   r_Numerator - [Output, Optional] If specified, this variable contains the
;       numerator of the zoom ratio.  Ex: 80.  The value will be zero (0) if
;       Edit control is not zoomable or if there was an error.
;
;   r_Denominator - [Output, Optional] If specified, this variable contains the
;       denominator of the zoom ratio.  Ex: 100.  The value will be zero (0) if
;       Edit control is not zoomable or if there was an error.
;
;   r_ZoomPct - [Output, Optional] If specified, this variable contains the
;       zoom ratio as a percentage.  Ex: 120.  The value will be zero (0) if
;       Edit control is not zoomable or if there was an error.
;
; Returns:
;
;   An AutoHotkey object (tests as TRUE) if successful, otherwise FALSE.  If
;   returned, the object contains the following properties:
;
;   Numerator - The numerator of the zoom ratio.  Ex: 80.  The value will be
;       zero (0) if Edit control is not zoomable or if there was an error.
;
;   Denominator - The denominator of the zoom ratio.  Ex: 100.  The value will
;       be zero (0) if Edit control is not zoomable or if there was an error.
;
;   ZoomPct - The zoom ratio as a percentage.  Ex: 120.  The value will be zero
;       (0) if the Edit control is not zoomable or if there was an error.
;
; Requirements:
;
;   Windows 10+
;
; Remarks:
;
;   This function will not return a usable values unless the ES_EX_ZOOMABLE
;   extended style is set on the Edit control.  If needed, use <Edit_IsZoomable>
;   to determine if the ES_EX_ZOOMABLE extended style has been set.
;
;-------------------------------------------------------------------------------
Edit_GetZoom(hEdit,ByRef r_Numerator:="",ByRef r_Denominator:="",ByRef r_ZoomPct:="")
    {
    Static EM_GETZOOM:=0x4E0  ;-- WM_USER+224

    ;-- Get Zoom
    r_Numerator:=r_Denominator:=0  ;-- Initialize jic SendMessage fails
    DllCall("SendMessage" . (A_IsUnicode ? "W":"A")
        ,"UPtr",hEdit
        ,"UInt",EM_GETZOOM
        ,"UInt*",r_Numerator
        ,"UInt*",r_Denominator)

    ;-- Populate output variables and return object
    r_ZoomPct:=r_Denominator ? Round((r_Numerator/r_Denominator)*100):0
    Return {Numerator:r_Numerator,Denominator:r_Denominator,ZoomPct:r_ZoomPct}
    }

;------------------------------
;
; Function: Edit_HasFocus
;
; Description:
;
;   Determine if the Edit control has functional input focus, aka "keyboard
;   focus".
;
; Returns:
;
;   TRUE if the Edit control has functional input focus, otherwise FALSE.
;
; Credit:
;
;   Adapted from an example in the AutoHotkey documentation.
;
; Remarks:
;
;   This function uses the *GetGUIThreadInfo* system function to determine if
;   the Edit control has focus or not.  It is very efficient and if needed, this
;   function can be called frequently (assuming there are no errors) without
;   degrading the response of the script.
;
;-------------------------------------------------------------------------------
Edit_HasFocus(hEdit)
    {
    Static Dummy72914086
          ,GUITHREADINFO

          ;-- Create and initialize GUITHREADINFO structure
          ,sizeofGUITHREADINFO:=A_PtrSize=8 ? 72:48
          ,Dummy1:=VarSetCapacity(GUITHREADINFO,sizeofGUITHREADINFO,0)
          ,Dummy2:=NumPut(sizeofGUITHREADINFO,GUITHREADINFO,0,"UInt")

    ;-- Collect GUI Thread Info
    if DllCall("GetGUIThreadInfo","UInt",0,"UPtr",&GUITHREADINFO)
        Return (hEdit=NumGet(GUITHREADINFO,A_PtrSize=8 ? 16:12,"UPtr"))
            ;-- hwndFocus

    ;-- Error
    outputdebug,
       (ltrim join`s
        Function: %A_ThisFunc% -
        Call to GetGUIThreadInfo failed.  A_LastError: %A_LastError%
       )

    Return False
    }

;------------------------------
;
; Function: Edit_Hide
;
; Description:
;
;   Hide the Edit control.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   For AutoHotkey GUIs, use the *<GUIControl at
;   https://www.autohotkey.com/docs/v1/lib/GuiControl.htm>* command for improved
;   efficiency.  Ex: GUIControl 33:Hide,%hEdit%
;
;   This command only hides the Edit control, it does not disable it.  To
;   prevent use of the Edit control's keyboard shortcut keys, be sure to disable
;   the Edit control as well.
;
;-------------------------------------------------------------------------------
Edit_Hide(hEdit)
    {
    Control Hide,,,ahk_id %hEdit%
    Return ErrorLevel ? False:True
    }

;------------------------------
;
; Function: Edit_HideBalloonTip
;
; Description:
;
;   Hide any balloon tip associated with the Edit control.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   This function is usually unnecessary.  A balloon tip attached to the Edit
;   control will usually auto-hide after short period of time (8 to 12 seconds).
;   In addition, the balloon tip will auto-hide if the contents of the control
;   are changed or if focus is moved to another control.
;
;-------------------------------------------------------------------------------
Edit_HideBalloonTip(hEdit)
    {
    Static EM_HIDEBALLOONTIP:=0x1504  ;-- ECM_FIRST+4
    SendMessage EM_HIDEBALLOONTIP,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_IsDisabled
;
; Description:
;
;   Return TRUE if the Edit control is disabled.
;
; Calls To Other Functions:
;
; * <Edit_GetStyle>
;
;-------------------------------------------------------------------------------
Edit_IsDisabled(hEdit)
    {
    Static WS_DISABLED:=0x8000000
    Return Edit_GetStyle(hEdit) & WS_DISABLED ? True:False
    }

;------------------------------
;
; Function: Edit_IsHidden
;
; Description:
;
;   Return TRUE if the Edit control is hidden.
;
; Calls To Other Functions:
;
; * <Edit_GetStyle>
;
;-------------------------------------------------------------------------------
Edit_IsHidden(hEdit)
    {
    Static WS_VISIBLE:=0x10000000
    Return Edit_GetStyle(hEdit) & WS_VISIBLE ? False:True
    }

;------------------------------
;
; Function: Edit_IsMultiline
;
; Description:
;
;   Return TRUE if the Edit control is multiline, otherwise FALSE.
;
;-------------------------------------------------------------------------------
Edit_IsMultiline(hEdit)
    {
    Static ES_MULTILINE:=0x4
    Return Edit_GetStyle(hEdit) & ES_MULTILINE ? True:False
    }

;------------------------------
;
; Function: Edit_IsReadOnly
;
; Description:
;
;   Return TRUE if the ES_READONLY style has been set on the Edit control,
;   otherwise FALSE.
;
;-------------------------------------------------------------------------------
Edit_IsReadOnly(hEdit)
    {
    Static ES_READONLY:=0x800
    Return Edit_GetStyle(hEdit) & ES_READONLY ? True:False
    }

;------------------------------
;
; Function: Edit_IsStyle
;
; Description:
;
;   Determine if a style or combination of styles has been set on the Edit
;   control.
;
; Parameters:
;
;   p_Style - See the *Style* section for more information.
;
; Returns:
;
;   TRUE if the specified style or combination of styles has been set on the
;   Edit control, otherwise FALSE.
;
;   FALSE is also returned if p_Style contain an invalid value.  A
;   developer-friendly message is dumped to debugger if this occurs.
;
; Style:
;
;   The p_Style parameter is used to identify which Edit control style or
;   combination of styles to test for.  See the function's static variables for
;   a list of possible style values.
;
;   Alternatively, a style name can be specified.  Ex: "Uppercase".  The
;   Microsoft style constant name less the "ES_" or "WS_" prefix can be used as
;   the style name.  For example, for the ES_NUMBER constant, "Number" can be
;   used as the style name.  To specify more than one style name, the names must
;   be delimited by a space character (Ex: "Multiline NoHideSel") or by the "|"
;   (pipe) character (Ex: "Right|Uppercase").  This syntax also supports a mix
;   of style names and style integer values.  Ex: "Right|0x4|ReadOnly".
;
;   A few notes.
;
;     * Style names (Ex: "Center") can only be used for constants defined in the
;       function's static variables which only includes constants for common
;       Edit control styles.
;
;     * The tests for style names is not case sensitive but invalid style names
;       are ignored.  Be sure to test to thoroughly.
;
;     * Styles not included in the function's static variables can also be
;       tested but the style's integer value must be specified.
;
;-------------------------------------------------------------------------------
Edit_IsStyle(hEdit,p_Style)
    {
    Static Dummy35192468

          ;-- Common Edit control styles
          ;   Note: See the Edit_Constants.ahk document for more information
          ,ES_LEFT       :=0x0  ;-- Cannot test this style.  It's the absence of ES_CENTER and ES_RIGHT.
          ,ES_CENTER     :=0x1
          ,ES_RIGHT      :=0x2
          ,ES_MULTILINE  :=0x4
          ,ES_UPPERCASE  :=0x8
          ,ES_LOWERCASE  :=0x10
          ,ES_PASSWORD   :=0x20
          ,ES_AUTOVSCROLL:=0x40
          ,ES_AUTOHSCROLL:=0x80
          ,ES_NOHIDESEL  :=0x100
          ,ES_COMBO      :=0x200
          ,ES_OEMCONVERT :=0x400
          ,ES_READONLY   :=0x800
          ,ES_WANTRETURN :=0x1000
          ,ES_NUMBER     :=0x2000
          ,WS_TABSTOP    :=0x10000
          ,WS_HSCROLL    :=0x100000
          ,WS_VSCROLL    :=0x200000

    ;-- If needed, convert style name(s)
    if p_Style is not Integer
        {
        t_Style:=0x0
        Loop Parse,p_Style,|%A_Tab%%A_Space%
            {
            ;-- Skip if whitespace (blank, null, tab, etc.)
            if A_LoopField is Space
                Continue

            ;-- If integer, use as-is
            if A_LoopField is Integer
                {
                t_Style|=A_LoopField
                Continue
                }

            ;-- Convert style name
            if A_LoopField is Alpha  ;-- Alpha characters only
                if ES_%A_LoopField% is not Space
                    t_Style|=ES_%A_LoopField%
                 else if WS_%A_LoopField% is not Space
                    t_Style:=WS_%A_LoopField%
            }

        p_Style:=t_Style
        }

    ;-- Bounce and return FALSE if p_Style contains no value
    if not p_Style  ;-- Zero, null, or blank
        {
        outputdebug,
           (ltrim join`s
            Function: %A_ThisFunc% - No styles found in p_Style
           )

        Return False
        }

    ;-- Get the Edit control's style
    ControlGet EditStyle,Style,,,ahk_id %hEdit%
    if ErrorLevel
        outputdebug,
           (ltrim join`s
            Function: %A_ThisFunc% - Unexpected error from ControlGet command.
           )

    ;-- Test if the Edit control has the specified style(s)
    Return EditStyle & p_Style ? True:False
    }

;------------------------------
;
; Function: Edit_IsWordWrap
;
; Description:
;
;   Return TRUE if word wrap is enabled on the Edit control, otherwise FALSE.
;
; Remarks:
;
;   Definition: A false positive is when this function indicates that word wrap
;   is enabled on the Edit control (i.e. returns TRUE), when in fact, word wrap
;   is not enabled.  A false negative is when this function indicates that word
;   wrap is not enabled on the Edit control (i.e. returns FALSE), when in fact,
;   word wrap is enabled.
;
;   This function can return a false positive by hiding the horizontal scroll
;   bar after the Edit control has been created.  Although this situation is
;   rare, it is a possibility.  One way to ensure that the function always
;   returns FALSE correctly (i.e. word wrap is not enabled) is to always include
;   the ES_AUTOHSCROLL style (0x80 or -Wrap) when the WS_HSCROLL style
;   (0x100000) is also included.  So far, no situations where a false negative
;   is returned have been identified.
;
;-------------------------------------------------------------------------------
Edit_IsWordWrap(hEdit)
    {
    Static Dummy82465173

          ;-- Styles
          ,ES_LEFT       :=0x0
          ,ES_CENTER     :=0x1
          ,ES_RIGHT      :=0x2
          ,ES_MULTILINE  :=0x4
          ,ES_AUTOHSCROLL:=0x80
          ,WS_HSCROLL    :=0x100000

    ;-- Get the current style
    Style:=Edit_GetStyle(hEdit)

    ;----------------------------------------------------------------------------
    ;
    ;   Note: The following tests must be performed in the current order.  All
    ;   tests assume that conditions from previous tests have been met.
    ;
    ;---------------------------------------------------------------------------

    ;-- FALSE if not multiline
    ;   Background: Word wrap is only an option of a multiline Edit control.
    if not (Style & ES_MULTILINE)
        Return False

    ;-- TRUE if ES_CENTER or ES_RIGHT style is set
    ;   Background: ES_AUTOHSCROLL is ignored by a multiline Edit control that
    ;   is not left-aligned.  Centered and right-aligned multiline Edit controls
    ;   cannot be horizontally scrolled.
    if Style & (ES_CENTER|ES_RIGHT)
        Return True

    ;-- FALSE if ES_AUTOHSCROLL style is set
    if Style & ES_AUTOHSCROLL
        Return False

    ;-- FALSE if WS_HSCROLL style is set
    ;   Background: ES_AUTOHSCROLL is automatically applied to a left-aligned,
    ;   multiline Edit control that has a WS_HSCROLL style.
    if Style & WS_HSCROLL
        Return False

    ;-- Otherwise, return TRUE
    Return True
    }

;------------------------------
;
; Function: Edit_IsZoomable
;
; Description:
;
;   Determine if the ES_EX_ZOOMABLE extended style has been set on the Edit
;   control.
;
; Returns:
;
;   TRUE if the ES_EX_ZOOMABLE extended style has been set on the Edit control,
;   otherwise FALSE.  FALSE is also returned if there is an error.
;
; Requirements:
;
;   Windows 10+
;
;-------------------------------------------------------------------------------
Edit_IsZoomable(hEdit)
    {
    Static Dummy14352728
          ,EM_GETEXTENDEDSTYLE:=0x150B  ;-- ECM_FIRST+11
          ,ES_EX_ZOOMABLE:=0x10

    SendMessage EM_GETEXTENDEDSTYLE,0,0,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel&ES_EX_ZOOMABLE ? True:False
    }

;------------------------------
;
; Function: Edit_LineFromChar
;
; Description:
;
;   Get the line in the Edit control that contains the specified character.
;
; Parameters:
;
;   p_CharPos - [Optional] The zero-based index of the character.  Ex: 726.  If
;       set to -1 (the default) or if not specified, the index of the current
;       line is retrieved.  See the *Observations* section for more information.
;
; Returns:
;
;   The zero-based index of the line containing the character index specified
;   by the p_CharPos parameter.
;
; Remarks:
;
;   If the p_CharPos parameter is set to -1 (the default), the function will
;   return the zero-based index of the current line (the line containing the
;   caret) in _all_ cases.  When multiple lines of text are selected, the
;   location of the caret will depend on how the text is selected.  If the text
;   is selected from top to bottom, the caret will be on the lower (bottom) line
;   of the selected text.  If selecting from bottom to top, the caret will on
;   the higher (top) line of the selected text.  This information is a bit
;   different than the Microsoft documentation but is more accurate.
;
;-------------------------------------------------------------------------------
Edit_LineFromChar(hEdit,p_CharPos:=-1)
    {
    Static EM_LINEFROMCHAR:=0xC9
    SendMessage EM_LINEFROMCHAR,p_CharPos,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_LineFromPos
;
; Description:
;
;   This function is the same as <Edit_CharFromPos> except that the line index
;   (r_LineIndex) is returned.
;
;-------------------------------------------------------------------------------
Edit_LineFromPos(hEdit,X,Y,ByRef r_CharPos:="",ByRef r_LineIndex:="")
    {
    Edit_CharFromPos(hEdit,X,Y,r_CharPos,r_LineIndex)
    Return r_LineIndex
    }

;------------------------------
;
; Function: Edit_LineIndex
;
; Description:
;
;   Get the character position of the first character of the specified line.
;
; Parameters:
;
;   p_LineIndex - [Optional] The zero-based index of the line.  If set to -1
;       (the default) or if not specified, the current line is used.
;
; Returns:
;
;   The character position of the first character of the specified line or -1 if
;   the specified line is greater than the total number of lines in the Edit
;   control.
;
;-------------------------------------------------------------------------------
Edit_LineIndex(hEdit,p_LineIndex:=-1)
    {
    Static EM_LINEINDEX:=0xBB
    SendMessage EM_LINEINDEX,p_LineIndex,0,,ahk_id %hEdit%
    Return ErrorLevel<<32>>32  ;-- Convert UInt to Int
    }

;------------------------------
;
; Function: Edit_LineIsVisible
;
; Description:
;
;   Determine if a line is visible.
;
; Parameters:
;
;   p_LineIndex - [Optional] The zero-based index of the line to check.  Set to
;        -1 (the default) or don't specify to check the current line.
;
;   p_Threshold - [Optional] The threshold for the last visible line in the
;       client area of the Edit control.  See the *Threshold* section in
;       <Edit_GetLastVisibleLine> for more information.
;
;   r_ScrollLines - [Output, Optional]  See the *Scroll Lines* section for more
;       information.
;
; Returns:
;
;   TRUE if the specified line is visible, otherwise FALSE.
;
; Calls To Other Functions:
;
; * <Edit_GetLastVisibleLine>
;
; Scroll Lines:
;
;   r_ScrollLines is an optional output parameter.  If specified, the variable
;   will contain the number of lines to scroll in order to make the
;   specified/current line visible.  The value will be negative to scroll up
;   (Ex: -3), positive to scroll down (Ex: 5), or 0 if the specified/current
;   line is already visible.  <Edit_LineScroll> or <Edit_Scroll> can be used to
;   scroll using this value.
;
;   If the p_LineIndex parameter contains a line index value that is greater
;   than the maximum line index, the function will return FALSE, i.e. the line
;   is not visible.  The r_ScrollLines output variable will contain the number
;   of lines necessary to make the last line visible.  This value can be 0.
;
; Remarks:
;
;   If used on a single-line Edit control, TRUE is returned if p_LineIndex is
;   set to -1 or 0, otherwise FALSE is returned.  If specified, the
;   r_ScrollLines output variable will contain 0.
;
;-------------------------------------------------------------------------------
Edit_LineIsVisible(hEdit,p_LineIndex:=-1,p_Threshold:="",ByRef r_ScrollLines:=0)
    {
    Static Dummy47096581
          ,EM_GETFIRSTVISIBLELINE:=0xCE
          ,EM_GETLINECOUNT:=0xBA
          ,EM_LINEFROMCHAR:=0xC9

    ;-- Line Index
    if (p_LineIndex<0)
        {
        ;-- Get the line index of the current line
        SendMessage EM_LINEFROMCHAR,-1,0,,ahk_id %hEdit%
        p_LineIndex:=ErrorLevel
        }

    ;-- First visible line
    ;   Note: This section was coded to improve the chances that the rest of the
    ;   function can be skipped.
    SendMessage EM_GETFIRSTVISIBLELINE,0,0,,ahk_id %hEdit%
    FirstVisibleLine:=ErrorLevel
    if (p_LineIndex<=FirstVisibleLine)
        {
        r_ScrollLines:=p_LineIndex-FirstVisibleLine
        Return p_LineIndex=FirstVisibleLine ? True:False
        }

    ;-- Last visible line
    LastVisibleLine:=Edit_GetLastVisibleLine(hEdit,p_Threshold)
    if (p_LineIndex>LastVisibleLine)
        {
        ;-- Determine the maximum line index
        SendMessage EM_GETLINECOUNT,0,0,,ahk_id %hEdit%
        MaxLineIndex:=ErrorLevel-1

        ;-- Calculate the number of scroll lines.  Can be 0 if p_LineIndex is
        ;   invalid.
        r_ScrollLines:=Min(p_LineIndex,MaxLineIndex)-LastVisibleLine
        Return False
        }

    ;-- Line is visible
    r_ScrollLines:=0
    Return True
    }

;------------------------------
;
; Function: Edit_LineLength
;
; Description:
;
;   Get the length of a line of text in the Edit control.
;
; Parameters:
;
;   p_LineIndex - [Optional] The zero-based line index of the desired line.  Set
;       to -1 (the default) or don't specify to use the current line.
;
; Returns:
;
;   The length, in characters, of the specified line.  If p_LineIndex is greater
;   than the index of the last line in the Edit control, the length of the last
;   (or only) line is returned.
;
; Remarks:
;
;   This function returns the length of a line, in characters, as it is
;   displayed in the Edit control.  The length does not include end-of-line
;   (EOL) characters (if any).
;
;-------------------------------------------------------------------------------
Edit_LineLength(hEdit,p_LineIndex:=-1)
    {
    Static Dummy15720896
          ,EM_GETLINECOUNT:=0xBA
          ,EM_LINEINDEX   :=0xBB
          ,EM_LINELENGTH  :=0xC1

    ;-- Get the characters index of the first character of the specified line
    SendMessage EM_LINEINDEX,p_LineIndex,0,,ahk_id %hEdit%
    CharPos:=ErrorLevel<<32>>32  ;-- Convert UInt to Int
    if (CharPos<0)  ;-- Invalid p_LineIndex value
        {
        ;-- Get the index of the last (or only) line
        SendMessage EM_GETLINECOUNT,0,0,,ahk_id %hEdit%
        p_LineIndex:=ErrorLevel-1

        ;-- Get the character index of the first character of the last line
        SendMessage EM_LINEINDEX,p_LineIndex,0,,ahk_id %hEdit%
        CharPos:=ErrorLevel<<32>>32  ;-- Convert UInt to Int
        }

    ;-- Get the line length
    SendMessage EM_LINELENGTH,CharPos,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_LineScroll
;
; Description:
;
;   Scroll the text vertically or horizontally in a multiline Edit control.
;
; Parameters:
;
;   xScroll, yScroll - See the *Scroll* section for more information.
;
; Scroll:
;
;   The optional xScroll and yScroll parameters are used to specify the number
;   of characters to scroll horizontally (xScroll) or the number of lines to
;   scroll vertically (yScroll).  Use a negative number to scroll to the left
;   (xScroll) or to scroll up (yScroll).  Use a positive number to scroll to the
;   right (xScroll) or to scroll down (yScroll).
;
;   If either of these parameters are set to 0 (the default) or are not
;   specified, scrolling for that parameter will not occur.
;
;   Alternatively, these parameters can contain one or more of the following
;   string values:
;
;       (start code)
;       Option  Description
;       ------   -----------
;       Left    Scroll to the left edge of the control.
;       Right   Scroll to the right edge of the control.
;       Top     Scroll to the top of the control.
;       Bottom  Scroll to the bottom of the control.
;       (end)
;
;   If more than one string option is specified, the options must be delimited
;   by a space character.  Ex: "Top Left".  The search for these string values
;   is not case sensitive but invalid string values are ignored.  Be sure to
;   test to thoroughly.
;
;   The xScroll parameter is processed first and then yScroll.  If either of
;   these parameters contains multiple values (Ex: "Top Left"), the values are
;   performed individually from left to right.  If there are conflicting values
;   (Ex: "Top Bottom"), actions for both values will be performed but the last
;   value specified will take precedence because it is performed last.
;
; Remarks:
;
;   This function does not return any values.  If it is important or valuable
;   to identify how many lines have scrolled, use <Edit_Scroll> instead.
;
;-------------------------------------------------------------------------------
Edit_LineScroll(hEdit,xScroll:=0,yScroll:=0)
    {
    Static Dummy34969125

          ;-- Horizontal scroll values
          ,SB_LEFT :=6
          ,SB_RIGHT:=7

          ;-- Vertical scroll values
          ,SB_TOP   :=6
          ,SB_BOTTOM:=7

          ;-- Messages
          ,EM_LINESCROLL:=0xB6
          ,WM_HSCROLL   :=0x114
          ,WM_VSCROLL   :=0x115

    if xScroll  ;-- Not null, blank, or zero (0)
        {
        if xScroll is Integer
            SendMessage EM_LINESCROLL,xScroll,0,,ahk_id %hEdit%
         else
            Loop Parse,xScroll,%A_Space%
                {
                if InStr(A_LoopField,"Left")
                    SendMessage WM_HSCROLL,SB_LEFT,0,,ahk_id %hEdit%
                 else if InStr(A_LoopField,"Right")
                    SendMessage WM_HSCROLL,SB_RIGHT,0,,ahk_id %hEdit%
                 else if InStr(A_LoopField,"Top")
                    SendMessage WM_VSCROLL,SB_TOP,0,,ahk_id %hEdit%
                 else if InStr(A_LoopField,"Bottom")
                    SendMessage WM_VSCROLL,SB_BOTTOM,0,,ahk_id %hEdit%
                }
        }

    if yScroll  ;-- Not null, blank, or zero (0)
        {
        if yScroll is Integer
            SendMessage EM_LINESCROLL,0,yScroll,,ahk_id %hEdit%
         else
            Loop Parse,yScroll,%A_Space%
                {
                if InStr(A_LoopField,"Left")
                    SendMessage WM_HSCROLL,SB_LEFT,0,,ahk_id %hEdit%
                 else if InStr(A_LoopField,"Right")
                    SendMessage WM_HSCROLL,SB_RIGHT,0,,ahk_id %hEdit%
                 else if InStr(A_LoopField,"Top")
                    SendMessage WM_VSCROLL,SB_TOP,0,,ahk_id %hEdit%
                 else if InStr(A_LoopField,"Bottom")
                    SendMessage WM_VSCROLL,SB_BOTTOM,0,,ahk_id %hEdit%
                }
        }
    }

;------------------------------
;
; Function: Edit_Paste
;
; Description:
;
;   Copy the current content of the clipboard to the Edit control at the
;   current caret position.  Text is only inserted if the clipboard contains
;   data in CF_TEXT format.
;
;-------------------------------------------------------------------------------
Edit_Paste(hEdit)
    {
    Static WM_PASTE:=0x302
    SendMessage WM_PASTE,0,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_PosFromChar
;
; Description:
;
;   Get the client area coordinates of a specified character in the Edit
;   control.
;
; Parameters:
;
;   p_CharPos - The zero-based index of the character.
;
;   r_X, r_Y - [Output, Optional] If specified, these variables contain the
;       coordinates of the requested point in the Edit control's client relative
;       to the upper-left corner of the client area.  See the *Remarks and
;       Observations* section for more information.
;
; Returns:
;
;   An AutoHotkey object with the following properties:
;
;   X, Y - The coordinates of the requested point in the Edit control's client
;       relative to the upper-left corner of the client area.
;
;   See the following sections for more information.
;
; End-Of-Line Characters:
;
;   If the specified index (p_CharPos parameter) is for a line delimiter
;   character, the returned coordinates indicate a point just beyond the last
;   visible character of the line.
;
;   For the text in most Edit controls, there are two end-of-line (EOL)
;   characters (CR+LF) per line.  The position of the first EOL character (CR)
;   will be the same as the position of second EOL character (LF).
;
;   What is true for end-of-line characters is also true for any non-printing
;   character in the text.  The position of any non-printing character will be
;   point just beyond the last visible character before the non-printing
;   character.  The position will not change if there are consecutive
;   non-printing characters.
;
; Remarks and Observations:
;
;   If the specified index (p_CharPos parameter) is invalid (Ex: -23) or is
;   greater than the index of the last character in the Edit control, -1 is
;   returned for all output variables and return properties.  To test for this
;   condition, test if both the X and Y return values are exactly -1.
;
;   The return values represent coordinates relative to upper-left corner of the
;   Edit control's client area.  However, the coordinates are for the entire
;   text document, not just the text that is visible in the Edit control.  If
;   the Y coordinate is negative (Ex: -25), the value represent a position
;   higher than the Edit control's client area.  The text must be scrolled up in
;   order for the character position to be visible.  If the Y coordinate is
;   larger that the height of the Edit control, the value represents a position
;   lower than the Edit control's client area.  The text must be scrolled down
;   in order for the character position to be visible.
;
;   In versions of Windows earlier than Windows 10, the return values from
;   EM_POSFROMCHAR may be different than Windows 10+ if the specified index
;   (p_CharPos parameter) is greater than the index of the last character in the
;   Edit control.  If needed, be sure to test this condition on older versions
;   of Windows.
;
;   Sending the EM_POSFROMCHAR message while operating other controls attached
;   to the Edit control (Ex: clicking or dragging the scroll bar) may interfere
;   with the operation of the control.  Usually the interference has only a
;   marginal effect, if anything.  More testing is needed.
;
;-------------------------------------------------------------------------------
Edit_PosFromChar(hEdit,p_CharPos,ByRef r_X:="",ByRef r_Y:="")
    {
    Static EM_POSFROMCHAR:=0xD6

    ;-- Collect position
    SendMessage EM_POSFROMCHAR,p_CharPos,0,,ahk_id %hEdit%

    ;-- Populate output variables
    r_X:=(ErrorLevel&0xFFFF)<<48>>48
        ;-- LOWORD of result and converted from UShort to Short
    r_Y:=(ErrorLevel>>16)<<48>>48
        ;-- HIWORD of result and converted from UShort to Short

    ;-- Return object with coordinate values
    Return {X:r_X,Y:r_Y}
    }

;------------------------------
;
; Function: Edit_ReadFile
;
; Description:
;
;   Load the contents of a file into the Edit control.
;
; Parameters:
;
;   p_File - The path of a text file.
;
;   p_Encoding - [Optional] See the *Encoding* section for more information.
;
;   p_Convert2DOS - [Optional] See the *Convert to DOS* section for more
;       information.
;
;   r_EOLFormat - [Optional] Contains the end-of-line (EOL) format.  If
;       specified, this variable contains the EOL format of the loaded file
;       which will be "DOS", "Unix", or "Mac".  This information is useful if
;       the contents of the Edit control will be converted back to the original
;       EOL format when the file is saved.
;
; Returns:
;
;   The number of characters loaded to Edit control (can be 0) if successful,
;   otherwise -1 if the file could not be found, -2 if the file could not be
;   opened, or -3 if the text could not be loaded to the Edit control (very
;   rare).
;
;   If the function fails, i.e. a negative value is returned (Ex: -1), a
;   developer-friendly message is dumped to the debugger.  Use a debugger or
;   debug viewer to see the message.
;
; Calls To Other Functions:
;
; * <Edit_Convert2DOS>
; * <Edit_SetText>
; * <Edit_SystemMessage>
;
; Convert To DOS:
;
;   By default, the Edit control uses the DOS/Windows end-of-line (EOL) format
;   which consists of the carriage return and a line feed (CR+LF) characters.
;   If a file is not already in the DOS/Windows format, the text will not
;   display correctly when it loaded to the Edit control.
;
;   The p_Convert2DOS parameter determines if the text from the file is
;   converted to the DOS/Windows format before it loaded to the Edit control.
;   Setting this parameter to TRUE will ensure that the text is the correct
;   format for the Edit control.  This conversion is essential if the file is in
;   a Unix (EOL=LF) or Mac (EOL=CR) format but it can also be beneficial if the
;   file is in a DOS/Unix mix where both the  DOS and Unix end-of-line
;   characters are used.
;
;   Conversion requires additional computer resources.  The extra time needed to
;   convert the text is not noticeable for small files, barely noticeable for
;   medium-sized files, but may be very noticeable for large and very large text
;   files.  See <Function Performance> for information on how to improve
;   performance.
;
;   Note: The Mac EOL format (CR) is only used on Mac OS version 9 and earlier.
;   Mac OS 10+ uses the Unix (LF) format.
;
; Encoding:
;
;   The optional p_Encoding parameter is used to specify the character encoding
;   name (Ex: "UTF-8") or code page identifier (Ex: "CP854") to use if the file
;   does not contain a UTF-8 or UTF-16 byte order mark (BOM).  If set to null
;   (the default) or if not specified, the current value of A_FileEncoding is
;   used.  Set to "CP0" to indicate the system default ANSI code page.  A list
;   of valid values for this parameter can be found <here at
;   https://www.autohotkey.com/docs/v1/lib/FileEncoding.htm>.
;
;   When reading a text file using AutoHotkey's standard file commands
;   (*<FileRead at https://www.autohotkey.com/docs/v1/lib/FileRead.htm>*,
;   *<FileReadLine at https://www.autohotkey.com/docs/v1/lib/FileReadLine.htm>*,
;   the *<Read at https://www.autohotkey.com/docs/v1/lib/File.htm#Read>* method
;   of AutoHotkey's File object, etc.), the file's byte order mark (BOM), if it
;   exists, takes precedence over whatever encoding the developer may specify,
;   if anything.  However, if the file has been encoding in some non-ANSI way
;   and file does not contain a byte order mark (BOM), the file will not decoded
;   correctly.  This is mentioned because many common programs/utilities will
;   automatically detect and decode a text file without a BOM, especially if the
;   file contains UTF-8 characters.  AutoHotkey file commands do not include a
;   mechanism to automatically identify and decode a non-ANSI text file so
;   specifying the correct encoding whether the file has a BOM or not is good
;   practice.
;
; Remarks:
;
;   This request will replace the Edit control with the contents of the
;   specified file.  Consequently, the modification flag is cleared and the undo
;   buffer is flushed.
;
;   If the p_Convert2DOS parameter is set to TRUE, the number of characters
;   loaded to the Edit control can be different that the number characters read
;   from the file.
;
;-------------------------------------------------------------------------------
Edit_ReadFile(hEdit,p_File,p_Encoding:="",p_Convert2DOS:=False,ByRef r_EOLFormat:="")
    {
    ;-- File exists?
    IfNotExist %p_File%
        {
        outputdebug Function: %A_ThiSFunc% - File "%p_File%" not found.
        Return -1
        }

    ;-- Open for read
    if not File:=FileOpen(p_File,"r",StrLen(p_Encoding) ? p_Encoding:A_FileEncoding)
        {
        Message:=Edit_SystemMessage(A_LastError)
        outputdebug,
           (ltrim join`s
            Function: %A_ThisFunc% -
            Unexpected return code from FileOpen function.
            A_LastError: %A_LastError% - %Message%
           )

        Return -2
        }

    ;-- Read the contents of the file into a variable
    Text:=File.Read()
    File.Close()

    ;-- Determine EOL format
    if Text Contains `r`n
        r_EOLFormat:="DOS"
     else
        if Text Contains `n
            r_EOLFormat:="UNIX"
         else
            if Text Contains `r
                r_EOLFormat:="MAC"
             else
                r_EOLFormat:="DOS"

    ;-- Convert EOL format?
    if p_Convert2DOS
        Text:=Edit_Convert2DOS(Text)

    ;-- Load text to the Edit control
    if not Edit_SetText(hEdit,Text)
        {
        outputdebug,
           (ltrim join`s
            Function: %A_ThisFunc% -
            Unable to load text to the Edit control
           )

        Return -3
        }

    ;-- Return to sender
    Return StrLen(Text)
    }

;------------------------------
;
; Function: Edit_ReplaceSel
;
; Description:
;
;   Replace the selected text with the specified text.  If nothing is selected,
;   the replacement text is inserted at the caret.
;
; Parameters:
;
;   p_Text - Text to replace selection with.
;
;   p_CanUndo - [Optional] If set to TRUE (the default) or if not specified, the
;       replace can be undone.
;
;-------------------------------------------------------------------------------
Edit_ReplaceSel(hEdit,p_Text:="",p_CanUndo:=True)
    {
    Static EM_REPLACESEL:=0xC2
    SendMessage EM_REPLACESEL,p_CanUndo,&p_Text,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_RemovePasswordChar
;
; Description:
;
;   Remove the password character for the Edit control.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   This function removes the ES_PASSWORD style from the Edit control.  This
;   will remove the password character and it allows the user to see the text
;   in the Edit control.
;
;   See the documentation in <Edit_SetPasswordChar> for more information.
;
;-------------------------------------------------------------------------------
Edit_RemovePasswordChar(hEdit)
    {
    Static EM_SETPASSWORDCHAR:=0xCC
    SPCRC:=DllCall("SendMessageW","UPtr",hEdit,"UInt",EM_SETPASSWORDCHAR,"UInt",0,"UInt",0)
    WinSet Redraw,,ahk_id %hEdit%  ;-- Force the style change to show
    Return SPCRC="FAIL" ? False:SPCRC
    }

;------------------------------
;
; Function: Edit_SelectAll
;
; Description:
;
;   Select all characters in the Edit control.
;
;-------------------------------------------------------------------------------
Edit_SelectAll(hEdit)
    {
    Static EM_SETSEL:=0x0B1
    SendMessage EM_SETSEL,0,-1,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_Scroll
;
; Description:
;
;   Scroll the text vertically in a multiline Edit control.
;
; Parameters:
;
;   p_Pages - [Optional] The number of pages to scroll.  Use a negative number
;       to page up and a positive number to page down.  If set to 0 (the
;       default) or if not specified, the Edit control is not scrolled by page.
;
;   p_Lines - [Optional] The number of lines to scroll.  Use a negative number
;       to scroll up and a positive number to scroll down.  If set to 0 (the
;       default) or if not specified, the Edit control is not scrolled by line.
;
; Returns:
;
;   The number of lines that the command scrolls.  The value will be negative if
;   scrolling up, positive if scrolling down, or zero (0) if no scrolling
;   occurred.
;
; Remarks:
;
;   Although the return value can provide useful information, this function can
;   be very inefficient when scrolling lines.  This is because the function
;   generates a scroll request for each line that is specified.  For example, if
;   the request is to scroll 50 lines, this function will generate up to 50
;   requests to scroll one line at a time.  In most cases, the user will see the
;   animation of the scroll occurring.  When scrolling a large number of lines,
;   the request can take up to 1 second, sometimes longer.  This can be
;   speeded up by increasing the script priority and by turning off redraw while
;   performing the request but in most cases, this is not an efficient use of
;   resources.  When scrolling lines and when the return value from this
;   function is not important, <Edit_LineScroll> should be used instead.
;
;-------------------------------------------------------------------------------
Edit_Scroll(hEdit,p_Pages:=0,p_Lines:=0)
    {
    Static Dummy94758102
          ,EM_SCROLL  :=0xB5
          ,SB_LINEUP  :=0x0     ;-- Scroll up one line
          ,SB_LINEDOWN:=0x1     ;-- Scroll down one line
          ,SB_PAGEUP  :=0x2     ;-- Scroll up one page
          ,SB_PAGEDOWN:=0x3     ;-- Scroll down one page

    ;-- Initialize
    ScrollLines:=0

    ;-- Pages
    Loop % Abs(p_Pages)
        {
        SendMessage EM_SCROLL,p_Pages>0 ? SB_PAGEDOWN:SB_PAGEUP,0,,ahk_id %hEdit%
        if not ErrorLevel
            Break

        ScrollLines+=((ErrorLevel&0xFFFF)<<48>>48)
            ;-- LOWORD of result and converted from UShort to Short
        }

    ;-- Lines
    Loop % Abs(p_Lines)
        {
        SendMessage EM_SCROLL,p_Lines>0 ? SB_LINEDOWN:SB_LINEUP,0,,ahk_id %hEdit%
        if not ErrorLevel
            Break

        ScrollLines+=((ErrorLevel&0xFFFF)<<48>>48)
            ;-- LOWORD of result and converted from UShort to Short
        }

    Return ScrollLines
    }

;------------------------------
;
; Function: Edit_ScrollCaret
;
; Description:
;
;   Scroll the caret into view in the Edit control.
;
; Observations:
;
;   This function does not return until scrolling (if any) has completed.
;
;-------------------------------------------------------------------------------
Edit_ScrollCaret(hEdit)
    {
    Static EM_SCROLLCARET:=0xB7
    SendMessage EM_SCROLLCARET,0,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_ScrollPage
;
; Description:
;
;   Scroll the Edit control by page.
;
; Parameters:
;
;   p_HPages - [Optional] The number of horizontal pages to scroll.  Use a
;       negative number to page left or a positive number to page right.  If
;       set to 0 (the default) or if not specified, the Edit control is not
;       scrolled horizontally.
;
;   p_VPages - [Optional] The number of vertical pages to scroll.  Set to a
;       negative number to page up or positive number to page down.  If set
;       to 0 (the default) or if not specified, the Edit control is not scrolled
;       vertically.
;
; Remarks:
;
;   This function duplicates some of the functionality of <Edit_Scroll>.  If
;   scrolling vertically and the return value is needed, use the <Edit_Scroll>
;   function instead.
;
;------------------------------------------------------------------------------
Edit_ScrollPage(hEdit,p_HPages:=0,p_VPages:=0)
    {
    Static Dummy35246789

          ;-- Horizontal scroll values
          ,SB_PAGELEFT :=2
          ,SB_PAGERIGHT:=3

          ;-- Vertical scroll values
          ,SB_PAGEUP  :=2
          ,SB_PAGEDOWN:=3

          ;-- Messages
          ,WM_HSCROLL :=0x114
          ,WM_VSCROLL :=0x115

    ;-- Horizontal
    Loop % Abs(p_HPages)
        SendMessage WM_HSCROLL,p_HPages>0 ? SB_PAGERIGHT:SB_PAGELEFT,0,,ahk_id %hEdit%

    ;-- Vertical
    Loop % Abs(p_VPages)
        SendMessage WM_VSCROLL,p_VPages>0 ? SB_PAGEDOWN:SB_PAGEUP,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetCaretIndex
;
; Description:
;
;   Set the position of the caret in the Edit control.
;
; Parameters:
;
;   p_CaretIndex - The new zero-based index value of the position of the caret.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.  FALSE is also returned if there is an
;   error.
;
; Requirements:
;
;   Window 10+
;
; Remarks and Observations:
;
;   If the requested caret index is out of the range of the text in the Edit
;   control, the index will be adjusted to fit inside the range of the text.
;
;   If the position of the requested caret index is out of range of the visible
;   text, the Edit control will automatically scroll the text so that caret is
;   visible.
;
;   If text was selected before calling this function, the text will be
;   unselected because of this operation.
;
;-------------------------------------------------------------------------------
Edit_SetCaretIndex(hEdit,p_CaretIndex)
    {
    Static EM_SETCARETINDEX:=0x1511  ;-- ECM_FIRST+17
    SendMessage EM_SETCARETINDEX,p_CaretIndex,0,,ahk_id %hEdit%
    Return ErrorLevel:="FAIL" ? False:True
    }

;------------------------------
;
; Function: Edit_SetCueBanner
;
; Description:
;
;   Set the textual cue, or tip, that is displayed by the Edit control to
;   prompt the user for information.
;
; Parameters:
;
;   p_Text - Cue banner text.
;
;   p_ShowWhenFocused - [Optional] Set to TRUE to show the cue banner even if
;       the Edit control has focus.  The default is FALSE, i.e. don't  show
;       when the Edit control has focus.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   Single-line Edit control only.
;
;-------------------------------------------------------------------------------
Edit_SetCueBanner(hEdit,p_Text,p_ShowWhenFocused:=False)
    {
    Static EM_SETCUEBANNER:=0x1501  ;-- ECM_FIRST+1

    ;-- Initialize
    wText:=p_Text  ;-- Working and Unicode copy

    ;-- Convert to Unicode if needed
    if !A_IsUnicode and StrLen(p_Text)
        {
        VarSetCapacity(wText,StrLen(p_Text)*2,0)
        StrPut(p_Text,&wText,"UTF-16")
        }

    ;-- Set cue banner
    SendMessage EM_SETCUEBANNER,p_ShowWhenFocused,&wText,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_SetDefaultMargins
;
; Description:
;
;   Set the margins of the Edit control to a narrow width calculated using the
;   text metrics of the control's current font.
;
; Remarks:
;
;   The default margins are the same margins that the Edit control had when
;   the control was created.
;
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_SetDefaultMargins(hEdit)
    {
    Static Dummy68391742
          ,EM_SETMARGINS :=0xD3
          ,EC_LEFTMARGIN :=0x1
          ,EC_RIGHTMARGIN:=0x2
          ,EC_USEFONTINFO:=0xFFFF

    Flags  :=EC_LEFTMARGIN|EC_RIGHTMARGIN
    Margins:=EC_USEFONTINFO<<16|EC_USEFONTINFO
    SendMessage EM_SETMARGINS,Flags,Margins,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetEndOfLine
;
; Description:
;
;   Set the end-of-line character used when a linebreak is inserted.
;
; Parameters:
;
;   p_EOLFlag - The end-of-line character used when a linebreak is inserted.
;       See the "Edit Control End-Of-Line Flags" in the Edit_Constants.ahk
;       document for a list of possible flag values.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Requirements:
;
;   Windows 10+
;
; Remarks:
;
;   When the end-of-line character is set to EC_ENDOFLINE_DETECTFROMCONTENT, the
;   Edit control will only detect end-of-line characters supported according to
;   its extended window style.  Use <Edit_SetExtendedStyle> to set the extended
;   styles for the Edit control.
;
;-------------------------------------------------------------------------------
Edit_SetEndOfLine(hEdit,p_EOLFlag)
    {
    Static EM_SETENDOFLINE:=0x150C  ;-- ECM_FIRST+12
    SendMessage EM_SETENDOFLINE,p_EOLFlag,0,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_SetExtendedStyle
;
; Description:
;
;   Set or remove an extended style for the Edit control.
;
; Parameters:
;
;   p_Mask - The mask to limit the style(s) to be set.
;
;   p_ExStyle - The extended style(s) to set.  Set to zero (0) to remove the
;       style(s).
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Requirements:
;
;   Windows 10+
;
; Observations and Programming Notes:
;
;   The EM_SETEXTENDEDSTYLE message sometimes returns 0x10 instead of 0x0 (S_OK)
;   when the request is successful.  This HRESULT code is not documented but it
;   is not an error code.  All HRESULT error codes have 8 hexadecimal digits and
;   they all begin with 0x8...  Ex: 0x80070005.  Sometime 0x10 is returned when
;   removing a style.  Sometimes 0x10 is returned when setting a style that has
;   already been set.  It is unknown exactly what all the conditions are where
;   0x10 is returned.  To accommodate for this undocumented but valid return
;   code, the function has modified to also return TRUE (i.e. successful) when
;   0x10 is returned from the EM_SETEXTENDEDSTYLE message.
;
;-------------------------------------------------------------------------------
Edit_SetExtendedStyle(hEdit,p_Mask,p_ExStyle)
    {
    Static Dummy69574821
          ,S_OK:=0x0
          ,EM_SETEXTENDEDSTYLE:=0x150A

    SendMessage EM_SETEXTENDEDSTYLE,p_Mask,p_ExStyle,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel=S_OK or ErrorLevel=0x10 ? True:False
    }

;------------------------------
;
; Function: Edit_SetFocus
;
; Description:
;
;   Set input focus to the specified Edit control.
;
; Parameters:
;
;   p_ActivateParent - [Optional] If set to TRUE, the function will call
;       <Edit_ActivateParent> which will activate the parent window if it is not
;       already active.  The default is FALSE.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Calls To Other Functions:
;
; * <Edit_ActivateParent>
; * <Edit_HasFocus>
;
; Remarks:
;
;   Functional input focus, aka keyboard focus, can only be achieved if the
;   control has focus _and_ the parent window is active (foremost).  If the
;   p_ActivateParent parameter is set to TRUE, this function will call
;   <Edit_ActivateParent> which will activate the parent window if is not
;   already active before setting input focus on the Edit control.
;
;   This function uses the AutoHotKey *<ControlFocus at
;   https://www.autohotkey.com/docs/v1/lib/ControlFocus.htm>* command to set
;   input focus.  See the AHK documentation for additional considerations (Ex:
;   SetControlDelay).
;
;   For AutoHotkey GUIs, the *<GUIControl at
;   https://www.autohotkey.com/docs/v1/lib/GuiControl.htm>* command can be used
;   instead.  Ex: GUIControl Focus,%hEdit%.  This command will automatically
;   activate the parent window if it not already active before setting input
;   focus on the Edit control.
;
;   This function can be used to set focus on any control.  Just specify the
;   handle to the desired control as the first parameter.  Ex:
;   Edit_SetFocus(hLV) where "hLV" is the handle to a ListView control.
;
; Programming Notes::
;
;   This function does not use the WM_NEXTDLGCTL message because it selects all
;   the text in the Edit control after focus has been set.  This action will
;   undo what may have already been selected and will reposition the caret.  The
;   WM_NEXTDLGCTL message can be useful for other controls (Ex: Buttons) but it
;   is counterproductive when used on the Edit control in most cases.
;
;-------------------------------------------------------------------------------
Edit_SetFocus(hEdit,p_ActivateParent:=False)
    {
    ;-- If requested, activate parent
    if p_ActivateParent
        if not Edit_ActivateParent(hEdit)
            Return False

    ;-- Does the control already have focus?
    if Edit_HasFocus(hEdit)
        Return True

    ;-- Set focus
    ControlFocus,,ahk_id %hEdit%
    Return ErrorLevel ? False:True
    }

;------------------------------
;
; Function: Edit_SetFont
;
; Description:
;
;   Set the font that the Edit control is to use when drawing text.
;
; Parameters:
;
;   hEdit - The handle to the Edit control.
;
;   hFont - The handle to a font.
;
;   p_Redraw - [Optional] Specifies whether the control should be redrawn
;       immediately upon setting the font.  If set to TRUE, the control will
;       redraw itself.  The default is FALSE.
;
; Remarks:
;
;   This function can be used to set the font on any control.  Just specify
;   the handle to the desired control as the first parameter.  Ex:
;   Edit_SetFont(hLV,hFont) where "hLV" is the handle to ListView control.
;
;-------------------------------------------------------------------------------
Edit_SetFont(hEdit,hFont,p_Redraw:=False)
    {
    Static WM_SETFONT:=0x30
    SendMessage WM_SETFONT,hFont,p_Redraw,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetLimitText
;
; Description:
;
;   Set the text limit of the Edit control.
;
; Parameters:
;
;   p_Limit - Set to the maximum number of characters the user can enter.  Ex:
;       10.  Set to 0 to remove the text limit.
;
; Remarks and Programming Notes:
;
;   This function uses the EM_SETLIMITTEXT message to set the text limit.  This
;   message limits the text the user can enter or paste in the Edit control.  It
;   does not truncate any text that is already in the Edit control when the
;   message is sent, nor does it affect the length of the text copied to the
;   Edit control by the WM_SETTEXT message.
;
;   Setting the p_Limit parameter to 0 will set the text limit on the Edit
;   control to the maximum which effectively removes the text limit.  The
;   maximum text limit is 0x7FFFFFFE (2,147,483,646) characters for single-line
;   Edit controls and 0xFFFFFFFF (4,294,967,295) characters for multiline Edit
;   controls.
;
;   For AutoHotkey GUI's, the +Limitnn and -Limit options can be used instead
;   of this function.
;
;   The EM_SETLIMITTEXT message is identical to the EM_LIMITTEXT message.
;
;   At this writing, the Microsoft documentation on the EM_SETLIMITTEXT message
;   is out of date and contains some incorrect information for recent versions
;   of Windows.  The documentation can be found here.
;
;      * <https://learn.microsoft.com/en-us/windows/win32/controls/em-limittext>
;
;   Warning: Although the EM_SETLIMITTEXT message can be sent to any Edit
;   control, not all programs will respond well to a change to the text limit.
;
;-------------------------------------------------------------------------------
Edit_SetLimitText(hEdit,p_Limit)
    {
    Static EM_SETLIMITTEXT:=0xC5
    SendMessage EM_SETLIMITTEXT,p_Limit,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetMargins
;
; Description:
;
;   Set the margins for the Edit control.
;
; Parameters:
;
;   p_LeftMargin, p_RightMargin - See the *Margins* section for more
;       information.
;
; Margins:
;
;   The p_Leftmargin and p_RightMargin parameters are used to set the left
;   and/or right margin, in pixels.  Ex: 10.  If set to null (the default) or if
;   not specified, the margin (left or right) is not modified.
;
;   To set the margin(s) to the default, set to the EC_USEFONTINFO value (0xFFFF
;   or 65535) or to "Default" or "UseFontInfo" (not case sensitive).  When the
;   EC_USEFONTINFO value is used, the EM_SETMARGINS message will set the margin
;   to a narrow width calculated using the text metrics of the control's current
;   font.
;
; Remarks:
;
;   This function uses the EM_SETMARGINS message without modifications.  This
;   message along with companion EM_GETMARGINS message is not DPI aware.  For a
;   DPI aware version of this function, use <Edit_SetMarginsInPixels> instead.
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_SetMargins(hEdit,p_LeftMargin:="",p_RightMargin:="")
    {
    Static Dummy20827935
          ,EM_SETMARGINS :=0xD3
          ,EC_LEFTMARGIN :=0x1
          ,EC_RIGHTMARGIN:=0x2
          ,EC_USEFONTINFO:=0xFFFF

    Flags  :=0
    Margins:=0
    if p_LeftMargin is not Space  ;-- Not null or blank
        {
        if (p_LeftMargin="Default" or p_LeftMargin="UseFontInfo")
            p_LeftMargin:=EC_USEFONTINFO

        if p_LeftMargin is Integer
            {
            Flags  |=EC_LEFTMARGIN
            Margins|=p_LeftMargin       ;-- LOWORD
            }
        }

    if p_RightMargin is not Space  ;-- Not null or blank
        {
        if (p_RightMargin="Default" or p_RightMargin="UseFontInfo")
            p_RightMargin:=EC_USEFONTINFO

        if p_RightMargin is Integer
            {
            Flags  |=EC_RIGHTMARGIN
            Margins|=p_RightMargin<<16  ;-- HIWORD
            }
        }

    if Flags
        SendMessage EM_SETMARGINS,Flags,Margins,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetMarginsInInches
;
; Description:
;
;   Set the width of the left and/or right margin, in inches, for the Edit
;   control.
;
;   S
;
; Parameters:
;
;   p_LeftMargin, p_RightMargin - See the *Margins* section for more
;       information.
;
; Margins:
;
;   The p_Leftmargin and p_RightMargin parameters are used to set the left
;   and/or right margin in inches.  Ex: 0.5 (1/2 inch).  If set to null (the
;   default) or if not specified, the margin (left or right) is not modified.
;
;   To set the margin(s) to the default, set to the EC_USEFONTINFO value (0xFFFF
;   or 65535) or to "Default" or "UseFontInfo" (not case sensitive).  When the
;   EC_USEFONTINFO value is used, the EM_SETMARGINS message will set the margin
;   to a narrow width calculated using the text metrics of the control's current
;   font.
;
; Remarks:
;
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_SetMarginsInInches(hEdit,p_LeftMargin:="",p_RightMargin:="")
    {
    Static Dummy45321786
          ,EM_SETMARGINS :=0xD3
          ,EC_LEFTMARGIN :=0x1
          ,EC_RIGHTMARGIN:=0x2
          ,EC_USEFONTINFO:=0xFFFF

    Flags  :=0
    Margins:=0
    if p_LeftMargin is not Space  ;-- Not null or blank
        {
        if (p_LeftMargin="Default" or p_LeftMargin="UseFontInfo")
            p_LeftMargin:=EC_USEFONTINFO

        if p_LeftMargin is Number
            {
            if (p_LeftMargin>0 and p_LeftMargin!=EC_USEFONTINFO)
                p_LeftMargin:=Round(p_LeftMargin*96)

            Flags  |=EC_LEFTMARGIN
            Margins|=p_LeftMargin       ;-- LOWORD
            }
        }

    if p_RightMargin is not Space  ;-- Not null or blank
        {
        if (p_RightMargin="Default" or p_RightMargin="UseFontInfo")
            p_RightMargin:=EC_USEFONTINFO

        if p_RightMargin is Number
            {
            if (p_RightMargin>0 and p_RightMargin!=EC_USEFONTINFO)
                p_RightMargin:=Round(p_RightMargin*96)

            Flags  |=EC_RIGHTMARGIN
            Margins|=p_RightMargin<<16  ;-- HIWORD
            }
        }

    if Flags
        SendMessage EM_SETMARGINS,Flags,Margins,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetMarginsInPixels
;
; Description:
;
;   Set the width of the left and/or right margin, in pixels, for the Edit
;   control.
;
; Parameters:
;
;   p_LeftMargin, p_RightMargin - [Optional] See the *Margins* section for more
;       information.
;
; Margins:
;
;   The p_Leftmargin and p_RightMargin parameters are used to set the left
;   and/or right margin, in pixels.  Ex: 10.  If set to null (the default) or if
;   not specified, the margin (left or right) is not modified.
;
;   To set the margin(s) to the default, set to the EC_USEFONTINFO value (0xFFFF
;   or 65535) or to "Default" or "UseFontInfo" (not case sensitive).  When the
;   EC_USEFONTINFO value is used, the EM_SETMARGINS message will set the margin
;   to a narrow width calculated using the text metrics of the control's current
;   font.
;
; Remarks and Observations:
;
;   Unlike <Edit_SetMargins>, this function is DPI aware.  The specified margin
;   values (p_LeftMargin and p_RightMargin parameters) will be scaled to the
;   DPI of the current computer.  For example, if the margin is set to 60 and
;   the DPI of the current computer is 120 DPI, the margin will be converted to
;   48.  If needed, the converted value is rounded to the nearest pixel.
;
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_SetMarginsInPixels(hEdit,p_LeftMargin:="",p_RightMargin:="")
    {
    Static Dummy34210897
          ,EM_SETMARGINS :=0xD3
          ,EC_LEFTMARGIN :=0x1
          ,EC_RIGHTMARGIN:=0x2
          ,EC_USEFONTINFO:=0xFFFF

    Flags  :=0
    Margins:=0
    if p_LeftMargin is not Space  ;-- Not null or blank
        {
        if (p_LeftMargin="Default" or p_LeftMargin="UseFontInfo")
            p_LeftMargin:=EC_USEFONTINFO

        if p_LeftMargin is Integer
            {
            if (p_LeftMargin>0 and p_LeftMargin!=EC_USEFONTINFO)
                p_LeftMargin:=Round(p_LeftMargin/A_ScreenDPI*96)

            Flags  |=EC_LEFTMARGIN
            Margins|=p_LeftMargin       ;-- LOWORD
            }
        }

    if p_RightMargin is not Space  ;-- Not null or blank
        {
        if (p_RightMargin="Default" or p_RightMargin="UseFontInfo")
            p_RightMargin:=EC_USEFONTINFO

        if p_RightMargin is Integer
            {
            if (p_RightMargin>0 and p_RightMargin!=EC_USEFONTINFO)
                p_RightMargin:=Round(p_RightMargin/A_ScreenDPI*96)

            Flags  |=EC_RIGHTMARGIN
            Margins|=p_RightMargin<<16  ;-- HIWORD
            }
        }

    if Flags
        SendMessage EM_SETMARGINS,Flags,Margins,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetModify
;
; Description:
;
;   Set or clear the modification flag for the Edit control.
;
; Parameters:
;
;   p_Flag - Set to TRUE to set the modification flag.  Set to FALSE to clear
;       the modification flag.
;
; Remarks:
;
;   The modification flag indicates whether the text within the control has been
;   modified.  Use <Edit_GetModify> to identify the current state of the
;   modification flag.
;
;-------------------------------------------------------------------------------
Edit_SetModify(hEdit,p_Flag)
    {
    Static EM_SETMODIFY:=0xB9
    SendMessage EM_SETMODIFY,p_Flag,0,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetPasswordChar
;
; Description:
;
;   Set or remove the password character for the Edit control.
;
; Parameters:
;
;   p_CharValue - [Optional] The decimal value of the character that is
;       displayed in place of the characters typed by the user.  The default is
;       an 9679 (black circle).  Set to 0 to remove the password character.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Credit:
;
;   The code for this function was adapted from a post on the old AutoHotkey
;   forum.  The link has been lost.  Author: Unknown
;
; Remarks:
;
;   This function adds, retains, or removes, the ES_PASSWORD style to/from the
;   Edit control and it sets, changes, or removes, the password character.
;
;   Once the Edit control has been created, the EM_SETPASSWORDCHAR message is
;   the only way to add or remove the ES_PASSWORD style from the Edit control.
;
;   Starting with AutoHotkey v1.1.24.04, this function is no longer needed to
;   remove the password character from the Edit control on an AutoHotkey GUI.
;   Use the standard *<GUIControlGet at
;   https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command to remove
;   the password character.  Ex: GUIControl -Password,%hEdit%
;
;   Most versions of Windows (XP+) can display Unicode characters on the
;   Edit control even if using the ANSI version of AutoHotkey.  For this reason,
;   this function should work correctly for all versions of AutoHotkey (ANSI and
;   Unicode) even when Unicode character values are specified.  Be sure to test
;   thoroughly.
;
;   On Windows 2000+, the ES_PASSWORD style cannot be removed once added unless
;   the request is made from the same process that created the control.
;
; Observations:
;
;   Although the Microsoft documentation and much of the documentation in this
;   library indicates that ES_PASSWORD style is only for single-line Edit
;   control, this style appears to also work on multiline Edit controls starting
;   with Window 7.
;
;-------------------------------------------------------------------------------
Edit_SetPasswordChar(hEdit,p_CharValue:=9679)
    {
    Static EM_SETPASSWORDCHAR:=0xCC
    SPCRC:=DllCall("SendMessageW","UPtr",hEdit,"UInt",EM_SETPASSWORDCHAR,"UInt",p_CharValue,"UInt",0)
    WinSet Redraw,,ahk_id %hEdit%  ;-- Force style change to show
    Return SPCRC="FAIL" ? False:SPCRC
    }

;------------------------------
;
; Function: Edit_SetReadOnly
;
; Description:
;
;   Add or remove the read-only style (ES_READONLY) on the Edit control.
;
; Parameters:
;
;   p_Flag - [Optional] Set to TRUE (the default) or don't specify to add the
;       ES_READONLY style.  Set to FALSE to remove the ES_READONLY style.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   For AutoHotkey GUIs, AutoHotkey commands and the +ReadOnly or -ReadOnly
;   options can be used to add or remove the read-only style.  When creating the
;   Edit control, just include one of these options.  Ex: gui,Add,Edit,w200
;   +ReadOnly.  After the Edit control has been created, use the *<GUIControlGet
;   at https://www.autohotkey.com/docs/v1/lib/GuiControlGet.htm>* command to add
;   or remove the read-only style.  Ex: GUIControl -ReadOnly,%hEdit%
;
;-------------------------------------------------------------------------------
Edit_SetReadOnly(hEdit,p_Flag:=True)
    {
    Static EM_SETREADONLY:=0xCF
    SendMessage EM_SETREADONLY,p_Flag,0,,ahk_id %hEdit%
    Return ErrorLevel ? True:False
    }

;------------------------------
;
; Function: Edit_SetTabStops
;
; Description:
;
;   Set the tab stops in a multiline Edit control.  When text is copied to the
;   control, any tab character in the text causes space to be generated up to
;   the next tab stop.
;
; Parameters:
;
;   p_NbrOfTabStops - [Optional] The number of tab stops.  Set to 0 (the
;       default) or don't specify to set the tab stops to the system default.
;       Set to 1 to have all tab stops set to the value of the p_DTU parameter
;       or 32 if the p_DTU parameter is not specified.  Any value greater than 1
;       will set that number of tab stops.
;
;   p_DTU - [Optional] See the *Dialog Template Units* section for more
;       information.
;
; Returns:
;
;   TRUE if tab stops are set, otherwise FALSE.
;
; Dialog Template Units:
;
;   The optional p_DTU parameter is used to the set the tab stops for the Edit
;   control.  The values used in the parameter are measured in Dialog Template
;   Units.
;
;   This parameter can contain a single value (Ex: 32), a string with a
;   comma-delimited list of values (Ex: "29,72,122,174"), or an AutoHotkey
;   object with a simple array of values (Ex: [150,180,205,255].  If this
;   parameter is not specified, the default is 32.
;
;   If the p_NbrOfTabStops parameter is set to 0 (the default), this parameter
;   is ignored.  If this parameter contains a single value (Ex: 30), all tab
;   stops will be set to a factor of that value (Ex: 30, 60, 90, etc.).
;   Otherwise, this parameter should contain values for all requested tab stops.
;
; Remarks:
;
;   See <Custom Tab Stops and Zoom> for more information.
;
;-------------------------------------------------------------------------------
Edit_SetTabStops(hEdit,p_NbrOfTabStops:=0,p_DTU:=32)
    {
    Static EM_SETTABSTOPS:=0xCB
    VarSetCapacity(TabStops,p_NbrOfTabStops*4,0)
    if IsObject(p_DTU)
        {
        NbrOfElements:=0  ;-- Not assuming a correctly formed simple array
        For Key,Value in p_DTU
            {
            NbrOfElements++
            if (A_Index<=p_NbrOfTabStops)
                NumPut(Value+0,TabStops,(A_Index-1)*4,"UInt")
            }

        if (NbrOfElements=1 and p_NbrOfTabStops>1)
            Loop %p_NbrOfTabStops%
                NumPut(Value*A_Index,TabStops,(A_Index-1)*4,"UInt")
        }
     else  ;-- A single value or a string of comma-delimited values
        if p_DTU Contains ,,
            {
            Loop Parse,p_DTU,`,,%A_Space%
                if (A_Index<=p_NbrOfTabStops)
                    NumPut(A_LoopField+0,TabStops,(A_Index-1)*4,"UInt")
            }
         else
            Loop %p_NbrOfTabStops%
                NumPut(p_DTU*A_Index,TabStops,(A_Index-1)*4,"UInt")

    SendMessage EM_SETTABSTOPS,p_NbrOfTabStops,&TabStops,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_SetText
;
; Description:
;
;   Set the text in the Edit control.
;
; Parameters:
;
;   p_Text - Text to set in the Edit control.
;
;   p_SetModify - [Optional] See the *Modification Flag* section for more
;       information.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Modification Flag:
;
;   The system automatically clears the modification flag whenever the Edit
;   control receives a WM_SETTEXT message.  The optional p_SetModify parameter
;   can be used to effectively override this behavior.
;
;   If the p_SetModify parameter is set to TRUE, the Edit control's modification
;   flag is set after the text is set.  If set to FALSE (the default) or if not
;   specified, the modification flag is not set (remains cleared) after the text
;   is set.
;
; Remarks:
;
;   This function uses the WM_SETTEXT message to replace all the text in the
;   Edit control (if any) with the text contained in the p_Text parameter.  The
;   WM_SETTEXT message also clears the modification and flushes the undo buffer.
;
;   If the ability to undo replacing all the text in the Edit is required, one
;   workaround is to call <Edit_SelectAll> to select all the text and then call
;   <Edit_ReplaceSel> to set the text in Edit control instead of calling this
;   function.  There may be some screen flickering when these functions are
;   called but the user can undo the change if needed.
;
;   This function is similar to the AutoHotkey *<ControlSetText at
;   https://www.autohotkey.com/docs/v1/lib/ControlSetText.htm>* command except
;   there is no delay after the command has executed.
;
;-------------------------------------------------------------------------------
Edit_SetText(hEdit,p_Text,p_SetModify:=False)
    {
    Static WM_SETTEXT:=0xC
    SendMessage WM_SETTEXT,0,&p_Text,,ahk_id %hEdit%
    if STRC:=ErrorLevel  ;-- Text set successfully
        if p_SetModify
            Edit_SetModify(hEdit,True)

    Return STRC  ;-- Return code from the WM_SETTEXT message
    }

;------------------------------
;
; Function: Edit_SetSel
;
; Description:
;
;   Select a range of characters in the Edit control.
;
; Parameters:
;
;   p_StartSelPos - [Optional] The zero-based index of the start character of
;       the selection.  If set to 0 (the default) or if not specified, the index
;       of the first character in the Edit control is used.  If set to -1, the
;       current selection (if any) will be deselected.
;
;   p_EndSelPos - [Optional] The zero-based index of the end character of the
;       selection plus 1.  If set to -1 (the default) or if not specified, the
;       position of the last character in the control is used.  Exception: If
;       p_StartSelPos is set to -1, this parameter is ignored.
;
; Remarks:
;
;   The start value can be greater than the end value.  If this occurs, the
;   lower of the two values becomes the start position and the remaining becomes
;   the end position.
;
;   If the Edit control has the ES_NOHIDESEL style, the selected text is
;   highlighted regardless of whether the control has focus.  Without the
;   ES_NOHIDESEL style, the selected text is only highlighted when the Edit
;   control has focus.
;
;-------------------------------------------------------------------------------
Edit_SetSel(hEdit,p_StartSelPos:=0,p_EndSelPos:=-1)
    {
    Static EM_SETSEL:=0xB1
    SendMessage EM_SETSEL,p_StartSelPos,p_EndSelPos,,ahk_id %hEdit%
    }

;------------------------------
;
; Function: Edit_SetStyle
;
; Description:
;
;   Add, remove, or toggle a style for the Edit control.
;
; Parameters:
;
;   p_Style - See the *Style* section for more information.
;
;   p_Option - [Optional] Set to "+" (the default) or "Add" to add the style,
;       "-" or "Remove" to remove the style, or "^" or "Toggle" to toggle the
;       style.
;
; Returns:
;
;   TRUE if the request completed successfully, otherwise FALSE.
;
;   FALSE is also returned if the p_Style or p_Option parameters contain an
;   invalid value.  A developer-friendly message is dumped to debugger if this
;   occurs.
;
; Style:
;
;   The p_Style parameter is used to identify which style will be added,
;   removed, or toggled by this function.  See the function's static variables
;   for a list of possible style values.
;
;   Alternatively, a style name can be specified.  Ex: "Uppercase".  The
;   Microsoft style constant name less the "ES_" prefix can be used as the style
;   name.  For example, for the ES_NUMBER constant, "Number" can be used as the
;   style name.  If this syntax is used and an invalid style name is specified,
;   the function will fail (return FALSE).  Be sure to test thoroughly to ensure
;   that a valid style name is used.
;
; Remarks:
;
;   There are only a few styles that can be modified by this function after
;   the Edit control has been created.  See the function's static variables for
;   a list.
;
;   If needed, use <Edit_IsStyle> to determine if a style is currently set.
;
;-------------------------------------------------------------------------------
Edit_SetStyle(hEdit,p_Style,p_Option:="+")
    {
    Static Dummy56432908

          ;-- Styles that can be modified after the Edit control has been
          ;   created
          ,ES_UPPERCASE :=0x8
          ,ES_LOWERCASE :=0x10
          ,ES_PASSWORD  :=0x20
                ;-- This style cannot be set here.  Use the Edit_SetPasswordChar
                ;   function to add or remove this style.
          ,ES_OEMCONVERT:=0x400
          ,ES_READONLY  :=0x800
                ;-- This style cannot be set here.  Use the Edit_SetReadOnly
                ;   function to add or remove this style.
          ,ES_WANTRETURN:=0x1000
          ,ES_NUMBER    :=0x2000

    ;-- If needed, convert the style name
    if p_Style is not Integer
        {
        ;-- Convert
        if p_Style is not Space ;-- Not null or blank
            if p_Style is Alpha ;-- Alpha characters only
                if ES_%p_Style% is not Space
                    p_Style:=ES_%p_Style%

        ;-- Bounce and return FALSE if invalid
        if p_Style is not Integer
            {
            outputdebug,
               (ltrim join`s
                Function: %A_ThisFunc% - Invalid p_Style name: %p_Style%
               )

            Return False
            }
        }

    ;-- If needed, convert the p_Option parameter
    if p_Option is Alpha
        {
        if (p_Option="Add")
            p_Option:="+"
         else if (p_Option="Remove")
            p_Option:="-"
         else if (p_Option="Toggle")
            p_Option:="^"

        ;-- Bounce and return FALSE if p_Option contains an invalid value
        if p_Option not in +,-,^
            {
            outputdebug,
               (ltrim join`s
                Function: %A_ThisFunc% - Invalid p_Option name: %p_Option%
               )

            Return False
            }
        }

    ;-- Add, remove, or toggle style
    Control Style,%p_Option%%p_Style%,,ahk_id %hEdit%
    Return ErrorLevel ? False:True
    }

;------------------------------
;
; Function: Edit_SetZoom
;
; Description:
;
;   Set the zoom ratio for a multiline Edit control.
;
; Parameters:
;
;   p_Numerator - The numerator of the zoom ratio.  Ex: 200.  If the denominator
;       (p_Denominator parameter) is assumed to be 100 (the default), this
;       parameter can used to set a value that represents a zoom ratio
;       percentage.
;
;   p_Denominator - [Optional] The denominator of the zoom ratio.  The default
;       is 100.
;
; Returns:
;
;   TRUE if the new zoom setting is accepted, otherwise FALSE.  FALSE is also
;   returned if there is an error.
;
; Requirements:
;
;   Windows 10+
;
; Remarks and Observations:
;
;   The Edit control must have the ES_EX_ZOOMABLE extended style set for this
;   function to have an effect.
;
;   Sometimes there are no noticeable changes to the Edit control if a very
;   small change to the zoom ratio is specified (Ex: Change from 19% to 20%) or
;   if a very small zoom ratio is specified (less than ~10%) or if a very large
;   zoom ratio is specified (Ex: 2000%).  Testing is recommended.
;
;   The original font size will determine the minimum zoom ratio.  For most
;   fonts, the minimum is 2% and for others it is 3% or more.  The smaller the
;   size of the original font, the larger the minimum zoom ratio will be.  10%
;   is usually a good minimum zoom ratio to use.
;
;   If custom tab stops are set, there may be issues when the Edit control is
;   zoomed.  See <Custom Tab Stops and Zoom> for more information.
;
;   If custom margins have been set, zooming will reset the margins to the
;   default.  See <Margins> for more information.
;
;   See <Zoom Issues> for more zoom issues.
;
;-------------------------------------------------------------------------------
Edit_SetZoom(hEdit,p_Numerator,p_Denominator:=100)
    {
    Static EM_SETZOOM:=0x4E1  ;-- WM_USER+225
    SendMessage EM_SETZOOM,p_Numerator,p_Denominator,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_Show
;
; Description:
;
;   Show the Edit control if it was previously hidden.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Remarks:
;
;   For AutoHotkey GUIs, use the *<GUIControl at
;   https://www.autohotkey.com/docs/v1/lib/GuiControl.htm>* command for improved
;   efficiency.  Ex: GUIControl MyGUI:Show,%hEdit%
;
;-------------------------------------------------------------------------------
Edit_Show(hEdit)
    {
    Control Show,,,ahk_id %hEdit%
    Return ErrorLevel ? False:True
    }

;------------------------------
;
; Function: Edit_ShowBalloonTip
;
; Description:
;
;   Display a balloon tip associated with the Edit control.
;
; Parameters:
;
;   p_Title - The balloon tip title.  Set to null for no title.
;
;   p_Text - The balloon tip text.
;
;   p_Icon - [Optional] The type of icon to associate with the balloon tip.  See
;       the *Icon* section for more information.
;
; Returns:
;
;   TRUE if successful, otherwise FALSE.
;
; Icon:
;
;   The optional p_Icon parameter is used to specify the type of icon to
;   associate with the balloon tip.  The default is 0 (no icon).  See the
;   function's static variables for a list of possible values.
;
;   Alternatively, an icon type "name" can be specified.  Ex: "Info".  The
;   Microsoft icon type constant name less the "TTI_" prefix can be used as the
;   icon type name.  For example, for the TTI_WARNING constant, "Warning" can be
;   used as the icon type name.  If this syntax is used and an invalid icon type
;   name is specified, TTI_NONE (0) will be set.  Be sure to test thoroughly to
;   ensure that a valid icon type name is used.
;
; Remarks:
;
;   Sending the EM_SHOWBALLOONTIP message will automatically move focus to the
;   designated Edit control.
;
;   If specified, the balloon tip icon (p_Icon parameter) will not be displayed
;   unless a balloon tip title (p_Title parameter) is also specified.
;
;   Important: A balloon tip will not show if the *EnableBalloonTips* registry
;   key is disabled (set to 0).  The key can be found here:
;
;     * HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\
;
;   Note: On Windows 10+, balloon tips are enabled by default and the
;   *EnableBalloonTips* registry key does not exist.  It can be created manually
;   if desired.  Set the value to 0 to disable.  Set the value to 1 to enable.
;   Disabling balloon tips is usually not recommended.
;
; Observations:
;
;   The EM_SHOWBALLOONTIP message does not fail (return FALSE) if the
;   "EnableBalloonTips" registry key is disabled (set to 0).
;
;-------------------------------------------------------------------------------
Edit_ShowBalloonTip(hEdit,p_Title,p_Text,p_Icon:=0)
    {
    Static Dummy81452369

          ;-- p_Icon values
          ,TTI_NONE   :=0
          ,TTI_INFO   :=1
          ,TTI_WARNING:=2
          ,TTI_ERROR  :=3

          ;-- p_Icon values
          ,TTI_INFO_LARGE   :=4
          ,TTI_InfoLarge    :=4  ;-- Alias
          ,TTI_WARNING_LARGE:=5
          ,TTI_WarningLarge :=5  ;-- Alias
          ,TTI_ERROR_LARGE  :=6
          ,TTI_ErrorLarge   :=6  ;-- Alias

          ;-- Messages
          ,EM_SHOWBALLOONTIP:=0x1503  ;-- ECM_FIRST+3

    ;-- Working and Unicode copies the title and text
    wTitle:=p_Title
    wText :=p_Text

    ;-- If necessary, convert the title and text to Unicode
    if not A_IsUnicode
        {
        if StrLen(p_Title)
            {
            VarSetCapacity(wTitle,StrLen(p_Title)*2,0)
            StrPut(p_Title,&wTitle,"UTF-16")
            }

        if StrLen(p_Text)
            {
            VarSetCapacity(wText,StrLen(p_Text)*2,0)
            StrPut(p_Text,&wText,"UTF-16")
            }
        }

    ;-- If needed, convert string icon flag
    if p_Icon is not Integer
        {
        p_Icon:=Trim(p_Icon," `f`n`r`t`v")
        StringReplace p_Icon,p_Icon,%A_Space%,,All
        StringReplace p_Icon,p_Icon,-,,All
        StringReplace p_Icon,p_Icon,_,,All

        ;-- Convert
        if p_Icon is not Space ;-- Not null or blank
            if p_Icon is Alpha ;-- Alpha characters only
                if TTI_%p_Icon% is not Space
                    p_Icon:=TTI_%p_Icon%

        ;-- Set to TTI_NONE if invalid
        if p_Icon is not Integer
            p_Icon:=TTI_NONE
        }


    ;-- Create and populate the EDITBALLOONTIP structure
    cbSize:=A_PtrSize=8 ? 32:16
    VarSetCapacity(EDITBALLOONTIP,cbSize)
    NumPut(cbSize, EDITBALLOONTIP,0,"Int")
    NumPut(&wTitle,EDITBALLOONTIP,A_PtrSize=8 ? 8:4,"UPtr")
    NumPut(&wText, EDITBALLOONTIP,A_PtrSize=8 ? 16:8,"UPtr")
    NumPut(p_Icon, EDITBALLOONTIP,A_PtrSize=8 ? 24:12,"Int")

    ;-- Show it
    SendMessage EM_SHOWBALLOONTIP,0,&EDITBALLOONTIP,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_SystemMessage
;
; Description:
;
;   Convert a system message number into a readable message.
;
; Type:
;
;   Internal function.  Subject to change.
;
;-------------------------------------------------------------------------------
Edit_SystemMessage(p_MessageNbr)
    {
    Static FORMAT_MESSAGE_FROM_SYSTEM:=0x1000

    ;-- Convert system message number into a readable message
    VarSetCapacity(Message,1024*(A_IsUnicode ? 2:1),0)
    DllCall("FormatMessage"
           ,"UInt",FORMAT_MESSAGE_FROM_SYSTEM           ;-- dwFlags
           ,"UPtr",0                                    ;-- lpSource
           ,"UInt",p_MessageNbr                         ;-- dwMessageId
           ,"UInt",0                                    ;-- dwLanguageId
           ,"Str",Message                               ;-- lpBuffer
           ,"UInt",1024                                 ;-- nSize (in TCHARS)
           ,"UPtr",0)                                   ;-- *Arguments

    ;-- Remove trailing CR+LF, if defined
    if (SubStr(Message,-1)="`r`n")
        StringTrimRight Message,Message,2

    ;-- Return system message
    Return Message
    }

;------------------------------
;
; Function: Edit_TextIsSelected
;
; Description:
;
;   Return TRUE if any text is selected, otherwise FALSE.
;
; Parameters:
;
;   r_StartSelPos, r_EndSelPos - [Output, Optional] If defined, these variables
;       contain the start and end character positions of the current selection
;       in the Edit control.
;
;-------------------------------------------------------------------------------
Edit_TextIsSelected(hEdit,ByRef r_StartSelPos:="",ByRef r_EndSelPos:="")
    {
    Static Dummy4509876
          ,s_StartSelPos
          ,s_EndSelPos
          ,Dummy1:=VarSetCapacity(s_StartSelPos,4)
          ,Dummy2:=VarSetCapacity(s_EndSelPos,4)

          ;-- Message
          ,EM_GETSEL:=0xB0

    ;-- Get select positions
    SendMessage EM_GETSEL,&s_StartSelPos,&s_EndSelPos,,ahk_id %hEdit%
    r_StartSelPos:=NumGet(s_StartSelPos,0,"UInt")
    r_EndSelPos  :=NumGet(s_EndSelPos,0,"UInt")
    Return r_StartSelPos=r_EndSelPos ? False:True
    }

;------------------------------
;
; Function: Edit_Undo
;
; Description:
;
;   Undo the last operation.
;
; Returns:
;
;   For a single-line Edit control, the return value is always TRUE.  For a
;   multiline Edit control, the return value is TRUE if the undo operation is
;   successful, otherwise FALSE.
;
;-------------------------------------------------------------------------------
Edit_Undo(hEdit)
    {
    Static EM_UNDO:=0xC7
    SendMessage EM_UNDO,0,0,,ahk_id %hEdit%
    Return ErrorLevel
    }

;------------------------------
;
; Function: Edit_WriteFile
;
; Description:
;
;   Write the contents of the Edit control to a file.  See the *File Processing*
;   section for more information.
;
; Parameters:
;
;   p_File - The file path.
;
;   p_Encoding - [Optional] The character encoding name (Ex: "UTF-16") or code
;       page identifier (Ex: "CP854") to encode the text written to the file.
;       Set to "CP0" to force the program to use the system default ANSI code
;       page.  If set null (the default) or if not specified, the current value
;       of A_FileEncoding is used.  A list of valid values for this parameter
;       can be found <here at
;       https://www.autohotkey.com/docs/v1/lib/FileEncoding.htm>.
;
;   p_Convert - [Optional] Convert end-of-line (EOL) format.  See the *Convert
;       End-Of-Line format* section for more information.
;
; Returns:
;
;   The number bytes (not characters) written to the file (can be zero),
;   otherwise -1 if the file could not be created or opened.
;
; Calls To Other Functions:
;
; * <Edit_Convert2DOS>
; * <Edit_Convert2Mac>
; * <Edit_Convert2Unix>
; * <Edit_GetText>
; * <Edit_SystemMessage>
;
; Convert End-Of-Line format::
;
;   The optional p_Convert parameter is used to instruct the function to
;   convert the text to a new/different end-of-line (EOL) format before the
;   text is written to the file.
;
;   If set to null (the default) or if not specified, no conversion is performed
;   before the text is written to the file.
;
;   If set to "D", "DOS", "W", or "Windows", the text is converted to the
;   DOS/Windows end-of-line format (CR+LF) before the text is written to the
;   file.  This format might be useful if the integrity of the end-of-line
;   character(s) is unknown or if converting from another EOL format.
;
;   If set to "M" or "Mac", the text is converted to the old Mac (OS 9 and
;   earlier) format (CR) before the text is written to the file.  This
;   end-of-line format is rarely used anymore.
;
;   If set to "U" or "Unix", the text is converted to the Unix format (LF)
;   before the text is written to the file.  This end-of-line format is used by
;   Unix and Mac (OS 10+).
;
;   Conversion requires additional computer resources.  The extra time needed to
;   convert the text is insignificant for small files, barely noticeable for
;   medium-sized files, but may be very noticeable for large and very large text
;   files.  See <Function Performance> for information on how to improve
;   performance.
;
; File Processing:
;
;   If the file (p_File parameter) does not exist, it will be created and the
;   contents of the Edit control will be written to the file.
;
;   If the file already exists, the contents of the file will be overwritten
;   with the contents of the Edit control.  All other attributes of the file are
;   not modified.  This includes the standard attributes like creation date but
;   if the file is on NTFS, it can include permissions, compression, encryption,
;   properties, etc.  To force a new file to be created, the existing file must
;   be deleted before calling this function.
;
;   In all cases, a byte order mark (BOM) is automatically added to the
;   beginning of the file if encoding (p_Encoding parameter or A_FileEncoding if
;   p_Encoding is null) is set to UTF-8 or UTF-16.
;
; Remarks:
;
;   If the function fails, i.e. returns -1, a developer-friendly message is
;   dumped to the debugger.  Use a debugger or debug viewer to see the message.
;
;-------------------------------------------------------------------------------
Edit_WriteFile(hEdit,p_File,p_Encoding:="",p_Convert:="")
    {
    ;-- Open file for write
    ;   Note: The file is created if it doesn't exist.  Otherwise, it is
    ;   overwritten.
    if not File:=FileOpen(p_File,"w",StrLen(p_Encoding) ? p_Encoding:A_FileEncoding)
        {
        Message:=Edit_SystemMessage(A_LastError)
        outputdebug,
           (ltrim join`s
            Function: %A_ThisFunc% -
            Unexpected return code from FileOpen function.
            A_LastError: %A_LastError% - %Message%
           )

        Return -1
        }

    ;-- Get text from the Edit control
    Text:=Edit_GetText(hEdit)

    ;-- If requested, convert EOL format
    if p_Convert
        {
        StringUpper,p_Convert,p_Convert,T  ;-- jic StringCaseSense is On
        if p_Convert in D,Dos,W,Windows
            Text:=Edit_Convert2DOS(Text)
         else if p_Convert in U,Unix
            Text:=Edit_Convert2Unix(Text)
         else if p_Convert in M,Mac
            Text:=Edit_Convert2Mac(Text)
         else
            outputdebug,
               (ltrim join`s
                Function: %A_ThisFunc% - Unsupported EOL format specified:
                %p_Convert%. Conversion not performed.
               )
        }

    ;-- Save to file
    NumberOfBytesWritten:=File.Write(Text)

    ;-- Close file and return
    File.Close()
    Return NumberOfBytesWritten
    }

;------------------------------
;
; Function: Edit_ZoomIn
;
; Description:
;
;   Increase the zoom ratio of the Edit control by a specified percentage.
;
; Parameters:
;
;   p_IncrementPct - [Optional] The amount to increase the zoom ratio as a
;       percentage.  The default is 10.  The default increment value matches the
;       zoom increment value used by the built-in Ctrl+WheelUp keyboard
;       shortcut.
;
;   p_MaxZoomPct - [Optional] The maximum zoom ratio as a percentage.  The
;       default is 9999, i.e. no practical maximum limit.
;
; Returns:
;
;   TRUE if the new zoom setting is accepted, otherwise FALSE.  FALSE is also
;   returned if zoom is not enabled or if there is an error (rare).
;
; Requirements:
;
;   Windows 10+
;
; Calls To Other Functions:
;
; * <Edit_GetZoom>
;
; Remarks:
;
;   The Edit control must have the ES_EX_ZOOMABLE extended style set for this
;   function to have an effect.
;
;   See <Margins> for more information.
;
; Observations:
;
;   Zooming in using the built-in Ctrl+WheelUp keyboard shortcut is limited to a
;   maximum zoom ratio of 500% of the original font size.  The EM_SETZOOM
;   message used by this function does not have that limitation.  The maximum
;   zoom ratio varies from font to font but it is usually from 1000% to 1500%,
;   sometimes a little more.  Once the maximum zoom ratio is reached, the
;   EM_SETZOOM message will have no effect on the font size used on the Edit
;   control.
;
;   There is no programmatic method to identify when the maximum zoom ratio for
;   a font has been reached.  Well, nothing that is very accurate.  After the
;   maximum zoom ratio has been reached, the font size will not change but the
;   Edit control will continue to accept requests to zoom in.  When zooming out,
;   the font size will not change until the zoom ratio is less than the maximum.
;
;   While the zoom ratio is less than 500%, the built-in Ctrl+WheelUp keyboard
;   shortcut will continue to work.  If 500% or greater, the Ctrl+WheelUp
;   keyboard shortcut will have no effect.
;
;-------------------------------------------------------------------------------
Edit_ZoomIn(hEdit,p_IncrementPct:=10,p_MaxZoomPct:=9999)
    {
    Static EM_SETZOOM:=0x4E1  ;-- WM_USER+225

    ;-- Get the current zoom factor
    ;   Bounce if there is no zoom percent (Error or zoom not enabled)
    Edit_GetZoom(hEdit,Numerator,Denominator,ZoomPct)
    if (ZoomPct=0)
        Return False

    ;-- If needed, reset values
    if (Numerator=Denominator)
        Numerator:=Denominator:=100

    ;-- Zoom in by the specified percentage
    Numerator:=Min(p_MaxZoomPct,Numerator+p_IncrementPct)

    ;-- Set zoom
    SendMessage EM_SETZOOM,Numerator,100,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_ZoomOut
;
; Description:
;
;   Decrease the zoom ratio of the Edit control by a specified percentage.
;
; Parameters:
;
;   p_DecrementPct - [Optional] The amount to decrease the zoom ratio as a
;       percentage.  The default is 10.  The default decrement value matches the
;       zoom decrement value used by the built-in Ctrl+WheelDown keyboard
;       shortcut.
;
;   p_MinZoomPct - [Optional] The minimum zoom ratio as a percentage.  The
;       default is 10.
;
; Returns:
;
;   TRUE if the new zoom setting is accepted, otherwise FALSE.  FALSE is also
;   returned if zoom is not enabled or if there is an error (rare).
;
; Requirements:
;
;   Windows 10+
;
; Calls To Other Functions:
;
; * <Edit_GetZoom>
;
; Remarks:
;
;   The Edit control must have the ES_EX_ZOOMABLE extended style set for this
;   function to have an effect.
;
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_ZoomOut(hEdit,p_DecrementPct:=10,p_MinZoomPct:=10)
    {
    Static EM_SETZOOM:=0x4E1  ;-- WM_USER+225

    ;-- Get the current zoom factor
    ;   Bounce if there is no zoom percent (Error or zoomable not enabled)
    Edit_GetZoom(hEdit,Numerator,Denominator,ZoomPct)
    if (ZoomPct=0)
        Return False

    ;-- If needed, reset values
    if (Numerator=Denominator)
        Numerator:=Denominator:=100

    ;-- Zoom out
    Numerator:=Max(p_MinZoomPct,Numerator-p_DecrementPct)

    ;-- Set zoom
    SendMessage EM_SETZOOM,Numerator,100,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }

;------------------------------
;
; Function: Edit_ZoomReset
;
; Description:
;
;   Set the zoom ratio on the Edit control to 100% which will restore the font
;   to the original font size.
;
; Returns:
;
;   TRUE if the zoom setting is accepted, otherwise FALSE.  FALSE is also
;   returned if there is an error.
;
; Requirements:
;
;   Windows 10+
;
; Remarks:
;
;   The Edit control must have the ES_EX_ZOOMABLE extended style set for this
;   function to have an effect.
;
;   See <Margins> for more information.
;
;-------------------------------------------------------------------------------
Edit_ZoomReset(hEdit)
    {
    Static EM_SETZOOM:=0x4E1  ;-- WM_USER+225
    SendMessage EM_SETZOOM,100,100,,ahk_id %hEdit%
    Return ErrorLevel="FAIL" ? False:ErrorLevel
    }
