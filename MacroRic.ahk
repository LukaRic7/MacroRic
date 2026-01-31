; =============================
;   GLOBALS
; =============================

; === Directives ===
#Requires AutoHotkey v2.0
#SingleInstance Force
#MaxThreadsPerHotkey 16
#MaxThreads 16
#NoTrayIcon

; === Constants ===
DATA_DIRECTORY      := A_AppData . "\LukaRicApps\MacroRic"
CONFIG_PATH         := DATA_DIRECTORY . "\config.ini"
SAVED_PATH          := DATA_DIRECTORY . "\saved" ; files use .mrs extension (MacroRicSave)
APPLICATION_VERSION := 2.0
APPLICATION_TITLE   := "MacroRic v" . APPLICATION_VERSION

; === State Variables ===
g_is_macro_active := False
g_macro_execution_rows := []
g_main_hotkey := ""
g_gui := ""

; =============================
;   DATASTORE CLASSES
; =============================

class DS_INI {
    ; Default config values
    static DEFAULT_VALUES := { Hotkey: "F1" }

    ; Make sure the config file exists, if not create it and set default values.
    static Init() { ; -> Void
        try {
            ; Verify the folder path existing
            if !DirExist(DATA_DIRECTORY)
                DirCreate(DATA_DIRECTORY)
            
            ; Verify the config file exists
            if !FileExist(CONFIG_PATH) {
                ; Iterate the default keys and values
                for key, value in DS_INI.DEFAULT_VALUES.OwnProps() {
                    ; Write the default values to the file
                    IniWrite(value, CONFIG_PATH, "Default", key)
                }
                return
            }

            ; Read the ini file, and create two arrays containing keys and values
            ini_keys := []
            ini_values := []
            for line in StrSplit(IniRead(CONFIG_PATH, "Default"), "`n") {
                ; Clean the line
                line := Trim(line)

                ; Skip empty lines
                if !line
                    continue

                ; Append the keys and values to the arrays
                kv := StrSplit(line, "=")
                ini_keys.Push(kv[1])
                ini_values.Push(kv[2])
            }

            ; Iterate over the default keys and values
            for key, value in DS_INI.DEFAULT_VALUES.OwnProps() {
                ; Verify the key and value
                if !ArrayContains(ini_keys, key) || ini_values[A_Index] == ""
                    IniWrite(value, CONFIG_PATH, "Default", key)
            }
        } catch {
            ; Notify the user that an error occurred
            ErrorOccurred("An error occurred while verifying the config file.`n"
                . "All config settings have been reset.`n`n"
                . "Your saved macros are NOT affected!")
            
            ; Delete the file and try again
            file := FileOpen(CONFIG_PATH, "w")
            if !IsObject(file)
                ErrorOccurred("An error occurred while resetting the config file.", true)

            ; Iterate the default keys and values
            for key, value in DS_INI.DEFAULT_VALUES.OwnProps() {
                ; Write the default values to the file
                IniWrite(value, CONFIG_PATH, "Default", key)
            }
        }
    }

    ; Write to a section=Default and key with the given value.
    static Write(key, value) { ; -> Void
        IniWrite(value, CONFIG_PATH, "Default", key)
    }

    ; Read from a section=Default and key.
    static Read(key) { ; -> Value: string
        return IniRead(CONFIG_PATH, "Default", key)
    }
}

class DS_TBL {
    ; Save the current macro to a file.
    static SaveMacro(file_path, data) { ; -> Void
        ; Make sure the saved folder exists
        if !DirExist(SAVED_PATH)
            DirCreate(SAVED_PATH)

        ; Check if the file exists, if so, delete it
        if FileExist(file_path)
            FileDelete(file_path)

        ; Compact the data rows
        compacted := ""
        for row in data {
            ; Compact the row entries
            for entry in row
                compacted .= entry . ";"

            ; Remove the last semicolon and append a newline
            compacted := SubStr(compacted, 1, StrLen(compacted) - 1) . "`n"
        }

        ; Create the file containing the save data
        FileAppend(compacted, file_path)
    }

    ; Load a macro from a save file.
    static ReadMacro(file_path) { ; -> Data: array[array]
        ; Read the file data
        data := FileRead(file_path)

        header := []
        rows := []

        loop parse data, "`n" {
            ; Trim the line and skip if its blank
            line := Trim(A_LoopField)
            if !line
                continue
            
            if A_Index == 1 {
                ; Parse the settings
                header := StrSplit(line, ";")
            } else {
                ; Parse the row
                rows.Push(StrSplit(line, ";"))
            }
        }

        ; Return the unpacked data
        return [header, rows]
    }
}

; =============================
;   GUI CLASS
; =============================

class MainGUI extends Gui {
    ; Constructor
    __New() {
        ; Call the Gui constructor.
        super.__New("+AlwaysOnTop -Resize", APPLICATION_TITLE)

        ; Set the closing event.
        this.OnEvent("Close", ObjBindMethod(this, "Close"))

        ; Create a location to store gui controls
        this.Widgets := {}

        ; Populate the GUI
        try {
            this.Populate_Controls()
            this.Populate_PostWait()
            this.Populate_Hold()
            this.Populate_MousePosition()
            this.Populate_RowControls()
            this.Populate_LV()
        } catch {
            ErrorOccurred("An error occurred while populating the GUI", true)
        }
        
        ; Set keyboard focus to context for ease of life
        this.Widgets.Context.Focus()

        this.Show("w740 h310")
    }

    ; === Context Menu Functions ===
    ContextMenu_Replace(*) { ; -> Void
        global g_macro_execution_rows

        ; Grab the selected row indexes
        selected := LV_GetSelected(this.LV)

        ; Calculate the wait time in milliseconds
        wait_ms := (this.Widgets.PostWaitHours.Value * 60 * 60 + this.Widgets.PostWaitMinutes.Value * 60
            + this.Widgets.PostWaitSeconds.Value) * 1000 + this.Widgets.PostWaitMilliseconds.Value

        ; Calculate the hold duration in milliseconds
        hold_ms := (this.Widgets.HoldHours.Value * 60 * 60 + this.Widgets.HoldMinutes.Value * 60
            + this.Widgets.HoldSeconds.Value) * 1000 + this.Widgets.HoldMilliseconds.Value
        
        ; Construct the mouse information if needed
        mouse_info := ""
        if this.Widgets.MoveMouse.Value
            mouse_info := "x" . this.Widgets.MouseX.Value . " y" . this.Widgets.MouseY.Value
                . (this.Widgets.ScreenRelative.Value ? " SR" : " WR") . " s" . this.Widgets.MouseSpeed.Value

        ; Iterate the selected indexes
        for index in selected {
            ; Modify the execution row
            g_macro_execution_rows[index] := { Context: this.Widgets.Context.Value, HoldMS: hold_ms, WaitMS: wait_ms,
                Repeats: this.Widgets.RowRepeats.Value, MoveMouse: this.Widgets.MoveMouse.Value,
                MouseX: this.Widgets.MouseX.Value, MouseY: this.Widgets.MouseY.Value, MouseSpeed: this.Widgets.MouseSpeed.Value,
                Relativity: this.Widgets.ScreenRelative.Value ? "Screen" : "Window" }

            ; Modify the list view row
            this.LV.Modify(index,, this.Widgets.Context.Value, FormatMS(wait_ms), FormatMS(hold_ms),
                this.Widgets.RowRepeats.Value, mouse_info)
        }

        this.UpdateEstimatedExecutionTime()
    }

    ContextMenu_MoveTop(*) { ; -> Void
        MsgBox("This feature isn't implemented!", APPLICATION_TITLE . " - Missing Feature", 64)
        ;this.MoveRows("top")
    }

    ContextMenu_MoveUp(*) { ; -> Void
        MsgBox("This feature isn't implemented!", APPLICATION_TITLE . " - Missing Feature", 64)
        ;this.MoveRows("up")
    }

    ContextMenu_MoveDown(*) { ; -> Void
        MsgBox("This feature isn't implemented!", APPLICATION_TITLE . " - Missing Feature", 64)
        ;this.MoveRows("down")
    }

    ContextMenu_MoveBottom(*) { ; -> Void
        MsgBox("This feature isn't implemented!", APPLICATION_TITLE . " - Missing Feature", 64)
        ;this.MoveRows("bottom")
    }

    ContextMenu_DeleteSelected(*) { ; -> Void
        ; Grab the selected row indexes
        selected := LV_GetSelected(this.LV)

        ; Iterate the selected rows
        for index in selected {
            ; Delete the rows from the global array, and from the list view
            g_macro_execution_rows.RemoveAt(index)
            this.LV.Delete(index)
        }

        ; Check if widgets should be disabled
        if this.LV.GetCount() < 1
            this.Widgets.SaveMacro.Opt("+Disabled")

        this.UpdateEstimatedExecutionTime()
    }

    ContextMenu_DeleteAll(*) { ; -> Void
        global g_macro_execution_rows

        ; Delete all list view rows
        this.LV.Delete()

        ; Clear all execution rows
        g_macro_execution_rows := []

        ; Set the save button state
        this.Widgets.SaveMacro.Opt("+Disabled")

        this.UpdateEstimatedExecutionTime()
    }

    ; === Callback Functions ===
    Callback_Hotkey(*) { ; -> Void
        global g_main_hotkey

        ; Get the new hotkey
        new_hotkey := this.Widgets.Hotkey.Value

        ; Update the config file
        DS_INI.Write("Hotkey", new_hotkey)

        ; Update the hotkey
        Hotkey("$" . g_main_hotkey, ToggleMacro, "Off")
        Hotkey("$" . new_hotkey, ToggleMacro, "On")

        ; Update the stop macro button text
        this.Widgets.StopMacro.Text := "Stop Macro (" . new_hotkey . ")"

        ; Change the global hotkey variable
        g_main_hotkey := new_hotkey
    }

    Callback_Context(*) { ; -> Void
        ; Get the length of the context
        enable := StrLen(this.Widgets.Context.Value) > 0

        ; Disable or enable the append row button
        this.Widgets.AppendRow.Opt((enable ? "-" : "+") . "Disabled")
    }

    Callback_PickPosition(*) { ; -> Void
        ; Hide the GUI
        this.Minimize()

        ; Mouse tracking function
        TrackMouse(*) {
            ; Grab the coordinates relative to the screen
            CoordMode("Mouse", "Screen")
            mouse_screen_x := 0
            mouse_screen_y := 0
            MouseGetPos(&mouse_screen_x, &mouse_screen_y)
            
            ; Grab the coordinates relative to the window
            CoordMode("Mouse", "Window")
            mouse_window_x := 0
            mouse_window_y := 0
            MouseGetPos(&mouse_window_x, &mouse_window_y)

            ; Display the current coords in a tooltip
            ToolTip("Screen: x" . mouse_screen_x . " y" . mouse_screen_y
                . "`nWindow: x" . mouse_window_x . " y" . mouse_window_y)
        }

        ; Called when the user picks a position
        OnPickPositionFinished(*) {
            ; Clear the hotkey
            Hotkey("$LButton", ObjBindMethod(this, "OnPickPositionFinished"), "Off")

            ; Clear the timer
            SetTimer(TrackMouse, 0)
            ToolTip()

            CoordMode("Mouse", this.Widgets.ScreenRelative.Value ? "Screen" : "Window")
            mouse_x := 0
            mouse_y := 0
            MouseGetPos(&mouse_x, &mouse_y)

            this.Widgets.MouseX.Value := mouse_x
            this.Widgets.MouseY.Value := mouse_y

            ; Display the GUI again
            this.Show()
        }
        
        ; Start a 60fps timer to track the mouse
        SetTimer(TrackMouse, 1000 / 120)

        ; Set the left mouse button to call a function
        Hotkey("$LButton", OnPickPositionFinished, "On")
    }

    Callback_MouseSpeed(*) { ; -> Void
        ; Get the value
        value := this.Widgets.MouseSpeed.Value

        ; So it doesn't raise an error when erasing the number to type a new value
        if value == ""
            return

        ; Clamp it to 100
        if value > 100
            this.Widgets.MouseSpeed.Value := 100
    }

    Callback_AppendRow(*) { ; -> Void
        ; Calculate the wait time in milliseconds
        wait_ms := (this.Widgets.PostWaitHours.Value * 60 * 60 + this.Widgets.PostWaitMinutes.Value * 60
            + this.Widgets.PostWaitSeconds.Value) * 1000 + this.Widgets.PostWaitMilliseconds.Value

        ; Calculate the hold duration in milliseconds
        hold_ms := (this.Widgets.HoldHours.Value * 60 * 60 + this.Widgets.HoldMinutes.Value * 60
            + this.Widgets.HoldSeconds.Value) * 1000 + this.Widgets.HoldMilliseconds.Value
        
        ; Construct the mouse information if needed
        mouse_info := ""
        if this.Widgets.MoveMouse.Value
            mouse_info := "x" . this.Widgets.MouseX.Value . " y" . this.Widgets.MouseY.Value
                . (this.Widgets.ScreenRelative.Value ? " SR" : " WR") . " s" . this.Widgets.MouseSpeed.Value
        
        ; Push the data to the execution rows
        g_macro_execution_rows.Push({ Context: this.Widgets.Context.Value, HoldMS: hold_ms, WaitMS: wait_ms,
            Repeats: this.Widgets.RowRepeats.Value, MoveMouse: this.Widgets.MoveMouse.Value,
            MouseX: this.Widgets.MouseX.Value, MouseY: this.Widgets.MouseY.Value, MouseSpeed: this.Widgets.MouseSpeed.Value,
            Relativity: this.Widgets.ScreenRelative.Value ? "Screen" : "Window" })
        
        ; Append the data to the list view
        this.LV.Add(, this.Widgets.Context.Value, FormatMS(wait_ms), FormatMS(hold_ms), this.Widgets.RowRepeats.Value, mouse_info)

        ; Set the save button state
        this.Widgets.SaveMacro.Opt("-Disabled")

        this.UpdateEstimatedExecutionTime()
    }

    Callback_LoadMacro(*) { ; -> Void
        global g_macro_execution_rows

        ; Minimize the GUI
        this.Minimize()

        ; Prompt the user to pick a file
        file_path := FileSelect(, SAVED_PATH)
        if !file_path {
            this.Show()
            return
        }

        ; Grab the file data and parse it
        data := DS_TBL.ReadMacro(file_path)

        ; Set the settings with the file data
        this.Widgets.Repeats.Value := data[1][1]
        this.Widgets.RowRepeats.Value := data[1][2]
        this.Widgets.Context.Value := data[1][3]
        this.Widgets.PostWaitHours.Value := data[1][4]
        this.Widgets.PostWaitMinutes.Value := data[1][5]
        this.Widgets.PostWaitSeconds.Value := data[1][6]
        this.Widgets.PostWaitMilliseconds.Value := data[1][7]
        this.Widgets.HoldHours.Value := data[1][8]
        this.Widgets.HoldMinutes.Value := data[1][9]
        this.Widgets.HoldSeconds.Value := data[1][10]
        this.Widgets.HoldMilliseconds.Value := data[1][11]
        this.Widgets.MoveMouse.Value := data[1][12]
        this.Widgets.MouseX.Value := data[1][13]
        this.Widgets.MouseY.Value := data[1][14]
        this.Widgets.ScreenRelative.Value := data[1][15]
        this.Widgets.MouseSpeed.Value := data[1][16]

        ; Delete all list view rows
        this.LV.Delete()

        ; Clear all execution rows
        g_macro_execution_rows := []

        ; Load the rows
        for row in data[2] {
            ; Append the row to the macro execution rows
            g_macro_execution_rows.Push({ Context: row[1], HoldMS: row[2], WaitMS: row[3], Repeats: row[4],
                MoveMouse: row[5], MouseX: row[6], MouseY: row[7], MouseSpeed: row[8], Relativity: row[9] })

            ; Construct the mouse information if needed
            mouse_info := ""
            if row[5]
                mouse_info := "x" . row[6] . " y" . row[7] . (row[9] == "Screen" ? "SR" : "WR") . " s" . row[8]

            ; Append the row to the list view
            this.LV.Add(, row[1], FormatMS(row[2]), FormatMS(row[3]), row[4], mouse_info)
        }

        ; Set the save button state
        this.Widgets.SaveMacro.Opt("-Disabled")

        ; Set the append row state
        if this.Widgets.Context.Value
            this.Widgets.AppendRow.Opt("-Disabled")

        this.UpdateEstimatedExecutionTime()

        ; Reshow the GUI
        this.Show()
    }

    Callback_SaveMacro(*) { ; -> Void
        ; Minimize the GUI
        this.Minimize()

        ; Recursive function for repeated prompting
        _PromptSaveFile(this) {
            ; Prompt the user for a file name
            file_name := InputBox("Write the name of the save file (no extensions)`n`nBlank names aren't saved!",
                APPLICATION_TITLE . " - Save Macro", "w300 h130")

            ; Confirm if the user typed a name
            if !file_name.Value {
                this.Show()
                return
            }

            ; Check if a file with that name already exists
            if FileExist(SAVED_PATH . "\" . file_name.Value . ".mrs") {
                ; Prompt the user weather to overwrite or write a new file name
                response := MsgBox("A saved config with this name already exists!`n`nOverwrite the file?",
                    APPLICATION_TITLE . " - Save Macro", 48+4)

                ; If the user wants to overwrite the existing save, return the name                
                if response == "Yes"
                    return file_name.Value

                ; Reprompt the user for a new file name
                _PromptSaveFile(this)
            }

            return file_name.Value
        }

        ; Prompt the user for a file name and validate it
        file_name := _PromptSaveFile(this)
        if !file_name {
            this.Show()
            return
        }

        ; Create a variable to hold the data
        data := []

        ; Copy the current setting configuration into the data variable
        data.Push([this.Widgets.Repeats.Value, this.Widgets.RowRepeats.Value, this.Widgets.Context.Value,
            this.Widgets.PostWaitHours.Value, this.Widgets.PostWaitMinutes.Value, this.Widgets.PostWaitSeconds.Value,
            this.Widgets.PostWaitMilliseconds.Value, this.Widgets.HoldHours.Value, this.Widgets.HoldMinutes.Value,
            this.Widgets.HoldSeconds.Value, this.Widgets.HoldMilliseconds.Value, this.Widgets.MoveMouse.Value,
            this.Widgets.MouseX.Value, this.Widgets.MouseY.Value, this.Widgets.ScreenRelative.Value, this.Widgets.MouseSpeed.Value])

        ; Insert all the execution rows into the data variable
        for row in g_macro_execution_rows {
            data.Push([row.Context, row.HoldMS, row.WaitMS, row.Repeats, row.MoveMouse, row.MouseX,
                row.MouseY, row.MouseSpeed, row.Relativity])
        }

        ; Save the data to file
        DS_TBL.SaveMacro(SAVED_PATH . "\" . file_name . ".mrs", data)

        ; Reshow the GUI
        this.Show()
    }

    ; === Population Functions ===
    Populate_Controls() { ; -> Void
        ; Groupbox
        this.AddGroupBox("x10 y5 w165 h297", "Controls")

        ; Hotkey
        this.AddText("x20 y25 w40 h20 0x200", "Hotkey:")
        this.Widgets.Hotkey := this.AddHotkey("x+5 yp w100 h20", g_main_hotkey)
        this.Widgets.Hotkey.OnEvent("Change", ObjBindMethod(this, "Callback_Hotkey"))

        ; Seperator
        this.AddText("x15 y+10 w155 h1 0x10")

        ; Repeat
        this.AddText("x20 y+10 w40 h20 0x200", "Repeat:")
        this.Widgets.Repeats := this.AddEdit("x+5 yp w65 h20 +Number Limit9", 1)
        this.Widgets.Repeats.OnEvent("Change", (*) => this.UpdateEstimatedExecutionTime())
        this.AddText("x+5 yp w30 h20 0x200", "Times")

        ; Seperator
        this.AddText("x15 y+10 w155 h1 0x10")

        ; Load macro
        this.Widgets.LoadMacro := this.AddButton("x20 y+10 w145 h25", "Load Macro")
        this.Widgets.LoadMacro.OnEvent("Click", ObjBindMethod(this, "Callback_LoadMacro"))

        ; Save macro
        this.Widgets.SaveMacro := this.AddButton("x20 y+5 wp hp +Disabled", "Save Macro")
        this.Widgets.SaveMacro.OnEvent("Click", ObjBindMethod(this, "Callback_SaveMacro"))

        ; Seperator
        this.AddText("x15 y+10 w155 h1 0x10")

        ; Stop macro
        this.Widgets.StopMacro := this.AddButton("x20 y+10 w145 h25 +Disabled", "Stop Macro (" . g_main_hotkey . ")")
        this.Widgets.StopMacro.OnEvent("Click", ToggleMacro)

        ; Seperator
        this.AddText("x15 y+10 w155 h1 0x10")

        ; Append row
        this.Widgets.AppendRow := this.AddButton("x20 y+10 w145 h25 +Disabled", "Append Row")
        this.Widgets.AppendRow.OnEvent("Click", ObjBindMethod(this, "Callback_AppendRow"))

        ; Seperator
        this.AddText("x15 y+10 w155 h1 0x10")

        ; Credit text
        this.AddText("x20 y+10 w145 h20 +Center", "Made by LukaRic")
    }

    Populate_PostWait() { ; -> Void
        ; Groupbox
        this.AddGroupBox("x185 y5 w284 h55", "Post-wait Duration")

        ; Hours
        this.AddText("x195 y25 w34 h20 0x200", "Hours:")
        this.Widgets.PostWaitHours := this.AddEdit("x+0 yp w30 hp +Number Limit2", 0)

        ; Minutes
        this.AddText("x+5 yp w28 hp 0x200", "Mins:")
        this.Widgets.PostWaitMinutes := this.AddEdit("x+0 yp w30 hp +Number Limit2", 0)

        ; Seconds
        this.AddText("x+5 yp w29 hp 0x200", "Secs:")
        this.Widgets.PostWaitSeconds := this.AddEdit("x+0 yp w30 hp +Number Limit2", 0)

        ; Milliseconds
        this.AddText("x+5 yp w38 hp 0x200", "MSecs:")
        this.Widgets.PostWaitMilliseconds := this.AddEdit("x+0 yp w30 hp +Number Limit3", 100)
    }

    Populate_Hold() { ; -> Void
        ; Groupbox
        this.AddGroupBox("x185 y65 w284 h55", "Hold Duration")

        ; Hours
        this.AddText("x195 y85 w34 h20 0x200", "Hours:")
        this.Widgets.HoldHours := this.AddEdit("x+0 yp w30 hp +Number Limit2", 0)

        ; Minutes
        this.AddText("x+5 yp w28 hp 0x200", "Mins:")
        this.Widgets.HoldMinutes := this.AddEdit("x+0 yp w30 hp +Number Limit2", 0)

        ; Seconds
        this.AddText("x+5 yp w29 hp 0x200", "Secs:")
        this.Widgets.HoldSeconds := this.AddEdit("x+0 yp w30 hp +Number Limit2", 0)

        ; Milliseconds
        this.AddText("x+5 yp w38 hp 0x200", "MSecs:")
        this.Widgets.HoldMilliseconds := this.AddEdit("x+0 yp w30 hp +Number Limit3", 0)
    }

    Populate_MousePosition() { ; -> Void
        ; Groupbox
        this.AddGroupBox("x479 y5 w250 h115", "Mouse Position")

        ; X coordinate
        this.AddText("x489 y25 w10 h20 0x200", "X:")
        this.Widgets.MouseX := this.AddEdit("+0 yp w40 hp +Number Limit4", 0)

        ; Y coordinate
        this.AddText("x+10 yp w10 hp 0x200", "Y:")
        this.Widgets.MouseY := this.AddEdit("+0 yp w40 hp +Number Limit4", 0)

        ; Pick position
        this.Widgets.PickPosition := this.AddButton("x+10 yp-2.5 w100 h25", "Pick Position")
        this.Widgets.PickPosition.OnEvent("Click", ObjBindMethod(this, "Callback_PickPosition"))

        ; Mouse relativity
        this.Widgets.ScreenRelative := this.AddRadio("x489 y+10 w112.5 h20 Checked1", "Screen Relative")
        this.AddRadio("x+5 yp wp hp", "Window Relative")

        ; Move speed
        this.AddText("x489 y+10 w70 h20 0x200", "Move Speed:")
        this.Widgets.MouseSpeed := this.AddEdit("x+0 yp w30 h20 +Number Limit3", 100)
        this.Widgets.MouseSpeed.OnEvent("Change", ObjBindMethod(this, "Callback_MouseSpeed"))

        ; Togglecheck
        this.Widgets.MoveMouse := this.AddCheckbox("x+17 yp w112 h20", "Move Mouse")
    }

    Populate_RowControls() { ; -> Void
        ; Context
        this.AddText("x187 y125 w45 h20 0x200", "Context:")
        this.Widgets.Context := this.AddEdit("x+0 yp w100 hp")
        this.Widgets.Context.OnEvent("Change", ObjBindMethod(this, "Callback_Context"))

        ; Row repeats
        this.AddText("x+10 yp w68 hp 0x200", "Repeat Row:")
        this.Widgets.RowRepeats := this.AddEdit("x+0 yp w65 h20 +Number Limit9", 1)
        this.AddText("x+5 yp w30 h20 0x200", "Times")

        ; Seperator
        this.AddText("x+10 yp w1 h23 0x11")

        ; Status
        this.Widgets.Statusbar := this.AddText("x+10 yp w150 h20 0x200 c505050", "EST: 0ms")
        this.Widgets.StatusTimer := this.AddText("x+5 yp w43 hp 0x200 Right c505050")
    }

    Populate_LV() { ; -> Void
        ; Listview
        this.LV := this.AddListView("x187 y+5 w542 h151 +NoSortHdr +NoSort +Grid -LV0x10 -WantF2", 
            ["Context", "Wait Duration (After)", "Hold Duration", "Repeats", "Move Mouse (Before)"])
        
        ; Set the header sizes
        for width in [119, 113, 113, 72, 121]
            this.LV.ModifyCol(A_Index, width)

        ; Build the context menu
        context_menu := Menu()
        context_menu.Add("Replace Values", ObjBindMethod(this, "ContextMenu_Replace"))
        context_menu.Add()
        context_menu.Add("Move To Top", ObjBindMethod(this, "ContextMenu_MoveTop"))
        context_menu.Add("Move Up", ObjBindMethod(this, "ContextMenu_MoveUp"))
        context_menu.Add("Move Down", ObjBindMethod(this, "ContextMenu_MoveDown"))
        context_menu.Add("Move To Bottom", ObjBindMethod(this, "ContextMenu_MoveBottom"))
        context_menu.Add()
        context_menu.Add("Delete Selected", ObjBindMethod(this, "ContextMenu_DeleteSelected"))
        context_menu.Add("Delete All", ObjBindMethod(this, "ContextMenu_DeleteAll"))
        context_menu.Default := "Delete All"

        ; When rightclicking the list view, show the menu
        this.LV.OnEvent("ContextMenu", (*) => (this.LV.GetCount("S") > 0 ? context_menu.Show() : ""))
    }

    ; === Other Functions ===
    ; Move list view rows up, down, top or bottom
    MoveRows(direction) {
        ; Make sure theres rows in the list view
        total := this.LV.GetCount()
        if !total
            return

        ; Make sure theres rows selected
        selected := LV_GetSelected(this.LV)
        if !selected.Length
            return

        ; 
        data := LV_BackupData(this.LV, selected)
        LV_DeleteRows(this.LV, selected)

        ; Determine new insertion position
        if (direction = "up")
            insertAt := Max(selected[1] - 1, 1)
        else if (direction = "down")
            insertAt := Min(selected[selected.Length] + 1, this.LV.GetCount() + 1)
        else if (direction = "top")
            insertAt := 1
        else if (direction = "bottom")
            insertAt := this.LV.GetCount() + 1
        else
            return

        ; Insert rows
        for rowData in data
            this.LV.Insert(insertAt++, "", rowData*)

        LV_RestoreSelection(this.LV, direction, selected, data.Length)
    }

    ; Works like Sleep(), except it updates the GUI status timer
    Wait(ms, update_var:=0, interval:=100) { ; -> Void
        ; Define variables
        start := A_TickCount
        remaining := ms

        ; Loop until theres no remaining time.
        while (remaining > 0) {
            ; Update the status timer
            this.Widgets.StatusTimer.Value := Round(Max(0, (ms - (A_TickCount - start)) / 1000), 1) . "s"

            ; Use actual sleep
            Sleep(interval)

            ; Set the new remaining time
            remaining := ms - (A_TickCount - start)
        }

        this.Widgets.StatusTimer.Value := ""
    }

    ; Calculate the estimated execution time
    UpdateEstimatedExecutionTime() { ; -> Void
        flag_mouse_move := false
        est_duration_ms := 0

        ; Check if the macro is looping infinitly
        if this.Widgets.Repeats.Value == 0 {
            this.Widgets.Statusbar.Value := "EST: Infinite"
            return
        }
        
        ; Iterate the execution rows
        for row in g_macro_execution_rows {
            ; Count the sleep time
            est_duration_ms += (row.WaitMS + row.HoldMS) * row.Repeats

            ; Raise the flag if the mouse should move
            if row.MoveMouse && row.MouseSpeed != 100
                flag_mouse_move := true
        }

        ; Multiply by macro repeats
        est_duration_ms *= (this.Widgets.Repeats.Value || 0)

        ; Check if the duration is above a month
        if est_duration_ms > 30 * 24 * 60 * 60 * 1000 {
            this.Widgets.Statusbar.Value := "EST: Infinite"
            return
        }

        ; Format to string
        est_duration_f := FormatMS(est_duration_ms)

        ; Set the statusbar text
        this.Widgets.Statusbar.Value := "EST: " . est_duration_f . (flag_mouse_move ? " (~)" : "")
    }

    ; Close the application.
    Close(*) { ; -> Void
        this.Destroy()
        ExitApp()
    }
}

; =============================
;   HELPER FUNCTIONS
; =============================

; Create a backup of the selected data and return it in an array
LV_BackupData(LV, selected) { ; -> Data: array[array[str]]
    data := []

    ; Iterate selected
    for idx in selected {
        ; Grab all the column text and push it to an array
        rowData := []
        Loop LV.GetCount("Col")
            rowData.Push(LV.GetText(idx, A_Index))

        ; Append the row data
        data.Push(rowData)
    }

    return data
}

; Delete selected rows
LV_DeleteRows(LV, selected) {
    ; Delete in reverse order
    for i, _ in selected {
        idx := selected[selected.Length - i + 1]
        LV.Delete(idx)
    }
}

; Moves the rows in the list view
LV_RestoreSelection(LV, direction, selected, count) {
    ; Clear all existing selections
    LV.Modify(0, "-Select")

    if (direction = "up")
        ; Shift up one selection, but dont go above the first row
        startIndex := Max(selected[1] - 1, 1)
    else if (direction = "down")
        ; Shift down, but down go beyond the final row
        startIndex := Min(selected[selected.Length] + 1, LV.GetCount())
    else if (direction = "top")
        ; Always start at the first row
        startIndex := 1
    else if (direction = "bottom")
        ; Always start at the, first of the last N rows
        startIndex := LV.GetCount() - count + 1

    ; Reselect the new moved rows at their new positions
    Loop count
        LV.Modify(startIndex + A_Index - 1, "+Select")

    ; Return keyboard focus
    LV.Focus()
}

; Get a list of the selected indexes from a list view
LV_GetSelected(LV) { ; -> Selection: array[int]
    selected := []
    RowNumber := 0

    ; Iterate the list view rows, grabbing the selected rows
    while (RowNumber := LV.GetNext(RowNumber))
        ; Insert from the front as to make it easier for implementations
        selected.InsertAt(1, RowNumber)
    
    return selected
}

; Format milliseconds into a dynamic time string
FormatMS(ms) { ; -> Format: string
    ; Extract h, m, s & ms
    hours := Floor(ms / (60 * 60 * 1000))
    mins := Floor(Mod(ms, (60 * 60 * 1000)) / (60 * 1000))
    secs := Floor(Mod(ms, (60 * 1000)) / 1000)
    ms := Mod(ms, 1000)

    ; Minify the output
    out := ""
    if (hours)
        out .= hours . "h "
    if (hours || mins)
        out .= mins . "m "
    if (hours || mins || secs)
        out .= secs . "s "
    if (ms)
        out .= ms . "ms"

    return out || "0ms"
}

; Notify user and log errors
ErrorOccurred(message, deadly:=false) { ; -> Void
    ; If the error is deadly, notify the user of this.
    if deadly
        message .= "`n`nTHIS ERROR IS CRITICAL, THE APPLICATION WILL TERMINATE."

    ; Notify the user.
    MsgBox(message, APPLICATION_TITLE . " - Error", deadly ? 16 : 48)

    if deadly
        ExitApp()
}

; Checks if a value exists in an array
ArrayContains(array, value) { ; -> Result: boolean
    ; Iterate the array
    for item in array {
        ; Check if the item matches the value
        if item == value
            return true
    }
    return false
}

; =============================
;   MAIN FUNCTIONS
; =============================

; Responsible for executing the macro (moving the mouse & pressing keys)
ExecuteMacro() {
    ; Iterate the execution rows
    loop g_gui.Widgets.Repeats.Value {
        ; Escape if needed
        if !g_is_macro_active {
            g_gui.Widgets.Statusbar.Value := "Execution Stopped!"
            return
        }

        for row in g_macro_execution_rows {
            ; Escape if needed
            if !g_is_macro_active {
                g_gui.Widgets.Statusbar.Value := "Execution Stopped!"
                return
            }
            
            loop row.Repeats {
                ; Escape if needed
                if !g_is_macro_active {
                    g_gui.Widgets.Statusbar.Value := "Execution Stopped!"
                    return
                }

                ; Check if the mouse should move
                if row.MoveMouse {
                    ; Set the statusbar
                    g_gui.Widgets.Statusbar.Value := "Move: x" . row.MouseX . " y" . row.MouseY
                        . (row.Relativity == "Screen" ? "SR" : "WR") . " s" . row.MouseSpeed

                    ; Set the coordmode
                    CoordMode("Mouse", row.Relativity)

                    ; Move the mouse
                    MouseMove(row.MouseX, row.MouseY, 100 - row.MouseSpeed)
                }

                ; Check if the context should be held
                if row.HoldMS > 0 {
                    ; Set the statusbar
                    g_gui.Widgets.Statusbar.Value := "Hold: " . row.Context

                    ; Strip the context
                    context_stripped := StrReplace(StrReplace(row.Context, "}", ""), "{", "")

                    ; Hold the context
                    Send("{" . context_stripped . " Down}")

                    g_gui.Wait(row.HoldMS)

                    ; Escape if needed
                    if !g_is_macro_active {
                        g_gui.Widgets.Statusbar.Value := "Execution Stopped!"
                        return
                    }

                    ; Lift the context
                    Send("{" . context_stripped . " Up}")
                } else {
                    ; Set the statusbar
                    g_gui.Widgets.Statusbar.Value := "Send: " . row.Context

                    ; Send the context
                    Send(row.Context)
                }

                ; Check if the macro should sleep post to executing
                if row.WaitMS > 0 {
                    g_gui.Wait(row.WaitMS)
                }
            }
        }
    }

    ; Set the statusbar
    g_gui.Widgets.Statusbar.Value := "Execution Complete!"

    ; Macro finished, so toggle it off
    ToggleMacro()
}

; Toggle the macro, if true passed, rows will start executing.
ToggleMacro(*) { ; -> Void
    global g_is_macro_active

    ; Make sure theres rows that can be executed
    if g_macro_execution_rows.Length < 1
        return

    ; Make sure the user cant start the macro while the window is active
    if !g_is_macro_active && WinActive(APPLICATION_TITLE)
        return

    ; Flip the state
    g_is_macro_active := !g_is_macro_active

    ; Change the disabled state of some widgets
    g_gui.Widgets.StopMacro.Opt((g_is_macro_active ? "-" : "+") . "Disabled")
    g_gui.LV.Opt((g_is_macro_active ? "+" : "-") . "Disabled")
    g_gui.Widgets.PickPosition.Opt((g_is_macro_active ? "+" : "-") . "Disabled")
    g_gui.Widgets.LoadMacro.Opt((g_is_macro_active ? "+" : "-") . "Disabled")
    g_gui.Widgets.SaveMacro.Opt((g_is_macro_active ? "+" : "-") . "Disabled")
    g_gui.Widgets.AppendRow.Opt((g_is_macro_active ? "+" : "-") . "Disabled")

    if g_is_macro_active
        ExecuteMacro()
}

; Main entry point
Main() { ; -> Void
    global g_main_hotkey, g_gui

    DS_INI.Init()

    g_main_hotkey := DS_INI.Read("Hotkey")
    Hotkey("$" . g_main_hotkey, ToggleMacro, "On")

    g_gui := MainGUI()
}

; Launch off!
Main()