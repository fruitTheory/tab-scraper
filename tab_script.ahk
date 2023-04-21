#NoEnv
#SingleInstance, Force

; Loop through tabs
Loop
{

    ; Ctrl+Tab to switch to next tab
    Send ^{Tab}

    ; Wait for the new tab to load
    Sleep, 1000

}

