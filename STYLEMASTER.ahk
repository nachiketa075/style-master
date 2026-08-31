 #Requires AutoHotkey v1.1.33+

#NoEnv

#SingleInstance Force

SendMode Input

SetWorkingDir %A_ScriptDir%

; ============================================================
; REUSABLE SINGLE-PICK LIST POPUP
; Shared by Shape/Quality/Setting/Brand. Options are "CODE |
; Description" lines (`n-separated); only CODE is returned.
; SearchPrefix pre-filters case-insensitively: exactly 1 match
; auto-returns with no popup, 0 or 2+ show the (filtered) list.
; Committed via Select button, Enter, or double-click.
; ============================================================

ShowSelectPopup(PopupTitle, OptionsList, SearchPrefix := "")
{
    global PopupChoice, PopupResult

    PopupResult := ""
    PopupChoice := ""

    ListToShow := OptionsList

    if (SearchPrefix != "")
    {
        FilteredList := ""
        MatchCount := 0
        LastMatch := ""

        Loop, Parse, OptionsList, `n
        {
            OptionLine := A_LoopField

            if (OptionLine = "")
                continue

            StringLower, OptionLineLower, OptionLine

            if InStr(OptionLineLower, SearchPrefix)
            {
                FilteredList .= OptionLine . "`n"
                MatchCount++
                LastMatch := OptionLine
            }
        }

        if (MatchCount = 1)
        {
            StringSplit, AutoPart, LastMatch, |
            PopupResult := Trim(AutoPart1)
            return PopupResult
        }
        else if (MatchCount > 1)
            ListToShow := Trim(FilteredList, "`n")
    }

    Gui, Popup:Destroy
    Gui, Popup:New, +AlwaysOnTop +ToolWindow +HwndPopupHwnd, %PopupTitle%
    Gui, Popup:Font, s10, Segoe UI
    Gui, Popup:+Delimiter`n

    Gui, Popup:Add, Text, x20 y15 w650 h25, Select one %PopupTitle%:
    Gui, Popup:Add, ListBox, x20 y45 w650 h450 vPopupChoice Choose1, %ListToShow%
    Gui, Popup:Add, Button, x270 y510 w150 h35 gPopupPick Default, Select

    Gui, Popup:Show, w690 h565, %PopupTitle%

    WinWaitClose, ahk_id %PopupHwnd%

    return PopupResult
}

return


; ==========================================
; COMMIT SELECTION (Select button / Enter)
; ==========================================

PopupPick:

Gui, Popup:Submit

if (PopupChoice = "")
{
    Gui, Popup:Show
    return
}

StringSplit, PopupPart, PopupChoice, |
PopupResult := Trim(PopupPart1)

Gui, Popup:Destroy

return


; ==========================================
; CLOSE WINDOW (X button / Alt+F4)
; ==========================================

PopupGuiClose:

PopupResult := ""

Gui, Popup:Destroy

return

^1::

; ============================================================
; 0. CLOSE EXISTING EXCEL IF IT IS ALREADY OPEN
; ============================================================

IfWinExist, ahk_class XLMAIN
{
    ; Try to connect to currently running Excel
    try
    {
        ExistingXL := ComObjActive("Excel.Application")

        ; Close all existing workbooks without saving
        Loop, % ExistingXL.Workbooks.Count
        {
            try
            {
                ExistingXL.Workbooks(1).Close(false)
            }
            catch
            {
            }
        }

        ; Quit existing Excel
        try
        {
            ExistingXL.Quit()
        }
        catch
        {
        }

        ExistingXL := ""
    }
    catch
    {
        ; If COM connection fails, force close Excel
        Process, Close, EXCEL.EXE
    }

    ; Wait for Excel to completely close
    WinWaitClose, ahk_class XLMAIN,, 10

    ; Extra safety: make sure EXCEL.EXE is gone
    Process, WaitClose, EXCEL.EXE, 10

    Sleep, 1500
}


; ============================================================
; 1. OPEN JPG TO EXCEL CONVERTER
; ============================================================

ConverterPath := "C:\Users\PC 4\Desktop\JPG To Excel Converter.lnk"

Run, %ConverterPath%


; ============================================================
; 2. WAIT UNTIL EXCEL WINDOW AND WORKBOOK ARE READY
; ============================================================

ExcelReady := false

Loop, 60
{
    IfWinExist, ahk_class XLMAIN
    {
        try
        {
            xl := ComObjActive("Excel.Application")

            if (xl.Workbooks.Count > 0)
            {
                xl.Visible := true

                WinActivate, ahk_class XLMAIN
                WinWaitActive, ahk_class XLMAIN,, 5

                if (xl.Workbooks.Count > 0)
                {
                    ExcelReady := true
                    break
                }
            }
        }
        catch
        {
        }
    }

    Sleep, 1000
}


if (!ExcelReady)
{
    MsgBox, 48, Error, Excel workbook is not ready.`nAutomation has been stopped.
    return
}


; ============================================================
; 3. CONNECT TO THE ACTIVE EXCEL WORKBOOK
; ============================================================

try
{
    xl := ComObjActive("Excel.Application")
    ws := xl.ActiveWorkbook.ActiveSheet
}
catch
{
    MsgBox, 48, Error, Unable to access the active Excel workbook.
    return
}


WinActivate, ahk_class XLMAIN
WinWaitActive, ahk_class XLMAIN,, 10

if ErrorLevel
{
    MsgBox, 48, Error, Unable to activate the Excel window.
    return
}

; 4. FIND EXISTING GEMI / JWEL SOLUTION

JwelPath := "\\jemmysoft-new\Jemysoft\False\Test\JwelSolution.exe"

Process, Exist, JwelSolution.exe
JwelPID := ErrorLevel

if (JwelPID)
{
    ; Already running - activate it instead of launching another.
    WinActivate, ahk_exe JwelSolution.exe
    WinWaitActive, ahk_exe JwelSolution.exe,, 10

    if ErrorLevel
    {
        MsgBox, 48, Error, JwelSolution is running but its window could not be activated.
        return
    }
}
else
{
   

    Run, %JwelPath%

    WinWait, ahk_exe JwelSolution.exe,, 30

    if ErrorLevel
    {
        MsgBox, 48, Error, JwelSolution could not be started within 30 seconds.
        return
    }

    WinActivate, ahk_exe JwelSolution.exe
    WinWaitActive, ahk_exe JwelSolution.exe,, 10

    if ErrorLevel
    {
        MsgBox, 48, Error, Unable to activate the JwelSolution window.
        return
    }
}

Sleep, 2000


; 5. GEMI INITIAL ACTIONS

MouseMove, 50, 42
Click, Left, 1
Sleep, 300

MouseMove, 60, 62
Click, Left, 1
Sleep, 300

MouseMove, 206, 66
Click, Left, 1
Sleep, 300

MouseMove, 94, 168
Click, Left, 1
Sleep, 2500


; 6. ENTER PRODUCT CODE

ProductCode := ws.Range("C2").Value
ProductCode := Trim(ProductCode)

StringSplit, ProductCodeParts, ProductCode, (
ProductCode := Trim(ProductCodeParts1)

if (ProductCode = "")
{
    MsgBox, 48, Error, Product Code in Excel C2 is empty.
    return
}
Sleep, 4000

Send, %ProductCode%
Sleep, 2000


; 7. CREATE NEW BLANK ROW

MouseMove, 548, 548
Click, Left, 1

Sleep, 750

Send, ^n

Sleep, 1250

Send, {Left 10}

Send, {Right}

Send, {Enter}

Sleep, 500


; 8. READ CATEGORY FROM EXCEL D2

try
{
    ProductType := ws.Range("D2").Value
    ProductType := Trim(ProductType)

    StringLower, ProductTypeLower, ProductType
    ProductTypeLower := Trim(ProductTypeLower)
}
catch
{
    MsgBox, 48, Error, Unable to read Category from Excel D2.
    return
}

if (ProductTypeLower = "")
{
    MsgBox, 48, Error, Category in Excel D2 is empty.
    return
}

try
{
    Pcs := ws.Range("P3").Value
    Pcs := Trim(Pcs)
}
catch
{
    MsgBox, 48, Error, Unable to read Pcs from Excel P3.
    return
}

try
{
    O3Data := ws.Range("O3").Value
    O3Data := Trim(O3Data)
}
catch
{
    MsgBox, 48, Error, Unable to read data from Excel O3.
    return
}


; 9. CATEGORY - CONVERT TO GEMI CODE

CategoryCode := ""

if (ProductTypeLower = "ring")
    CategoryCode := "RNG"

else if (ProductTypeLower = "earrings" || ProductTypeLower = "earring")
    CategoryCode := "EAR"

else if (ProductTypeLower = "pendant")
    CategoryCode := "PND"

else if (ProductTypeLower = "necklace" || ProductTypeLower = "neclace")
    CategoryCode := "NCK"

else if (ProductTypeLower = "bracelet" || ProductTypeLower = "braclet")
    CategoryCode := "BRC"

else if (ProductTypeLower = "bangle" || ProductTypeLower = "bangel")
    CategoryCode := "BNG"

else if (ProductTypeLower = "head" || ProductTypeLower = "hd")
    CategoryCode := "HD"

else if (ProductTypeLower = "broch" || ProductTypeLower = "brooch" || ProductTypeLower = "brch")
    CategoryCode := "BRCH"

else if (ProductTypeLower = "crown" || ProductTypeLower = "crn")
    CategoryCode := "CRN"

else if (ProductTypeLower = "cuff link" || ProductTypeLower = "cufflink" || ProductTypeLower = "cuff-link" || ProductTypeLower = "cfl")
    CategoryCode := "CFL"

else
    CategoryCode := ProductType



; 10. CATEGORY - SEND TO GEMI

StringUpper, CategoryCode, CategoryCode

Send, {F4}
Sleep, 1500

if (CategoryCode = "EAR")
{
    Send, e
    Sleep, 500
    Send, {Enter}
}
else
{
    Loop, Parse, CategoryCode
    {
        Send, {Raw}%A_LoopField%
        Sleep, 150
    }

    Sleep, 500
    Send, {Enter}
}

Sleep, 500
; ============================================================
; 11. SUB CATEGORY - AUTOMATIC CATEGORY-BASED SELECTION
; ============================================================

Send, {Tab}
Sleep, 1000

SubCategorySelected := ""
SubCategoryFirstLetter := ""
SubCategoryDownCount := 0

StringUpper, CategoryCode, CategoryCode
CategoryCode := Trim(CategoryCode)

; Column W = Sub Category
ExcelSubCategory := ws.Range("W2").Value
ExcelSubCategory := Trim(ExcelSubCategory)
StringLower, ExcelSubCategory, ExcelSubCategory


; ============================================================
; CATEGORY-BASED SUB CATEGORY MAPPING
; ============================================================

; RING
if (CategoryCode = "RNG")
{
    if (ExcelSubCategory = "anniversary")
    {
        SubCategorySelected := "ANNIV"
    }
    else if (ExcelSubCategory = "bridal ring")
    {
        SubCategorySelected := "BRD-RNG"
    }
    else if (ExcelSubCategory = "engagement ring")
    {
        SubCategorySelected := "ENG-RNG"
    }
    else if (ExcelSubCategory = "fashion")
    {
        SubCategorySelected := "FASH"
    }
    else if (ExcelSubCategory = "gents ring")
    {
        SubCategorySelected := "GNT-RNG"
    }
    else if (ExcelSubCategory = "guard round")
    {
        SubCategorySelected := "GRD-RND"
    }
    else if (ExcelSubCategory = "ladies ring")
    {
        SubCategorySelected := "LDS-RNG"
    }
    else if (ExcelSubCategory = "wedding ring")
    {
        SubCategorySelected := "WED-RNG"
    }
    else if (ExcelSubCategory = "wpr ring")
    {
        SubCategorySelected := "WPR-RNG"
    }
}


; EARRING
else if (CategoryCode = "EAR")
{
    if (ExcelSubCategory = "earring")
    {
        SubCategorySelected := "EAR"
    }
    else if (ExcelSubCategory = "huggie earring")
    {
        SubCategorySelected := "HUG-EAR"
    }
    else if (ExcelSubCategory = "stud earring")
    {
        SubCategorySelected := "STUD-EAR"
    }
}


; PENDANT
else if (CategoryCode = "PND")
{
    if (ExcelSubCategory = "pendant")
    {
        SubCategorySelected := "PND"
    }
}


; NECKLACE
else if (CategoryCode = "NCK")
{
    if (ExcelSubCategory = "necklace")
    {
        SubCategorySelected := "NCK"
    }
    else if (ExcelSubCategory = "necklace pendant")
    {
        SubCategorySelected := "NCK-PND"
    }
}


; BRACELET
else if (CategoryCode = "BRC")
{
    if (ExcelSubCategory = "bracelet")
    {
        SubCategorySelected := "BRC"
    }
    else if (ExcelSubCategory = "bracelet pendant")
    {
        SubCategorySelected := "BRC-PND"
    }
}


; BANGLE
else if (CategoryCode = "BNG")
{
    if (ExcelSubCategory = "bangle")
    {
        SubCategorySelected := "BNG"
    }
}


; HEAD
else if (CategoryCode = "HD")
{
    if (ExcelSubCategory = "head")
    {
        SubCategorySelected := "HD"
    }
}


; BROOCH
else if (CategoryCode = "BRCH")
{
    if (ExcelSubCategory = "brooch")
    {
        SubCategorySelected := "BRCH"
    }
}


; CUFFLINK
else if (CategoryCode = "CFL")
{
    if (ExcelSubCategory = "cufflink")
    {
        SubCategorySelected := "CFL"
    }
}


; CROWN
else if (CategoryCode = "CRN")
{
    SubCategorySelected := ""
}


; ============================================================
; AUTOMATIC JEWELSOLUTION SELECTION
; ============================================================

if (SubCategorySelected != "")
{
    ; First letter + down count
    if (SubCategorySelected = "ANNIV")
    {
        SubCategoryFirstLetter := "a"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "BRD-RNG")
    {
        SubCategoryFirstLetter := "b"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "ENG-RNG")
    {
        SubCategoryFirstLetter := "e"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "FASH")
    {
        SubCategoryFirstLetter := "f"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "GNT-RNG")
    {
        SubCategoryFirstLetter := "g"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "GRD-RND")
    {
        SubCategoryFirstLetter := "g"
        SubCategoryDownCount := 1
    }
    else if (SubCategorySelected = "LDS-RNG")
    {
        SubCategoryFirstLetter := "l"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "WED-RNG")
    {
        SubCategoryFirstLetter := "w"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "WPR-RNG")
    {
        SubCategoryFirstLetter := "w"
        SubCategoryDownCount := 1
    }
    else if (SubCategorySelected = "EAR")
    {
        SubCategoryFirstLetter := "e"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "HUG-EAR")
    {
        SubCategoryFirstLetter := "h"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "STUD-EAR")
    {
        SubCategoryFirstLetter := "s"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "PND")
    {
        SubCategoryFirstLetter := "p"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "NCK")
    {
        SubCategoryFirstLetter := "n"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "NCK-PND")
    {
        SubCategoryFirstLetter := "n"
        SubCategoryDownCount := 1
    }
    else if (SubCategorySelected = "BRC")
    {
        SubCategoryFirstLetter := "b"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "BRC-PND")
    {
        SubCategoryFirstLetter := "b"
        SubCategoryDownCount := 1
    }
    else if (SubCategorySelected = "BNG")
    {
        SubCategoryFirstLetter := "b"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "HD")
    {
        SubCategoryFirstLetter := "h"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "BRCH")
    {
        SubCategoryFirstLetter := "b"
        SubCategoryDownCount := 0
    }
    else if (SubCategorySelected = "CFL")
    {
        SubCategoryFirstLetter := "c"
        SubCategoryDownCount := 0
    }


    ; Automatic selection
    Send, %SubCategoryFirstLetter%
    Sleep, 500

    Loop, %SubCategoryDownCount%
    {
        Send, {Down}
        Sleep, 200
    }

    Send, {Enter}
    Sleep, 2000
}


Goto, ContinueAfterSubCategory


ContinueAfterSubCategory:

Sleep, 500

; 12. CODE FIELD
Send, {Tab}

Sleep, 125

CodeToType := RegExReplace(ProductCode, "\{[^}]*\}", "")
CodeToType := Trim(CodeToType)

Loop, Parse, CodeToType
{
    Send, {Raw}%A_LoopField%
    Sleep, 30
}

Sleep, 1250

Send, {Enter}

Sleep, 500

; 13. UNIT FIELD

Send, {Tab 2}

Sleep, 500

Send, {F4}

Sleep, 2000

if (ProductTypeLower = "earrings" || ProductTypeLower = "earring")
{
    Send, p
    Sleep, 500
    Send, {Down}
    Sleep, 500
    Send, {Enter}
}
else
{
    Send, p
    Sleep, 500
    Send, {Enter}
}

Sleep, 1500

; 14. METAL FIELD
Send, {Tab}

Sleep, 1750
; READ METAL FROM EXCEL I2
try

{

    MetalType := ws.Range("I2").Value

    MetalType := Trim(MetalType)

    StringLower, MetalTypeLower, MetalType

    MetalTypeLower := Trim(MetalTypeLower)

}

catch

{

    MsgBox, 48, Error, Excel I2 read nahi ho saka.

    return

}
;  DETECT METAL CODE
MetalCode := ""

if InStr(MetalTypeLower, "silver")

{

    MetalCode := "S"

}

else if InStr(MetalTypeLower, "gold")

{

    MetalCode := "G"

}

else if InStr(MetalTypeLower, "10kt")

{

    MetalCode := "G"

}

else if InStr(MetalTypeLower, "9kt")

{

    MetalCode := "G"

}

else if InStr(MetalTypeLower, "14kt")

{

    MetalCode := "G"

}

else if InStr(MetalTypeLower, "18kt")

{

    MetalCode := "G"

}

else if InStr(MetalTypeLower, "22kt")

{

    MetalCode := "G"

}
else if InStr(MetalTypeLower, "24kt")

{

    MetalCode := "G"

}

else if InStr(MetalTypeLower, "alloy")

{

    MetalCode := "A"

}

else if InStr(MetalTypeLower, "brass")

{

    MetalCode := "B"

}

else if InStr(MetalTypeLower, "palladium")

{

    MetalCode := "L"

}

else if InStr(MetalTypeLower, "platinum")

{

    MetalCode := "P"

}

else

{

    MsgBox, 48, Unknown Metal, Excel I2 mein metal identify nahi hua:`n`n%MetalType%

    return

}
Send, %MetalCode%

Sleep, 700

Send, {Enter}

Sleep, 700

; 15. KT FIELD
; ============================================================
; KT FIELD - AUTOMATIC FROM EXCEL I2
; ============================================================

Send, {Tab}
Sleep, 500

KTSelected := ""

; ============================================================
; GOLD KT
; ============================================================

if (MetalCode = "G")
{
    if RegExMatch(MetalTypeLower, "03\s*kt")
    {
        KTSelected := "G03"
    }
    else if RegExMatch(MetalTypeLower, "08\s*kt")
    {
        KTSelected := "G08"
    }
    else if RegExMatch(MetalTypeLower, "09\s*kt")
    {
        KTSelected := "G09"
    }
    else if RegExMatch(MetalTypeLower, "10\s*kt")
    {
        KTSelected := "G10"
    }
    else if RegExMatch(MetalTypeLower, "14\s*kt")
    {
        KTSelected := "G14"
    }
    else if RegExMatch(MetalTypeLower, "18\s*kt")
    {
        KTSelected := "G18"
    }
    ; Check G24R before G24
    else if RegExMatch(MetalTypeLower, "24\s*kt\s*r")
    {
        KTSelected := "G24R"
    }
    else if RegExMatch(MetalTypeLower, "24\s*kt")
    {
        KTSelected := "G24"
    }
    else
    {
        MsgBox, 48, KT Error, Unable to identify the Gold KT value from Excel I2.`n`nMetal Type: %MetalType%
        return
    }
}


; SILVER
else if (MetalCode = "S")
{
    KTSelected := "SL925"
}
; ALLOY
else if (MetalCode = "A")
{
    KTSelected := "A"
}
; BRASS
else if (MetalCode = "B")
{
    KTSelected := "B"
}
; PLATINUM
else if (MetalCode = "P")
{
    KTSelected := "P"
}
; PALLADIUM - no KT required
else if (MetalCode = "L")
{
    KTSelected := ""
}
; UNKNOWN METAL CODE
else
{
    MsgBox, 48, KT Error, Unknown Metal Code.`n`nMetal Code: %MetalCode%`nMetal Type: %MetalType%
    return
}


; ============================================================
; SELECT KT IN GEMI
; ============================================================

if (KTSelected != "")
{
    ; Open KT dropdown, type code slowly, confirm
    Send, {F4}
    Sleep, 500

    Loop, Parse, KTSelected
    {
        Send, {Raw}%A_LoopField%
        Sleep, 75
    }

    Sleep, 750

    Send, {Enter}
    Sleep, 1500
}


Goto, ContinueAfterKT


ContinueAfterKT:

; 16. COLOR FIELD

Send, {Tab}
Sleep, 1000

ColorSelected := ""
ColorChoice := ""

ColorOptions := "R | Rose`nTT | TwoTone`nW | White`nY | Yellow"

Gui, Color:Destroy
Gui, Color:New, +AlwaysOnTop +ToolWindow +HwndColorHwnd, Select Color
Gui, Color:Font, s10, Segoe UI
Gui, Color:+Delimiter`n

Gui, Color:Add, Text, x20 y15 w300 h25, Select Color:
Gui, Color:Add, ListBox, x20 y45 w300 h120 vColorChoice Choose1, %ColorOptions%
Gui, Color:Add, Button, x95 y185 w150 h35 gColorOK Default, Select

Gui, Color:Show, w340 h240, Select Color

WinWaitClose, ahk_id %ColorHwnd%

if (ColorSelected = "")
    return

WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 750

Send, %ColorSelected%
Sleep, 750

Send, {Enter}
Sleep, 1500

Goto, ContinueAfterColor


ColorOK:

Gui, Color:Submit

if (ColorChoice = "")
{
    Gui, Color:Show
    return
}

StringSplit, ColorParts, ColorChoice, |

ColorSelected := Trim(ColorParts1)

Gui, Color:Destroy
return


ColorGuiClose:

ColorSelected := ""

Gui, Color:Destroy
return


ContinueAfterColor:
; 17. SIZE FIELD
Send, {Tab}

Sleep, 2500
;  SIZE
SizeSelected := ""

SizeChoice := ""

SizeOptions := "B100 | Brc 10.00 Inch`n"

SizeOptions .= "B600 | Brc 6.00 Inch`n"

SizeOptions .= "B625 | Brc 6.25 Inch`n"

SizeOptions .= "B650 | Brc 6.50 Inch`n"

SizeOptions .= "B675 | Brc 6.75 Inch`n"

SizeOptions .= "B700 | Brc 7.00 Inch`n"

SizeOptions .= "B725 | Brc 7.25 Inch`n"

SizeOptions .= "B750 | Brc 7.5 Inch`n"

SizeOptions .= "B775 | Brc 7.75 Inch`n"

SizeOptions .= "B800 | Brc 8.00 Inch`n"

SizeOptions .= "B825 | Brc 8.25 Inch`n"

SizeOptions .= "B850 | Brc 8.50 Inch`n"

SizeOptions .= "B875 | Brc 8.75 Inch`n"

SizeOptions .= "B900 | Brc 9.00 Inch`n"

SizeOptions .= "B950 | Brc 9.50 Inch`n"

SizeOptions .= "D | D`n"

SizeOptions .= "E | E`n"

SizeOptions .= "E1/2 | E1/2`n"

SizeOptions .= "EJ1/2 | EJ1/2`n"

SizeOptions .= "EN | EN`n"

SizeOptions .= "F | F`n"

SizeOptions .= "F1/2 | F1/2`n"

SizeOptions .= "G | G`n"

SizeOptions .= "G1/2 | G1/2`n"

SizeOptions .= "H | H`n"

SizeOptions .= "H1/2 | H1/2`n"

SizeOptions .= "I | I`n"

SizeOptions .= "I1/2 | I1/2`n"

SizeOptions .= "J | J`n"

SizeOptions .= "J+ | J+`n"

SizeOptions .= "J1/2 | J1/2`n"

SizeOptions .= "K | K`n"

SizeOptions .= "K1/2 | K1/2`n"

SizeOptions .= "L | L`n"

SizeOptions .= "L1/2 | L1/2`n"

SizeOptions .= "M1/2 | M1/2`n"

SizeOptions .= "N | N`n"

SizeOptions .= "N/A | NOT APPLICABLE`n"

SizeOptions .= "N1/2 | N1/2`n"

SizeOptions .= "N17 | Nck 17 Inch`n"

SizeOptions .= "N18 | Nck 18 Inch`n"

SizeOptions .= "N19 | Nck 19 Inch`n"

SizeOptions .= "N20 | Nck 20 Inch`n"

SizeOptions .= "O | O`n"

SizeOptions .= "O1/2 | O1/2`n"

SizeOptions .= "P | P`n"

SizeOptions .= "P1/2 | P1/2`n"

SizeOptions .= "Q | Q`n"

SizeOptions .= "Q1/2 | Q1/2`n"

SizeOptions .= "R | R`n"

SizeOptions .= "R1/2 | R1/2`n"

SizeOptions .= "R56CM | R56CM`n"

SizeOptions .= "S | S`n"

SizeOptions .= "S1/2 | S1/2`n"

SizeOptions .= "STD | STD`n"

SizeOptions .= "T | T`n"

SizeOptions .= "T1/2 | T1/2`n"

SizeOptions .= "U | U`n"

SizeOptions .= "U1/2 | U1/2`n"

SizeOptions .= "US10.00 | US10.00`n"

SizeOptions .= "US10.25 | US10.25`n"

SizeOptions .= "US10.50 | US10.50`n"

SizeOptions .= "US10.75 | US10.75`n"

SizeOptions .= "US11.00 | US11.00`n"

SizeOptions .= "US11.25 | US11.25`n"

SizeOptions .= "US11.50 | US11.50`n"

SizeOptions .= "US11.75 | US11.75`n"

SizeOptions .= "US12.00 | US12.00`n"

SizeOptions .= "US12.25 | US12.25`n"

SizeOptions .= "US12.50 | US12.50`n"

SizeOptions .= "US12.75 | US12.75`n"

SizeOptions .= "US13.00 | US13.00`n"

SizeOptions .= "US13.25 | US13.25`n"

SizeOptions .= "US13.50 | US13.50`n"

SizeOptions .= "US14.00 | US14.00`n"

SizeOptions .= "US14.25 | US14.25`n"

SizeOptions .= "US14.50 | US14.50`n"

SizeOptions .= "US15.00 | US15.00`n"

SizeOptions .= "US15.25 | US15.25`n"

SizeOptions .= "US15.50 | US15.50`n"

SizeOptions .= "US17.00 | US17.00`n"

SizeOptions .= "US2.00 | US2.00`n"

SizeOptions .= "US2.25 | US2.25`n"

SizeOptions .= "US2.50 | US2.50`n"

SizeOptions .= "US2.75 | US2.75`n"

SizeOptions .= "US3.00 | US3.00`n"

SizeOptions .= "US3.25 | US3.25`n"

SizeOptions .= "US3.50 | US3.50`n"

SizeOptions .= "US4.00 | US4.00`n"

SizeOptions .= "US4.25 | US4.25`n"

SizeOptions .= "US4.50 | US4.50`n"

SizeOptions .= "US4.75 | US4.75`n"

SizeOptions .= "US5.00 | US5.00`n"

SizeOptions .= "Y | Y`n"

SizeOptions .= "Y1/2 | Y1/2`n"

SizeOptions .= "Z | Z`n"

SizeOptions .= "Z+1 | Z+1`n"

SizeOptions .= "Z+2 | Z+2`n"

SizeOptions .= "Z+2 1/2 | Z+2 1/2"

Gui, Size:New, +AlwaysOnTop +ToolWindow +HwndSizeHwnd, Select Size

Gui, Size:Font, s10, Segoe UI

Gui, Size:+Delimiter`n

Gui, Size:Add, Text, x20 y15 w300 h25, Select Size:

Gui, Size:Add, ListBox, x20 y45 w300 h300 vSizeChoice, %SizeOptions%

Gui, Size:Add, Button, x95 y360 w150 h35 gSizeOK Default, Select

Gui, Size:Show, w340 h415, Select Size

WinWaitClose, ahk_id %SizeHwnd%

if (SizeSelected = "")

    return

WinActivate, ahk_exe JwelSolution.exe

WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel

    return

Sleep, 1250

Send, %SizeSelected%

Sleep, 1250

Send, {Enter}

Sleep, 2500

Goto, ContinueAfterSize


SizeOK:

Gui, Size:Submit

if (SizeChoice = "")

{

    Gui, Size:Show

    return

}

StringSplit, SizeParts, SizeChoice, |

SizeSelected := Trim(SizeParts1)

Gui, Size:Destroy


return


SizeGuiClose:

SizeSelected := ""

Gui, Size:Destroy

return


ContinueAfterSize:



; 28. COLLECTION FIELD

Send, {Tab 2}
Sleep, 1000


; Read Collection from Excel A2
CollectionSearch := ws.Range("A2").Value
CollectionSearch := Trim(CollectionSearch)

if (CollectionSearch = "")
{
    MsgBox, 48, Collection Error, Excel A2 collection value is empty.
    return
}


; First 4 characters, lowercased
CollectionSearch := SubStr(CollectionSearch, 1, 4)

StringLower, CollectionSearch, CollectionSearch

CollectionSearch := Trim(CollectionSearch)


; Collection options
CollectionOptions := "AG|AG DIAMOND`n"
CollectionOptions .= "BE|Brilliant Earth`n"
CollectionOptions .= "BEACON|Beacon`n"
CollectionOptions .= "BLK RHD|Black Rhodium`n"
CollectionOptions .= "BLUNLE|Blue Nile`n"
CollectionOptions .= "BUBBLE|BUBBLE`n"
CollectionOptions .= "BUBBLE NEW|BUBBLE NEW`n"
CollectionOptions .= "CARO|Caro`n"
CollectionOptions .= "CORE LINE|Core Line`n"
CollectionOptions .= "DOTCOM|Dotcom`n"
CollectionOptions .= "EJ|EJ`n"
CollectionOptions .= "ENSEMBLE|Ensemble`n"
CollectionOptions .= "ETERNAL SMILE|Eternal Smile.`n"
CollectionOptions .= "EVERUS|EverUs`n"
CollectionOptions .= "FC|FC`n"
CollectionOptions .= "FOUNDE|Founde`n"
CollectionOptions .= "FRED MEYER|Fred Meyer`n"
CollectionOptions .= "HELZBERG|Helzberg`n"
CollectionOptions .= "I|I`n"
CollectionOptions .= "IMP STUDD|IMP STUDD`n"
CollectionOptions .= "IMP STUDD NEW|IMP STUDD NEW`n"
CollectionOptions .= "JA|JAMES ALLEN`n"
CollectionOptions .= "JARED|Jared`n"
CollectionOptions .= "JARED-LCD|Jared-LCD`n"
CollectionOptions .= "JRA|Jared Royal Asscher`n"
CollectionOptions .= "JV|JV`n"
CollectionOptions .= "KAY|Kay`n"
CollectionOptions .= "KAY FASHION|Kay Fashion`n"
CollectionOptions .= "KAY-GENDER-NEUTRAL|Kay-Gender Neutral`n"
CollectionOptions .= "KAY-OUTLET|Kay outlet`n"
CollectionOptions .= "KOHL-US134|KOHL-US134`n"
CollectionOptions .= "LAB GROWN|Lab Grown`n"
CollectionOptions .= "LONDON|London`n"
CollectionOptions .= "LOOP LOGO|Loop Logo`n"
CollectionOptions .= "MAYORS|Mayors`n"
CollectionOptions .= "MILESTONE|Milestone`n"
CollectionOptions .= "MILESTONE-NEW STY|Milestone-New Style`n"
CollectionOptions .= "NEIL LANE|NEIL LANE`n"
CollectionOptions .= "NEIL LANE NEW|NEIL LANE NEW`n"
CollectionOptions .= "NEW FOUNDE|New Founde`n"
CollectionOptions .= "NEW HELZBERG|New Helzberg`n"
CollectionOptions .= "NEW STARS BY THE YARD|New Stars By The Yard`n"
CollectionOptions .= "NEW STY|NEW STYLE`n"
CollectionOptions .= "NEW STY MAYORS|New Style Mayors`n"
CollectionOptions .= "NEW STY-BEACON|New Style Beacon`n"
CollectionOptions .= "NEW STY-BLK RHD|New Style Black Rhodium`n"
CollectionOptions .= "NEW STY-KAY FASH|New Style Kay Fashion`n"
CollectionOptions .= "NEW STY-MP|NEW STY With Miracle Plate`n"
CollectionOptions .= "NEW STY-RUTHENIUM PLT|New Style Ruthenium Plating`n"
CollectionOptions .= "NEW YORK|New York`n"
CollectionOptions .= "PARIS|Paris`n"
CollectionOptions .= "R2NET|R2NET`n"
CollectionOptions .= "R2NET-BN|R2NET-BN`n"
CollectionOptions .= "R2NET-JA|R2NET-JA`n"
CollectionOptions .= "RC|RC`n"
CollectionOptions .= "RD|RD`n"
CollectionOptions .= "RD-KENFLD|Reeds Kenfield`n"
CollectionOptions .= "RI|Riddles`n"
CollectionOptions .= "ROYASS|Royal Asscher`n"
CollectionOptions .= "ROYASS-HK|Royal Asscher-Hong Kong`n"
CollectionOptions .= "ROYASS-NY|Royal Asscher-New York`n"
CollectionOptions .= "RUTHENIUM PLT|Ruthenium Plating`n"
CollectionOptions .= "SD|Super diamond`n"
CollectionOptions .= "SD-MP|Super Diamond With Miracle Plate`n"
CollectionOptions .= "SDC-MP|SDC With Miracle Plate`n"
CollectionOptions .= "SGN-LVSTR|Signet Love Story`n"
CollectionOptions .= "SJ|SJ`n"
CollectionOptions .= "SOLITAIRE|Solitaire`n"
CollectionOptions .= "Special Ord|Special Order`n"
CollectionOptions .= "ST|ST`n"
CollectionOptions .= "ST-US81|ST-US81`n"
CollectionOptions .= "ST-US88|ST-US88`n"
CollectionOptions .= "STARS BY THE YARD|Stars By The Yard`n"
CollectionOptions .= "SW|SIGNET WEB`n"
CollectionOptions .= "TJ|TJ`n"
CollectionOptions .= "UK-DVS|UK Division`n"
CollectionOptions .= "UNSPOKEN|UNSPOKEN`n"
CollectionOptions .= "UNSPOKEN-NEW|UNSPOKEN-NEW`n"
CollectionOptions .= "UNSPOKEN-NEW STY|UNSPOKEN-NEW STYLE.`n"
CollectionOptions .= "US117|US117`n"
CollectionOptions .= "US71.COM|US71.COM`n"
CollectionOptions .= "US73|US73`n"
CollectionOptions .= "VALINA|Valina`n"
CollectionOptions .= "VALINA-ALY|Valina-Alloy`n"
CollectionOptions .= "VALINA-SIL|Valina-Silver`n"
CollectionOptions .= "VERA|Vera`n"
CollectionOptions .= "VERA-NEW STYLE|Vera-New Style`n"
CollectionOptions .= "VW-LV|Vera Wang Love`n"
CollectionOptions .= "XLAB|XLAB (Lgd & Dia Combination)`n"
CollectionOptions .= "ZALES|Zales`n"
CollectionOptions .= "ZALES-LCD|Zales-LCD`n"
CollectionOptions .= "ZALES-LCD-KLD|Zales-LCD-KLD (KLEINFELD)`n"
CollectionOptions .= "ZC|Zales Canada`n"
CollectionOptions .= "ZC-US71C|ZC-US71C`n"
CollectionOptions .= "ZC-US71C-B10|ZC-US71C-B10`n"
CollectionOptions .= "ZC-US71CDOS|ZC-US71CDOS`n"
CollectionOptions .= "ZC-US71CGRDO|ZC-US71CGRDO`n"
CollectionOptions .= "ZCAN|ZCAN`n"
CollectionOptions .= "ZD|Zales Division`n"
CollectionOptions .= "ZD-US71|ZD-US71`n"
CollectionOptions .= "ZD-US71-B10|ZD-US71-B10`n"
CollectionOptions .= "ZI|Zale dotcom`n"
CollectionOptions .= "ZO-US85|ZO-US85`n"
CollectionOptions .= "ZO-US85-GR|ZO-US85-GR`n"
CollectionOptions .= "ZO-US85SM|ZO-US85SM"


; Filter by the search prefix
FilteredCollectionOptions := ""
CollectionMatchCount := 0
CollectionSingleMatch := ""

Loop, Parse, CollectionOptions, `n, `r
{
    StringSplit, CollectionParts, A_LoopField, |

    CollectionCodeCandidate := Trim(CollectionParts1)
    CollectionNameCandidate := Trim(CollectionParts2)

    StringLower, CodeCheck, CollectionCodeCandidate
    StringLower, NameCheck, CollectionNameCandidate

    CodeCheck := RegExReplace(CodeCheck, "[^a-z0-9]+", "")
    NameCheck := RegExReplace(NameCheck, "[^a-z0-9]+", "")

    StringReplace, NameCheck, NameCheck, diamonds, diamond, All

    if (InStr(CodeCheck, CollectionSearch) || InStr(NameCheck, CollectionSearch))
    {
        CollectionMatchCount++

        CollectionSingleMatch := A_LoopField

        if (FilteredCollectionOptions = "")
            FilteredCollectionOptions := A_LoopField
        else
            FilteredCollectionOptions .= "`n" A_LoopField
    }
}


; No match - ask the user
if (CollectionMatchCount = 0)
{
    InputBox, CollectionSelected, New Collection, No existing collection found.`n`nSearch:`n%CollectionSearch%`n`nEnter Collection manually:

    if ErrorLevel
        return

    CollectionSelected := Trim(CollectionSelected)

    if (CollectionSelected = "")
    {
        MsgBox, 48, Collection Error, Collection cannot be empty.
        return
    }

    StringLower, CollectionSelected, CollectionSelected

    WinActivate, ahk_exe JwelSolution.exe
    WinWaitActive, ahk_exe JwelSolution.exe,, 10

    if ErrorLevel
        return

    Sleep, 1000

    Send, %CollectionSelected%

    Sleep, 1000

    Send, {Enter}

    Sleep, 2000

    Goto, ContinueAfterCollection
}


; Exactly one match - select it automatically
if (CollectionMatchCount = 1)
{
    StringSplit, CollectionParts, CollectionSingleMatch, |

    CollectionSelected := Trim(CollectionParts1)

    StringLower, CollectionSelected, CollectionSelected

    WinActivate, ahk_exe JwelSolution.exe
    WinWaitActive, ahk_exe JwelSolution.exe,, 10

    if ErrorLevel
        return

    Sleep, 1000

    Send, %CollectionSelected%

    Sleep, 1000

    Send, {Enter}

    Sleep, 2000

    Goto, ContinueAfterCollection
}


; Multiple matches - show popup
CollectionSelected := ""
CollectionChoice := ""

Gui, Collection:Destroy
Gui, Collection:New, +AlwaysOnTop +ToolWindow +HwndCollectionHwnd, Select Collection
Gui, Collection:Font, s10, Segoe UI
Gui, Collection:+Delimiter`n

Gui, Collection:Add, Text, x20 y15 w420 h25, Multiple collections found. Select one:

Gui, Collection:Add, ListBox, x20 y45 w420 h400 vCollectionChoice Choose1, %FilteredCollectionOptions%

Gui, Collection:Add, Button, x155 y460 w150 h35 gCollectionOK Default, Select

Gui, Collection:Show, w460 h515, Select Collection


WinWaitClose, ahk_id %CollectionHwnd%


if (CollectionSelected = "")
    return


WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 1000

StringLower, CollectionSelected, CollectionSelected

Send, %CollectionSelected%

Sleep, 1000

Send, {Enter}

Sleep, 1000

Goto, ContinueAfterCollection


CollectionOK:

Gui, Collection:Submit

if (CollectionChoice = "")
{
    Gui, Collection:Show
    return
}

StringSplit, CollectionParts, CollectionChoice, |

CollectionSelected := Trim(CollectionParts1)

StringLower, CollectionSelected, CollectionSelected

Gui, Collection:Destroy

return


CollectionGuiClose:

CollectionSelected := ""

Gui, Collection:Destroy

return


ContinueAfterCollection:
; 19. DESIGN CODE

Send, {Tab 3}
Sleep, 500

DesignCode := ws.Range("V2").Value
DesignCode := Trim(DesignCode)

if (DesignCode = "")
{
    MsgBox, 48, Design Code Error, Active Excel V2 contains no Design Code.
    return
}

; Remove anything after the first (
DesignCode := RegExReplace(DesignCode, "\(.+$", "")
DesignCode := Trim(DesignCode)

; Send Design Code to GEMI
Send, %DesignCode%

Sleep, 550
Send, {Enter}
Sleep, 200


MouseMove, 687, 718 
Click, Left
Sleep, 566
MouseMove, 938, 604
Click, Left
Sleep, 566


Send, +{F10}
Sleep, 500
Send, {Down 2}
Send, {Enter}

Sleep, 1000

;Remark feild 1

Sleep, 300


KeyWait, Enter, D
Sleep, 500



; REMARK 2

Send, {Tab}
Sleep, 200

StringUpper, DesignCodeUpper, DesignCode
Send, %DesignCodeUpper%

Send, {Enter}


ComponentCount := ws.Range("E2").Value
ComponentCount := Trim(ComponentCount)

ComponentDetails := ws.Range("F2").Value
ComponentDetails := Trim(ComponentDetails)


; Remove numbers after colon
ComponentDetails := RegExReplace(ComponentDetails, ":[ ]*[0-9]+", "")

; Replace semicolons with spaces
ComponentDetails := RegExReplace(ComponentDetails, "[;]+", " ")

; Remove multiple spaces
ComponentDetails := RegExReplace(ComponentDetails, "\s+", " ")

ComponentDetails := Trim(ComponentDetails)


; Convert component details to uppercase
StringUpper, ComponentDetailsUpper, ComponentDetails


ExtraComponentName := ws.Range("G2").Value
ExtraComponentName := Trim(ExtraComponentName)

; Remove numbers after colon
ExtraComponentName := RegExReplace(ExtraComponentName, ":[ ]*[0-9]+", "")

; Replace semicolons with spaces
ExtraComponentName := RegExReplace(ExtraComponentName, "[;]+", " ")

; Remove multiple spaces
ExtraComponentName := RegExReplace(ExtraComponentName, "\s+", " ")

ExtraComponentName := Trim(ExtraComponentName)

; Convert extra component name to uppercase
StringUpper, ExtraComponentNameUpper, ExtraComponentName


Send, %ComponentDetailsUpper%

Sleep, 700
Send, {Enter}

Sleep, 1500


if (ComponentCount >= 2)
{
    ; Assembly pricing remark
    Send, %ComponentCount% PIECES ASSEMBLY CHARGE INCLUDE IN PRICING
    Sleep, 200
    Send, {Enter}

    Sleep, 200

    ; Use actual extra component name
    Send, ASSEMBLY EXTRA %ExtraComponentNameUpper%
    Sleep, 300
    Send, {Enter}

    Sleep, 300
}



; REMARK 3

Send, {Tab}
Sleep, 500

Remark3 := ""

if (ColorSelected = "W")
{
    Remark3 := ""
}
else if (ColorSelected = "Y")
{
   
    Remark3 := "NO RHODIUM"
}
else if (ColorSelected = "R")
{
    
    Remark3 := "NO RHODIUM"
}
else if (ColorSelected = "TT")
{
   

    InputBox, Remark3, Remark 3, Please enter Remark 3 message:

    if ErrorLevel
        return

    Remark3 := Trim(Remark3)

    if (Remark3 = "")
        return
}

if (Remark3 != "")
    Send, %Remark3%

Sleep, 750
Send, {Enter}
Sleep, 1500




; REMARK 4 - STAMPING DETAILS

Send, {Tab}
Sleep, 750

Remark4Selected := ""
Remark4Choice := ""

; Stamping detail options
Remark4Options := "14K ""SDC LOGO"" CZ Resin`n"
Remark4Options .= "10K ""SDC LOGO"" CZ Resin`n"
Remark4Options .= "18K ""SDC LOGO"" CZ Resin`n"
Remark4Options .= "925 ""SDC LOGO"" CZ Resin`n"
Remark4Options .= "14K ""SDC LOGO"" CZ`n"
Remark4Options .= "10K ""SDC LOGO"" CZ`n"
Remark4Options .= "18K ""SDC LOGO"" CZ`n"
Remark4Options .= "14K ""SDC LOGO""`n"
Remark4Options .= "10K ""SDC LOGO""`n"
Remark4Options .= "18K ""SDC LOGO""`n"
Remark4Options .= "LASER WITH KLEINFELD LOGO`n"
Remark4Options .= "14K ""vera wang Love"" CZ Resin`n"
Remark4Options .= "10K ""vera wang Love"" CZ Resin`n"
Remark4Options .= "18K ""vera wang Love"" CZ Resin`n"
Remark4Options .= "925 ""vera wang Love"" CZ Resin`n"
Remark4Options .= "14K ""vera wang Love"" CZ`n"
Remark4Options .= "10K ""vera wang Love"" CZ`n"
Remark4Options .= "18K ""vera wang Love"" CZ`n"
Remark4Options .= "14K ""vera wang Love""`n"
Remark4Options .= "10K ""vera wang Love""`n"
Remark4Options .= "18K ""vera wang Love""`n"
Remark4Options .= "14K ""vera wang Wish"" CZ Resin`n"
Remark4Options .= "10K ""vera wang Wish"" CZ Resin`n"
Remark4Options .= "18K ""vera wang Wish"" CZ Resin`n"
Remark4Options .= "925 ""vera wang Wish"" CZ Resin`n"
Remark4Options .= "14K ""vera wang Wish"" CZ`n"
Remark4Options .= "10K ""vera wang Wish"" CZ`n"
Remark4Options .= "18K ""vera wang Wish"" CZ`n"
Remark4Options .= "925 ""vera wang Wish"" CZ`n"
Remark4Options .= "14K ""vera wang Wish""`n"
Remark4Options .= "10K ""vera wang Wish""`n"
Remark4Options .= "18K ""vera wang Wish""`n"
Remark4Options .= "925 ""vera wang Wish""`n"
Remark4Options .= "14K""Loop Logo""`n"
Remark4Options .= "925 ""Loop Logo""`n"
Remark4Options .= "10K ""Loop Logo""`n"
Remark4Options .= "18K ""Loop Logo""`n"
Remark4Options .= "14K""Loop Logo"" Resin`n"
Remark4Options .= "925 ""Loop Logo"" Resin`n"
Remark4Options .= "10K ""Loop Logo"" Resin`n"
Remark4Options .= "18K ""Loop Logo"" Resin`n"
Remark4Options .= "14K ""Ever US Logo""`n"
Remark4Options .= "925 ""vera wang Wish""`n"
Remark4Options .= "14K ""SAY I DO"" CZ Resin`n"
Remark4Options .= "925 ""SAY I DO"" CZ Resin`n"
Remark4Options .= "14K ""SAY I DO"" CZ`n"
Remark4Options .= "925 ""SAY I DO"" CZ`n"
Remark4Options .= "14K ""SAY I DO""`n"
Remark4Options .= "925 ""SAY I DO""`n"
Remark4Options .= "Royal Asscher Logo`n"
Remark4Options .= "925 ""SDC LOGO"" CZ`n"
Remark4Options .= "14KT-CZ Mile stone Logo`n"
Remark4Options .= "14KT-CZ & VERA SIGNATURE LOGO`n"
Remark4Options .= "VALINA 925/CZ ( DO BLACK ANTIQUE)`n"
Remark4Options .= "4KT-CZ & SDC LOGO & EVERLEDGER LOGO`n"
Remark4Options .= "14KT-CZ & SDC LOGO & EVERLEDGER LOGO`n"
Remark4Options .= "14KT-CZ & SDC LOGO`n"
Remark4Options .= "925-CZ & VERA WANG LOVE & SDC LOGO`n"
Remark4Options .= "14KT-CZ & VERA WANG LOVE & SDC LOGO`n"
Remark4Options .= "925-CZ & SDC LOGO`n"
Remark4Options .= "925-CZ & VERA WANG LOVE & SDC LOGO`n"
Remark4Options .= "14KT SD LAB & LAB CREATED`n"
Remark4Options .= "18KT-CZ & SDC LOGO`n"
Remark4Options .= "PT950-CZ & SDC LOGO`n"
Remark4Options .= "14KT & SDC LOGO & CZ CENTER`n"
Remark4Options .= "14KT & SDC LOGO & CHOSEN LOGO"


Gui, Remark4:Destroy
Gui, Remark4:New, +AlwaysOnTop +ToolWindow +HwndRemark4Hwnd, Select Stamping Details
Gui, Remark4:Font, s10, Segoe UI
Gui, Remark4:+Delimiter`n
Gui, Remark4:Add, Text, x20 y15 w650 h25, Select one Stamping Detail:
Gui, Remark4:Add, ListBox, x20 y45 w650 h450 vRemark4Choice Choose1, %Remark4Options%
Gui, Remark4:Add, Button, x270 y510 w150 h35 gRemark4OK Default, Select
Gui, Remark4:Show, w690 h565, Select Stamping Details

WinWaitClose, ahk_id %Remark4Hwnd%

if (Remark4Selected = "")
    return

WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 750

Send, %Remark4Selected%

Sleep, 750

Send, {Enter}

Sleep, 1500

Goto, ContinueAfterRemark4


Remark4OK:

Gui, Remark4:Submit

if (Remark4Choice = "")
{
    Gui, Remark4:Show
    return
}

Remark4Selected := Trim(Remark4Choice)

Gui, Remark4:Destroy

return



Remark4GuiClose:

Remark4Selected := ""

Gui, Remark4:Destroy

return


ContinueAfterRemark4:

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; save ;;;;;;;;;;;;;;;;;;;;
Sleep, 681

Send,{Tab}
Sleep, 200
Send, {Enter}
Sleep, 2000


; 21 Metal code

MouseMove, 687, 718 
Click, Left
Sleep, 566
Send ^n
Sleep, 1000
Send, %MetalCode%
Send, {Enter}
Sleep, 500
Send, {Enter}
Send, {Left 19}
Send, {Right 7}
Send, {Enter}
Sleep, 200

; READ METAL WEIGHT FROM EXCEL I2 (value after "=") AND SEND IT
try
{
    MetalWtText := ws.Range("I2").Value
    MetalWtText := Trim(MetalWtText)
}
catch
{
    MsgBox, 48, Error, Excel I2 read nahi ho saka.
    return
}

if !RegExMatch(MetalWtText, "=\s*([\d.]+)", MetalWtMatch)
{
    MsgBox, 48, Weight Error, Unable to identify the weight value from Excel I2.`n`n%MetalWtText%
    return
}

Send, %MetalWtMatch1%
Sleep, 200

Send, {Tab 2}
Sleep, 100
Send, {Enter}
Sleep, 100
Send, %MetalCode%
Sleep, 100
Send, {Enter}
Sleep, 500
Send, {Tab}
Sleep, 100
Send, {Enter}

; Search-select KT by typing it into the dropdown
if (KTSelected != "")
{
    Send, {F4}
    Sleep, 1000

    Loop, Parse, KTSelected
    {
        SendInput, %A_LoopField%
        Sleep, 300
    }

    Sleep, 700
    Send, {Enter}
    Sleep, 700
}

Send, {Tab}
Send, %ColorSelected%
Sleep, 2000

; 22. PART / COMPONENT DETAILS

ComponentList := []

AddComponents(ComponentText)
{
    global ComponentList

    ComponentText := Trim(ComponentText)

    if (ComponentText = "")
        return

    Loop, Parse, ComponentText, `;
    {
        Entry := Trim(A_LoopField)

        if (Entry = "")
            continue

        if InStr(Entry, ":")
        {
            StringSplit, Part, Entry, :

            ComponentName := Trim(Part1)
            ComponentQuantity := Trim(Part2)
        }
        else
        {
            ComponentName := Trim(Entry)
            ComponentQuantity := ""
        }

        if (ComponentName = "")
            continue

        Component := {}
        Component.Name := ComponentName
        Component.Quantity := ComponentQuantity

        ComponentList.Push(Component)
    }
}


try
{
    ComponentDetails := ws.Range("F2").Value
    ComponentDetails := Trim(ComponentDetails)
}
catch
{
    MsgBox, 48, Error, Unable to read Component Details from Excel F2.
    return
}

AddComponents(ComponentDetails)

try
{
    ExtraComponents := ws.Range("G2").Value
    ExtraComponents := Trim(ExtraComponents)
}
catch
{
    MsgBox, 48, Error, Unable to read Extra Components from Excel G2.
    return
}

AddComponents(ExtraComponents)

if (ComponentList.MaxIndex() = "")
{
    MsgBox, 48, Component Error, No component details were found in Excel F2 or G2.
    return
}

; KT is required for Part/Component Details
if (KTSelected = "")
{
    MsgBox, 48, KT Error, KT value is required for Part / Component Details.
    return
}


; Open Part Detail workbench (Ctrl+P only once)
Send, ^p
Sleep, 500

MouseMove, 2000, 719
Sleep, 200

Click, Left
Sleep, 300

; Process each component
ComponentIndex := 0

for Index, Component in ComponentList
{
    ComponentIndex++

    ComponentName := Trim(Component.Name)
    ComponentQuantity := Trim(Component.Quantity)

    ; Generate part detail code (A, B, C... then C26+)
    if (ComponentIndex <= 25)
    {
        CodeLetter := Chr(64 + ComponentIndex)
        PartDetailCode := "C" . CodeLetter
    }
    else
    {
        PartDetailCode := "C" . (ComponentIndex + 1)
    }

    ; New row per component
    Send, ^n
    Sleep, 1600

    ; Code, lowercase, character by character
    StringLower, PartDetailCodeLower, PartDetailCode

    Loop, Parse, PartDetailCodeLower
    {
        SendInput, %A_LoopField%
        Sleep, 150
    }

    Sleep, 300

    Send, {Enter}
    Sleep, 500

    ; Name
    Send, {Tab}
    Sleep, 1000

    StringLower, ComponentNameLower, ComponentName
    ComponentNameLower := Trim(ComponentNameLower)

    Send, %ComponentNameLower%

    Sleep, 1000

    Send, {Tab}
    Sleep, 1000


    ; Pieces
    if (ComponentQuantity != "")
    {
        Send, %ComponentQuantity%
        Sleep, 1000
    }

    Send, {Tab 3}
    Sleep, 500

    ; Metal - same code for every row
    if (MetalCode = "")
    {
        MsgBox, 48, Metal Error, Metal Code is missing.
        return
    }

    StringLower, MetalCodeLower, MetalCode

    Loop, Parse, MetalCodeLower
    {
        SendInput, %A_LoopField%
        Sleep, 150
    }

    Sleep, 300

    Send, {Tab}
    Sleep, 500

    ; KT - same code for every row. Search via F4 dropdown
    ; (typing straight into the cell landed on the combo's
    ; single-character jump, e.g. G03 instead of G14).
    if (KTSelected != "")
    {
        Send, {F4}
        Sleep, 1000

        Loop, Parse, KTSelected
        {
            SendInput, %A_LoopField%
            Sleep, 300
        }

        Sleep, 700

        Send, {Enter}
        Sleep, 700
    }

    Send, {Tab}
    Sleep, 500


    ; Color - same code for every row
    if (ColorSelected != "")
    {
        StringLower, ColorSelectedLower, ColorSelected

        Loop, Parse, ColorSelectedLower
        {
            SendInput, %A_LoopField%
            Sleep, 150
        }

        Sleep, 300
    }

    Send, {Enter}
    Sleep, 750
}

MouseMove, 623, 1110
Click, Left


; ============================================================
; PER-ROW STONE/SETTING LOOP
; Excel rows 3, 4, 5... each describe one stone/setting entry.
; Everything below reads that row's own K/L/M/N/O/P/T/U cells
; and stops at the first row with no Location in column K.
; ============================================================

StoneRow := 3

Loop
{
    ; Stop once this row has no Location - that means the data ran out.
    StoneRowLocation := ws.Range("K" . StoneRow).Value
    StoneRowLocation := Trim(StoneRowLocation)

    if (StoneRowLocation = "")
        break


; ========================================================
; 23 : Laboour Feild
; ========================================================

    

    Send ^n
    Sleep, 500
    Send, {Enter}
    Sleep, 100
    Send, {Enter}
    Sleep, 100
    Send, {Enter}
    Sleep, 100

    Send, {Tab 2}
    Sleep, 1000


; Shape options (`n-separated, same order as GEMI's dropdown)
ShapeOptions := "ASCT | Asscher Cut`n"
ShapeOptions .= "BRL | Barrel`n"
ShapeOptions .= "BUG | Baguette`n"
ShapeOptions .= "BULLET | Bullet Cut`n"
ShapeOptions .= "CADILLAC | Cadillac`n"
ShapeOptions .= "CBCU | Cabochon Cushion`n"
ShapeOptions .= "CBCURT | Cabochon Cushion Rectangle`n"
ShapeOptions .= "CBOVL | Cabochon Oval`n"
ShapeOptions .= "CBPRC | Cabochon Princess`n"
ShapeOptions .= "CBRD | Cabochon Round`n"
ShapeOptions .= "CHKBRCR | Checker Briolite Cushion Rectangle`n"
ShapeOptions .= "CHKBRCU | Checker Briolite Cushion`n"
ShapeOptions .= "CHKBRMRQ | Checker Briolite Marquise`n"
ShapeOptions .= "CHKBROVL | Checker Briolite Oval`n"
ShapeOptions .= "CHKBRPER | Checker Briolite Pear`n"
ShapeOptions .= "CHKBRRD | Checker Briolite Round`n"
ShapeOptions .= "CHKCU | Checker Cushion`n"
ShapeOptions .= "CHKCURT | Checker Cushion Rectangle`n"
ShapeOptions .= "CHKMRQ | Checker Marquise`n"
ShapeOptions .= "CHKOVL | Checker Oval`n"
ShapeOptions .= "CHKPER | Checker Pear`n"
ShapeOptions .= "CHKRD | Checker Round`n"
ShapeOptions .= "CHMPGE-FCT | Champagne Diamond Round Cut`n"
ShapeOptions .= "CURT | Cushion Rectangle`n"
ShapeOptions .= "CUSH | Cushion Cut`n"
ShapeOptions .= "DUTCH | Dutch Cut`n"
ShapeOptions .= "ELONGATED-CUSH | Elongated Cushion`n"
ShapeOptions .= "ELONGATED-HEXA | Elongated Hexagon`n"
ShapeOptions .= "ELONGATED-TRIGL | Elongated Triangle`n"
ShapeOptions .= "EMC | Emerald Cut`n"
ShapeOptions .= "EPAULLETE | Epaullete`n"
ShapeOptions .= "FCT | Round Full Cut`n"
ShapeOptions .= "HERT | Heart`n"
ShapeOptions .= "HEXA | Hexagon`n"
ShapeOptions .= "HFMN | Half Moon`n"
ShapeOptions .= "KITE | Kite`n"
ShapeOptions .= "LLYPRD | Lollipop Round`n"
ShapeOptions .= "LOZENGE | Lozenge`n"
ShapeOptions .= "MOVAL | Moval`n"
ShapeOptions .= "MRQ | Marquise`n"
ShapeOptions .= "OTG | Octagon`n"
ShapeOptions .= "OVL | Oval Cut`n"
ShapeOptions .= "PENTA | Pentagon`n"
ShapeOptions .= "PER | Pear Cut`n"
ShapeOptions .= "PRC | Princess Cut`n"
ShapeOptions .= "RA-ASCT | Royal Asscher Cut`n"
ShapeOptions .= "RA-CUSH | Royal Cushion Cut`n"
ShapeOptions .= "RA-FCT | Royal Full Cut`n"
ShapeOptions .= "RA-OVL | Royal Oval Cut`n"
ShapeOptions .= "RA-PER | Royal Pear Cut`n"
ShapeOptions .= "RDNT | Radiant Cut`n"
ShapeOptions .= "RRA-ASCT | Royal Rose Asscher Cut`n"
ShapeOptions .= "RRA-FCT | Royal Rose Full Cut`n"
ShapeOptions .= "RRA-OVL | Royal Rose Oval Cut`n"
ShapeOptions .= "RRA-PER | Royal Rose Pear Cut`n"
ShapeOptions .= "SCT | Single Cut`n"
ShapeOptions .= "SHIELD-CT | Shield Cut`n"
ShapeOptions .= "TAP | Tapper`n"
ShapeOptions .= "TRIGL | Triangle`n"
ShapeOptions .= "TRILL | Trillion`n"
ShapeOptions .= "TRPZ | Trapezoid"

; Shape from Excel L: Round/Oval map to a fixed GEMI code, no
; popup needed. Other shapes pre-filter the popup with the
; first word (lower-cased) - one match auto-fills, else the
; filtered list is shown for the user/client to pick.

try
{
    ShapeRowValue := ws.Range("L" . StoneRow).Value
    ShapeRowValue := Trim(ShapeRowValue)
}
catch
{
    MsgBox, 48, Error, Unable to read Excel L%StoneRow%.
    return
}

if (ShapeRowValue = "Round")
    ShapeSelected := "FCT"
else if (ShapeRowValue = "Oval")
    ShapeSelected := "OVL"
else
{
    RegExMatch(ShapeRowValue, "^[A-Za-z0-9]+", ShapeSearchWord)
    if (ShapeSearchWord != "")
        ShapeSearch := ShapeSearchWord
    else
        ShapeSearch := ShapeRowValue

    StringLower, ShapeSearch, ShapeSearch

    ShapeSelected := ShowSelectPopup("Shape", ShapeOptions, ShapeSearch)
}


; Type the picked Shape into GEMI's Shape cell
WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 1250

if (ShapeSelected != "")
{
    Send, {F4}
    Sleep, 500

    ; F4 may not have opened the dropdown (focus wasn't on the
    ; cell) - dismiss GEMI's "mandatory field" popup and retry.
    IfWinExist, ahk_class #32770 ahk_exe JwelSolution.exe
    {
        WinActivate
        Send, {Enter}
        Sleep, 500
        WinActivate, ahk_exe JwelSolution.exe
        WinWaitActive, ahk_exe JwelSolution.exe,, 10
        Sleep, 300
        Send, {F4}
        Sleep, 500
    }

    StringLower, ShapeSelected, ShapeSelected
    SendRaw, %ShapeSelected%
    Sleep, 300

    ; Tab commits it (Enter would trigger GEMI's row-validate shortcut)
    Send, {Tab}
    Sleep, 500
}


; Quality options (`n-separated, same order as GEMI's dropdown)

QualityOptions := "A+ | A+`n"
QualityOptions .= "A1 | A1`n"
QualityOptions .= "A2 | A2`n"
QualityOptions .= "BE-LG | BE-LG`n"
QualityOptions .= "BK-GH/SI | BK-GH/SI ( SDCC )`n"
QualityOptions .= "BK-SI | BK-SI`n"
QualityOptions .= "BKP2 | BKP2`n"
QualityOptions .= "BLP2 | BLP2`n"
QualityOptions .= "BR-Q7 | BR-Q7`n"
QualityOptions .= "BR-SI | BR-SI`n"
QualityOptions .= "BR-W12C | BR-W12C`n"
QualityOptions .= "BRP2 | BRP2`n"
QualityOptions .= "C1/C2-VS-SI | C1/C2-VS-SI ( Hening Diam )`n"
QualityOptions .= "C3/C6-VS-SI | C3/C6-VS-SI ( Hening Diam )`n"
QualityOptions .= "CSQ1 | CSQ1`n"
QualityOptions .= "CSQ2 | CSQ2`n"
QualityOptions .= "CSQ2A | CSQ2A`n"
QualityOptions .= "CSQ3 | CSQ3`n"
QualityOptions .= "CSQ4 | CSQ4`n"
QualityOptions .= "CSQ5 | CSQ5`n"
QualityOptions .= "CZ | CZ`n"
QualityOptions .= "CZ-BK-SDC | CZ-BK-SDC`n"
QualityOptions .= "CZ-CHM-SDC | CZ-CHM-SDC`n"
QualityOptions .= "CZ-SDC | CZ-SDC`n"
QualityOptions .= "D2 | D2`n"
QualityOptions .= "DB-HI/SI-DC | DB-HI/SI-DC`n"
QualityOptions .= "DB-HI/SI-JC | DB-HI/SI-JC`n"
QualityOptions .= "DB/HI-SI | DB/HI-SI`n"
QualityOptions .= "DB/I1-I2 | DB/I1-I2`n"
QualityOptions .= "DF/VVS2 | DF/VVS2`n"
QualityOptions .= "DPRP2 | DPRP2 ( SDC-RC )`n"
QualityOptions .= "EF/I1-I2 | EF/I1-I2`n"
QualityOptions .= "EF/VS | EF/VS`n"
QualityOptions .= "F-VS2 | F-VS2`n"
QualityOptions .= "F-VVS1 | F-VVS1 ( TJG )`n"
QualityOptions .= "F-W1 | F/SI1-SI2`n"
QualityOptions .= "F/SI-2 | F/SI-2 ( TJG )`n"
QualityOptions .= "FC74W2 | FC74W2- CARO ONLY`n"
QualityOptions .= "FM | FM`n"
QualityOptions .= "G2 | G2`n"
QualityOptions .= "GD-I3 | GD-I3 ( Grey Dia )`n"
QualityOptions .= "GH/SI1-SI2 | GH/SI1-SI2 ( Hening Dia )`n"
QualityOptions .= "GH/VS | GH/VS`n"
QualityOptions .= "GIA | GIA`n"
QualityOptions .= "GIA-GH/VS | GIA-GH/VS`n"
QualityOptions .= "GIA-HI/SI | GIA-HI/SI`n"
QualityOptions .= "GIA-LG-DEF/VVS1-2 | GIA-LG-DEF/VVS1-2 ( LAB GROWN )`n"
QualityOptions .= "GIA-LG-G+/VS | GIA-LG-G+/VS ( LAB GROWN )`n"
QualityOptions .= "H2 | H2 ( H/SI2 )`n"
QualityOptions .= "HI/1 | HI/1`n"
QualityOptions .= "HI/1-I2 | HI/1-I2`n"
QualityOptions .= "HI/I2 | HI/I2`n"
QualityOptions .= "HI/I2-I3 | HI/I2-I3`n"
QualityOptions .= "HI/I2-L | HI/I2-L`n"
QualityOptions .= "HI/SI | HI/SI`n"
QualityOptions .= "HI/SI-BR | HI/SI-BR`n"
QualityOptions .= "HI/SI-JC | HI/SI-JC`n"
QualityOptions .= "HI/VS | HI/VS`n"
QualityOptions .= "I1 | I1`n"
QualityOptions .= "I2 | I2`n"
QualityOptions .= "I2-I3 | I2-I3`n"
QualityOptions .= "ID-SI | ID-SI ( IDEAL CUT )`n"
QualityOptions .= "IGI-HI/SI | IGI-HI/SI`n"
QualityOptions .= "IGI-LG-F+/VS | IGI-LG-F+/VS`n"
QualityOptions .= "JK/I1 | JK/I1`n"
QualityOptions .= "JK/I1-I2 | JK/I1-I2`n"
QualityOptions .= "JK/I2 | JK/I2`n"
QualityOptions .= "JK/VS-SI | JK/VS-SI`n"
QualityOptions .= "K-VVS2 | K-VVS2`n"
QualityOptions .= "LG-D/VVS2 | LG-D/VVS2`n"
QualityOptions .= "LG-DE/VVS | FOR SDCC ONLY`n"
QualityOptions .= "LG-DEF/VVS1-2 | LG-DEF/VVS1-2 ( LAB GROWN )`n"
QualityOptions .= "LG-DG/VS | LG-DG/VS ( LAB GROWN )`n"
QualityOptions .= "LG-E+/VVS+ | LG-E+/VVS+`n"
QualityOptions .= "LG-E/VVS2 | FOR SDCC ONLY`n"
QualityOptions .= "LG-F+/VS | LG-F+/VS`n"
QualityOptions .= "LG-F/SI1-2 | LG-F/SI1-2 ( TJG )`n"
QualityOptions .= "LG-F/VVS | LG-F/VVS`n"
QualityOptions .= "LG-G+/VS | LG-G+/VS`n"
QualityOptions .= "LG-GH/SI | LG-GH/SI ( LAB GROWN )`n"
QualityOptions .= "LG-GHI/SI-I1 | LG-GHI/SI-I1 ( LAB GROWN QLT I+-I1+ )`n"
QualityOptions .= "LG-VS/SI | FOR TJG ONLY`n"
QualityOptions .= "LG-VS1 | FOR SDCC ONLY`n"
QualityOptions .= "LG-VS2 | FOR SDCC ONLY`n"
QualityOptions .= "LG-VVS2 | FOR SDCC ONLY`n"
QualityOptions .= "LM/I1 | LM/I1`n"
QualityOptions .= "LM/I1-I2 | LM/I1-I2`n"
QualityOptions .= "LM/I2 | LM/I2`n"
QualityOptions .= "MIX | MIX`n"
QualityOptions .= "N/A | N/A`n"
QualityOptions .= "N4 | N4`n"
QualityOptions .= "N5+ | N5+`n"
QualityOptions .= "PDQ5 | PDQ5`n"
QualityOptions .= "PK-TRT | Pink Treated Dia`n"
QualityOptions .= "PROMO | Promo`n"
QualityOptions .= "PROMO-HI/1 | PROMO-HI/1`n"
QualityOptions .= "PROMO-HI/1-I2 | PROMO-HI/1-I2`n"
QualityOptions .= "PROMO-HI/SI | PROMO-HI/SI`n"
QualityOptions .= "PROMO-HI/SI-I1 | PROMO-HI/SI-I1`n"
QualityOptions .= "PROMO-I2 | PROMO-I2`n"
QualityOptions .= "PROMO-I3 | PROMO-I3`n"
QualityOptions .= "PW3 | Pw3`n"
QualityOptions .= "Q5 | Q5`n"
QualityOptions .= "Q5-GL | Q5-GL`n"
QualityOptions .= "Q5-IDL | Q5-IDL ( IDEAL CUT )`n"
QualityOptions .= "Q6 | Q6`n"
QualityOptions .= "Q6+ | Q6+`n"
QualityOptions .= "Q7 | Q7`n"
QualityOptions .= "Q7+ | Q7+`n"
QualityOptions .= "Q8 | Q8`n"
QualityOptions .= "Q9 | Q9`n"
QualityOptions .= "SC-23 | SC-23`n"
QualityOptions .= "SC-S1 | SC-S1`n"
QualityOptions .= "SC12 | SC12`n"
QualityOptions .= "SC2 | SC2`n"
QualityOptions .= "SC3 | SC3`n"
QualityOptions .= "SC4 | SC4`n"
QualityOptions .= "SC5 | SC5`n"
QualityOptions .= "SC6 | SC6`n"
QualityOptions .= "SCN4 | SCN4`n"
QualityOptions .= "SD-A+ | SD-A+`n"
QualityOptions .= "SD-BKI3 | SD-BKI3`n"
QualityOptions .= "SD-DRJA | SD-DRJA`n"
QualityOptions .= "SD-I/I2-I3 | SD-I/I2-I3`n"
QualityOptions .= "SD-I1DL | I1-IDEAL CUT`n"
QualityOptions .= "SD-II2- | SD-II2-`n"
QualityOptions .= "SD-II3 | SD-II3`n"
QualityOptions .= "SD-IRJA | SD-IRJA`n"
QualityOptions .= "SD-N4B | SD-N4B`n"
QualityOptions .= "SD-NRJA | SD-NRJA`n"
QualityOptions .= "SD-RJB | SD-RJB`n"
QualityOptions .= "SD5 | SD5`n"
QualityOptions .= "SD6 | SD6`n"
QualityOptions .= "SD7 | SD7`n"
QualityOptions .= "SD8 | SD8`n"
QualityOptions .= "SDC-Q6 | SDC-Q6`n"
QualityOptions .= "SDC-W2 | SDC-W2`n"
QualityOptions .= "SDW2 | SDW2`n"
QualityOptions .= "STD | Standerd`n"
QualityOptions .= "STD-JM | STD-JM ( JIT )`n"
QualityOptions .= "STD-RC | STD-RC ( JIT )`n"
QualityOptions .= "TLB-3A | TLB-3A`n"
QualityOptions .= "TLC-4A | TLC-4A`n"
QualityOptions .= "TTLB-I1 | TTLB-I1`n"
QualityOptions .= "TTLB-SI1 | TTLB-SI1`n"
QualityOptions .= "VS-SI | VS-SI ( YASH BHAI )`n"
QualityOptions .= "W1 | W1`n"
QualityOptions .= "W1-2 | W1-2`n"
QualityOptions .= "W2 | W2`n"
QualityOptions .= "W2-3 | W2-3`n"
QualityOptions .= "W3 | W3`n"
QualityOptions .= "W4 | W4`n"
QualityOptions .= "W6 | FCW-6A ( SDCC )`n"
QualityOptions .= "W6N | W6N"

QualitySelected := ShowSelectPopup("Quality", QualityOptions)


; Type the picked Quality into GEMI's Quality cell
WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 500

if (QualitySelected != "")
{
    Send, {F4}
    Sleep, 500

    ; F4 may not have opened the dropdown (focus wasn't on the
    ; cell) - dismiss GEMI's "mandatory field" popup and retry.
    IfWinExist, ahk_class #32770 ahk_exe JwelSolution.exe
    {
        WinActivate
        Send, {Enter}
        Sleep, 500
        WinActivate, ahk_exe JwelSolution.exe
        WinWaitActive, ahk_exe JwelSolution.exe,, 10
        Sleep, 300
        Send, {F4}
        Sleep, 500
    }

    StringLower, QualitySelected, QualitySelected
    SendRaw, %QualitySelected%
    Sleep, 300

    ; Tab commits it (Enter would trigger GEMI's row-validate shortcut)
}

; Size select
Sleep, 100
Send, {Tab}
Sleep, 200
Send, {Enter}
Sleep, 300
Send, {Tab 3}

; Dismiss GEMI's row-validate popup if the Enter above triggered it
IfWinExist, ahk_class #32770 ahk_exe JwelSolution.exe
{
    WinActivate
    Send, {Enter}
    Sleep, 500
    WinActivate, ahk_exe JwelSolution.exe
    WinWaitActive, ahk_exe JwelSolution.exe,, 10
    Sleep, 300
}




ShapeCheckValue := ""
SizeValue := ""

try
{
    ; Shape (Round/Oval/etc.) comes from L for this row - check
    ; that to decide where the actual size value comes from.
    ShapeCheckValue := ws.Range("L" . StoneRow).Value
    ShapeCheckValue := Trim(ShapeCheckValue)

    if (ShapeCheckValue = "Round")
        SizeValue := ws.Range("N" . StoneRow).Value
    else if (ShapeCheckValue = "Oval")
        SizeValue := ws.Range("M" . StoneRow).Value
    else
        SizeValue := ws.Range("M" . StoneRow).Value

    SizeValue := Trim(SizeValue)
}
catch
{
    MsgBox, 48, Error, Unable to read Excel L%StoneRow%/M%StoneRow%/N%StoneRow%.
    return
}

if (SizeValue != "")
{
    StringLower, SizeValue, SizeValue
    SendRaw, %SizeValue%
}


Sleep, 500
Send, {Tab}
Sleep, 100
Send, {Enter}
Sleep, 300

; Same dismiss check as above - clear GEMI's row-validate popup
; if this Enter triggered it, before moving on to Pieces.
IfWinExist, ahk_class #32770 ahk_exe JwelSolution.exe
{
    WinActivate
    Send, {Enter}
    Sleep, 500
    WinActivate, ahk_exe JwelSolution.exe
    WinWaitActive, ahk_exe JwelSolution.exe,, 10
    Sleep, 300
}

;;;;;;;;;;;;;;; Pieces tab
Send, {Tab 2}
Sleep, 200


; ==========================================
; PIECES FROM EXCEL P
; ==========================================

try
{
    PiecesValue := ws.Range("P" . StoneRow).Value
    PiecesValue := Trim(PiecesValue)
}
catch
{
    MsgBox, 48, Error, Unable to read Excel P%StoneRow%.
    return
}

if (PiecesValue != "")
{
    SendRaw, %PiecesValue%
    Sleep, 300
}
Send, {Enter}
Sleep, 300

;;; WEIGHT

Send, {Tab}
Sleep, 200


; ==========================================
; WEIGHT FROM EXCEL O
; ==========================================

try
{
    WeightValue := ws.Range("O" . StoneRow).Value
    WeightValue := Trim(WeightValue)
}
catch
{
    MsgBox, 48, Error, Unable to read Excel O%StoneRow%.
    return
}

if (WeightValue != "")
    SendRaw, %WeightValue%

Sleep, 200
Send, {Enter}
Sleep, 200

;;;;;;;;;;; Setiing MetalType
Send, {Tab 2}
Sleep, 200


; Setting options (`n-separated, same order as GEMI's dropdown)

SettingOptions := "Articulate | Articulate`n"
SettingOptions .= "Bezel | Bezel Set`n"
SettingOptions .= "Cent Channel Prg | Cent Channel Prong`n"
SettingOptions .= "Cent Pave | Cent Pave`n"
SettingOptions .= "Cent Shared Prong | Center Shared Prong`n"
SettingOptions .= "Cent- Channel | Center Channel Set`n"
SettingOptions .= "Cent- Claw Prong | Center Claw Prong Set`n"
SettingOptions .= "Cent- Double Claw Prong | Center Double Claw Prong`n"
SettingOptions .= "Cent- Flush | Center- Flush`n"
SettingOptions .= "Cent- Prong | Center Prong`n"
SettingOptions .= "Cent- Prong ( SDC-H ) | Center Prong Set ( SDC-H )`n"
SettingOptions .= "Cent- Single Claw Prong | Center Single Claw Prong`n"
SettingOptions .= "Cent-Bezel | Cent-Bezel`n"
SettingOptions .= "Cent-French Pave | Center French Pave`n"
SettingOptions .= "Cent-Kite Prong | Cent-Kite Prong`n"
SettingOptions .= "Cent-M-Nick Set | Center Miracle Nick Set`n"
SettingOptions .= "Cent-M-Plate Prg | Centre Miracle Plate Prong`n"
SettingOptions .= "Cent-Mcr Spt Prg | Center Micro split prong`n"
SettingOptions .= "Cent-Partial Bezel | Cent-Partial Bezel`n"
SettingOptions .= "Cent-Pressure | Cent-Pressure Set`n"
SettingOptions .= "Cent-Talon Claw Prong | Cent-Talon Claw Prong`n"
SettingOptions .= "Cent-Triangular Prg | Center-Triangular Prong`n"
SettingOptions .= "Cent-TriangularSpPrg | Center Triangular Split Prong`n"
SettingOptions .= "Cent-V-Prong | Center V Prong`n"
SettingOptions .= "Cent-V-Prong-8 | Center V Prong Eight`n"
SettingOptions .= "Channel | Channel Set`n"
SettingOptions .= "Channel Prong | Channel Prong`n"
SettingOptions .= "Claw Prong | Claw Prong`n"
SettingOptions .= "Cut Down | Cut Down`n"
SettingOptions .= "Dancing Prong | Dancing Prong`n"
SettingOptions .= "Double Drill | Double Drill`n"
SettingOptions .= "Drill Bit | Drill Bits`n"
SettingOptions .= "Eternal Set | Eternal Set`n"
SettingOptions .= "Fan Prong | Fan Prong`n"
SettingOptions .= "Fancy Bezel | Fancy Bezel`n"
SettingOptions .= "Fancy Cent Bezel | Fancy Cent Bezel`n"
SettingOptions .= "Fancy Cent Channel | Fancy Cent Channel`n"
SettingOptions .= "Fancy Cent Channel Prg | Fancy Center Channel Prong`n"
SettingOptions .= "Fancy Cent Claw-V Prong | Fancy Center Claw V Prong Set`n"
SettingOptions .= "Fancy Cent Double Claw-Heart Prong | Fancy Cent Double Claw-Heart Prong`n"
SettingOptions .= "Fancy Cent Flush | Fancy Center Flush`n"
SettingOptions .= "Fancy Cent Kite Prg | Fancy Cent Kite Prg`n"
SettingOptions .= "Fancy Cent M-Nick Set | Fancy Cent Miracle Nick Set`n"
SettingOptions .= "Fancy Cent Mcr Spt Prg | Fancy Center Micro split prong`n"
SettingOptions .= "Fancy Cent Shared Prong | Fancy Center Shared Prong`n"
SettingOptions .= "Fancy Cent Split Prong | Fancy Cent Split Prong`n"
SettingOptions .= "Fancy Cent Talon Claws prong | Fancy Cent Talon Claws prong`n"
SettingOptions .= "Fancy Cent V-Prong | Fancy Cent V-Prong`n"
SettingOptions .= "Fancy Cent V-Prong-8 | Fancy Cent V-Prong Eight`n"
SettingOptions .= "Fancy Cent- Claw Prong | Fancy Center Claw Prong Set`n"
SettingOptions .= "Fancy Cent- Double Claw Prong | Fancy Center Double Claw Prong`n"
SettingOptions .= "Fancy Cent- Double Prong | Fancy Center Double Prong`n"
SettingOptions .= "Fancy Cent- Single Claw Prong | Fancy Center Single Claw Prong`n"
SettingOptions .= "Fancy Cent-Pressure | Fancy Cent-Pressure Set`n"
SettingOptions .= "Fancy Cent-Triangular Prg | Fancy Center-Triangular Prong`n"
SettingOptions .= "Fancy CentPrg | FancyCenterProng`n"
SettingOptions .= "Fancy Channel | Fancy Channel`n"
SettingOptions .= "Fancy Channel Prg | FancyChannelProng`n"
SettingOptions .= "Fancy Claw Prong | Fancy Claw Prong`n"
SettingOptions .= "Fancy Flush | FancyFlush`n"
SettingOptions .= "Fancy Inv | Fancy Invisible`n"
SettingOptions .= "Fancy Nick Set | Fancy Nick Set`n"
SettingOptions .= "Fancy Pave | Fancy Pave`n"
SettingOptions .= "Fancy Pressure | Fancy Pressure`n"
SettingOptions .= "Fancy Prong | Fancy Prong`n"
SettingOptions .= "Fancy Shared Prong | Fancy Shared Prong`n"
SettingOptions .= "Fancy Split Prong | Fancy Split Prong`n"
SettingOptions .= "Fancy Stick Cent Prg | Fancy Stick Centre Prong`n"
SettingOptions .= "Fancy Talon Claw Prong | Fancy Talon Claw Prong`n"
SettingOptions .= "Fancy V-Prong | Fancy V-Prong`n"
SettingOptions .= "FancyTriangular Prong | FancyTriangular Prong`n"
SettingOptions .= "Flush | Flush Set`n"
SettingOptions .= "French Pave | French Pave`n"
SettingOptions .= "Full Drill | Full Drill`n"
SettingOptions .= "Half Drill | Half Drill`n"
SettingOptions .= "Invisible | Invisible`n"
SettingOptions .= "JTCLK | Just Click`n"
SettingOptions .= "L Drill | L Drill`n"
SettingOptions .= "Micro Pave | Micro Pave`n"
SettingOptions .= "Mini Prong | Mini Prong Set`n"
SettingOptions .= "Miracle Plate | Miracle Plate`n"
SettingOptions .= "Nick Set | Nick Set`n"
SettingOptions .= "Nuovo | Nuovo`n"
SettingOptions .= "P-Bezel | Platinum Bezel`n"
SettingOptions .= "P-Cent- Channel | PLT-Center Channel`n"
SettingOptions .= "P-Cent- Claw Prong | PLT-Center- Claw Prong`n"
SettingOptions .= "P-Cent- Double Claw Prong | PLT- Cent- Double Claw prong`n"
SettingOptions .= "P-Cent- Prong | PLT-Cent- Prong`n"
SettingOptions .= "P-Cent- Shared Prong | PLT-Cent-Shared Prong`n"
SettingOptions .= "P-Cent-Bezel | Platinum Cent-Bezel`n"
SettingOptions .= "P-Cent-Mcr Spt Prong | PLT-Cent Mcr Splt Prong`n"
SettingOptions .= "P-Channel | PLT Channel`n"
SettingOptions .= "P-Channel Prong | PLT Channel Prong`n"
SettingOptions .= "P-Claw Prong | PLT Claw Prong`n"
SettingOptions .= "P-Cut Down | P-Cut Down`n"
SettingOptions .= "P-Fancy Bezel | PLT Fancy Bezel`n"
SettingOptions .= "P-Fancy Cent Bezel | PLT- Fancy Cent Bezel`n"
SettingOptions .= "P-Fancy Cent Channel | PLT-Fancy Cent Channel`n"
SettingOptions .= "P-Fancy Cent Mcr Spt Prong | PLT-Fancy Cent- Claw Prong`n"
SettingOptions .= "P-Fancy Cent- Double Claw Prong | P-Fancy Cent Double Claw Prong`n"
SettingOptions .= "P-Fancy Cent- Shared Prong | PLT-Fancy Cent- Shared Prong`n"
SettingOptions .= "P-Flush | P-Flush`n"
SettingOptions .= "P-French Pave | PLT French Pave`n"
SettingOptions .= "P-Invisible | PLT-Invisible`n"
SettingOptions .= "P-Micro Pave | PLT Micro Pave`n"
SettingOptions .= "P-Micro Prong | PLT Micro Prong`n"
SettingOptions .= "P-Nuovo | P-Nuovo`n"
SettingOptions .= "P-Pave | PLT Pave`n"
SettingOptions .= "P-Plate Prong | Platinum Plate Prong`n"
SettingOptions .= "P-Prong | PLT Prong`n"
SettingOptions .= "P-Shared Prong | PLT Shared Prong`n"
SettingOptions .= "P-Split Prong | PLT-Split Prong`n"
SettingOptions .= "P-Swallow Tail | PLT-Swallow Tail`n"
SettingOptions .= "P-TriangularSpPr | PLT Triangular Split Prong`n"
SettingOptions .= "P-V-Prong | Platinum V- Prong`n"
SettingOptions .= "Partial Bezel | Partial Bezel`n"
SettingOptions .= "Pave | Pave Set`n"
SettingOptions .= "Pin Point | Pin Point`n"
SettingOptions .= "Plate Prong | Plate Prong`n"
SettingOptions .= "Pressure | Pressure Set`n"
SettingOptions .= "Prong | Prong Set`n"
SettingOptions .= "Prong ( SDC-H ) | Prong Set ( SDC-H )`n"
SettingOptions .= "S-Bezel | Silver Bezel`n"
SettingOptions .= "S-Cent Channel | S-Cent Channel`n"
SettingOptions .= "S-Cent Channel Prong | Silver Cent Channel Prong`n"
SettingOptions .= "S-Cent Fan Prong | Silver Center Fan Prong`n"
SettingOptions .= "S-Cent- Claw Prong | S-Cent- Claw Prong`n"
SettingOptions .= "S-Cent- Pave | Silver Cent Pave`n"
SettingOptions .= "S-Cent- Prong | Silver Cent- Prong`n"
SettingOptions .= "S-Cent-Bezel | S-Cent-Bezel`n"
SettingOptions .= "S-Cent-M-Plate Prg | Silver Centre Miracle Plate Prong`n"
SettingOptions .= "S-Cent-Mcr Spt Prg | Silver Center Micro Split Prong`n"
SettingOptions .= "S-Cent-Pressure | S-Cent Pressure Set`n"
SettingOptions .= "S-Cent-Talon Claw Prong | Silver Cent-Talon Claw Prong`n"
SettingOptions .= "S-Channel | Silver Channel`n"
SettingOptions .= "S-Channel Prong | Silver Channel Prong`n"
SettingOptions .= "S-Claw Prong | Silver-Claw Prong`n"
SettingOptions .= "S-Cut Down | Silver Cut Down`n"
SettingOptions .= "S-Fan Prong | Silver Fan Prong`n"
SettingOptions .= "S-Fancy Bezel | S-Fancy Bezel`n"
SettingOptions .= "S-Fancy Cent Bezel | S-Fancy Cent Bezel`n"
SettingOptions .= "S-Fancy Cent Channel | S-Fancy Cent Channel`n"
SettingOptions .= "S-Fancy Cent Channel Prg | S-Fancy Center Channel Prong`n"
SettingOptions .= "S-Fancy Cent Claw Prong | Silver Fancy Center Claw Prong`n"
SettingOptions .= "S-Fancy Cent Claw-V Prong | Silver Fancy Center Claw-V Prong`n"
SettingOptions .= "S-Fancy Cent Flush | Silver Fancy Cent Flush`n"
SettingOptions .= "S-Fancy Cent Prg | Silver Fancy CentProng`n"
SettingOptions .= "S-Fancy Cent Split Prong | Silver Fancy Centre Split Prong`n"
SettingOptions .= "S-Fancy Cent Talon Claws prong | Silver Fancy Cent Talon Claws prong`n"
SettingOptions .= "S-Fancy Channel | Silver Fancy Channel`n"
SettingOptions .= "S-Fancy Channel Prg | Silver Fancy Channel Prong`n"
SettingOptions .= "S-Fancy Claw Prong | Silver Fancy Claw Prong`n"
SettingOptions .= "S-Fancy Prong | Silver Fancy Prong`n"
SettingOptions .= "S-Fancy Split Prong | Silver Fancy Split Prong`n"
SettingOptions .= "S-Fancy Talon Claw Prong | Silver Fancy Talon Claw Prong`n"
SettingOptions .= "S-Flush | Silver Flush`n"
SettingOptions .= "S-INVISIBLE | Silver Invisible`n"
SettingOptions .= "S-Micro Pave | Silver Micro Pave`n"
SettingOptions .= "S-Miracle Plate | Silver Miracle Plate`n"
SettingOptions .= "S-Nick Set | Silver Nick Set`n"
SettingOptions .= "S-Pave | Silver Pave`n"
SettingOptions .= "S-Plate Prong | Silver Plate Prong`n"
SettingOptions .= "S-Pressure | Silver Pressure`n"
SettingOptions .= "S-Prong | Silver Prong`n"
SettingOptions .= "S-Split Prong | S-Split Prong`n"
SettingOptions .= "S-Stick | Silver Stick Set`n"
SettingOptions .= "S-Talon Claw Prong | Silver Talon Claw Prong`n"
SettingOptions .= "Shared Prong | Shared Prong`n"
SettingOptions .= "Single Drill | Single Drill`n"
SettingOptions .= "Split Prong | Split Prong`n"
SettingOptions .= "Split Prong - Grain Tool | Split Prong - Grain Tool`n"
SettingOptions .= "Stick | Stick Set`n"
SettingOptions .= "Stick Cent Prong | Stick Centre Prong`n"
SettingOptions .= "Stick Prong | Stick Prong`n"
SettingOptions .= "Swallow Tail | Swallow Tail`n"
SettingOptions .= "T Drill | T Drill`n"
SettingOptions .= "Talon Claw Prong | Talon Claw Prong`n"
SettingOptions .= "TriangularSpPr | Triangular Split Prong`n"
SettingOptions .= "V Drill | V Drill`n"
SettingOptions .= "V-Prong | V-Prong Set`n"
SettingOptions .= "V-Prong French Pave | V-Prong French Pave`n"
SettingOptions .= "Y Drill | Y Drill"


try
{
    SettingSearch := ws.Range("T" . StoneRow).Value
    SettingSearch := Trim(SettingSearch)
}
catch
{
    MsgBox, 48, Error, Unable to read Excel T%StoneRow%.
    return
}

; T3 word break is a space or hyphen - take the leading word either way
RegExMatch(SettingSearch, "^[A-Za-z0-9]+", SettingSearchWord)
if (SettingSearchWord != "")
    SettingSearch := SettingSearchWord

StringLower, SettingSearch, SettingSearch

SettingSelected := ShowSelectPopup("Setting", SettingOptions, SettingSearch)

; Type the picked Setting into GEMI's Setting cell
WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 300

if (SettingSelected != "")
{
    Send, {F4}
    Sleep, 500

    ; GEMI's search can't handle a typed space (eg. "Claw Prong")
    ; - it reopens the dropdown instead of continuing the search.
    ; Type up to (not including) the first space and stop there;
    ; a code with no space types in full and commits with Enter.
    ; A code with a space leaves the dropdown open, filtered to
    ; the first word, for the user to finish picking manually.
    StringLower, SettingSelectedLower, SettingSelected

    HitSpace := false

    Loop, Parse, SettingSelectedLower
    {
        if (A_LoopField = " ")
        {
            HitSpace := true
            break
        }

        SendInput, %A_LoopField%
        Sleep, 200
    }

    Sleep, 300

    if !HitSpace
        Send, {Enter}

    Sleep, 500
}

;;;    setting type

Send, {Tab}


; Setting Type from Excel U: only HS/NA/WS in GEMI, each starting
; with a different letter, so just type U's first letter + Enter.

try
{
    SettingTypeValue := ws.Range("U" . StoneRow).Value
    SettingTypeValue := Trim(SettingTypeValue)
}
catch
{
    MsgBox, 48, Error, Unable to read Excel U%StoneRow%.
    return
}

if (SettingTypeValue != "")
{
    SettingTypeLetter := SubStr(SettingTypeValue, 1, 1)
    StringLower, SettingTypeLetter, SettingTypeLetter

    SendInput, %SettingTypeLetter%
    Sleep, 300
    Send, {Enter}
    Sleep, 500
}



; Automatic Stone Size selection
Send, {Tab}
Sleep, 200
Send, {F4}
Sleep, 500

; Read Stone Size from the active Excel row
ShapeCheckValue := ""
SizeValue := ""

try
{
    ShapeCheckValue := ws.Range("L" . StoneRow).Value
    ShapeCheckValue := Trim(ShapeCheckValue)

    if (ShapeCheckValue = "Round")
        SizeValue := ws.Range("N" . StoneRow).Value
    else
        SizeValue := ws.Range("M" . StoneRow).Value

    SizeValue := Trim(SizeValue)
}
catch
{
    MsgBox, 48, Error,
    (
    Unable to read Stone Size from Excel.

    Row: %StoneRow%
    Shape: %ShapeCheckValue%

    Round = Column N
    Oval/Other = Column M
    )
    return
}


if (SizeValue != "")
{
    ; Strip spaces and "mm", take first 3 chars, lowercase, type
    StringReplace, SizeSearchValue, SizeValue, %A_Space%,, All
    StringReplace, SizeSearchValue, SizeSearchValue, mm,, All
    StringLeft, SizeSearch3, SizeSearchValue, 3
    StringLower, SizeSearch3, SizeSearch3

    SendInput, %SizeSearch3%
    Sleep, 800

    ; Stop here - do NOT send Enter/F4/Tab. User picks the exact result.
}
else
{
    MsgBox, 48, Stone Size,
    (
    Stone Size is empty in Excel.

    Row: %StoneRow%
    Shape: %ShapeCheckValue%
    )
}

Sleep, 300


; Brand
Send, {Tab 2}
Sleep, 200

BrandOptions := "CEN-ST | Center Stone`n"
BrandOptions .= "ENG | Engagement Ring`n"
BrandOptions .= "WED | Wedding Band"

BrandSelected := ShowSelectPopup("Brand", BrandOptions)

WinActivate, ahk_exe JwelSolution.exe
WinWaitActive, ahk_exe JwelSolution.exe,, 10

if ErrorLevel
    return

Sleep, 300


if (BrandSelected != "")
{
    ; Open GEMI Brand dropdown
    Send, {F4}
    Sleep, 500

    ; Type selected brand code
    StringLower, BrandSelected, BrandSelected
    SendRaw, %BrandSelected%

    Sleep, 300

    ; Commit selection
    Send, {Enter}
    Sleep, 500
}
    StoneRow++

}


;;;;;;;;;;;;;;;;;Labour charge
; ==========================================
; LABOUR CHARGE - DEFAULT STYL
; ==========================================

MouseMove, 600, 1338
Click, Left
Sleep, 300

Send, {F4}
Sleep, 300

; Select default option: STYL
Send, style
Sleep, 300

Send, {Enter}
Sleep, 500

Send, {Tab}
Sleep, 200

Send, {Enter}


; ==========================================
; PART DETAIL CODE
; ==========================================

; Normalize SubCategory
StringLower, CurrentSubCategory, ExcelSubCategory
CurrentSubCategory := Trim(CurrentSubCategory)


; ==========================================
; ENGAGEMENT RING
; ==========================================

if (CurrentSubCategory = "engagement ring")
{
    if (ComponentCount = 1)
    {
        PartDetailCode := "E-CFP-1"
    }
    else if (ComponentCount = 2 || ComponentCount = 3)
    {
        PartDetailCode := "E-CFP-2&3"
    }
    else if (ComponentCount = 4)
    {
        PartDetailCode := "E-CFP-4"
    }
    else if (ComponentCount = 5)
    {
        PartDetailCode := "E-CFP-5"
    }
    else if (ComponentCount = 6)
    {
        PartDetailCode := "E-CFP-6"
    }
    else if (ComponentCount = 7)
    {
        PartDetailCode := "E-CFP-7"
    }
    else if (ComponentCount = 8)
    {
        PartDetailCode := "E-CFP-8"
    }
    else if (ComponentCount = 9 || ComponentCount = 10)
    {
        PartDetailCode := "E-CFP-9&10"
    }
    else if (ComponentCount = 11 || ComponentCount = 12)
    {
        PartDetailCode := "E-CFP-11&12"
    }
    else if (ComponentCount >= 13 && ComponentCount <= 15)
    {
        PartDetailCode := "E-CFP-13 TO15"
    }
    else if (ComponentCount = 16)
    {
        PartDetailCode := "E-CFP-16"
    }
    else if (ComponentCount = 17)
    {
        PartDetailCode := "E-CFP-17"
    }
    else if (ComponentCount = 18)
    {
        PartDetailCode := "E-CFP-18"
    }
    else if (ComponentCount = 19)
    {
        PartDetailCode := "E-CFP-19"
    }
    else if (ComponentCount = 20)
    {
        PartDetailCode := "E-CFP-20"
    }
}


; ==========================================
; ALL OTHER SUB CATEGORIES
; ==========================================

else
{
    if (ComponentCount >= 1 && ComponentCount <= 20)
    {
        PartDetailCode := "CFP-" . ComponentCount
    }
}


; ==========================================
; SAFETY CHECK
; ==========================================

if (PartDetailCode = "")
{
    MsgBox, 48, Part Detail Code,
    (
    Unable to determine Part Detail Code.

    Sub Category:
    %ExcelSubCategory%

    Component Count:
    %ComponentCount%
    )

    return
}


; ==========================================
; SEND CODE TO GEMI
; ==========================================

StringLower, PartDetailCodeLower, PartDetailCode

Loop, Parse, PartDetailCodeLower
{
    SendInput, %A_LoopField%
    Sleep, 150
}

Sleep, 300





















Return


^R::
Reload