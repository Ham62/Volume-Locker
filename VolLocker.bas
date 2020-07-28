#define fbc -s gui res\VolLocker.rc

#include "windows.bi"
#include "win\commctrl.bi"
#include "win\shellapi.bi"
#include "win\mmsystem.bi"
#include "crt.bi"

#include "LockControl.bi"
#include "NewLockDialog.bas"

#define ShowDebug 0

#define _SysVolToPercent( _iVol ) cint((( _iVol ) * 100) / &HFFFF)

' In 4 years I've never figured out the exact formula to use to get these values
' to line up with the Windows mixer UI... this is the closest I could find.
' If you have a better formula please share it! :)
#define _PercentToSysVol( _iVol ) iif (( _iVol ) = 100, 65535, ((( _iVol )*656)+(( _iVol )\3)))

Declare Sub WinMain(hInstance as HINSTANCE, hPrevInstance as HINSTANCE, szCmdLine as PSTR, iCmdShow as Integer)
Declare Sub SaveSettings()
Declare Sub LoadSettings()
enum WindowControls
    wcMain
    
    wcTabControl  ' Tab control
    wcTabCloseBtn ' Button to close tab
    
    wcWelcomeLbl
        
    wcVolGrp
        
    wcVolumeLbl
    wcVolLevelLbl
    wcVolLockEnable
    wcVolumeSlider
    
    wcMuteGrp
    
    wcMuteLockEnabled
    wcMuteState
    
    wcVersionLbl
    
    wcLast
end enum

enum TrayMenuIDs
    IDM_SHOWHIDE
    IDM_SEPARATOR
    IDM_ABOUT
    IDM_EXIT
end enum

Const WINDOW_WIDTH = 400, WINDOW_HEIGHT = 265
Dim shared as NOTIFYICONDATA nidApp
Dim shared as hwnd CTL(wcLast)   'Control handles
Dim Shared as HINSTANCE hInstance
Dim Shared as HWND hWndMain, hWndNewLock
Dim Shared as String szAppName
Dim Shared as String szCaption
Dim Shared as String szIconTxt

Declare Function GetLineState(device as DeviceControl, dwControlID as DWORD) as DWORD
Declare Function OpenDevice(iDev as integer, device as DeviceControl) as integer
Declare Sub SetLineState(device as DeviceControl, dwControlID as DWORD, dwValue as DWORD)
Declare Sub CloseDevice(device as DeviceControl)
Declare Sub EnumInputDevices()

Dim shared as integer SelectedDevice = 0 ' 0 is default device

Dim shared as integer iTotalDevices
Redim shared as DeviceControl devices(0)

EnumInputDevices() 'Store device names in array

hInstance = GetModuleHandle(NULL)
szAppName = "Volume Locker"
szCaption = "Volume Locker"
szIconTxt = "Volume Locker"

'Launch into WinMain()
WinMain(hInstance, NULL, GetCommandLine(), SW_NORMAL)

Sub ShowControls(device as DeviceControl)
    ' Hide welcome message
    ShowWindow(CTL(wcWelcomeLbl), SW_HIDE)

    ' Show volume/mute groups
    ShowWindow(CTL(wcVolGrp), SW_SHOW)
    ShowWindow(CTL(wcMuteGrp), SW_SHOW)
    
    ' Update and enable volume controls
    Dim as integer iVolume
    if (device.iControlFlags AND LOCK_VOLUME_ENABLED) then
        iVolume = _SysVolToPercent(device.iLockVolume)
        EnableWindow(CTL(wcVolumeSlider), TRUE)
        EnableWindow(CTL(wcVolumeLbl), TRUE)
        EnableWindow(CTL(wcVolLevelLbl), TRUE)
    else
        iVolume = _SysVolToPercent(GetLineState(device, device.iVolControlID))
        EnableWindow(CTL(wcVolumeSlider), FALSE)
        EnableWindow(CTL(wcVolumeLbl), FALSE)
        EnableWindow(CTL(wcVolLevelLbl), FALSE)
    end if                
    SendMessage(CTL(wcVolumeSlider), TBM_SETPOS, TRUE, iVolume)
    SetWindowText(CTL(wcVolLevelLbl), Str(iVolume))
    SendMessage(CTL(wcVolLockEnable), BM_SETCHECK, _
                iif((device.iControlFlags AND LOCK_VOLUME_ENABLED), BST_CHECKED, BST_UNCHECKED), 0)
    
    ' Update mute controls
    Dim as integer iMuteState
    if (device.iControlFlags AND LOCK_MUTE_ENABLED) then
        iMuteState = device.iLockMute
        EnableWindow(CTL(wcMuteState), TRUE)
    else
        iMuteState = GetLineState(device, device.iMuteControlID)
        EnableWindow(CTL(wcMuteState), FALSE)
    end if
    SendMessage(CTL(wcMuteState), BM_SETCHECK, iif(iMuteState, BST_CHECKED, BST_UNCHECKED), 0)
    SendMessage(CTL(wcMuteLockEnabled), BM_SETCHECK, _
                iif((device.iControlFlags AND LOCK_MUTE_ENABLED), BST_CHECKED, BST_UNCHECKED), 0)
End Sub

Sub ShowStartPage()
    Dim as zstring ptr pszWelcomeMsg = _
        @(!"Welcome to Volume Locker V1.3!\r\n\r\n" + _
          !"You don't currently have any device locks configured.\r\n" + _
          !"Click on the '+' tab to select a new device to configure.\r\n\r\n" + _
          !"Volume Locker (C) Graham Downey 2017-2020")
         
    SetWindowText(CTL(wcWelcomeLbl), pszWelcomeMsg)
    ShowWindow(CTL(wcWelcomeLbl), SW_SHOW)
    
    ShowWindow(CTL(wcVolGrp), SW_HIDE)
    ShowWindow(CTL(wcMuteGrp), SW_HIDE)
    'ShowWindow(CTL(wcVersionLbl), SW_HIDE)
End Sub

' Allows subclassing of controls by sending messages back to parent window
Function BubbleUp(hWnd as HWND, iMsg as long, wParam as WPARAM, lParam as LPARAM) as LRESULT
    Select Case iMsg
    Case WM_COMMAND, WM_HSCROLL
        SendMessage(GetParent(hWnd), iMsg, wParam, lParam)
    End Select
    return CallWindowProc(cast(any ptr, GetWindowLong(hWnd, GWL_USERDATA)), hWnd, iMsg, wParam, lParam)
End Function

Function WndProc(hWnd as HWND, iMsg as uLong, wParam as WPARAM, lParam as LPARAM) as LRESULT
    static as HFONT fntDefault, fntSmall, fntTiny
    static as HMENU hPopupMenu
    static as TC_ITEM tcItem

    Select Case iMsg
    Case WM_CREATE
        
        '**** Center window on desktop ****'
        'Calculate Client Area Size
        Dim as rect RcWnd = any, RcCli = Any, RcDesk = any
        GetClientRect(hWnd, @RcCli)
        GetClientRect(GetDesktopWindow(), @RcDesk)
        GetWindowRect(hWnd, @RcWnd)
        'Window Rect is in SCREEN coordinate.... make right/bottom become WID/HEI
        with RcWnd
            .right -= .left: .bottom -= .top
            .right += (.right-RcCli.right)  'Add difference cli/wnd
            .bottom += (.bottom-RcCli.bottom)   'add difference cli/wnd
            var CenterX = (RcDesk.right-.right)\2
            var CenterY = (RcDesk.bottom-.bottom)\2
            SetWindowPos(hwnd,null,CenterX,CenterY,.right,.bottom,SWP_NOZORDER)
        end with

        InitCommonControls()

        '**** Create default fonts ****'
        var hDC = GetDC(hWnd)
        var nHeight = -MulDiv(10, GetDeviceCaps(hDC, LOGPIXELSY), 72) 'calculate size matching DPI
        var nSmallHeight = -MulDiv(8, GetDeviceCaps(hDC, LOGPIXELSY), 72) 'calculate size matching DPI
        var nTinyHeight = -MulDiv(7, GetDeviceCaps(hDC, LOGPIXELSY), 72) 'calculate size matching DPI
        
        fntDefault = CreateFont(nHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, _
                            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
                            DEFAULT_PITCH, "Verdana")

        fntSmall = CreateFont(nSmallHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, _
                            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
                            DEFAULT_PITCH, "Verdana")
                                
        fntTiny = CreateFont(nTinyHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, _
                            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
                            DEFAULT_PITCH, "Verdana")
        
        ' Macro for creating window controls
        #define CreateControl(mID , mExStyle , mClass , mCaption , mStyle , mX , mY , mWid , mHei) CTL(mID) = CreateWindowEx(mExStyle,mClass,mCaption,mStyle,mX,mY,mWid,mHei,hwnd,cast(hmenu,mID),hInstance,null)
        #define CreateTabControl(mID , mParent, mExStyle , mClass , mCaption , mStyle , mX , mY , mWid , mHei) CTL(mID) = CreateWindowEx(mExStyle,mClass,mCaption,mStyle,mX,mY,mWid,mHei,mParent,cast(hmenu,mID),hInstance,null)
        const cBase = WS_VISIBLE OR WS_CHILD
        const cTabControl = cBase OR WS_CLIPSIBLINGS OR TCS_OWNERDRAWFIXED'OR TCS_MULTILINE
        const cTabCloseBtn = cBase OR BS_OWNERDRAW
        const cGrpBox = cBase OR BS_GROUPBOX OR WS_CLIPSIBLINGS
        const cTrackBar = cBase OR TBS_TOP
        const cCheckBox = cBase OR BS_AUTOCHECKBOX
        const cCheckBoxR = cCheckBox OR BS_RIGHTBUTTON
        
        ' Device tab controls
        CreateControl(wcTabControl, NULL, WC_TABCONTROL, "", cTabControl, 0, 0, rcWnd.right, rcWnd.bottom)
        CreateControl(wcTabCloseBtn, NULL, WC_BUTTON, "x", cTabCloseBtn, 0, 0, 14, 14)

        ' Welcome message label
        CreateTabControl(wcWelcomeLbl, CTL(wcTabControl), NULL, WC_STATIC, "", cBase, 10, 30, 170, 20)

        ' **** Volume Control Group ****
        CreateTabControl(wcVolGrp, CTL(wcTabControl), WS_EX_TRANSPARENT, WC_BUTTON, "Volume", cGrpBox, 5, 5, WINDOW_WIDTH-10, WINDOW_HEIGHT-10)

        CreateTabControl(wcVolumeLbl, CTL(wcVolGrp), WS_EX_TRANSPARENT, WC_STATIC, "Volume:", cBase, 19, 110, 170, 20)
        CreateTabControl(wcVolLevelLbl, CTL(wcVolGrp), WS_EX_TRANSPARENT, WC_STATIC, "69", cBase, 195, 110, 30, 20)
        CreateTabControl(wcVolLockEnable, CTL(wcVolGrp), WS_EX_TRANSPARENT, WC_BUTTON, "Enable Volume Lock", cCheckBox, 19, 180, 160, 20)
        CreateTabControl(wcVolumeSlider, CTL(wcVolGrp), NULL, TRACKBAR_CLASS, "", cTrackBar, 10, 130, 380, 30)

        SendMessage(ctl(wcVolumeSlider), TBM_SETRANGE, TRUE, MAKELONG(0, 100))  'Set slider range
        SendMessage(ctl(wcVolumeSlider), TBM_SETPAGESIZE, 0, 4)                 'Set page size
        SendMessage(ctl(wcVolumeSlider), TBM_SETSEL, FALSE, MAKELONG(0, 1))    'Set selection range
        
        
        ' **** Mute Control Group ****
        CreateTabControl(wcMuteGrp, CTL(wcTabControl), WS_EX_TRANSPARENT, WC_BUTTON, "Mute", cGrpBox, 5, 5, WINDOW_WIDTH-10, WINDOW_HEIGHT-10)

        CreateTabControl(wcMuteLockEnabled, CTL(wcMuteGrp), WS_EX_TRANSPARENT, WC_BUTTON, "Enable Mute Lock", cCheckBox, 19, 180, 160, 20)
        CreateTabControl(wcMuteState, CTL(wcMuteGrp), WS_EX_TRANSPARENT, WC_BUTTON, "Device Muted:", cCheckBoxR, 19, 180, 130, 20)

        SetWindowLong(CTL(wcTabControl), GWL_USERDATA, SetWindowLong(CTL(wcTabControl), GWL_WNDPROC, cast(LONG, @BubbleUp)))
        SetWindowLong(CTL(wcVolGrp), GWL_USERDATA, SetWindowLong(CTL(wcVolGrp), GWL_WNDPROC, cast(LONG, @BubbleUp)))
        SetWindowLong(CTL(wcMuteGrp), GWL_USERDATA, SetWindowLong(CTL(wcMuteGrp), GWL_WNDPROC, cast(LONG, @BubbleUp)))

        '**** Version string ****'
        CreateTabControl(wcVersionLbl, CTL(wcTabControl), WS_EX_TRANSPARENT, WC_STATIC, !"Volume locker V1.3\nGraham Downey (C) 2017-20", cBase, 215, 225, 173, 30)

        ShowStartPage()
        ShowWindow(CTL(wcTabCloseBtn), SW_HIDE)

        ' Create default tabs
        tcItem.mask = TCIF_TEXT OR TCIF_PARAM
        tcItem.lParam = -1 ' -1 means placeholder tab
        tcItem.pszText = @"[ Welcome ]"
        TabCtrl_InsertItem(CTL(wcTabControl), 0, @tcItem)        
        tcItem.pszText = @"+"
        TabCtrl_InsertItem(CTL(wcTabControl), 1, @tcItem)

        'Set fonts for all controls
        for CNT as integer = wcMain to wcLast-1
            SendMessage(CTL(CNT), WM_SETFONT, cast(WPARAM, fntDefault), true)
        next CNT

        SendMessage(CTL(wcTabControl), WM_SETFONT, cast(WPARAM, fntSmall), true)
        SendMessage(CTL(wcVersionLbl), WM_SETFONT, cast(WPARAM, fntSmall), true)
        SendMessage(CTL(wcTabCloseBtn), WM_SETFONT, cast(WPARAM, fntTiny), true)

        ' Create popup menu
        hPopupMenu = CreatePopupMenu() ' Create popup menu
        InsertMenu(hPopupMenu, &HFFFFFFFF, MF_BYPOSITION, IDM_SHOWHIDE, "Show/Hide Window")
        InsertMenu(hPopupMenu, &HFFFFFFFF, MF_SEPARATOR, IDM_SEPARATOR, "")
        InsertMenu(hPopupMenu, &HFFFFFFFF, MF_BYPOSITION, IDM_ABOUT, "About")
        InsertMenu(hPopupMenu, &HFFFFFFFF, MF_BYPOSITION, IDM_EXIT, "Exit")
                
        ReleaseDC(hWnd, hDC)
        return 0

    Case WM_DRAWITEM
        Dim as DRAWITEMSTRUCT ptr pDIS = Cast(DRAWITEMSTRUCT ptr, lParam)
        var hDC = pDIS->hDC
        
        Select Case wParam
        Case wcTabCloseBtn ' 'x' button to close active tab
            dim as rect r = pDis->rcItem
            
            var hPen = CreatePen(PS_SOLID, 0, GetSysColor(COLOR_WINDOWFRAME))
            var hOldBrush = SelectObject(hDC, GetSysColorBrush(COLOR_BTNFACE))
            var hOldPen = SelectObject(hDC, hPen)            
            var fntOld = SelectObject(hDC, fntTiny)
            
            Rectangle(hDC, r.left, r.top, r.right, r.bottom)
            DrawText(hDC, "X", 1, @r, DT_SINGLELINE OR DT_CENTER OR DT_VCENTER)
            
            SelectObject(hDC, hOldBrush)
            SelectObject(hDC, hOldPen)
            SelectObject(hDC, fntOld)
            DeleteObject(hPen)
            
        Case wcTabControl ' Draw a tab
            Select Case pDIS->itemAction
            Case ODA_DRAWENTIRE, ODA_SELECT
                dim as integer iCurTab = TabCtrl_GetCurSel(CTL(wcTabControl))
                dim as RECT r = pDIS->rcItem
                
                ' Draw background of tab
                var hPen = CreatePen(PS_SOLID, 0, GetSysColor(COLOR_BTNFACE))
                var hOldBrush = SelectObject(hDC, GetSysColorBrush(COLOR_BTNFACE))
                var hOldPen = SelectObject(hDC, hPen)
                
                Rectangle(hDC, r.left, r.top, r.right, r.bottom)
                
                SelectObject(hDC, hOldPen)
                SelectObject(hDC, hOldBrush)
                DeleteObject(hPen)

                ' Get control text & lParam value
                tcItem.mask = TCIF_PARAM OR TCIF_TEXT
                tcItem.pszText = malloc(MAXPNAMELEN)
                tcItem.cchTextMax = MAXPNAMELEN                
                TabCtrl_GetItem(CTL(wcTabControl), pDIS->itemID, @tcItem)
                
                ' If not a placeholder, add 'x' button
                if pDIS->itemID = iCurTab AndAlso tcItem.lParam >= 0 then
                    r.left += 3: r.right -= 20                    
                    SetWindowPos(CTL(wcTabCloseBtn), HWND_TOP, _
                                r.right+2, r.top+4, 0, 0, _
                                SWP_SHOWWINDOW	OR SWP_NOACTIVATE OR SWP_NOSIZE)
                else
                    ' Hide 'x' button if not applicable on current tab
                    ShowWindow(CTL(wcTabCloseBtn), SW_HIDE)
                end if
                    
                ' Draw text on tab
                DrawText(hDC, tcItem.pszText, -1, @r, _
                        DT_CENTER OR DT_VCENTER OR DT_SINGLELINE _
                        OR DT_END_ELLIPSIS OR DT_NOPREFIX)
                        
                free(tcItem.pszText)        
            End Select
        End Select

    ' Notifies us when mixer settings for an open device change
    Case MM_MIXM_CONTROL_CHANGE
        print "Line state changed: ";lParam
        var hMixer = cast(HMIXER, wParam)
        var dwControlID = cast(DWORD, lParam)

        ' Determine which device's state has changed
        dim as integer iDev = -1
        for iD as integer = 0 to iTotalDevices
            if devices(iD).hMixer = hMixer then 
                iDev = iD: exit for
            end if
        next iD
        
        ' Device not found (??)
        if iDev = -1 then return 0
        
        ' Determine which line has had it's state changed and correct it
        with devices(iDev)
            select case dwControlID
            case .iVolControlID
                print "Volume state changed!"
                if (.iControlFlags AND LOCK_VOLUME_ENABLED) then
                    var iNewVolume = _SysVolToPercent(GetLineState(devices(iDev), .iVolControlID))
                    if iNewVolume <> _SysVolToPercent(.iLockVolume) then
                        SetLineState(devices(iDev), .iVolControlID, .iLockVolume)
                    end if
                else        
                    var iVol = _SysVolToPercent(GetLineState(devices(iDev), .iVolControlID))

                    ' Check if device changed is currently open tab
                    tcItem.mask = TCIF_PARAM
                    var iCurTab = TabCtrl_GetCurSel(CTL(wcTabControl))
                    TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)
                    
                    if tcItem.lParam = iDev then
                        SendMessage(ctl(wcVolumeSlider), TBM_SETPOS, TRUE, iVol)
                        SetWindowText(CTL(wcVolLevelLbl), Str(iVol))
                    end if
                end if
                
            case .iMuteControlID
                print "Mute state changed!"
                var iMuteState = GetLineState(devices(iDev), .iMuteControlID)
                if (.iControlFlags AND LOCK_MUTE_ENABLED AndAlso iMuteState <> .iLockMute) then
                    SetLineState(devices(iDev), .iMuteControlID, .iLockMute)
                else
                    .iLockMute = iMuteState ' Set mute state checkbox
                    SendMessage(CTL(wcMuteState), BM_SETCHECK, iif(iMuteState, BST_CHECKED, BST_UNCHECKED), 0)
                end if
                
            end select
        end with
    
    case WM_GETMINMAXINFO:
        Dim as LPMINMAXINFO lpMMI = cast(LPMINMAXINFO, lParam)
        lpMMI->ptMinTrackSize.x = 330
        lpMMI->ptMinTrackSize.y = 233
        lpMMI->ptMaxTrackSize.y = 233
    
    Case WM_SIZE
        dim as integer iWid = LOWORD(lParam)
        dim as integer iHei = HIWORD(lParam)
        SetWindowPos(CTL(wcTabControl), NULL, NULL, NULL, iWid, iHei, SWP_NOZORDER OR SWP_NOMOVE)
        SetWindowPos(CTL(wcWelcomeLbl), NULL, NULL, NULL, iWid-20, iHei-40, SWP_NOZORDER OR SWP_NOMOVE)


        ' **** Volume Control Group ****
        SetWindowPos(CTL(wcVolGrp), NULL, 10, 25, iWid-20, iHei\2-20, SWP_NOZORDER)
        iWid -= 20
        
        ' Volume control
        SetWindowPos(CTL(wcVolumeLbl), NULL, 10, 20, 50, 20, SWP_NOZORDER)
        SetWindowPos(CTL(wcVolLevelLbl), NULL, 70, 20, 30, 20, SWP_NOZORDER)
        SetWindowPos(CTL(wcVolLockEnable), NULL, iWid-160, 20, 150, 20, SWP_NOZORDER)
        SetWindowPos(CTL(wcVolumeSlider), NULL, 5, 40, iWid-15, 40, SWP_NOZORDER)
        
        ' **** Mute Control Group ****
        SetWindowPos(CTL(wcMuteGrp), NULL, 10, 10+iHei\2, iWid, iHei\3-20, SWP_NOZORDER)
        
        ' Mute control
        SetWindowPos(CTL(wcMuteLockEnabled), NULL, 10, 20, 140, 20, SWP_NOZORDER)
        SetWindowPos(CTL(wcMuteState), NULL, iWid-130, 20, 120, 20, SWP_NOZORDER)

        SetWindowPos(CTL(wcVersionLbl), NULL, 10, iHei-35, 190, 30, SWP_NOZORDER)
        
    Case WM_NOTIFY
        static as integer iLastTab = 0
        var pNotifyData = cast(LPNMHDR, lParam)
        Select Case pNotifyData->code            
        Case TCN_SELCHANGING
            ' Store the last tab we were in
            iLastTab = TabCtrl_GetCurSel(CTL(wcTabControl))
            
        Case TCN_SELCHANGE
            var iTotalTabs = TabCtrl_GetItemCount(CTL(wcTabControl))
            var iCurTab = TabCtrl_GetCurSel(CTL(wcTabControl))
            
            ' If tab is '+' tab
            if iCurTab = iTotalTabs-1 then
                ' Open a dialog box here to add a new lock
                ShowWindow(hWndNewLock, SW_SHOW)
                EnableWindow(hWndMain, FALSE)
                
                TabCtrl_SetCurSel(CTL(wcTabControl), iLastTab)
                InvalidateRect(CTL(wcTabControl), NULL, TRUE)
            else
                tcItem.mask = TCIF_PARAM
                TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)
                ShowControls(devices(tcItem.lParam))
            end if
        End Select
    
    Case WM_HSCROLL
        tcItem.mask = TCIF_PARAM
        var iCurTab = TabCtrl_GetCurSel(CTL(wcTabControl))
        TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)        
        dim as integer iDev = tcItem.lParam        
        
        Select case lParam ' hwndScrollBar
        Case CTL(wcVolumeSlider)
            with devices(iDev)
                ' Convert to percents for easy calculations
                var iPercent = _SysVolToPercent(.iLockVolume)
                Select Case LOWORD(wParam) ' nScrollCode
                Case SB_BOTTOM ' Right side
                    iPercent = 100
                Case SB_TOP    ' Left side
                    iPercent = 0
                Case SB_LINELEFT
                    if iPercent > 0 then iPercent -= 1
                Case SB_LINERIGHT
                    if iPercent < 100 then iPercent += 1
                Case SB_PAGELEFT
                    var iPageSize = SendMessage(cast(HWND, lParam), TBM_GETPAGESIZE, 0, 0)
                    iPercent -= iPageSize
                    if iPercent < 0 then iPercent = 0
                Case SB_PAGERIGHT
                    var iPageSize = SendMessage(cast(HWND, lParam), TBM_GETPAGESIZE, 0, 0)
                    iPercent += iPageSize
                    if iPercent > 100 then iPercent = 100
                Case SB_THUMBPOSITION, SB_THUMBTRACK
                    iPercent = HIWORD(wParam) ' nPos
                    
                Case SB_ENDSCROLL ' Don't care about this message
                    return 0
                    
                End Select
                .iLockVolume = _PercentToSysVol(iPercent) ' Convert back to SysVol
                
                SetWindowText(CTL(wcVolLevelLbl), Str(iPercent))
                SetLineState(devices(iDev), .iVolControlID, .iLockVolume)
            end with
        End Select
    
    Case WM_COMMAND
        Select Case lParam 'hwndCtl
        Case 0 'Popup menu messages
            Select Case LOWORD(wParam) ' wID
            Case IDM_ABOUT
                MessageBox(hWnd, !"Volume Locker\r\n" + _
                                 !"Version 1.3\r\n\r\n" + _
                                 !"More great software available at:\r\n" + _ 
                                 !"http://grahamdowney.com/\r\n\r\n" + _
                                 !"Written by Graham Downey (C) 2017-2020", _
                                 "About Volume Locker", _
                                 MB_ICONINFORMATION)
                
            Case IDM_SHOWHIDE
                if IsWindowVisible(hWnd) then ' Hide if showing, open if minimized
                    ShowWindow(hWnd, SW_HIDE)
                else
                    ShowWindow(hWnd, SW_RESTORE) 'show window
                    SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE OR SWP_NOMOVE) 'move to top
                    SetForegroundWindow(hWnd) 'Set as forground
                end if
                
            Case IDM_EXIT
                SendMessage(hWnd, WM_CLOSE, 0, 0) 'Close the program
            End Select
            
        Case Else
            Select Case HIWORD(wParam)     ' wNotifyCode
            Case BN_CLICKED
                var iCurTab = TabCtrl_GetCurSel(CTL(wcTabControl))
                
                tcItem.mask = TCIF_PARAM
                TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)
                
                Select Case LOWORD(wParam) ' wID
                Case wcTabCloseBtn
                    TabCtrl_DeleteItem(CTL(wcTabControl), iCurTab)
                    devices(tcItem.lParam).iControlFlags AND= NOT(LOCK_OPEN)
                    CloseDevice(devices(tcItem.lParam)) ' Close mixer
                    
                    if TabCtrl_GetItemCount(CTL(wcTabControl)) = 1 then
                        ' Create default placeholder tab
                        tcItem.mask = TCIF_TEXT OR TCIF_PARAM
                        tcItem.lParam = -1
                        tcItem.pszText = @"[ Welcome ]"
                        TabCtrl_InsertItem(CTL(wcTabControl), 0, @tcItem)
                        ShowStartPage()
                    end if
                    
                    if iCurTab > 0 then iCurTab -= 1
                    TabCtrl_SetCurSel(CTL(wcTabControl), iCurTab)
                    InvalidateRect(CTL(wcTabControl), NULL, TRUE)
                    
                Case wcVolLockEnable
                    tcItem.mask = TCIF_PARAM
                    TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)
                    
                    var iLockState = SendMessage(CTL(wcVolLockEnable), BM_GETCHECK, 0, 0)
                    if iLockState = BST_CHECKED then
                        devices(tcItem.lParam).iControlFlags OR= LOCK_VOLUME_ENABLED
                        EnableWindow(CTL(wcVolumeSlider), TRUE)
                        EnableWindow(CTL(wcVolumeLbl), TRUE)
                        EnableWindow(CTL(wcVolLevelLbl), TRUE)
                    else
                        devices(tcItem.lParam).iControlFlags AND= NOT(LOCK_VOLUME_ENABLED)
                        EnableWindow(CTL(wcVolumeSlider), FALSE)
                        EnableWindow(CTL(wcVolumeLbl), FALSE)
                        EnableWindow(CTL(wcVolLevelLbl), FALSE)
                    end if
                    
                Case wcMuteLockEnabled
                    tcItem.mask = TCIF_PARAM
                    TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)
                    
                    with devices(tcItem.lParam)
                        var iLockState = SendMessage(CTL(wcMuteLockEnabled), BM_GETCHECK, 0, 0)
                        if iLockState = BST_CHECKED then
                            .iControlFlags OR= LOCK_MUTE_ENABLED
                            EnableWindow(CTL(wcMuteState), TRUE)
                        else
                            .iControlFlags AND= NOT(LOCK_MUTE_ENABLED)
                            EnableWindow(CTL(wcMuteState), FALSE)
                        end if
                    end with
                    
                Case wcMuteState
                    tcItem.mask = TCIF_PARAM
                    TabCtrl_GetItem(CTL(wcTabControl), iCurTab, @tcItem)
                    
                    with devices(tcItem.lParam)
                        var iEnabled = SendMessage(CTL(wcMuteState), BM_GETCHECK, 0, 0)
                        .iLockMute = (iEnabled <> FALSE)
                        SetLineState(devices(tcItem.lParam), .iMuteControlID, iEnabled)
                    end with
                End Select
            End Select
            
        End Select
        
    Case UM_ADDDEVICE ' Sent by NewLockDialog to add device lock tab
        tcItem.mask = TCIF_PARAM
        
        ' Check if the open tab was placeholder
        var iTotalTabs = TabCtrl_GetItemCount(CTL(wcTabControl))
        if iTotalTabs = 2 then
            TabCtrl_GetItem(CTL(wcTabControl), 0, @tcItem)
            if tcItem.lParam = -1 then ' It was placeholder, delete it
                TabCtrl_DeleteItem(CTL(wcTabControl), 0)
                iTotalTabs -= 1
            end if
        end if
        
        ' Make sure this isn't a duplicate of already locked devices
        for i as integer = 0 to iTotalTabs-1
            TabCtrl_GetItem(CTL(wcTabControl), i, @tcItem)
            if tcItem.lParam = lParam then return 0
        next i
                
        ' Add new item to tab control
        tcItem.mask = TCIF_TEXT OR TCIF_PARAM
        tcItem.lParam = lParam
        tcItem.pszText = @devices(lParam).szName
        TabCtrl_InsertItem(CTL(wcTabControl), iTotalTabs-1, @tcItem)
        
        ' Select new tab
        TabCtrl_SetCurSel(CTL(wcTabControl), iTotalTabs-1)
        InvalidateRect(CTL(wcTabControl), NULL, TRUE)
        
        ' Open device and update controls
        OpenDevice(lParam, devices(lParam))
        with devices(lParam)
            .iControlFlags OR= LOCK_OPEN
            if (.iControlFlags AND LOCK_VOLUME_ENABLED) then
                SetLineState(devices(lParam), .iVolControlID, .iLockVolume)
            end if
        end with
        ShowControls(devices(lParam))

        
    'Tray icon events
    Case UM_SHELLICON 
        Select Case LOWORD(lParam)
        Case WM_LBUTTONDBLCLK
            if IsWindowVisible(hWnd) then ' Hide if showing, open if minimized
                ShowWindow(hWnd, SW_HIDE)
            else
                ShowWindow(hWnd, SW_RESTORE) 'show window
                SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE OR SWP_NOMOVE) 'move to top
                SetForegroundWindow(hWnd) 'Set as forground
            end if
            
        Case WM_RBUTTONDOWN, WM_LBUTTONDOWN
            Dim as POINT ClickPt            
            GetCursorPos(@ClickPt)
            TrackPopupMenu(hPopupMenu, TPM_BOTTOMALIGN OR TPM_LEFTALIGN, ClickPt.X, ClickPt.Y, 0, hWnd, NULL)
        End Select

    'Remove from taskbar when minimized
    Case WM_SYSCOMMAND
        Select Case wParam
        Case SC_MINIMIZE    ' Minimize message
            ShowWindow(hWnd, SW_HIDE)
            SetForegroundWindow(hWnd)
            return 0
            
        End Select
    
    Case WM_DESTROY
        Shell_NotifyIcon(NIM_DELETE, @nidApp) ' Delete tray icon
        DestroyMenu(hPopupMenu) ' Delete popup menu
        DestroyWindow(hWndNewLock)
        DeleteObject(fntDefault)
        DeleteObject(fntSmall)
        PostQuitMessage(0)
        return 0
    End Select
    
    return DefWindowProc(hWnd, iMsg, wParam, lParam)
End Function
   
Sub WinMain(hInstance as HINSTANCE, hPrevInstance as HINSTANCE, _
            szCmdLine as PSTR, iCmdShow as Integer)
            
    Dim as HWND       hWnd
    Dim as MSG        msg
    Dim as WNDCLASSEX wcls

    #if ShowDebug
        AllocConsole() 'Show console
    #endif

    ' Only allow one instance at a time
    var hwndRunning = FindWindow(szAppName, NULL)
    if hwndRunning then
        ShowWindow(hwndRunning, SW_RESTORE)
        SetForegroundWindow(hwndRunning)
        SetFocus(hWndRunning)
        system
    end if

    wcls.cbSize        = sizeof(WNDCLASSEX)
    wcls.style         = CS_HREDRAW OR CS_VREDRAW
    wcls.lpfnWndProc   = @WndProc
    wcls.cbClsExtra    = 0
    wcls.cbWndExtra    = 0
    wcls.hInstance     = hInstance
    wcls.hIcon         = LoadIcon(hInstance, "FB_PROGRAM_ICON") 
    wcls.hCursor       = LoadCursor(NULL, IDC_ARROW)
    wcls.hbrBackground = cast(HBRUSH, COLOR_BTNFACE + 1)
    wcls.lpszMenuName  = NULL
    wcls.lpszClassName = strptr(szAppName)
    wcls.hIconSm       = LoadIcon(hInstance, "FB_PROGRAM_ICON")
    
    if (RegisterClassEx(@wcls) = FALSE) then
        Print "Error! Failed to register window class ", Hex(GetLastError())
        sleep: system
    end if
        
    const WINDOW_STYLE = WS_CLIPCHILDREN OR WS_OVERLAPPEDWINDOW
    
    hWnd = CreateWindow(szAppName, _            ' window class name
                        szCaption, _            ' Window caption
                        WINDOW_STYLE, _         ' Window style
                        CW_USEDEFAULT, _        ' Initial X position
                        CW_USEDEFAULT, _        ' Initial Y Posotion
                        WINDOW_WIDTH, _         ' Window width
                        WINDOW_HEIGHT, _        ' Window height
                        NULL, _                 ' Parent window handle
                        NULL, _                 ' Window menu handle
                        hInstance, _            ' Program instance handle
                        NULL)                   ' Creation parameters
                        
    if hWnd = NULL then system
    
    hWndMain = hWnd 

    ShowWindow(hWnd, iCmdShow)
    UpdateWindow(hWnd)
    
    'Setup tray icon
    nidApp.cbSize = SizeOf(NOTIFYICONDATA) 'NOTIFYICONDATA_V1_SIZE
    nidApp.hWnd = hWnd
    nidApp.uID = 1
    nidApp.uFlags = NIF_ICON OR NIF_MESSAGE OR NIF_TIP
    nidApp.hIcon = LoadIcon(hInstance, "FB_PROGRAM_ICON")
    nidApp.uCallbackMessage = UM_SHELLICON
    nidApp.szTip = szAppName
    Shell_NotifyIcon(NIM_ADD, @nidApp)
        
    hWndNewLock = NewLockDialog.initNewLockDialog(hWnd, @devices(0), iTotalDevices)
    
    SetTimer(hWnd, 1, 10, cast(TIMERPROC, @WndProc))
    
    LoadSettings()
    
    'while (GetMessage(@msg, NULL, 0, 0))
    while (msg.message <> WM_QUIT)
        while (PeekMessage(@msg, NULL, 0, 0, PM_REMOVE))
            TranslateMessage(@msg)
            DispatchMessage(@msg)
        wend
        
        for iD as integer = 0 to iTotalDevices-1
            with devices(iD)
                ' Only check device if it's loaded in GUI
                if .iControlFlags AND LOCK_OPEN then
                    var iMuteState = GetLineState(devices(iD), .iMuteControlID)
                    var iCheckState = SendMessage(CTL(wcMuteState), BM_GETCHECK, 0, 0)
                    if iMuteState <> .iLockMute then
                        if .iControlFlags AND LOCK_MUTE_ENABLED then
                             SetLineState(devices(iD), .iMuteControlID, .iLockMute)
                        else
                            .iLockMute = iMuteState
                            SendMessage(CTL(wcMuteState), BM_SETCHECK, iif(iMuteState, BST_CHECKED, BST_UNCHECKED), 0)
                        end if
                    end if
                    
                    
                end if
            end with
        next iD
        
        sleep 1,1
    wend
    
    SaveSettings()
    KillTimer(hWnd, 1)
    UnregisterClass(szAppName, hInstance)
    system msg.wParam
End Sub

Sub EnumInputDevices()
    print "Enumerating input devices..."
    
    iTotalDevices = waveInGetNumDevs()    'Get number of input devices
    Redim devices(iTotalDevices-1)
    
    if iTotalDevices = 0 then
        MessageBox(NULL, !"Warning!\r\nNo wave input devices found!", "Volume Locker", MB_ICONWARNING)'MB_ICONINFORMATION)
    end if
        
    ' Get device mixer info
    Dim as MIXERLINE mxl
    mxl.cbStruct = sizeof(MIXERLINE)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN
    
    for iD as integer = 0 to iTotalDevices-1
        with devices(iD)

            var result = mixerGetLineInfo(cast(HMIXEROBJ, iD), @mxl, MIXER_OBJECTF_WAVEIN OR MIXER_GETLINEINFOF_COMPONENTTYPE)
            if (result <> MMSYSERR_NOERROR) then
                print "Failed to get line info for device "; iD; ": 0x"; hex(result)
            end if
            
            strcpy(.szName, mxl.Target.szPname)
            
            print iD;") ";.szName
        end with
    next iD
End Sub

Function OpenDevice(iDev as integer, device as DeviceControl) as integer
    with device                
        ' Open mixer for specified waveIn device
        var result = mixerOpen(@.hMixer, iDev, cast(DWORD_PTR, hWndMain), NULL, _
                                MIXER_OBJECTF_WAVEIN OR CALLBACK_WINDOW)
        if (result <> MMSYSERR_NOERROR) then
            print "Failed to open mixer for ";iDev;": 0x";hex(result)
            return result
        end if
                
        'Get mixer line controls
        Dim as MIXERLINECONTROLS mxlc
        mxlc.dwLineID = &HFFFF0000 ' Line ID containing volume/mute controls
                                   ' This constant is the master vol on pre-vista
                
        Dim as MIXERCONTROL mxc   ' Control info returned in this struct
        mxlc.cControls = 1        ' Only passing one control struct
        mxlc.pamxctrl = @mxc
        
        mxc.cbStruct  = SizeOf(MIXERCONTROL)
        mxlc.cbmxctrl = SizeOf(MIXERCONTROL)
        mxlc.cbStruct = SizeOf(MIXERLINECONTROLS)
        
        ' Get volume control ID first
        mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_VOLUME
        result = mixerGetLineControls(cast(HMIXEROBJ, .hMixer), @mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        if (result = MMSYSERR_NOERROR) then
            .iVolControlID = mxc.dwControlID
        else
            print MIXERR_INVALCONTROL, MIXERR_INVALLINE, MMSYSERR_INVALFLAG, MMSYSERR_INVALPARAM, MMSYSERR_NODRIVER
            print "Error #";result;" getting ";.szName;" volume control!"
        end if
        
        ' Get mute control ID
        mxlc.dwControlType = MIXERCONTROL_CONTROLTYPE_MUTE
        result = mixerGetLineControls(cast(HMIXEROBJ, .hMixer), @mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        if (result = MMSYSERR_NOERROR) then
            .iMuteControlID = mxc.dwControlID
        else
            print "Error #";result;" getting ";.szName;" mute control!"
        end if
    end with
    
    return MMSYSERR_NOERROR
End Function

Sub CloseDevice(device as DeviceControl)
    with device
        mixerClose(.hMixer)
        .hMixer = 0
    end with
End Sub

function GetLineState(device as DeviceControl, dwControlID as DWORD) as DWORD
    with device
                       
        Dim as MIXERCONTROLDETAILS mxcd
        mxcd.dwControlID = dwControlID ' control ID
        mxcd.cChannels = 1             ' Apply volume uniformly to all channels
        mxcd.cMultipleItems = 0        ' Only setting one value
           
        ' Control status returned in this struct
        Dim as MIXERCONTROLDETAILS_UNSIGNED details
        mxcd.paDetails = @details
        mxcd.cbDetails = sizeof(MIXERCONTROLDETAILS_UNSIGNED)
        mxcd.cbStruct  = sizeof(MIXERCONTROLDETAILS)
        
        var result = mixerGetControlDetails(cast(HMIXEROBJ, .hMixer), @mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
 
        if (result <> MMSYSERR_NOERROR) then
            print "Error #"; result; " getting line info for "; .szName
        end if
       
        return details.dwValue ' Return line control value
    end with
end function

Sub SetLineState(device as DeviceControl, dwControlID as DWORD, dwValue as DWORD)
    with device
        Dim as MIXERCONTROLDETAILS mxcd
        mxcd.dwControlID = dwControlID ' control ID
        mxcd.cChannels = 1             ' Apply volume uniformly to all channels
        mxcd.cMultipleItems = 0        ' Only setting one value
            
        ' Set the volume level for the line
        Dim as MIXERCONTROLDETAILS_UNSIGNED details
        details.dwValue = dwValue
    
        mxcd.paDetails = @details
        mxcd.cbDetails = sizeof(details)
        mxcd.cbStruct  = sizeof(MIXERCONTROLDETAILS)
            
        var result = mixerSetControlDetails(cast(HMIXEROBJ, .hMixer), @mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
        if (result <> MMSYSERR_NOERROR) then
            print "Error #"; result; " setting line info for "; .szName
        end if
    end with
End Sub

Sub SaveSettings()
    dim as string sCfgPath = exePath+"\config.ini"
    Dim as integer iSection = 0 ' Start at first section
    
    ' [Lock n]
    for iD as integer = 0 to iTotalDevices-1
        ' Only write config for opened devices
        if (devices(iD).iControlFlags AND LOCK_OPEN) = 0 then continue for
        
        dim as string sSection = "Lock "+str(iSection+1) ' .ini Section name
        
        with devices(iD)
            ' Device name
            WritePrivateProfileString(sSection, "deviceName", .szName, sCfgPath)
            
            ' Volume
            dim as string sTF = iif((.iControlFlags AND LOCK_VOLUME_ENABLED), "true", "false") 
            WritePrivateProfileString(sSection, "volEnabled", sTF, sCfgPath)
            WritePrivateProfileString(sSection, "volume", str(_SysVolToPercent(.iLockVolume)), sCfgPath)
            
            ' Mute
            sTF = iif((.iControlFlags AND LOCK_MUTE_ENABLED), "true", "false") 
            WritePrivateProfileString(sSection, "muteEnabled", sTF, sCfgPath)
            sTF = iif(.iLockMute, "true", "false") 
            WritePrivateProfileString(sSection, "mute", sTF, sCfgPath)
        end with
        
        iSection += 1 ' Go to next section
    next iD
    
    ' [default]
    WritePrivateProfileString("default", "configuredLocks", str(iSection), sCfgPath)
End Sub

Sub LoadSettings()
    dim as string sCfgPath = exePath+"\config.ini"

    if dir(sCfgPath) = "" then ' Create config if it doesn't exist
        if open(sCfgPath for output as #1) then
            print sCfgPath
            print "Error creating config file!"
        else
            ' Write default config
            print #1, !"[default]\r\nconfiguredLocks=0\r\n"
            close #1
        end if
    end if

    dim as integer iCurLock = 0
    
    var iLoadLocks = GetPrivateProfileInt("default", "configuredLocks", 0, sCfgPath)
    for iL as integer = 1 to iLoadLocks            
        dim as string sSection = "Lock "+str(iL) ' .ini Section name
        
        dim as integer iDeviceID = -1
        
        ' Check if device still exists on the system
        Dim as zstring*256 szTmp = any
        
        ' Get device name
        var iLen = GetPrivateProfileString(sSection, "deviceName", "", szTmp, 255, sCfgPath)
        if iLen <= 0 then continue for ' If no device name skip device
        
        ' check if device exists on system
        for iD as integer = 0 to iTotalDevices-1
            if szTmp = devices(iD).szName then
                ' Can't configure same device twice, ignore duplicates
                if (devices(iD).iControlFlags AND LOCK_OPEN) then 
                    dim as string sErr = !"Duplicate device configuration found!\r\n"""+szTmp+!"""\r\nIgnoring duplicates"
                    MessageBox(NULL, sErr, "Volume Locker -- Warning", MB_ICONWARNING)
                    continue for, for
                end if
                
                ' Save device ID and exit for
                iDeviceID = iD
                exit for
            end if
        next iD
    
        ' Display error if device not found on system
        dim as string sErr = !"Previously configured device not found!\r\n"""+szTmp+""""
        if iDeviceID < 0 then MessageBox(NULL, sErr, "Volume Locker -- Error", MB_ICONERROR)
    
        dim as zstring*6 szTF = any
        with devices(iDeviceID)
            ' Volume lock enabled?
            iLen = GetPrivateProfileString(sSection, "volEnabled", "true", szTF, 5, sCfgPath)
            .iControlFlags OR= iif(lcase(szTF) = "true", LOCK_VOLUME_ENABLED, 0)
            
            ' Volume to lock to
            .iLockVolume = _PercentToSysVol(GetPrivateProfileInt(sSection, "volume", 100, sCfgPath))
            
            ' Mute lock enabled?
            iLen = GetPrivateProfileString(sSection, "muteEnabled", "false", szTF, 5, sCfgPath)
            .iControlFlags OR= iif(lcase(szTF) = "true", LOCK_MUTE_ENABLED, 0)
            
            ' Mute lock state
            iLen = GetPrivateProfileString(sSection, "mute", "false", szTF, 5, sCfgPath)
            .iLockMute = lcase(szTF) = "true"
            
            .iControlFlags OR= LOCK_OPEN
        end with
        
        ' Add configured tab
        PostMessage(hWndMain, UM_ADDDEVICE, NULL, iDeviceID)
        
        iCurLock += 1 ' Lock configured, go to next one
    next iL
End Sub

