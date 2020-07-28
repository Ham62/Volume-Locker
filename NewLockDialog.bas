namespace NewLockDialog
#include once "windows.bi"
#include once "LockControl.bi"

enum WindowControls_
    wcMain
  
    wcDevSelect ' Device drop down menu
    wcDevSelectLbl
  
    wcOkBtn
    wcCancelBtn
  
    wcLast
end enum
const WINDOW_WIDTH = 290, WINDOW_HEIGHT = 100
dim shared as HWND CTL(wcLast)    'controls
Dim shared as HWND hWndOwner
Dim Shared as HINSTANCE hInstance

Dim shared as integer iTotalDevices
Dim shared as DeviceControl ptr pDevices

Function ProgDialogProc(hWnd as HWND, iMsg as uLong, wParam as WPARAM, lParam as LPARAM) as LRESULT   
    static as hFont fntDefault, fntSmall
    
    Select Case iMsg
    Case WM_CREATE        
        '**** Center window on desktop ****'
        Scope   'Calculate Client Area Size
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
        end Scope           
        
        InitCommonControls()
        
        ' Create window controls
        #define CreateControl(mID, mExStyle, mClass, mCaption, mStyle, mX, mY, mWid, mHei) CTL(mID) = CreateWindowEx(mExStyle,mClass,mCaption,mStyle,mX,mY,mWid,mHei,hwnd,cast(hmenu,mID),hInstance,null)
        const cBase = WS_CHILD OR WS_VISIBLE
        const cLabelStyle = cBase
        const cDropList = cBase OR CBS_DROPDOWNLIST OR CBS_HASSTRINGS OR CBS_NOINTEGRALHEIGHT OR WS_VSCROLL
        
        dim as integer iPadding = 10
        dim as integer iHei = 24, iWid = 60
                 
        CreateControl(wcOKBtn, NULL, WC_BUTTON, "OK", cBase, _
                      WINDOW_WIDTH\4-iWid\2, WINDOW_HEIGHT-iHei-iPadding, _
                      iWid, iHei)
                      
        CreateControl(wcCancelBtn, NULL, WC_BUTTON, "Cancel", cBase, _
                      (WINDOW_WIDTH\4)*3-iWid\2, WINDOW_HEIGHT-iHei-iPadding, _
                      iWid, iHei)

        ' Populate list of devices and pre-select first device
        CreateControl(wcDevSelectLbl, NULL, WC_STATIC, "Select Device:", cBase, iPadding, iPadding/2, WINDOW_WIDTH-(iPadding*2), 20)
        CreateControl(wcDevSelect, NULL, WC_COMBOBOX, "", cDropList, iPadding, iPadding*3, WINDOW_WIDTH-(iPadding*2), 100)
        For iD as integer = 0 to iTotalDevices-1
            SendMessage(ctl(wcDevSelect), CB_ADDSTRING, 0, cast(.LPARAM, @pDevices[iD].szName))
        Next iD
        SendMessage(ctl(wcDevSelect), CB_SETCURSEL, 0, 0)

        ' Create fonts
        var hDC = GetDC(hWnd)
        var nHeight = -MulDiv(9, GetDeviceCaps(hDC, LOGPIXELSY), 72) 'calculate size matching DPI
        fntDefault = CreateFont(nHeight,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,0,0,"Verdana")
        
        nHeight = -MulDiv(7, GetDeviceCaps(hDC, LOGPIXELSY), 72)
        fntSmall = CreateFont(nHeight,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,0,0,0,0,"Verdana")

        for iCTL as integer = wcMain to wcLast-1
            SendMessage(CTL(iCTL), WM_SETFONT, cast(.WPARAM, fntDefault), TRUE)
        next iCTL
                
        ReleaseDC(hWnd, hDC)
        return 0
     
     
    Case WM_CTLCOLORSTATIC
        var hdcStatic = cast(HDC, wParam)
        SetBkMode(hdcStatic, TRANSPARENT)
        return cast(integer, GetSysColorBrush(COLOR_BTNFACE))

    Case WM_KEYDOWN
        if wParam = VK_RETURN then
            var iCurDevice = SendMessage(CTL(wcDevSelect), CB_GETCURSEL, 0, 0)
            PostMessage(hWndOwner, UM_ADDDEVICE, NULL, iCurDevice)
            PostMessage(hWnd, WM_CLOSE, 0, 0)
        end if
            

    Case WM_COMMAND
        Select Case HIWORD(wParam)     ' wNotifyCode
        Case BN_CLICKED
            Select Case LOWORD(wParam) ' wID
            Case wcOKBtn
                var iCurDevice = SendMessage(CTL(wcDevSelect), CB_GETCURSEL, 0, 0)
                PostMessage(hWndOwner, UM_ADDDEVICE, NULL, iCurDevice)
                PostMessage(hWnd, WM_CLOSE, 0, 0)
                
            Case wcCancelBtn
                PostMessage(hWnd, WM_CLOSE, 0, 0)

            End Select
        End Select

    Case WM_CLOSE                
        ' Set parent to active window and hide this dialog
        EnableWindow(hWndOwner, TRUE)
        SetActiveWindow(hWndOwner)
        ShowWindow(hWnd, SW_HIDE)
        return 0

    Case WM_DESTROY
        DeleteObject(fntDefault)
        DeleteObject(fntSmall)
        
    End Select
    
    return DefWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Function initNewLockDialog(hWndParent as HWND, pDevs as DeviceControl ptr, iTotalDevs as integer, hInst as .HINSTANCE = NULL) as HWND
    static as zstring ptr szClass = @"NewLockDialog"
    Dim as HWND       hWnd
    Dim as WNDCLASSEX wcls
        
    hWndOwner = hWndParent
    if hInst = NULL then hInst = GetModuleHandle(NULL)
    hInstance = hInst
    
    iTotalDevices = iTotalDevs
    pDevices = pDevs
    
    wcls.cbSize        = sizeof(WNDCLASSEX)
    wcls.style         = CS_HREDRAW OR CS_VREDRAW
    wcls.lpfnWndProc   = cast(WNDPROC, @ProgDialogProc)
    wcls.cbClsExtra    = 0
    wcls.cbWndExtra    = 0
    wcls.hInstance     = hInstance
    wcls.hIcon         = LoadIcon(hInstance, "FB_PROGRAM_ICON") 
    wcls.hCursor       = LoadCursor(NULL, IDC_ARROW)
    wcls.hbrBackground = cast(HBRUSH, COLOR_BTNFACE + 1)
    wcls.lpszMenuName  = NULL
    wcls.lpszClassName = szClass
    wcls.hIconSm       = LoadImage(hInstance, "FB_PROGRAM_ICON", _
                                   IMAGE_ICON, 16, 16, LR_DEFAULTSIZE)
    
    if (RegisterClassEx(@wcls) = FALSE) then
        Print("Error! Failed to register window class " & Hex(GetLastError()))
        sleep: system
    end if
    
    const WINDOW_STYLE = WS_OVERLAPPEDWINDOW XOR WS_THICKFRAME XOR WS_MAXIMIZEBOX
    
    hWnd = CreateWindow(szClass, _              ' window class name
                        "Select New Device", _   ' Window caption
                        WINDOW_STYLE, _         ' Window style
                        CW_USEDEFAULT, _        ' Initial X position
                        CW_USEDEFAULT, _        ' Initial Y Posotion
                        WINDOW_WIDTH, _         ' Window width
                        WINDOW_HEIGHT, _        ' Window height
                        hWndParent, _           ' Parent window handle
                        NULL, _                 ' Window menu handle
                        hInstance, _            ' Program instance handle
                        NULL)                   ' Creation parameters


    return hWnd

End Function

end namespace

