enum UserMessages
    UM_SHELLICON = WM_USER+1
    UM_ADDDEVICE
end enum

enum ControlFlags
    LOCK_VOLUME_ENABLED = 1
    LOCK_MUTE_ENABLED   = 2
    LOCK_OPEN           = 4
end enum

type DeviceControl field=1
    as HMIXER  hMixer         ' Handle of mixer for device
    
    as integer iVolControlID  ' ID of volume control
    as integer iMuteControlID ' ID of mute control
    
    as integer iControlFlags  ' Control flags
    as integer iLockVolume    ' Volume we're setting device to (SysUnits)
    as integer iLockMute      ' State of mute lock
            
    as zString*MAXPNAMELEN szName ' Device name
end type

