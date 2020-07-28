# Volume Locker
Volume Locker allows you to quickly and easily set and "lock" the state of the input devices on your system, preventing pesky programs from changing the volume or muting/unmuting your devices.

### INSTALLING AND USING:
After clicking the '+' tab to show a list of available input devices, you
will be brought to a screen mirroring the device's current volume/mute
state. Once you enable the mute or volume locks for a device, Volume
Locker takes over the mixer for that input and will keep it set at the
configured volume/mute state, blocking other programs from changing it
until the locks are released.

When closing Volume Locker, the program automatically saves the loaded
lock configurations to disk and automatically loads them back the next
time you restart the program.

### COMPILING:
The project is written in [FreeBASIC](http://www.freebasic.net/) so you will need to download and install the compiler before you can build Volume Locker.<br />
Volume Locker can be compiled with the command:<br />
```fbc VolLocker.bas -s gui res\VolLocker.rc ```

### CONTACT:
You can contact me (Ham62) via IRC in ##freebasic on irc.freenode.net<br />
If you find any bugs you can raise them through the Github issues tab

### LICENSE:
    VOLUME LOCKER IS FREEWARE.  IT MAY BE USED AND
    DISTRIBUTED FREELY.  NO FEE MAY BE CHARGED FOR ITS 
    USE OR DISTRIBUTION.  NO WARRANTIES ARE GRANTED 
    WITH THIS PRODUCT. BY USING THIS PRODUCT, YOU AGREE 
    TO THESE LICENSES.
    
