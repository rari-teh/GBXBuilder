# GBXBuilder
This program creates, modifies and removes metadata on Game Boy (Color) ROMs for use with GBX-compliant emulators, such as [hhugboy](https://github.com/tzlion/hhugboy). GBX is a footered Game Boy (Color) ROM format that includes metadata like mapper type and RAM size. This is pretty much only useful for emulating unlicensed cartridges, which often have wrong metadata on the internal header to make piracy more difficult.

The GBX ROM format was created by [taizou](https://github.com/tzlion). The file format specification is available [here](http://hhug.me/gbx/1.0).

This program uses truncatefile.bas by [countingpine](https://github.com/countingpine) and LongToByteArray by [FreeVBCode](https://www.freevbcode.com/).

Made with Microsoft Visual Basic 6.0

## Requirement
This program requires the Common Dialog ActiveX Control (COMDLG32.OCX). If you get a missing file error, here’s what you should do:

### 64-bit
1. Get the file from [here](https://www.ocxme.com/files/comdlg32/ac9bd4138ba1cece3c25f62166b0ba70). If you don’t trust a random library download website, you can also get it yourself by opening [this MSI file from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=7030) with 7-Zip and extracting it from there.
2. Move COMDLG32.OCX to C:\Windows\SysWOW64.
3. Open the Command Prompt as administrator and run `C:\Windows\SysWOW64\regsvr32.exe C:\Windows\SysWOW64\COMDLG32.OCX`.
4. Reboot.

### 32-bit
1. Get the file from [here](https://www.ocxme.com/files/comdlg32/ac9bd4138ba1cece3c25f62166b0ba70). If you don’t trust a random library download website, you can also get it yourself by opening [this MSI file from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=7030) with 7-Zip and extracting it from there.
2. Move COMDLG32.OCX to C:\Windows\System32.
3. Open the Command Prompt as administrator and run `regsvr32 COMDLG32.OCX`.
4. Reboot.
