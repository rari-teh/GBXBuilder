# GBXBuilder
This program creates, modifies and removes metadata on Game Boy (Color) ROMs for use with GBX-compliant emulators, such as [hhugboy](https://github.com/tzlion/hhugboy). GBX is a footered Game Boy (Color) ROM format that includes metadata like mapper type and RAM size. This is pretty much only useful for emulating unlicensed cartridges, which often have wrong metadata on the internal header to make piracy more difficult.

The GBX ROM format was created by [taizou](https://github.com/tzlion). The file format specification is available [here](http://hhug.me/gbx/1.0).

This program uses truncatefile.bas by [countingpine](https://github.com/countingpine) and LongToByteArray by [FreeVBCode](https://www.freevbcode.com/).

Made with Microsoft Visual Basic 6.0

## Requirement
This program requires the Common Dialog ActiveX Control (COMDLG32.OCX). If you get a missing file error, you should install [Visual Basic 6.0 Runtime Plus](https://sourceforge.net/projects/vb6extendedruntime/files/latest/download).