


::
::	Move the files older the /minage:7 (days) to a new folder.
::


robocopy.exe D:\Temp D:\archive-older-then\d_temp *.* /minage:7 /z /e /move /r:1 /w:1 /np



::
::	Compress the folder D:\archive-older-then
::

7za.exe a -r D:\archive-older-then\d_temp.7z D:\archive-older-then\d_temp\*.* 