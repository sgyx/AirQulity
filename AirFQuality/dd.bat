copy .\\MSCOMM.SRG %windir%\system32
copy .\\MSCOMM32.DEP %windir%\system32
copy .\\MSCOMM32.oca %windir%\system32
copy .\\mscomm32.ocx %windir%\system32

regsvr32 mscomm32.ocx