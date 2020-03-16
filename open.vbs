Set TI = CreateObject("WMPlayer.OCX.7" )
Set CDROM = TI.cdromCollection
if CDROM.Count >= 1 then
For i = 0 to CDROM.Count - 1
CDROM.Item(i).Eject
Next ' CDTRAY
End If
