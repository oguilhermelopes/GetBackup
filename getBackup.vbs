On Error Resume Next
Set objFSO  = CreateObject("Scripting.FileSystemObject")
Set log = objFSO.OpenTextFile("log.txt", 8, True)

Function getBackup(ip, path)
getStringDate = "N/A"
getStringSize = "N/A"
getStringPath = "N/A"
getSize = "N/A"

getStringDate = objFSO.GetFile("\\"& ip &"\"& path &"").DateLastModified
getStringDate = objFSO.GetFile("\\"& ip &"\"& path &"").DateLastModified
getStringSize = objFSO.GetFile("\\"& ip &"\"& path &"").Size/1024/1024/1024
getStringPath = objFSO.GetFile("\\"& ip &"\"& path &"").Path
getSize = Split(getStringSize,",")

coleta.writeline("Name:"& ip)
coleta.writeline("Size:"& left(getStringSize,5)&"GB")
coleta.writeline("Path:"& getStringPath)
coleta.writeline("Date:"& getStringDate)
coleta.writeline("---------")
coleta.writeline("")
End Function

'Set Server IP, Path and Name of .bak
call getBackup("10.0.0.1","d$\backup\teste.bak")

wscript.echo "End of Collection"
