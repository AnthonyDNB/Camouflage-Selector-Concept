Set objShell = CreateObject("WScript.Shell")

filename = objShell.ExpandEnvironmentStrings("%USERPROFILE%") &"\AppData\Local\KillHouseGames\DoorKickers2\roster.xml"
filename2 = objShell.ExpandEnvironmentStrings("%USERPROFILE%") &"\AppData\Local\KillHouseGames\DoorKickers2\roster2.xml"
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename,1)
Set f2 = fso.createTextFile(filename2)
test = "        <Trooper class=""Assault"">"
Do Until f.AtEndOfStream
    line = f.ReadLine
    If test = line Then
      f2.WriteLine line
      For i = 1 to 18
        f2.WriteLine f.ReadLine
      Next
      line = f.ReadLine
		IF line = "                <MountedGun/>" then
			f2.WriteLine "                <MountedGun name=""default_camo_assaulter""/>"
		Else
			f2.WriteLine line
		End IF
    Else
      f2.WriteLine line
    End IF
Loop
f.Close
f2.Close
fso.DeleteFile filename
fso.MoveFile filename2, filename


Set f = fso.OpenTextFile(filename,1)
Set f2 = fso.createTextFile(filename2)
test = "        <Trooper class=""Support"">"
Do Until f.AtEndOfStream
    line = f.ReadLine
    If test = line Then
      f2.WriteLine line
      For i = 1 to 18
        f2.WriteLine f.ReadLine
      Next
      line = f.ReadLine
		IF line = "                <MountedGun/>" then
			f2.WriteLine "                <MountedGun name=""default_camo_support""/>"
		Else
			f2.WriteLine line
		End IF
    Else
      f2.WriteLine line
    End IF
Loop

f.Close
f2.Close
fso.DeleteFile filename
fso.MoveFile filename2, filename
Set f = fso.OpenTextFile(filename,1)
Set f2 = fso.createTextFile(filename2)
test = "        <Trooper class=""Marksman"">"
Do Until f.AtEndOfStream
    line = f.ReadLine
    If test = line Then
      f2.WriteLine line
      For i = 1 to 18
        f2.WriteLine f.ReadLine
      Next
      line = f.ReadLine
		IF line = "                <MountedGun/>" then
			f2.WriteLine "                <MountedGun name=""default_camo_marksman""/>"
		Else
			f2.WriteLine line
		End IF
    Else
      f2.WriteLine line
    End IF
Loop
f.Close
f2.Close
fso.DeleteFile filename
fso.MoveFile filename2, filename
Set f = fso.OpenTextFile(filename,1)
Set f2 = fso.createTextFile(filename2)
test = "        <Trooper class=""Grenadier"">"
Do Until f.AtEndOfStream
    line = f.ReadLine
    If test = line Then
      f2.WriteLine line
      For i = 1 to 18
        f2.WriteLine f.ReadLine
      Next
      line = f.ReadLine
		IF line = "                <MountedGun/>" then
			f2.WriteLine "                <MountedGun name=""default_camo_grenadier""/>"
		Else
			f2.WriteLine line
		End IF
    Else
      f2.WriteLine line
    End IF
Loop

f.Close
f2.Close
fso.DeleteFile filename
fso.MoveFile filename2, filename
Set f = fso.OpenTextFile(filename,1)
Set f2 = fso.createTextFile(filename2)
test = "        <Trooper class=""Undercover"">"
Do Until f.AtEndOfStream
    line = f.ReadLine
    If test = line Then
      f2.WriteLine line
      For i = 1 to 18
        f2.WriteLine f.ReadLine
      Next
      line = f.ReadLine
		IF line = "                <MountedGun/>" then
			f2.WriteLine "                <MountedGun name=""default_outfit_undercover""/>"
		Else
			f2.WriteLine line
		End IF
    Else
      f2.WriteLine line
    End IF
Loop

f.Close
f2.Close
fso.DeleteFile filename
fso.MoveFile filename2, filename
Set f = fso.OpenTextFile(filename,1)
Set f2 = fso.createTextFile(filename2)
test = "        <Trooper class=""BlackOps"">"
Do Until f.AtEndOfStream
    line = f.ReadLine
    If test = line Then
      f2.WriteLine line
      For i = 1 to 18
        f2.WriteLine f.ReadLine
      Next
      line = f.ReadLine
		IF line = "                <MountedGun/>" then
			f2.WriteLine "                <MountedGun name=""default_outfit_blackops""/>"
		Else
			f2.WriteLine line
		End IF
    Else
      f2.WriteLine line
    End IF
Loop

f.Close
f2.Close

MsgBox "The default Camos/Outfits have been equipped to your vanilla squad(s)",vbOKOnly,"Camouflage Selector Concept"


fso.DeleteFile filename
fso.MoveFile filename2, filename
