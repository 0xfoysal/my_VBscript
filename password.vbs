set shell = createobject("wscript.shell")
password = inputbox("Enter a password")
if password="123" then shell.run("notepad.exe") else msgbox("incorrect password")