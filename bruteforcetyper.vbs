set shell = createobject("wscript.shell")
strnum = inputbox("Number of characters: ")
msgbox "Click OK to start"
for i=1 to strnum
shell.sendkeys("ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz0123456789:;?!""$%'(),-. ")
next