set shell = createobject("wscript.shell") 
strtext = inputbox("Text to type: ")
for i=1 To Len(strtext)
    shell.sendkeys(Mid(strtext,i,1))
    wscript.sleep(10)
next