set shell = createobject("wscript.shell") 
strtext = inputbox("Text to type: ")
shell.sendkeys(strtext)