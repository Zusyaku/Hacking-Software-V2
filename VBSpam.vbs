set shell = createobject ("wscript.shell") 
strtext = inputbox ("Please Enter The Message You Want To Be Spammed") 
strtimes = inputbox ("How Many Threads Would You Like To Spam It?") 
strspeed = inputbox ("How Fast Do You Want To Send The Message? (1000 = 1 per sec, 100 = 10 per sec etc)") 
strtimeneed = inputbox ("How Many Seconds Do You Want To Delay The Spammer So You Can Get To You Victims Inbox?") 
If not isnumeric (strtimes & strspeed & strtimeneed) then 
msgbox "You entered something else then a number on Times, Speed and/or Time need. shutting down" 
wscript.quit 
End If 
strtimeneed2 = strtimeneed * 1000 
do 
msgbox "You have " & strtimeneed & " seconds to get to your input area where you are going to spam." 
wscript.sleep strtimeneed2 
shell.sendkeys ("Spambot activated" & "{enter}") 
for i=0 to strtimes 
shell.sendkeys (strtext & "{enter}") 
wscript.sleep strspeed 
Next 
shell.sendkeys ("Spambot deactivated" & "{enter}") 
wscript.sleep strspeed * strtimes / 10 
returnvalue=MsgBox ("Would You Like To Spam The Victim Again With The Same Information?",36) 
If returnvalue=6 Then 
Msgbox "Ok Spambot will activate again -Coded by: GrIFt3r_EdITz.inc" 
End If 
If returnvalue=7 Then 
msgbox "Deactivating Spammer" 
wscript.quit 
End IF 
loop 
