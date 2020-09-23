<div align="center">

## Find Free Socket in Winsock


</div>

### Description

Take a winsock control that is an array and use this function to find an availible sock/create a new one. Just thought I would submit something of mine to the site.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[NDW](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ndw.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ndw-find-free-socket-in-winsock__1-22373/archive/master.zip)





### Source Code

<BR><BR><font face="verdana" size="1">If you would like to create a server or other such program, you can initiate many connections over the same port by using this type of function. By creating an index of your Winsock you can create multiple connections over the same port...<BR><BR>Create a winsock control and set index to 0. Example of code use:<BR><BR>
iSock = FindUserSocket(sckMyWinsock)<BR>
sckMyWinsock(iSock).Accept RequestID
<BR><BR>Function FindUserSocket(sckWinsock As Winsock) As Long<BR>
 Dim iSock As Integer<BR>
 For iSock = 1 To sckUser.Count - 1<BR>
 If sckWinsock(iSock).State = sckClosed Then<BR>
 GoTo SockFound<BR>
 End If<BR>
 Next iSock<BR>
 GoTo MakeSock<BR>
 Exit Function<BR>
SockFound:<BR>
 FindFreeSocket = CLng(iSock)<BR>
 Exit Function<BR>
MakeSock:<BR>
 iSock = sckWinsock.Count<BR>
 Load sckWinsock(iSock)<BR>
 FindFreeSocket = CLng(iSock)<BR>
 Exit Function<BR>
End Function<BR></font>

