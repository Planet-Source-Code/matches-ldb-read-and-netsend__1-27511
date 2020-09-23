<div align="center">

## LDB Read and NetSend


</div>

### Description

I wanted a way to send messages to the users who are in an access database. This code opens and reads the ldb file for the database (that you specify on the form), reads the user names and sends a message....all without API calls. You can do a lot more with this obviously (I trimmed it down a lot for this forum).

You will need one text box for the source of the ldb file (named txtLocation), one command button, and one text box for your message (named txtMsg).
 
### More Info
 
To my knowledge, the NetSend Shell command only works with Windows NT.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matches](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matches.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matches-ldb-read-and-netsend__1-27511/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
'This simple code uses no api calls. it simply opens an ldb file that
'you choose (me.txtLocation) and places the the data in an Array
'For each variable in the array it shoots off a batch 'netsend'
'with a message you supply (me.txtmsg)
Dim strText As String
Dim vArray As Variant
Dim vParse As Variant
Dim iCount As Integer
   Open Me.txtLocation For Input As #1
   Input #1, strText
   For x = 1 To 25 'get rid of spaces-replace with single space
     strText = Replace(strText, " ", " ")
   Next
   strText = Replace(strText, " ", ",") 'replace all single spaces with a comma
   vArray = Split(strText, ",", -1) 'find all commas and split into an array
   For Each vParse In vArray
     iCount = iCount + 1 ' Get every other variable in Array (odd numbers)
     If iCount Mod 2 <> 0 Then
       RunBatch = Shell("net send " & _
           vParse & " "" " & txtmsg & """", vbNormalFocus)
     End If
   Next
   MsgBox "Message Sent"
   Close #1
End Sub
```

