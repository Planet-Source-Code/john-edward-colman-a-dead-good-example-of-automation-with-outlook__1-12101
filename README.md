<div align="center">

## A dead good example of Automation with Outlook


</div>

### Description

This demonstration program gives examples of how you can control Outlook using AUTOMATION to create mail, contacts and appointments.

You can adapt this code to create the other outlook items.
 
### More Info
 
I supplied the sub-routines but you'll need to create text boxes and command buttons to call them.

Creates Outlook objects

You'll still need to supply Outlook login info.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Edward Colman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-edward-colman.md)
**Level**          |Intermediate
**User Rating**    |4.4 (71 globes from 16 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-edward-colman-a-dead-good-example-of-automation-with-outlook__1-12101/archive/master.zip)

### API Declarations

```
Perform the following menu operations:
   Project > References
Put a check next to:
   "Microsoft Outlook 8.0 Object Library" or equivalent
```


### Source Code

```
Option Explicit
'Create an object to refererence the Outlook App.
'This is simular to a pointer and is declared in this way...
'...to allow early binding, making the code more efficient.
Private o1 As Outlook.Application
Private Sub Form_Load()
  'Create an instance of Outlook
  Set o1 = New Outlook.Application
End Sub
Private Sub Form_Terminate()
  'Comment out this line if you don't want to close Outlook
  o1.Quit
  'The next line frees up the memory used
  Set o1 = Nothing
End Sub
Private Sub CreateEmail(Recipient As String, Subject As String, Body As String, Attach As String)
  'Create a reference to a mail item
  Dim e1 As Outlook.MailItem
  'Create a new mail item
  Set e1 = o1.CreateItem(olMailItem)
  'Set a few of the many possible message parameters.
  e1.To = Recipient
  e1.Subject = Subject
  e1.Body = Body
  'This is how you add attatchments
  If Attach <> vbNullString Then
    e1.Attachments.Add Path
  End If
  'Commit the message
  e1.Send
  'Free up the space
  Set e1 = Nothing
End Sub
Private Sub CreateContact(Name As String, Nick As String, Email As String)
  'Create a reference to a Contact item
  Dim e1 As Outlook.ContactItem
  'Create a new contact item
  Set e1 = o1.CreateItem(olContactItem)
  'Set a few of the many possible contact parameters.
  e1.FullName = Name
  e1.NickName = Nick
  e1.Email1Address = Email
  'Commit the contact
  e1.Save
  'Free up the space
  Set e1 = Nothing
End Sub
Private Sub CreateAppointment(StartTime As Date, Endtime As Date, Subject As String, Location As String)
  'Create a reference to a Appointment item
  Dim e1 As Outlook.AppointmentItem
  'Create a new appointment item
  Set e1 = o1.CreateItem(olAppointmentItem)
  'Set a few of the many possible appointment parameters.
  e1.Start = StartTime
  e1.End = Endtime
  e1.Subject = Subject
  e1.Location = Location
  'If you want to set a list of recipients, do it like this
  'e1.Recipients.Add Name
  'Commit the appointment
  e1.Send
  'Free up the space
  Set e1 = Nothing
End Sub
```

