' Outlook Mail Merge Attachment
' Modified by Timothy Law to deal with Sender's address
' 
' This script adds an attachment to all the emails that are currently
' in the Microsoft Office Outlook outbox. The script is tested with 
' Microsoft Outlook 2003, 2007, 2010 and 2013. 
' 
' Usage:
' 1.  Create your mail merge and be sure the messages are kept in the 
'     outbox (Work Offline).
' 2.  Execute (Double-Click) 'Outlook Mail Merge Attachment.vbs',
' 2a. select the attachment,
' 2b. the scripts now adds the selected file to all the emails in 
'     the outbox.
' 3.  Send the emails by working Online.
' 
' The emails are send by passing keystrokes. Please do not touch the keyboard or mouse while in process.
'
' For more information, visit http://omma.sourceforge.net or contact
' westerveld@users.sourceforge.net.
'
' Version 1.1.9 Beta, 26 October 2013
'
' Copyright (C) 2006-2013 Wouter Westerveld
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'

SubOutlookMailMergeAttachment

Sub SubOutlookMailMergeAttachment		
	' Script version
	strProgamName = "OMMA"
	strProgamVersion = "Outlook Mail Merge Attachment (v1.1.9 Beta)"
	
	' Set manual line-breaks in message box texts for windoes versions < 6.
	strBoxCr = vbCrLf
	On Error Resume Next
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 	
	For each System in SystemSet 				
		If System.Version >= 6 Then		
			strBoxCr = ""
		End If
		sWindowsVersion = System.Caption
	Next 
	On Error Goto 0
		
	' Welcome dialog
    strDialog = "Copyrighted by Wouter Westerveld (OMMA Version 1.1.9)"  & vbCrLf & "Modified by Tim Law"  & vbCrLf & "Lisa doing MailMerge with Attachment on Behalf of Susan" & vbCrLf & vbCrLf & _
    "Instructions before Running this Script (Click cancel here if you haven't done so):" & vbCrLf & _
    "1. Outlook > Send/Receive > Check WorkOffline" & vbCrLf & _
    "2. Be sure that your Outbox is empty" & vbCrLf & _
    "3. Do MailMerge from Word until Finish" & vbCrLf & _
    "4. Find all MailMerge mails in Outlook Outbox" & vbCrLf & _
    "5. Run this script again and hit OK" & vbCrLf & _
    "6. Wait for Attachment dialog, choose files or not" & vbCrLf & _
    "7. Check that your Outbox email has been set to be sent from Susan" & vbCrLf & vbCrLf & _
                "Click OK to Begin:"

	'''''''''''''''''''''''''''''''''''''''''''''''
	' Initialize, load objects, check
	'''''''''''''''''''''''''''''''''''''''''''''''

    If MsgBox(strDialog, vbOKCancel + vbInformation, strProgamName) = vbCancel Then
        ' fout  
	    Exit Sub                  
    End If     
        
    ' Outlook and Word Constants
    intFolderOutbox = 4
    msoFileDialogOpen = 1
    
    
    ' Load requied objects
    Set WshShell = WScript.CreateObject("WScript.Shell")	' Windows Shell
    Set ObjWord = CreateObject("Word.Application")      ' File Open dialog    
    Set ObjOlApp = CreateObject("Outlook.Application")      ' Outlook
    Set ns = ObjOlApp.GetNamespace("MAPI")                  ' Outlook
    Set box = ns.GetDefaultFolder(intFolderOutbox)          ' Outlook                  	    
         
    ' Check if we can detect problems in the outlook configuration
    sProblems = ""    
    sBuild = Left(ObjOlApp.Version, InStr(1, ObjOlApp.Version, ".") + 1)
    
    ' check spelling check just before sending
    On Error Resume Next
    r = WshShell.RegRead("HKCU\Software\Microsoft\Office\" & sBuild & "\Outlook\Options\Spelling\Check")    
    If Not(Err) And (r = 1) Then
    	sProblems = sProblems & _    	
    	"Your Outlook spell check is configured such that it gives a pop-up box when sending emails. Please disable " & strBoxCr & _
    	"the 'Always check spelling before sending' option in your Outlook. (ErrorCode = 101)" & vbCrLf &vbCrLf
    End If
    On Error Goto 0
    
	' For outlook 2000, 2002, 2003
	If sBuild = "9.0" Or sBuild = "10.0" Or sBuild = "11.0" Then
	
	    ' Check for word as email editor.
	    On Error Resume Next
		intEditorPrefs = WshShell.RegRead("HKCU\Software\Microsoft\Office\" & sBuild & "\Outlook\Options\Mail\EditorPreference")		
		If Not(Err) Then
			If intEditorPrefs = 131073 Or intEditorPrefs = 196609 Or intEditorPrefs = 65537 Then
				' HTML = 131072, HTML & Word To Edit = 131073, Rich Text = 196610, Rich Text & Word To Edit = 196609, Plain Text = 65536, Plain Text & Word To Edit = 65537			
				sProblems = sProblems & _			
				"Your Outlook is configured to use Word as email editor. Please change this to the internal outlook editor in " & strBoxCr & _
				"your outlook settings. (ErrorCode = 102)" & vbCrLf &vbCrLf				
			End If
		End If		
		On Error Goto 0
	End If

	If sProblems <> "" Then				    
		sProblems = "The OMMA script detected settings in your Outlook settings that need to be changed for the software to work." & vbCrLf & vbCrLf & sProblems
		MsgBox 	sProblems, vbExclamation, strProgamName			
		'fout
		Exit Sub
	End If

        
    ' Check if there are messages
    If box.Items.Count = 0 Then
        MsgBox "There are no messages in the Outbox.", vbExclamation, strProgamName           
       	' fout
       Exit Sub
    End If
    
    ' Give a warning if there already is an attachment
    If box.Items(1).Attachments.Count > 0 Then
        If MsgBox("The first email in your outbox has already " & box.Items(1).Attachments.Count & " attachment(s). Do you want to continue?", vbOKCancel + vbQuestion, strProgamName) = vbCancel Then
            ' fout  
		    Exit Sub            
        End If
    End If
        
    
        
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Ask user for Filenames, add atachment, and 
    ' Add attachment and save email
    '''''''''''''''''''''''''''''''''''''''''''''''     
        
        
    ' Ask user to open a file
    ' Select the attachment filename 
    
	ObjWord.ChangeFileOpenDirectory(CreateObject("Wscript.Shell").SpecialFolders("Desktop"))	
	ObjWord.FileDialog(msoFileDialogOpen).Title = "Attach file(s)..."
	ObjWord.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
	
	
	okEscape = False	
	If ObjWord.FileDialog(1).Show = -1 Then
		If ObjWord.FileDialog(1).SelectedItems.Count > 0 Then		
			okEscape = True
		End If 
	End If 
	
	If Not okEscape Then
	'	ObjWord.Quit
		MsgBox "Cancel was pressed, no attachments where added.", vbExclamation, strProgamName
	'	Exit Sub   	
	End If 
	
    WScript.Sleep(800)               
        
    ' Add the attachment to each email
    For Each Item In box.Items        
    	For Each objFile in ObjWord.FileDialog(1).SelectedItems
            Item.Attachments.Add(objFile)   
        Next
        ' Set Sender's email address
        Item.SentOnBehalfOfName = "SLamoreaux@solomonpage.com"
        Item.Save
    Next 

	ObjWord.Quit
 	
 	'''''''''''''''''''''''''''''''''''''''''''''''
 	' Send the emails using keystrokes
 	'''''''''''''''''''''''''''''''''''''''''''''''
 	
    For i = 1 to box.Items.Count
        
        ' Wait 5 extra seconds after 50 emails
        If (i Mod 50) = 0 Then
    		WScript.Sleep(5000)    	
        End If
        
        ' Open email
        Set objItem = box.Items(i)
		Set objInspector = objItem.GetInspector
		objInspector.Activate		
		WshShell.AppActivate(objInspector.Caption)		
		objInspector.Activate
	
		' wait upto 10 seconds until the window has focus		
		okEscape = False
		For j = 1 To 100
			WScript.Sleep(100)
			If (objInspector Is ObjOlApp.ActiveWindow) Then
				okEscape = True
				Exit For
			End	If
		Next
		If Not(okEscape) Then			        		
	        MsgBox "Internal error while opening email in outbox. Please read the how-to and the troubleshooting sections in the " & strBoxCr & "documentation. (ErrorCode = 103)", vbError, strProgamName
	       ' fout
	       Exit Sub			
		End If
		
		' send te email by typing ALT+S
		WshShell.SendKeys("%S")
						
		' wait upto 10 seconds for the sending to complete
		okEscape = False
		For j = 1 To 100
			WScript.Sleep(100)
			boolSent = False
			On Error Resume Next
			boolSent = objItem.Sent
			If Err Then
				boolSent = True
			End	If
			On Error Goto 0
			If boolSent Then
				okEscape = True
				Exit For
			End	If
		Next						
		If Not(okEscape) Then					
			' Error			       
	        MsgBox "Internal error while sending email. Perhaps the email window was not activated. Please read the how-to and " & strBoxCr & "the troubleshooting sections in the documentation. (ErrorCode = 104)", vbExclamation, strProgamName
	       ' fout
	       Exit Sub						
		End If
		
		
	
    Next 
 
    ' Finished    
'    strDialog = "Successfully added the attachment to " & box.Items.Count & " emails." & vbCrLf & vbCrLf & _    	
'    	"OMMA is free software, please let the author know whether OMMA worked properly. " &strBoxCr & _
'    	"Did you already fill the feedback form?" & vbCrLf & vbCrLf & _ 
'    	"Answer 'No' will open the feedback form in your browser."  & vbCrLf & _  
'    	"Answer 'Yes' just exit the script." 
    	
    MsgBox("Successfully proccessed " & box.Items.Count & " emails.")
    
'    If MsgBox(strDialog, vbYesNo + vbInformation, strProgamName) = vbNo Then
'		WshShell.Run "http://omma.sourceforge.net/feedback.php?worksok=yes&verOmma=" & escape(strProgamVersion) & "&verWindows=" & escape(sWindowsVersion) & "&verOutlook=" & escape(sBuild)
'    End If         
    
End Sub


