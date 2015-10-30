

'Check File
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim cnf
Dim cnf2
Dim dir As String
Dim dir2 As String
Set cnf = CreateObject("Scripting.FileSystemObject")
Set cnf2 = CreateObject("Scripting.FileSystemObject")
dir = "S:\he\ADMINISTRATION\COMPLAINTS\ComplaintDB\" & Me.parcelBox.Value
dir2 = "S:\he\ADMINISTRATION\COMPLAINTS\ComplaintDB\" & Me.parcelBox.Value & "\" & Me.ComboBox1.Value
If Not cnf.FolderExists(dir) Then
cnf.CreateFolder (dir)
If Not cnf2.FolderExists(dir2) Then
cnf2.CreateFolder (dir2)


End If
End If

'Screenshot Userform2
''''''''''''''''


Dim myPath As String
    myPath = dir2
    DoEvents
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    DoEvents
    Workbooks.Add
    Application.Wait Now + TimeValue("00:00:01")
    ActiveSheet.PasteSpecial Format:="Bitmap", Link:=False, DisplayAsIcon:=False
    ActiveSheet.Range("A1").Select
    ActiveSheet.PageSetup.Orientation = xlLandscape
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
            myPath & UserForm2.ComboBox3.Value & ".pdf", Quality _
            :=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    'ActiveSheet.Export FileName:="C:\Users\he_wooster\Desktop\ClipBoardToPic.jpg", FilterName:="jpg"
    'UserForm2.Hide
    '    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWorkbook.Close False
    '    OptionButton2.Value = False
    '    Me.Show vbModeless


   Public newpath As String
    newpath = myPath & UserForm2.ComboBox3.Value & ".pdf"

Me.Hide
UserForm3.Show
