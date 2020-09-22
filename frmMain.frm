VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "DataFX"
   ClientHeight    =   3195
   ClientLeft      =   3060
   ClientTop       =   4080
   ClientWidth     =   2625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2625
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   1800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
      Picture         =   "frmMain.frx":08CA
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   1323
      BandCount       =   2
      _CBWidth        =   2625
      _CBHeight       =   750
      _Version        =   "6.7.8988"
      Child1          =   "tbrMain"
      MinHeight1      =   330
      Width1          =   1605
      NewRow1         =   0   'False
      Child2          =   "picTableName"
      MinHeight2      =   330
      Width2          =   4005
      NewRow2         =   -1  'True
      Begin VB.PictureBox picTableName 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   2370
         TabIndex        =   3
         Top             =   390
         Width           =   2370
         Begin VB.ComboBox cmbTableName 
            Height          =   315
            Left            =   0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   390
         End
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open a Database"
               Object.Tag             =   "Open"
               ImageKey        =   "Open"
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy Field Name"
               Object.Tag             =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               Object.Tag             =   "Find"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "FindNext"
               Object.ToolTipText     =   "Find Next"
               Object.Tag             =   "FindNext"
               ImageKey        =   "FindNext"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "DataSheet"
               Object.ToolTipText     =   "DataSheet View"
               Object.Tag             =   "DataSheet"
               ImageKey        =   "DataSheet"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Front"
               Object.ToolTipText     =   "Stay On Top"
               Object.Tag             =   "Front"
               ImageKey        =   "Front"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B7A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":424E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43AA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A7E
            Key             =   "Front"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B92
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CB2
            Key             =   "FindNext"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E12
            Key             =   "DataSheet"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4128
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileName 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Developed by Chris George
' 03/19/02
'
' Feel free to use any code here in any of your projects
' If you like this program, all I ask is that you please leave a comment at planet-source-code.com


' Disclaimer:
' I am not responsable for any damages to any file you open with this program

' Note: This program makes use of the DBGrid32.ocx
'       If you don't have it, view the readme file to find out how to get it.
'       Don't save if you loaded this program without the ocx because you will
'       lose the grid control on frmGrid.  Exit VB without saving and get the
'       control.  Then come back in and it should work fine.


Option Explicit      'all variables must be defined

Public DBPath As String                 'Database path (DAO) / Connection String (ADO)
Public WindowPos As Integer             'Indicates if window is to stay on top
Public LastSearch As String             'Last string searched for
Public WholeWord As Integer             'Indicates if searching for a whole word or part
Public DataSourceType As DSTypeEnum     'Indicates if using ADO or DAO
Public AskAgain As Integer              'Indicates if the type of datasource should be asked every time you press open

'This enumeration is used to indicate the type of datasource to use
Public Enum DSTypeEnum
    dsADO = 0
    dsDAO = 1
End Enum

Private Sub cbrMain_HeightChanged(ByVal NewHeight As Single)
    Me.Form_Resize
End Sub

Private Sub cmbTableName_Click()
    'get fields from the table
    Dim i As Integer
    Dim j As Integer
    'DAO Objects
    Dim MyDAODB As DAO.Database
    Dim MyDAORecSet As DAO.Recordset
    'ADO Objects
    Dim MyADODB As New ADODB.Connection
    Dim MyADORecSet As New ADODB.Recordset
    
    On Error GoTo Err_Handle
    
    If Me.DataSourceType = dsDAO Then
        'open the database
        Set MyDAODB = OpenDatabase(DBPath)
        Set MyDAORecSet = MyDAODB.OpenRecordset(Me.cmbTableName.Text)
    Else
        'open the datasource
        MyADODB.Open DBPath
        MyADORecSet.Open "Select * from [" & Me.cmbTableName.Text & "]", MyADODB, adOpenDynamic, adLockBatchOptimistic
    End If
    
    'clear the list
    Me.lvwFields.ListItems.Clear
    
    'clear the column headers
    Me.lvwFields.ColumnHeaders.Clear
    
    If Me.DataSourceType = dsDAO Then
        Me.lvwFields.ColumnHeaders.Add , , "Field Name"
        Me.lvwFields.ColumnHeaders.Add , , "Description"
    Else
        Me.lvwFields.ColumnHeaders.Add , , "Field Name"
    End If
    
    'get all of the field properties
    If Me.DataSourceType = dsDAO Then
        'DAO
        For i = 0 To MyDAORecSet.Fields(0).Properties.Count - 1
            If MyDAORecSet.Fields(0).Properties(i).Name <> "Name" And MyDAORecSet.Fields(0).Properties(i).Name <> "Value" And MyDAORecSet.Fields(0).Properties(i).Name <> "Description" Then
                Me.lvwFields.ColumnHeaders.Add , , MyDAORecSet.Fields(0).Properties(i).Name
            End If
        Next
    Else
        'ADO
        For i = 0 To MyADORecSet.Fields(0).Properties.Count - 1
            Me.lvwFields.ColumnHeaders.Add , , MyADORecSet.Fields(0).Properties(i).Name
        Next
    End If
    
    'list all of the fields
    
    On Error Resume Next
    If Me.DataSourceType = dsDAO Then
        For j = 0 To MyDAORecSet.Fields.Count
            'DAO
            Me.lvwFields.ListItems.Add , "K" & j, MyDAORecSet.Fields(j).Name
            Me.lvwFields.ListItems("K" & j).SubItems(1) = MyDAORecSet.Fields(j).Properties("Description").Value
            For i = 2 To Me.lvwFields.ColumnHeaders.Count - 1
                On Error Resume Next
                Me.lvwFields.ListItems("K" & j).SubItems(i) = MyDAORecSet.Fields(Me.lvwFields.ListItems("K" & j).Text).Properties(Me.lvwFields.ColumnHeaders(i + 1)).Value
            Next
        Next
    Else
        For j = 0 To MyADORecSet.Fields.Count
            'ADO
            Me.lvwFields.ListItems.Add , "K" & j, MyADORecSet.Fields(j).Name
            For i = 1 To Me.lvwFields.ColumnHeaders.Count - 1
                On Error Resume Next
                Me.lvwFields.ListItems("K" & j).SubItems(i) = MyADORecSet.Fields(Me.lvwFields.ListItems("K" & j).Text).Properties(i - 1).Value
            Next
        Next
    End If

    
    Me.lvwFields.BackColor = vbWindowBackground
    'release database objects
    Set MyDAORecSet = Nothing
    Set MyDAODB = Nothing
    Set MyADORecSet = Nothing
    Set MyADODB = Nothing
        
    Exit Sub
    
Err_Handle:
    MsgBox "Error Opening Database: " & Error, vbExclamation + vbOKOnly
    'release database objects
    Set MyDAORecSet = Nothing
    Set MyDAODB = Nothing
    Set MyADORecSet = Nothing
    Set MyADODB = Nothing
    Me.lvwFields.ListItems.Clear
    
End Sub

Public Function FileExists(strPath As String) As Boolean
    'checks to see if a file exists
    FileExists = Not (Dir(strPath) = "")
End Function

Private Sub Form_Load()
    On Error Resume Next
    'resize the combo box
    Me.cmbTableName.Width = Me.picTableName.Width
    'load the recent file list
    Me.LoadRecentFiles
    'show the form
    Me.Visible = True
    
    'default to DAO
    Me.DataSourceType = dsDAO
    'show the datatype dialog window
    Me.ShowDataTypeDialog
    
Cancel:
End Sub

Public Sub Form_Resize()
    'resize the controls when the form is resized
    On Error Resume Next
    Me.lvwFields.Top = Me.cbrMain.Top + Me.cbrMain.Height + 25
    Me.lvwFields.Width = Me.Width - 100
    Me.lvwFields.Height = Me.Height - Me.lvwFields.Top - sbrMain.Height - 400
End Sub

Public Sub LoadRecentFiles()
    Dim FileNum As Integer
    Dim i As Integer
    Dim Index As Integer
    Dim strInput As String
    Dim strArray() As String
    On Error GoTo Err_Handle
    
    'get a free file number
    FileNum = FreeFile
    Index = 0
    'This sub will go get the most recent files
    If FileExists(App.Path & "\rfiles") Then
        Open App.Path & "\rfiles" For Input As #FileNum
            Do
                'loop through the file and read in each line
                Line Input #FileNum, strInput
                'get the next menu to load
                If Index > 0 Then Load Me.mnuFileName(Index)
                strArray = Split(strInput, "|")
                'set the caption to the filename
                Me.mnuFileName(Index).Caption = Trim(strArray(0))
                Me.mnuFileName(Index).Tag = Trim(strArray(1))
                Index = Index + 1
            Loop Until EOF(FileNum)
        Close #FileNum
    End If
    
Err_Handle:
    Close
End Sub

Private Sub lvwFields_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Me.lvwFields.Sorted = True
    If Me.lvwFields.SortKey = ColumnHeader.Index - 1 Then
        If Me.lvwFields.SortOrder = lvwAscending Then
            Me.lvwFields.SortOrder = lvwDescending
        Else
            Me.lvwFields.SortOrder = lvwAscending
        End If
    End If
    Me.lvwFields.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwFields_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.sbrMain.Panels(1).Text = ""
End Sub

Private Sub lvwFields_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectCopy
    Data.SetData Me.lvwFields.SelectedItem.Text
End Sub

Public Sub OpenADODB(Optional ConnectionString As String)
    'This sub opens an ODBC Datasource selected by the user
    Dim MyDB As New ADODB.Connection
    Dim MyRecSet As New ADODB.Recordset
    Dim i As Integer
    Dim RecCount As Integer
    Dim Count As Integer
    Dim NewFile As Boolean
    Dim DSName As String
    On Error GoTo Cancel

    'if a connectionstring was not passed then it is a new file
    If ConnectionString = "" Then NewFile = True

    'set the connection string
    MyDB.ConnectionString = ConnectionString
    'show the user the open data source dialog
    MyDB.Properties("Prompt") = adPromptComplete
    'use client side cursor
    MyDB.CursorLocation = adUseClient
    'open the datasource
    MyDB.Open
    
    On Error GoTo Err_Handle
    
    'set the global connection string
    DBPath = MyDB.ConnectionString
    
    'clear the combo box
    Me.cmbTableName.Clear
    
    'open the table schema to get table names
    Set MyRecSet = MyDB.OpenSchema(adSchemaTables)
    'make sure there are tables
    If Not MyRecSet.EOF Then
        'get a count
        MyRecSet.MoveLast
        RecCount = MyRecSet.RecordCount
        MyRecSet.MoveFirst
        
        'load the combo box with table names
        For i = 1 To RecCount
            If UCase(Mid(MyRecSet!Table_Name, 1, 4)) <> "MSYS" And UCase(Mid(MyRecSet!Table_Name, 1, 1)) <> "~" Then
                Me.cmbTableName.AddItem MyRecSet!Table_Name
            End If
            MyRecSet.MoveNext
        Next
    End If
    
    'not all providers can provide views so make sure to have error handling
    On Error Resume Next
    
    'open the query schema to get query names
    
    Set MyRecSet = MyDB.OpenSchema(adSchemaViews)
    'make sure there are tables
    If Not MyRecSet.EOF Then
        'get a count
        MyRecSet.MoveLast
        RecCount = MyRecSet.RecordCount
        MyRecSet.MoveFirst
        
        'load the combo box with table names
        For i = 1 To RecCount
            If UCase(Mid(MyRecSet!View_Name, 1, 4)) <> "MSYS" And UCase(Mid(MyRecSet!View_Name, 1, 1)) <> "~" Then
                Me.cmbTableName.AddItem MyRecSet!View_Name
            End If
            MyRecSet.MoveNext
        Next
    End If
    
    'select the first item in the list
    Me.cmbTableName.ListIndex = 0
    
    'set the caption
    Me.Caption = "DataFX [" & MyDB.Properties(0).Value & "]"
        
    If NewFile = True Then
        'load a new menu item for the file
        Count = Me.mnuFileName.Count
        If Count < 6 Then
            'load a new menu item if the first menu is already filled with a file name
            If Me.mnuFileName(0).Caption <> "-" Then
                Load mnuFileName(Count)
            Else
                Count = Count - 1
            End If

            For i = Count To 1 Step -1
                mnuFileName(i).Caption = mnuFileName(i - 1).Caption
                mnuFileName(i).Tag = mnuFileName(i - 1).Tag
            Next
            mnuFileName(0).Caption = "ADO - " & MyDB.Properties(0).Value
            mnuFileName(0).Tag = "ADO - " & MyDB.ConnectionString
        Else
            For i = 5 To 1 Step -1
                mnuFileName(i).Caption = mnuFileName(i - 1).Caption
                mnuFileName(i).Tag = mnuFileName(i - 1).Tag
            Next
            mnuFileName(0).Caption = "ADO - " & MyDB.Properties(0).Value
            mnuFileName(0).Tag = "ADO - " & MyDB.ConnectionString
        End If
    End If
    
    'save the recent file list
    Me.SaveRecentFiles
    
    'enable buttons on the toolbar
    Me.tbrMain.Buttons(2).Enabled = True
    Me.tbrMain.Buttons(3).Enabled = True
    Me.tbrMain.Buttons(4).Enabled = True
    Me.tbrMain.Buttons(5).Enabled = True
    
    'release data objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing

    Exit Sub

Err_Handle:
    MsgBox "Error Opening DataSource: " & Error, vbExclamation + vbOKOnly
        
Cancel:
    'release data objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing
    
End Sub

Public Sub OpenDAODB(Optional FileName As String)
    'This sub opens a database selected by the user
    Dim MyDB As DAO.Database
    Dim MyRecSet As DAO.Recordset
    Dim i As Integer
    Dim Count As Integer
    Dim NewFile As Boolean
    
    On Error GoTo Cancel
    If FileName = "" Then
        'set the filter for the common dialog window for access database
        cdMain.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        cdMain.Filter = "Microsoft Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
        cdMain.FilterIndex = 0
        'get the filename to open
        cdMain.ShowOpen
        FileName = cdMain.FileName
        NewFile = True
    End If
    
    On Error GoTo Err_Handle
    
    'make sure the file exists
    If FileExists(FileName) = False Then
        MsgBox FileName & " does not exist.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    'set the global dbpath
    DBPath = FileName
    
    'attempt to open the database
    Set MyDB = OpenDatabase(FileName)
    
    'clear the combo box
    Me.cmbTableName.Clear
    
    'load the combo box with table names
    For i = 0 To MyDB.TableDefs.Count - 1
        If UCase(Mid(MyDB.TableDefs(i).Name, 1, 4)) <> "MSYS" And UCase(Mid(MyDB.TableDefs(i).Name, 1, 1)) <> "~" Then
            Me.cmbTableName.AddItem MyDB.TableDefs(i).Name
        End If
    Next
    
    'load the combo box with query names
    For i = 0 To MyDB.QueryDefs.Count - 1
        If UCase(Mid(MyDB.QueryDefs(i).Name, 1, 4)) <> "MSYS" And UCase(Mid(MyDB.QueryDefs(i).Name, 1, 1)) <> "~" Then
            Me.cmbTableName.AddItem MyDB.QueryDefs(i).Name
        End If
    Next
    
    'select the first item in the list
    Me.cmbTableName.ListIndex = 0
    
    'set the caption
    Me.Caption = "DataFX [" & Dir(FileName) & "]"
    
    If NewFile = True Then
        'load a new menu item for the file
        Count = Me.mnuFileName.Count
        If Count < 6 Then
            'load a new menu item if the first menu is already filled with a file name
            If Me.mnuFileName(0).Caption <> "-" Then
                Load mnuFileName(Count)
            Else
                Count = Count - 1
            End If
            
            For i = Count To 1 Step -1
                mnuFileName(i).Caption = mnuFileName(i - 1).Caption
                mnuFileName(i).Tag = mnuFileName(i - 1).Tag
            Next
            mnuFileName(0).Caption = "DAO - " & FileName
            mnuFileName(0).Tag = "DAO"
        Else
            For i = 5 To 1 Step -1
                mnuFileName(i).Caption = mnuFileName(i - 1).Caption
                mnuFileName(i).Tag = mnuFileName(i - 1).Tag
            Next
            mnuFileName(0).Caption = "DAO - " & FileName
            mnuFileName(0).Tag = "DAO"
        End If
    End If
    'save the recent file list
    Me.SaveRecentFiles
    
    'enable buttons on the toolbar
    Me.tbrMain.Buttons(2).Enabled = True
    Me.tbrMain.Buttons(3).Enabled = True
    Me.tbrMain.Buttons(4).Enabled = True
    Me.tbrMain.Buttons(5).Enabled = True
    
    'release database objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing
        
    Exit Sub
    
Err_Handle:
    MsgBox "Error opening database: " & Error, vbExclamation + vbOKOnly

Cancel:
    'release database objects
    Set MyRecSet = Nothing
    Set MyDB = Nothing

End Sub


Private Sub mnuFileName_Click(Index As Integer)
    'if the user clicks one of the files on the menu then
    'set it to DAO
    Dim myDataType As String
    Dim ConnString As String
    
    'get the type of data (DAO/ADO)
    myDataType = Mid(Me.mnuFileName(Index).Tag, 1, 3)
    'if it is ADO then get the connection string
    If myDataType = "ADO" Then
        ConnString = Mid(Me.mnuFileName(Index).Tag, 7, Len(Me.mnuFileName(Index).Tag) - 6)
        Me.DataSourceType = dsADO
    Else
        Me.DataSourceType = dsDAO
    End If
    
    If myDataType = "DAO" Then
        If InStr(1, Me.mnuFileName(Index).Caption, ":") > 0 Then
            Me.OpenDAODB Mid(Me.mnuFileName(Index).Caption, 7, Len(Me.mnuFileName(Index).Caption) - 6)
        Else
            Me.OpenDAODB App.Path & "\" & Mid(Me.mnuFileName(Index).Caption, 7, Len(Me.mnuFileName(Index).Caption) - 6)
        End If
    Else
        Me.OpenADODB ConnString
    End If
End Sub

Private Sub picTableName_Resize()
    On Error Resume Next
    'resize the combo box when the container is resized
    Me.cmbTableName.Width = Me.picTableName.Width
End Sub

Public Sub SaveRecentFiles()
    Dim FileNum As Integer
    Dim i As Integer
    Dim Index As Integer
    Dim strInput As String
    On Error GoTo Err_Handle
    
    'get a free file number
    FileNum = FreeFile
    
    'This sub will save the most recent files

    Open App.Path & "\rfiles" For Output As #FileNum
        For i = 0 To Me.mnuFileName.Count - 1
            Print #FileNum, Me.mnuFileName(i).Caption & "|" & Me.mnuFileName(i).Tag
        Next
    Close #FileNum
    
    Exit Sub
    
Err_Handle:
    Close
End Sub

Public Sub ShowDataTypeDialog()
    'show the open dialog
    frmPrompt.Show vbModal, Me
    With frmPrompt
        If .pCancel = True Then
            Unload frmPrompt
            Exit Sub
        End If
        
        'store the value of the don't show checkbox
        Me.AskAgain = .chkDontAsk
        
        If .optDataSource(0).Value = True Then
            Me.DataSourceType = dsADO
            Me.OpenADODB
        Else
            Me.DataSourceType = dsDAO
            Me.OpenDAODB
        End If
        'unload the prompt window
        Unload frmPrompt
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lRetVal As Long
    Dim SearchFor As String
    Dim SearchIndex As Integer
    
    On Error GoTo Err_Handle
    
    Select Case Button.Tag
        Case "Open"
            'if the user has not requested to not be asked again then
            'show the data type dialog window,
            'otherwise show the corresponding open dialog window
            
            If Me.AskAgain = vbUnchecked Then
                'show the open dialog window
                Me.ShowDataTypeDialog
            Else
                'show the open dialog window
                If Me.DataSourceType = dsDAO Then
                    Me.OpenDAODB
                Else
                    Me.OpenADODB
                End If
            End If
        Case "Copy"
            'copy the selected field name
            If Me.lvwFields.ListItems.Count > 0 Then
                Clipboard.Clear
                Clipboard.SetText Me.lvwFields.SelectedItem.Text
            End If
        Case "Find"
            If Me.lvwFields.ListItems.Count < 0 Then Exit Sub
            'put the last item searched for in the search box
            'and then select it so the user can quickly delete it
            frmSearch.txtSearch.Text = LastSearch
            frmSearch.txtSearch.SelStart = 0
            frmSearch.txtSearch.SelLength = Len(LastSearch)
            frmSearch.chkWholeWord = WholeWord
            'show the dialog window to get what the user is searching for
            frmSearch.Show vbModal, Me
            With frmSearch
                'if the user clicked cancel then exit the sub
                If .pCancel = True Then Exit Sub
                LastSearch = .txtSearch
                WholeWord = .chkWholeWord.Value
                SearchFor = .txtSearch.Text
                'if the user doesn't have anything selected then make it 1
                'otherwise start from the next item on the list with the search
                'so the user can find the next item
                If Me.lvwFields.SelectedItem Is Nothing Then
                    SearchIndex = 1
                Else
                    SearchIndex = Me.lvwFields.SelectedItem.Index + 1
                    If SearchIndex > Me.lvwFields.ListItems.Count Then SearchIndex = 1
                End If
                'if the user checked find whole word only then only search for the whole word, else do a partial search
                If .chkWholeWord.Value = vbChecked Then
                    Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwWhole)
                Else
                    Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwPartial)
                End If
                'if nothing was found, tell the user, otherwise display it
                If Me.lvwFields.SelectedItem Is Nothing Then
                    MsgBox "String was not found.", vbInformation + vbOKOnly
                Else
                    Me.lvwFields.SelectedItem.EnsureVisible
                End If
            End With
        Case "FindNext"
            If Me.lvwFields.ListItems.Count < 0 Then Exit Sub
            'if the user doesn't have anything selected then make it 1
            'otherwise start from the next item on the list with the search
            'so the user can find the next item
            If Me.lvwFields.SelectedItem Is Nothing Then
                SearchIndex = 1
            Else
                SearchIndex = Me.lvwFields.SelectedItem.Index + 1
                If SearchIndex > Me.lvwFields.ListItems.Count Then SearchIndex = 1
            End If
            SearchFor = LastSearch
            'if the user checked find whole word only then only search for the whole word, else do a partial search
            If WholeWord = vbChecked Then
                Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwWhole)
            Else
                Set Me.lvwFields.SelectedItem = Me.lvwFields.FindItem(SearchFor, lvwText, SearchIndex, lvwPartial)
            End If
            'if nothing was found, tell the user, otherwise display it
            If Me.lvwFields.SelectedItem Is Nothing Then
                MsgBox "String was not found.", vbInformation + vbOKOnly
            Else
                Me.lvwFields.SelectedItem.EnsureVisible
            End If
        Case "Front"
            If Me.WindowPos = HWND_TOPMOST Then
                lRetVal = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
                Me.WindowPos = HWND_NOTOPMOST
                Button.Image = "Front"
                Button.ToolTipText = "Stay on Top"
            Else
                lRetVal = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
                Me.WindowPos = HWND_TOPMOST
                Button.Image = "Back"
                Button.ToolTipText = "Don't Stay on Top"
            End If
            Me.cbrMain.Refresh
        Case "DataSheet"
            'make sure a table is selected
            If Trim(Me.cmbTableName) <> "" Then
                'create a new datasheet
                Dim NewSheet As New frmGrid
                If Me.DataSourceType = dsDAO Then
                    NewSheet.DataSourceType = dsDAO
                    NewSheet.MyDAODB.DatabaseName = Me.DBPath
                    NewSheet.MyDAODB.RecordSource = Me.cmbTableName.Text
                    NewSheet.MyDAODB.Refresh
                    NewSheet.sbrMain.Panels(1).Text = NewSheet.MyDAODB.Recordset.RecordCount & " Records"
                    NewSheet.DAODBGrid.Visible = True
                    NewSheet.TableName = "Select * from [" & Me.cmbTableName.Text & "]"
                    NewSheet.Caption = "DataSheet - " & Me.cmbTableName.Text
                    NewSheet.Visible = True
                Else
                    NewSheet.DataSourceType = dsADO
                    NewSheet.Visible = True
                    NewSheet.ADODBGrid.Visible = True
                    NewSheet.MyADODB.ConnectionString = Me.DBPath
                    NewSheet.MyADODB.RecordSource = "Select * from [" & Me.cmbTableName.Text & "]"
                    NewSheet.MyADODB.Refresh
                    NewSheet.ADODBGrid.Refresh
                    NewSheet.sbrMain.Panels(1).Text = NewSheet.MyADODB.Recordset.RecordCount & " Records"
                    NewSheet.TableName = "Select * from [" & Me.cmbTableName.Text & "]"
                    NewSheet.Caption = "DataSheet - " & Me.cmbTableName.Text
                    NewSheet.MyADODB.Refresh
                End If
            End If
    End Select
    
    Exit Sub
    
Err_Handle:
    MsgBox "Error: " & Error, vbExclamation + vbOKOnly
    On Error Resume Next
    Unload NewSheet
End Sub

Private Sub tbrMain_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Me.PopupMenu Me.mnuFile, , Button.Left + Me.tbrMain.Left + Me.cbrMain.Left, Button.Top + Button.Height + Me.tbrMain.Top + Me.cbrMain.Top
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set the status bar caption to the tooltip text of the button
    Dim i As Integer
    
    For i = 1 To Me.tbrMain.Buttons.Count
        With Me.tbrMain.Buttons(i)
            If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height Then
                Me.sbrMain.Panels(1).Text = .ToolTipText
                Exit Sub
            End If
        End With
    Next
    Me.sbrMain.Panels(1).Text = ""
End Sub

