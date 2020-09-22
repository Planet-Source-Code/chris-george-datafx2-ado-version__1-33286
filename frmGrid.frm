VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmGrid 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Data"
   ClientHeight    =   5775
   ClientLeft      =   10650
   ClientTop       =   5685
   ClientWidth     =   6000
   Icon            =   "frmGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDBGrid.DBGrid DAODBGrid 
      Bindings        =   "frmGrid.frx":000C
      Height          =   1695
      Left            =   0
      OleObjectBlob   =   "frmGrid.frx":0022
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid ADODBGrid 
      Bindings        =   "frmGrid.frx":09F7
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc MyADODB 
      Height          =   330
      Left            =   120
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ADO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5520
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7514
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   3240
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0A0D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0B25
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0C3D
            Key             =   "First"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0D55
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0E6D
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0F85
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":109D
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":11B5
            Key             =   "Ascending"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":12CD
            Key             =   "Descending"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":13E5
            Key             =   "Front"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1505
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1625
            Key             =   "DeleteFilter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1739
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   6000
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "tbrMain"
      MinHeight1      =   330
      Width1          =   3840
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "First"
               Object.ToolTipText     =   "Go to the First Record"
               Object.Tag             =   "First"
               ImageKey        =   "First"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Previous"
               Object.ToolTipText     =   "Go to the Previous Record"
               Object.Tag             =   "Previous"
               ImageKey        =   "Previous"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Next"
               Object.ToolTipText     =   "Go to the Next Record"
               Object.Tag             =   "Next"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Last"
               Object.ToolTipText     =   "Go to the Last Record"
               Object.Tag             =   "Last"
               ImageKey        =   "Last"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Add"
               Object.ToolTipText     =   "Add a New Record"
               Object.Tag             =   "Add"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete the Selected Record"
               Object.Tag             =   "Delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Filter"
               Object.ToolTipText     =   "Filter by SQL"
               Object.Tag             =   "Filter"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DeleteFilter"
               Object.ToolTipText     =   "Remove the Filter"
               Object.Tag             =   "DeleteFilter"
               ImageKey        =   "DeleteFilter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Refresh"
               Object.Tag             =   "Refresh"
               ImageKey        =   "Refresh"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Front"
               Object.ToolTipText     =   "Stay on Top"
               Object.Tag             =   "Front"
               ImageKey        =   "Front"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Data MyDAODB 
      Caption         =   "DAO"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2460
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TableName As String
Public WindowPos As Integer
Public DataSourceType As DSTypeEnum     'Indicates if using ADO or DAO

Private Sub ADODBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Cancel
    'show the current record and the record count in the status bar
    If Me.DataSourceType = dsADO Then
        Me.sbrMain.Panels(1).Text = "Record " & Me.ADODBGrid.Row + 1 & " of " & Me.MyADODB.Recordset.RecordCount
    End If
    Exit Sub
    
Cancel:
    Me.sbrMain.Panels(2).Text = ""
    
End Sub

Private Sub DAODBGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Cancel
    'show the field description in the status bar
    If Me.DataSourceType = dsDAO Then Me.sbrMain.Panels(2).Text = MyDAODB.Recordset.Fields(Me.DAODBGrid.Col).Properties("Description").Value
    Exit Sub
    
Cancel:
    Me.sbrMain.Panels(2).Text = ""
    
End Sub

Private Sub DAODBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Cancel
    'show the current record and the record count in the status bar
    If Me.DataSourceType = dsDAO Then
        Me.sbrMain.Panels(1).Text = "Record " & Me.DAODBGrid.Row + 1 & " of " & Me.MyDAODB.Recordset.RecordCount
        'show the field description in the status bar
        Me.sbrMain.Panels(2).Text = MyDAODB.Recordset.Fields(Me.DAODBGrid.Col).Properties("Description").Value
    End If
    Exit Sub
    
Cancel:
    Me.sbrMain.Panels(2).Text = ""
    
End Sub

Private Sub ADODBGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.sbrMain.Panels(2).Text = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'resize the grid to fit on the form
    With Me.DAODBGrid
        .Width = Me.Width - 150
        .Height = Me.Height - .Top - Me.sbrMain.Height - 400
    End With
    With Me.ADODBGrid
        .Width = Me.Width - 150
        .Height = Me.Height - .Top - Me.sbrMain.Height - 400
    End With

End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo Err_Handle
    
    Select Case Button.Tag
        Case "First"
            If Me.DataSourceType = dsDAO Then
                Me.MyDAODB.Recordset.MoveFirst
            Else
                Me.MyADODB.Recordset.MoveFirst
            End If
        Case "Next"
            If Me.DataSourceType = dsDAO Then
                If Not Me.MyDAODB.Recordset.EOF Then Me.MyDAODB.Recordset.MoveNext
            Else
                If Not Me.MyADODB.Recordset.EOF Then Me.MyADODB.Recordset.MoveNext
            End If
        Case "Previous"
            If Me.DataSourceType = dsDAO Then
                If Not Me.MyDAODB.Recordset.BOF Then Me.MyDAODB.Recordset.MovePrevious
            Else
                If Not Me.MyADODB.Recordset.BOF Then Me.MyADODB.Recordset.MovePrevious
            End If
        Case "Last"
            If Me.DataSourceType = dsDAO Then
                Me.MyDAODB.Recordset.MoveLast
            Else
                Me.MyADODB.Recordset.MoveLast
            End If
        Case "Add"
            If Me.DataSourceType = dsDAO Then
                Me.MyDAODB.Recordset.AddNew
            Else
                Me.MyADODB.Recordset.AddNew
            End If
        Case "Delete"
            If Me.DataSourceType = dsDAO Then
                If Not Me.MyDAODB.Recordset.EOF And Not Me.MyDAODB.Recordset.BOF Then Me.MyDAODB.Recordset.Delete
                Me.MyDAODB.Refresh
            Else
                If Not Me.MyADODB.Recordset.EOF And Not Me.MyADODB.Recordset.BOF Then Me.MyADODB.Recordset.Delete
                Me.MyADODB.Refresh
            End If
        Case "Filter"
            'set the text of the sql textbox to the current recordsource
            If Me.DataSourceType = dsDAO Then
                frmSQL.txtSQL = Me.MyDAODB.RecordSource
            Else
                frmSQL.txtSQL = Me.MyADODB.RecordSource
            End If
            'select the text so the user can quickly delete it
            frmSQL.txtSQL.SelStart = 0
            frmSQL.txtSQL.SelLength = Len(frmSQL.txtSQL)
            'show the filter dialog window
            frmSQL.Show vbModal, Me
            'if the user pressed cancel then unload the form and exit
            If frmSQL.pCancel = True Then
                Unload frmSQL
                Exit Sub
            End If
            'reset the recordsource
            If Me.DataSourceType = dsDAO Then
                Me.MyDAODB.RecordSource = frmSQL.txtSQL
                Me.MyDAODB.Refresh
            Else
                Me.MyADODB.RecordSource = frmSQL.txtSQL
                Me.MyADODB.Refresh
            End If
            'unload the dialog window
            Unload frmSQL
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
        Case "DeleteFilter"
            'reset the recordsource back to the original table
            If Me.DataSourceType = dsDAO Then
                Me.MyDAODB.RecordSource = Me.TableName
                Me.MyDAODB.Refresh
            Else
                Me.MyADODB.RecordSource = Me.TableName
                Me.MyADODB.Refresh
            
            End If
        Case "Refresh"
            If Me.DataSourceType = dsDAO Then
                Me.MyDAODB.Refresh
            Else
                Me.MyADODB.Refresh
            End If
    End Select
    
    Exit Sub
    
Err_Handle:
    On Error Resume Next
    MsgBox "Error: " & Error, vbExclamation + vbOKOnly
    If Me.DataSourceType = dsDAO Then
        Me.MyDAODB.RecordSource = Me.TableName
        Me.MyDAODB.Refresh
    Else
        Me.MyADODB.RecordSource = Me.TableName
        Me.MyADODB.Refresh
    End If
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'set the status bar caption to the tooltip text of the button
    
    For i = 1 To Me.tbrMain.Buttons.Count
        With Me.tbrMain.Buttons(i)
            If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height Then
                Me.sbrMain.Panels(2).Text = .ToolTipText
                Exit Sub
            End If
        End With
    Next
    Me.sbrMain.Panels(2).Text = ""
End Sub
