VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connector Line"
   ClientHeight    =   8010
   ClientLeft      =   2790
   ClientTop       =   2325
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   7800
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1111
      ButtonWidth     =   2672
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Create Object"
            Key             =   "Object"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Create Relationship"
            Key             =   "Relationship"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update File"
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlArrow 
      Left            =   8400
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":0000
            Key             =   "TOP"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":03CE
            Key             =   "LEFT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":0798
            Key             =   "RIGHT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConnector.frx":0B64
            Key             =   "BOTTOM"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblObject 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Here And Move The Object"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6315
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Line lneLeft 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   6435
      X2              =   7710
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line lneMiddle 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   7725
      X2              =   7725
      Y1              =   1710
      Y2              =   2880
   End
   Begin VB.Line lneRight 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   7695
      X2              =   9060
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Image imgArrow 
      Height          =   240
      Index           =   0
      Left            =   7020
      Picture         =   "frmConnector.frx":0F33
      Top             =   1575
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bStartMove As Boolean   ''' move a window wherever clicked
Dim sXClickPos As Single    ''' Current x pos
Dim sYClickPos As Single    ''' current Y Pos
Dim objLine As clsLine
Public ConfigINIFileName As String
Dim ObjectCount As Integer, RelationCount As Integer
Dim LastControl As Integer
Dim FSO As New FileSystemObject
Dim intRS(1000) As Integer, intRD(1000) As Integer
Dim Rel1 As Integer, Rel2 As Integer

Private Sub Form_Load()


    Set objLine = New clsLine
    ConfigINIFileName = App.Path & "\Config.ini"
    If Not FSO.FileExists(ConfigINIFileName) Then
        MsgBox "Ini File Missing"
        End
    End If
    
    Initialize

End Sub
Private Sub Initialize()

    ObjectCount = 0
    
    RelationCount = 0
    LastControl = 0
    LoadOldObjects
    LoadOldLines
    ConnectOldLines
    Rel1 = 0
    Rel2 = 0
End Sub
Private Sub ConnectOldLines()
    Dim intTotal As Integer, Cnt As Integer, temp As String * 255, n As Integer, intWidth As Integer, intHeight As Integer
    Dim intSource As Integer, intDest As Integer, intLine As Integer
    Dim strText As String, strArray, innerLoop As Integer
    
    RelationCount = GetPrivateProfileInt("Global", "RelationCount", "0", ConfigINIFileName)
    
    For Cnt = 1 To RelationCount
        n = GetPrivateProfileString("Relations", Cnt, "", temp, 255, ConfigINIFileName)
        strText = Trim(Mid(temp, 1, n))
        
        strArray = Split(strText, ",")
            
            intRS(Cnt) = Val(strArray(0))
            intRD(Cnt) = Val(strArray(1))
            
            objLine.drawConnectLine lblObject(intRS(Cnt)), lblObject(intRD(Cnt)).Left, lblObject(intRD(Cnt)).Top, Cnt, Me
            
    Next
    
End Sub
Private Function AlreadyExists(ByVal Obj1 As Integer, ByVal Obj2 As Integer) As Boolean
    Dim Cnt As Integer
    AlreadyExists = False
        For Cnt = 1 To RelationCount
            If intRS(Cnt) = Obj1 And intRD(Cnt) = Obj2 Then
                AlreadyExists = True
                Exit Function
            End If
        Next
End Function
Private Sub CreateRelationship()
On Error GoTo Err
    If Rel1 > 0 And Rel2 > 0 Then
        If Rel1 <> Rel2 Then
            If AlreadyExists(Rel1, Rel2) = False Then
                ConnectNewLine Rel1, Rel2
            Else
                GoTo Err
            End If
        Else
            GoTo Err
        End If

    Else
        GoTo Err
    End If

Rel1 = 0
Rel2 = 0

Exit Sub
Err:
MsgBox ("Unable to Create Relationship")
Rel1 = 0
Rel2 = 0
End Sub
Private Sub ConnectNewLine(ByVal Obj1 As Integer, ByVal Obj2 As Integer)
    LoadNewLine (0)
    'RelationCount = RelationCount + 1
    
    intRS(RelationCount) = Obj1
    intRD(RelationCount) = Obj2
    
            
    objLine.drawConnectLine lblObject(Obj1), lblObject(Obj2).Left, lblObject(Obj2).Top, RelationCount, Me
    
End Sub

Private Sub UpdateGlobaltoFile()
    
    WritePrivateProfileString "Global", "ObjectCount", CStr(ObjectCount), ConfigINIFileName
    WritePrivateProfileString "Global", "RelationCount", CStr(RelationCount), ConfigINIFileName
    
End Sub
Private Sub UpdateRelationtoFile()
    Dim Cnt As Integer, strText As String
    
    
    For Cnt = 1 To RelationCount
        strText = intRS(Cnt) & "," & intRD(Cnt)
        WritePrivateProfileString "Relations", CStr(Cnt), strText, ConfigINIFileName
    Next
    
    
End Sub

Private Sub UpdateObjecttoFile()
    Dim Cnt As Integer, strText As String
    
    
    For Cnt = 1 To ObjectCount
        WritePrivateProfileString "Settings", Cnt & "_Left", CStr(lblObject(Cnt).Left), ConfigINIFileName
        WritePrivateProfileString "Settings", Cnt & "_Top", CStr(lblObject(Cnt).Top), ConfigINIFileName
        WritePrivateProfileString "Settings", Cnt & "_Text", CStr(lblObject(Cnt).Caption), ConfigINIFileName
    Next
    
    
End Sub

Private Sub UpdatetoFile()
    UpdateGlobaltoFile
    UpdateObjecttoFile
    UpdateRelationtoFile

End Sub


Private Sub LoadNewObject(ByVal intVal As Integer)
        If intVal = 0 Then
            ObjectCount = ObjectCount + 1
        Else
            ObjectCount = intVal
        End If
        
        Load lblObject(ObjectCount)
        lblObject(ObjectCount).Move 4000, 1000
        
        lblObject(ObjectCount).Visible = True
        lblObject(ObjectCount).Refresh
        LastControl = ObjectCount
    
End Sub
Private Sub LoadOldObjects()

    Dim intTotal As Integer, Cnt As Integer, temp As String * 255, n As Integer, intWidth As Integer, intHeight As Integer
    Dim intLeft As Integer, intTop As Integer, strText As String, strBG As String, intBG As Double, intFG As Double
    Dim intFS As Integer, strFN As String, intPage As Integer, LC As Integer
    
    ObjectCount = GetPrivateProfileInt("Global", "ObjectCount", "0", ConfigINIFileName)
    For Cnt = 1 To ObjectCount
          
        intLeft = GetPrivateProfileInt("Settings", Cnt & "_Left", "0", ConfigINIFileName)
        intTop = GetPrivateProfileInt("Settings", Cnt & "_Top", "0", ConfigINIFileName)
        n = GetPrivateProfileString("Settings", Cnt & "_Text", "", temp, 255, ConfigINIFileName)
        strText = Trim(Mid(temp, 1, n))
            'If intLeft > 0 Or intTop > 0 Then
                LoadNewObject (Cnt)
                lblObject(Cnt).Caption = strText
                lblObject(Cnt).Move intLeft, intTop
                
            'End If
                
    Next
End Sub

Private Sub LoadNewLine(ByVal intVal As Integer)
        If intVal = 0 Then
            RelationCount = RelationCount + 1
        Else
            RelationCount = intVal
        End If
        
        Load lneLeft(RelationCount)
        Load lneMiddle(RelationCount)
        Load lneRight(RelationCount)
        Load imgArrow(RelationCount)
               
        lneLeft(RelationCount).Visible = True
        lneMiddle(RelationCount).Visible = True
        lneRight(RelationCount).Visible = True
        imgArrow(RelationCount).Visible = True
        
    
End Sub
Private Sub LoadOldLines()

    Dim intTotal As Integer, Cnt As Integer, temp As String * 255, n As Integer, intWidth As Integer, intHeight As Integer
    Dim intLeft As Integer, intTop As Integer, strText As String, strBG As String, intBG As Double, intFG As Double
    Dim intFS As Integer, strFN As String, intPage As Integer, LC As Integer
    
    RelationCount = GetPrivateProfileInt("Global", "RelationCount", "0", ConfigINIFileName)
    
    For Cnt = 1 To RelationCount
        LoadNewLine (Cnt)
                
    Next
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Do you want to save changes", vbYesNo) = vbYes Then
        UpdatetoFile
    End If
End Sub

Private Sub lblObject_Click(Index As Integer)
    If Rel1 = 0 Then
        Rel1 = Index
    ElseIf Rel1 > 0 Then
        Rel2 = Index
    End If
End Sub

Private Sub lblObject_DblClick(Index As Integer)
Dim strText As String
    strText = InputBox("Enter Text for the Object", , lblObject(Index).Caption)
    If Trim(strText) <> "" Then
        lblObject(Index).Caption = strText
    End If
End Sub

Private Sub lblObject_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    bStartMove = True
    ''' get the mouse click X,Y position
    sXClickPos = X
    sYClickPos = Y
Else
    bStartMove = False
End If
End Sub
Private Sub MoveCL(SourceCtrl As Integer, X As Single, Y As Single)

    Dim Cnt As Integer
    
    If bStartMove Then
        '''''' Move the object with respect to the clicked position
        If sXClickPos > X Then
            lblObject(SourceCtrl).Left = lblObject(SourceCtrl).Left - (sXClickPos - X)
       ElseIf sXClickPos < X Then
            lblObject(SourceCtrl).Left = lblObject(SourceCtrl).Left + (X - sXClickPos)
       End If
       If sYClickPos < Y Then
            lblObject(SourceCtrl).Top = lblObject(SourceCtrl).Top + (Y - sYClickPos)
       ElseIf sYClickPos > Y Then
            lblObject(SourceCtrl).Top = lblObject(SourceCtrl).Top - (sYClickPos - Y)
       End If
       '''''''''
       For Cnt = 1 To RelationCount
            If intRS(Cnt) = SourceCtrl Then
                objLine.moveConnectLine lblObject(SourceCtrl), lblObject(intRD(Cnt)), Cnt, Me, True
            ElseIf intRD(Cnt) = SourceCtrl Then
                objLine.moveConnectLine lblObject(SourceCtrl), lblObject(intRS(Cnt)), Cnt, Me, False
            
            End If
       Next
       
    End If

End Sub

Private Sub lblObject_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    MoveCL Index, X, Y
End If
End Sub

Private Sub lblObject_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bStartMove = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Object" Then
        LoadNewObject (0)
    
    ElseIf Button.Key = "Relationship" Then
        CreateRelationship
    ElseIf Button.Key = "File" Then
        UpdatetoFile
        
    End If
End Sub
