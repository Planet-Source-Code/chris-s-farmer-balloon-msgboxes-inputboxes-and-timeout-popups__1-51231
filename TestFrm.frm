VERSION 5.00
Begin VB.Form TestFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating Balloons"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "TestFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Generate Code for this balloon"
      Height          =   585
      Left            =   3750
      TabIndex        =   38
      Top             =   5040
      Width           =   1485
   End
   Begin VB.TextBox TC 
      ForeColor       =   &H00C00000&
      Height          =   2115
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "TestFrm.frx":000C
      Top             =   5685
      Width           =   8040
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Box Settings"
      Height          =   5625
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   8190
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   345
         Left            =   7695
         TabIndex        =   46
         ToolTipText     =   "Clear Caption and Prompt boxes ready for you to enter your own."
         Top             =   345
         Width           =   330
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   210
         LargeChange     =   15
         Left            =   195
         Max             =   75
         SmallChange     =   5
         TabIndex        =   44
         Top             =   4215
         Value           =   60
         Width           =   2430
      End
      Begin VB.Frame Frame6 
         Caption         =   "Auto Timed Msgbox"
         Height          =   1440
         Left            =   2280
         TabIndex        =   39
         Top             =   1425
         Visible         =   0   'False
         Width           =   4080
         Begin VB.CommandButton Command3 
            Caption         =   "Demonstrate positioning"
            Height          =   450
            Left            =   135
            TabIndex        =   43
            Top             =   885
            Width           =   1995
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2580
            MaxLength       =   2
            TabIndex        =   40
            Text            =   "5"
            Top             =   255
            Width           =   405
         End
         Begin VB.Label Label12 
            Caption         =   "Seconds"
            Height          =   195
            Left            =   3090
            TabIndex        =   42
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label11 
            Caption         =   "Time to display the message for:"
            Height          =   195
            Left            =   195
            TabIndex        =   41
            Top             =   285
            Width           =   2520
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Don't display a pointer (Centre the baloon on the screen.)"
         Height          =   435
         Left            =   165
         TabIndex        =   36
         Top             =   3690
         Width           =   2430
      End
      Begin VB.Frame Frame5 
         Caption         =   "Baloon Color"
         Height          =   1710
         Left            =   3780
         TabIndex        =   29
         Top             =   2880
         Width           =   1455
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   630
            MaxLength       =   3
            TabIndex        =   35
            Text            =   "204"
            Top             =   1185
            Width           =   510
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   630
            MaxLength       =   3
            TabIndex        =   33
            Text            =   "255"
            Top             =   720
            Width           =   510
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   630
            MaxLength       =   3
            TabIndex        =   31
            Text            =   "255"
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label9 
            Caption         =   "Blue:"
            Height          =   255
            Index           =   2
            Left            =   255
            TabIndex        =   34
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label Label9 
            Caption         =   "Green:"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   32
            Top             =   735
            Width           =   525
         End
         Begin VB.Label Label9 
            Caption         =   "Red:"
            Height          =   255
            Index           =   0
            Left            =   255
            TabIndex        =   30
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Click toTest "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   5310
         TabIndex        =   27
         Top             =   2955
         Width           =   2730
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   150
         TabIndex        =   26
         Text            =   "Don't display this dialog again."
         Top             =   3285
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show a checkbox with caption:"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   3075
         Width           =   2580
      End
      Begin VB.ListBox List2 
         Height          =   1230
         ItemData        =   "TestFrm.frx":0012
         Left            =   6420
         List            =   "TestFrm.frx":0033
         TabIndex        =   23
         ToolTipText     =   "Note the icon will not display in the IDE (compile to see the icon)"
         Top             =   1635
         Width           =   1620
      End
      Begin VB.Frame Frame3 
         Caption         =   "Default Button"
         Height          =   1440
         Left            =   4740
         TabIndex        =   16
         Top             =   1440
         Width           =   1620
         Begin VB.OptionButton Option1 
            Caption         =   "Third Button"
            Enabled         =   0   'False
            Height          =   225
            Index           =   2
            Left            =   105
            TabIndex        =   19
            Top             =   990
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Second Button"
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   18
            Top             =   690
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "First Button"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   17
            Top             =   390
            Value           =   -1  'True
            Width           =   1425
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Custom Buttons Text"
         Enabled         =   0   'False
         Height          =   1440
         Left            =   2295
         TabIndex        =   9
         Top             =   1440
         Width           =   2370
         Begin VB.TextBox Text3 
            Height          =   255
            Index           =   2
            Left            =   765
            TabIndex        =   12
            Top             =   1050
            Width           =   1470
         End
         Begin VB.TextBox Text3 
            Height          =   255
            Index           =   1
            Left            =   765
            TabIndex        =   11
            Top             =   705
            Width           =   1470
         End
         Begin VB.TextBox Text3 
            Height          =   255
            Index           =   0
            Left            =   765
            TabIndex        =   10
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label Label5 
            Caption         =   "button 3"
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   15
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label5 
            Caption         =   "button 2"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   14
            Top             =   735
            Width           =   705
         End
         Begin VB.Label Label5 
            Caption         =   "button 1"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   13
            Top             =   390
            Width           =   705
         End
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "TestFrm.frx":0095
         Left            =   150
         List            =   "TestFrm.frx":00AE
         TabIndex        =   8
         Top             =   1665
         Width           =   2040
      End
      Begin VB.TextBox Text2 
         Height          =   720
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "TestFrm.frx":0114
         Top             =   675
         Width           =   7875
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3915
         TabIndex        =   4
         Text            =   "This is the title to use."
         Top             =   90
         Width           =   4125
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "TestFrm.frx":018E
         Left            =   1065
         List            =   "TestFrm.frx":019B
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   90
         Width           =   1650
      End
      Begin VB.Frame Frame4 
         Caption         =   "Input box Settings:"
         Height          =   1440
         Left            =   2280
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   4065
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   135
            TabIndex        =   22
            Top             =   585
            Width           =   3750
         End
         Begin VB.Label Label6 
            Caption         =   "Default Text:"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   345
            Width           =   1350
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Dark      -Dropshadow intensity-    light"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   90
         TabIndex        =   45
         Top             =   4425
         Width           =   2685
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Return Value:"
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   5310
         TabIndex        =   28
         Top             =   5070
         Width           =   2715
      End
      Begin VB.Label Label7 
         Caption         =   "Icon"
         Height          =   180
         Left            =   6435
         TabIndex        =   24
         Top             =   1425
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Buttons"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   1470
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Message"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   465
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Caption title: "
         Height          =   255
         Left            =   2925
         TabIndex        =   3
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Type of box"
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   120
         Width           =   915
      End
   End
End
Attribute VB_Name = "TestFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Dim tx$


Select Case Combo1.ListIndex
    Case 0 'msgbox
        Text1.Text = "Msgbox Title"
        tx = "This is the message to be displayed. The balloon will automatically resize itself (Verically)according"
        tx = tx & " to the size of this text."
        tx = tx & " The pointer of the balloon will appear at the cursor position. The cursor position is checked to"
        tx = tx & " see in which quarter of the screen it exists and then the balloon is displayed in it's best position."
        tx = tx & vbCrLf & vbCrLf & "When you click and drag the balloon you will notice the pointer and drop shadow dissapear"
        tx = tx & " whilst dragging. The drop shadow reapears when the mouse is released but since the balloon has been moved there"
        tx = tx & " no point to showing the pointer again."
        Text2.Text = tx
        Frame4.Visible = False
        Frame2.Visible = True
        Frame3.Visible = True
        List1.Enabled = True
        Frame6.Visible = False
        List1.ListIndex = 0
        Check1.Enabled = True

    Case 1 'inputbo
        Text1.Text = "Inputbox Title"
        tx = tx & "This is the message to display in the Input balloon. The return value of the inputbox is the string"
        tx = tx & " entered by the user into the textbox displayed. You can display an icon and checkbox in this balloon also."
        
        Text2.Text = tx
        Frame4.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
        List1.ListIndex = 2
        List1.Enabled = False
        Frame6.Visible = False
        Check1.Enabled = True
    
    Case 2 'timed msgbox
        'no buttons and no return value
        Text1.Text = "Auto Timed Msgbox"
        tx = "This is the message to display. The msgbox balloon will automatically close itself after the set period "
        tx = tx & "of time. In order to display more than one auto timed msgbox at the same time you will need to modify "
        tx = tx & "the code, (see notes in the TimerProc Sub)."
        
        Text2.Text = tx
        Frame4.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        List1.ListIndex = -1
        List1.Enabled = False
        Frame6.Visible = True
        Check1.Value = 0
        Check1.Enabled = False
End Select
End Sub


Private Sub Command1_Click()

Dim X%, nRet$, n%, msg$, Title$, nFlags As MsgBox_Flags, nCol As OLE_COLOR, CheckTxt$, nPointer As Boolean
Title = Text1.Text
msg = Text2.Text

Select Case Combo1.ListIndex
    Case 0 'msgbox
         
         If List1.ListIndex = -1 Then Exit Sub
         nFlags = List1.ItemData(List1.ListIndex)
         
         'validate choices
         If List1.ListIndex = List1.ListCount - 1 Then 'just in case it wasn't entered
            If Text3(0).Text = "" Then
                Text3(0).Text = "OK"
                Option1(0).Value = True
            End If
            If (Text3(1).Text = "" And Option1(0).Value = False) Then Option1(0).Value = True
            If (Text3(2).Text = "" And Option1(2).Value = True) Then Option1(0).Value = True
         End If
         
         If Option1(0).Value = True Then
            nFlags = nFlags Or vbDefaultButton1
         ElseIf Option1(1).Value = True Then
            nFlags = nFlags Or vbDefaultButton2
         Else
            nFlags = nFlags Or vbDefaultButton3
         End If
         If List2.ListIndex <> -1 Then nFlags = nFlags Or List2.ItemData(List2.ListIndex)
         
         nCol = RGB(CInt(Text6(0).Text), CInt(Text6(1).Text), CInt(Text6(2).Text))
         If Check1.Value = 1 Then CheckTxt = Text5.Text Else CheckTxt = ""
         
         
         'call the msgbox which is modal
        X = msgFrm.msg_Box(msg, nFlags, Title, CBool(Check2.Value), CheckTxt, Text3(0).Text, Text3(1).Text, Text3(2).Text, nCol, , , HScroll1.Value)
         
        Label8.Caption = ""
        If X > 100 Then
            Label8.Caption = "Checkbox was checked" & Chr(10)
            X = X - 100
        End If
        If List1.ListIndex <> List1.ListCount - 1 Then
            Select Case X
                Case vbYes:     Label8.Caption = Label8.Caption & "Yes was pressed"
                Case vbNo:      Label8.Caption = Label8.Caption & "No was pressed"
                Case vbAbort:   Label8.Caption = Label8.Caption & "Abort was pressed"
                Case vbCancel:  Label8.Caption = Label8.Caption & "Cancel was pressed"
                Case vbOK:      Label8.Caption = Label8.Caption & "OK was pressed"
                Case vbRetry:   Label8.Caption = Label8.Caption & "Retry was pressed"
                Case vbIgnore:  Label8.Caption = Label8.Caption & "Ignore was pressed"
            End Select
        Else
            Select Case X
                Case 1: Label8.Caption = Label8.Caption & "Button 1 (" & Text3(0).Text & ") was pressed."
                Case 2: Label8.Caption = Label8.Caption & "Button 2 (" & Text3(1).Text & ") was pressed."
                Case 3: Label8.Caption = Label8.Caption & "Button 3 (" & Text3(2).Text & ") was pressed."
            End Select
        End If
        
        
    Case 1 'inputbox
        
        If List2.ListIndex <> -1 Then nFlags = List2.ItemData(List2.ListIndex)
        nCol = RGB(CInt(Text6(0).Text), CInt(Text6(1).Text), CInt(Text6(2).Text))
        If Check1.Value = 1 Then CheckTxt = Text5.Text Else CheckTxt = ""
        
        nRet = msgFrm.input_box(msg, nFlags, Title, Text4.Text, CBool(Check2.Value), CheckTxt, nCol, , , HScroll1.Value)
        Label8.Caption = ""
        If Right(nRet, 1) = "¬" Then
            Label8.Caption = "Checkbox is checked" & Chr(10)
            nRet = Left(nRet, Len(nRet) - 1)
        End If
        If nRet <> "" Then
            Label8.Caption = Label8.Caption & "Returned string is : " & nRet
        Else
            Label8.Caption = "The inputbox was canceled."
        End If
        
     Case 2 'TIMED MSGBOX
        If List2.ListIndex <> -1 Then nFlags = List2.ItemData(List2.ListIndex)
        nCol = RGB(CInt(Text6(0).Text), CInt(Text6(1).Text), CInt(Text6(2).Text))
        
        Call msgFrm.MsgAuto_box(msg, nFlags, Title, CInt(Text7.Text), CBool(Check2.Value), nCol, , , HScroll1.Value)
        
     
End Select


End Sub



Private Sub Command2_Click()
Dim txt$, txt1$
With TC
    If Combo1.ListIndex = 0 Then
        .Text = "Dim nPrompt$,nTitle$,nRet%,CheckTxt$,nFlags As MsgBox_Flags" & vbCrLf
        .Text = .Text & "nPrompt = " & Chr(34) & Text2.Text & Chr(34) & vbCrLf
        .Text = .Text & "nTitle = " & Chr(34) & Text1.Text & Chr(34) & vbCrLf
        txt = "nFlags = " & List1.List(List1.ListIndex)
        If Option1(0).Value = True Then
            txt = txt & " OR vbDefaultButton1"
        ElseIf Option1(1).Value = True Then
            txt = txt & " OR vbDefaultButton2"
        Else
            txt = txt & " OR vbDefaultButton3"
        End If
        If List2.ListIndex <> -1 Then txt = txt & " OR " & List2.List(List2.ListIndex)
        .Text = .Text & txt & vbCrLf
        If Check1.Value = 1 Then
            .Text = .Text & "CheckTxt = " & Chr(34) & Text5.Text & Chr(34)
        Else
            .Text = .Text & "CheckTxt = vbNullString"
        End If
            
        .Text = .Text & vbCrLf & "nRet = msgFrm.msg_box(nPrompt, nFlags, nTitle," & CBool(Check2.Value) & "," & "CheckTxt" & "," & Chr(34) & Text3(0).Text & Chr(34) & "," & Chr(34) & Text3(1).Text & Chr(34) & "," & Chr(34) & Text3(2).Text & Chr(34) & ",RGB(" & CInt(Text6(0).Text) & "," & CInt(Text6(1).Text) & "," & CInt(Text6(2).Text) & "),,, " & HScroll1.Value & ")"
        .Text = .Text & vbCrLf
        If Check1.Value = 1 Then
        .Text = .Text & "if nRet > 100 Then" & vbCrLf
        .Text = .Text & "    'Checkbox was checked DO WHAT YOU WILL HERE!" & vbCrLf
        .Text = .Text & "    nRet = nRet - 100" & vbCrLf
        .Text = .Text & "End If"
        End If
        .Text = .Text & vbCrLf & "Select case nRet" & vbCrLf
        
        Select Case List1.ListIndex
            Case 1 'vbYesNo
                .Text = .Text & "    Case vbYes:'user selected YES, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbNo:'user selected NO, add code here! " & vbCrLf & vbCrLf
            Case 2 'vbOKCancel
                .Text = .Text & "    Case vbOK:'user selected OK, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbCancel:'user selected CANCEL, add code here! " & vbCrLf & vbCrLf
            Case 3 'vbYesNoCancel
                .Text = .Text & "    Case vbYes:'user selected YES, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbNo:'user selected NO, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbCancel:'user selected CANCEL, add code here! " & vbCrLf & vbCrLf
            Case 4 'RetryCancel
                .Text = .Text & "    Case vbRetry:'user selected RETRY, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbCancel:'user selected CANCEL, add code here! " & vbCrLf & vbCrLf
            Case 5 'AbortRetryIgnore
                .Text = .Text & "    Case vbAbort:'user selected ABORT, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbRetry:'user selected RETRY, add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case vbIgnore:'user selected IGNORE, add code here! " & vbCrLf & vbCrLf
            Case 6 'custombuttons
                .Text = .Text & "    Case 1:'user selected " & Text3(0).Text & " add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case 2:'user selected " & Text3(1).Text & " add code here! " & vbCrLf & vbCrLf
                .Text = .Text & "    Case 3:'user selected " & Text3(2).Text & " add code here! " & vbCrLf & vbCrLf
        End Select
        .Text = .Text & "End Select" & vbCrLf
        
    ElseIf Combo1.ListIndex = 1 Then 'input box
        .Text = "Dim nPrompt$, nTitle$, nRet$" & vbCrLf
        .Text = .Text & "nPrompt = " & Chr(34) & Text2.Text & Chr(34) & vbCrLf
        .Text = .Text & "nTitle = " & Chr(34) & Text1.Text & Chr(34) & vbCrLf
    
        If List2.ListIndex <> -1 Then txt = List2.List(List2.ListIndex) Else txt = "vbOKCancel"
        
        nCol = RGB(CInt(Text6(0).Text), CInt(Text6(1).Text), CInt(Text6(2).Text))
        If Check1.Value = 1 Then CheckTxt = Chr(34) & Text5.Text & Chr(34) Else CheckTxt = ""
        
        If Text4.Text = "" Then txt1 = "" Else txt1 = Chr(34) & Text4.Text & Chr(34)
        'nRet = msgFrm.input_box(msg, txt, Title, Text4.text, CBool(Check2.Value), CheckTxt, nCol)
        .Text = .Text & vbCrLf & "nRet = msgFrm.input_box(nPrompt, " & txt & ", nTitle," & txt1 & "," & CBool(Check2.Value) & "," & CheckTxt & ",RGB(" & CInt(Text6(0).Text) & "," & CInt(Text6(1).Text) & "," & CInt(Text6(2).Text) & "),,, " & HScroll1.Value & ")"
        .Text = .Text & vbCrLf
        
        If CheckTxt <> "" Then
            .Text = .Text & "If Right(nRet, 1) = " & Chr(34) & "¬" & Chr(34) & " Then" & vbCrLf
            .Text = .Text & "      'Checkbox is checked: do code here!" & vbCrLf
            .Text = .Text & "      nRet = Left(nRet, Len(nRet) - 1)" & vbCrLf
            .Text = .Text & "End If" & vbCrLf
        End If
        .Text = .Text & "If nRet <> vbNullString Then 'nRet is the string the user entered" & vbCrLf
        .Text = .Text & "           'Enter code to manipulate returned string here !" & vbCrLf
        .Text = .Text & "End If"
        
    ElseIf Combo1.ListIndex = 2 Then 'auto timed msgbox
    
        .Text = "Dim nPrompt$, nTitle$" & vbCrLf & vbCrLf
        .Text = .Text & "nPrompt = " & Chr(34) & Text2.Text & Chr(34) & vbCrLf
        .Text = .Text & "nTitle = " & Chr(34) & Text1.Text & Chr(34) & vbCrLf
    
         If List2.ListIndex <> -1 Then txt = List2.List(List2.ListIndex) Else txt = "vbOKOnly "



'            If List2.ListIndex <> -1 Then nFlags = List2.ItemData(List2.ListIndex)
 '       nCol = RGB(CInt(Text6(0).Text), CInt(Text6(1).Text), CInt(Text6(2).Text))
        
        .Text = .Text & vbCrLf
        .Text = .Text & "'the X,Y position of the pointer may be set by entering the screen coordinates (pixels) in the following call." & vbCrLf
        .Text = .Text & "Call msgFrm.MsgAuto_box(nPrompt," & txt & ", nTitle," & CInt(Text7.Text) & ", " & CBool(Check2.Value) & ", RGB(" & CInt(Text6(0).Text) & "," & CInt(Text6(1).Text) & "," & CInt(Text6(2).Text) & "),,, " & HScroll1.Value & ")"
    
    
    End If
End With

End Sub


Private Sub Command3_Click()
Dim nPrompt$, nTitle$, X As Single, Y As Single

nPrompt = "The code behind this button shows how you can position the auto message box to appear anywhere on the screen. This is positioned on the close button of this form."
nTitle = "10 second Timed Msgbox"

X = (TestFrm.Left + TestFrm.Width - 300) / Screen.TwipsPerPixelX
Y = (TestFrm.Top + 300) / Screen.TwipsPerPixelY

'the X,Y position of the pointer may be set by entering the screen coordinates (pixels) in the following call.
Call msgFrm.MsgAuto_box(nPrompt, vbInformation, nTitle, 10, False, RGB(205, 255, 204), X, Y)

End Sub


Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Form_Load()
Dim tx$

 
 
 Top = (Screen.Height - Height) / 2
 Left = (Screen.Width - Width) / 2

Combo1.ListIndex = 0
List1.ListIndex = 0

tx = "This window allows you to generate the code to display the current balloon." & vbCrLf & vbCrLf
tx = tx & "Add the BubbleFrm.frm and MsgBox_Flags.bas files to your project. "
tx = tx & vbCrLf & vbCrLf
tx = tx & "You can now use the 'Generate Code for this Balloon' button to generate the code "
tx = tx & "that you can cut and paste into your application."
TC.Text = tx

End Sub


Private Function GetHexChar(sChar As String) As Integer
   'Check for Hex Char
   Select Case sChar
   Case "0" To "9"
      GetHexChar = Val(sChar)
   Case "A" To "F"
      GetHexChar = Asc(sChar) - 55
   Case Else
      Exit Function
   End Select
End Function

Private Sub List1_Click()
Dim n%
If List1.Selected(List1.ListCount - 1) Then
    Frame2.Enabled = True
    For n = 0 To 2
        Text3(n).Enabled = True
        Label5(n).Enabled = True
    Next n
Else
    Frame2.Enabled = False
    For n = 0 To 2
        Text3(n).Text = ""
        Text3(n).Enabled = False
        Label5(n).Enabled = False
    Next n
End If

Select Case List1.ListIndex
    Case 0
        Option1(1).Enabled = False
        Option1(2).Enabled = False
    Case 1, 2, 4
        Option1(1).Enabled = True
        Option1(2).Enabled = False
    Case Else
        Option1(1).Enabled = True
        Option1(2).Enabled = True
End Select

End Sub


