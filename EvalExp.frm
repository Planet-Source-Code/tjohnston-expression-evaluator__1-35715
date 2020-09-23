VERSION 5.00
Begin VB.Form EvalExp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expression Evaluator"
   ClientHeight    =   4230
   ClientLeft      =   6555
   ClientTop       =   4935
   ClientWidth     =   5565
   Icon            =   "EvalExp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5565
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   315
      Left            =   4380
      TabIndex        =   10
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   315
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   1260
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   315
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   540
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Evaluate"
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   3795
   End
   Begin VB.Label lblresult 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   3810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   8
      Top             =   2940
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expression"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   7
      Top             =   2100
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variable List"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   6
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "EvalExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub EvalExpression()
 
   On Error GoTo matherr
   
   ' import all variables and their values
   vlist = "|"
   lc = List1.ListCount - 1
   For t = 0 To lc
      vlist = vlist & UCase$(List1.List(t)) & "|"
   Next
   
   e = Text1

   ' remove all spaces to avoid problems
   While InStr(e, " ") > 0
      pp = InStr(e, " ")
      e = Left$(e, pp - 1) & Right$(e, Len(e) - pp)
   Wend

   Do
      ' locate next ( ) section
      pp = InStr(e, ")")
      ppc = InStr(e, "(")
      If (ppc > 0 And pp = 0) Or (pp > 0 And ppc = 0) Then
         MsgBox "Parenthesis missmatch - check your opens and closes!", 16, "Stein Seal Advisor Message"
         Exit Sub ' error parsing - open/close parenthesis missmatch
      End If
      ppop = InStr(e, "/") + InStr(e, "*") + InStr(e, "^") + InStr(e, "+") + InStr(e, "-")
      If pp = 0 And ppop = 0 Then Exit Do
      pp2 = 0
      If pp > 0 Then pp2 = InStrRev(e, "(", pp)
      If pp > 0 And pp2 > 0 Then ee = Mid$(e, pp2 + 1, (pp - pp2 - 1)) Else ee = e
      ' evaluate expression
      Do
         If InStr(ee, "/") + InStr(ee, "*") + InStr(ee, "^") + InStr(ee, "+") + InStr(ee, "-") > 0 Then
            ' follow 'my dear aunt sally' method if multiple operations found in same ( ) group
            ppe = 0
            If ppe = 0 Then If InStr(ee, "^") > 0 Then ppe = InStr(ee, "^"): tp = "^"
            If ppe = 0 Then If InStr(ee, "*") > 0 Then ppe = InStr(ee, "*"): tp = "*"
            If ppe = 0 Then If InStr(ee, "/") > 0 Then ppe = InStr(ee, "/"): tp = "/"
            If ppe = 0 Then If InStr(ee, "+") > 0 Then ppe = InStr(ee, "+"): tp = "+"
            If ppe = 0 Then If InStr(ee, "-") > 0 Then ppe = InStr(ee, "-"): tp = "-"
            If ppe > 0 Then
               s = GetValues(ee, ppe, vlist)
               l1 = Val(Delim(s, "|"))
               l2 = Val(Delim(s, "|"))
               v1 = Val(Delim(s, "|"))
               v2 = Val(s)
               If tp = "^" Then r = v1 ^ v2
               If tp = "*" Then r = v1 * v2
               If tp = "/" Then r = v1 / v2
               If tp = "+" Then r = v1 + v2
               If tp = "-" Then r = v1 - v2
               r = Trim$(Str$(r))
               ' replace original expression with final value
               ee = Left$(ee, l1 - 1) & r & Right$(ee, Len(ee) - l2)
            End If
         Else
            ' replace entire expression with final value - check if a math function exists
            ' immediately outside () expression or () is just used to force the calculation order
            r = Val(ee)
            pp3 = pp2 - 3
            If pp3 >= 1 Then
               ppf = UCase$(Mid$(e, pp3, 3))
               If InStr("/ABS/ATN/COS/EXP/FIX/INT/LOG/SGN/SIN/SQR/TAN/", "/" & ppf & "/") > 0 Then
                  pp2 = pp2 - 3
                  If ppf = "ABS" Then r = Abs(r)
                  If ppf = "ATN" Then r = Atn(r)
                  If ppf = "COS" Then r = Cos(r)
                  If ppf = "EXP" Then r = Exp(r)
                  If ppf = "FIX" Then r = Fix(r)
                  If ppf = "INT" Then r = Int(r)
                  If ppf = "LOG" Then r = Log(r)
                  If ppf = "SGN" Then r = Sgn(r)
                  If ppf = "SIN" Then r = Sin(r)
                  If ppf = "SQR" Then r = Sqr(r)
                  If ppf = "TAN" Then r = Tan(r)
               End If
            End If
            r = Trim$(Str$(r))
            If pp2 > 0 And pp > 0 Then
               e = Left$(e, pp2 - 1) & r & Right$(e, Len(e) - pp)
            Else
               e = r
            End If
            Exit Do
         End If
      Loop
   Loop
   
   lblresult = e
   
endeval:
   On Error GoTo 0
   
   Exit Sub
   
matherr:
   MsgBox "An error occurred '" & Error$ & "' while trying to evaluate this formula.", 16, "Stein Seal Advisor Message"
   Resume endeval
   
End Sub
Function Delim(s, ByVal d)

   ' return left portion of string 's' prior to first
   ' occurance of delimiting character 'd'
   '
   ' strip string of leftmost portion, including
   ' delimiting character to prepare for next function call

   p = InStr(s, d)
   If p > 0 Then
      l = Left$(s, InStr(s, d) - 1)
      s = Right$(s, Len(s) - InStr(s, d))
   Else
      l = "" ' error - delimiter char not found, return empty string
   End If

   Delim = l

End Function
Function GetValues(ByVal ee, ByVal ppe, ByVal vlist) As String

   ' get variable or value to left of operand
   pp1 = ppe - 1
   vflag1 = False
   Do While pp1 > 0
      a = Asc(UCase$(Mid$(ee, pp1)))
      If Not ((a >= 65 And a <= 97) Or (a >= 48 And a <= 57) Or a = 46) Then pp1 = pp1 + 1: Exit Do
      If a >= 65 And a <= 97 Then vflag1 = True
      pp1 = pp1 - 1
   Loop
   If pp1 = 0 Then pp1 = 1
   vleft = Mid$(ee, pp1, (ppe - pp1))
   If vflag1 Then
      ' alpha variable found - locate corrosponding value
      ppp = InStr(vlist, "|" & UCase$(vleft) & "=")
      If ppp > 0 Then
         eee = Right$(vlist, Len(vlist) - ppp)
         xxx = Delim(eee, "=")
         vleft = Delim(eee, "|")
      End If
   End If
   
   ' get variable or value to right of operand
   pp2 = ppe + 1
   vflag2 = False
   Do While pp2 <= Len(ee)
      a = Asc(UCase$(Mid$(ee, pp2)))
      If Not ((a >= 65 And a <= 97) Or (a >= 48 And a <= 57) Or a = 46) Then pp2 = pp2 - 1: Exit Do
      If a >= 65 And a <= 97 Then vflag2 = True
      pp2 = pp2 + 1
   Loop
   If pp2 > Len(ee) Then pp2 = Len(ee)
   vright = Mid$(ee, ppe + 1, pp2 - ppe)
   If vflag2 Then
      ' alpha variable found - locate corrosponding value
      ppp = InStr(vlist, "|" & UCase$(vright) & "=")
      If ppp > 0 Then
         eee = Right$(vlist, Len(vlist) - ppp)
         xxx = Delim(eee, "=")
         vright = Delim(eee, "|")
      End If
   End If

   GetValues = pp1 & "|" & pp2 & "|" & vleft & "|" & vright
   
End Function
Sub Variable_Add()

   e = InputBox("Enter Variable and It's Value (ex. X=10)", "Enter Variable")
   If e > "" Then List1.AddItem e
   
End Sub
Sub Variable_Delete()

   If List1.ListIndex = -1 Then Exit Sub

   List1.RemoveItem List1.ListIndex
   
End Sub
Sub Variable_Update()

   If List1.ListIndex = -1 Then Exit Sub
   
   e = InputBox("Enter Variable and It's Value (ex. X=10)", "Enter Variable", List1.Text)
   If e > "" Then List1.List(List1.ListIndex) = e
   
End Sub
Private Sub Command1_Click()

   EvalExpression
   
End Sub
Private Sub Command2_Click(Index As Integer)

   If Index = 0 Then Variable_Add
   If Index = 1 Then Variable_Update
   If Index = 2 Then Variable_Delete
   
End Sub
Private Sub Command3_Click()

   End
   
End Sub
Private Sub Form_Load()

   ' center form on screen
   Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
   
End Sub
Private Sub List1_DblClick()

   Variable_Update
   
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      KeyAscii = 0
      Call Command1_Click
   End If
   
End Sub
