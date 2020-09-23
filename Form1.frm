VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CnTime Test Form"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDateTime 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblData 
      Caption         =   "Data Goes Here..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblDateTime 
      AutoSize        =   -1  'True
      Caption         =   "Enter A Date/Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   200
      Width           =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    txtDateTime.Text = Now()
    Call cmdConvert_Click
End Sub



Private Sub cmdConvert_Click()
'    lblData.Caption = "GetTimeZoneOffset: " & GetTimeZoneOffset()
'    lblData.Caption = lblData.Caption & vbCrLf & "GetCurrentTimeZoneOffset: " & GetCurrentTimeZoneOffset()
'    Exit Sub
    
    
    Dim dDate As Date
    Dim dblTimestamp As Double

    On Error GoTo cmdConvert_Error
    dDate = CDate(txtDateTime.Text)
    dblTimestamp = DateToTimestamp(dDate)

    lblData.Caption = "Timestamp of Now():    " & Timestamp() & vbCrLf
'    lblData.Caption = "CDbl'd Now():    " & CDbl(Now()) & vbCrLf
    lblData.Caption = lblData.Caption & vbCrLf

    lblData.Caption = lblData.Caption & "Date Entered:          " & dDate & vbCrLf
    lblData.Caption = lblData.Caption & "Date's Timestamp:      " & dblTimestamp & vbCrLf
    lblData.Caption = lblData.Caption & "Timestamp to Date:     " & TimestampToDate(dblTimestamp) & vbCrLf
    lblData.Caption = lblData.Caption & vbCrLf

    lblData.Caption = lblData.Caption & "TimeZoneOffset:                 " & GetTimeZoneOffset() & vbCrLf
    lblData.Caption = lblData.Caption & "CurrentTimeZoneOffset:          " & GetCurrentTimeZoneOffset() & vbCrLf
    lblData.Caption = lblData.Caption & vbCrLf

    lblData.Caption = lblData.Caption & "Is Entered Year a Leap Year?:   " & isLeapYear(Year(dDate)) & vbCrLf
    lblData.Caption = lblData.Caption & "Days In Entered Month:          " & DaysInMonth(Month(dDate), Year(dDate)) & vbCrLf
    lblData.Caption = lblData.Caption & vbCrLf
    Exit Sub

cmdConvert_Error:
    Call MsgBox("Please type in a valid date into the text box!", vbExclamation + vbOKOnly, "Invalid Date Provided")
End Sub


'April 4, 1966, 12:15:25
'
'Time:       25 + 900 + 43200
'Month:      7776000
'Day:        345600
'Year:       94608000
'
'Total:      102773725
'
'
'Neg 'd Total:    -94608000 + -23414400 + -42275 = -118064675

