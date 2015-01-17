VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBJSON Test Form"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2730
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRunJSC 
      Caption         =   "Run JSONScript Program"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   960
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReadJSON 
      Caption         =   "Read JSON Data From File"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdObjToJSON 
      Caption         =   "Test JSON Object"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed


Private Sub cmdObjToJSON_Click()

   Dim p As Object
   
   Dim sInputJson As String
   sInputJson = "{ width: '200', frame: false, height: 130, bodyStyle:'background-color: #ffffcc;',buttonAlign:'right', items: [{ xtype: 'form',  url: '/content.asp'},{ xtype: 'form2',  url: '/content2.asp'}] }"
   
   MsgBox "Input JSON string: " & sInputJson
   
   ' sets p
   Set p = JSON.parse(sInputJson)
   
   MsgBox "Parsed object output: " & JSON.toString(p)
   
   MsgBox "Get Bodystyle data: " & p.Item("bodyStyle")
   
   MsgBox "Get Form Url data: " & p.Item("items").Item(1).Item("url")
   
   
   p.Item("items").Item(1).Add "ExtraItem", "Extra Data Value"
   
   MsgBox "Parsed object output with added item: " & JSON.toString(p)
   
   
End Sub

Private Sub cmdReadJSON_Click()


   Dim p As Object
   
   cd.ShowOpen
   
   If cd.FileName <> "" Then
      Set p = JSON.parse(ReadTextFile(cd.FileName))
      If Not (p Is Nothing) Then
         If JSON.GetParserErrors <> "" Then
            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
         Else
            MsgBox "Base item count: " & p.Count
            MsgBox "JSON toString: " & Left(JSON.toString(p), 1000)
         End If
      Else
         MsgBox "An error occurred parsing " & cd.FileName
      End If
   End If
   
      
End Sub

Private Sub cmdRunJSC_Click()
 
   Dim JSC As New cJSONScript
   
   Dim p As Object
   cd.InitDir = App.Path
   
   cd.ShowOpen
   
   If cd.FileName <> "" Then
      MsgBox JSC.Eval(ReadTextFile(cd.FileName)), vbInformation, "Program Output"
   End If
   
End Sub


Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
   
      handle = FreeFile
      Open sFilePath For Binary As #handle
      ReadTextFile = Space$(LOF(handle))
      Get #handle, , ReadTextFile
      Close #handle
      
   End If
   
End Function


