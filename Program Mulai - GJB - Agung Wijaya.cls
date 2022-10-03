VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Mulai_Click()
Range("B13").Value = 0 '0
delta_t = Range("B9").Value  '0,1

While Range("B13").Value < 9
Range("B13").Value = Range("B13").Value + delta_t
DoEvents
Wend
End Sub
