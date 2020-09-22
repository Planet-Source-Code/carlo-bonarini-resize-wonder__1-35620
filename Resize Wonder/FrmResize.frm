VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmResize 
   Caption         =   "Resize Wonder by fingolfin@inwind.it"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicResizeTV 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   8000
      Left            =   1920
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7965
      ScaleWidth      =   45
      TabIndex        =   0
      ToolTipText     =   "Click to resize"
      Top             =   0
      Width           =   70
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   7995
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   14102
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGrid 
      Height          =   4005
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   20
      Cols            =   3
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      BandDisplay     =   1
      RowSizingMode   =   1
      MouseIcon       =   "FrmResize.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGrid 
      Height          =   4005
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   3960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   20
      Cols            =   3
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      BandDisplay     =   1
      RowSizingMode   =   1
      MouseIcon       =   "FrmResize.frx":031A
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "FrmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Resize As Boolean

Private Sub Form_Resize()
'I don 't want the form to be different from these values
On Error Resume Next:
If FrmResize.Height < 8000 Then
FrmResize.Height = 8000
End If
If FrmResize.Width < 8000 Then
FrmResize.Width = 8000
End If
    
On Error Resume Next:
If Error = 380 Then Exit Sub
TV.Height = Me.Height - 700
PicResizeTV.Height = Me.Height - 700

For A = 0 To 2
'this settings are only an example, you can add any control you like with settings you like
FlexGrid.Item(A).Width = Me.Width - TV.Width - 200 - PicResizeTV.Width
FlexGrid.Item(A).Height = TV.Height / 2
FlexGrid.Item(A).ColWidth(0) = FlexGrid.Item(A).Width / 25
FlexGrid.Item(A).ColWidth(1) = FlexGrid.Item(A).Width / 4
FlexGrid.Item(A).ColWidth(2) = FlexGrid.Item(A).ColWidth(1) * 4
FlexGrid.Item(A).TextMatrix(0, 1) = "Resize  Wonder !!!"
FlexGrid.Item(A).TextMatrix(0, 2) = "Click on red line to resize flexgrids!!!"
For B = 1 To 20
FlexGrid.Item(A).TextMatrix(B, 0) = B  'this only generates numbers on first column
FlexGrid.Item(A).TextMatrix(B, 1) = "Click on red line to resize flexgrids!!!"
FlexGrid.Item(A).TextMatrix(B, 2) = "Both flexgrids will be resized without exiting from the form!!!"
TV.Nodes.Add , tvwFirst, , FlexGrid.Item(0).TextMatrix(A, 2)
Next B
Next A


End Sub

Private Sub PicResizeTV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Resize = True 'the user has clicked the pic = the user wants to resize the treeview
PicResizeTV.BackColor = &H80000012

End Sub

Private Sub PicResizeTV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next:
If Resize = True Then ' the user has clicked the pic = the user wants to resize the treeview
'sets  treeview width when moving the pic
TV.Width = TV.Width + X
For I = 0 To 2
'this moves the pic exactly where the treeview is moved
PicResizeTV.Move TV.Width, PicResizeTV.Top, PicResizeTV.Width, PicResizeTV.Height

'this contract the flexgrid, so he does't exit from the form
'(increase the value 200 if you want to add more space between
'flexgrid and right size of form, even if 200 is optimal for showing correctly scroll bars)
FlexGrid.Item(I).Width = Me.Width - TV.Width - 200

'this moves the flexgrid
 FlexGrid.Item(I).Move TV.Width + 100, FlexGrid.Item(I).Top, FlexGrid.Item(I).Width, FlexGrid.Item(I).Height


Me.Caption = "Resize Wonder by fingolfin@inwind.it" & Space(1) & "Resizing Treeview -" _
 & Space(1) & "Width:" & Space(1) & TV.Width & Space(1) & ";" & Space(1) & "Moving FlexGrids -" _
  & Space(1) & TV.Width + 100 & Space(8) & Date & Space(1) & Time
  
PicResizeTV.ToolTipText = "Treeview width:" & Space(1) & TV.Width & Space(1) & _
  ";" & Space(1) & "FlexGrids width:" & Space(1) & FlexGrid.Item(I).Width
Next I

'if you add array of control
'you can resize as much controls as you like with this routine only!
'remember to edit the for..next loop in this routine and also at form resize event if
'you need to resize all controls at startup
'Note that you can add this routine to the mousemove event of the treeview (or listview, or button, etc)
'but the treeview needs to be clicked so if the mouse pointer is the default arrow the
'user don't understands if he can resize the control or not even if you tell it on a tooltiptext
'and if the mouse pointer is the sizeWEpointer the treeview looks not very good, so I added this picture
'that resizes the control and can be put right or left or top or bottom any control that needs mouse clicking

'please leave comment on PSC or send me an email: fingolfin@inwind.it

'**********************************************************************************
'PS.: sorry for my terrible english, I'm from Italy!!!

End If
End Sub


Private Sub PicResizeTV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Resize = False 'the user has lift up mouse button = the user no longer wants to resize the treeview
PicResizeTV.BackColor = &HFF&
Me.Caption = "Resize Wonder by fingolfin@inwind.it"
End Sub


