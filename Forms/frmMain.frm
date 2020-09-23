VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   8490.001
      Top             =   5790
   End
   Begin VB.PictureBox PB 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   0
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   0
      Top             =   0
      Width           =   7395
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8040
      Top             =   5790
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CUSTOMVERTEX
    x As Single         'x in screen space.
    y As Single         'y in screen space.
    z As Single         'normalized z.
    rhw As Single       'normalized z rhw.
    color As Long       'vertex color.
End Type


Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)

Private DX As DirectX8
Private D3D As Direct3D8
Private D3DX As D3DX8     ' Root objects for d3dx
Private D3DDevice As Direct3DDevice8
Private Mode As D3DDISPLAYMODE
Private VertexBuffer As Direct3DVertexBuffer8
Private VertexBuffer2 As Direct3DVertexBuffer8

Private mFont As New StdFont
Private D3DFont As D3DXFont
Private D3DFontDesc As IFont
Private ScreenRect As RECT

Private mX As Long
Private mY As Long

Private FPS As Long
Private FPS2 As Long



Private Sub Form_Load()
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    If D3D Is Nothing Then Exit Sub

    'Retrieve the current display mode
    'by using the Direct3D8.GetAdapterDisplayMode method
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode

    'By filling in the fields of the D3DPRESENT_PARAMETERS object,
    'you can specify how you want your 3-D application to behave.
    'The CreateDevice sample project sets its Windowed member to 1(True),
    'its SwapEffect member to D3DSWAPEFFECT_COPY_VSYNC,
    'and its BackBufferFormat member to mode.Format.
    Dim D3DPP As D3DPRESENT_PARAMETERS
    D3DPP.Windowed = 1
    D3DPP.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    D3DPP.BackBufferFormat = Mode.Format

    'creates the device with the default adapter by using the D3DADAPTER_DEFAULT flag
    'assuming tehre is only one physical adapter installed
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, PB.hWnd, _
                                         D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
    'if you tell the system to use hardware vertex processing by specifying
    'D3DCREATE_HARDWARE_VERTEXPROCESSING, you will see a significant performance gain
    'on video cards that support hardware vertex processing

    If D3DDevice Is Nothing Then Exit Sub
    
    
    

    
    With ScreenRect
        .Top = 0
        .Left = 0
        .Right = PB.ScaleWidth
        .bottom = PB.ScaleHeight
    End With
    
    With mFont
        .Name = "Arial"
        .Size = 18
        .Bold = True
    End With
    
    Set D3DFontDesc = mFont
    Set D3DFont = D3DX.CreateFont(D3DDevice, D3DFontDesc.hFont)

    mX = 200
    mY = 200

End Sub

Private Sub Form_Resize()
    PB.Height = frmMain.ScaleHeight
End Sub

Private Sub PB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mX = x
    mY = y
End Sub

Private Sub Form_Terminate()
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set DX = Nothing
End Sub

Private Sub Render()
    Dim Vertices(2) As CUSTOMVERTEX
    Dim Lines(7) As CUSTOMVERTEX
    Dim VertexSizeInBytes As Long
    Dim VertexSizeInBytes2 As Long
    
    With Vertices(0): .x = 25 + mX: .y = mY: .z = 0.5: .rhw = 1: .color = vbGreen: End With
    With Vertices(1): .x = 50 + mX: .y = 50 + mY: .z = 0.5: .rhw = 1: .color = vbRed: End With
    With Vertices(2): .x = mX: .y = 50 + mY: .z = 0.5: .rhw = 1: .color = vbBlue: End With

    With Lines(0): .x = mX: .y = mY: .z = 0: .color = RGB(255, 0, 0): End With
    With Lines(1): .x = mX + 100: .y = mY: .z = 0: .color = RGB(255, 0, 0): End With
    With Lines(2): .x = mX + 100: .y = mY: .z = 0: .color = RGB(0, 255, 0): End With
    With Lines(3): .x = mX + 100: .y = mY + 100: .z = 0: .color = RGB(0, 255, 0): End With
    With Lines(4): .x = mX + 100: .y = mY + 100: .z = 0: .color = RGB(0, 0, 255): End With
    With Lines(5): .x = mX: .y = mY + 100: .z = 0:  .color = RGB(0, 0, 255): End With
    With Lines(6): .x = mX: .y = mY + 100: .z = 0: .color = RGB(0, 0, 0): End With
    With Lines(7): .x = mX: .y = mY: .z = 0: .color = RGB(0, 0, 0): End With
    

    'Triangle
    VertexSizeInBytes = LenB(Vertices(0))
    Set VertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 3, _
                 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If VertexBuffer Is Nothing Then Exit Sub

    D3DVertexBuffer8SetData VertexBuffer, 0, VertexSizeInBytes * 3, 0, Vertices(0)

    'Lines
    VertexSizeInBytes2 = LenB(Lines(0))
    Set VertexBuffer2 = D3DDevice.CreateVertexBuffer(VertexSizeInBytes2 * 8, _
                 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If VertexBuffer2 Is Nothing Then Exit Sub

    D3DVertexBuffer8SetData VertexBuffer2, 0, VertexSizeInBytes2 * 8, 0, Lines(0)

    'clear screen and set back color
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HFFFFFF, 1#, 0

    ' Begin the scene.
    D3DDevice.BeginScene
    
    'Draw mouse triangle
        ' Rendering of scene objects occur here.
        D3DDevice.SetStreamSource 0, VertexBuffer, VertexSizeInBytes
        
        'let Direct3D know what vertex shader to use
        D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
        
        'render the vertices in the vertex buffer
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
    'End Draw mouse triangle
    'Draw lines
        ' Rendering of scene objects occur here.
        D3DDevice.SetStreamSource 0, VertexBuffer2, VertexSizeInBytes2
        
        'let Direct3D know what vertex shader to use
        D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
        
        'render the vertices in the vertex buffer
        D3DDevice.DrawPrimitive D3DPT_LINELIST, 0, 4
    'end Draw lines

    
        'Draw frame rate as text
        D3DX.DrawText D3DFont, &HFF000000, "Frame rate : " + CStr(FPS2), ScreenRect, DT_LEFT Or DT_TOP

    ' End the scene.
    D3DDevice.EndScene
   
    ' Swap buffers and present primary
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub


Private Sub Timer1_Timer()
    Render
    FPS = FPS + 1
End Sub

Private Sub Timer2_Timer()
    FPS2 = FPS
    FPS = 0
End Sub
