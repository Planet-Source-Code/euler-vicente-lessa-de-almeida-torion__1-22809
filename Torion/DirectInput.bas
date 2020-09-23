Attribute VB_Name = "Module3"
Public DI As DirectInput              ' Nosso Objeto DXInput
Public DITeclado As DirectInputDevice ' Nosso objeto para o teclado
Public DIMouse As DirectInputDevice   ' Nosso objeto para o mouse
Public Sub ler_teclado()

    Dim KeyboardState(0 To 255) As Byte
    
    DITeclado.Acquire
    DITeclado.GetDeviceState 256, KeyboardState(0)
    
    'Teclado
    If (KeyboardState(DIK_SPACE)) <> 0 Then
        atira
    End If
    If (KeyboardState(DIK_UP)) <> 0 Then
        If Module2.torion.y <= 48 Then
            Module2.torion.y = 48
        Else
            Module2.torion.y = Module2.torion.y - Module2.torion.velocidade
        End If
    End If
    If (KeyboardState(DIK_DOWN)) <> 0 Then
        If Module2.torion.y + ALTURA_TORION >= 475 Then
            Module2.torion.y = 480 - ALTURA_TORION
        Else
            Module2.torion.y = Module2.torion.y + Module2.torion.velocidade
        End If
    End If
    If (KeyboardState(DIK_LEFT)) <> 0 Then
        If Module2.torion.x <= 5 Then
            Module2.torion.x = 0
        Else
            Module2.torion.x = Module2.torion.x - Module2.torion.velocidade
        End If
    End If
    If (KeyboardState(DIK_RIGHT)) <> 0 Then
        If Module2.torion.x + LARGURA_TORION >= 640 Then
            Module2.torion.x = 640 - LARGURA_TORION
        Else
            Module2.torion.x = Module2.torion.x + Module2.torion.velocidade
        End If
    End If

    

End Sub

Public Sub destroi_DI()

    Set DITeclado = Nothing
    Set DIMouse = Nothing
    Set DI = Nothing
    
End Sub
Public Sub inicializa_DI()

    Set DI = DX.DirectInputCreate()
    
    ' Cria o dispositivo de entrada (teclado)
    Set DITeclado = DI.CreateDevice("GUID_SysKeyboard")
    ' Associa a estrutura de informação para o teclado (DIKEYBOARDSTATE)
    DITeclado.SetCommonDataFormat DIFORMAT_KEYBOARD
    ' Se estamos sem o foco , setamos para modo não exclusivo
    DITeclado.SetCooperativeLevel Form1.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    ' "Adquire" o teclado
    DITeclado.Acquire
    ' Cria o dispositivo de entrada (mouse)
    Set DIMouse = DI.CreateDevice("GUID_SysMouse")
    ' Associa a estrutura de informação para o mouse (DIMOUSESTATE)
    DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
    ' Se estamos sem o foco , setamos para modo não exclusivo
    DIMouse.SetCooperativeLevel Form1.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    ' "Adquire" o mouse
    DIMouse.Acquire


End Sub

