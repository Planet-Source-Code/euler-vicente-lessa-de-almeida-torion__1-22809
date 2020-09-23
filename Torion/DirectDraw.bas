Attribute VB_Name = "Module1"
'----------------- API -------------------------
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'----------------- CONSTANTES -------------------------
Public Const SRCCOPY = &HCC0020

'----------------- DIRECT DRAW -------------------------
Public DX           As New DirectX7 'Váriavel Basica do DirectX7
Public DDraw        As DirectDraw7 'Váriavel do DirectDraw
Public DDSPrimary   As DirectDrawSurface7 'DDSPrimary é a surface principal,
                                        'o que está sendo mostrado no monitor
Public DDSBack As DirectDrawSurface7 'DDSBack é o Back buffer,onde são feitas as
                                     'alterações (mudança de quadro de animação e movimentação)
Public DDSDESC As DDSURFACEDESC2
Public Caps As DDSCAPS2 'Caps é uma variavel que armazena algumas flags para criar o backbuffer
Public DDClrKey As DDCOLORKEY 'Declaramos a Color Key, que permite a transparência das imagens
Public DDSMar_tile As DirectDrawSurface7
Public DDSLava_tile As DirectDrawSurface7
Public DDSSeafloor_tile As DirectDrawSurface7
Public DDSGround_tile As DirectDrawSurface7
Public DDSGrama_tile As DirectDrawSurface7
Public DDSNuvem  As DirectDrawSurface7
Public DDSPainel As DirectDrawSurface7
Public DDSTorion As DirectDrawSurface7
Public DDSNumero As DirectDrawSurface7
Public DDSTiro   As DirectDrawSurface7
Public DDSArma1  As DirectDrawSurface7
Public DDSArma2  As DirectDrawSurface7
Public DDSArma3  As DirectDrawSurface7
Public DDSMissil As DirectDrawSurface7
Public DDSExtras As DirectDrawSurface7
Public DDSInimigos As DirectDrawSurface7
Public DDSTiroInimigo As DirectDrawSurface7
Public DDSExplosao0 As DirectDrawSurface7
Public DDSTorionLogo As DirectDrawSurface7
Public Sub destroi_DD()

    DDraw.RestoreDisplayMode
    DDraw.SetCooperativeLevel Form1.hWnd, DDSCL_NORMAL

    Set DDSNuvem = Nothing
    Set DDSPainel = Nothing
    Set DDSTiro = Nothing
    Set DDSArma1 = Nothing
    Set DDSArma2 = Nothing
    Set DDSArma3 = Nothing
    Set DDSMissil = Nothing
    Set DDSExtras = Nothing
    Set DDSInimigos = Nothing
    Set DDSTorion = Nothing
    Set DDSNumero = Nothing
    Set DDSTiroInimigo = Nothing
    Set DDSExplosao0 = Nothing
    Set DDSTorionLogo = Nothing
    Set DDSPedra_tile = Nothing
    Set DDSBack = Nothing
    Set DDSPrimary = Nothing
    Set DDraw = Nothing


End Sub

Public Sub inicializa_DD()
  
  'Aqui criamos o objeto do DirectDraw
  Set DDraw = DX.DirectDrawCreate("")
  'SetCooperativeLevel criará o directdraw no formulário, em modo FullScreen (tela cheia)
  DDraw.SetCooperativeLevel Form1.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
  'Usamos SetDisplayMode para selecionar a resolução 640x480 e 16 bits de cores
  DDraw.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
  
  'Vamos descrever a nova surface a ser criada, que é a principal (que aparece no monitor)
  With DDSDESC
    'Primeiro setamos as Flags, para podermos usar as propriedades BackBufferCount e Caps
    .lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    'No Caps, especificamos mais propriedades da surface, PRIMARYSURFACE pra dizer que ela é primaria (será exibida no monitor)
    'FLIP para podermos usar o backbuffer e COMPLEX pq é necessário sempre que usamos PRIMARYSURFACE
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    'Especificamos agora, no BackBufferCount que esta surface tem apenas 1 backbuffer
    .lBackBufferCount = 1
  End With
  
  'Agora, vamos criar a surface com o comando CreateSurface do DirectDraw, e passamos o argumento DDSDESC, que é a descrição que acabamos de setar
  Set DDSPrimary = DDraw.CreateSurface(DDSDESC)

  'Setamos o Caps para BACKBUFFER, pois agora iremos criar o BackBuffer da Surface Principal
  Caps.lCaps = DDSCAPS_BACKBUFFER
  'Criamos o BackBuffer utilizando o comando GetAttachedSurface da Surface Principal, passando como argumento o Caps que esta setado para BACKBUFFER
  Set DDSBack = DDSPrimary.GetAttachedSurface(Caps)


  'Tudo que for preto na surface será considerado cor transparente
  DDClrKey.high = vbBlack
  DDClrKey.low = vbBlack
  
  
  Set DDSNuvem = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\nuvem.gif", True)
  DDSNuvem.SetColorKey DDCKEY_SRCBLT, DDClrKey
  
  Set DDSPainel = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\painel.gif", True)
  
  Set DDSNumero = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\numeros.gif", True)
  
  Set DDSTorion = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\torion.gif", True)
  DDSTorion.SetColorKey DDCKEY_SRCBLT, DDClrKey
  
  Set DDSTiro = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\tiro.gif", True)
  DDSTiro.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSArma1 = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\arma1.bmp", True)
  DDSArma1.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSArma2 = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\arma2.bmp", True)
  DDSArma2.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSArma3 = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\arma3.gif", True)
  DDSArma3.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSExtras = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\extras.gif", True)
  DDSExtras.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSMissil = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\missel.gif", True)
  DDSMissil.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSInimigos = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\inimigos.gif", True)
  DDSInimigos.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSTiroInimigo = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\tiro_inimigo.bmp", True)
  DDSTiroInimigo.SetColorKey DDCKEY_SRCBLT, DDClrKey

  Set DDSExplosao0 = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\Explosao0.gif", True)
  DDSExplosao0.SetColorKey DDCKEY_SRCBLT, DDClrKey
  
  Set DDSTorionLogo = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\torion.bmp")
    
    

End Sub

Public Function CreateDDSFromBitmap(dd As DirectDraw7, ByVal strFile As String, Optional VideoMem As Boolean) As DirectDrawSurface7
    '
    'Essa função foi retirada do jogo Space Shooter 2K.
    'Programado por Adam "Gollum" Lonnberg
    '
    
    'This function creates a direct draw surface from any valid file format that loadpicture uses, and returns
    'the newly created surface
    
    
    Dim ddsd As DDSURFACEDESC2                                              'Surface description
    Dim dds As DirectDrawSurface7                                           'Created surface
    Dim hdcPicture As Long                                                  'Device context for picture
    Dim hdcSurface As Long                                                  'Device context for surface
    Dim Picture As StdPicture                                                'stdole2 StdPicture object
    
    Set Picture = LoadPicture(strFile)                                      'Load the bitmap
    
    With ddsd                                                               'Fill the surface description
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH                    'Tell Direct Draw that the caps element is valid, the height element is valid, and the width element is valid
        If VideoMem Then                                                    'If the videomem flag is set, then
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY  'Create the surface in video memory
        Else
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY 'Otherwise, creat the surface in system memory
        End If
        .lWidth = Screen.ActiveForm.ScaleX(Picture.Width, vbHimetric, vbPixels)
                                                                            'The width of the surface is set by scaling from the stdpicture objects vbhimetric scale mode to pixels
        .lHeight = Screen.ActiveForm.ScaleY(Picture.Height, vbHimetric, vbPixels)
                                                                            'The height of the surface is set by scaling from the stdpicture objects vbhimetric scale mode to pixels
    End With
    
    
    Set dds = dd.CreateSurface(ddsd)                                        'Create the surface
    hdcPicture = CreateCompatibleDC(ByVal 0&)                               'Create a memory device context
    SelectObject hdcPicture, Picture.Handle                                 'Select the bitmap into this memory device
    dds.restore                                                             'Restore the surface
    hdcSurface = dds.GetDC                                                  'Get the surface's DC
    StretchBlt hdcSurface, 0, 0, ddsd.lWidth, ddsd.lHeight, hdcPicture, 0, 0, Screen.ActiveForm.ScaleX(Picture.Width, vbHimetric, vbPixels), Screen.ActiveForm.ScaleY(Picture.Height, vbHimetric, vbPixels), SRCCOPY
                                                                            'Copy from the memory device to the DirectDrawSurface
    dds.ReleaseDC hdcSurface                                                'Release the surface's DC
    DeleteDC hdcPicture                                                     'Release the memory device context
    Set Picture = Nothing                                                   'Release the picture object
    Set CreateDDSFromBitmap = dds                                           'Sets the function to the newly created direct draw surface
    

End Function

