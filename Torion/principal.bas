Attribute VB_Name = "Module2"
'---------------- API -----------------
Public Declare Function IntersectRect Lib "user32" (ByRef r As RECT, ByRef r2 As RECT, ByRef r3 As RECT) As Long

'---------------- TIPOS -----------------
Type tipo_tile
    x       As Integer
    y       As Integer
    Existe  As Boolean
    YoffSet As Integer
End Type

Type tipo_nuvem
    x           As Integer
    y           As Integer
    Existe      As Boolean
    velocidade  As Single
End Type

Type tipo_torion
    x           As Single
    y           As Single
    velocidade  As Single
    arma_atual  As Single
    frame_atual As Single
    num_frames  As Integer
    tiros_laser As Single
    vidas       As Single
    invencibil  As Integer
End Type

Type tipo_tiro
    x           As Single
    y           As Single
    poder       As Single
    Existe      As Boolean
End Type

Type tipo_dados_arma
    numero_frames   As Single
    poder           As Single
    largura         As Integer
    altura          As Integer
End Type

Type tipo_arma
    x               As Integer
    y               As Integer
    movimento_x     As Integer
    movimento_y     As Integer
    Existe          As Boolean
    frame_atual     As Single
    alvo            As Single
    angulo          As Single
End Type

Type tipo_inimigo
    x               As Single
    y               As Single
    Existe          As Boolean
    formacao        As Single
    tipo            As Single
    danos           As Single
    angulo          As Single
    x_inicial       As Single
    extra           As Integer
End Type

Type tipo_tiro_inimigo
    x           As Single
    y           As Single
    angulo      As Double
    Existe      As Boolean
End Type


Type tipo_dados_inimigo
    resistencia     As Single
    atira           As Boolean
    pontos          As Integer
End Type

Type tipo_editor_fase
    formacao        As Single
    inicia          As Long 'starta a formação quando o contador do jogo chegar nesse ponto
    tipo_inimigo    As Single
    num_inimigos    As Single
    extra           As Integer
End Type

Type tipo_jogo
    placar              As Long
    fase_atual          As Single
    contador            As Long
    indice_editor_fase  As Single
    nivel               As Single '1-facil;2-normal;3-difícil
    vida_extra          As Long 'indica qual o placar tem que atingir para ganhar vida extra
    status              As Integer '1-abertura;2-jogando;3-tela congratulações;4-fim de fase;5-fim de jogo;6-pause
    numero_fases        As Integer
    item_abertura       As Integer
    FPS                 As Boolean
    recorde             As Double
End Type

Type tipo_extras
    x           As Integer
    y           As Integer
    extra       As Integer
    Existe      As Boolean
End Type

Type tipo_explosao                      'Define valores para as explosões na tela
    x As Single                         'X da exposao
    y As Single                         'Y da explosao
    Existe As Boolean                   'Determina se o objeto existe
    frame_atual As Integer              'Indice que determinará qual frame será exibido
End Type


'---------------- VARIAVEIS AUXILIARES -----------------
Public explosao(14)             As tipo_explosao
Public jogo                     As tipo_jogo
Public EmptyRect                As RECT
Public Tiles(108)               As tipo_tile
Public nuvem(3)                 As tipo_nuvem
Public torion                   As tipo_torion 'É a nossa nave
Public tiro(17)                 As tipo_tiro
Public arma(2)                  As tipo_arma
Public tipo_arma(1 To 5)        As tipo_dados_arma
Public tipo_inimigo(7)          As tipo_dados_inimigo
Public inimigo(20)              As tipo_inimigo
Public tiro_inimigo(15)         As tipo_tiro_inimigo
Public editor_fase()            As tipo_editor_fase
Public extras                   As tipo_extras
Public opcao                    As Integer
Private lTimer As Long                  ' stores the last time tick
Private lFPS As Integer                 ' stores the last number of frames per second
Private lFPSCounter As Integer          ' counts the frames per second
Public tmpString                As String

'--------------- CONSTANTES ----------------------
Public Const LARGURA_TORION = 30             'Largura da nave do jogador
Public Const ALTURA_TORION = 35              'altura da nave do jogador
Public Const placar_vida_extra = 100000      'a cada 100000 pontos uma vida extra
Public Sub InitFPS()
    lTimer = DX.TickCount
End Sub

Public Function FPS() As Long
    
    If lTimer + 1000 <= DX.TickCount Then
        '
        ' store the current tick count
        '
        lTimer = DX.TickCount
        '
        ' increase the FPS counter, because this was also a frame, and store it in
        ' our FPS variable
        '
        lFPS = lFPSCounter + 1
        '
        ' reset FPS counter
        '
        lFPSCounter = 0
    Else
        '
        ' less than one second => increase frame counter
        '
        lFPSCounter = lFPSCounter + 1
    End If
    FPS = lFPS
    
End Function


Public Sub atualiza_nave_torion()
    '
    'Essa sub atualiza os dados da nave do jogador
    '
    Dim SrcRect As RECT     'retangulo fonte que representará o frame corrente

    'Preenche os valores do retangulo que conterá o frame corrente
    With SrcRect
        .Top = 0              'Top é o espaco (em pixels) da imagem até o topo da tela
        .Bottom = .Top + ALTURA_TORION     'Bottom é a altura da imagem a ser "pegada"
        .Left = LARGURA_TORION * torion.frame_atual   'Left é o espaco (em pixels) da imagem até a esquerda da tela
        .Right = .Left + LARGURA_TORION 'Right é a largura da imagem a ser "pegada"
    End With
    
    'Testa se for o ultimo frame volta para o 1º
    If torion.frame_atual >= torion.num_frames Then
        torion.frame_atual = 0
    Else
        torion.frame_atual = torion.frame_atual + 1           'incrementa o frame
    End If
   
    
    'blit a nave para o Back Buffer com efeito de transparencia
    DDSBack.BltFast torion.x, torion.y, DDSTorion, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY


End Sub

Public Sub atualiza_tile(surface As DirectDrawSurface7)

    Dim TileRect As RECT
    Dim tile_cont   As Integer

    With TileRect
        .Top = 0
        .Bottom = .Top + 48
        .Left = 0
        .Right = .Left + 64
    End With
    
    'Rolagem de fundo do tile
    For tile_cont = 0 To 108
        If Not Tiles(tile_cont).Existe Then
            Tiles(tile_cont).Existe = True
        End If

        If Tiles(tile_cont).y + 48 > 480 Then
            Tiles(tile_cont).YoffSet = Tiles(tile_cont).YoffSet + 1
            TileRect.Top = 0: TileRect.Bottom = 48 - Tiles(tile_cont).YoffSet: TileRect.Left = 0: TileRect.Right = 64
        Else
            TileRect.Top = 0: TileRect.Bottom = 48: TileRect.Left = 0: TileRect.Right = 64
            Tiles(tile_cont).YoffSet = 0
        End If
        If Tiles(tile_cont).y > 480 Then
            Tiles(tile_cont).Existe = False
            Tiles(tile_cont).y = 1
        End If
        If Tiles(tile_cont).Existe Then
            DDSBack.BltFast Tiles(tile_cont).x, Tiles(tile_cont).y, surface, TileRect, DDBLTFAST_WAIT
        End If
        Tiles(tile_cont).y = Tiles(tile_cont).y + 1
    Next


End Sub

Public Sub atualiza_painel()

    Dim PainelRect   As RECT
    Dim VelocidRect  As RECT
    Dim NumeroRect   As RECT
    Dim extrasRect   As RECT
    Dim cont         As Single
    Dim x            As Single
    Dim aux          As Single
    
    'Painel
    DDSBack.BltFast 0, 0, DDSPainel, PainelRect, DDBLTFAST_WAIT
    
    
    'Mostrador de velocidade
    VelocidRect.Top = 15: VelocidRect.Bottom = 25
    VelocidRect.Left = 585: VelocidRect.Right = VelocidRect.Left + 40
    DDSBack.BltColorFill VelocidRect, RGB(100, 200, 200)
    
    VelocidRect.Top = 16: VelocidRect.Bottom = 24
    VelocidRect.Left = 585: VelocidRect.Right = VelocidRect.Left + torion.velocidade * 5
    DDSBack.BltColorFill VelocidRect, RGB(255, 0, 0)
    
    Form1.FontName = "Impact"
    Form1.FontSize = 8
    Form1.FontBold = False
    DDSBack.SetForeColor RGB(255, 255, 250)
    DDSBack.SetFont Form1.Font
    DDSBack.DrawText 514, 12, "V E L O C I D A D E", False
    
    DDSBack.DrawText 514, 23, "I N V E N C I B I L I D A D E: " & torion.invencibil, False
    
    
    'Mostrador de vidas
    Form1.FontName = "Impact"
    Form1.FontSize = 16
    Form1.FontBold = True
    DDSBack.SetForeColor RGB(255, 255, 250)
    DDSBack.SetFont Form1.Font
    
    DDSBack.DrawText 488, 12, Str(torion.vidas), False
    
    
    If jogo.status = 2 Then
        If torion.invencibil <> 0 And torion.invencibil <= 5 Then
            Form1.FontSize = 48
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.SetFont Form1.Font
            DDSBack.DrawText 300, 240, torion.invencibil, False
        End If
    End If
    'Placar
    For cont = Len(Trim(Str(jogo.placar))) To 1 Step -1
        With NumeroRect
            .Top = 0
            .Bottom = 22
            .Left = VBA.Val(Mid(Trim(Str(jogo.placar)), cont, 1)) * 12
            .Right = .Left + 12
        End With
        x = 150 - aux * 12
        DDSBack.BltFast x, 14, DDSNumero, NumeroRect, DDBLTFAST_WAIT
        aux = aux + 1
    Next cont
    
    'Arma atual
    If torion.arma_atual <> 0 Then
        With extrasRect
            .Top = 0
            .Bottom = 32
            .Left = 16 * (torion.arma_atual - 1)
            .Right = .Left + 16
        End With
        DDSBack.BltFast 258, 9, DDSExtras, extrasRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    
End Sub

Public Sub atualiza_nuvem()
    
    Dim NuvemRect As RECT
    Dim nuvem_cont  As Single
    Dim TempX      As Integer
    Dim TempY      As Integer
    
    For nuvem_cont = 0 To 3
        If Not nuvem(nuvem_cont).Existe Then
            nuvem(nuvem_cont).Existe = True
            nuvem(nuvem_cont).x = (400 * Rnd) + 100
            nuvem(nuvem_cont).velocidade = (3 * Rnd) + 3
            nuvem(nuvem_cont).y = 0
        End If

        NuvemRect.Top = 0: NuvemRect.Bottom = 48: NuvemRect.Left = 0: NuvemRect.Right = NuvemRect.Left + 68
        
        If nuvem(nuvem_cont).x + NuvemRect.Right > 640 Then
          TempX = Abs(nuvem(nuvem_cont).x - 640)
          NuvemRect.Right = TempX
        End If

        If nuvem(nuvem_cont).y + NuvemRect.Bottom > 480 Then
          TempY = Abs(nuvem(nuvem_cont).y - 480)
          NuvemRect.Bottom = TempY
        End If
        
        If nuvem(nuvem_cont).x < 0 Then
          TempX = Abs(nuvem(nuvem_cont).x)
          NuvemRect.Left = TempX
        End If
        
        If nuvem(nuvem_cont).y < 0 Then
          TempY = Abs(nuvem(nuvem_cont).y)
          NuvemRect.Top = TempY
        End If

        If nuvem(nuvem_cont).x >= 0 And nuvem(nuvem_cont).y >= 0 Then
          DDSBack.BltFast nuvem(nuvem_cont).x, nuvem(nuvem_cont).y, DDSNuvem, NuvemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Else
          If nuvem(nuvem_cont).x < 0 Then DDSBack.BltFast 0, nuvem(nuvem_cont).y, DDSNuvem, NuvemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If nuvem(nuvem_cont).y < 0 Then DDSBack.BltFast nuvem(nuvem_cont).x, 0, DDSNuvem, NuvemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If nuvem(nuvem_cont).x < 0 And nuvem(nuvem_cont).y < 0 Then DDSBack.BltFast 0, 0, DDSNuvem, NuvemRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
        If nuvem(nuvem_cont).y > 480 Then
            nuvem(nuvem_cont).Existe = False
            nuvem(nuvem_cont).y = 0
        End If
        nuvem(nuvem_cont).y = nuvem(nuvem_cont).y + nuvem(nuvem_cont).velocidade
    Next
    
End Sub
Public Function ChecaColisaoPIXEL(Surface1 As DirectDrawSurface7, RectSurface1 As RECT, RectSurface2 As RECT, Surface2 As DirectDrawSurface7) As Boolean
    '
    'Essa função foi adaptada do projeto VB Open Source
    'de autoria de Claudio Lins.
    '

    Dim blnPPCollision          As Boolean
    Dim RectColidido            As RECT
    Dim RectSurface1Colidido    As RECT
    Dim RectSurface2Colidido    As RECT
    Dim bitSurface1()           As Byte
    Dim bitSurface2()           As Byte
    Dim DDColorK                As DDCOLORKEY
    
    blnPPCollision = False

    'Checa a colisão
     If IntersectRect(RectColidido, RectSurface1, RectSurface2) Then
        'se chegou aqui colidiram os retângulos
        'Pega os retângulos colididos de ambas superficies
        With RectSurface1Colidido
            .Top = RectColidido.Top - RectSurface1.Top
            .Bottom = RectColidido.Bottom - RectSurface1.Top
            .Left = RectColidido.Left - RectSurface1.Left
            .Right = RectColidido.Right - RectSurface1.Left
        End With
        With RectSurface2Colidido
            .Top = RectColidido.Top - RectSurface2.Top
            .Bottom = RectColidido.Bottom - RectSurface2.Top
            .Left = RectColidido.Left - RectSurface2.Left
            .Right = RectColidido.Right - RectSurface2.Left
        End With
    
        'Determina a largura a altura da area colidida
        intWidth = RectColidido.Right - RectColidido.Left - 1
        intHeight = RectColidido.Bottom - RectColidido.Top - 1
    
        'necessário ...
        Surface1.Lock RectSurface1Colidido, DDSDESC, DDLOCK_READONLY Or DDLOCK_WAIT, 0
        Surface1.GetLockedArray bitSurface1
    
        Surface2.Lock RectSurface2Colidido, DDSDESC, DDLOCK_READONLY Or DDLOCK_WAIT, 0
        Surface2.GetLockedArray bitSurface2
    
    
        'checa pixel por pixel se está colidindo , se o pixel do rect for
        'diferente da cor de transparência ,etão colidiu ...
        For i = 0 To intWidth
            For j = 0 To intHeight
                'se ambas supercies não são transparentes neste pixel ...
                If (bitSurface1(i + RectSurface1Colidido.Left, j + RectSurface1Colidido.Top) <> CByte(DDColorK.high)) And (bitSurface2(i + RectSurface2Colidido.Left, j + RectSurface2Colidido.Top) <> CByte(DDColorK.high)) Then blnPPCollision = True
                '... então temos colisão!!
                If blnPPCollision = True Then Exit For
            Next j
            If blnPPCollision = True Then Exit For
        Next i
    
        'Destrava as superficies ...
        Surface1.Unlock RectSurface1Colidido
        Surface2.Unlock RectSurface2Colidido
    End If
   
    'retorna se houve ou não colisão ...
    If blnPPCollision = True Then
        ChecaColisaoPIXEL = True
    Else
        ChecaColisaoPIXEL = False
    End If

End Function

Public Sub atualiza_extras()
    
    Dim Ret_extras  As RECT
    Dim TempX       As Integer
    Dim TempY       As Integer
    
    If extras.Existe Then
        With Ret_extras
            .Top = 0
            .Bottom = 32
            .Left = (extras.extra - 1) * 16
            .Right = .Left + 16
        End With
        
        extras.y = extras.y + 2
        
        If extras.x + 16 > 640 Then
          TempX = Abs(extras.x - 640)
          Ret_extras.Right = TempX
        End If
    
        If extras.y + Ret_extras.Bottom > 480 Then
          TempY = Abs(extras.y - 480)
          Ret_extras.Bottom = TempY
        End If
        
        If extras.x < 0 Then
          TempX = Abs(extras.x)
          Ret_extras.Left = TempX
        End If
        
        If extras.y < 0 Then
          TempY = Abs(extras.y)
          Ret_extras.Top = TempY
        End If
    
    
        If extras.x >= 0 And extras.y >= 0 Then
          DDSBack.BltFast extras.x, extras.y, DDSExtras, Ret_extras, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Else
          If extras.x < 0 Then DDSBack.BltFast 0, extras.y, DDSExtras, Ret_extras, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If extras.y < 0 Then DDSBack.BltFast extras.x, 0, DDSExtras, Ret_extras, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If extras.x < 0 And extras.y < 0 Then DDSBack.BltFast 0, 0, DDSExtras, Ret_extras, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
        
        If extras.y > 480 Then
            extras.y = 0
            extras.Existe = False
        End If
    End If

End Sub
Public Sub vida_extra()

    If jogo.placar >= jogo.vida_extra Then
        torion.vidas = torion.vidas + 1
        jogo.vida_extra = jogo.vida_extra + placar_vida_extra
        dsVidaExtra.SetCurrentPosition 0
        dsVidaExtra.Play DSBPLAY_DEFAULT
    End If

End Sub

Public Sub testa_colisoes()

    Dim cont_loop_inimigo       As Integer
    Dim cont_loop_tiro          As Integer
    Dim Retangulo_torion        As RECT
    Dim Retangulo_inimigo       As RECT
    Dim Retangulo_tiro          As RECT
    Dim RectColidido            As RECT
    Dim Retangulo_extras        As RECT
    Dim aux                     As Integer
    
    'Lembrete : As posições (X,Y) correspondem ao canto superior esquerdo da figura
    With Retangulo_torion 'define o retangulo da nossa nave
        .Top = torion.y  'o Topo da nave corresponde ao Y
        .Bottom = .Top + ALTURA_TORION 'A parte inferior é o topo + a altura da nave
        .Left = torion.x  'A esquerda corresponde ao X
        .Right = .Left + LARGURA_TORION 'A direita é a posição X + a largura da nave
    End With

    With Retangulo_extras
        .Top = extras.y
        .Bottom = .Top + 32
        .Left = extras.x
        .Right = .Left + 16
    End With
    'testa se nos pegamos algum armamento extra
    If extras.Existe Then
        If IntersectRect(RectColidido, Retangulo_extras, Retangulo_torion) Then
            jogo.placar = jogo.placar + 50
            vida_extra 'testa se vai ganhar uma vida extra
            dsExtra.SetCurrentPosition 0 'seta a posição do buffer para 0
            dsExtra.Play DSBPLAY_DEFAULT 'Som da explosao
            extras.Existe = False
            If extras.extra <= 5 Then 'se for arma especial
                If extras.extra = 4 Then
                    arma(0).angulo = 0
                    arma(1).angulo = 90
                    arma(2).angulo = 180
                End If
                torion.arma_atual = extras.extra
                For aux = 0 To UBound(arma)
                    arma(aux).Existe = False
                Next aux
            ElseIf extras.extra = 6 Then 'invencibilidade temporaria
                torion.invencibil = 15 '15 segundos de invencibilidade
                torion.num_frames = 4 'Para dar o efeito da nave piscando
            ElseIf extras.extra = 7 Then 'tiro extra
                If torion.tiros_laser < 3 Then
                    torion.tiros_laser = torion.tiros_laser + 1
                End If
            ElseIf extras.extra = 8 Then
                If torion.velocidade < 8 Then
                    torion.velocidade = torion.velocidade + 1
                End If
            End If
        End If
    End If

    For cont_loop_inimigo = 0 To UBound(inimigo) 'laço atraves dos inimigos
        If inimigo(cont_loop_inimigo).Existe Then
            With Retangulo_inimigo
                .Top = inimigo(cont_loop_inimigo).y
                .Bottom = .Top + 46
                .Left = inimigo(cont_loop_inimigo).x
                .Right = .Left + 35
            End With
            'Testa se o inimigo bateu em Torion
            If torion.invencibil = 0 Then 'está vulneravel
                If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_torion, DDSTorion) Then
                    inimigo(cont_loop_inimigo).Existe = False
                    preenche_vetor_explosao torion.x, torion.y
                    torion.vidas = torion.vidas - 1
                    'Se Vc for atingido, dá um tempo de invencibilidade para não
                    'correr o risco de que assim que voltar ser atingido novamente.
                    torion.invencibil = 7
                    torion.arma_atual = 1 'volta para a arma  inicial
                    torion.tiros_laser = 1
                    torion.num_frames = 4 'Para dar o efeito da nave piscando
                    If torion.vidas = 0 Then
                        'fim de jogo
                        para_midi
                        jogo.status = 5
                        editor_fase(1).inicia = jogo.contador
                    End If
                End If
            End If
            For cont_loop_tiro = 0 To UBound(tiro) 'Laço através dos tiros
                If tiro(cont_loop_tiro).Existe Then
                    With Retangulo_tiro 'Pego retangulo desse tiro
                        .Top = tiro(cont_loop_tiro).y
                        .Bottom = .Top + 10
                        .Left = tiro(cont_loop_tiro).x
                        .Right = .Left + 4
                    End With
                    'Testa se o inimigo foi atingido por algum tiro
                    If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_tiro, DDSTiro) Then
                        tiro(cont_loop_tiro).Existe = False
                        inimigo(cont_loop_inimigo).danos = inimigo(cont_loop_inimigo).danos + tiro(cont_loop_tiro).poder
                        If inimigo(cont_loop_inimigo).danos >= tipo_inimigo(inimigo(cont_loop_inimigo).tipo).resistencia Then
                            preenche_vetor_explosao inimigo(cont_loop_inimigo).x, inimigo(cont_loop_inimigo).y
                            inimigo(cont_loop_inimigo).Existe = False
                            jogo.placar = jogo.placar + tipo_inimigo(inimigo(cont_loop_inimigo).tipo).pontos
                            vida_extra 'testa se vai ganhar uma vida extra
                            If Not extras.Existe And inimigo(cont_loop_inimigo).extra <> 0 Then
                                extras.x = inimigo(cont_loop_inimigo).x
                                extras.y = inimigo(cont_loop_inimigo).y
                                extras.extra = inimigo(cont_loop_inimigo).extra
                                extras.Existe = True
                            End If
                        End If
                    End If
                End If
            Next cont_loop_tiro
            If torion.arma_atual <> 0 Then
                For cont_loop_tiro = 0 To UBound(arma) 'Laço através dos armas especiais
                    If arma(cont_loop_tiro).Existe Then
                        With Retangulo_tiro 'Pego retangulo desse tiro
                            .Top = arma(cont_loop_tiro).y
                            .Bottom = .Top + tipo_arma(torion.arma_atual).altura
                            .Left = arma(cont_loop_tiro).x
                            .Right = .Left + tipo_arma(torion.arma_atual).largura
                        End With
                        'Testa se o inimigo foi atingido por alguma arma especial
                        If torion.arma_atual = 1 Then
                            If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_tiro, DDSArma1) Then
                                atualiza_valores_colisao cont_loop_tiro, cont_loop_inimigo
                            End If
                        ElseIf torion.arma_atual = 2 Then
                            If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_tiro, DDSArma2) Then
                                atualiza_valores_colisao cont_loop_tiro, cont_loop_inimigo
                            End If
                        ElseIf torion.arma_atual = 3 Then
                            If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_tiro, DDSArma3) Then
                                atualiza_valores_colisao cont_loop_tiro, cont_loop_inimigo
                            End If
                        ElseIf torion.arma_atual = 4 Then
                            If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_tiro, DDSArma3) Then
                                atualiza_valores_colisao cont_loop_tiro, cont_loop_inimigo
                            End If
                        ElseIf torion.arma_atual = 5 Then
                            If ChecaColisaoPIXEL(DDSInimigos, Retangulo_inimigo, Retangulo_tiro, DDSMissil) Then
                                atualiza_valores_colisao cont_loop_tiro, cont_loop_inimigo
                            End If
                        End If
                    End If
                Next cont_loop_tiro
            End If
        End If
    Next cont_loop_inimigo

    For cont_loop_inimigo = 0 To UBound(tiro_inimigo) 'laço atraves dos tiros dos inimigos
        If tiro_inimigo(cont_loop_inimigo).Existe Then
            With Retangulo_inimigo 'retangulo do tiro do inimigo
                .Top = tiro_inimigo(cont_loop_inimigo).y
                .Bottom = .Top + 8
                .Left = tiro_inimigo(cont_loop_inimigo).x
                .Right = .Left + 8
            End With
            'Testa se o tiro do inimigo atingiu Torion
            If torion.invencibil = 0 Then
                If ChecaColisaoPIXEL(DDSTiroInimigo, Retangulo_inimigo, Retangulo_torion, DDSTorion) Then
                    tiro_inimigo(cont_loop_inimigo).Existe = False
                    preenche_vetor_explosao torion.x, torion.y
                    'Se Vc for atingido, dá um tempo de invencibilidade para não
                    'correr o risco de que assim que voltar ser atingido novamente.
                    torion.invencibil = 7
                    torion.arma_atual = 1 'volta para a arma  inicial
                    torion.tiros_laser = 1
                    torion.num_frames = 4 'Para dar o efeito da nave piscando
                    torion.vidas = torion.vidas - 1
                    If torion.vidas = 0 Then
                        'fim de jogo
                        para_midi
                        jogo.status = 5
                        editor_fase(1).inicia = jogo.contador
                    End If
                End If
            End If
            If torion.arma_atual = 4 Then 'o campo de força nos protejo dos tiros inimigos
                For cont_loop_tiro = 0 To UBound(arma) 'Laço através dos armas especiais
                    If arma(cont_loop_tiro).Existe Then
                        With Retangulo_tiro 'Pego retangulo desse tiro
                            .Top = arma(cont_loop_tiro).y
                            .Bottom = .Top + tipo_arma(torion.arma_atual).altura
                            .Left = arma(cont_loop_tiro).x
                            .Right = .Left + tipo_arma(torion.arma_atual).largura
                        End With
                        If IntersectRect(RectColidido, Retangulo_inimigo, Retangulo_tiro) Then
                            tiro_inimigo(cont_loop_inimigo).Existe = False
                        End If
                    End If
                Next cont_loop_tiro
            End If
        End If
    Next cont_loop_inimigo
    
End Sub
Public Sub atualiza_valores_colisao(cont_loop_tiro As Integer, cont_loop_inimigo As Integer)

    arma(cont_loop_tiro).Existe = False
    inimigo(cont_loop_inimigo).danos = inimigo(cont_loop_inimigo).danos + tipo_arma(torion.arma_atual).poder
    If inimigo(cont_loop_inimigo).danos >= tipo_inimigo(inimigo(cont_loop_inimigo).tipo).resistencia Then
        preenche_vetor_explosao inimigo(cont_loop_inimigo).x, inimigo(cont_loop_inimigo).y
        inimigo(cont_loop_inimigo).Existe = False
        jogo.placar = jogo.placar + tipo_inimigo(inimigo(cont_loop_inimigo).tipo).pontos
        vida_extra 'testa se vai ganhar uma vida extra
        If Not extras.Existe And inimigo(cont_loop_inimigo).extra <> 0 Then
            extras.x = inimigo(cont_loop_inimigo).x
            extras.y = inimigo(cont_loop_inimigo).y
            extras.extra = inimigo(cont_loop_inimigo).extra
            extras.Existe = True
        End If
    End If
    arma(cont_loop_tiro).Existe = False

End Sub
Public Sub reseta_jogo()
    Dim a As Single
    Dim strtemp As String
    
    para_midi
    
    InitFPS
    
    'Define valores iniciais de Torion
    torion.x = 300
    torion.y = 400
    torion.frame_atual = 0
    torion.velocidade = 4
    torion.vidas = 3
    
    torion.tiros_laser = 1
    torion.arma_atual = 1
    torion.invencibil = 0
    torion.num_frames = 2
    
    'Recorde atual
    file1 = FreeFile
    Open App.Path & "\recordes.txt" For Input As #file1
    Input #file1, strtemp
    
    jogo.recorde = Val(Mid(strtemp, InStr(1, strtemp, ";") + 1, Len(strtemp)))
    jogo.placar = 0
    jogo.fase_atual = 1
    jogo.contador = 0
    jogo.indice_editor_fase = 0
    jogo.vida_extra = placar_vida_extra
    jogo.status = 2 'jogando
    jogo.numero_fases = 5
    
    Close #file1

    inicializa_fase



End Sub
Public Sub define_extras()

    'Define valores das armas
    tipo_arma(1).numero_frames = 2
    tipo_arma(1).poder = 2
    tipo_arma(1).altura = 15
    tipo_arma(1).largura = 12

    tipo_arma(2).numero_frames = 1
    tipo_arma(2).poder = 3
    tipo_arma(2).altura = 10
    tipo_arma(2).largura = 30
    
    tipo_arma(3).numero_frames = 7
    tipo_arma(3).poder = 5
    tipo_arma(3).altura = 16
    tipo_arma(3).largura = 18

    tipo_arma(4).numero_frames = 1
    tipo_arma(4).poder = 4
    tipo_arma(4).altura = 16
    tipo_arma(4).largura = 18
    
    tipo_arma(5).numero_frames = 2
    tipo_arma(5).poder = 2
    tipo_arma(5).altura = 25
    tipo_arma(5).largura = 27.5


    'define os inimigos
    tipo_inimigo(0).atira = True
    tipo_inimigo(0).resistencia = 4
    tipo_inimigo(0).pontos = 1000
    
    tipo_inimigo(1).atira = False
    tipo_inimigo(1).resistencia = 4
    tipo_inimigo(1).pontos = 1000
    
    tipo_inimigo(2).atira = True
    tipo_inimigo(2).resistencia = 4
    tipo_inimigo(2).pontos = 1000
    
    tipo_inimigo(3).atira = False
    tipo_inimigo(3).resistencia = 1
    tipo_inimigo(3).pontos = 500
    
    tipo_inimigo(4).atira = True
    tipo_inimigo(4).resistencia = 6
    tipo_inimigo(4).pontos = 2000

    tipo_inimigo(5).atira = False
    tipo_inimigo(5).resistencia = 10
    tipo_inimigo(5).pontos = 3000

    tipo_inimigo(6).atira = True
    tipo_inimigo(6).resistencia = 7
    tipo_inimigo(6).pontos = 2500

    tipo_inimigo(7).atira = True
    tipo_inimigo(7).resistencia = 8
    tipo_inimigo(7).pontos = 2750


End Sub



Public Sub abertura_evla()
    
    Dim DDSEVLA_logo As DirectDrawSurface7
    Dim dsEvla       As DirectSoundBuffer
    Dim EVLARect     As RECT
    Dim x            As Integer
    Dim y            As Single
    
    
    With DDSDESC
      .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
      .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
      .lWidth = 218
      .lHeight = 105
    End With
    
    Set DDSEVLA_logo = DDraw.CreateSurfaceFromFile(App.Path & "\graficos\evla_logo.bmp", DDSDESC)
    
    Set dsEvla = DS.CreateSoundBufferFromFile(App.Path & "\sons\evla_intro.wav", bdesc, DSwaveFormat)
    
    y = -150
    x = 200
    
    dsEvla.SetCurrentPosition 0
    dsEvla.Play DSBPLAY_DEFAULT
    
    For cont = 0 To 170
            
        DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
        
        With EVLARect
            .Top = 0
            .Bottom = .Top + 105
            .Left = 0
            .Right = .Left + 218
        End With
        
        If x + 16 > 640 Then
          TempX = Abs(x - 640)
          EVLARect.Right = TempX
        End If

        If y + EVLARect.Bottom > 480 Then
          TempY = Abs(y - 480)
          EVLARect.Bottom = TempY
        End If

        If x < 0 Then
          TempX = Abs(x)
          EVLARect.Left = TempX
        End If

        If y < 0 Then
          TempY = Abs(y)
          EVLARect.Top = TempY
        End If
    
        If x >= 0 And y >= 0 Then
          DDSBack.BltFast x, y, DDSEVLA_logo, EVLARect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        Else
          If x < 0 Then DDSBack.BltFast 0, y, DDSEVLA_logo, EVLARect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
          If y < 0 Then DDSBack.BltFast x, 0, DDSEVLA_logo, EVLARect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
          If x < 0 And y < 0 Then DDSBack.BltFast 0, 0, DDSEVLA_logo, EVLARect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        End If
            
        DDSPrimary.Flip Nothing, 0
        
        y = y + 2
    
    Next cont
    
    TocaMidi "abertura.mid" 'toca musica continuamente
    
    lngTime = DX.TickCount
    Do Until DX.TickCount > lngTime + 1000 'espera 3 segundos
    Loop
    
    Set DDSEVLA_logo = Nothing
    Set dsEvla = Nothing

End Sub
Public Sub inclui_recorde()
    
    Dim cont               As Integer
    Dim array_recordes(9)  As String
    Dim strtemp            As String

    tmpString = ""
    jogo.status = 7
    While jogo.status = 7
        DoEvents
        DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
        DDSBack.SetForeColor RGB(255, 255, 250)
        Form1.FontBold = False
        Form1.FontSize = 16
        DDSBack.SetFont Form1.Font
        DDSBack.DrawText 100, 100, "NOVO RECORDE !!!! ", False
        DDSBack.DrawText 100, 150, "DIGITE SEU NOME : " & tmpString & "_", False
        
        DDSPrimary.Flip Nothing, 0
    Wend
    
    cont = 0
    'file 1 lê o arquivo
    file1 = FreeFile
    Open App.Path & "\recordes.txt" For Input As #file1
    Do While Not EOF(file1) And cont < 10
        Input #file1, strtemp
        array_recordes(cont) = strtemp
        cont = cont + 1
    Loop
    
    Close #file1
    
    'file 2 grava
    file2 = FreeFile
    Open App.Path & "\recordes.txt" For Output As #file2
    'grava o recorde atual primeiro
    Print #file2, tmpString & ";" & Str(jogo.placar)
    cont = 0
    Do While cont < 10 And Trim(array_recordes(cont)) <> ""
        Print #file2, array_recordes(cont)
        cont = cont + 1
    Loop
    Close #file2
    
End Sub
Public Sub recordes()
    Dim cont    As Integer
    Dim strtemp As String
    Dim str_aux As String

    DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
    DDSBack.DrawLine 1, 35, 640, 35
    DDSBack.SetForeColor RGB(255, 255, 250)
    Form1.FontBold = True
    DDSBack.SetFont Form1.Font
    DDSBack.DrawText 10, 10, "OS 10 MAIORES PLACARES ", False

    
    cont = 0
    file1 = FreeFile
    Open App.Path & "\recordes.txt" For Input As #file1
    'Erase editor_fase
    Do While Not EOF(file1) And cont < 10
        Input #file1, strtemp
        
        str_aux = Mid(strtemp, 1, InStr(1, strtemp, ";") - 1)
        DDSBack.DrawText 100, 40 * (cont + 1), str_aux, False
        
        str_aux = Mid(strtemp, InStr(1, strtemp, ";") + 1, Len(strtemp))
        DDSBack.DrawText 450, 40 * (cont + 1), str_aux, False
        
        cont = cont + 1
    Loop
    
    Close #file1
    
    Form1.FontSize = 10
    Form1.FontBold = False
    DDSBack.SetFont Form1.Font
    DDSBack.SetForeColor RGB(255, 255, 250)
    DDSBack.DrawText 10, 450, "ESC PARA VOLTAR AO MENU  ", False


End Sub
Public Sub instrucoes()

    Dim extrasRect As RECT
    Dim a          As Integer
    
    DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
    DDSBack.SetForeColor RGB(255, 255, 250)
    DDSBack.DrawText 10, 10, "INSTRUÇÕES  ", False
    Form1.FontSize = 10
    Form1.FontBold = True
    DDSBack.SetFont Form1.Font
    DDSBack.DrawText 10, 40, "CONTROLES : ", False
    DDSBack.DrawText 10, 60, "TECLAS CURSORAS PARA MOVER A NAVE E A BARRA DE ESPAÇO PARA ATIRAR ", False
    DDSBack.DrawText 10, 75, "TECLE 'ESC' PARA ENCERRAR O JOGO , E A TECLA 'P' PARA PAUSAR.", False
    DDSBack.DrawText 10, 120, "EXTRAS : ", False
    DDSBack.DrawText 10, 140, "SÃO OBTIDOS AO SE DESTRUIR DETERMINADOS INIMIGOS. VALEM 50 PONTOS CADA ", False
    For a = 0 To 7
        With extrasRect
            .Top = 0
            .Bottom = 32
            .Left = 16 * a
            .Right = .Left + 16
        End With
        DDSBack.BltFast 10 + a * 80, 170, DDSExtras, extrasRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        If a < 5 Then
            Form1.FontSize = 6
            Form1.FontBold = False
            DDSBack.SetFont Form1.Font
            DDSBack.DrawText 10 + a * 80, 205, "ARMA " & (a + 1), False
            DDSBack.DrawText 10 + a * 80, 215, "PODER : " & tipo_arma(a + 1).poder, False
        ElseIf a = 5 Then
            DDSBack.DrawText 10 + a * 80, 205, "15s DE", False
            DDSBack.DrawText 10 + a * 80, 215, "INVENCIBILIDADE", False
        ElseIf a = 6 Then
            DDSBack.DrawText 10 + a * 80, 205, "TIRO EXTRA", False
        ElseIf a = 7 Then
            DDSBack.DrawText 10 + a * 80, 205, "VELOCIDADE", False
            DDSBack.DrawText 10 + a * 80, 215, "EXTRA", False
        End If
    Next a
    Form1.FontSize = 10
    Form1.FontBold = True
    DDSBack.SetFont Form1.Font
    DDSBack.DrawText 10, 280, "INIMIGOS : ", False
    For a = 0 To 7
        With extrasRect
            .Top = 0
            .Bottom = .Top + 46
            .Left = a * 35
            .Right = .Left + 35
        End With
        DDSBack.BltFast 10 + a * 80, 300, DDSInimigos, extrasRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Form1.FontSize = 6
        Form1.FontBold = False
        DDSBack.SetFont Form1.Font
        DDSBack.DrawText 10 + a * 80, 345, "PODER : " & tipo_inimigo(a).resistencia, False
        DDSBack.DrawText 10 + a * 80, 355, "PONTOS:" & tipo_inimigo(a).pontos, False
        
    Next a
    
    
    Form1.FontSize = 6
    Form1.FontBold = False
    DDSBack.SetFont Form1.Font
    DDSBack.DrawText 10, 450, "ESC PARA VOLTAR AO MENU  ", False
    DDSBack.DrawLine 1, 35, 640, 35
    
    
End Sub
Public Sub opcoes()

        DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
        DDSBack.DrawLine 1, 35, 640, 35
        DDSBack.SetForeColor RGB(255, 255, 250)
        Form1.FontBold = True
        DDSBack.SetFont Form1.Font
        DDSBack.DrawText 10, 10, "OPÇÕES DE JOGO  ", False
        
        If opcao = 1 Then
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 90, 150, "DIFICULDADE : ", False
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 100, 200, "EXIBIR FPS   : ", False
        ElseIf opcao = 2 Then
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 90, 150, "DIFICULDADE : ", False
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 100, 200, "EXIBIR FPS   : ", False
        End If
        If jogo.nivel = 1 Then
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 250, 150, "FÁCIL  ", False
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 320, 150, "NORMAL  ", False
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 430, 150, "DIFÍCIL  ", False
        ElseIf jogo.nivel = 2 Then
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 250, 150, "FÁCIL  ", False
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 320, 150, "NORMAL  ", False
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 430, 150, "DIFÍCIL  ", False
        ElseIf jogo.nivel = 3 Then
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 250, 150, "FÁCIL  ", False
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 320, 150, "NORMAL  ", False
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 430, 150, "DIFÍCIL  ", False
        End If
        
        If jogo.FPS Then
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 250, 200, "SIM ", False
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 300, 200, "NÃO ", False
        Else
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 250, 200, "SIM ", False
            DDSBack.SetForeColor RGB(255, 20, 10)
            DDSBack.DrawText 300, 200, "NÃO ", False
        End If
        Form1.FontSize = 10
        Form1.FontBold = False
        DDSBack.SetFont Form1.Font
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 10, 450, "ESC PARA VOLTAR AO MENU  ", False


End Sub
Public Sub main()

Dim Rect_origem As RECT
Dim Rect_dest As RECT
Dim x As Integer
Dim texto  As String


x = 650
opcao = 1
jogo.FPS = True
jogo.nivel = 2

inicializaDX
abertura_evla
jogo.status = 1
jogo.item_abertura = 1
define_extras

While jogo.status = 1 'tela de abertura
    DoEvents
    
    TocaMidi "abertura.mid" 'toca musica continuamente
    
    DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
        
        
    Form1.FontBold = False
    DDSBack.SetForeColor RGB(255, 255, 250)
    Form1.FontSize = 10
    Form1.FontName = "MS Sans Serif"
    Form1.FontBold = False
    DDSBack.SetFont Form1.Font
    
    
    DDSBack.DrawText 90, 305, "2001 EVLA SOFTWARE", False
    texto = "TORION É UM PROJETO SEM FINS LUCRATIVOS , PORTANTO É TOTALMENTE FREEWARE. "
    texto = texto & " ALGUNS ARQUIVOS GRÁFICOS, ASSIM COMO A TRILHA SONORA E OS EFEITOS SONOROS SÃO DE AUTORIA DE TERCEIROS.     "
    texto = texto & " VISITEM O SITE DA EVLA SOFTWARE : www.evla.hpg.com.br     "
    texto = texto & "   --- CRÉDITOS --- PROGRAMAÇÃO : EULER V. L. DE ALMEIDA  "
    texto = texto & "    GRÁFICOS : ANTHARES SOFTWARE , ARI FELDMAN ( LICENSA Nº 200.223.31.73-972937899 ) , EULER V. L. DE ALMEIDA , SEGA , TECHO SOFT - JAPAN     "
    texto = texto & "    AGRADECIMENTOS ESPECIAIS :  ANDRÉ LUIZ SILVA , PROGRAMADORES DE JOGOS , PROJETO VB OPEN SOURCE E OS DESENVOLVEDORES DO JOGO SPACE SHOOTER 2K "
    DDSBack.DrawText x, 450, texto, False
    
    Form1.FontSize = 16
    DDSBack.SetFont Form1.Font
    If opcao = 1 Then
        DDSBack.SetForeColor RGB(255, 20, 10)
        DDSBack.DrawText 250, 335, "NOVO JOGO", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 355, "OPÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 375, "RECORDES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 395, "INSTRUÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 415, "FIM", False
    ElseIf opcao = 2 Then
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 335, "NOVO JOGO", False
        DDSBack.SetForeColor RGB(255, 20, 10)
        DDSBack.DrawText 250, 355, "OPÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 375, "RECORDES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 395, "INSTRUÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 415, "FIM", False
    ElseIf opcao = 3 Then
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 335, "NOVO JOGO", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 355, "OPÇÕES", False
        DDSBack.SetForeColor RGB(255, 20, 10)
        DDSBack.DrawText 250, 375, "RECORDES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 395, "INSTRUÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 415, "FIM", False
    ElseIf opcao = 4 Then
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 335, "NOVO JOGO", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 355, "OPÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 375, "RECORDES", False
        DDSBack.SetForeColor RGB(255, 20, 10)
        DDSBack.DrawText 250, 395, "INSTRUÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 415, "FIM", False
    ElseIf opcao = 5 Then
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 335, "NOVO JOGO", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 355, "OPÇÕES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 375, "RECORDES", False
        DDSBack.SetForeColor RGB(255, 255, 250)
        DDSBack.DrawText 250, 395, "INSTRUÇÕES", False
        DDSBack.SetForeColor RGB(255, 20, 10)
        DDSBack.DrawText 250, 415, "FIM", False
    End If
        
    With Rect_origem
        .Top = 0
        .Bottom = 59
        .Left = 0
        .Right = .Left + 187
    End With
        
    With Rect_dest
        .Top = 50
        .Bottom = 300
        .Left = 10
        .Right = .Left + 587
    End With
        
    DDSBack.Blt Rect_dest, DDSTorionLogo, Rect_origem, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
    
    x = x - 2
        
    If x < Len(texto) * -8 Then
        x = 650
    End If
        
    'Menu Opções
    If jogo.item_abertura = 2 Then
        opcoes
    ElseIf jogo.item_abertura = 3 Then 'recordes
        recordes
    ElseIf jogo.item_abertura = 4 Then
        instrucoes
    End If
        
    DDSPrimary.Flip Nothing, 0 'Passo do Back Buffer para a Surface principal
    
    While jogo.status <> 1
        
        DoEvents
        
        DDSBack.BltColorFill EmptyRect, 0 'Limpa o back buffer
        
        atualiza_fase
        
        If jogo.status = 2 Then 'jogo normal
            TocaMidi "fase" & jogo.fase_atual & ".mid" 'toca musica continuamente
            ler_teclado
        ElseIf jogo.status = 3 Then 'tela de congratulações
            TocaMidi "congrats.mid" 'musica da vitoria
            Form1.FontSize = 24
            DDSBack.SetFont Form1.Font
            DDSBack.DrawText 130, 200, "CONGRATULAÇÕES PILOTO !!!!", False
            DDSBack.DrawText 50, 250, "VOCÊ VENCEU TODAS AS FASES DE TORION", False
            DDSBack.DrawText 130, 300, "BÔNUS 50000 x " & torion.vidas & " = " & torion.vidas * 50000, False
        ElseIf jogo.status = 4 Then 'tela de fim de fase
            If editor_fase(jogo.indice_editor_fase - 1).inicia + 8 < jogo.contador Then
                para_midi
                jogo.status = 2 'volta para a jogo
                jogo.indice_editor_fase = 0
                jogo.contador = 0
                jogo.fase_atual = jogo.fase_atual + 1
                inicializa_fase
            Else
                If (seg Is Nothing) And (segstate Is Nothing) Then 'Se o segmento da musica está vazio . . .
                    Set seg = loader.LoadSegment("fim_fase.mid") 'Carrega MIDI passado pela função
                    Set segstate = perf.PlaySegment(seg, 0, 0) 'Toca MIDI carregado no Segmento
                End If
                Form1.FontSize = 26
                DDSBack.SetFont Form1.Font
                DDSBack.DrawText 190, 250, "FASE " & jogo.fase_atual & " COMPLETA", False
            End If
        ElseIf jogo.status = 5 Then 'tela de fim de jogo
            TocaMidi "fim.mid"
            Form1.FontSize = 26
            DDSBack.SetFont Form1.Font
            DDSBack.DrawText 220, 250, "F I M  D O  J O G O", False
        End If
        
        If torion.arma_atual = 1 Then
            atualiza_armas DDSArma1
        ElseIf torion.arma_atual = 2 Then
            atualiza_armas DDSArma2
        ElseIf torion.arma_atual = 3 Then
            atualiza_armas DDSArma3
        ElseIf torion.arma_atual = 4 Then
            atualiza_arma4
        ElseIf torion.arma_atual = 5 Then
            atualiza_armas DDSMissil
        End If
        
        atualiza_tiro 'Esse é o tiro normal da nave que pode ter até 3 simultâneos
        
        atualiza_nave_torion
        
        atualiza_inimigos
        
        atualiza_tiro_inimigo
        
        atualiza_extras
        
        If jogo.status = 2 Then 'jogo normal
            testa_colisoes
        End If
        
        atualiza_explosao
        
        atualiza_nuvem
    
        atualiza_painel
        
        If jogo.status = 6 Then 'PAUSE
            Form1.FontSize = 26
            DDSBack.SetFont Form1.Font
            DDSBack.DrawText 250, 250, "P A U S A", False
            DDSPrimary.Flip Nothing, 0 'Passo do Back Buffer para a Surface principal
            While jogo.status = 6
                DoEvents
            Wend
        End If
        
        If jogo.FPS Then
            Form1.FontSize = 10
            Form1.FontBold = False
            DDSBack.SetFont Form1.Font
            DDSBack.SetForeColor RGB(255, 255, 250)
            DDSBack.DrawText 10, 460, "FPS: " & FPS, False
        End If
        
        DDSPrimary.Flip Nothing, 0 'Passo do Back Buffer para a Surface principal
    
    Wend
    
Wend


End Sub
Public Sub atualiza_tiro()
    
    Dim intCount  As Integer
    Dim EmptyRect As RECT   'retangulo
    
    'tiro da nossa nave
    intCount = 0
    Do Until intCount > UBound(tiro)  'Loop através do vetor do raio laser
        If tiro(intCount).Existe Then  'Se o laser existir . . .
            tiro(intCount).y = tiro(intCount).y - 10 'Incrementa a posicao Y
            If tiro(intCount).y < 0 Then 'Se o laser sair da tela
                tiro(intCount).Existe = False 'O laser não existe mais
                tiro(intCount).y = 0 'Limpa posição Y
                tiro(intCount).x = 0 'Limpa posição X
            Else
                 DDSBack.BltFast tiro(intCount).x, tiro(intCount).y, DDSTiro, EmptyRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
        intCount = intCount + 1
    Loop


End Sub
Public Sub atualiza_arma4()

    Dim ArmaRect            As RECT
    Dim TempX               As Integer
    Dim TempY               As Integer
    Dim intCount            As Single
    
        
    'Preenche os valores do retangulo que conterá o frame corrente
    With ArmaRect
        .Top = 0
        .Bottom = .Top + tipo_arma(torion.arma_atual).altura
        .Left = tipo_arma(torion.arma_atual).largura * arma(intCount).frame_atual
        .Right = .Left + tipo_arma(torion.arma_atual).largura
    End With
    
    For intCount = 0 To UBound(arma)
        If arma(intCount).angulo <= 359 Then
            arma(intCount).angulo = arma(intCount).angulo + 0.3
        Else
            arma(intCount).angulo = 0
        End If
        
        'calcula a nova posição
        arma(intCount).x = (Cos(arma(intCount).angulo) * 80) + torion.x
        arma(intCount).y = (Sin(arma(intCount).angulo) * 80) + torion.y
        
        arma(intCount).Existe = True
        
        If arma(intCount).x + ArmaRect.Right > 640 Then
          TempX = Abs(arma(intCount).x - 640)
          ArmaRect.Right = TempX
        End If
    
        If arma(intCount).y + ArmaRect.Bottom > 480 Then
          TempY = Abs(arma(intCount).y - 480)
          ArmaRect.Bottom = TempY
        End If
    
        If arma(intCount).x < 0 Then
          TempX = Abs(arma(intCount).x)
          ArmaRect.Left = TempX
        End If
    
        If arma(intCount).y < 0 Then
          TempY = Abs(arma(intCount).y)
          ArmaRect.Top = TempY
        End If
    
        If arma(intCount).x >= 0 And arma(intCount).y >= 0 Then
          DDSBack.BltFast arma(intCount).x, arma(intCount).y, DDSArma3, ArmaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Else
          If arma(intCount).x < 0 Then DDSBack.BltFast 0, arma(intCount).y, DDSArma3, ArmaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If arma(intCount).y < 0 Then DDSBack.BltFast arma(intCount).x, 0, DDSArma3, ArmaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
          If arma(intCount).x < 0 And arma(intCount).y < 0 Then DDSBack.BltFast 0, 0, DDSArma3, ArmaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    Next

End Sub
Public Sub atualiza_tiro_inimigo()
    Dim cont                 As Single
    Dim rect_tiro_inimigo    As RECT

    With rect_tiro_inimigo
        .Top = 0
        .Bottom = 8
        .Left = 0
        .Right = 8
    End With

    For cont = 0 To UBound(tiro_inimigo)
        If tiro_inimigo(cont).Existe Then
            If tiro_inimigo(cont).y < 0 Then
                tiro_inimigo(cont).Existe = False
            ElseIf tiro_inimigo(cont).y > 480 Then
                tiro_inimigo(cont).Existe = False
            ElseIf tiro_inimigo(cont).x > 640 Then
                tiro_inimigo(cont).Existe = False
            ElseIf tiro_inimigo(cont).x < 0 Then
                tiro_inimigo(cont).Existe = False
            Else
                tiro_inimigo(cont).x = tiro_inimigo(cont).x - 5 * Sin(tiro_inimigo(cont).angulo / 57.2957795130824)
                tiro_inimigo(cont).y = tiro_inimigo(cont).y + 5 * Cos(tiro_inimigo(cont).angulo / 57.2957795130824)
                DDSBack.BltFast tiro_inimigo(cont).x, tiro_inimigo(cont).y, DDSTiroInimigo, rect_tiro_inimigo, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next cont

End Sub
Public Sub atualiza_armas(surface As DirectDrawSurface7)
    '
    'Exibe as armas na tela
    '
    Dim ArmaRect  As RECT
    Dim intCount  As Integer
    Dim direcao_missel As Single
    Static ArmaCounter As Byte
    Dim cont_inimigos      As Integer
    
    intCount = 0
    Do Until intCount > UBound(arma)  'Loop através do vetor da arma
        If arma(intCount).Existe Then
            If arma(intCount).y < 0 Then
                arma(intCount).Existe = False
            ElseIf arma(intCount).y > 480 Then
                arma(intCount).Existe = False
            ElseIf arma(intCount).x > 640 Then
                arma(intCount).Existe = False
            ElseIf arma(intCount).x < 0 Then
                arma(intCount).Existe = False
            Else
                If torion.arma_atual = 5 Then 'Missel teleguiado
                    'IA do míssel
                    If arma(intCount).alvo = 999 Then 'se nao tiver nenhum alvo no momento
                        'Seleciona um alvo
                        For cont_inimigos = 0 To UBound(inimigo)
                            If inimigo(cont_inimigos).Existe Then
                                arma(intCount).alvo = cont_inimigos
                                Exit For
                            Else
                                arma(intCount).alvo = 999
                                'se não encontrar nenhum alvo segue em frente
                                arma(intCount).movimento_y = -10
                                arma(intCount).movimento_x = 0
                                direcao_missel = 1
                            End If
                        Next cont_inimigos
                    Else
                        If inimigo(arma(intCount).alvo).Existe Then
                            If inimigo(arma(intCount).alvo).x > arma(intCount).x Then
                                If inimigo(arma(intCount).alvo).y > arma(intCount).y Then
                                    arma(intCount).movimento_y = 10
                                    arma(intCount).movimento_x = 10
                                    direcao_missel = 5
                                ElseIf inimigo(arma(intCount).alvo).y < arma(intCount).y Then
                                    arma(intCount).movimento_y = -10
                                    arma(intCount).movimento_x = 10
                                    direcao_missel = 0
                                ElseIf inimigo(arma(intCount).alvo).y = arma(intCount).y Then
                                    arma(intCount).movimento_y = 0
                                    arma(intCount).movimento_x = 10
                                    direcao_missel = 6
                                End If
                            ElseIf inimigo(arma(intCount).alvo).x < arma(intCount).x Then
                                If inimigo(arma(intCount).alvo).y > arma(intCount).y Then
                                    arma(intCount).movimento_y = 10
                                    arma(intCount).movimento_x = -10
                                    direcao_missel = 3
                                ElseIf inimigo(arma(intCount).alvo).y < arma(intCount).y Then
                                    arma(intCount).movimento_y = -10
                                    arma(intCount).movimento_x = -10
                                    direcao_missel = 4
                                ElseIf inimigo(arma(intCount).alvo).y = arma(intCount).y Then
                                    arma(intCount).movimento_y = 0
                                    arma(intCount).movimento_x = -10
                                    direcao_missel = 7
                                End If
                            ElseIf inimigo(arma(intCount).alvo).x = arma(intCount).x Then
                                If inimigo(arma(intCount).alvo).y > arma(intCount).y Then
                                    arma(intCount).movimento_y = 10
                                    arma(intCount).movimento_x = 0
                                    direcao_missel = 2
                                ElseIf inimigo(arma(intCount).alvo).y < arma(intCount).y Then
                                    arma(intCount).movimento_y = -10
                                    arma(intCount).movimento_x = 0
                                    direcao_missel = 1
                                End If
                            End If
                        Else
                            arma(intCount).alvo = 999
                            'se não encontrar nenhum alvo segue em frente
                            arma(intCount).movimento_y = -10
                            arma(intCount).movimento_x = 0
                            direcao_missel = 1
                        End If
                    End If
                    
                    With ArmaRect
                        .Top = tipo_arma(torion.arma_atual).altura * arma(intCount).frame_atual
                        .Bottom = .Top + tipo_arma(torion.arma_atual).altura
                        .Left = tipo_arma(torion.arma_atual).largura * direcao_missel
                        .Right = .Left + tipo_arma(torion.arma_atual).largura
                    End With
                Else
                    With ArmaRect
                        .Top = 0
                        .Bottom = .Top + tipo_arma(torion.arma_atual).altura
                        .Left = tipo_arma(torion.arma_atual).largura * arma(intCount).frame_atual
                        .Right = .Left + tipo_arma(torion.arma_atual).largura
                    End With
                End If
                
                'Define para qual direcao o tiro sai
                arma(intCount).x = arma(intCount).x + arma(intCount).movimento_x
                arma(intCount).y = arma(intCount).y + arma(intCount).movimento_y
                                               
                'Testa se for o ultimo frame volta para o 1º
                If arma(intCount).frame_atual = tipo_arma(torion.arma_atual).numero_frames Then
                    arma(intCount).frame_atual = 0
                Else
                    ArmaCounter = ArmaCounter + 1
                    If ArmaCounter = 10 Then
                        arma(intCount).frame_atual = arma(intCount).frame_atual + 1           'incrementa o frame
                        ArmaCounter = 0
                    End If
                End If
                DDSBack.BltFast arma(intCount).x, arma(intCount).y, surface, ArmaRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
        intCount = intCount + 1
    Loop



End Sub
Public Sub atualiza_explosao()
    '
    'Mostra as explosões na tela
    '
    Dim intCount            As Integer
    Dim Retangulo_explosao  As RECT

    
    For intCount = 0 To UBound(explosao)  'Loop através do vetor
        If explosao(intCount).Existe Then  'Se a explosao existir . . .
            If explosao(intCount).frame_atual <= 7 Then    'Se não chegou no ultimo frame
                With Retangulo_explosao 'Retangulo da explosao
                    .Top = 0
                    .Bottom = .Top + 30
                    .Left = 34 * explosao(intCount).frame_atual
                    .Right = .Left + 34
                End With
                DDSBack.BltFast explosao(intCount).x, explosao(intCount).y, DDSExplosao0, Retangulo_explosao, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                'Incrementa o frame para dar a animação à explosao
                 explosao(intCount).frame_atual = explosao(intCount).frame_atual + 1           'incrementa o frame
                
            Else 'Se a explosao acabou
                explosao(intCount).Existe = False
                explosao(intCount).frame_atual = 0
            End If
        End If
    Next
    
End Sub

Public Sub preenche_vetor_explosao(x As Single, y As Single)
    '
    'Essa fucao tem como objetivo percorrer o vetor de explosões
    'até encontrar um posicao no array que não exista nenhuma explosao
    'acontecendo. Depois, atualiza os valores para depois ser exibido na tela
    'pela função atualiza_explosao()
    '
    
    Dim intCount As Integer
    

    For intCount = 0 To UBound(explosao) 'Loop até achar uma posição do vetor vazia
        If explosao(intCount).Existe = False Then   'Se essa posicao do array tiver vazia . . .
            With explosao(intCount) 'cria uma nova explosao na tela
                .Existe = True  ' a explosao está ativa
                .x = x         'Posição X da explosao
                .y = y         'Posição Y da explosao
                .frame_atual = 0      '1º frame da explosao
            End With
            dsExplosao.SetCurrentPosition 0 'seta a posição do buffer para 0
            dsExplosao.Play DSBPLAY_DEFAULT 'Som da explosao
            Exit For 'sai do loop
        End If
    Next

End Sub


Public Sub atira()
    '
    'Essa função é executada a partir do evento disparado pelo usuário(Teclando espaço)
    'Sua finalidade é atualizar os valores do array de forma que possa ser exibida
    'depois pela função atualiza_armas()
    '
    Dim intCount As Single
    Dim cont     As Single
    Static TiroCounter As Byte
    Static ArmaCounter As Byte
    Dim KeyboardState(0 To 255) As Byte
    Static balanco As Long
    Dim cont_inimigos   As Integer
        
    DITeclado.Acquire
    DITeclado.GetDeviceState 256, KeyboardState(0)
    
    'Tiro da nave do jogador
    TiroCounter = TiroCounter + 1
    If TiroCounter >= 7 Then 'Marreta para o tiro sair com um intervalo de um para o outro
        For cont = 1 To torion.tiros_laser
            intCount = 0
            Do Until intCount > UBound(tiro)  'Loop até achar uma posição do vetor vazia
                If tiro(intCount).Existe = False Then  'Se essa posicao do array tiver vazia . . .
                    With tiro(intCount) 'cria um novo laser ativo na tela
                        .Existe = True  ' o laser existe
                        .poder = 1
                        'Se a nave só tiver um laser, então o tiro sai do meio da nave
                        If torion.tiros_laser = 1 Then
                            .x = torion.x + ((LARGURA_TORION \ 2)) - 1
                        End If
                        'Se tiver 2 laser, os tiros saem um do lado do outro
                        If torion.tiros_laser = 2 Then
                            If cont = 1 Then
                                .x = torion.x + ((LARGURA_TORION \ 2)) - 14
                            Else
                                .x = torion.x + ((LARGURA_TORION \ 2)) + 10
                            End If
                        End If
                        If torion.tiros_laser = 3 Then
                            If cont = 1 Then
                                .x = torion.x + ((LARGURA_TORION \ 2)) - 14
                            ElseIf cont = 2 Then
                                .x = torion.x + ((LARGURA_TORION \ 2)) - 2
                            Else
                                .x = torion.x + ((LARGURA_TORION \ 2)) + 10
                            End If
                        End If
                        .y = torion.y   'A mesma posicao do Y da nave
                    End With
                    dsTiro.SetCurrentPosition 0
                    dsTiro.SetPan 0
                    dsTiro.Play DSBPLAY_DEFAULT
                    Exit Do 'sai do loop
                End If
                intCount = intCount + 1
            Loop
        Next cont
        TiroCounter = 0
    End If
    
    
    ArmaCounter = ArmaCounter + 1
    If ArmaCounter >= 10 Then 'Marreta para o tiro sair com um intervalo de um para o outro
        intCount = 0
        Do Until intCount > UBound(arma)  'Loop até achar uma posição do vetor vazia
            If arma(intCount).Existe = False Then  'Se essa posicao do array tiver vazia . . .
                With arma(intCount) 'cria um novo laser ativo na tela
                    .Existe = True  ' o laser existe
                    .x = torion.x + ((LARGURA_TORION \ 2)) - 6
                    .y = torion.y - 10  'A mesma posicao do Y da nave
                End With
    
                If balanco = -1000 Then
                    balanco = 1000
                Else
                    balanco = -1000
                End If
                If torion.arma_atual = 1 Then
                    dsArma1.SetCurrentPosition 0
                    dsArma1.SetPan balanco
                    dsArma1.Play DSBPLAY_DEFAULT
                
                    arma(intCount).movimento_x = 0
                    arma(intCount).movimento_y = -10
                    Exit Do
                ElseIf torion.arma_atual = 2 Then
                    arma(intCount).movimento_x = 0
                    If (KeyboardState(DIK_LEFT)) <> 0 Then
                        arma(intCount).movimento_x = -10
                        dsArma2.SetPan 1000
                    End If
                    If (KeyboardState(DIK_RIGHT)) <> 0 Then
                        arma(intCount).movimento_x = 10
                        dsArma2.SetPan -1000
                    End If
                    arma(intCount).movimento_y = -10
                    dsArma2.SetCurrentPosition 0
                    dsArma2.SetPan 0
                    dsArma2.Play DSBPLAY_DEFAULT
                    
                    Exit Do 'sai do loop
                ElseIf torion.arma_atual = 3 Then
                    dsArma3.SetPan 0
                    If (KeyboardState(DIK_UP)) <> 0 Then
                        arma(intCount).movimento_x = 0
                        arma(intCount).movimento_y = -10
                    ElseIf (KeyboardState(DIK_DOWN)) <> 0 Then
                        arma(intCount).movimento_x = 0
                        arma(intCount).movimento_y = 10
                    ElseIf (KeyboardState(DIK_LEFT)) <> 0 Then
                        arma(intCount).movimento_x = -10
                        arma(intCount).movimento_y = 0
                        dsArma3.SetPan 1000
                    ElseIf (KeyboardState(DIK_RIGHT)) <> 0 Then
                        arma(intCount).movimento_x = 10
                        arma(intCount).movimento_y = 0
                        dsArma3.SetPan -1000
                    Else
                        arma(intCount).movimento_x = 0
                        arma(intCount).movimento_y = -10
                    End If
                    dsArma3.SetCurrentPosition 0
                    dsArma3.Play DSBPLAY_DEFAULT
                    Exit Do 'sai do loop
                ElseIf torion.arma_atual = 4 Then
                    arma(intCount).movimento_x = 0
                    arma(intCount).movimento_y = -10
                    Exit Do
                
                ElseIf torion.arma_atual = 5 Then
                    dsMissel.SetCurrentPosition 0
                    dsMissel.SetPan balanco
                    dsMissel.Play DSBPLAY_DEFAULT
                    arma(intCount).alvo = 999
                    Exit Do
                End If
            End If
            intCount = intCount + 1
        Loop
        ArmaCounter = 0
    End If
    
End Sub
Public Function GetAngle(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double

    Dim XComp As Single
    Dim YComp As Single
    
    'Find the angle between the 2 coords
    XComp = X2 - X1
    YComp = Y1 - Y2
    If Sgn(YComp) > 0 Then GetAngle = Atn(XComp / YComp)
    If Sgn(YComp) < 0 Then GetAngle = Atn(XComp / YComp) + PI
    
    GetAngle = GetAngle * 57.2957795130823
    
End Function

Public Sub atualiza_inimigos()
    Dim TempX               As Integer
    Dim TempY               As Integer
    Dim Retangulo_inimigo   As RECT
    Dim count               As Single
    Dim count_tiro          As Single
    
    For count = 0 To UBound(inimigo)
        If inimigo(count).Existe Then
            If inimigo(count).formacao = 1 Then
                formacao1 (count)
            ElseIf inimigo(count).formacao = 2 Then
                formacao2 (count)
            ElseIf inimigo(count).formacao = 3 Then
                formacao3 (count)
            ElseIf inimigo(count).formacao = 4 Then
                formacao4 (count)
            ElseIf inimigo(count).formacao = 5 Then
                formacao5 (count)
            End If
            'Se o inimigo pode atirar, então manda bala em nóis
            If tipo_inimigo(inimigo(count).tipo).atira Then
                If inimigo(count).y <= torion.y Then
                    For count_tiro = 0 To UBound(tiro_inimigo) - (jogo.numero_fases - jogo.fase_atual)
                        If Not tiro_inimigo(count_tiro).Existe Then
                            With tiro_inimigo(count_tiro)
                                .Existe = True
                                .x = inimigo(count).x
                                .y = inimigo(count).y
                                .angulo = GetAngle(.x, .y, torion.x, torion.y)
                            End With
                            Exit For
                        End If
                    Next count_tiro
                End If
            End If
            
            With Retangulo_inimigo
                .Top = 0
                .Bottom = .Top + 46
                .Left = inimigo(count).tipo * 35
                .Right = .Left + 35
            End With
            
            If inimigo(count).x + 35 > 640 Then
              TempX = Abs(inimigo(count).x - 640)
              Retangulo_inimigo.Right = TempX
            End If
        
            If inimigo(count).y + Retangulo_inimigo.Bottom > 480 Then
              TempY = Abs(inimigo(count).y - 480)
              Retangulo_inimigo.Bottom = TempY
            End If
            
            If inimigo(count).x < 0 Then
              TempX = Abs(inimigo(count).x)
              Retangulo_inimigo.Left = TempX
            End If
            
            If inimigo(count).y < 0 Then
              TempY = Abs(inimigo(count).y)
              Retangulo_inimigo.Top = TempY
            End If
        
            If inimigo(count).x >= 0 And inimigo(count).y >= 0 Then
              DDSBack.BltFast inimigo(count).x, inimigo(count).y, DDSInimigos, Retangulo_inimigo, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            Else
              If inimigo(count).x < 0 Then DDSBack.BltFast 0, inimigo(count).y, DDSInimigos, Retangulo_inimigo, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
              If inimigo(count).y < 0 Then DDSBack.BltFast inimigo(count).x, 0, DDSInimigos, Retangulo_inimigo, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
              If inimigo(count).x < 0 And inimigo(count).y < 0 Then DDSBack.BltFast 0, 0, DDSInimigos, Retangulo_inimigo, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
            
            If inimigo(count).y > 480 Then
                inimigo(count).Existe = False
            End If

        End If
    Next count

End Sub
Public Sub inicia_formacao(numero_componentes As Single, tipo_inimigo As Single, formacao As Single, extra As Integer)
    Dim numero_inimigos As Single
    Dim count           As Single
    Dim aux             As Integer
    Dim extra_gravado   As Boolean
    
    Randomize
    extra_gravado = False

    aux = Rnd * 250 + 150
    For count = 0 To UBound(inimigo)
        If Not inimigo(count).Existe Then
            inimigo(count).Existe = True
            inimigo(count).formacao = formacao
            inimigo(count).tipo = tipo_inimigo
            inimigo(count).danos = 0
            If Not extra_gravado Then
                inimigo(count).extra = extra
                extra_gravado = True
            Else
                inimigo(count).extra = 0
            End If
            If formacao = 1 Then
                inimigo(count).x = Rnd * 500 + 35
                inimigo(count).y = 0
            ElseIf formacao = 2 Then
                If numero_inimigos < 1 Then
                    inimigo(count).x = 640
                Else
                    inimigo(count).x = 640 + numero_inimigos * 70
                End If
                inimigo(count).y = 430
                inimigo(count).angulo = 500
            ElseIf formacao = 3 Then
                inimigo(count).x = 100
                If numero_inimigos < 1 Then
                    inimigo(count).y = 0
                Else
                    inimigo(count).y = numero_inimigos * 50 * -1
                End If
            ElseIf formacao = 4 Then
                inimigo(count).x_inicial = aux
                inimigo(count).x = x_inicial
                If numero_inimigos < 1 Then
                    inimigo(count).y = 0
                Else
                    inimigo(count).y = numero_inimigos * 50 * -1
                End If
            ElseIf formacao = 5 Then
                inimigo(count).x = 100
                If numero_inimigos < 1 Then
                    inimigo(count).y = 0
                Else
                    inimigo(count).y = numero_inimigos * 50 * -1
                End If
            End If
            numero_inimigos = numero_inimigos + 1
        End If
        If numero_inimigos = numero_componentes Then Exit For
    Next count


End Sub
Public Sub formacao2(count As Single) 'Inimigos fazendo uma rodinha

    Dim distancia   As Single
    
    distancia = 200

    If inimigo(count).x >= 250 And inimigo(count).angulo = 500 Then
        inimigo(count).x = inimigo(count).x - 2 - jogo.nivel - jogo.fase_atual
    Else
        If inimigo(count).angulo = 500 Then
            inimigo(count).angulo = 90
        End If
        'calcula a nova posição
        inimigo(count).x = (Cos(inimigo(count).angulo) * distancia) + 300
        inimigo(count).y = (Sin(inimigo(count).angulo) * distancia) + 220
        If inimigo(count).angulo >= 360 Then
            inimigo(count).angulo = 1
        Else
            inimigo(count).angulo = inimigo(count).angulo + 0.05
        End If
    End If
            


End Sub
Public Sub formacao5(count As Single) 'inimigo descendo na diagonal

    inimigo(count).y = inimigo(count).y + jogo.nivel * 2 + jogo.fase_atual
    inimigo(count).x = inimigo(count).x + jogo.nivel * 2 + jogo.fase_atual


End Sub
Public Sub formacao4(count As Single) 'inimigos descendo em espiral

    inimigo(count).y = inimigo(count).y + jogo.nivel + jogo.fase_atual
    inimigo(count).x = inimigo(count).x_inicial + Sin(inimigo(count).y / 7) * 50 - 50

End Sub
Public Sub formacao1(count As Single) 'inimigos decendo em linha reta
        
    inimigo(count).y = inimigo(count).y + jogo.nivel * 2 + jogo.fase_atual

End Sub
Public Sub formacao3(count As Single) 'inimigos decendo em zigue-zague

        If inimigo(count).y < 100 Then
            inimigo(count).y = inimigo(count).y + jogo.nivel * 2 + jogo.fase_atual
        End If
        If inimigo(count).y >= 100 And inimigo(count).y <= 200 Then
            inimigo(count).x = inimigo(count).x + jogo.nivel * 2 + jogo.fase_atual
        End If
        If inimigo(count).x >= 200 And inimigo(count).x <= 400 Then
            inimigo(count).y = inimigo(count).y + jogo.nivel * 2 + jogo.fase_atual
        End If
        If inimigo(count).y > 200 Then
            inimigo(count).x = inimigo(count).x + jogo.nivel * 2 + jogo.fase_atual
        End If
        If inimigo(count).x > 400 Then
            inimigo(count).y = inimigo(count).y + jogo.nivel * 2 + jogo.fase_atual
        End If

End Sub

Public Sub atualiza_fase()
    
    Dim cont            As Integer
    Dim existe_inimigo  As Boolean
    
    If jogo.fase_atual = 1 Then
        atualiza_tile DDSMar_tile
    ElseIf jogo.fase_atual = 2 Then
        atualiza_tile DDSLava_tile
    ElseIf jogo.fase_atual = 3 Then
        atualiza_tile DDSSeafloor_tile
    ElseIf jogo.fase_atual = 4 Then
        atualiza_tile DDSGround_tile
    ElseIf jogo.fase_atual = 5 Then
        atualiza_tile DDSGrama_tile
    End If
    
    'Cria a formacao de ataque dos inimigos
    If jogo.indice_editor_fase <= UBound(editor_fase) Then
        If editor_fase(jogo.indice_editor_fase).inicia = jogo.contador Then
                inicia_formacao editor_fase(jogo.indice_editor_fase).num_inimigos, editor_fase(jogo.indice_editor_fase).tipo_inimigo, editor_fase(jogo.indice_editor_fase).formacao, editor_fase(jogo.indice_editor_fase).extra
                If UBound(editor_fase) >= jogo.indice_editor_fase Then
                    jogo.indice_editor_fase = jogo.indice_editor_fase + 1
                End If
        End If
    Else ' fim da fase atual
        If jogo.status <> 5 Then 'se não é fim do jogo
            'testa para ver se ainda tem algum inimigo na tela
            existe_inimigo = False
            For cont = 0 To UBound(inimigo)
                If inimigo(cont).Existe Then
                    existe_inimigo = True
                    Exit For
                End If
            Next cont
            If Not existe_inimigo Then
                If jogo.status = 2 Then
                    editor_fase(jogo.indice_editor_fase - 1).inicia = jogo.contador
                    para_midi
                    Set seg = Nothing
                    Set segstate = Nothing
                    'tela de fim de fase
                    jogo.status = 4
                End If
                If jogo.fase_atual + 1 > jogo.numero_fases And jogo.status = 4 Then
                    'fim do jogo
                    jogo.status = 3
                    jogo.placar = jogo.placar + 50000 * torion.vidas
                End If
            End If
        End If
    End If

End Sub
Public Sub mostra_mensagem(texto As String, x As Integer, y As Integer)
    '
    'Mostra na tela os textos com efeitos especiais
    '
    Dim cont    As Integer
    Dim cor     As Long
    Dim aux     As Long
    Dim recte   As RECT

    
    Form1.FontSize = 48
    DDSBack.SetFont Form1.Font
    
    'A idéia é ir incrementando a paleta de cores até a cor desejada, e depois
    'decrementá-la até desaparecer
    For cont = 1 To 40
        DDSBack.BltColorFill recte, 0 'Limpa o BackBuffer
        cor = DX.CreateColorRGB(cont / 40, cont / 40, cont / 40) 'Retorna cor
        DDSBack.SetForeColor cor 'Seta a cor para o backbuffer
        DDSBack.SetFont Form1.Font 'Seta o tipo de alfabeto para o backbuffer
        DDSBack.DrawText x, y, texto, False 'Escreve no backbuffer
        aux = DX.TickCount 'Serve para dar aquela paradinha, pois é muito rápido
        Do
        Loop Until aux < DX.TickCount - 100
        DDSPrimary.Flip Nothing, 0 'Passa do BackBuffer p/ a surface principal
    Next cont
    
    For cont = 40 To 1 Step -1
        DDSBack.BltColorFill recte, 0 'Limpa o BackBuffer
        cor = DX.CreateColorRGB(cont / 40, cont / 40, cont / 40) 'Retorna cor
        DDSBack.SetForeColor cor 'Seta a cor para o backbuffer
        DDSBack.SetFont Form1.Font 'Seta o tipo de alfabeto para o backbuffer
        DDSBack.DrawText x, y, texto, False 'Escreve no backbuffer
        aux = DX.TickCount 'Serve para dar aquela paradinha, pois é muito rápido
        Do
        Loop Until aux < DX.TickCount - 100
        DDSPrimary.Flip Nothing, 0 'Passa do BackBuffer p/ a surface principal
    Next cont


End Sub

Public Sub destroi_fase()
    
    If jogo.fase_atual = 1 Then
        Set DDSMar_tile = Nothing
    ElseIf jogo.fase_atual = 2 Then
        Set DDSLava_tile = Nothing
    ElseIf jogo.fase_atual = 3 Then
        Set DDSSeafloor_tile = Nothing
    End If

End Sub
Public Sub zera_objetos()

    Dim cont        As Integer
    
    extras.Existe = False
    
    For cont = 0 To UBound(tiro_inimigo)
        tiro_inimigo(cont).Existe = False
    Next cont
    
    For cont = 0 To UBound(nuvem)
        nuvem(cont).Existe = False
    Next cont

    For cont = 0 To UBound(tiro)
        tiro(cont).Existe = False
    Next cont
    
    For cont = 0 To UBound(arma)
        arma(cont).Existe = False
    Next cont
    
    For cont = 0 To UBound(explosao)
        explosao(cont).Existe = False
    Next cont
    
    For cont = 0 To UBound(inimigo)
        inimigo(cont).Existe = False
    Next cont
    
    arma(0).angulo = 0
    arma(1).angulo = 90
    arma(2).angulo = 180


End Sub

Public Sub inicializa_fase()
    Dim tile_cont   As Single
    Dim x           As Single
    Dim y           As Single
    Dim strtemp     As String

    
    'A idéia é so criar as surfaces que serão efetivamente usadas em cada fase para
    'economizar memória de vídeo
    If jogo.fase_atual = 1 Then
        Set DDSMar_tile = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\mar_tile.bmp", True)
    ElseIf jogo.fase_atual = 2 Then
        Set DDSLava_tile = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\lava_tile.bmp", True)
    ElseIf jogo.fase_atual = 3 Then
        Set DDSSeafloor_tile = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\seafloor_tile.gif", True)
    ElseIf jogo.fase_atual = 4 Then
        Set DDSGround_tile = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\ground_tile.bmp", True)
    ElseIf jogo.fase_atual = 5 Then
        Set DDSGrama_tile = CreateDDSFromBitmap(DDraw, App.Path & "\graficos\grama_tile.bmp", True)
    End If

    torion.x = 300
    torion.y = 400
    
    zera_objetos

    'Lê o editor da fase
    file1 = FreeFile
    Open App.Path & "\fases\fase" & jogo.fase_atual & ".txt" For Input As #file1
    cont = 0
    'Erase editor_fase
    Do While Not EOF(file1)
        ReDim Preserve editor_fase(cont)
        Input #file1, strtemp
        editor_fase(cont).formacao = Val(Mid(strtemp, 1, 2))
        editor_fase(cont).tipo_inimigo = Val(Mid(strtemp, 4, 1))
        editor_fase(cont).num_inimigos = Val(Mid(strtemp, 6, 1))
        editor_fase(cont).extra = Val(Mid(strtemp, 8, 1))
        editor_fase(cont).inicia = Val(Mid(strtemp, 10, 5))
        cont = cont + 1
    Loop

    'Recria os tiles do cenário de fundo
    x = 0
    y = 0
    For tile_cont = 0 To 108
        If x > 10 Then
            x = 0
            y = y + 1
        End If
        Tiles(tile_cont).x = x * 64
        Tiles(tile_cont).y = y * 48
        Tiles(tile_cont).Existe = True
        Tiles(tile_cont).YoffSet = 0
        x = x + 1
    Next


End Sub

Public Sub inicializaDX()
    
    DoEvents

    inicializa_DD 'Inicializa Direct Draw
    inicializa_DI 'Inicializa o Direct Input
    inicializaDS 'Inicializa o Direct Sound
    inicializaDM 'Inicializa o Direct Music

End Sub


