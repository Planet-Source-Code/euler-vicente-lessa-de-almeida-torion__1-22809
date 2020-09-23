VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If jogo.status = 1 Then 'se estiver na abertura
    If jogo.item_abertura = 1 Then
        If KeyCode = vbKeyDown Then
            If opcao = 5 Then
                opcao = 1
            Else
                opcao = opcao + 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            If opcao = 1 Then
                opcao = 5
            Else
                opcao = opcao - 1
            End If
        ElseIf KeyCode = vbKeySpace Then
            If opcao = 1 Then
                reseta_jogo 'novo jogo
            ElseIf opcao = 2 Then
                opcao = 1
                jogo.item_abertura = 2 'opções
            ElseIf opcao = 3 Then
                jogo.item_abertura = 3 'opções
            ElseIf opcao = 4 Then
                jogo.item_abertura = 4
            ElseIf opcao = 5 Then
                End
            End If
        End If
    ElseIf jogo.item_abertura = 2 Then 'tela de opçoes
        If KeyCode = vbKeyDown Then
            If opcao = 2 Then
                opcao = 1
            Else
                opcao = opcao + 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            If opcao = 1 Then
                opcao = 2
            Else
                opcao = opcao - 1
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If opcao = 1 Then
                If jogo.nivel > 1 Then 'dificuldade
                    jogo.nivel = jogo.nivel - 1
                Else
                    jogo.nivel = 3
                End If
            ElseIf opcao = 2 Then 'FPS
                If jogo.FPS Then
                    jogo.FPS = False
                Else
                    jogo.FPS = True
                End If
            End If
        ElseIf KeyCode = vbKeyRight Then
            If opcao = 1 Then
                If jogo.nivel < 3 Then 'dificuldade
                    jogo.nivel = jogo.nivel + 1
                Else
                    jogo.nivel = 1
                End If
            ElseIf opcao = 2 Then 'FPS
                If jogo.FPS Then
                    jogo.FPS = False
                Else
                    jogo.FPS = True
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If jogo.status = 2 Then 'se estiver jogando
        If KeyAscii = vbKeyEscape Then
            para_midi
            jogo.status = 1 'volta para a abertura
            jogo.fase_atual = 0
        ElseIf KeyAscii = 112 Or KeyAscii = 80 Then 'TECLA P
            jogo.status = 6
            dsPausa.SetCurrentPosition 0
            dsPausa.Play DSBPLAY_DEFAULT
        End If
    ElseIf jogo.status = 1 Then 'se estiver na abertura
        If jogo.item_abertura = 1 Then
            If KeyAscii = vbKeyEscape Then 'esc termina o jogo
                End
            End If
        Else
            If KeyAscii = vbKeyEscape Then
                jogo.item_abertura = 1
                opcao = 1
            End If
        End If
    ElseIf jogo.status = 3 Then 'se estiver na tela de congratulações
        If editor_fase(jogo.indice_editor_fase - 1).inicia + 5 < jogo.contador Then
            If jogo.recorde < jogo.placar Then
                inclui_recorde
            End If
            destroi_fase
            para_midi
            jogo.status = 1 'volta para a abertura
            jogo.fase_atual = 0
        End If
    ElseIf jogo.status = 5 Then 'se estiver na tela de fim de jogo
        If editor_fase(1).inicia + 5 < jogo.contador Then
            If jogo.recorde < jogo.placar Then
                inclui_recorde
            End If
            destroi_fase
            para_midi
            jogo.status = 1 'volta para a abertura
            jogo.fase_atual = 0
        End If
    ElseIf jogo.status = 6 Then 'se estiver pausado
        If KeyAscii = 112 Or KeyAscii = 27 Or KeyAscii = 80 Then 'TECLA P OU ESC
            jogo.status = 2
            dsPausa.SetCurrentPosition 0
            dsPausa.Play DSBPLAY_DEFAULT
        End If
    ElseIf jogo.status = 7 Then 'se estiver digitando o nome do recordista
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            destroi_fase
            para_midi
            jogo.status = 1 'volta para a abertura
            jogo.fase_atual = 0
        Else
            If KeyAscii = 8 Then 'Back space
                tmpString = Mid(tmpString, 1, Len(tmpString) - 1)
            Else
                tmpString = tmpString & VBA.Chr$(KeyAscii)
            End If
        End If
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    para_midi
    destroi_DS
    destroi_DD
    destroi_DI
    Set DX = Nothing

End Sub

Private Sub Timer1_Timer()

    If jogo.status <> 6 Then
        jogo.contador = jogo.contador + 1
    End If

    If Module2.torion.invencibil > 0 Then
        Module2.torion.invencibil = Module2.torion.invencibil - 1
    Else
        Module2.torion.num_frames = 2
    End If

End Sub
