Attribute VB_Name = "Module5"
'----------------- DIRECT Music -------------------------
Public perf As DirectMusicPerformance                   'DirectMusic Performance object
Public seg As DirectMusicSegment                        'DirectMusic Segment
Public segstate As DirectMusicSegmentState              'DirectMusic Segment State
Public loader As DirectMusicLoader                      'DirectMusic Loader
Public Sub TocaMidi(MidiString As String)
    '
    'Essa fun��o toca as musicas do jogo continuamente
    '
    If jogo.fase_atual = 3 Or MidiString = "abertura.mid" Then 'porque a musica da fase 3 est� baixa
        perf.SetMasterVolume (1200)
    Else
        perf.SetMasterVolume (400)
    End If
    If (seg Is Nothing) And (segstate Is Nothing) Then 'Se o segmento da musica est� vazio . . .
        loader.SetSearchDirectory App.Path & "\musicas\" ' Busca as musicas em app.path
        Set seg = loader.LoadSegment(MidiString) 'Carrega MIDI passado pela fun��o
        Set segstate = perf.PlaySegment(seg, 0, 0)  'Toca MIDI carregado no Segmento
    Else 'Se o segmento j� foi carregado . . .
        If Not perf.IsPlaying(seg, segstate) Then 'Se o MIDI n�o est� tocando mais, carrega tudo
                                                  'novamente para que a musica toque continuamente
            loader.SetSearchDirectory App.Path & "\musicas\" ' Busca as musicas em app.path
            
            Set seg = loader.LoadSegment(MidiString) 'Carrega MIDI passado pela fun��o
            Set segstate = perf.PlaySegment(seg, 0, 0)  'Toca MIDI carregado no Segmento
        End If
    End If
    
End Sub

Public Sub para_midi()
    '
    'Para de tocar a musica do jogo
    '
    If Not (seg Is Nothing) And Not (segstate Is Nothing) Then 'Para a m�sica
        If perf.IsPlaying(seg, segstate) Then 'Se a musica est� tocando . . .
            Call perf.Stop(seg, segstate, 0, 0) 'Para a musica que est� tocando
        End If
    End If


End Sub

Public Sub inicializaDM()
    '
    'Inicializa Direct Music
    '
    Set loader = DX.DirectMusicLoaderCreate() 'Cria um novo DMusic Loader
    Set perf = DX.DirectMusicPerformanceCreate() 'Cria um novo objeto DMusic Performance
    Call perf.Init(Nothing, 0) 'Initializa o objeto
    perf.SetPort -1, 16 'Set the default port to 4 sets(64) of voices
    Call perf.SetMasterAutoDownload(True)
    


End Sub

