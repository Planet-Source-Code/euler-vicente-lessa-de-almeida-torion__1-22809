Attribute VB_Name = "Module4"
Public DS As DirectSound                            'Objeto Direct Sound
Public dsMissel As DirectSoundBuffer
Public dsArma1 As DirectSoundBuffer
Public dsArma2 As DirectSoundBuffer
Public dsArma3 As DirectSoundBuffer
Public dsTiro As DirectSoundBuffer
Public dsExplosao As DirectSoundBuffer
Public dsExtra As DirectSoundBuffer
Public dsVidaExtra As DirectSoundBuffer
Public dsPausa As DirectSoundBuffer
Public bdesc As DSBUFFERDESC      'variavel que contem a descricao do direct sound buffer
Public DSwaveFormat As WAVEFORMATEX

Public Sub destroi_DS()

    Set dsMissel = Nothing
    Set dsArma1 = Nothing
    Set dsArma2 = Nothing
    Set dsArma3 = Nothing
    Set dsTiro = Nothing
    Set dsExplosao = Nothing
    Set dsExtra = Nothing
    Set dsVidaExtra = Nothing
    Set dsPausa = Nothing
    Set DS = Nothing

End Sub

Public Sub inicializaDS()
    '
    'Essa sub inicializa todos os efeitos sonoros usados no jogo
    '
    'Cria o objeto direct sound usando o mecanismo default de som
    Set DS = DX.DirectSoundCreate("")
    'seta o cooperative level o form do jogo
    DS.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
   
    bdesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    With DSwaveFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 22050
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
   

    'Carrega todos os .wav usados no jogo
    Set dsMissel = DS.CreateSoundBufferFromFile(App.Path & "\sons\missel.wav", bdesc, DSwaveFormat)
    
    Set dsArma1 = DS.CreateSoundBufferFromFile(App.Path & "\sons\arma1.wav", bdesc, DSwaveFormat)
    
    Set dsArma2 = DS.CreateSoundBufferFromFile(App.Path & "\sons\arma2.wav", bdesc, DSwaveFormat)
    
    Set dsArma3 = DS.CreateSoundBufferFromFile(App.Path & "\sons\arma3.wav", bdesc, DSwaveFormat)
    
    Set dsTiro = DS.CreateSoundBufferFromFile(App.Path & "\sons\tiro.wav", bdesc, DSwaveFormat)

    Set dsExplosao = DS.CreateSoundBufferFromFile(App.Path & "\sons\explosao.wav", bdesc, DSwaveFormat)

    Set dsExtra = DS.CreateSoundBufferFromFile(App.Path & "\sons\extra.wav", bdesc, DSwaveFormat)

    Set dsVidaExtra = DS.CreateSoundBufferFromFile(App.Path & "\sons\vida_extra.wav", bdesc, DSwaveFormat)

    Set dsPausa = DS.CreateSoundBufferFromFile(App.Path & "\sons\pausa.wav", bdesc, DSwaveFormat)

End Sub
