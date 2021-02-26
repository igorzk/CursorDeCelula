VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CursorDeCelula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lin As Long
Private col As Long
Private Sh As Worksheet
Private Const PROCURA_MAXBRANCAS As Long = 100

Private Sub Class_Initialize()

    lin = 1
    col = 1
    Err = False
    Set Sh = ThisWorkbook.Sheets(1)
    
End Sub

Public Property Get Erro() As Boolean

    If lin = Application.Rows.Count + 1 Or col = Application.Columns.Count + 1 Then
        Erro = True
    Else
        Erro = False
    End If

End Property

Public Sub SinalizaErro()
    
    lin = Application.Rows.Count + 1
    col = Application.Columns.Count + 1

End Sub

Public Property Get Celula() As Range

    Set Celula = Sh.Cells(lin, col)

End Property

Public Property Set Planilha(ByVal Plan As Worksheet)

    Set Sh = Plan

End Property

Public Property Get Planilha() As Worksheet

    Set Planilha = Sh

End Property

Public Property Get Linha() As Long

    Linha = lin

End Property

Public Property Let Linha(ByVal NumLinha As Long)

    If NumLinha > 0 And NumLinha <= Application.Rows.Count + 1 Then
        lin = NumLinha
    Else
        RaiseErrorEnderecoNegativo
    End If

End Property

Public Function IncrementaLinha(Optional ByVal NumeroDeLinhas As Long = 1) As CursorDeCelula

    Me.Linha = Me.Linha + NumeroDeLinhas
    
    Set IncrementaLinha = Me

End Function

Public Property Get Coluna() As Long

    Coluna = col

End Property

Public Property Let Coluna(ByVal NumColuna As Long)

    If NumColuna > 0 And NumColuna <= Application.Columns.Count + 1 Then
        col = NumColuna
    Else
        RaiseErrorEnderecoNegativo
    End If

End Property

Public Function IncrementaColuna(Optional ByVal NumeroDeColunas As Long = 1) As CursorDeCelula

    Me.Coluna = Me.Coluna + NumeroDeColunas
    
    Set IncrementaColuna = Me

End Function

Public Property Get EnderecoAbsoluto() As String

    EnderecoAbsoluto = Me.Celula.Address

End Property

Public Property Get EnderecoRelativo() As String

    EnderecoRelativo = Replace(Me.Celula.Address, "$", vbNullString)

End Property

Public Property Get ReferenciaCompleta() As String

    ReferenciaCompleta = "'" & Me.Planilha.name & "'!" & Me.EnderecoAbsoluto

End Property

Public Function Clona() As CursorDeCelula
    
    Dim NovoCursor As New CursorDeCelula
    
    NovoCursor.MudaEndereco Me.Planilha, Me.Linha, Me.Coluna
    
    Set Clona = NovoCursor

End Function

Public Function MudaEndereco(Optional ByVal Plan As Worksheet = Nothing, _
                                Optional ByVal NumLinha As Long = -1, Optional ByVal NumColuna As Long = -1) As CursorDeCelula

    If Not Plan Is Nothing Then
        Set Me.Planilha = Plan
    End If
    
    If NumLinha <> -1 Then
        Me.Linha = NumLinha
    End If
    
    If NumColuna <> -1 Then
        Me.Coluna = NumColuna
    End If
    
    Set MudaEndereco = Me

End Function

Public Function AchaExpressaoPartindoDeMeuEndereco(ByVal Expressao As Variant) As CursorDeCelula

    Dim NovoCursor As CursorDeCelula
    Dim LinBrancas As Long
    Dim ColBrancas As Long
    Dim AchouConteudoLinha As Boolean
    
    Set NovoCursor = Me.Clona
    
    Do While NovoCursor.Celula.value <> Expressao And LinBrancas <= PROCURA_MAXBRANCAS
        ColBrancas = 0
        AchouConteudoLinha = False
        If IsEmpty(NovoCursor.Celula) Or NovoCursor.Celula.value = vbNullString Then
            ColBrancas = ColBrancas + 1
        Else
            AchouConteudoLinha = True
        End If
        NovoCursor.IncrementaColuna
        Do While NovoCursor.Celula.value <> Expressao And ColBrancas <= PROCURA_MAXBRANCAS
            If IsEmpty(NovoCursor.Celula) Or NovoCursor.Celula.value = vbNullString Then
                ColBrancas = ColBrancas + 1
            Else
                ColBrancas = 0
                AchouConteudoLinha = True
            End If
            NovoCursor.IncrementaColuna
        Loop
        If AchouConteudoLinha = False Then
            ColBrancas = ColBrancas + 1
        Else
            ColBrancas = 0
        End If
        If NovoCursor.Celula.value <> Expressao Then
            NovoCursor.IncrementaLinha
            NovoCursor.Coluna = Me.Coluna
        End If
    Loop
    
    If NovoCursor.Celula.value <> Expressao Then
        NovoCursor.SinalizaErro
    End If
    
    Set AchaExpressaoPartindoDeMeuEndereco = NovoCursor
    
End Function

Public Function AchaNaMinhaColunaEnderecoComExpressao(ByVal Expressao As Variant) As CursorDeCelula

    Dim NovoCursor As CursorDeCelula
    Dim brancas As Long
    
    Set NovoCursor = Me.Clona
        
    brancas = 0
    Do While NovoCursor.Celula.value <> Expressao And brancas <= PROCURA_MAXBRANCAS
        If IsEmpty(NovoCursor.Celula) Or NovoCursor.Celula.value = vbNullString Then
            brancas = brancas + 1
        Else
            brancas = 0
        End If
        NovoCursor.IncrementaLinha
    Loop
    
    If NovoCursor.Celula.value <> Expressao Then
        NovoCursor.SinalizaErro
    End If
    
    Set AchaNaMinhaColunaEnderecoComExpressao = NovoCursor

End Function

Public Function AchaNaMinhaColunaEnderecoContendoExpressao(ByVal Expressao As Variant) As CursorDeCelula

    Dim NovoCursor As CursorDeCelula
    Dim brancas As Long
    
    Set NovoCursor = Me.Clona
        
    brancas = 0
    Do While Not InStr(NovoCursor.Celula.value, Expressao) > 0 And brancas <= PROCURA_MAXBRANCAS
        If IsEmpty(NovoCursor.Celula) Or NovoCursor.Celula.value = vbNullString Then
            brancas = brancas + 1
        Else
            brancas = 0
        End If
        NovoCursor.IncrementaLinha
    Loop
    
    If Not InStr(NovoCursor.Celula.value, Expressao) > 0 Then
        NovoCursor.SinalizaErro
    End If
    
    Set AchaNaMinhaColunaEnderecoContendoExpressao = NovoCursor

End Function


Public Function AchaNaMinhaLinhaEnderecoComExpressao(ByVal Expressao As Variant) As CursorDeCelula

    Dim NovoCursor As CursorDeCelula
    Dim brancas As Long
    
    Set NovoCursor = Me.Clona
        
    brancas = 0
    Do While NovoCursor.Celula.value <> Expressao And brancas <= PROCURA_MAXBRANCAS
        If IsEmpty(NovoCursor.Celula) Or NovoCursor.Celula.value = vbNullString Then
            brancas = brancas + 1
        Else
            brancas = 0
        End If
        NovoCursor.IncrementaColuna
    Loop
    
    If NovoCursor.Celula.value <> Expressao Then
        NovoCursor.SinalizaErro
    End If
    
    Set AchaNaMinhaLinhaEnderecoComExpressao = NovoCursor

End Function

Public Function AchaNaMinhaLinhaEnderecoContendoExpressao(ByVal Expressao As Variant) As CursorDeCelula

    Dim NovoCursor As CursorDeCelula
    Dim brancas As Long
    
    Set NovoCursor = Me.Clona
        
    brancas = 0
    Do While Not InStr(NovoCursor.Celula.value, Expressao) > 0 And brancas <= PROCURA_MAXBRANCAS
        If IsEmpty(NovoCursor.Celula) Or NovoCursor.Celula.value = vbNullString Then
            brancas = brancas + 1
        Else
            brancas = 0
        End If
        NovoCursor.IncrementaColuna
    Loop
    
    If Not InStr(NovoCursor.Celula.value, Expressao) > 0 Then
        NovoCursor.SinalizaErro
    End If
    
    Set AchaNaMinhaLinhaEnderecoContendoExpressao = NovoCursor

End Function

Private Sub RaiseErrorEnderecoNegativo()

    Err.Raise 438, TypeName(Me), TypeName(Me) & ": Endereço inválido"

End Sub
