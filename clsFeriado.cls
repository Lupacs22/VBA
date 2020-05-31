VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFeriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ProximaData As Date
Private FeriadosSalvos As Object

Private Sub Class_Initialize()

    Set FeriadosSalvos = CreateObject("scripting.dictionary")
    
    FeriadosSalvos.Add "Confraternização", "01/01"
    FeriadosSalvos.Add "Aniversario", "25/01"
    FeriadosSalvos.Add "Tiradentes", "21/04"
    FeriadosSalvos.Add "Trabalho", "01/05"
    FeriadosSalvos.Add "Constitucionalista", "09/09"
    FeriadosSalvos.Add "Independência", "07/09"
    FeriadosSalvos.Add "Aparecida", "12/10"
    FeriadosSalvos.Add "Finados", "02/11"
    FeriadosSalvos.Add "República", "15/11"
    FeriadosSalvos.Add "Consciência", "20/11"
    FeriadosSalvos.Add "Preparação", "23/12"
    FeriadosSalvos.Add "Natal", "25/12"
    FeriadosSalvos.Add "Último", "31/12"
    
End Sub

Public Function adicionarFeriado(Nome As String, Data As String)
    FeriadosSalvos.Add Nome, Data
End Function

Public Function proximoDiaUtil(ByVal DataAtual As Date, Optional ByVal quantidade = 1) As Date
    
    ProximaData = DateAdd("d", quantidade, DataAtual)
        
    FeriadosVariaveis
        
    Do While tipoDia(ProximaData) <> 0
        ProximaData = ProximaData + 1 * Sgn(quantidade)
    Loop
    
    proximoDiaUtil = ProximaData
    
    FeriadosSalvos.Remove "Segunda"
    FeriadosSalvos.Remove "Carnaval"
    FeriadosSalvos.Remove "Santa"
    FeriadosSalvos.Remove "Pascoa"
    FeriadosSalvos.Remove "Corpus"
    
End Function

Public Function tipoDia(Data As Date) As Integer
    ' 0 - Dia útil | 1 - Fim de semana | 2 - Feriado
    
    Dim key As Variant
    Dim Dia As Integer, Mes As Integer
        
    ' Determina se é fim de semana
    If Application.WorksheetFunction.Weekday(CDate(Data)) = 1 Or _
    Application.WorksheetFunction.Weekday(CDate(Data)) = 7 Then
        tipoDia = 1
        Exit Function
    End If
    
    ' Determina se é feriado
    For Each key In FeriadosSalvos.Keys
        Dia = Split(FeriadosSalvos(key), "/")(0)
        Mes = Split(FeriadosSalvos(key), "/")(1)
        If Day(Data) = Dia And Month(Data) = Mes Then
            tipoDia = 2
            Exit Function
        End If
    Next key
    
    ' Retorna dia útil
    tipoDia = 0
    
End Function

Private Sub FeriadosVariaveis()
    
    Dim Ano As Integer
    Dim k1 As Integer
    Dim k2 As Integer
    Dim k3 As Integer
    Dim k4 As Integer
    Dim k5 As Integer
    Dim k6 As Integer
    Dim k7 As Integer
    Dim DiaPascoa As Integer
    Dim MesPascoa As Integer
    
    Dim Santa As Date
    Dim Segunda As Date
    Dim Carnaval As Date
    Dim Corpus As Date
    
    Ano = Year(ProximaData)
    
    k6 = 24
    k7 = 5
    
    k1 = Ano Mod 19
    k2 = Ano Mod 4
    k3 = Ano Mod 7
    k4 = (19 * k1 + k6) Mod 30
    k5 = ((2 * k2) + (4 * k3) + (6 * k4) + k7) Mod 7
        
    If 22 + k4 + k5 <= 31 Then
        DiaPascoa = 22 + k4 + k5
        MesPascoa = 3
    ElseIf k4 + k5 - 9 > 25 Then
        DiaPascoa = k4 + k5 - 9 - 7
        MesPascoa = 4
    Else
        DiaPascoa = k4 + k5 - 9
        MesPascoa = 4
    End If
        
    Santa = DateAdd("d", -2, CDate(DiaPascoa & "/" & MesPascoa & "/" & Ano))
    Segunda = DateAdd("d", -48, CDate(DiaPascoa & "/" & MesPascoa & "/" & Ano))
    Carnaval = DateAdd("d", -47, CDate(DiaPascoa & "/" & MesPascoa & "/" & Ano))
    Corpus = DateAdd("d", 60, CDate(DiaPascoa & "/" & MesPascoa & "/" & Ano))
    
    FeriadosSalvos.Add "Segunda", Day(Segunda) & "/" & Month(Segunda)
    FeriadosSalvos.Add "Carnaval", Day(Carnaval) & "/" & Month(Carnaval)
    FeriadosSalvos.Add "Santa", Day(Santa) & "/" & Month(Santa)
    FeriadosSalvos.Add "Pascoa", DiaPascoa & "/" & MesPascoa
    FeriadosSalvos.Add "Corpus", Day(Corpus) & "/" & Month(Corpus)
    
End Sub

