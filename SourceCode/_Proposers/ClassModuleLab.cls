'--- ClassModule: Proposer
' Salve este código em um novo módulo de classe chamado "Proposer"

Option Explicit

Private pName As String
Private pParty As String
Private pCity As String

' Propriedade: 

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

' Propriedade: Partido
Public Property Get Party() As String
    Party = pParty
End Property

Public Property Let Party(value As String)
    pParty = value
End Property

' Propriedade: Cidade
Public Property Get City() As String
    City = pCity
End Property

Public Property Let City(value As String)
    pCity = value
End Property

' Método: Retorna uma descrição formatada do proponente
Public Function GetDescription() As String
    GetDescription = pName & " (" & pParty & ") - " & pCity
End Function