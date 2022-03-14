Option Explicit

'set string values
Private sProduct As String
Private sYear198990 As String
Private sYear201819 As String

'set individual properties
Property Get product() As String
    product = sProduct
End Property

Property Get year198990() As String
    year198990 = sYear198990
End Property

Property Get year201819() As String
    year201819 = sYear201819
End Property

'set all class properties called from module
Public Sub SetAll(ByVal pProduct As String, _
    ByVal pYear198990 As String, _
    ByVal pYear201819 As String)
    
    sProduct = pProduct
    sYear198990 = pYear198990
    sYear201819 = pYear201819

End Sub
