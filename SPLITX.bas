Function SPLITX(texto As String, direccion As Integer, delimitador1 As String, Optional delimitador2 As String = "", Optional delimitador3 As String = "") As Variant
    Dim partes() As String
    Dim i As Integer
    Dim salida() As Variant
    Dim delimitadores As String
    
    ' Hace la función volátil, es decir, se recalcula siempre que algo cambie en la hoja
    Application.Volatile
    
    ' Construimos la cadena de delimitadores
    delimitadores = delimitador1
    If delimitador2 <> "" Then delimitadores = delimitadores & "|" & delimitador2
    If delimitador3 <> "" Then delimitadores = delimitadores & "|" & delimitador3
    
    ' Dividimos el texto usando los delimitadores
    partes = SplitMultiDelims(texto, delimitadores)
    
    ' Dependiendo de la dirección (0 para filas, 1 para columnas)
    If direccion = 0 Then
        ReDim salida(1 To 1, 1 To UBound(partes) + 1)
        For i = LBound(partes) To UBound(partes)
            salida(1, i + 1) = partes(i)
        Next i
    ElseIf direccion = 1 Then
        ReDim salida(1 To UBound(partes) + 1, 1 To 1)
        For i = LBound(partes) To UBound(partes)
            salida(i + 1, 1) = partes(i)
        Next i
    End If
    
    SPLITX = salida
End Function

' Función auxiliar para dividir el texto usando múltiples delimitadores
Function SplitMultiDelims(Text As String, Delims As String) As Variant
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    RE.Global = True
    RE.IgnoreCase = True
    RE.Pattern = "[" & Delims & "]"
    SplitMultiDelims = RE.Replace(Text, "§")
    SplitMultiDelims = Split(SplitMultiDelims, "§")
End Function

' Función para registrar las opciones de ayuda de la función SPLITX
Sub RegistrarFuncionSPLITX()
    Application.MacroOptions _
        Macro:="SPLITX", _
        Description:="Divide el texto en celdas adyacentes usando uno o múltiples delimitadores (hasta tres). Esta función permite dividir un texto en partes más pequeñas y distribuirlas en celdas adyacentes, ya sea en filas o columnas, según los delimitadores especificados.", _
        Category:="Texto", _
        ArgumentDescriptions:=Array( _
            "texto: El texto que se va a dividir.", _
            "direccion: 0 para filas, 1 para columnas.", _
            "delimitador1: El primer delimitador.", _
            "delimitador2: (Opcional) El segundo delimitador.", _
            "delimitador3: (Opcional) El tercer delimitador.")
End Sub
