Attribute VB_Name = "M�dulo1"
Function SPLITX(texto As String, direccion As Integer, delimitador1 As String, Optional delimitador2 As String = "", Optional delimitador3 As String = "") As Variant
Attribute SPLITX.VB_Description = "Divide el texto en celdas adyacentes usando uno o m�ltiples delimitadores (hasta tres). Esta funci�n permite dividir un texto en partes m�s peque�as y distribuirlas en celdas adyacentes, ya sea en filas o columnas, seg�n los delimitadores especificados."
Attribute SPLITX.VB_ProcData.VB_Invoke_Func = " \n7"
    Dim partes() As String
    Dim i As Integer
    Dim salida() As Variant
    Dim celda As Range
    Dim delimitadores As String
    Dim ocupado As Boolean
    
    ' Hace la funci�n vol�til, es decir, que se recalcula siempre que algo cambie en la hoja
    Application.Volatile
    
    ' Construimos la cadena de delimitadores
    delimitadores = delimitador1
    If delimitador2 <> "" Then delimitadores = delimitadores & "|" & delimitador2
    If delimitador3 <> "" Then delimitadores = delimitadores & "|" & delimitador3
    
    ' Dividimos el texto usando los delimitadores
    partes = SplitMultiDelims(texto, delimitadores)
    
    ' Dependiendo de la direcci�n (0 para filas, 1 para columnas)
    If direccion = 0 Then
        ReDim salida(1 To 1, 1 To UBound(partes) + 1)
        ' Verifica si las celdas adyacentes hacia la derecha est�n vac�as
        For i = LBound(partes) To UBound(partes)
            Set celda = Application.Caller.Offset(0, i)
            If Not IsEmpty(celda) And celda.Address <> Application.Caller.Address Then
                MsgBox "Error: La celda " & celda.Address & " ya contiene datos.", vbExclamation
                ocupado = True
                Exit For
            End If
        Next i
        
        ' Si no hay celdas ocupadas, coloca los resultados
        If Not ocupado Then
            For i = LBound(partes) To UBound(partes)
                salida(1, i + 1) = partes(i)
            Next i
        End If
        
    ElseIf direccion = 1 Then
        ReDim salida(1 To UBound(partes) + 1, 1 To 1)
        ' Verifica si las celdas adyacentes hacia abajo est�n vac�as
        For i = LBound(partes) To UBound(partes)
            Set celda = Application.Caller.Offset(i, 0)
            If Not IsEmpty(celda) And celda.Address <> Application.Caller.Address Then
                MsgBox "Error: La celda " & celda.Address & " ya contiene datos.", vbExclamation
                ocupado = True
                Exit For
            End If
        Next i
        
        ' Si no hay celdas ocupadas, coloca los resultados
        If Not ocupado Then
            For i = LBound(partes) To UBound(partes)
                salida(i + 1, 1) = partes(i)
            Next i
        End If
    End If
    
    SPLITX = salida
End Function

' Funci�n auxiliar para dividir el texto usando m�ltiples delimitadores
Function SplitMultiDelims(Text As String, Delims As String) As Variant
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    RE.Global = True
    RE.IgnoreCase = True
    RE.Pattern = "[" & Delims & "]"
    SplitMultiDelims = RE.Replace(Text, "�")
    SplitMultiDelims = Split(SplitMultiDelims, "�")
End Function

' Funci�n para brindar ayuda
Sub RegistrarFunciones()
    Application.MacroOptions _
        Macro:="SPLITX", _
        Description:="Divide el texto en celdas adyacentes usando uno o m�ltiples delimitadores (hasta tres). Esta funci�n permite dividir un texto en partes m�s peque�as y distribuirlas en celdas adyacentes, ya sea en filas o columnas, seg�n los delimitadores especificados.", _
        Category:="Texto", _
        ArgumentDescriptions:=Array( _
            "texto: El texto que se va a dividir. Este es el texto completo que deseas separar en partes m�s peque�as.", _
            "direccion: 0 para filas, 1 para columnas. Este argumento determina si las partes del texto se distribuir�n horizontalmente (en filas) o verticalmente (en columnas).", _
            "delimitador1: El primer delimitador. Este es el car�cter o cadena de caracteres que se utilizar� para dividir el texto.", _
            "delimitador2: (Opcional) El segundo delimitador. Puedes especificar un segundo delimitador adicional para dividir el texto.", _
            "delimitador3: (Opcional) El tercer delimitador. Puedes especificar un tercer delimitador adicional para dividir el texto.")
End Sub



