Private Sub CommandButton1_Click()

Dim path As String
Dim fname As String

path = "C:\Users\cristian.gomez\Desktop\"
fname = Range("N13")
'fname = invno & " - " & Range("A4")

Application.DisplayAlerts = False

Hoja1.Copy
ActiveSheet.Shapes("CommandButton1").Delete
ActiveSheet.Shapes("CommandButton2").Delete
    
    With ActiveWorkbook
        .SaveAs Filename:=path & fname, FileFormat:=51
        .Close
    End With

MsgBox "Archivo creado en el Escritorio: " + fname

'MsgBox "Your next invoice number is " & invno + 1

'Range("D3") = invno + 1

ThisWorkbook.Save

Application.DisplayAlerts = True

End Sub

Private Sub CommandButton2_Click()
Dim FileToOpen As Variant
Dim OpenBook As Workbook

'Dim LibroDestino As Workbook
'Set LibroDestino = ThisWorkbook

Application.ScreenUpdating = False

    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xls*),*xls*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        
        'Nombre CC
        OpenBook.Worksheets("Completar datos").Range("D3").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N13").PasteSpecial (xlPasteValues)
        
        'Nombre CC
        OpenBook.Worksheets("Completar datos").Range("D3").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N14").PasteSpecial (xlPasteValues)
        
        'Rut CC
        OpenBook.Worksheets("Completar datos").Range("D5").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N15").PasteSpecial (xlPasteValues)
        
        'Domicilio
        OpenBook.Worksheets("Completar datos").Range("D10").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N16").PasteSpecial (xlPasteValues)
        
        'Comuna
        OpenBook.Worksheets("Completar datos").Range("D11").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N17").PasteSpecial (xlPasteValues)
        
        'Rep. Legal
        OpenBook.Worksheets("Completar datos").Range("D6").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N23").PasteSpecial (xlPasteValues)
        
        'Rut Rep. Legal
        OpenBook.Worksheets("Completar datos").Range("D7").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N24").PasteSpecial (xlPasteValues)
 
        'B correo electrónico
        OpenBook.Worksheets("Completar datos").Range("D21").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N31").PasteSpecial (xlPasteValues)
       
        'Rut Rep. Legal
        OpenBook.Worksheets("Completar datos").Range("D7").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N32").PasteSpecial (xlPasteValues)
        
        'B correo electrónico
        OpenBook.Worksheets("Completar datos").Range("D21").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N39").PasteSpecial (xlPasteValues)
        
        'B Nombre Banco
        OpenBook.Worksheets("Completar datos").Range("D18").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N40").PasteSpecial (xlPasteValues)
    
        'B N de cuenta
        OpenBook.Worksheets("Completar datos").Range("D20").Copy
        Workbooks("Nueva Ficha Proveedores.xlsm").Worksheets("T3").Range("N41").PasteSpecial (xlPasteValues)
        
        OpenBook.Close False
    
    End If
    

Application.ScreenUpdating = False
End Sub

Private Sub CommandButton3_Click()
    
    'Dim LibroActual As Workbook
    'For Each LibroActual In Workbooks
     '   LibroActual.Close SaveChanges:=False
     'Next LibroActual
    'ThisWorkbook.Close SaveChanges:=False
    'ThisWorkbook.Close
    'Application.DisplayAlerts = False
   Dim wb As Workbook
   Set wb = ActiveWorkbook
   Range("M:M").Copy Range("N:N")
   wb.Save
   Application.Quit
   wb.Close
    
End Sub
    
