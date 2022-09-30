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

Sub Get_Data_From_File3()
Dim FileToOpen As Variant
Dim OpenBook As Workbook

'Dim LibroDestino As Workbook
'Set LibroDestino = ThisWorkbook

Application.ScreenUpdating = False

    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xls*),*xls*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        
        'Name
        OpenBook.Worksheets("Completar datos").Range("D3").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N13")
        
        'Name company
        OpenBook.Worksheets("Completar datos").Range("D3").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N14")
        
        'Id Compnay
        OpenBook.Worksheets("Completar datos").Range("D5").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N15")
        
        'Address
        OpenBook.Worksheets("Completar datos").Range("D10").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N16")
        
        'Location
        OpenBook.Worksheets("Completar datos").Range("D11").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N17")
        
        'Legal representative
        OpenBook.Worksheets("Completar datos").Range("D6").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N23")
        
        'Id Legal representative
        OpenBook.Worksheets("Completar datos").Range("D7").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N24")
 
        'Email
        OpenBook.Worksheets("Completar datos").Range("D21").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N41")
        
        'Bank 
        OpenBook.Worksheets("Completar datos").Range("D18").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N42")
    
        'Bank Account
        OpenBook.Worksheets("Completar datos").Range("D20").Copy Workbooks("Nueva Ficha Proveedores.xlsx").Worksheets("T3").Range("N43")
       
              
        OpenBook.Close False
    
    End If

Application.ScreenUpdating = False

End Sub
    
Private Sub CommandButton3_Click()
    
    'Dim LibroActual As Workbook
        
    'For Each LibroActual In Workbooks
    
     '   LibroActual.Close SaveChanges:=False
    Application.Quit
    ThisWorkbook.Close SaveChanges:=False
     'Next LibroActual
    'ThisWorkbook.Close SaveChanges:=False
    'ThisWorkbook.Close
    'Application.DisplayAlerts = False
    
End Sub

    
