Attribute VB_Name = "INTERFAZ"
Public autECLSession As Object
Public autECLPS As Object
Public autECLOIA As Object

Private Function Col_Letter(lngCol As Variant)
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Function a_reg()
    
    ubicacion = ActiveCell.Address(False, False)
    fila = ActiveCell.Row
    columna = Replace(ubicacion, fila, "")
    Range(Sheets("VAR").Range("B2") & fila).Select 'A
    cuenta = ActiveCell & ""
    
    
    If cuenta = "" Then
        Exit Function
    End If
    
    'automation variables
    Set autECLSession = CreateObject("pcomm.auteclsession")
    Set autECLPS = CreateObject("Pcomm.auteclps")
    Set autECLOIA = CreateObject("Pcomm.autecloia")
        
    'selecting the open session
    autECLSession.SetConnectionByName ("" & Sheets("VAR").Range("B1"))
        
    'from here the terminal logic begins in which the keystroke and reception of data from the terminal is managed based on coordinates of rows and columns
    If Trim(LTrim(autECLSession.autECLPS.GetText(2, 22, 27) & "")) = "Revisar Maestro de Clientes" Then
    
        
        If IsEmpty(fila) Or IsNull(autECLSession) Then
            Range(Sheets("VAR").Range("B24") & fila) = "NO SE PUEDE EJECUTAR LA MACRO" 'L
            Exit Function
        Else
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[eraseeof]", 9, 5
            autECLSession.autECLOIA.WaitForInputReady
            
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "" & cuenta, 9, 5
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[enter]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            
            If Trim(LTrim(autECLSession.autECLPS.GetText(3, 27, 7) & "")) = "WGEOCAR" Then
                autECLSession.autECLPS.SendKeys "[enter]"
                autECLSession.autECLOIA.WaitForAppAvailable
                autECLSession.autECLOIA.WaitForInputReady
            End If
            
            autECLSession.autECLPS.SendKeys "1"
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[enter]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[enter]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[enter]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[enter]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady

            autECLSession.autECLPS.SendKeys "[pf2]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[right]"
            autECLSession.autECLOIA.WaitForInputReady
           
            cta = Trim(LTrim(autECLSession.autECLPS.GetText(3, 27, 7) & ""))
            
            If cuenta = cta Then
               

                'data collection
                'DIRECCION
                Range(Sheets("VAR").Range("B4") & fila) = LTrim(Trim("" & autECLSession.autECLPS.GetText(14, 18, 44)))
                'GEOCODIGO
                Range(Sheets("VAR").Range("B6") & fila) = Format(LTrim(Trim("" & autECLSession.autECLPS.GetText(18, 13, 2))), "00") & "." & Format(LTrim(Trim("" & autECLSession.autECLPS.GetText(18, 45, 2))), "00") & "." & Format(LTrim(Trim("" & autECLSession.autECLPS.GetText(19, 13, 2))), "00") & "." & Format(LTrim(Trim("" & autECLSession.autECLPS.GetText(20, 7, 4))), "000") & "." & Format(LTrim(Trim("" & autECLSession.autECLPS.GetText(20, 73, 7))), "0000000")
                'NAME
                Range(Sheets("VAR").Range("B3") & fila) = LTrim(Trim("" & autECLSession.autECLPS.GetText(3, 35, 35)))
                'DNI
                Range(Sheets("VAR").Range("B5") & fila) = "'" & Trim("" & autECLSession.autECLPS.GetText(4, 10, 13))
                
                'MEDIDOR
                Range(Sheets("VAR").Range("B7") & fila) = Trim("" & autECLSession.autECLPS.GetText(6, 11, 20))
                
                
                autECLSession.autECLOIA.WaitForInputReady
                autECLSession.autECLPS.SendKeys "[pf2]"
                autECLSession.autECLOIA.WaitForAppAvailable
                autECLSession.autECLOIA.WaitForInputReady
                autECLSession.autECLPS.SendKeys "[right]"
                autECLSession.autECLOIA.WaitForInputReady
                

                'STATUS
                Range(Sheets("VAR").Range("B8") & fila) = LTrim(Trim("" & autECLSession.autECLPS.GetText(10, 30, 20)))
                
                
                
                autECLSession.autECLOIA.WaitForInputReady
                autECLSession.autECLPS.SendKeys "[pf12]"
                autECLSession.autECLOIA.WaitForAppAvailable
                autECLSession.autECLOIA.WaitForInputReady
                autECLSession.autECLPS.SendKeys "[right]"
                autECLSession.autECLOIA.WaitForInputReady
        
            Else
            
                Range(Sheets("VAR").Range("B3") & fila) = "LA CUENTA / CODIGO DE CLIENTE / SERVICIO / SUMINISTRO " & cuenta & " NO ESTA EN EL SISTEMA..." 'L
                buscarMedidor = False
            
            End If
        
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[pf12]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[pf12]"
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            
          
            
            autECLSession.autECLOIA.WaitForAppAvailable
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[up]"
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[tab]"
            autECLSession.autECLOIA.WaitForInputReady
            autECLSession.autECLPS.SendKeys "[eraseeof]"
            autECLSession.autECLOIA.WaitForInputReady
     
       
        End If
        
    End If
    'end of logic
        
End Function

Sub MasiveSearch()

    Dim celda As Range

    If Selection.Cells.Rows.Count > 1 Then
            
        RF = Selection.Cells(Selection.Cells.Rows.Count, 1).Row
        RI = Selection.Cells(1, 1).Row
            
        COLU = Col_Letter(Selection.Cells.Column)
        
        For Each celda In Range(COLU & RI & ":" & COLU & RF).SpecialCells(xlCellTypeVisible)
            celda.Select
            If ActiveCell <> "" Then
            
                
                a_reg
                    


            End If
        Next
    Else
        If ActiveCell <> "" Then
        
                        
            a_reg


            
        End If
    End If
    
End Sub
