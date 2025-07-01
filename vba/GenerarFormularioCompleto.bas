' === Crear formulario con campos y lógica ===
Sub GenerarFormularioCompleto()
    Dim frm As Object
    Dim ctrl As MSForms.Control
    Dim nombres(), i As Integer

    ' Crear el formulario
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3) ' vbext_ct_MSForm
    frm.Name = "UserForm1"

    ' Lista de controles (TextBoxes y ComboBoxes)
    nombres = Array( _
        "txtInvoice", "txtStartDate", "txtAmount", "txtACH", "txtSentTo", _
        "cmbActionDone", "txtNote", "cmbSystem", "cmbCollector", "cmbCashAppOwner", _
        "cmbStatusConsolidation", "cmbStatusCashApp", "txtFollowUp", "txtAnswer" _
    )
    
    For i = 0 To UBound(nombres)
        If Left(nombres(i), 3) = "txt" Then
            Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1", nombres(i))
        Else
            Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1", nombres(i))
        End If
        ctrl.Top = 10 + (i * 20)
        ctrl.Left = 100
        ctrl.Width = 150
        
        ' Label
        Set lbl = frm.Designer.Controls.Add("Forms.Label.1", "lbl" & nombres(i))
        lbl.Caption = nombres(i)
        lbl.Top = ctrl.Top
        lbl.Left = 10
        lbl.Width = 90
    Next i

    ' Botón de envío
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1", "btnSubmit")
    ctrl.Caption = "Enviar"
    ctrl.Top = ctrl.Top + 30
    ctrl.Left = 100
    ctrl.Width = 80

    ' === Código del formulario ===
    With frm.CodeModule
        .InsertLines .CountOfLines + 1, Formulario_InitializeCode()
        .InsertLines .CountOfLines + 1, Formulario_SubmitCode()
    End With

    ' === Crear módulo para abrir formulario ===
    Dim modCode As Object
    Set modCode = ThisWorkbook.VBProject.VBComponents.Add(1)
    modCode.Name = "modFormulario"

    With modCode.CodeModule
        .InsertLines 1, "Sub AbrirFormulario()" & vbCrLf & _
                         "    UserForm1.Show" & vbCrLf & _
                         "End Sub" & vbCrLf & vbCrLf & CrearHojaDataCode()
    End With

    MsgBox "Formulario creado exitosamente.", vbInformation
End Sub

' === Código del formulario: UserForm_Initialize ===
Function Formulario_InitializeCode() As String
Formulario_InitializeCode = _
"Private Sub UserForm_Initialize()" & vbCrLf & _
"    cmbActionDone.AddItem ""Asking Data Entry""" & vbCrLf & _
"    cmbActionDone.AddItem ""Asking Billing""" & vbCrLf & _
"    cmbSystem.AddItem ""ACCOUNTABLE HEALTHCARE STAFFING""" & vbCrLf & _
"    cmbSystem.AddItem ""Adventist Health Adventist Health System/West - Aya Invoices""" & vbCrLf & _
"    cmbCollector.AddItem ""Yoselyn Lizano""" & vbCrLf & _
"    cmbCollector.AddItem ""Tracy Clemons""" & vbCrLf & _
"    cmbCashAppOwner.AddItem ""Masis Alvarado Alfredo""" & vbCrLf & _
"    cmbCashAppOwner.AddItem ""Navarrete Sandino Alma""" & vbCrLf & _
"    cmbStatusConsolidation.AddItem ""Completed""" & vbCrLf & _
"    cmbStatusConsolidation.AddItem ""Pending""" & vbCrLf & _
"    cmbStatusCashApp.AddItem ""Completed""" & vbCrLf & _
"    cmbStatusCashApp.AddItem ""Pending Collector""" & vbCrLf & _
"End Sub"
End Function

' === Código del formulario: btnSubmit_Click ===
Function Formulario_SubmitCode() As String
Formulario_SubmitCode = _
"Private Sub btnSubmit_Click()" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    Dim nextRow As Long" & vbCrLf & _
"    Set ws = ActiveSheet" & vbCrLf & _
"    If ws Is Nothing Then MsgBox ""No hay hoja activa."": Exit Sub" & vbCrLf & _
"    nextRow = ws.Cells(ws.Rows.Count, ""A"").End(xlUp).Row + 1" & vbCrLf & _
"    With ws" & vbCrLf & _
"        .Cells(nextRow, 1).Value = txtInvoice.Value" & vbCrLf & _
"        .Cells(nextRow, 2).Value = txtStartDate.Value" & vbCrLf & _
"        .Cells(nextRow, 3).Formula = ""=IF(O"" & nextRow & ""=""""Completed"""",IF(C"" & nextRow & ""="""",NOW(),C"" & nextRow & ""),"""""")""" & vbCrLf & _
"        .Cells(nextRow, 4).Value = txtAmount.Value" & vbCrLf & _
"        .Cells(nextRow, 5).Value = txtACH.Value" & vbCrLf & _
"        .Cells(nextRow, 6).Value = txtSentTo.Value" & vbCrLf & _
"        .Cells(nextRow, 7).Value = cmbActionDone.Value" & vbCrLf & _
"        .Cells(nextRow, 8).Value = txtNote.Value" & vbCrLf & _
"        .Cells(nextRow, 9).Value = cmbSystem.Value" & vbCrLf & _
"        .Cells(nextRow, 10).Formula = ""=IF(I"" & nextRow & ""="""",Data!$I$2,VLOOKUP(I"" & nextRow & "",Data!$B:$C,2,FALSE))""" & vbCrLf & _
"        .Cells(nextRow, 11).Value = cmbCollector.Value" & vbCrLf & _
"        .Cells(nextRow, 12).Value = cmbCashAppOwner.Value" & vbCrLf & _
"        .Cells(nextRow, 13).Formula = ""=IF(I"" & nextRow & ""="""",$AL$2,VLOOKUP(I"" & nextRow & "",Data!$B:$E,4,FALSE))""" & vbCrLf & _
"        .Cells(nextRow, 14).Formula = ""=IF(D"" & nextRow & ""=""""", """"", IF(D"" & nextRow & ">=0, """"OA"""", """"SP""""))""" & vbCrLf & _
"        .Cells(nextRow, 15).Value = cmbStatusConsolidation.Value" & vbCrLf & _
"        .Cells(nextRow, 16).Value = cmbStatusCashApp.Value" & vbCrLf & _
"        .Cells(nextRow, 17).Value = txtFollowUp.Value" & vbCrLf & _
"        .Cells(nextRow, 18).Value = txtAnswer.Value" & vbCrLf & _
"        .Cells(nextRow, 19).Formula = ""=IF(P"" & nextRow & ""=""""Completed"""", """"✔"""", IF(B"" & nextRow & ""="""", """"""", TODAY()-S"" & nextRow & "))""" & vbCrLf & _
"        .Cells(nextRow, 19).NumberFormat = ""[Red]# """"dias pendientes"""";[Color10] # """"dias restantes"""";[Blue] """"hoy""""" & vbCrLf & _
"    End With" & vbCrLf & _
"    MsgBox ""Datos guardados en "" & ws.Name" & vbCrLf & _
"    Unload Me" & vbCrLf & _
"End Sub"
End Function

' === Código para crear hoja "Data" con ejemplo ===
Function CrearHojaDataCode() As String
CrearHojaDataCode = _
"Sub CrearHojaData()" & vbCrLf & _
"    Dim ws As Worksheet" & vbCrLf & _
"    On Error Resume Next" & vbCrLf & _
"    Set ws = Sheets(""Data"")" & vbCrLf & _
"    If ws Is Nothing Then" & vbCrLf & _
"        Set ws = Sheets.Add: ws.Name = ""Data""" & vbCrLf & _
"        ws.Range(""B1:E1"").Value = Array(""System"", ""Billing Owner"", ""Cash App Owner"", ""Cash App CC"")" & vbCrLf & _
"        ws.Range(""B2:E2"").Value = Array(""ACCOUNTABLE HEALTHCARE STAFFING"", ""Irene Medrano"", ""Masis Alvarado Alfredo"", ""Daisuke Hara / Laura Alfaro"")" & vbCrLf & _
"        ws.Range(""B3:E3"").Value = Array(""Adventist Health Adventist Health System/West - Aya Invoices"", ""Karina Zamora"", ""Montero Lobo Barbara"", ""Oscar Carballo"")" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub"
End Function
