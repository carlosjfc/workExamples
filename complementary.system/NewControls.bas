Attribute VB_Name = "NewControls"
Option Compare Database

Sub NewControles()
    Dim frm As Form
    Dim ctlLabel As control, ctlText As control
    Dim intDataX As Integer, intDataY As Integer
    Dim intLabelX As Integer, intLabelY As Integer

    ' Create new form with Orders table as its record source.
    Set frm = CreateForm
    frm.RecordSource = "StudentsCourses"
    ' Set positioning values for new controls.
    intLabelX = 100
    intLabelY = 100
    intDataX = 1000
    intDataY = 100
    ' Create unbound default-size text box in detail section.
    Set ctlText = CreateControl(frm.NAME, acTextBox, , "", "", _
        intDataX, intDataY)
    ' Create child label control for text box.
    Set ctlLabel = CreateControl(frm.NAME, acLabel, , _
         ctlText.NAME, "NewLabel", intLabelX, intLabelY)
    ' Restore form.
    DoCmd.Restore
End Sub

Sub CrearFormulario()
' **************
' Código de prueba
' eduardo@olaz.net
' Junio de 2002
' **************

Dim frm As Form
Dim strFormulario As String
Const conFilas As Long = 16
Const conColumnas As Long = 16
Const conAncho As Long = 400
Const conAlto As Long = 400
Const conSeparacion As Long = 50
Const conMargenX As Long = 40
Const conMargenY As Long = 40
Const conComilla As String = """"
Const conIncrementoColor As Long = 65536
Dim lngFila As Long
Dim lngColumna As Long
Dim lngX As Long
Dim lngY As Long
Dim ctlEtiqueta As control
Dim ctlLabel As control, ctlText As control
Dim ctlBoton As control
Dim intDataX As Integer, intDataY As Integer
Dim intLabelX As Integer, intLabelY As Integer
Dim colColor As Long
Dim aControles() As control
Dim strCodigoBotonSalir As String
Dim mdlFormulario As Module

ReDim aControles(conColumnas, conFilas)
' Crea el nuevo formulario
Set frm = CreateForm
colColor = 0 * conIncrementoColor \ 16
With frm
.RecordSelectors = False
.NavigationButtons = False
.Width = 7300
.ScrollBars = 0
.DividingLines = False
.MinMaxButtons = 0
End With

For lngFila = 0 To conFilas - 1
lngY = lngY + conAncho + conSeparacion
For lngColumna = 0 To conColumnas - 1
lngX = lngX + conAlto + conSeparacion
Set aControles(lngColumna, lngFila) = CreateControl(frm.NAME, acLabel, , "", "", _
lngX, lngY, conAncho, conAlto)
Set ctlEtiqueta = aControles(lngColumna, lngFila)
With ctlEtiqueta
.BackColor = colColor
.BackStyle = 1
End With
colColor = colColor + conIncrementoColor
Next lngColumna

lngX = 0
Next lngFila

Set ctlBoton = CreateControl(frm.NAME, acCommandButton, , "", "", 8000, 8000, 1500, 500)
ctlBoton.Caption = "Salir"
ctlBoton.NAME = "cmdSalir"

Set mdlFormulario = frm.Module
strCodigoBotonSalir = "Private sub " & ctlBoton.NAME & "_Click()" & _
vbCrLf & _
vbCrLf & _
"msgbox " & conComilla & "Cierro el formulario" & conComilla & _
" & me.name " & _
vbCrLf & _
" docmd.Close" & _
vbCrLf & _
"End Sub"
With mdlFormulario
.InsertText strCodigoBotonSalir
End With
' Restaura el formulario.
DoCmd.Restore
Erase aControles
Set ctlEtiqueta = Nothing
strFormulario = frm.NAME
DoCmd.Save acForm, strFormulario
DoCmd.OpenForm strFormulario

Debug.Print Forms(strFormulario).Width
Set frm = Nothing
End Sub


