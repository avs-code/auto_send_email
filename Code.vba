Option Explicit
Public strExtension As String

'
'Autor: https://github.com/arv187
'v27.11.19.1233 (versión día.mes.año.horaMinutos)
'
'Para que este código funcione debes activar la referencia:
'Microsoft CDO for Windows 2000 en Herramientas > Referencia (cdosys.dll)
'Documentación de Collaboration Data Objects (CDO)
'https://msdn.microsoft.com/en-us/library/ms872853.aspx
'Más información sobre las librerias CDO:
'https://blogs.msdn.microsoft.com/webdav_101/2015/05/18/about-cdo-for-windows-2000-cdosys/
'https://blogs.itpro.es/exceleinfo/2018/03/27/enviar-emails-de-gmail-o-dominio-propio-desde-excel-usando-cdo-y-vba-sin-tener-un-cliente-de-correo-configurado/
'
'Solo si usas google como email Gmail, autorizar acceso:
'https://www.google.com/settings/security/lesssecureapps
'
'
'Celda D = DNI/NIE/Pasaporte, Celda E = email.
'Formacion-Acreditaciones = Nombre hoja con datos.
'strLocation = ruta final de archivo pdf
'Secfolder = ruta seleccionada a carpeta de archivos
'recuento = cuenta nº celdas con DNI
'i cuenta nº de veces que se ha hecho el For.

'Mostramos un mensaje en un cuadro de diálogo para advertir al usuario antes de continuar con el envío, si pulsa sí se ejecuta lo que hay tras la línea 29 vbYes.
Private Sub CommandButton1_Click()
Dim Respuesta As Integer
Respuesta = MsgBox("¿Has puesto la ruta a los archivos PDF y hecho la comprobación?", vbYesNo + vbExclamation)
Select Case Respuesta
    Case vbYes

'Variables y tipos de datos
Dim MiCorreo As CDO.Message
Dim Rango As Range, i As Long, cell As Range
Dim recuento As Integer
Dim username As String
username = f_username
Dim password As String
password = f_password
Dim htmlcuerpo As String
Dim htmlfirma As String
Dim firma2 As String
Dim strLocation As String
Dim Destinatario As String
Dim Correo As String
Dim si_cuestionario As String

'Definiendo contenido de firma2
firma2 = "<p><span style=""font-size: x-small;"">Este correo electrónico y, en su caso, cualquier fichero anexo al mismo, contiene información de carácter confidencial exclusivamente dirigida a su destinatario o destinatarios. Si no es Vd. el destinatario del mensaje, le ruego lo destruya sin hacer copia digital o física, comunicando al emisor por esta misma vía la recepción del presente mensaje. Gracias</span></p>"

'Si el campo cuestionario está vacio (longitud 0) no aparece ningún texto de cuestionario, Si tiene algo escrito se adjunta mensaje predefinido+lo escrito en casilla cuestionario.
If Len(cuestionario.Value) = 0 Then
si_cuestionario = ""
Else
si_cuestionario = "Le adjuntamos el enlace al cuestionario de valoración de la actividad formativa con el objetivo de poder mejorarla de cara a próximas ediciones. " & "<a href=" & cuestionario & " target=""_blank"">Cuestionario</a>"
End If

'Si la ruta hacía el archivo PDF está vacía (longitud 0), entonces ejecuta mensaje de advertencia. En caso contrario (Else) ejecuta el programa de envío.
If Len(f_ruta.Value) = 0 Then
    MsgBox "No ha escogido la ruta donde se encuentran los archivos PDF", vbExclamation, "Ruta no seleccionada"
    
    Else

    With Worksheets("Formacion-Acreditaciones") 'Con la hoja Formacion-Acreditaciones seleccionamos el rango E2 a toda la columna E y contamos las celdas desde la última hacía arriba que contengan celdas de tipo visible
    Set Rango = .Range("E2:E" & .Cells(.Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
    End With

    Application.ScreenUpdating = False 'Desactivamos la actualización de pantalla para hacer más rápido el proceso

'El programa realiza un bucle recorriendo cada celda de DNI, email, en el rango dado anteriormente, columna E.

    For Each cell In Rango
        i = cell.Row
        
'Si en DNI no encuentra ningún dato manda un mensaje y salta a la L141 (siguiente fila). En caso contrario comprueba el email, si parece valido configura el servidor de correo, Set MiCorreo, si no es valido L131 muestra mensaje de error.
      
            If cell.Offset(0, -1).Value = "" Then
                MsgBox "No se encuentra el DNI-NIE-Pasaporte de: " & vbCrLf & cell.Offset(0, -3).Value & ", " & cell.Offset(0, -4).Value & vbCrLf & vbCrLf & "Anótelo ahora y envíelo más tarde de forma manual con su gestor de correo.", vbExclamation, "Error DNI-NIE-Pasaporte"
                GoTo siguiente_fila

            Else


check_email:
                          
                        If cell.Value Like "?*@?*.?*" Then
                        
                            Set MiCorreo = New CDO.Message
                            '
                            On Error Resume Next
                            With MiCorreo.Configuration.Fields
                           .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'indica si se usa cifrado ssl
                           .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
                           .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "nombre" 'nombre del servidor smtp
                           .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 0 'puerto del servidor smtp
                           .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1=cdoSendUsingPickup 2=cdoSendUsingPort
                           .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = username 'variable que contiene el nombre de usuario de la cuenta de email
                           .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password 'variable que contiene el password de usuario de la cuenta de email
                           .Update
                            End With
                
                            'Elementos del correo
                            'Asunto
                            Dim Asunto As String
                            Asunto = f_asunto.Value
                            'nombre destinatario
                            Destinatario = cell.Offset(0, -4).Value & " " & cell.Offset(0, -3).Value
                            'correo destinatario
                            Correo = cell.Value
                            strLocation = f_ruta.Value & "\" & Worksheets("Formacion-Acreditaciones").Cells(cell.Row, "D").Value & strExtension
                
                            'Cuerpo del mensaje
                            '
                            htmlcuerpo = "<p>" & "Estimado/a " & Destinatario & "<br><br>" & "TEXTO: <br>" & Worksheets("Formacion-Acreditaciones").Cells(cell.Row, "F").Value & "<br><br>" & si_cuestionario & "<br><br>" & "Atentamente:"
                
                            htmlfirma = "<p>" & "FIRMA" & "<hr /></div>" & firma2
                            
                            
                            With MiCorreo 'En este bloque unimos los elementos del correo
                            .Subject = Asunto
                            .From = username
                            .To = Correo
                            '.CC = "correo@dominio.com" 'otros destinatarios que queremos tengan constancia de este correo electrónico.
                            .BCC = username '"otrocorreo@dominio.com" 'ocultar los destinatarios en los mensajes de correo electrónico.
                            '.TextBody = Para cuerpo formato texto plano.
                            .HTMLBody = htmlcuerpo & htmlfirma 'Para cuerpo formato html.
                
                            'Añadimos el elemento adjunto que está contenido en la variable strLocation.
                            .AddAttachment strLocation
                            End With
                
                            On Error GoTo 0
                            '
                            MiCorreo.Send
                
                            Set MiCorreo = Nothing
                                                    
                            With ThisWorkbook.Sheets("Formacion-Acreditaciones") 'Hacemos recuento para mostrar el total y el progreso del envío en un mensaje que se cierra automáticamente cada 1 segundo (indicado en cell.Value, 1,).
                            recuento = .Range("D2", .Range("D" & .Rows.Count).End(xlUp)).Rows.Count
                            CreateObject("wscript.shell").Popup "email nº " & i - 1 & "/" & recuento & " enviado a: " & cell.Value, 1, "Macro enviar acreditaciones", 64
                            End With
                            
                        Else

                            MsgBox "No se encuentra la dirección de Email correspondiente al DNI-NIE-Pasaporte: " & vbCrLf & cell.Offset(0, -1).Value & vbCrLf & vbCrLf & "Y por tanto no se ha enviado el correo al destinatario" & vbCrLf & cell.Offset(0, -3).Value & ", " & cell.Offset(0, -4).Value & vbCrLf & vbCrLf & "Anótelo ahora y envíelo más tarde de forma manual con su gestor de correo.", vbExclamation, "Error Emails"
                       
                        End If
                            
            End If

    Application.ScreenUpdating = True
        
siguiente_fila:

    Next cell

MsgBox "Tarea finalizada." & vbCrLf & vbCrLf & "Compruebe que no haya ningún correo devuelto en la bandeja de entrada de su cuenta del gestor de correo. (Motivo: Correo rechazado por el servidor)", vbInformation, "FIN"

End If

    Case vbNo
        'Desde aquí la ejecucción en caso de pulsar el botón No en el cuadro de diálogo que advierte al usuario antes de continuar.
End Select
End Sub
'
'Aquí se indica la ruta a los archivos PDF.
Private Sub CommandButton2_Click()
'
Dim Secfolder As String
With Application.FileDialog(msoFileDialogFolderPicker)
.Title = "Get folder"
.ButtonName = "Aceptar"
'Ruta por defecto, dejar vacío para que el programa compruebe si el usuario ha puesto o no una ruta.
.InitialFileName = ""
  If .Show = -1 Then
  'Si se escoge una carpeta y se hace clic en aceptar la ruta se guarda en la variable Secfolder.
  Secfolder = .SelectedItems(1)
  'Se muestra la ubicación de la carpeta que se ha escogido en la celda f_ruta.
  f_ruta.Value = Secfolder
  
Else
  'cancel clicked
  
End If
End With
End Sub
'
'En el formulario al pulsar el botón cerrar se descarga de memoria el formulario
Private Sub button_cerrar_Click()
Unload Enviar
End Sub
'
'COMPROBADOR, comprueba si se ha escogido una ruta, entonces recorre el rango de la columna *DNI a la busqueda de archivos *.report.pdf en la ruta indicada que correspondan con el DNI de cada fila, además comprueba la columna email.
Private Sub comprobar_pdf_Click()

Dim Secfolder As String
Dim Rango As Range, i As Long, cell As Range

If Len(f_ruta.Value) = 0 Then
    MsgBox "No ha escogido la ruta donde se encuentran los archivos PDF", vbExclamation, "Ruta no seleccionada"
    
    Else

    With Worksheets("Formacion-Acreditaciones")
    Set Rango = .Range("E2:E" & .Cells(.Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
    End With

    Application.ScreenUpdating = False
    Dim strLocation As String
    For Each cell In Rango
        i = cell.Row

    If cell.Value Like "?*@?*.?*" Then
    strLocation = f_ruta.Value & "\" & Worksheets("Formacion-Acreditaciones").Cells(cell.Row, "D").Value & strExtension

                    If Dir(strLocation) <> "" Then
                    GoTo next_comprobacion

                    Else
                        MsgBox "No se encuentra el archivo PDF adjunto en la ruta escogida correspondiente a: " & vbCrLf & cell.Offset(0, -1).Value & vbCrLf & cell.Offset(0, -3).Value & ", " & cell.Offset(0, -4).Value & vbCrLf & vbCrLf & "O la celda DNI-NIE-Pasaporte está vacía" & vbCrLf & vbCrLf & "Corríjalo y vuelva a comprobar.", vbExclamation, "Error Archivo adjunto"
                    End If
    Else
    MsgBox "No se encuentra la dirección de Email correspondiente al DNI-NIE-Pasaporte: " & vbCrLf & cell.Offset(0, -1).Value & vbCrLf & cell.Offset(0, -3).Value & ", " & cell.Offset(0, -4).Value & vbCrLf & vbCrLf & "Corríjalo y vuelva a comprobar.", vbExclamation, "Error Emails"

    End If
next_comprobacion:
    Next cell
    Application.ScreenUpdating = True

fin_comprobacion:
    MsgBox "Comprobación finalizada." & vbCrLf & vbCrLf & "Si han aparecido errores corríjalos y vuelva a comprobar", vbInformation, "Comprobación Archivo adjunto"

End If
End Sub

Sub CommandButton_ext_report_Click()
strExtension = ".report.pdf"
End Sub

Sub CommandButton_ext_pdf_Click()
strExtension = ".pdf"
End Sub

Sub CommandButton_ext_input_Click()
strExtension = InputBox("Introduce la extension del archivo (El texto detrás del DNI/NIE/Pasaporte) incluídos puntos y otros simbolos que pudiera contener, P.Ejemplo: Para 28856425.input.pdf debe escribir .input.pdf")
End Sub

'
'Cambia los caracteres introducidos por asteríscos
Private Sub f_password_Change()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label11_Click()

End Sub

'
Private Sub Label6_Click()

End Sub

Private Sub Label9_Click()

End Sub

'
'Botón para desvelar la contraseña
Private Sub ver_pass_Click()
MsgBox Me.f_password
End Sub
