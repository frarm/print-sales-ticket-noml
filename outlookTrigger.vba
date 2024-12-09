Public Sub RunPowerShellScript(MyMail As MailItem)
    Dim scriptPath As String
    Dim param1 As String
    Dim param2 As String
    Dim param3 As String
    Dim shellCommand As String

    ' Ruta del script de PowerShell
    scriptPath = "D:\Java\print-sales-ticket-noml\print-sales-ticket-noml.ps1"

    ' Parámetros del script
    param1 = "valor1"
    param2 = "valor2"
    param3 = "valor3"

    ' Comando para ejecutar PowerShell con parámetros
    shellCommand = "powershell.exe -File """ & scriptPath & """ -arg1 """ & param1 & """ -arg2 """ & param2 & """ -arg3 """ & param3 & """"

    ' Ejecuta el comando
    Shell shellCommand, vbNormalFocus
End Sub

