param (
    [string]$arg1,
    [string]$arg2,
    [string]$arg3
)

# Definir el path del JAR
$jarPath = "D:/Java/print-sales-ticket/target/print-sales-ticket-0.0.1-SNAPSHOT.jar"
$jarArgs = "-jar $jarPath $arg1 $arg2 $arg3"

# Mostrar los argumentos y el comando para depuración
Write-Output "Ejecutando comando: java $jarArgs"

# Ejecutar el programa Java que genera el documento Word con argumentos
& java -jar $jarPath $arg1 $arg2 $arg3

#Crear un objeto de aplicación de Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Abrir el documento DOCX
$doc = $word.Documents.Open("D:\Java\print-sales-ticket\target\classes\Ticket.docx")

# Imprimir el documento
$doc.PrintOut()

# Cerrar el documento sin guardar cambios
$doc.Close([ref]$false)

# Salir de la aplicación de Word
$word.Quit()