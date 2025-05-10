param (
    [string]$actionType = "GP/NOGP",
    [string]$customerName,
    [string]$district,
    [string]$address,
    [string]$phone,
    [string]$date,
    [string]$productName,
    [string]$observation
)

# Definir el path del JAR
$jarPathImprimir = "D:\Java\print-sales-ticket-noml\target\print-sales-ticket-noml-0.0.1.jar"
$jarPathRegistrar = "D:\Java\register-gopack\target\register-gopack-0.0.1-SNAPSHOT.jar"
$jarArgs = "$productName $customerName $district"

# Mostrar los argumentos y el comando para depuración
Write-Output "Ejecutando comando: $jarArgs"

# Ejecutar el programa Java que genera el documento Word con argumentos
& java -jar $jarPathImprimir $productName $customerName $district

# Si actionType es "Y", ejecutar un programa adicional
if ($actionType -eq "GP") {
    Write-Output "Registrando producto en gopack..."
    & java -jar $jarPathRegistrar $customerName $district $address $phone $date $productName $observation
}

# Crear un objeto de aplicación de Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Abrir el documento DOCX
$doc = $word.Documents.Open("D:\Java\print-sales-ticket-noml\input-output\Ticket.docx")

# Imprimir el documento
$doc.PrintOut()

# Cerrar el documento sin guardar cambios
$doc.Close([ref]$false)

# Salir de la aplicación de Word
$word.Quit()