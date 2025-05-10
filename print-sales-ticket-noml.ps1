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

$jarPathImprimir = "D:\Java\print-sales-ticket-noml\target\print-sales-ticket-noml-0.0.1.jar"
$jarPathRegistrar = "D:\Java\register-gopack\target\register-gopack-0.0.1-SNAPSHOT.jar"
$jarArgs = "$productName $customerName $district"

Write-Output "Ejecutando comando: $jarArgs"

& java -jar $jarPathImprimir $productName $customerName $district

if ($actionType -eq "GP") {
    Write-Output "Registrando producto en gopack..."
    & java -jar $jarPathRegistrar $customerName $district $address $phone $date $productName $observation
}

# No se puede imprimir en word sin que sea pdf
$word = New-Object -ComObject Word.Application
$word.Visible = $false

$doc = $word.Documents.Open("D:\Java\print-sales-ticket-noml\input-output\Ticket.docx")
$doc.PrintOut()
$doc.Close([ref]$false)
$word.Quit()