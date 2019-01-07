# Get report list to print
$outputs=(Get-ChildItem output\*.fodg)

foreach($invoice in $outputs){

    # Print to Default Printer
    Start-Process $invoice -Verb Print | Stop-Process
    # Print to your prefer printer
    #Start-Process soffice.exe -ArgumentList "-p $invoce.Fullname"
    # After print and Delete
    Remove-Item $invoice

}