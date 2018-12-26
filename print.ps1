# Get report list to print
$outputs=(Get-ChildItem output\*.fodg)

foreach($invoice in $outputs){

    # Print to Default Printer
    Start-Process $invoice -Verb Print | Stop-Process
    # After print and Delete
    Remove-Item $invoice

}