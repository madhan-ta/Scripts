param(
    [string] $computername = $env:computername,
    [string] $logname = "system",
    [string] $newest = 1000,
    [parameter(Mandatory, HelpMessage = "enter the path name for the report.")]
    [string] $path
)
$data = Get-EventLog -LogName $logname -EntryType Error -Newest $newest -ComputerName $computername | Group-Object -Property source -NoElement

#html report
$title = "Eventlog Analysis"
$footer = "<h5> report run $(get-date)</h5>"
$data | Sort-Object -Property count, name -Descending | Select-Object count, name | ConvertTo-Html -Title $title -PreContent "<h1>Hostname=$env:computername <br/> $logname</h1>" -PostContent $footer | Out-File $path
Invoke-Item $path