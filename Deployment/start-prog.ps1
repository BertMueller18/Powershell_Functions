 initializes the exe and assigns it to a variable
$Install = $([WMICLASS]"\\$server\ROOT\CIMV2:win32_process").Create("[path to .exe]")

# $checkproc continuously checks on the status of the process until it stops - 
be sure to use the -ErrorAction parameter to silence the output unless you want an error to pop 
when it finally sees the process isn't running anymore   

Do {
    $checkproc = Get-process -computer $Server -Id $Install.ProcessId -ErrorAction SilentlyContinue   
}
While ($checkproc -ne $null)

#more code here if you want - this will run after the do-while loop sees the .exe is no longer running

