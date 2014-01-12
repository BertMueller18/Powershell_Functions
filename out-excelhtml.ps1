# write HTML intro code:

$begin =

{

    '<table>'

    '<tr>'

    '<th>DisplayName</th><th>Status</th><th>Required</th><th>Dependent</th>'

    '</tr>'

}

# this is executed for each data object:

$process =

{

    if ($_.Status -eq 'Running')

    {

        $style = '<td style="color:green; ; font-family:Segoe UI; font-size:14pt">'

    }

    else

    {

        $style = '<td style="color:red; font-family:Segoe UI; font-size:14pt">'

    }

      '<tr>'

    '{0}{1}</td><td>{2}</td><td>{3}</td><td>{4}</td>' -f $style, $_.DisplayName, $_.Status, ($_.RequiredServices -join ','), ($_.DependentServices -join ',')

    '</tr>'

}

# finish HTML fragment:

$end =

{

    '</table>'

}

$Path = "$env:temp\tempfile.html"

# get all services and create custom HTML report:

Get-Service |

  ForEach-Object -Begin $begin -Process $process -End $end |

  Set-Content -Path $Path -Encoding UTF8

# feed HTML report into Excel:

Start-Process -FilePath 'C:\Program Files*\Microsoft Office\Office*\EXCEL.EXE' -ArgumentList $Path

