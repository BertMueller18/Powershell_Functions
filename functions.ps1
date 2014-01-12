function Remove-Diacritics([string]$String)
{
    $objD = $String.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
 
    for ($i = 0; $i -lt $objD.Length; $i++) {
        $c = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($objD[$i])
        if($c -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
          [void]$sb.Append($objD[$i])
        }
      }
 
    return("$sb".Normalize([Text.NormalizationForm]::FormC))
}

function Change-Umlaut{
	[CmdletBinding()]
	[OutputType([System.String])]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Inputstring
	)
	try {
    $Inputstring = $Inputstring.Replace('ä', 'ae').Replace('ö', 'oe').Replace('ü','ue').Replace('Ä', 'Ae').Replace('Ö', 'Oe').Replace('Ü', 'Ue').Replace('ß', 'ss').Replace('-','').Replace(' ','').Replace('.','').Replace('é','e').Replace('á','a').Replace('Ê','E')		
	$InputString = Remove-Diacritics($Inputstring)
    return $Inputstring
    
	}
	catch {
		throw
	}
}