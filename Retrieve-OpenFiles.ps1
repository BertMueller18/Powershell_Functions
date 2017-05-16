
    param ([string]$server='localhost',[string[]]$keywords,[switch]$refresh,[switch]$group)


    # Refresh the data else search in existing data


    if (!($collection) -or $refresh){

        $netfile = [ADSI]"WinNT://$server/LanmanServer"
        $allopenfiles = $netfile.Invoke("Resources") 

        $global:collection = @()
        $allopenfiles | foreach {

            # Using try/catch because of translation errors in the InvokeMember methods, they can be ignored

            try{
                $global:collection += New-Object PsObject -Property @{
                    Id = $_.GetType().InvokeMember("Name", ‘GetProperty’, $null, $_, $null)
                    itemPath = $_.GetType().InvokeMember("Path", ‘GetProperty’, $null, $_, $null)
                    UserName = $_.GetType().InvokeMember("User", ‘GetProperty’, $null, $_, $null)
                    LockCount = $_.GetType().InvokeMember("LockCount", ‘GetProperty’, $null, $_, $null)
                    Server = $server
                } # psobject
            }catch{
            } # end catch

        } # end foreach

    } # end if lastrun check


    $list = @()

    $keywords | foreach {

        [string]$keyword = $_
        $keywordmatch = $collection | ? { $_.itemPath -match $keyword} 

        if (!$group){

            $list = $keywordmatch | sort itempath | ft username,itempath -AutoSize

        } else {

            $keywordgroup = $keywordmatch | group username | sort count -des 
            $keywordcount = 0
            $keywordgroup | % {$keywordcount++}

            $list += "`n$keywordcount users found with file paths containing keyword $keyword`n"
            $list += $keywordgroup | ft name,count -auto

        } # end if group
   
    } # end foreach keyword
  
    $list


### EXAMPLES
########### 

# Retrieve-OpenFiles -keywords @('user','application','somewordinpath') -server 'fileserver2' -retrieve -group

# Retrieves files with multiple keywords in an array for server fileserver2, specify the group switch to group on username with
# the open files count per user. Specify refresh to fetch the realtime open files data from the fileserver otherwise cached data is used.
