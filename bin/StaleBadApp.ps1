#Script date: 20-Oct-2023
#Author: Israel Rojas Valera

#region Execution Time
$TimeQuery = ConvertFrom-AFRelativeTime -RelativeTime "*"
$TimeQuery = $TimeQuery.ToLocalTime() 
#endregion

#region Directories
$CurrentBinPath = $PSScriptRoot
$CurrentLogPath = (get-item $CurrentBinPath).parent.FullName + "\log"
$CurrentLogFilePath = $CurrentLogPath + "\StaleBadApp_Log.csv"
$CurrentConfigPath = (get-item $CurrentBinPath).parent.FullName + "\config"
$CurrentOutputPath = (get-item $CurrentBinPath).parent.FullName + "\output"
$CurrentConfigFile = (get-item $CurrentBinPath).parent.FullName + "\config\parameters.xml"
#endregion

(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Inicio Ejecucion" | Out-File $CurrentLogFilePath -Append

#region Config read
[xml]$xmlConfig = Get-Content -Path $CurrentConfigFile
$XML_PIDAServer = $xmlConfig.config.PIDAConnection.PIDAServerName
$XML_PIDAServerUser = $xmlConfig.config.PIDAConnection.PIDAServerUser
$XML_PIDAServerUserPass = $xmlConfig.config.PIDAConnection.PIDAServerUserPass
$XML_AFServerName = $xmlConfig.config.AFConnection.AFServerName
$XML_AFServerUser = $xmlConfig.config.AFConnection.AFServerUser
$XML_AFServerUserPass = $xmlConfig.config.AFConnection.AFServerUserPass
$XML_AFDataBase = $xmlConfig.config.AFConnection.AFDataBase
$XML_AFRoot = $xmlConfig.config.AFConnection.AFRoot
$XML_AFResultTable = $xmlConfig.config.AFConnection.AFResultTable
$XML_StaleStartTime = $xmlConfig.config.StaleConfig.StaleStartTime
$XML_StaleEndTime = $xmlConfig.config.StaleConfig.StaleEndTime
$XML_StaleOutputTag = $xmlConfig.config.StaleConfig.OutputTag
$XML_TotalTags = $xmlConfig.config.OverallResults.TotalTags
$XML_TotalStaleTags = $xmlConfig.config.OverallResults.TotalStaleTags
$XML_TotalBadTags = $xmlConfig.config.OverallResults.TotalBadTags

#endregion

function CheckStaleBadTags{
	#Query DA Time
	$QueryDATime = ConvertFrom-AFRelativeTime -RelativeTime "*"
	$QueryDATime = $QueryDATime.ToLocalTime()
	
	#Global Variables
	$N_Tags = 0
	$N_Tags_Stale = 0
	$N_Tags_Bad = 0
    $CurrentOutputStaleFile = $CurrentOutputPath + "\StaleTags.csv"
    $CurrentOutputBadFile = $CurrentOutputPath + "\BadTags.csv"
    #Array Lists
    New-Object -TypeName System.Collections.ArrayList
    $arrlist_BadStatus = [System.Collections.ArrayList]@()
    $arrlist_BadStatus_OutputTags = [System.Collections.ArrayList]@()
    $arrlist_BadStatus_Count = [System.Collections.ArrayList]@()
    $arrlist_BadStatus_fromPI = [System.Collections.ArrayList]@()
    
    #Clean Output files
    $CleanOutputFilesPath = $CurrentOutputPath + "\*.csv"
    Remove-Item -Path $CleanOutputFilesPath -Force -Recurse

    #PI Data Archive Connection
	try{
		#Credential encryption
		$PIDA_secure_pass = ConvertTo-SecureString -String $XML_PIDAServerUserPass -AsPlainText -Force
		$PIDA_credentials = New-Object System.Management.Automation.PSCredential ($XML_PIDAServerUser, $PIDA_secure_pass)
		$PIDA_Connection = Connect-PIDataArchive $XML_PIDAServer -WindowsCredential $PIDA_credentials
	}catch { 
		$e = $_.Exception
		$msg = $e.Message
		while ($e.InnerException) {
			$e = $e.InnerException
			$msg += "`n" + $e.Message
			(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Error: " + $msg | Out-File $CurrentLogFilePath -Append
			}
	}
    #PI AF Connection
	try{
		#Credential encryption
		$AF_secure_pass = ConvertTo-SecureString -String $XML_AFServerUserPass -AsPlainText -Force
		$AF_credentials = New-Object System.Management.Automation.PSCredential ($XML_AFServerUser, $AF_secure_pass)
		$AFServer = Get-AFServer $XML_AFServerName
        $AF_Connection = Connect-AFServer -WindowsCredential $AF_credentials -AFServer $AFServer
        #Get AF Database $ Root Element
        $AFDB = Get-AFDatabase -Name $XML_AFDatabase -AFServer $AFServer
        $AFRootElement = Get-AFElement -AFDatabase $AFDB -Name "Results"
	}catch { 
		$e = $_.Exception
		$msg = $e.Message
		while ($e.InnerException) {
			$e = $e.InnerException
			$msg += "`n" + $e.Message
			(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Error: " + $msg | Out-File $CurrentLogFilePath -Append
			}
	}
    
    #AF Table Cleaning    
    $AF_Tables = $AFDB.Tables
    if ($AF_Tables["Stale Bad Results"].Table.Rows.Count -gt 0){
        $AF_Tables["Stale Bad Results"].Table.Rows.Clear()
    }

    # Bad Status list XML Reading
    $xmlConfig.Config.BadConfig.Status | ForEach-Object {
        #Read Status
        $XML_BadValue = $_.Value
        #Read Output Tag
        $XML_BadResultTag = $_.OutputTag
        $arrlist_BadStatus += $XML_BadValue
        $arrlist_BadStatus_OutputTags += $XML_BadResultTag
        $arrlist_BadStatus_Count += 0
    }

    #Stale Time Limits
    $TimeLimit_StartTime = ConvertFrom-AFRelativeTime -RelativeTime $XML_StaleStartTime
    $TimeLimit_StartTime = $TimeLimit_StartTime.ToLocalTime()
    $TimeLimit_EndTime = ConvertFrom-AFRelativeTime -RelativeTime $XML_StaleEndTime
    $TimeLimit_EndTime = $TimeLimit_EndTime.ToLocalTime()   


    #PI Tag check
    $PIPoints = Get-PIPoint -Name "*" -Connection $PIDA_Connection
    $LastTimeQuery = $TimeQuery

    $PIPoints | ForEach-Object {

        $CurrTime = Get-Date

        If ($LastTimeQuery -lt $CurrTime){
            $adv = [math]::round(100*$N_Tags/$PIPoints.Count,2)
            Write-Host $adv "%"
            $LastTimeQuery = $CurrTime.AddSeconds(5)
        }
        $N_Tags = $N_Tags + 1
        try{         
            $TagName = $_.Point.Name
            $PISnapshotValue = Get-PIValue -PIPoint $_ -Time (ConvertFrom-AFRelativeTime -RelativeTime "*") -ArchiveMode Previous
            $LastArcVal_TS = $PISnapshotValue.TimeStamp                    
            $LastArcVal_TS = $LastArcVal_TS.ToLocalTime()				
            $LastArcVal_Val = $PISnapshotValue.Value
            $LastArcVal_Status = [boolean]$PISnapshotValue.IsGood             
            #Stale

            $IsStale = $false
            $IsBad = $false
            $IsStaleBad = $false

            If ($LastArcVal_TS -gt $TimeLimit_StartTime -and $LastArcVal_TS -lt $TimeLimit_EndTime) {
                #Stale
                #Write-Host "Stale"
                $IsStale = $true
				If ($N_Tags_Stale -eq 0){
					"TagName,TimeStamp,Value" | Out-File $CurrentOutputStaleFile -Append
				}
                $N_Tags_Stale = $N_Tags_Stale + 1                    
                If ($LastArcVal_Status) {
                    $TagName + "," +  $LastArcVal_TS + "," + $LastArcVal_Val | Out-File $CurrentOutputStaleFile -Append
                }else{
                    #Bad Value
                    #Write-Host "Bad"
                    $IsBad = $true
                    $stateSetID = $LastArcVal_Val.StateSet
                    $stateID = $LastArcVal_Val.State
                    If ($stateSetID -eq 0){ #It is a System Digital State (Error)
                        If ($N_Tags_Bad -eq 0){
					        "TagName,TimeStamp,Value" | Out-File $CurrentOutputBadFile -Append
				        }
                        $N_Tags_Bad = $N_Tags_Bad + 1
                        $digitalState = Get-PIDigitalStateSet -ID $stateSetID -Connection $PIDA_Connection
                        $resultState = $digitalState[$stateID]
                    
                        #Add value to the Dig State Array
                        $arrlist_BadStatus_fromPI += $resultState
                        $TagName + "," +  $LastArcVal_TS + "," + $resultState | Out-File $CurrentOutputStaleFile -Append
                        $TagName + "," +  $LastArcVal_TS + "," + $resultState | Out-File $CurrentOutputBadFile -Append
                    }
                }
            }else{
                If ($LastArcVal_Status) {
                    #Do Nothing
                }
                else {

                    #Bad Value
                    #Write-Host "Bad"
                    $IsBad = $true
                    $stateSetID = $LastArcVal_Val.StateSet
                    $stateID = $LastArcVal_Val.State
                    If ($stateSetID -eq 0){ #It is a System Digital State (Error)
                        If ($N_Tags_Bad -eq 0){
					        "TagName,TimeStamp,Value" | Out-File $CurrentOutputBadFile -Append
				        }
                        $N_Tags_Bad = $N_Tags_Bad + 1
                        $digitalState = Get-PIDigitalStateSet -ID $stateSetID -Connection $PIDA_Connection
                        $resultState = $digitalState[$stateID]
                    
                        #Add value to the Dig State Array
                        $arrlist_BadStatus_fromPI += $resultState
                        $TagName + "," +  $LastArcVal_TS + "," + $resultState | Out-File $CurrentOutputBadFile -Append
                    }
                }
            }
            
            #Add value to the Results PI AF Table
            If ($IsStale){
                If ($IsBad){
                    $AF_Tables["Stale Bad Results"].Table.Rows.Add($XML_PIDAServer,$TagName,$resultState,$LastArcVal_TS,"X","X") | Out-Null
                }else{
                    $AF_Tables["Stale Bad Results"].Table.Rows.Add($XML_PIDAServer,$TagName,$LastArcVal_Val,$LastArcVal_TS,"X","") | Out-Null
                }
            }else{
                If ($IsBad){
                    $AF_Tables["Stale Bad Results"].Table.Rows.Add($XML_PIDAServer,$TagName,$resultState,$LastArcVal_TS,"","X") | Out-Null
                }
            }                    
        }
        catch { 
		$e = $_.Exception
		$msg = $e.Message
		while ($e.InnerException) {
			$e = $e.InnerException
			$msg += "`n" + $e.Message
			(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Error: " + $msg | Out-File $CurrentLogFilePath -Append
			}
	    }
    }

    #Check In AF Database (Table insertions)
    $AF_Checkin = New-AFCheckIn($AFDB)

    #Process Bad Tag List
    for ($i=0; $i -lt $arrlist_BadStatus_fromPI.Length; $i++){        
        for ($j=0; $j -lt $arrlist_BadStatus.Length; $j++) {
            if ($arrlist_BadStatus_fromPI[$i] -eq $arrlist_BadStatus[$j]){
                $arrlist_BadStatus_Count[$j] = $arrlist_BadStatus_Count[$j] + 1
            }
        }            
    } 

    #Write Results to PI
    try{
        for ($i=0; $i -lt $arrlist_BadStatus.Length; $i++){
            Add-PIValue -PointName $arrlist_BadStatus_OutputTags[$i] -Time $TimeQuery -Value $arrlist_BadStatus_Count[$i] -WriteMode Append -Connection $PIDA_Connection -Buffer
        }
        Add-PIValue -PointName $XML_TotalTags -Time $TimeQuery -Value $N_Tags -WriteMode Append -Connection $PIDA_Connection -Buffer
        Add-PIValue -PointName $XML_TotalStaleTags -Time $TimeQuery -Value $N_Tags_Stale -WriteMode Append -Connection $PIDA_Connection -Buffer
        Add-PIValue -PointName $XML_TotalBadTags -Time $TimeQuery -Value $N_Tags_Bad -WriteMode Append -Connection $PIDA_Connection -Buffer
    }
    catch { 
		$e = $_.Exception
		$msg = $e.Message
		while ($e.InnerException) {
			$e = $e.InnerException
			$msg += "`n" + $e.Message
			(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Error: " + $msg | Out-File $CurrentLogFilePath -Append
		}
	}

    #Write Results to PI AF    
    $AFDatabase = Get-AFDatabase -AFServer $AF_Connection -Name $XML_AFDataBase
    $RootAFElement = Get-AFElement -AFDatabase $AFDatabase -Name $XML_AFRoot
    $AFElement = Get-AFElement -AFElement $RootAFElement -Name $XML_PIDAServer
    $ChildAttribute_StaleOutputFile = Get-AFAttribute -AFElement $AFElement -Name "Stale Output File"
    $ChildAttribute_BadOutputFile = Get-AFAttribute -AFElement $AFElement -Name "Bad Output File"
    $ChildAttribute_LastCheck = Get-AFAttribute -AFElement $AFElement -Name "Last Check"   
    $AFOutputStaleFile = New-Object OSIsoft.AF.Asset.AFFile
    If (Test-Path $CurrentOutputStaleFile){
        $AFOutputStaleFile.upload($CurrentOutputStaleFile)
        $ChildAttribute_StaleOutputFile.SetValue($AFOutputStaleFile)
    }
    $AFOutputBadFile = New-Object OSIsoft.AF.Asset.AFFile
    If (Test-Path $CurrentOutputBadFile){
        $AFOutputBadFile.upload($CurrentOutputBadFile)
        $ChildAttribute_BadOutputFile.SetValue($AFOutputBadFile)
    }
    $LastCheckAF = New-Object OSIsoft.AF.Asset.AFValue
    $LastCheckAF.Value = $TimeQuery
    $LastCheckAF.Timestamp = $TimeQuery
    $ChildAttribute_LastCheck.SetValue($LastCheckAF)

    #Write Results to Log
    (Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Total Stale Tags: " + $N_Tags_Stale | Out-File $CurrentLogFilePath -Append
    (Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Total Bad Tags: " + $N_Tags_Bad | Out-File $CurrentLogFilePath -Append
	(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Termino Ejecucion" | Out-File $CurrentLogFilePath -Append

    #Disconnections
    $AF_Connection = Disconnect-AFServer -AFServer $AFServer
    $PIDA_Connection.Disconnect()
}

#Main
CheckStaleBadTags	