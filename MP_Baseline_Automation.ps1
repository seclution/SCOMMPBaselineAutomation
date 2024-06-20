<#
.SYNOPSIS
    This script is used to override specific properties of a given SCOM management pack.

.DESCRIPTION
    Creates for "Disabled" a Monitor- or RuleOverride (Enabled = "false")
    Creates for "Debug" a Monitor- or RuleOverride and a second override for Debug Managementgroup (Enabled = "true" for GroupTarget *.DebugGrp)
    Creates for "Activated" a Monitor- or RuleOverride (Enabled = "true")

.PREREQUISITES
    - The OperationsManager module must be installed on the system.
    - Install OpsMgrExtended from https://github.com/tyconsulting/OpsMgrExtended-PS-Module/tree/master via "Install-Module -Name OpsMgrExtended".
    - Permissions to modify the SCOM environment.
    - Execute the script on your SCOM management server.

.FILES
    - Source MP as XML-File.
    - MP Viewer.
    - MPTool PS script.

.TO DO
    - Export your MP with MPViewer as an Excel XML to the same directory as the script with SourceMPName (sourceMP properties "ID", e.g., Microsoft.SQLServer.Windows.Monitoring.xml).
    - Ensure only one XML file exists.                                                                                                                                                         
    - Open the XML and modify the worksheets "Monitors - Unit" and "Rules" with "Enabled" = "Debug" or "Disabled".
    - Export again as Excel XML.
    - Run the MPTool script!
    - Check new overrides in the Operations Manager Console.

.FUNCTIONALITY
    - Imports the OperationsManager module.
    - Loads an XML file.
    - Searches for specific worksheets in the XML file.
    - Retrieves specific cells in the worksheets.
    - Creates new override objects.
    - Sets their properties.
    - Saves the changes to the override management pack.

.FLAGS

    -Scope      | Allowed Values: AddOnly and DeleteOnly
                | AddOnly: Only Add Overrides without deleting old overrides, autocreated by this Script
                | DeleteOnly: Only Delete old override, which are autocreated by this script
    -Debuglog   | Enables detailed debug information


.EXAMPLE
    .\Override-SCOMProperties.ps1
    This command runs the script with deleting old overrides and then create new ones.
    
    .\Override-SCOMProperties.ps1 -DebugLog -Scope AddOnly
    This command runs the script without deleting old overrides and with detailed Debugmessages.


.NOTES
    This script is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License.
    To view a copy of this license, visit https://creativecommons.org/licenses/by-sa/4.0/.

    You are free to:
    - Share: copy and redistribute the material in any medium or format for any purpose, even commercially.
    - Adapt: remix, transform, and build upon the material for any purpose, even commercially.

    Under the following conditions:
    - Attribution: You must give appropriate credit, provide a link to the license, and indicate if changes were made. You may do so in any reasonable manner, but not in any way that suggests the licensor endorses you or your use.
    - ShareAlike: If you remix, transform, or build upon the material, you must distribute your contributions under the same license as the original.

    For more information, please visit https://creativecommons.org/licenses/by-sa/4.0/.

.TOBECONTINUED
	- add remote mg connection

#>

# Bind the parameters for the script
[CmdletBinding()]
param(
    [ValidateSet("AddOnly", "DeleteOnly")]
    [string]$Scope,
    [switch]$DebugLog
)

# Import Moduls
Import-Module OpsMgrExtended
Import-Module OperationsManager


# Set Variables
$scriptName = $MyInvocation.MyCommand.Name
$toolName = [System.IO.Path]::GetFileNameWithoutExtension($scriptName)
$ErrorActionPreference = 'Continue'
$devCompany = "Seclution GmbH & Co. KG"
$comment = "Overriden by: $($env:USERNAME); `nOn: $(Get-Date) `nwith $($toolName) `nscripted by $($devCompany)"

Write-Host "`n Starting $toolName..." -BackgroundColor DarkGray

# Set the DebugPreference to Continue if the -Debug flag is used
if ($DebugLog) {
    $DebugPreference = 'Continue'
}


# Check existing XML Files
Try {
    # Get the directory of the currently executing script
    $scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    # Get all XML files in the script directory
    $xmlFiles = Get-ChildItem -Path $scriptDirectory -Filter *.xml

    # Check the number of XML files found
    if ($xmlFiles.Count -eq 1) {
        $file = $xmlFiles[0]
        $sourceManagementPackName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        Write-Host "`nFound XML file: $($file.FullName)"
        Write-Debug "Filename without extension: $sourceManagementPackName"
    } elseif ($xmlFiles.Count -gt 1) {
        Write-Host "Error: There should only be one XML file in the same folder as the script."
        Read-Host -Prompt "Drücken Sie die Eingabetaste, um das Fenster zu schließen"
        exit 1
    } else {
        Write-Host "No XML files found in the script directory."
        Read-Host -Prompt "Drücken Sie die Eingabetaste, um das Fenster zu schließen"
        exit 1
    }
} catch {
    throw "Error while processing XML files: $_"
}

    
try{    
# Mainaction

    # Create MP if not available
    try {
        Write-Host "`nCreate Managment Pack if not available..."
        $sourceMP = Get-SCOMManagementPack -Name $sourceManagementPackName
        $overrideMPName = "$($sourceManagementPackName).Override"
        $error.clear()        
        Write-Debug "Checking if MP $($overrideMPName) exists..."

        $OverrideManagementPackExists = Get-SCOMManagementPack -Name $overrideMPName 
        if ($OverrideManagementPackExists -eq $NULL)
        {
            $OverrideManagementPackDisplayName = "$($sourceMP.DisplayName) Override"
            $OverrideManagementPackDescription = "Override-MP $($sourceMP.Description)"
            Write-Debug "Creating new Managementpack with name $($OverrideManagementPackDisplayName)"
            $MG = Get-SCOMManagementGroup
            $MPStore = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackFileStore
            $MP = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPack($overrideMPName, $OverrideManagementPackDisplayName, (New-Object Version(1, 0, 0)), $MPStore)
            $MG.ImportManagementPack($MP)
            $MP.DisplayName = $OverrideManagementPackDisplayName
            $MP.Description = $OverrideManagementPackDescription
            $MP.Verify()
            $MP.AcceptChanges()
            $OverrideManagementPackExists = Get-SCOMManagementPack -Name $overrideMPName
            Write-Host "Created new Managementpack with name $($OverrideManagementPackDisplayName)"
        }
        else {
            Write-Debug "OverrideMP `"$($overrideMPName)`" already exists "
        }   
    }
    catch {
        throw "Error while creating MPOverride-Objekts: $_"
    }
 
    # Create Group if not available
    try {
        Write-Host "Create Override Debug Group if not available..."
        $DebugGrp = "$($sourceMP.Name).DebugGrp"
        if ($OverrideManagementPackExists -ne $NULL)
        {
            try { 
                $DebugGrpExists = Get-SCOMGroup -DisplayName $DebugGrp
                Write-Debug "Checking if Group $($DebugGrp) exists..."
                if ($DebugGrpExists -eq $NULL) {
                    $create = new-omcomputergroup -sdk "localhost" -MPname $overrideMPName -computergroupname $DebugGrp -computergroupdisplayname $DebugGrp
                    Write-Host "Created new Group with name $($DebugGrp)"
                }
                else
                {
                    Write-Debug "OverrideGrp `"$($DebugGrp)`" already exists " 
                }
            }
            catch {
                Write-Error "DebugGrp Exception"
            }
        }
        else {
            Write-Error "MP not found"
        }

    } catch {
        Write-Error "MP/DebugGrp Exception"
    }

    # If AddOnly Flag is not set, delete old overrides
    if ($Scope -ne "AddeOnly")
    {
        $overrides = Get-SCOMOverride | Where-Object { $_.GetManagementPack().Name -eq $overrideMPName }
        foreach ($override in $overrides) {
            if ($override.Name -like "AutoCreated.*"){
            Write-Host "Deleting $($override), please wait..."
            Remove-OMOverride -OverrideName $override.Name -SDK localhost
            }
        }
    }

    # If DeleteOnly Flag is not set, add new Overrides from XML
    if ($Scope -ne "DeleteOnly")
    {
        # Load XML-File
        try {

            # Load XML-File
            [xml]$xmlContent = Get-Content -Path $file.FullName
            Write-Debug "Loaded XML-File successfully"
            # Create NamespaceManager
            $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
            $nsManager.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")

        } catch{
            throw "Error while loading XML-File: $_"
        }

        $overrideMP = $OverrideManagementPackExists

        # Create Monitors    
        try {
            Write-Host "`n`nMonitors" -BackgroundColor DarkRed
            # Only for Worksheet "Monitors - Unit"
            $worksheet = $xmlContent.DocumentElement.SelectSingleNode("//ss:Worksheet[@ss:Name='Monitors - Unit']", $nsManager)


            if ($worksheet -ne $null) {
                Write-Debug "Found Worksheet: Monitors - Unit"
                $worksheet.SelectNodes("ss:Table/ss:Row", $nsManager) | ForEach-Object {
                    if ($_.SelectNodes("preceding-sibling::ss:Row", $nsManager).Count -eq 0)
                    {
                        foreach ($cell in $_.SelectNodes("ss:Cell", $nsManager)) {
                            if ($cell.SelectSingleNode("ss:Data", $nsManager).InnerText -eq "Enabled") {
                                $EnabledIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + 1)
                            }
                            if ($cell.SelectSingleNode("ss:Data", $nsManager).InnerText -eq "ObjectRef") {
                                $ObjRefIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + 1)
                            }
                        }
                    } else {
                        $EnabledValue = ""
                        $ObjRefValue = ""
                        $addindex = 1
				        foreach ($cell in $_.SelectNodes("ss:Cell", $nsManager)) {
					        # Get the column index (if specified)
					        $index = $cell.Attributes["ss:Index"]
					        $columnIndex = if ($index) { [int]$index.Value } else { $null }
                    
					        # Determine the actual column position
					        if ($columnIndex -eq $null) {
						        $columnIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + $addindex)
					        }
                            # Ignore empty rows
                            else
                            {
                                $columnIndexValue = $columnIndex
                                $columnIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + $addindex)
                                $addindex = $addindex + ($columnIndexValue - $columnIndex)
                                $columnIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + $addindex)
                            }

					        # Check if the cell is Enabled-Column or ObjRef-column
					        if ($columnIndex -eq $EnabledIndex) {
						        $EnabledValue = $cell.SelectSingleNode("ss:Data", $nsManager).InnerText
					        } elseif ($columnIndex -eq $ObjRefIndex) {
						        $ObjRefValue = $cell.SelectSingleNode("ss:Data", $nsManager).InnerText
					        }
				        }


                        if ($EnabledValue -eq "Debug") {
                            $parts = $ObjRefValue -split ";"
                            $DisabledMonitorName = $parts[1]
                            Write-Host "`nFound `"$($EnabledValue)`" for Monitor: $DisabledMonitorName" -BackgroundColor DarkGreen
                            Write-Debug "DisabledMonitorRef: $($ObjRefValue)"
                            $monitor = $sourceMP | Get-SCOMMonitor | Where-Object {$_.Name -eq $DisabledMonitorName}
                            Write-Debug "monitor: $($monitor)"
                            $targetClass = Get-ScomClass -Id $monitor.Target.Id
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName =  "AutoCreated." + $monitor.Name + ".Override"
                            Write-Debug "Creating Override for Monitor with Name: $($overrideName)  ..."
                            # Create a new override object in the appropriate management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackMonitorPropertyOverride($overrideMP, $overrideName)
                                $override.Monitor = $monitor
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "false"
                                Write-Host "Created Override with value: `"$($override.Value)`" for Monitor with Name: $($overrideName)"
                            } catch {
                                throw "Error while creating MonitorOverrideObjekt: $_"
                            }
					
                            Write-Debug "DebugMonitorRef: $($ObjRefValue)"
                            $parts = $ObjRefValue -split ";"
                            $DebugMonitorName = $parts[1]
                            Write-Debug "DebugMonitorName: $($DebugMonitorName)"                 
                            $monitor = $sourceMP | Get-SCOMMonitor | Where-Object {$_.Name -eq $DebugMonitorName}
                            Write-Debug "monitor: $($monitor)"
                            $targetClassInstance = Get-SCOMClassInstance -DisplayName $DebugGrp
                            Write-Debug "targetClassInstance: $($targetClassInstance)"
                            $targetClass = get-scomclass -id $targetClassInstance.MonitoringClassIds
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $monitor.Name + ".Debug.Override"
                            Write-Debug "Creating Debug Override for Monitor with Name: $($overrideName)"

                            # Create a new override object in the debug management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackMonitorPropertyOverride($overrideMP, $overrideName)
                                $override.Monitor = $monitor
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.ContextInstance = $targetClassInstance.Id
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "true"
                                Write-Host "Created Debug Override with Value: `"$($override.Value)`" for Monitor with Name: $($overrideName) and Group: $($DebugGrp)"
                            } catch {
                                throw "Error while creating DebugMonitorOverrideObjekt: $_"
                            }
                        }
                        elseif ($EnabledValue -eq "Disabled") {
                            $parts = $ObjRefValue -split ";"
                            $DisabledMonitorName = $parts[1]
                            Write-Host "`nFound `"$($EnabledValue)`" for Monitor: $DisabledMonitorName" -BackgroundColor DarkGreen
                            Write-Debug "DisabledMonitorRef: $($ObjRefValue)"
                            $monitor = $sourceMP | Get-SCOMMonitor | Where-Object {$_.Name -eq $DisabledMonitorName}
                            Write-Debug "monitor: $($monitor)"
                            $targetClass = Get-ScomClass -Id $monitor.Target.Id
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $monitor.Name + ".Override"
                            Write-Debug "Creating Override for Monitor with Name: $($overrideName)"
                            # Create a new override object in the appropriate management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackMonitorPropertyOverride($overrideMP, $overrideName)
                                $override.Monitor = $monitor
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "false"
                                Write-Host "Created Disabled Override with value: `"$($override.Value)`" for Monitor with Name: $($overrideName)"
                            } catch {
                                throw "Error while creating DisabledMonitorOverrideObjekt: $_"
                            }
                        }
                        elseif ($EnabledValue -eq "Enabled") {
                            $parts = $ObjRefValue -split ";"
                            $ActivatedMonitorName = $parts[1]
                            Write-Host "`nFound `"$($EnabledValue)`" for Monitor: $ActivatedMonitorName" -BackgroundColor DarkGreen
                            Write-Debug "DisabledMonitorRef: $($ObjRefValue)"
                            $monitor = $sourceMP | Get-SCOMMonitor | Where-Object {$_.Name -eq $ActivatedMonitorName}
                            Write-Debug "monitor: $($monitor)"
                            $targetClass = Get-ScomClass -Id $monitor.Target.Id
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $monitor.Name + ".Override"
                            Write-Debug "Creating Activated Override for Monitor with Name: $($overrideName)"
                            # Create a new override object in the appropriate management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackMonitorPropertyOverride($overrideMP, $overrideName)
                                $override.Monitor = $monitor
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "true"
                                Write-Host "Created Enabled Override with value: `"$($override.Value)`" for Monitor with Name: $($overrideName)"
                            } catch {
                                throw "Error while EnabledMonitorOverrideObjekt: $_"
                            }
                        }
                    }
                }
            }
            else {
                Write-Host "Can not find Worksheet 'Monitors - Unit'"
            }
        }catch{
            throw "Error while creating Override Montitors: $_"
        }

        # Create Rules
        try{
            Write-Host "`n`nRules" -BackgroundColor DarkRed
            # Only for Worksheet "Rules"
            $worksheet = $xmlContent.DocumentElement.SelectSingleNode("//ss:Worksheet[@ss:Name='Rules']", $nsManager)
        
            if ($worksheet -ne $null) {
                Write-Debug "Found Worksheet: Rules"
                $worksheet.SelectNodes("ss:Table/ss:Row", $nsManager) | ForEach-Object {
                    if ($_.SelectNodes("preceding-sibling::ss:Row", $nsManager).Count -eq 0)
                    {
                        foreach ($cell in $_.SelectNodes("ss:Cell", $nsManager)) {
                            if ($cell.SelectSingleNode("ss:Data", $nsManager).InnerText -eq "Enabled") {
                                $EnabledIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + 1)
                            }
                            if ($cell.SelectSingleNode("ss:Data", $nsManager).InnerText -eq "ObjectRef") {
                                $ObjRefIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + 1)
                            }
                        }
                    } else {
                        $EnabledValue = ""
                        $ObjRefValue = ""
                        $addindex = 1
				        foreach ($cell in $_.SelectNodes("ss:Cell", $nsManager)) {
					        # Get the column index (if specified)
					        $index = $cell.Attributes["ss:Index"]
					        $columnIndex = if ($index) { [int]$index.Value } else { $null }

					        # Determine the actual column position
					        if ($columnIndex -eq $null) {
						        $columnIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + $addindex)
					        }
                            else
                            {
                                $columnIndexValue = $columnIndex
                                $columnIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + $addindex)
                                $addindex = $addindex + ($columnIndexValue - $columnIndex)
                                $columnIndex = ($cell.SelectNodes("preceding-sibling::ss:Cell", $nsManager).Count + $addindex)
                            }

					        # Check if the cell is Enabled-Column or ObjRef-Column
					        if ($columnIndex -eq $EnabledIndex) {
						        $EnabledValue = $cell.SelectSingleNode("ss:Data", $nsManager).InnerText
					        } elseif ($columnIndex -eq $ObjRefIndex) {
						        $ObjRefValue = $cell.SelectSingleNode("ss:Data", $nsManager).InnerText
					        }
				        }


                       if ($EnabledValue -eq "Debug") {
                            $parts = $ObjRefValue -split ";"
                            $DisabledRulesName = $parts[1]
                            Write-Host "`nFound `"$($EnabledValue)`" for Rule: $DisabledRulesName" -BackgroundColor DarkGreen
                            Write-Debug "DisabledRulesRef: $($ObjRefValue)"
                            $rule = $sourceMP | Get-SCOMRule | Where-Object {$_.Name -eq $DisabledRulesName}
                            Write-Debug "rule: $($rule)"
                            $targetClass = Get-ScomClass -Id $rule.Target.Id
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $rule.Name + ".Override"
                            Write-Debug "Creating Override for Rule with Name: $($overrideName)"
                            # Create a new override object in the appropriate management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackRulePropertyOverride($overrideMP, $overrideName)
                                $override.Rule = $rule
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "false"
                                Write-Host "Created Override with value: `"$($override.Value)`" for Rule with Name: $($overrideName)"
                            } catch {
                                throw "Error while creating RuleOverrideObjekt: $_"
                            }
                  
                            Write-Debug "DebugRulesRef: $($ObjRefValue)"
                            $parts = $ObjRefValue -split ";"
                            $DebugRulesName = $parts[1]
                            Write-Debug "DebugRulesName: $($DebugRulesName)"
                            $rule = $sourceMP | Get-SCOMRule | Where-Object {$_.Name -eq $DebugRulesName}
                            Write-Debug "Rule: $($rule)"
                            $targetClassInstance = Get-SCOMClassInstance -DisplayName $DebugGrp
                            Write-Debug "targetClassInstance: $($targetClassInstance)"
                            $targetClass = get-scomclass -id $targetClassInstance.MonitoringClassIds
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $rule.Name + ".Debug.Override"
                            Write-Debug "Creating Debug Override for Rule with Name: $($overrideName)"

                            # Create a new override object in the debug management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackRulePropertyOverride($overrideMP, $overrideName)
                                $override.Rule = $rule
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.ContextInstance = $targetClassInstance.Id
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "true"
                                Write-Host "Created Debug Override with Value: `"$($override.Value)`" for Monitor with Name: $($overrideName) and Group: $($DebugGrp)"
                            } catch {
                                throw "Error while creating DebugRuleOverrideObjekt: $_"
                            }
                        }
                        elseif ($EnabledValue -eq "Disabled") {
                            $parts = $ObjRefValue -split ";"
                            $DisabledRulesName = $parts[1]
                            Write-Host "`nFound `"$($EnabledValue)`" for Rule: $DisabledRulesName" -BackgroundColor DarkGreen
                            Write-Debug "DisabledRulesRef: $($ObjRefValue)"
                            $rule = $sourceMP | Get-SCOMRule | Where-Object {$_.Name -eq $DisabledRulesName}
                            Write-Debug "rule: $($rule)"
                            $targetClass = Get-ScomClass -Id $rule.Target.Id
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $rule.Name + ".Override"
                            Write-Debug "Creating Override for Rule with Name: $($overrideName)"
                            # Create a new override object in the appropriate management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackRulePropertyOverride($overrideMP, $overrideName)
                                $override.Rule = $rule
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "false"
                                Write-Host "Created Disabled Override with value: `"$($override.Value)`" for Rule with Name: $($overrideName)"
                            } catch {
                                Write-Host "Error while creating DisabledRuleOverrideObjekt: $_"
                            }
                        }
                        elseif ($EnabledValue -eq "Enabled") {
                            $parts = $ObjRefValue -split ";"
                            $ActivatedRulesName = $parts[1]
                            Write-Host "`nFound `"$($EnabledValue)`" for Rule: $ActivatedRulesName" -BackgroundColor DarkGreen
                            Write-Debug "DisabledRulesRef: $($ObjRefValue)"
                            $rule = $sourceMP | Get-SCOMRule | Where-Object {$_.Name -eq $ActivatedRulesName}
                            Write-Debug "rule: $($rule)"
                            $targetClass = Get-ScomClass -Id $rule.Target.Id
                            Write-Debug "targetClass: $($targetClass)"
                            $overrideName = "AutoCreated." + $rule.Name + ".Override"
                            Write-Debug "Creating Override for Rule with Name: $($overrideName)"
                            # Create a new override object in the appropriate management pack
                            try {
                                $override = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPackRulePropertyOverride($overrideMP, $overrideName)
                                $override.Rule = $rule
                                $override.Property = 'Enabled'
                                $override.Context = $targetClass
                                $override.Description = $comment
                                $override.DisplayName = $overrideName
                                $override.Value = "true"
                                Write-Host "Created Disabled Override with value: `"$($override.Value)`" for Rule with Name: $($overrideName)"
                            } catch {
                                Write-Host "Error while creating EnabledRuleOverrideObjekt: $_"
                            }
                        }
                    }
                }
            } 
            else {
                Write-Host "Can not find Worksheet 'Monitors - Unit'"
            } 
        }catch{
            throw "Error while creating Override Rules: $_"
        }

		# Verify and safe MP
        try {
            # Verify the override management packs
            Write-Host "`nVerifying override management packs..."
            $overrideMP.Verify()
    
            # Save the changes to the override and debug management packs
            Write-Host "Saving changes to override management packs..."
            $overrideMP.AcceptChanges()
 
            Write-Host "`nScript completed successfully!" -BackgroundColor DarkGray
        } catch {
                Write-Host "Error while verifying or safing changes: $_"
        }
    }
}catch {
	Write-Host "Error: $_"
}
    

# Keep Session open
Read-Host -Prompt "`nPress Enter to close"
