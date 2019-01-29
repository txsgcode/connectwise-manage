<#
    .Synopsis
        Checks ConnectWise time entries in a particular reporting period for common errors.

    .NOTES
        Version:        1.0
        Author:         Randy Lucas
        Creation Date:  1/28/2019
        Purpose/Change: Initial script development
    	
    .DESCRIPTION
        This script connects to the Connectwise database, which will typically be similar to cwwebapp_<company_name>
        (e.g. cwwebapp_tsg)

        *REQUIRED
        The user running the script will need at least read only access to the database and will need the 
        SQLSERVER powershell module. The easiest way to obtain the SQLSERVER module is to download and 
        install the SQL Server Management Studios (SSMS). This is a free download from Microsoft.

        Checks time entries in a particular reporting period for the following types of errors:
            - Possible overlapping time entries
            - Blank time entries
            - Entries not set to "Onsite" that are surrounded by "Travel To" and "Travel From" work type
            - Entries set to work type "Travel From" that are not set to billable type "No Charge"

    .PARAMETER Name
        Runs the report against a specified individual. Enter the ConnectWise username.
        (e.g. jdoe)

    .PARAMETER WeeksAgo
        Defines how far in the past the query will be performed. Results only for that pay period are returned.
        (e.g. Entering 1 would return results for last week, 2 would return results for the week prior to last week)

    .PARAMETER SQLServer
        Specify the name of the SQL Server instance (no need to include instance name if default instance).
        (e.g. CONNECTWISE or CONNECTWISE\INSTANCE)

    .PARAMETER Database
        Specify the name of the ConnectWise sQL Database to query against.
        (e.g. cwwebapp_mycompany)

    .EXAMPLE
        Example command to launch the script for all employees for the previous pay period.

        Get-TimeSheetErrors -SQLServer CONNECTWISE -Database cwwebapp_com -WeeksAgo 1

    .EXAMPLE
        Example command to launch the script for a single employee for the previous pay period.
        
        Get-TimeSheetErrors -SQLServer CONNECTWISE -Database cwwebapp_com -WeeksAgo 1 -Name jdoe

    .EXAMPLE
        Example to launch the script and output results to screen:
        
        Get-TimeSheetErrors -SQLServer <CONNECTWISE> -Database <cwwebapp_com> -WeeksAgo <1> -Name <jdoe> | | Select Name, EntryStart, EntryEnd, EntryHours, PrevDeduction, ActualHours, WorkType, Billable, Error | FT

    .EXAMPLE
        Example to launch the script, and export results to HTML file or CSV with specified parameters:

        $reportData = Get-TimeSheetErrors -SQLServer <CONNECTWISE> -Database <cwwebapp_com> -WeeksAgo <1>

        $Header = @"
            <Style>
                TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; font-family: sans-serif; font-size: 12px;}
                TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #336699; color: #ffffff; font-family: sans-serif; font-size: 12px;}
                TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black; font-family: sans-serif; font-size: 12px;}
            </Style>
        "@

        $reportData | ConvertTo-Html -Property Name, EntryStart, EntryEnd, EntryHours, PrevDeduction, ActualHours, WorkType, Billable, Error -Head $Header | Out-File D:\report.html
        $reportData | Select-Object Name, EntryStart, EntryEnd, EntryHours, PrevDeduction, ActualHours, WorkType, Billable, Error | Export-Csv D:\report.csv -NoTypeInformation

    .DISCLAIMER 
        THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
        RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>

function Get-TimeSheetErrors {

param(
    [Parameter(Mandatory=$true)]
    [string]
    $SQLServer,

    [Parameter(Mandatory=$true)]
    [string]
    $Database,

    [Parameter(Mandatory=$true)]
    [int16]
    $WeeksAgo = 1,

    [Parameter(Mandatory=$false)]
    [string]
    $Name = $null
)


# Import the SQLServer Powershell Module
Import-Module SqlServer


    function Get-ReportingPeriod {
    
        # This function queries the correct pay period information from Connectwise, based on the weeksAgo parameter.
        [CmdletBinding()]
        [OutputType([psobject])]
        param(
            [Parameter(Mandatory=$true)]
            [string]
            $SQLServer,

            [Parameter(Mandatory=$true)]
            [string]
            $Database,

            [Parameter(Mandatory=$false)]
            [int16]
            $weeksAgo = 0
        )

        # Get Time Entry Periods from ConnectWise Database not older than 2 years.
        $timePeriodInfo = @()
        $threeYearsAgo = (Get-Date (Get-Date).AddYears(-2) -Format "M/d/yyyy")
    
        $timePeriodScriptBlock = 
            "SELECT *
                FROM [cwwebapp_tsg].[dbo].[TE_Period]
                WHERE (Date_End > '$threeYearsAgo')"

        $timePeriodInfo = (Invoke-Sqlcmd -Query $timePeriodScriptBlock -ServerInstance $SQLServer)

        # This is the heart of the function, where we get the needed ConnectWise time period.
        $currentWeek = $null
        $currentYear = (Get-Date).year
        $currentTimePeriods = $timeperiodInfo | 
            Where-Object {($_.Date_End -ge (Get-Date 1`/1/$currentYear)) -and ($_.Date_Start -le (Get-Date 1/1/$($currentYear + 1)))} | 
            Where-Object {if (!(($_.period -eq 1) -and ((get-date $_.Date_Start -f MM/dd/yyyy) -like "12*$currentyear"))) {$_}}

            # Determines if the current date falls within a particular time period.
            foreach ($timeperiod in $currentTimePeriods) {      

                # Tests against the time period to see if the current date falls within that period.
                if (((Get-Date) -ge $timeperiod.Date_Start) -and ((Get-Date) -le (Get-Date $timeperiod.Date_End).AddDays(1))) {

                    $currentWeek = $timeperiod

                    if ($weeksago -ge 1) {
                        # Calculates numbers to subtract based on $weeksago variable.
                        $daysago = $weeksago * 7
                        $periodago = $weeksago

                        # Calculates previous time period start and end date.
                        $prevweekStartDate = Get-Date (get-date $currentWeek.Date_Start).AddDays(-$daysago) -Format MM/dd/yyyy
                        $prevweekEndDate = Get-Date (get-date $currentWeek.Date_End).AddDays(-$daysago) -Format MM/dd/yyyy
                        $prevweekPeriod = ($currentWeek.Period) - $periodago
 
                        # Builds a custom object to store the time period data.
                        $prevTimePeriod = New-Object -TypeName psobject -Property @{ 
                            StartDate = $prevweekStartDate
                            EndDate = $prevweekEndDate
                            Period = $prevweekPeriod
                        }
                    }
                
                    else {
                        # Builds a custom object to store the time period data.
                        $prevTimePeriod = New-Object -TypeName psobject -Property @{
                            StartDate = (Get-Date $currentWeek.Date_Start -f MM/dd/yyyy)
                            EndDate = (Get-Date $currentWeek.Date_End -f MM/dd/yyyy)
                            Period = $currentWeek.Period
                        }
                    }
                }

                # This else statement would typically only appear if the Time Period table for the current year was not populated in CW.
                elseif ($currentWeek -eq $null) {
                    $prevTimePeriod = "Week calculation error"
                }
            }

        # Output the previous time period information.
        $prevTimePeriod | select Period,StartDate,EndDate   

    }


    function Get-TimeEntryErrors {

    [CmdletBinding()]
    [OutputType([psobject])]

    param(
        [Parameter(Mandatory=$false)]
        [string]
        $Name
    )
        # This section begins the main queries and checks for the time entries.
        $timeResults = @()

        $startDate = $reportingPeriod.startDate
        $endDate = $reportingPeriod.endDate

            # Queries the database for time entries for the chosen time period.
            $timeScriptBlock = 
            "SELECT *
                FROM [cwwebapp_tsg].[dbo].[v_rpt_Time]
                    WHERE (Date_Start >= '$startDate')
                      AND (Date_Start <= '$endDate')
                    ORDER BY date_start, time_start"

            $memberTime = (Invoke-Sqlcmd -Query $timeScriptBlock -ServerInstance $SQLServer) | Sort-Object Member_ID, Date_Start, Time_Start

            if ($name) {$members = $name}
            else {$members = $memberTime | Select-Object -ExpandProperty Member_ID -Unique}

        # This section builds out the results list based on the various checks.

         foreach ($member in $members) { 

            $memberEntries = $memberTime | Where-Object {$_.Member_ID -eq $member}
            $prevEntry = $null
            $endPoint = "01-01-2000"

            foreach ($entry in $memberEntries) {

                $entryStartTrue = $entry.Time_Start_UTC.tolocaltime()
                $entryStart = Get-Date $entrystarttrue -Format "MM-dd-yyyy hh:mm tt"
                $entryEndTrue = $entry.Time_End_UTC.tolocaltime()
                $entryEnd = Get-Date $entryendtrue -Format "MM-dd-yyyy hh:mm tt"
                $entryHours = "{0:N2}" -f ((((Get-Date $entryend) - (Get-Date $entrystart)).TotalMinutes) / 60)
                $deduction = $entryhours - $entry.hours_actual

                $entryObj = New-Object -TypeName psobject -Property @{
                    Name = $entry.member_id;
                    EntryStart = $entrystart;
                    EntryEnd = $entryend;
                    EntryHours = $entryhours; 
                    Deduction = $deduction;
                    PrevDeduction = $preventry.Deduction; 
                    ActualHours = $entry.hours_actual;
                    Type = $entry.work_type;
                    Error = "-";
                    Notes = $entry.notes;
                    WorkType = $entry.work_type;
                    Billable = $entry.option_id;
                }    

                if (!($entryObj.Type -eq "Clock In/Out") -and $entryObj.ActualHours -ne "0.00") {

                    if ((Get-Date $entryobj.entrystart) -lt ($preventry.entryend)) {
            
                        if ($prevEntry.Deduction -eq $entryObj.ActualHours) {
                            $entryObj.Error = "Deducted"
                            $overlap = $false
                        }

                        Else {
                            $entryObj.Error = "Possible Overlap"
                            $overlap = $true
                            $timeResults += $entryObj
                        }
                    }

                    elseif (($overlap -ne $false) -and ((Get-Date $entryobj.entrystart) -lt (Get-Date $endpoint))) {
                        $entryObj.Error = "Possible Overlap"
                        $timeResults += $entryObj
                    }


                    if ($entryObj.notes -like $null) {             
                        $entryObj.Error = "Blank"                
                        $timeResults += $entryObj
                    }

                    if (($entryObj.WorkType -like "Travel From") -and ($entryobj.Billable -notlike "NC")) {
                        $entryObj.Error = "No Charge"
                        $timeResults += $entryObj
                    }

                    if (($prevEntry.WorkType -like "Travel To") -and ($entryObj.WorkType -notlike "Onsite")) {
                        $entryObj.Error = "Onsite"
                        $timeResults += $entryObj
                    }

                    if ((Get-Date $entryObj.EntryEnd) -gt ($endPoint)) { $endpoint = $entryobj.EntryEnd }

                    $prevEntry = $entryObj
                }
            }
        }
    
        $timeResults
    }


# Launches the function to get the reporting period.
$reportingPeriod = Get-ReportingPeriod -weeksAgo $weeksAgo -SQLServer $SqlServer -Database $Database

# Launches the function to get the time entries. If Name is null, all employees will be returned.
Get-TimeEntryErrors -Name $Name

}
