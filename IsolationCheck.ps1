# External Dependencies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

# Project: Isolation Notifications
# Creation Date: 8-15-2022
# Contributor(s): Tony Jordan
# Desc: Present a large notification to user to notify them of upcoming isolation. 
# Make consequences of isolation apparent to user.

# First create function to allow for recursive scheduling of this script in task scheduler.
$arg = '-File ' + $MyInvocation.MyCommand.Path;
function recursiveSchedule() {
    # $time variable holds date as of the running of this script.
    $time = Get-Date;

    # Schedule script to run based on when the application is first launched/other times this script has been ran.
    if($time.Hours -gt 6) {
        if($time.Hours -le 18) {
            # Create a task timeline (Daily at 12 PM)
            $taskTrigger = New-ScheduledTaskTrigger -Daily -At 12PM
            $taskTrigger
        }
        else {
            # Create a task timeline (Daily at 12 PM)
            $taskTrigger = New-ScheduledTaskTrigger -Daily -At 12AM
            $taskTrigger
        }
    }
    else {
        # Create a task timeline (Daily at 12 PM)
        $taskTrigger = New-ScheduledTaskTrigger -Daily -At 12AM
        $taskTrigger
    }

    $taskName = "Isolation_Check"
    $desc = "Test to see if device is nearing isolation";

    $taskAction = New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument $arg
    $taskAction

    # Test to see if scheduled task already exists
    $taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -like $taskName}

    if(-not($taskExists)) {
        Register-ScheduledTask `
            -TaskName $taskName `
            -Action $taskAction `
            -Trigger $taskTrigger `
            -Description $desc `
            -Settings (New-ScheduledTaskSettingsSet -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries) `
    }

}

# Create task to re-run this script once a day
recursiveSchedule($e);

$global:resetMode = 1;
$tempv1 = Get-Date;
$tempv2 = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime;                      # Collect uptime information
$global:Uptime = $tempv1 - $tempv2;


$isoMessage = (quarantine /timeline)[2];                                     # Collect remediation/isolation information from windows.
$updates_needed = (quarantine /listupdates);                        # Collect info regarding which services/applications need to be updated.

$isoDate = ($isoMessage).SubString(($isoMessage).Indexof("IsolationDeadline: ") + ("IsolationDeadline: ").Length);       # Extract only isolation timeline information (for now)
$CurrentTime = (Get-Date).toString("yyyy/MM/dd HH:mm:ss");                                                                  # Extract current time information

# Function for comparing times to see if isolation is 3 or less days from taking effect.

function Compare-Time($timeleft) {
    
    $temp = $timeleft;
    if($temp.ToString().Contains("-")) {                                                                                                # Conditional that checks to see if isolation date has passed or not, and displays corresponding dialogue in console.
        DisplayCurrentIsoWindow($timeleft)
        Write-Output("Device has been isolated for: " + $temp.toString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'"));
    }
    else {
        if($timeleft.Days -lt 3) {
            Write-Output("Time until device is isolated: " + $temp.toString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'"));
            DisplayFutureIsoWindow($timeleft)
        }
    }

    return $timeleft;
}

function DisplayFutureIsoWindow($dateString) {
    
    # Create a new form
    $form = New-Object system.Windows.Forms.Form

    # Define the size, title and background color
    $form.ClientSize = '1000,1000'
    $form.WindowState = 'Maximized'
    $form.text = "Isolation ALERT"
    $form.BackColor = "Black"


    # Create conditional to test whether or not reset mode is turned on for this application
    if($resetMode -eq 0) {
        # Create Labels to inform user on important information.
        $TTI = New-Object system.Windows.Forms.Label;                                             #TTI = Time 'Till Isolation
        $TTI.text = "You have " + $dateString.Days + " day(s), " + $dateString.Hours + " hour(s), " + $dateString.Minutes + " minute(s), and " + $dateString.Seconds + " second(s) until your device is isolated. " +
            "Isolation will result in complete loss of network access, and, by extension, a major loss in productivity.`n`nIn order to remedy this, your laptop will restart in the allotted time below. Please save all content before this timer concludes. This restart cannot be avoided by closing this window. If your laptop is still not" +
            " in compliance after this restart, please escalate with local IT.";

        $timer = New-Object System.Windows.Forms.TextBox;
        $FONT = New-Object System.Drawing.Font("Arial", 80, [System.Drawing.FontStyle]::Bold);
        $timer.AutoSize = $False;
        $timer.font = $FONT;
        $timer.ForeColor = "#F61302";
        $timer.BackColor = "Black";
        $timer.text = "10:00";
        #$timer.Location = New-Object System.Drawing.Point(100, 100);
        $timer.Size = New-Object System.Drawing.Size(300, 200);
        $timer.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center;
        $timer.ReadOnly = $True;

        $timer.Dock = [System.Windows.Forms.DockStyle]::Bottom;
        $form.Controls.Add($timer);
    }
    else {
        # Create Labels to inform user on important information.
        $TTI = New-Object system.Windows.Forms.Label;
        $report = "";
        foreach ($line in $updates_needed) {
            if ($line.Contains("Update Name") ) {
                continue;
            }
            elseif ($line.Contains("---------")) {
                continue;
            }
            else {
                $report = $report + "`n`n" + $line;
            }
        }                                            
        $TTI.text = "You have " + $dateString.Days + " day(s), " + $dateString.Hours + " hour(s), " + $dateString.Minutes + " minute(s), and " + $dateString.Seconds + " second(s) until your device is isolated. " +
            "Isolation will result in complete loss of network access (which means you won't be able to access the internet at all), and by extension, a major loss in productivity.`n`nIn order to remedy this, your laptop will need to complete these update(s): `n" + $report + "Your device may need to restart in order to complete some of these updates, information regarding such will be " +
            "available in your ACME application. Many update issues occur due to devices not being restarted for several weeks. Your device has been up for:`n" + $Uptime.toString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'") + "`n`n If you have any questions, please direct them to your site's local IT department.";

    }
    $TTI.AutoSize = $False;
    #$TTI.Location = New-Object System.Drawing.Point(20, 115);
    $TTI.ForeColor = "#F61302";
    $TTI.BackColor = "Transparent";
    $FONT = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold);
    $TTI.font = $FONT;
    $TTI.TextAlign = [System.Drawing.ContentAlignment]::TopCenter;

    $TTI.Dock = [System.Windows.Forms.DockStyle]::Fill;

    $form.Controls.Add($TTI);

    # Display the form
    [void]$form.ShowDialog()
}

function DisplayCurrentIsoWindow($dateString) {
    # Create a new form
    $form = New-Object system.Windows.Forms.Form

    # Define the size, title and background color
    $form.ClientSize = '1000,1000'
    $form.WindowState = 'Maximized'
    $form.text = "Isolation ALERT"
    $form.SelectionColor = "#ffffff"
    $form.BackColor = "Black"

    # Create Labels to inform user on important information.
    $TTI = New-Object system.Windows.Forms.Label;                                             #TTI = Time 'Till Isolation
    $TTI.text = "Your device has been isolated for " + $dateString.Days + " day(s), " + $dateString.Hours + " hour(s), " + $dateString.Minutes + " minute(s), and " + $dateString.Seconds + " second(s). It currently has no network access. " +
    "Your device will need to be restarted in an attempt to remedy, however, local IT might need to take a closer look at your device to make it fully operational.";
    $TTI.AutoSize = $False;
    #$TTI.Location = New-Object System.Drawing.Point(20, 115);
    $TTI.ForeColor = "#F61302";
    $FONT = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold);
    $TTI.font = $FONT;
    $TTI.TextAlign = [System.Drawing.ContentAlignment]::TopCenter;
    $TTI.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $form.Controls.Add($TTI);



    # Display the form
    [void]$form.ShowDialog();
}

if($isoDate -ne "N/A"){                                                                               # Conditional for checking whether or not device is close to isolation.

    $timeleft = NEW-TIMESPAN -Start ([DateTime]::ParseExact($CurrentTime, "yyyy/MM/dd HH:mm:ss", $null)) -End ([DateTime]::ParseExact($isoDate, "yyyy-MM-dd HH:mm", $null))  # Create timespan between isolation date and current date to see how many days until isolation goes through.
    $tl = Compare-Time($timeleft);
}
else { 

    Write-Output("There is currently no timeline available for isolation. (Which means you're probably just fine)");
    #$timeleft = NEW-TIMESPAN -Start ([DateTime]::ParseExact($CurrentTime, "yyyy/MM/dd HH:mm:ss", $null)) -End ([DateTime]"09/23/2022 16:50");
    #`$tl = Compare-Time($timeleft);
}

