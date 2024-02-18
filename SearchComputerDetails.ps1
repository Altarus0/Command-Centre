
<#
======================================================================================
        Search Computer Details

        Version 1.0

======================================================================================
#>

Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing


#hiding powershell

$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
Function Show-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
Function Hide-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}

# variables
$Logfile = "\\file.gt.local\users$\Adrian.Chow\Code\Logs\SearchComputerDetails.log"
[string] $permissionflags
$ComputerData = @{}
$ComputerData.log = @()
$ComputerData.ID = @()
$ComputerData.Password = @()


# logging function
function WriteLog
{
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
}

Hide-Powershell

# creating window
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Computer Details'
$main_form.Width = 540
$main_form.Height = 410
$main_form.AutoSize = $true
$main_form.MaximizeBox = $false;
$main_form.StartPosition = 'CenterScreen'
$main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

#starting text on top
$StartText = New-Object System.Windows.Forms.Label
$StartText.Text = "Please input a computer serial number:"
$StartText.Location  = New-Object System.Drawing.Point(10,8)
$StartText.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$StartText.AutoSize = $true
$main_form.Controls.Add($StartText)

# window text for computer input
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "Computer:"
$Label1.Location  = New-Object System.Drawing.Point(20,55)
$Label1.AutoSize = $true
$Label1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($Label1)

# text box object for computer input
$TextBox1 = New-Object System.Windows.Forms.Textbox
$TextBox1.Width = 410
$TextBox1.Location  = New-Object System.Drawing.Point(95,53)
$TextBox1.Multiline = $false
$TextBox1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($TextBox1)

# Results textbox below the credentials
$TextBoxResult = New-Object System.Windows.Forms.Textbox
$TextBoxResult.Text = "Please enter a serial number."
$textBoxResult.Multiline = $true 
$TextBoxResult.Location  = New-Object System.Drawing.Point(10,110)
$TextBoxResult.Size = New-Object System.Drawing.Size(500,90)
$TextBoxResult.ReadOnly = $true
$TextBoxResult.Font = [System.Drawing.Font]::new('Segoe UI', 10,[System.Drawing.FontStyle]::Italic)
$main_form.Controls.Add($TextBoxResult)

# search button
$SearchButton = New-Object System.Windows.Forms.Button
$SearchButton.Location = New-Object System.Drawing.Size(390,320)
$SearchButton.Size = New-Object System.Drawing.Size(120,33)
$SearchButton.Text = "Search"
$SearchButton.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$SearchButton.backcolor = 'White'
$SearchButton.FlatStyle = 'Flat'
$main_form.Controls.Add($SearchButton)



# Laebel for Combobox
$ComboBoxLabel = New-Object System.Windows.Forms.Label
$ComboBoxLabel.Width = 210
$ComboBoxLabel.Location  = New-Object System.Drawing.Point(10,210)
$ComboBoxLabel.text = "Bitlocker recovery timestamps:"
$ComboBoxLabel.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)


# Dropdown box with Times in alphabetical order
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Width = 290
$ComboBox.text = " -  -  -  -  -  -  -  -  -  -  -  -"
$ComboBox.Location  = New-Object System.Drawing.Point(220, 210)


# Bitlocker Recovery Details

$RecoveryPasswordData= New-Object System.Windows.Forms.Textbox
$RecoveryPasswordData.Width = 500
$RecoveryPasswordData.Height = 60
$RecoveryPasswordData.Location  = New-Object System.Drawing.Point(10,240)
$RecoveryPasswordData.text = $null
$RecoveryPasswordData.Multiline = $true 
$RecoveryPasswordData.ReadOnly = $true 
$RecoveryPasswordData.Font = [System.Drawing.Font]::new('Segoe UI', 10)


# Outline for Credentials
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Text = 'Credentials'
$groupBox1.Location = New-Object System.Drawing.Point(10,30)
$groupBox1.Size = New-Object System.Drawing.Size(510,60)
$groupBox1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($groupBox1)


$TextBox1.Add_KeyDown( {if ($PSItem.KeyCode -eq "Enter") {
        $SearchButton.PerformClick()
        }
    }
)

$SearchButton.Add_KeyDown( {if ($PSItem.KeyCode -eq "Enter") {
    $SearchButton.PerformClick()
    }
}
)

$SearchButton.Add_Click({
    $TextBoxResult.text = "Please wait, this may take up to 10 seconds to complete..."
    $SearchButton.Enabled = $false
    $main_form.Controls.Remove($ComboBoxLabel)
    $main_form.Controls.Remove($combobox)
    $main_form.Controls.Remove($RecoveryPasswordData)
    $comboBox.Items.Clear()
    $TextBoxResult.Refresh()
    $RecoveryPasswordData.text = "xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx"
    $ComboBox.text = " -  -  -  -  -  -  -  -  -  -  -  -"

    #Test to make sure email is existing
    $Computer = '*' + $TextBox1.text + '*'
    $testComputer = Get-ADComputer -Filter {name -like $Computer} -properties CanonicalName
    $ComputerLocation = $testComputer.CanonicalName
    $ComputerData.clear
    if ($testComputer.Enabled -eq $True) {
        $PCEnabled = 'Enabled'
        } else {
        $PCEnabled = 'Disabled'
     }

        
    # Populates Bitlocker details into an array
    if ($testComputer -ne $null) {

        $BitLockerDetails = Get-ADObject -Filter {(objectclass -eq 'msFVE-RecoveryInformation')} -Properties whenCreated, msFVE-RecoveryPassword | Where-Object {$_.DistinguishedName -like $Computer} | Sort-Object whenCreated -Descending

        foreach ($BitLockerDetail in $BitLockerDetails) {
            $computerDate = $BitLockerDetail.Name.Substring(0,10)
            $computerTime = $BitLockerDetail.Name.Substring(11,8)
            $ComputerPasswordID = $BitLockerDetail.Name.Substring(26,36)
            $ComputerRecoveryPassword = $BitLockerDetail.'msFVE-RecoveryPassword'
            $ComboBoxLog = "$computerDate" +" " + "$ComputerTime"
            $combobox.Items.Add($ComboBoxLog) 
            $ComputerData.log += "$ComboBoxLog"
            $ComputerData.ID += "$ComputerPasswordID"
            $ComputerData.Password += "$ComputerRecoveryPassword"

            }

        $TextBoxResult.text = "Search Complete, these are your search results:"
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("PC Name: ")
        $TextBoxResult.AppendText($Textbox1.text)
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("Location: $ComputerLocation")
        $TextBoxResult.AppendText("`r`n")
        $TextBoxResult.Appendtext("Status: $PCEnabled")
        $TextBoxResult.Refresh()
        $main_form.Controls.Add($ComboBoxLabel)
        $main_form.Controls.Add($ComboBox)
        $main_form.Controls.Add($RecoveryPasswordData)
        

        Add-Content $Logfile "-------------------------------------------"
        WriteLog $env:username
        WriteLog "Laptop SN searched:"
        WriteLog $textBox1.text
    }
    else {
        $data = $textbox1.text
        $TextBoxResult.Text = "$data is not valid computer serial number.Please check the text inputted."
        $TextBoxResult.Refresh()
        Add-Content $Logfile "-------------------------------------------"
        WriteLog $env:username
        WriteLog "Invalid SN searched:"
        WriteLog $textBox1.text
    }
        $SearchButton.Enabled = $true
        $main_form.Refresh()
})


$ComboBox.add_SelectedIndexChanged({
    $RecoveryPasswordData.text = $null
    $RecoveryPasswordData.Refresh()
    for ($i = 0; $i -lt $ComputerData.log.Count; $i++) {
        if($comboBox.text -eq $ComputerData.log[$i]) {
            $RecoveryPasswordData.Appendtext("Time: ")
            $RecoveryPasswordData.AppendText($ComputerData.log[$i])
            $RecoveryPasswordData.AppendText("`r`n")
            $RecoveryPasswordData.Appendtext("ID: ")
            $RecoveryPasswordData.AppendText($ComputerData.ID[$i])               
            $RecoveryPasswordData.AppendText("`r`n")
            $RecoveryPasswordData.AppendText("Password: ") 
            $RecoveryPasswordData.AppendText($ComputerData.Password[$i]) 
            break
        }
    }
})


$main_form.ShowDialog()

