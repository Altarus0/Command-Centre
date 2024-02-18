
<#
======================================================================================
        Update Delegate Access 

        Version 1.0
        Allows for updates to the delegate access via inputting the
        two user's information into the selected slots. Will be able 
        to work for any calendar mailbox that exists on Exchange
        online.

        Version 1.1
        - Added refresh on label3.text to show pending sign.
        - Added function to hide powershell
        - Updated name and design

        Version 1.2
        - Changed Colour Scheme
        - Added Disabled button fuction to avoid double clicking.

        Version 1.3
        -Added "View" button to allow for system to query the calendar delegates for the user they inputted.
        - Opens up a dialog box viewing the details.
        - Updated text format and sizing of winodw.
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
$Logfile = "\\file.gt.local\users$\Adrian.Chow\Code\Logs\UpdateDelegateAccess.log"
[string] $permissionflags
[string] $User = ""

# logging function
function WriteLog
{
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $LogFile -value $LogMessage
}

Connect-ExchangeOnline
Hide-Powershell

# creating window
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='User Delegation'
$main_form.Width = 450
$main_form.Height = 300
$main_form.AutoSize = $true
$main_form.MaximizeBox = $false;
$main_form.StartPosition = 'CenterScreen'
$main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

#starting text on top
$StartText = New-Object System.Windows.Forms.Label
$StartText.Text = "Please enter the users you would like to delegate access to."
$StartText.Location  = New-Object System.Drawing.Point(10,8)
$StartText.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$StartText.AutoSize = $true
$main_form.Controls.Add($StartText)

# window text for user input
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "User:"
$Label1.Location  = New-Object System.Drawing.Point(20,55)
$Label1.AutoSize = $true
$Label1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($Label1)

# text box object for user input
$TextBox1 = New-Object System.Windows.Forms.Textbox
$TextBox1.Width = 310
$TextBox1.Location  = New-Object System.Drawing.Point(95,53)
$TextBox1.Multiline = $false
$TextBox1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($TextBox1)

# window text for delegate input
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Delegate:"
$Label2.Location  = New-Object System.Drawing.Point(20,87)
$Label2.AutoSize = $true
$Label2.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($Label2)

# text box object for delegate input
$TextBox2 = New-Object System.Windows.Forms.Textbox
$TextBox2.Width = 310
$Textbox2.Location  = New-Object System.Drawing.Point(95,85)
$TextBox2.Multiline = $false
$TextBox2.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($TextBox2)


# checkbox for invite forwarding
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Location = New-Object System.Drawing.Point(10,130)
$checkBox1.Size = New-Object System.Drawing.Size(170,20)
$checkBox1.Text = "Can View Private Items?"
$checkBox1.CheckAlign = 'MiddleLeft'
$checkBox1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($checkBox1)

# window text below the text box
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = "Please enter a user."
$Label3.Location  = New-Object System.Drawing.Point(10,155)
$Label3.AutoSize = $true
$Label3.Font = [System.Drawing.Font]::new('Segoe UI', 10,[System.Drawing.FontStyle]::Italic)
$main_form.Controls.Add($Label3)


# remove button
$RemoveButton = New-Object System.Windows.Forms.Button
$RemoveButton.Location = New-Object System.Drawing.Size(80,210)
$RemoveButton.Size = New-Object System.Drawing.Size(100,33)
$RemoveButton.Text = "Remove"
$RemoveButton.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$RemoveButton.backcolor = 'White'
$RemoveButton.FlatStyle = 'Flat'
$main_form.Controls.Add($RemoveButton)


# view button
$ViewButton = New-Object System.Windows.Forms.Button
$ViewButton.Location = New-Object System.Drawing.Size(195,210)
$ViewButton.Size = New-Object System.Drawing.Size(100,33)
$ViewButton.Text = "View"
$ViewButton.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$ViewButton.backcolor = 'White'
$ViewButton.FlatStyle = 'Flat'
$main_form.Controls.Add($ViewButton)


# update button
$UpdateButton = New-Object System.Windows.Forms.Button
$UpdateButton.Location = New-Object System.Drawing.Size(310,210)
$UpdateButton.Size = New-Object System.Drawing.Size(100,33)
$UpdateButton.Text = "Update"
$UpdateButton.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$UpdateButton.backcolor = 'White'
$UpdateButton.FlatStyle = 'Flat'
$main_form.Controls.Add($UpdateButton)


# Outline for Credentials
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Text = 'Credentials'
$groupBox1.Location = New-Object System.Drawing.Point(10,30)
$groupBox1.Size = New-Object System.Drawing.Size(410,93)
$groupBox1.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($groupBox1)


# Dialog Box for View
$dialogBox = New-Object System.Windows.Forms.Form
$dialogBox.Size = New-Object System.Drawing.Size(850,380)
$dialogBox.Name = 'Results'
$dialogBox.AutoSize = $true
$dialogBox.MaximizeBox = $false
$dialogBox.StartPosition = 'CenterScreen'
$dialogBox.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

#Label for dialog box
$DiaLogLabel = New-Object System.Windows.Forms.Label
$DiaLogLabel.Text = 'Current Delegate Access'
$DiaLogLabel.Location  = New-Object System.Drawing.Point(15,5)
$DiaLogLabel.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$DiaLogLabel.AutoSize = $true
$dialogBox.Controls.Add($DiaLogLabel)

# Dialog text in box
$DialogText = New-Object System.Windows.Forms.RichTextBox
$DialogText.Multiline = $true
$DialogText.WordWrap = $false
$DialogText.ScrollBars = 'Vertical'
$DialogText.ReadOnly = 'True'
$DialogText.Location = New-Object System.Drawing.Size(15,30)
$DialogText.Size = New-Object System.Drawing.Size(800,250)
$DialogText.Font = [System.Drawing.Font]::new('Consolas', 9)
$dialogBox.Controls.Add($DialogText)

#OK button in box
$DialogOKButton = New-Object System.Windows.Forms.Button
$DialogOKButton.Location = New-Object System.Drawing.Size(710,290)
$DialogOKButton.Size = New-Object System.Drawing.Size(100,33)
$DialogOKButton.Text = "OK"
$DialogOKButton.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$DialogOKButton.backcolor = 'White'
$DialogOKButton.FlatStyle = 'Flat'
$dialogBox.Controls.Add($DialogOKButton)


$RemoveButton.Add_Click(
    {
        $Label3.Text = "Pending..."
        $RemoveButton.Enabled = $false
        $ViewButton.Enabled = $false
        $UpdateButton.Enabled = $false
        $Label3.Refresh()

        #Test to make sure email is existing
        $User = $textBox1.text + ":\Calendar"
        $Delegate = $textBox2.text
        $testUser = Get-MailboxFolderPermission -Identity $User -erroraction silentlycontinue
        $testDelegate = Get-MailboxFolderPermission -Identity $Delegate -erroraction silentlycontinue

        if (($testUser -ne $null) -and ($testDelegate -ne $null)) {

            Remove-MailboxFolderPermission -Identity $user -User $delegate -confirm:$false
            $Label3.Text = "Delegate access has been successfully removed."
            $Label3.Refresh()
            Add-Content $Logfile "-------------------------------------------"
            WriteLog $env:username
            WriteLog "Removal of access had been successfully executed for:"
            WriteLog $textBox1.text
            WriteLog $textBox2.text
        }
        else {
            $Label3.Text = "Please input valid users."
            $Label3.Refresh()
            Add-Content $Logfile "-------------------------------------------"
            WriteLog $env:username
            WriteLog "Invalid Data inputted:"
            WriteLog $textBox1.text
            WriteLog $textBox2.text
        }

        $RemoveButton.Enabled = $true
        $ViewButton.Enabled = $true
        $UpdateButton.Enabled = $true
    }
)


$ViewButton.Add_Click(
{
    $Label3.Text = "Pending..."
    $RemoveButton.Enabled = $false
    $ViewButton.Enabled = $false
    $UpdateButton.Enabled = $false
    $Label3.Refresh()
    
    #Test to make sure email is existing
    $User = $textBox1.text + ":\Calendar"
    $testUser = Get-MailboxFolderPermission -Identity $User 
    
    if ($testUser -ne $null) {
        $DialogText.text = $testUser | Out-String
        $DialogText.Refresh()
        $Label3.Text = "Results have been retrieved."
        $Label3.Refresh()
        $dialogBox.ShowDialog()    
        Add-Content $Logfile "-------------------------------------------"
        WriteLog $env:username
        WriteLog "Retrieved data for:"
        WriteLog $textBox1.text
    }
    else {
        $Label3.Text = "Please input valid users."
        $Label3.Refresh()
        Add-Content $Logfile "-------------------------------------------"
        WriteLog $env:username
        WriteLog "Invalid Data inputted:"
        WriteLog $textBox1.text
    }
    $RemoveButton.Enabled = $true
    $ViewButton.Enabled = $true
    $UpdateButton.Enabled = $true
}
)

$DialogOKButton.Add_Click( 
{
    $dialogBox.Close()
}
)

$UpdateButton.Add_Click(
{
    $Label3.Text = "Pending, please wait for 5 seconds..."
    $RemoveButton.Enabled = $false
    $ViewButton.Enabled = $false
    $UpdateButton.Enabled = $false
    $Label3.Refresh()

    #Test to make sure email is existing
    $User = $textBox1.text + ":\Calendar"
    $Delegate = $textBox2.text
    $testUser = Get-MailboxFolderPermission -Identity $User 
    $testDelegate = Get-MailboxFolderPermission -Identity $Delegate 

    if (($testUser -ne $null) -and ($testDelegate -ne $null)) {

        #checkbox check
        if ($checkBox1.Checked) {
            Add-MailboxFolderPermission -Identity $user -User $delegate -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems -erroraction silentlycontinue
            Set-MailboxFolderPermission -Identity $user -User $delegate -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems -erroraction silentlycontinue
        }
        else {
            Add-MailboxFolderPermission -Identity $user -User $delegate -AccessRights Editor -erroraction silentlycontinue
            Set-MailboxFolderPermission -Identity $user -User $delegate -AccessRights Editor -erroraction silentlycontinue
        }

        $Label3.Text = "Delegate access has been successfully completed."
        $Label3.Refresh()
        Add-Content $Logfile "-------------------------------------------"
        WriteLog $env:username
        WriteLog "The Delegate access had been successfully executed for:"
        WriteLog $textBox1.text
        WriteLog $textBox2.text
    }
    else {
        $Label3.Text = "Please input valid users."
        $Label3.Refresh()
        Add-Content $Logfile "-------------------------------------------"
        WriteLog $env:username
        WriteLog "Invalid Data inputted:"
        WriteLog $textBox1.text
        WriteLog $textBox2.text
    }
        $RemoveButton.Enabled = $true
        $ViewButton.Enabled = $true
        $UpdateButton.Enabled = $true
    
}
)

$main_form.ShowDialog()
