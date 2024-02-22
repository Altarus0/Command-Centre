<#
======================================================================================
        AD Sync to Office 365 for Profile Picture Script

        Version 1.0

        This script was created so that you could manually choose
        a specific user to update, if their AD account was not
        showing their latest photo.


        Sources the string data from the ThumbnailPhoto attribute
        of the AD user and pushes the data onto Azure AD.

        Version 1.1
        - Added refresh on label3.text to show pending sign.
        - Added function to hide powershell
        - Updated name and design       
        
        UPDATE: CODE WILL BE UNUSABLE FROM MARCH 2024 ONWARDS. PLEASE REFER TO SITE:
        https://knowledge-junction.in/2023/10/11/microsoft-365-major-update-exchangepowershell-retirement-of-tenant-admin-cmdlets-to-get-set-and-remove-userphotos/
        https://learn.microsoft.com/en-us/microsoft-365/admin/add-users/change-user-profile-photos?view=o365-worldwide#manage-user-photos-in-microsoft-graph-powershell
        https://learn.microsoft.com/en-us/graph/api/profilephoto-update?view=graph-rest-1.0&tabs=powershell

======================================================================================
#>


Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# variables
$Logfile = "\\file.gt.local\users$\Adrian.Chow\Code\Logs\UpdateProfilePictures.log"

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

#Creating Window
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Username for Photo Update'
$main_form.Width = 405
$main_form.Height = 205
$main_form.AutoSize = $true
$main_form.MaximizeBox = $false;
$main_form.StartPosition = 'CenterScreen'
$main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Please choose from the AD users"
$Label.Location  = New-Object System.Drawing.Point(15,10)
$Label.AutoSize = $true
$Label.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($Label)

#dropdown box with AD users in alphabetical order
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Width = 360
$Users = get-aduser -filter * -Properties SamAccountName -SearchBase "OU=Win10,OU=Staff,DC=GT,DC=local" | Sort-Object SamAccountName 
Foreach ($User in $Users) {
    $ComboBox.Items.Add($User.SamAccountName) | out-null;
}
$ComboBox.Location  = New-Object System.Drawing.Point(15,40)
$ComboBox.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($ComboBox)

# text box below
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Results:"
$Label2.Location  = New-Object System.Drawing.Point(15,70)
$Label2.AutoSize = $true
$Label2.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($Label2)

$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = ""
$Label3.Location  = New-Object System.Drawing.Point(70,70)
$Label3.AutoSize = $true
$Label3.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$main_form.Controls.Add($Label3)

# update button
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(275,120)
$Button.Size = New-Object System.Drawing.Size(100,33)
$Button.Text = "Update"
$Button.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button.backcolor = 'White'
$Button.FlatStyle = 'Flat'
$main_form.Controls.Add($Button)


$Button.Add_Click( {
    $user = $ComboBox.selectedItem
    $Label3.Text = "Pending, This may take around 10 seconds..."
    $Button.Enabled = $false
    $Label3.Refresh();
    if ($User -ne $null) {
    Set-UserPhoto -Identity $user -PictureData (Get-ADUser $user -Properties thumbnailPhoto).thumbnailPhoto -Confirm:$false
    $Label3.Text = $comboBox.selectedItem + " has been successfully updated."
    WriteLog "The profile picture has been updated for"
    WriteLog $user
    }
    else {
        $Label3.text = "Please choose a user."
    }
    $Button.Enabled = $true
}
)

$main_form.ShowDialog()

