

Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

Hide-Powershell

# creating window
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Service Desk Command Centre'
$main_form.Width = 450
$main_form.Height = 375
$main_form.AutoSize = $true
# $main_form.MaximizeBox = $false;
$main_form.StartPosition = 'CenterScreen'
$main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D

#starting text on top
$StartText = New-Object System.Windows.Forms.Label
$StartText.Text = "Please choose from the following prompts:"
$StartText.Location  = New-Object System.Drawing.Point(10,8)
$StartText.Font = [System.Drawing.Font]::new('Segoe UI', 10)
$StartText.AutoSize = $true
$main_form.Controls.Add($StartText)

$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(10,40)
$Button1.Size = New-Object System.Drawing.Size(200,70)
$Button1.Text = "Update User's Profile Pictures"
$Button1.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button1.backcolor = 'White'
$Button1.FlatStyle = 'Flat'
$main_form.Controls.Add($Button1)

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(220,40)
$Button2.Size = New-Object System.Drawing.Size(200,70)
$Button2.Text = "Update Calendar Delegate Access"
$Button2.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button2.backcolor = 'White'
$Button2.FlatStyle = 'Flat'
$main_form.Controls.Add($Button2)

$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(10,120)
$Button3.Size = New-Object System.Drawing.Size(200,70)
$Button3.Text = "Find Computer Status and Bitlocker Password"
$Button3.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button3.backcolor = 'White'
$Button3.FlatStyle = 'Flat'
$main_form.Controls.Add($Button3)

$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(220,120)
$Button4.Size = New-Object System.Drawing.Size(200,70)
$Button4.autosize = $true
$Button4.Text = "(to be coded)"
$Button4.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button4.backcolor = 'White'
$Button4.FlatStyle = 'Flat'
$main_form.Controls.Add($Button4)

$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(10,200)
$Button5.Size = New-Object System.Drawing.Size(200,70)
$Button5.Text = "(To be coded)"
$Button5.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button5.backcolor = 'White'
$Button5.FlatStyle = 'Flat'
$main_form.Controls.Add($Button5)

$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(220,200)
$Button6.Size = New-Object System.Drawing.Size(200,70)
$Button6.autosize = $true
$Button6.Text = "(To be coded)"
$Button6.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$Button6.backcolor = 'White'
$Button6.FlatStyle = 'Flat'
$main_form.Controls.Add($Button6)

$pictureBox1 = New-Object System.Windows.Forms.PictureBox
$pictureBox1.BackgroundImage = [System.Drawing.Image]::FromFile('.\Resources\GT-Mobius 40x40.png')
$pictureBox1.Location = New-Object System.Drawing.Size(10,280)
$pictureBox1.Size = New-Object System.Drawing.Size(40,40)
$pictureBox1.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$main_form.Controls.Add($pictureBox1)

$GTtext = New-Object System.Windows.Forms.Label
$GTtext.Text = "Command Center v1.0"
$GTtext.Location  = New-Object System.Drawing.Point(60,300)
$GTtext.AutoSize = $true
$GTtext.ForeColor = 'Gray'
$GTtext.Font = [System.Drawing.Font]::new('Segoe UI', 10, [System.Drawing.FontStyle]::Italic)
$main_form.Controls.Add($GTtext)



$Button1.Add_Click(
    {
    $Button1.Enabled = $false
    $Button2.Enabled = $false
    $Button3.Enabled = $false
    $Button4.Enabled = $false
    $Button5.Enabled = $false
    $Button6.Enabled = $false
    $Button1.Text = "Running..."
    $main_form.Refresh()
    start-process powershell.exe .\UpdateProfilePictures.ps1
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
    Start-Sleep 3
    $Button1.Enabled = $true
    $Button2.Enabled = $true
    $Button3.Enabled = $true
    $Button4.Enabled = $true
    $Button5.Enabled = $true
    $Button6.Enabled = $true
    $Button1.text = "Update User's Profile Pictures"
    $main_form.Refresh()
    }
)

$Button2.Add_Click(
    {
    $Button1.Enabled = $false
    $Button2.Enabled = $false
    $Button3.Enabled = $false
    $Button4.Enabled = $false
    $Button5.Enabled = $false
    $Button6.Enabled = $false
    $Button2.Text = "Running..."
    $main_form.Refresh()
    start-process powershell.exe .\UpdateDelegateAccess.ps1
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
    Start-Sleep 3
    $Button1.Enabled = $true
    $Button2.Enabled = $true
    $Button3.Enabled = $true
    $Button4.Enabled = $true
    $Button5.Enabled = $true
    $Button6.Enabled = $true
    $Button2.text = "Update Calendar Delegate Access"
    $main_form.Refresh()
    }
)

$Button3.Add_Click(
    {
    $Button1.Enabled = $false
    $Button2.Enabled = $false
    $Button3.Enabled = $false
    $Button4.Enabled = $false
    $Button5.Enabled = $false
    $Button6.Enabled = $false
    $Button3.Text = "Running..."
    $main_form.Refresh()
    start-process powershell.exe .\SearchComputerDetails.ps1 
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
    Start-Sleep 3
    $Button1.Enabled = $true
    $Button2.Enabled = $true
    $Button3.Enabled = $true
    $Button4.Enabled = $true
    $Button5.Enabled = $true
    $Button6.Enabled = $true
    $Button3.text = "Find Computer Bitlocker Recovery Password"
    $main_form.Refresh()
    }
)


$main_form.ShowDialog()
