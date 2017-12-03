function textbox ($purpose){
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "Data Entry Form"
    $form.Size = New-Object System.Drawing.Size(300,200) 
    $form.StartPosition = "CenterScreen"

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20) 
    $label.Size = New-Object System.Drawing.Size(280,20) 
    $label.Text = "Please enter the " + $purpose + " here:"
    $form.Controls.Add($label) 

    $textBox = New-Object System.Windows.Forms.TextBox 
    $textBox.Location = New-Object System.Drawing.Point(10,40) 
    $textBox.Size = New-Object System.Drawing.Size(260,20) 
    $form.Controls.Add($textBox) 

    $form.Topmost = $True

    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $x = $textBox.Text
        $x
    }
}

$url = textbox ("course's URL")
$user = textbox ("site's username")
$password = textbox ("site's password")


function Find-Folders {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = "C:\"
    $browse.ShowNewFolderButton = $false
    $browse.Description = "Select the target directory"

    $loop = $true
    while($loop)
    {
        if ($browse.ShowDialog() -eq "OK")
        {
        $loop = $false
		
		#Insert your script here
		
        } else
        {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                #Ends script
                return
            }
        }
    } 
    $browse.SelectedPath
    $browse.Dispose()
} 

$directory= Find-Folders

cd $directory
youtube-dl -u $user -p $password $url


$list=get-childitem -Path $directory -filter "*.mp4"| sort CreationTime |Select-Object -ExpandProperty name
$i=1;
foreach ($file in $list)
{
    Rename-Item -Path $directory\$file -NewName $i'.'$file
    $i++;
}
get-childitem -Path $directory -filter "*"| sort CreationTime |Select-Object -ExpandProperty name
