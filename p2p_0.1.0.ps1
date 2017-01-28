$title = "PS offline To Arranged Directory"
$message = "Do you want to create a copy of the videos or move the videos to a new location?"

$Move = New-Object System.Management.Automation.Host.ChoiceDescription "&Move", `
    "Deletes all the files in the folder."

$Copy = New-Object System.Management.Automation.Host.ChoiceDescription "&Copy", `
    "Retains all the files in the folder."


$options = [System.Management.Automation.Host.ChoiceDescription[]]($Move, $Copy)

$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

$sourcePath = "$env:USERPROFILE\AppData\Roaming\Pluralsight\videos\"
do
{    
$items = Get-ChildItem -Path $sourcePath | sort LastWriteTime | Select LastWriteTime
 
# enumerate the items array
foreach ($item in $items)
{
            Write-Host $item.LastWriteTime
}

Write-Host "Look at the time up in the table and select the first and last files"
Write-Host "If you will write the date only, the time will be 00:00:00"

$date= read-host "Please enter date & time of the fisrt file downloaded on this course (ie: 'MM/DD/YY 09:00:00'):"

$date = $date -as [datetime]

if (!$date) { "Not A valid date and time"}

} while ($date -isnot [datetime])

$date

$begin = $date

do
{    
$date= read-host "Please enter date & time of the last file downloaded on this course (ie: 'MM/DD/YY 09:00:59'):"

$date = $date -as [datetime]

if (!$date) { "Not A valid date and time"}

} while ($date -isnot [datetime])

$date

$end = $date

#End of part 1

$pathsList = @()

$files = Get-ChildItem $sourcePath -Recurse -Filter "*.mp4" |Where-Object { $_.LastWriteTime -gt $begin -and $_.LastWriteTime -lt $end}| sort LastWriteTime |select FullName   
foreach ($item in $files)
    {
     $pathsList += $item.FullName
    }

#End of part 2
Write-Host "Please select the csv file which maps the videos and directories"

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$selection = Get-FileName
$csvFile = Import-Csv $selection

# End of part 3

$regex = '(\d+m\s\d+s)'
$videosCount = $csvFile.videos.Length
$directoryCount = $csvFile.directory.Length
if (($videosCount -ne $pathsList.Length) -or ($videosCount -ne $directoryCount))
{
    [System.Windows.Forms.MessageBox]::Show("there are " + $pathsList.Length + " videos and " + $videosCount + " video names in the csv
    and " + $directoryCount + " directories(same or different) int the csv file. Please fix it and try again later")
    exit
}

Write-Host "Please Select The Directory For The New Files"
Function Select-FolderDialog
{
    param([string]$Description="Select Folder",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

$destDir = Select-FolderDialog

$directories = @($csvFile.directory[0])

for ($i = 1; $i -le $csvFile.directory.length - 1; $i++)
{
    if ($csvFile.directory[$i] -ne $directories[-1])
    {
        $directories += $csvFile.directory[$i]
    }
}

foreach ($newFolder in $directories)
{
   
   New-Item "$destDir\$newFolder" -ItemType directory
}
Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)]
    [String]$Name
  )

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($Name -replace $re)
}


switch ($result)
    {
        0 
        {
            $serialNum = 0
            $lastDir = $csvFile.directory[0]
            for ($i=0; $i -le $csvFile.videos.Length - 1; $i++)
            {
                if ($csvFile.videos[$i] -match $regex)
                {
                    $video = $csvFile.videos[$i] -replace $regex,""
                }
                if ($csvFile.directory[$i] -eq $lastDir)
                {
                    $serialNum++
                }
                    else
                    {
                        $serialNum=1
                        $lastDir = $csvFile.directory[$i]
                    }
                    $newdir = $csvFile.directory[$i]
                $video = Remove-InvalidFileNameChars $video
                Move-Item $pathsList[$i] "$destDir\$newdir\$serialNum. $video.mp4"
                     
            }
          }
    
    
        1 
        {
            "You selected Copy."
            $serialNum = 0
            $lastDir = $csvFile.directory[0]
            for ($i=0; $i -le $csvFile.videos.Length - 1; $i++)
            {
                if ($csvFile.videos[$i] -match $regex)
                {
                    $video = $csvFile.videos[$i] -replace $regex,""
                }
                if ($csvFile.directory[$i] -eq $lastDir)
                {
                    $serialNum++
                }
                    else
                    {
                        $serialNum=1
                        $lastDir = $csvFile.directory[$i]
                    }
                    $newdir = $csvFile.directory[$i]
                $video = Remove-InvalidFileNameChars $video
                Copy-Item $pathsList[$i] "$destDir\$newdir\$serialNum. $video.mp4"
            }
           
        }
    }

    $dirs = Get-ChildItem -Path $destDir | sort CreationTime |select Name 
    $i = 0
    foreach ($dir in $dirs)
    {
        $i++
        $newname = $dir.Name
        Rename-Item "$destDir\$newname" -NewName "$destDir\$i. $newname"
        # Write-Host "$destDir\$newname" -NewName "$destDir\$i. $newname"
    }
    
    Write-Host "Done. Press any key to continue"

    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")