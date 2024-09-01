# Load assemblies
Add-Type -AssemblyName System.Windows.Forms

function GetMainMenuItems {
    param(
        $getMainMenuItemsArgs
    )

    $menuItem1 = New-Object Playnite.SDK.Plugins.ScriptMainMenuItem
    $menuItem1.Description = "Export selected games"
    $menuItem1.FunctionName = "SunshineExport"
    $menuItem1.MenuSection = "@Sunshine App Export"
    
    return $menuItem1
}

function SunshineExport {
    param(
        $scriptMainMenuItemActionArgs
    )

    $windowCreationOptions = New-Object Playnite.SDK.WindowCreationOptions
    $windowCreationOptions.ShowMinimizeButton = $false
    $windowCreationOptions.ShowMaximizeButton = $false
    
    $window = $PlayniteApi.Dialogs.CreateWindow($windowCreationOptions)
    $window.Title = "Sunshine App Export (RefNull Fork)"
    $window.SizeToContent = "WidthAndHeight"
    $window.ResizeMode = "NoResize"
    
    # Set content of a window. Can be loaded from xaml, loaded from UserControl or created from code behind
    [xml]$xaml = @"
<UserControl
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

    <UserControl.Resources>
        <Style TargetType="TextBlock" BasedOn="{StaticResource BaseTextBlockStyle}" />
    </UserControl.Resources>

    <StackPanel Margin="16,16,16,16">
        <!-- Sunshine apps.json Path -->
        <TextBlock Text="Enter your Sunshine apps.json path" 
        Margin="0,0,0,8" 
        VerticalAlignment="Center"/>

        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBox x:Name="SunshinePath"
            Grid.Column="0"
            Margin="0,0,8,0"
            Width="320"
            Height="36"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Top"/>

            <Button x:Name="BrowseButtonPath"
            Grid.Column="1"
            Width="72"
            Height="36"
            HorizontalAlignment="Right"
            Content="Browse"/>
        </Grid>

        <!-- Sunshine PREP DO Script -->
        <TextBlock Text="[OPTIONAL] Enter your DO Script path" 
        Margin="0,16,0,8" 
        VerticalAlignment="Center"/>

        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBox x:Name="SunshineDoPath"
            Grid.Column="0"
            Margin="0,0,8,0"
            Width="320"
            Height="36"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Top"/>

            <Button x:Name="BrowseButtonDo"
            Grid.Column="1"
            Width="72"
            Height="36"
            HorizontalAlignment="Right"
            Content="Browse"/>
        </Grid>

        <!-- Sunshine PREP-CMD UNDO Script -->
        <TextBlock Text="[OPTIONAL] Enter your UNDO Script path" 
        Margin="0,16,0,8" 
        VerticalAlignment="Center"/>

        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBox x:Name="SunshineUndoPath"
            Grid.Column="0"
            Margin="0,0,8,0"
            Width="320"
            Height="36"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Top"/>

            <Button x:Name="BrowseButtonUndo"
            Grid.Column="1"
            Width="72"
            Height="36"
            HorizontalAlignment="Right"
            Content="Browse"/>
        </Grid>

        <!-- Finished Message and Export Button -->
        <TextBlock x:Name="FinishedMessage" 
        Margin="0,0,0,24" 
        VerticalAlignment="Center"
        Visibility="Collapsed"/>

        <Button x:Name="OKButton"
        IsDefault="true"
        Height= "36"
        Margin="0,5,0,0"
        HorizontalAlignment="Center"
        Content="Export Games"/>
    </StackPanel>
</UserControl>
"@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window.Content = [Windows.Markup.XamlReader]::Load($reader)

    $sunshinePath = "$Env:ProgramW6432\Sunshine\config\apps.json"    
    $inputFieldSunshineJSON = $window.Content.FindName("SunshinePath")
    # Set the initial path if the user has not provided a path
    if (-not $inputFieldSunshineJSON.Text) {
        $inputFieldSunshineJSON.Text = $sunshinePath
    }

    $doPath = ""    
    $inputFieldSunshineDO = $window.Content.FindName("SunshineDoPath")
    # Set the initial path if the user has not provided a path
    if (-not $inputFieldSunshineDO.Text) {
        $inputFieldSunshineDO.Text = $doPath
    }

    $undoPath = ""    
    $inputFieldSunshineUNDO = $window.Content.FindName("SunshineUndoPath")
    # Set the initial path if the user has not provided a path
    if (-not $inputFieldSunshineUNDO.Text) {
        $inputFieldSunshineUNDO.Text = $undoPath
    }
    
    # Event handler - BrowseButtonPath
    $browseButton = $window.Content.FindName("BrowseButtonPath")
    $browseButton.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = Split-Path $inputFieldSunshineJSON.Text -Parent
            $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            $openFileDialog.Title = "Select apps.json"
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $inputFieldSunshineJSON.Text = $openFileDialog.FileName
            }
        })

    # Event handler - BrowseButtonDo
    $browseButton = $window.Content.FindName("BrowseButtonDo")
    $browseButton.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = Split-Path $inputFieldSunshineJSON.Text -Parent
            $openFileDialog.Filter = "All files (*.*)|*.*"
            $openFileDialog.Title = "Select DO Script"
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $inputFieldSunshineDO.Text = $openFileDialog.FileName
            }
        })

    # Event handler - BrowseButtonUndo
    $browseButton = $window.Content.FindName("BrowseButtonUndo")
    $browseButton.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = Split-Path $inputFieldSunshineJSON.Text -Parent
            $openFileDialog.Filter = "All files (*.*)|*.*"
            $openFileDialog.Title = "Select UNDO Script"
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $inputFieldSunshineUNDO.Text = $openFileDialog.FileName
            }
        })

    # Attach a click event handler
    $button = $window.Content.FindName("OKButton")
    $button.Add_Click({
            if ($button.Content -eq "Dismiss") {
                $window.Close()
            }
            else {
                $sunshinePath = $inputFieldSunshineJSON.Text
                $sunshinePath = $sunshinePath -replace '"', ''
                $doPath = $inputFieldSunshineDO.Text
                $doPath = $doPath -replace '"', ''
                $undoPath = $inputFieldSunshineUNDO.Text
                $undoPath = $undoPath -replace '"', ''
                
                $count = DoWork($sunshinePath)
                $button.Content = "Dismiss"

                $finishedMessage = $window.Content.FindName("FinishedMessage")
                $finishedMessage.Text = ("Created {0} Sunshine app shortcuts" -f $count)
                $finishedMessage.Visibility = "Visible"
            }
        })
    
    # Set owner if you need to create modal dialog window
    $window.Owner = $PlayniteApi.Dialogs.GetCurrentAppWindow()
    $window.WindowStartupLocation = "CenterOwner"
    
    # Use Show or ShowDialog to show the window
    $window.ShowDialog()
}

function GetGameIdFromCmd([string]$cmd) {
    $parts = $cmd -split " --start "
    if ($parts.Count -gt 1) {
        return ($parts[1] -split " ")[0]
    }
    else {
        return ""
    }
}

function DoWork($sunshinePath, $doPath, $undoPath) {
    # Load assemblies
    Add-Type -AssemblyName System.Drawing
    $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
    
    # Set paths
    $playniteExecutablePath = Join-Path -Path $PlayniteApi.Paths.ApplicationPath -ChildPath "Playnite.DesktopApp.exe"
    $appAssetsPath = Join-Path -Path $env:LocalAppData -ChildPath "Sunshine Playnite App Export\Apps"
    if (!(Test-Path $appAssetsPath -PathType Container)) {
        New-Item -ItemType Container -Path $appAssetsPath -Force
    }

    # Set creation counter
    $count = 0

    $json = ConvertFrom-Json (Get-Content $sunshinePath -Raw)

    foreach ($game in $PlayniteApi.MainView.SelectedGames) {
        $gameLaunchCmd = "`"$playniteExecutablePath`" --start $($game.id)"

        # Set cover path and create blank file
        $sunshineGameCoverPath = [System.IO.Path]::Combine($appAssetsPath, $game.id, "box-art.png")
        if (!(Test-Path $sunshineGameCoverPath -PathType Container)) {
            New-Item -ItemType File -Path $sunshineGameCoverPath -Force | Out-Null
        }

        if ($null -ne $game.CoverImage) {

            $sourceCover = $PlayniteApi.Database.GetFullFilePath($game.CoverImage)
            if (($game.CoverImage -notmatch "^http") -and (Test-Path $sourceCover -PathType Leaf)) {

                if ([System.IO.Path]::GetExtension($game.CoverImage) -eq ".png") {
                    Copy-Item $sourceCover $sunshineGameCoverPath -Force
                }
                else {
                    # Convert cover image to compatible PNG image format
                    try {
                        # Create a BitmapImage and load the JPG image
                        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                        $bitmap.BeginInit()
                        $bitmap.UriSource = New-Object System.Uri($sourceCover, [System.UriKind]::Absolute)
                        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                        $bitmap.EndInit()
                    
                        # Create a PngBitmapEncoder and add the BitmapFrame
                        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
                        $frame = [System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap)
                        $encoder.Frames.Add($frame)
                    
                        # Save the PNG image
                        $fileStream = New-Object System.IO.FileStream($sunshineGameCoverPath, [System.IO.FileMode]::Create)
                        $encoder.Save($fileStream)
                        $fileStream.Close()
                    }
                    catch {
                        if ($null -ne $bitmap) { $bitmap = $null }
                        if ($null -ne $fileStream) { $fileStream.Close() }
                        $errorMessage = $_.Exception.Message
                        $__logger.Info("Error converting cover image of `"$($game.name)`". Error: $errorMessage")
                    }
                }

                $ids = @()
                foreach ($app in $json.apps) {
                    if ($app.id) {
                        $ids += $app.id
                    }
                }

                $id = Get-Random
                while ($ids.Contains($id.ToString())) {
                    $id = Get-Random
                }

                $newApp = New-Object -TypeName psobject
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "name" -Value $game.name
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "detached" -Value @($gameLaunchCmd)
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "image-path" -Value $sunshineGameCoverPath
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "id" -Value $id.ToString()

                # Add the prep-cmd section with doPath and undoPath
                $prepCmd = New-Object -TypeName psobject
                Add-Member -InputObject $prepCmd -MemberType NoteProperty -Name "do" -Value $doPath
                Add-Member -InputObject $prepCmd -MemberType NoteProperty -Name "undo" -Value $undoPath
                Add-Member -InputObject $prepCmd -MemberType NoteProperty -Name "elevated" -Value "false"

                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "prep-cmd" -Value @($prepCmd)
                
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "exclude-global-prep-cmd" -Value "false"
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "elevated" -Value "false"
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "auto-detach" -Value "true"
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "wait-all" -Value "true"
                Add-Member -InputObject $newApp -MemberType NoteProperty -Name "exit-timeout" -Value "5"

                $json.apps = $json.apps | ForEach-Object {
                    if ($_.detached) {
                        $gameId = GetGameIdFromCmd($_.detached[0])
                        if ($gameId -eq $game.id) {
                            $newApp
                        }
                        else {
                            $_
                        }
                    }
                    else {
                        $_
                    }
                }

                if (!($json.apps | Where-Object { 
                            if ($_.detached) {
                                $gameId = GetGameIdFromCmd($_.detached[0])
                                return $gameId -eq $game.id
                            }
                            else {
                                return $false
                            }
                        })) {
                    [object[]]$json.apps += $newApp
                }
            }
        }

        $count++
    }

    $jsonObj = ConvertTo-Json $json -Depth 100
    # Write this using utf8-noBOM, which depending on PS version, is not supported.
    # so as a workaround, we'll use WriteAllLines which defaults to utf8-noBOM
    [System.IO.File]::WriteAllLines("$env:TEMP\apps.json", $jsonObj)


    $result = [System.Windows.Forms.MessageBox]::Show("You will be prompted for administrator rights, as Sunshine now requires administrator rights in order to modify the apps.json file.", "Administrator Required", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return 0
    }
    else {
        Start-Process powershell.exe  -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -Command `"Copy-Item -Path $env:TEMP\apps.json -Destination '$sunshinePath'`"" -WindowStyle Hidden
    }


    return $count
}
