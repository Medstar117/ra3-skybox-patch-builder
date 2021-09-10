#Requires -Version 2

if ($PSCommandPath -eq $Null) {
    $PSCommandPath = $MyInvocation.MyCommand.Definition
}

if ($PSScriptRoot -eq $Null) {
    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
}

# Tip: If you want to debug this tool, it is recommended to start this powershell script directly
# Or remove -WindowStyle Hidden from the .bat file
# Removing that will allow you to see the output of the script

### Compilation steps
$global:willClearBuiltDirectory = $True
$global:willGenerateCubeMap = $True
$global:willCompilePatch = $True
$global:willCompilePatchLodLevels = $True
$global:willCopyAdditionalFiles = $True
$global:willCreateBigFile = $True

### Various parameters
# Tool path
$global:htmlPath = Join-Path (Join-Path $PSScriptRoot "panorama-to-cubemap") "index.html"
$global:cmftPath = Join-Path (Join-Path $PSScriptRoot "cmft") "cmftRelease.exe"
$global:wrathEdPath = Join-Path (Join-Path $PSScriptRoot "WrathEdDebug") "WrathEd.exe"

# Folder path
$global:patchDirectory = Join-Path $PSScriptRoot "static-patch"
$global:additionalFilesDirectory = Join-Path $patchDirectory "additional"
$global:generatedDirectory = Join-Path $patchDirectory "generated"
$global:builtDirectory = Join-Path $patchDirectory "built"
$global:basePatchStreamDirectory = Join-Path $patchDirectory "base-patch-streams"

# The path where the generated BIG file is initially stored
$global:outputDirectory = Join-Path $PSScriptRoot "output"
$global:outputBigPath = Join-Path $outputDirectory "Skybox.big"

# WrathEd compilation parameters
$global:inputXml = Join-Path $patchDirectory "static.xml"
$global:newStreamVersion = ".sky"
$global:basePatchStreamName = "static.12.manifest"

# Automatically generated skybox mapping
$global:outputCubeMap = Join-Path $generatedDirectory "skybox"
$global:skyboxXml = "$outputCubeMap.xml"
$global:skyboxXmlContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<AssetDeclaration xmlns="uri:ea.com:eala:asset" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<Includes>
	</Includes>
	<Texture id="EVDefault" File="skybox.dds" Type="CubeTexture" />
</AssetDeclaration>
"@

### UI text
$global:mainTitle = "RAAA Skybox Patch Generator"
$global:mainDescription = @"
Lanyi's skybox patch generator, which should be compatible with most mods.

Requires some basic knowledge about modding Red Alert 3 (like how to get the game or another mod to load a BIG file).

Usage:
  1) Click "Create CubeMap File" to convert a 2:1 texture into a cubemap.
  2) Click "Create Skybox Patch", which will open a window asking to select the cubemap created in step 1.
  3) If the correct file was selected in step 2, the skybox patch can be generated.
"@
$global:cancelDescription = "If it is taking too long to generate the skybox map, consider clicking the `“Cancel`” button and try again."
$global:htmlButtonText = "Create CubeMap File"
$global:compileButtonText = "Create Skybox Patch"
$global:cancelButtonText = "Cancel"
$global:showAdvancedButtonText = "Show Advanced Options"
$global:hideAdvancedButtonText = "Hide Advanced Options"
$global:compilePatchText = "Compile patch"
$global:compilePatchLodLevelsText = "Compile low and medium LOD patches"
$global:basePatchStreamDescription = "Create a patch based on this manifest"
$global:newStreamVersionText = "New manifest version number"
$global:editThisScriptText = @"
If you want to make further changes, consider directly modifying the
<Hyperlink x:Name="ThisScriptLink" NavigateUri="$PSCommandPath">
    $((Get-Item $PSCommandPath).Name)
</Hyperlink>
file (can be opened directly in Notepad; requires rebooting $mainTitle after modification）
"@
$global:statusMessage = "is{0}"
$global:statusFailedMessage = "{0}failed"
$global:clearBuiltDirectoryStatus = "Clear last compiled file"
$global:generateCubeMapStatus = "Process skybox cubemap"
$global:wedStatus = "Compiling"
$global:copyAdditionalFilesStatus = "Copy extra files"
$global:createBigFileStatus = "Create BIG file"
$global:emptyBigMessage = "No files added to the patch's BIG file, or maybe something else went wrong"
$global:saveFailedMessage = "Failed to save BIG file: {0}"
$global:chooseSkyboxTextureTitle = "Choose a skybox map"
$global:skyboxTextureFilter = "Skybox texture （*.png;*.tga;*.jpg;*.bmp;*.dds;*.hdr）|*.png;*.tga;*.jpg;*.bmp;*.dds;*.hdr|All files （*.*）|*.*"
$global:saveBigFileTitle = "Save BIG file"
$global:bigFileFilter = "BIG file （*.big）|*.big|All files （*.*）|*.*"
$global:creditsText = @"
<Hyperlink NavigateUri="https://github.com/lanyizi/ra3-skybox-patch-builder">
    $mainTitle
</Hyperlink> v0.11
<LineBreak />
This generator uses
<Hyperlink NavigateUri="https://github.com/lanyizi/panorama-to-cubemap">
    panorama-to-cubemap
</Hyperlink>,<Hyperlink NavigateUri="https://github.com/dariomanesku/cmft">
    cmft
</Hyperlink> and
<Hyperlink NavigateUri="https://github.com/Qibbi/WrathEd2012">
    WrathEd
</Hyperlink> as well.
"@

$xaml = [xml]@"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="$mainTitle" Width="380" Height="420">
    <ScrollViewer Margin="0" VerticalScrollBarVisibility="Auto">
        <StackPanel Orientation="Vertical" Margin="8">
            <TextBlock x:Name="MainDescription" Margin="4" TextWrapping="Wrap" />
            <Button x:Name="HtmlButton"
                Margin="4" HorizontalAlignment="Left"
                Content="$htmlButtonText"
            />
            <Button x:Name="CompileButton"
                Margin="4" HorizontalAlignment="Left"
                Content="$compileButtonText"
            />
            <Button x:Name="CancelButton"
                Margin="4" HorizontalAlignment="Left" Visibility="Collapsed"
                Content="$cancelButtonText"
            />
            <TextBlock x:Name="CancelDescription" Margin="4" TextWrapping="Wrap" Visibility="Collapsed" />
            <TextBlock x:Name="StatusDescription" Margin="4" TextWrapping="Wrap">
                $creditsText
            </TextBlock>
            <Button x:Name="ToggleAdvancedButton"
                Margin="4,8,4,4" HorizontalAlignment="Left"
                Content="$showAdvancedButtonText"
            />
            <StackPanel x:Name="AdvancedPanel" Orientation="Vertical" Margin="4" Visibility="Collapsed">
                <TextBlock Margin="4" TextWrapping="Wrap">
                    $editThisScriptText
                </TextBlock>
                <CheckBox x:Name="ToggleClearBuiltDirectory" 
                    Margin="4" HorizontalAlignment="Left" Content="$clearBuiltDirectoryStatus" 
                />
                <CheckBox x:Name="ToggleGenerateCubeMap" 
                    Margin="4" HorizontalAlignment="Left" Content="$generateCubeMapStatus" 
                />
                <CheckBox x:Name="ToggleCompilePatch" 
                    Margin="4" HorizontalAlignment="Left" Content="$compilePatchText" 
                />
                <CheckBox x:Name="ToggleCompilePatchLodLevels" 
                    Margin="4" HorizontalAlignment="Left" Content="$compilePatchLodLevelsText" 
                />
                <CheckBox x:Name="ToggleCopyAdditionalFiles" 
                    Margin="4" HorizontalAlignment="Left" Content="$copyAdditionalFilesStatus" 
                />
                <CheckBox x:Name="ToggleCreateBigFile" 
                    Margin="4" HorizontalAlignment="Left" Content="$createBigFileStatus" 
                />
                <Label Margin="4,8,4,0" HorizontalAlignment="Left" Content="$basePatchStreamDescription" />
                <TextBox x:Name="BasePatchStreamNameInput" Margin="12,0" />
                <Label Margin="4,8,4,0" HorizontalAlignment="Left" Content="$newStreamVersionText" />
                <TextBox x:Name="NewStreamVersionInput" Margin="12,0" />
            </StackPanel>
        </StackPanel>
    </ScrollViewer>
</Window>
"@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[Windows.Forms.Application]::EnableVisualStyles()

function Initialize-Wpf($window, $nativeWindow) {
    $window.Add_SourceInitialized({
        # Used to get the native handle of the window
        $interopHelper = New-Object Windows.Interop.WindowInteropHelper -ArgumentList $window
        $nativeWindow.AssignHandle($interopHelper.Handle)
    }.GetNewClosure())

    # Adds support for clicking on hyperlinks and support for checkboxes
    $window.Add_Loaded({
        function Add-HyperlinkEventHandler($owner) {
            $count = [Windows.Media.VisualTreeHelper]::GetChildrenCount($owner)
            for ($i = 0; $i -lt $count; $i = $i + 1) {
                $child = [Windows.Media.VisualTreeHelper]::GetChild($owner, $i);
                if ($child -eq $null) {
                    continue
                }
                if ($child.GetType().Equals([Windows.Controls.TextBlock])) {
                    foreach ($inline in $child.Inlines) {
                        if ($inline.GetType().Equals([Windows.Documents.Hyperlink])) {
                            $inline.Add_RequestNavigate({ 
                                param ($sender, $e)
                                if ($e.Uri.IsFile) {
                                    Start-Process explorer.exe -ArgumentList "/select,`"$($e.Uri.LocalPath)`""
                                }
                                else {
                                    Start-Process $e.Uri
                                }
                            })
                        }
                    }
                }
                if ($child.GetType().Equals([Windows.Controls.CheckBox])) {
                    $varname = $child.Name.Replace("Toggle", "will")
                    $variable = Get-Variable $varname -Scope Global
                    $child.IsChecked = $variable.Value
                    $child.Add_Click({ 
                        $variable.Value = ($child.IsChecked -eq $True) 
                    }.GetNewClosure())
                }
                Add-HyperlinkEventHandler $child
            }
        }
        Add-HyperlinkEventHandler $window
    }.GetNewClosure())

    $window.FindName("MainDescription").Text = $mainDescription
    $window.FindName("CancelDescription").Text = $cancelDescription
    $htmlButton = $window.FindName("HtmlButton")
    $compileButton = $window.FindName("CompileButton")
    $cancelButton = $window.FindName("CancelButton")
    $cancelDescription = $window.FindName("CancelDescription")
    $statusDescription = $window.FindName("StatusDescription")
    $toggleAdvancedButton = $window.FindName("ToggleAdvancedButton")
    $advancedPanel = $window.FindName("AdvancedPanel")
    $basePatchStreamNameInput = $window.FindName("BasePatchStreamNameInput")
    $newStreamVersionInput = $window.FindName("NewStreamVersionInput")

    $context = @{
        NativeWindow = $nativeWindow
    }

    $currentlyTrackedProcesses = New-Object Collections.Generic.List[object]
    $context.ChangeTrackedProcesses = {
        param ($newValues)

        if ($newValues -eq $Null) {
            $currentlyTrackedProcesses.Clear()
        }
        else {
            foreach ($process in $newValues) {
                $currentlyTrackedProcesses.Add($process)
            }
        }
        if ($currentlyTrackedProcesses.Count -gt 0) {
            $cancelButton.Visibility = [Windows.Visibility]::Visible
        }
        else {
            $cancelButton.Visibility = [Windows.Visibility]::Collapsed
        }
    }.GetNewClosure()
    & $context.ChangeTrackedProcesses $Null

    $context.SetStatus = {
        param ($statusText)
        $context.StatusText = $statusText
        $statusDescription.Text = [string]::Format($statusMessage, $context.StatusText)
        $statusDescription = [Windows.Visibility]::Visible
    }.GetNewClosure()

    $context.Complete = {
        param ($succeeded)
        $statusDescription.Text = ""
        $compileButton.IsEnabled = $True
        $advancedPanel.IsEnabled = $True
        if (-not $succeeded) {
            # 假如是由用户自己取消的 那就不需要下面的弹框报错了
            if (-not $context.IsCancelled) {
                $what = [string]::Format($statusFailedMessage, $context.StatusText)
                [Windows.Forms.MessageBox]::Show($what, $mainTitle)
            }
        }
    }.GetNewClosure()

    $htmlButton.Add_Click({
        & $htmlPath
    })

    $compileButton.Add_Click({
        $compileButton.IsEnabled = $False
        $advancedPanel.IsEnabled = $False
        $context.IsCancelled = $False
        Start-PatchBuild $context $compileButton.Dispatcher
    }.GetNewClosure())

    $cancelButton.Add_Click({
        foreach ($process in $currentlyTrackedProcesses) {
            $process.Kill()
        }
        $context.IsCancelled = $True
        & $context.ChangeTrackedProcesses $Null
    }.GetNewClosure())

    $toggleAdvancedButton.Add_Click({
        if ($advancedPanel.Visibility -eq [Windows.Visibility]::Visible) {
            $advancedPanel.Visibility = [Windows.Visibility]::Collapsed
            $toggleAdvancedButton.Content = $showAdvancedButtonText
        }
        else {
            $advancedPanel.Visibility = [Windows.Visibility]::Visible
            $toggleAdvancedButton.Content = $hideAdvancedButtonText
        }
    }.GetNewClosure())

    $basePatchStreamNameInput.Text = $basePatchStreamName
    $basePatchStreamNameInput.Add_TextChanged({
        (Get-Variable "basePatchStreamName" -Scope Global).Value = $basePatchStreamNameInput.Text
    }.GetNewClosure())

    $newStreamVersionInput.Text = $newStreamVersion
    $newStreamVersionInput.Add_TextChanged({
        (Get-Variable "newStreamVersion" -Scope Global).Value = $newStreamVersionInput.Text
    }.GetNewClosure())

    $context.ShowCancelText = { $cancelDescription.Visibility = [Windows.Visibility]::Visible }.GetNewClosure()
    $context.HideCancelText = { $cancelDescription.Visibility = [Windows.Visibility]::Collapsed }.GetNewClosure()
}

function global:Start-PatchBuild($context, $dispatcher) {

    $context.DoEvents = {
        # Allows the program to update text while not in the foreground
        $dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [Action]{});
    }.GetNewClosure()

    $context.SynchronizationContext = New-Object Windows.Threading.DispatcherSynchronizationContext -ArgumentList $dispatcher

    # Used for clearing temporary files
    $context.ClearBuiltDirectory = {
        & $context.SetStatus $clearBuiltDirectoryStatus
        & $context.DoEvents
        # Delete the file, but do not delete the file in the soft link (delete the soft link itself)
        function Clear-MyDirectory($currentFolder) {
            foreach ($child in (Get-ChildItem $currentFolder)) {
                $isDirectory = ($child.Attributes -band [IO.FileAttributes]::Directory) -ne 0
                $isReparsePoint = ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
                if ($isDirectory) {
                    if ($isReparsePoint) {
                        # Only delete the soft link itself
                        cmd.exe /c rmdir $child.FullName
                        continue
                    }
                    else {
                        Clear-MyDirectory $child.FullName
                    }
                }
                Remove-Item $child.FullName -Force
            }
        }
        $target = New-Item -ItemType Directory -Force -Path $builtDirectory
        Clear-MyDirectory $target.FullName
    }.GetNewClosure()

    # CubeMap used to create the skybox
    $context.GenerateCubeMap = {
        & $context.SetStatus $generateCubeMapStatus
        $skyboxTexturePath = Get-SkyboxTexturePath $context.NativeWindow
        if ($skyboxTexturePath -eq $Null) {
            & $context.Complete $True
            return
        }

        New-Item -ItemType Directory -Path $generatedDirectory -Force | Out-Null

        $cmftProcess = Generate-SkyboxCubeMap $skyboxTexturePath $context.SynchronizationContext
        & $context.ChangeTrackedProcesses $cmftProcess
        $cmftProcess.Add_Exited($context.OnCubeMapGenerationEnd)
        $cmftProcess.Start()
        & $context.ShowCancelText
    }.GetNewClosure()

    $context.OnCubeMapGenerationEnd = {
        param ($sender)

        & $context.ChangeTrackedProcesses $Null
        & $context.HideCancelText
        if (-not $sender.Succeeded) {
            & $context.Complete $False
            return
        }

        $skyboxXmlFile = New-Item -ItemType File -Path $skyboxXml -Force
        [IO.File]::WriteAllText($skyboxXmlFile.FullName, $skyboxXmlContent)
        & $context.StartWrathEd
    }.GetNewClosure()

    # Used to start WrathEd and to compile patches
    $context.StartWrathEd = {
        & $context.SetStatus $wedStatus
        Start-WrathEd $context.SynchronizationContext $context.ChangeTrackedProcesses $context.OnWrathEdCompleted
    }.GetNewClosure()

    $context.OnWrathEdCompleted = {
        param ($succeeded)

        if (-not $succeeded) {
            & $context.Complete $False
            return
        }

        if ($willCopyAdditionalFiles) {
            & $context.SetStatus $copyAdditionalFilesStatus
            & $context.DoEvents

            $additionalFiles = Join-Path $additionalFilesDirectory "*"
            Copy-Item -Path $additionalFiles -Destination $builtDirectory -Force -Recurse
            if (-not $?) {
                & $context.Complete $False
                return        
            }
        }

        if ($willCreateBigFile) {
            & $context.SetStatus $createBigFileStatus
            & $context.DoEvents

            Create-BigFile $builtDirectory $context.NativeWindow
        }
        & $context.Complete $True
    }.GetNewClosure()

    if ($willClearBuiltDirectory) {
        & $context.ClearBuiltDirectory
    }

    if ($willGenerateCubeMap) {
        & $context.GenerateCubeMap
    }
    else {
        & $context.StartWrathEd
    }
}

function global:Get-SkyboxTexturePath($nativeWindow) {
    $openFileDialog = New-Object Windows.Forms.OpenFileDialog
    try {
        $openFileDialog.Title = $chooseSkyboxTextureTitle
        $openFileDialog.Filter = $skyboxTextureFilter
        if ($openFileDialog.ShowDialog($nativeWindow) -eq [Windows.Forms.DialogResult]::OK) {
            return $openFileDialog.FileName
        }
    }
    finally {
        $openFileDialog.Dispose()
    }
    return $Null
}

function global:Generate-SkyboxCubeMap($texturePath, $synchronizationContext) {
    $args = @(
        "--input `"$texturePath`""
        "--filter radiance"
        "--edgefixup warp"
        "--srcFaceSize 0"
        "--excludeBase true"
        "--mipCount 9"
        "--generateMipChain false"
        "--glossScale 17"
        "--glossBias 3"
        "--lightingModel blinnbrdf"
        "--dstFaceSize 0"
        "--numCpuProcessingThreads 4"
        "--useOpenCL true"
        "--clVendor anyGpuVendor"
        "--deviceType gpu"
        "--inputGammaNumerator 1.0"
        "--inputGammaDenominator 1.0"
        "--outputGammaNumerator 1.0"
        "--outputGammaDenominator 1.0"
        "--output0 `"$outputCubeMap`""
        "--output0params dds,bgra8,cubemap"
    ) -join " "
    return [JobSupport]::Prepare($cmftPath, $args, $synchronizationContext)
}

function global:Start-WrathEd($synchronizationContext, $changeTrackedProcesses, $onCompleted) {

    $builtDataDirectory = [IO.Path]::Combine($builtDirectory, "data")
    New-Item -ItemType Directory -Force -Path $builtDataDirectory | Out-Null

    function Get-LodName($originalPath, $lodPostFix) {
        $directory = [IO.Path]::GetDirectoryName($originalPath)
        $stem = [IO.Path]::GetFileNameWithoutExtension($originalPath)
        $stems = $stem -split "(?=[^\w])"
        $stems[$stems.Length - 2] += $lodPostFix
        $stem = -join $stems
        $newFileName = $stem + [IO.Path]::GetExtension($originalPath)
        return Join-Path $directory $newFileName
    }
    $inputXmlM = Get-LodName $inputXml "_m"
    $inputXmlL = Get-LodName $inputXml "_l"
    $basePatchStream = Join-Path $basePatchStreamDirectory $basePatchStreamName
    $basePatchStreamM = Get-LodName $basePatchStream "_m"
    $basePatchStreamL = Get-LodName $basePatchStream "_l"

    function Get-WedArguments($xml, $bps) {
        $bpsName = (Get-Item $bps).Name
        return @(
            "-gameDefinition:`"Red Alert 3`""
            "-compile:`"$xml`""
            "-out `"$builtDataDirectory`""
            "-version:`"$newStreamVersion`""
            "-bps:`"$bpsName,$bps`""
        ) -join " "
    }

    $context = @{
        Args = (Get-WedArguments $inputXml $basePatchStream)
        ArgsM = (Get-WedArguments $inputXmlM $basePatchStreamM)
        ArgsL = (Get-WedArguments $inputXmlL $basePatchStreamL)
        ChangeTrackedProcesses = $changeTrackedProcesses
        OnCompleted = $onCompleted 
        SynchronizationContext = $synchronizationContext

        Steps = @()
        StepCounter = 0
    }

    $context.LaunchWrathEd = {
        param ($context, $wedArgs)
        [Console]::WriteLine("WED Arguments: $($wedArgs)");
        $process = [JobSupport]::Prepare($wrathEdPath, $wedArgs, $context.SynchronizationContext)
        & $context.ChangeTrackedProcesses $process
        $process.Add_Exited($context.StepEnd)
        $process.WorkingDirectory = $builtDirectory
        $process.Start()
    }

    if ($willCompilePatch) {
        $context.Steps += {
            param ($context)
            & $context.LaunchWrathEd $context $context.Args
        }
    }

    if ($willCompilePatchLodLevels) {
        $context.Steps += {
            param ($context)
            & $context.LaunchWrathEd $context $context.ArgsM
        }
        $context.Steps += {
            param ($context)
            & $context.LaunchWrathEd $context $context.ArgsL
        }
    }

    $context.StepEnd = {
        param ($sender)
        $succeeded = $sender.Succeeded
        & $context.ChangeTrackedProcesses $Null
        # WrathEd automatically generates a stream of stringhashes, which is not needed, so delete it
        $stringHashes = Join-Path $builtDataDirectory "stringhashes.*"
        Remove-Item $stringHashes
        if (-not $succeeded) {
            & $context.OnCompleted $False
            return
        }
        $context.StepCounter = $context.StepCounter + 1
        if ($context.StepCounter -lt $context.Steps.Length) {
            & $context.Steps[$context.StepCounter] $context
        }
        else {
            & $context.OnCompleted $True
            return
        }
    }.GetNewClosure()

    if ($context.Steps.Length -gt 0) {
        & $context.Steps[0] $context
    }
    else {
        & $context.OnCompleted $True
    }
}

function global:Create-BigFile($sourceDirectory, $nativeWindow) {
    $list = New-Object Collections.Generic.List[HashTable]
    Get-BigFileList $list "" (Get-Item "$sourceDirectory")
    if ($list.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show([string]::Format($emptyBigMessage, $Error[0]), $mainTitle)
        return
    }

    $outputFile = New-Item $outputBigPath -ItemType File -Force
    $output = [IO.File]::Open($outputFile.FullName, [IO.FileMode]::Create)
    try {
        function Write-ByteArray($array) {
            $output.Write($array, 0, $array.Length)
        }

        function Write-BigEndianValue($v) {
            $array = [BitConverter]::GetBytes($v)
            if ([BitConverter]::IsLittleEndian) {
                [Array]::Reverse($array)
            }
            Write-ByteArray $array
        }

        function Check-StreamPosition() {
            if ($output.Position -gt [UInt32]::MaxValue) {
                throw [IO.IOException]"BIG File too large"
            }
        }

        # BIG header
        Write-ByteArray ([Text.Encoding]::ASCII.GetBytes("BIG4"))
        # Skip file size first
        $output.Position += 4
        # Number of files
        Write-BigEndianValue ([UInt32]($list.Count))
        # Skip the "first file location" first
        $output.Position += 4

        # File list
        foreach ($entry in $list) {
            & Check-StreamPosition
            $entry.EntryOffset = $output.Position
            # Skip size and position first
            $output.Position += 8
            # Write down the file name first
            Write-ByteArray $entry.PathBytes
            # 0-terminate string
            $output.WriteByte([byte]0)
        }

        $firstEntryOffset = $output.Position
        # Write file contents
        $buffer = New-Object byte[] 81920
        foreach ($entry in $list) {
            $fromFile = $entry.File.OpenRead()
            try {
                $entry.FileOffset = $output.Position
                $bytesRead = 0
                do {
                    & Check-StreamPosition
                    $bytesRead = $fromFile.Read($buffer, 0, $buffer.Length)
                    $output.Write($buffer, 0, $bytesRead)
                }
                while ($bytesRead -gt 0)
                $entry.FileSize = ($output.Position - $entry.FileOffset)
            }
            finally {
                $fromFile.Dispose()
            }
        }

        $bigFileSize = $output.Position

        # Go back to the beginning
        $output.Position = 4
        Write-BigEndianValue ([UInt32]($bigFileSize))
        $output.Position += 4
        Write-BigEndianValue ([UInt32]($firstEntryOffset))

        # Finish writing the list of files
        foreach ($entry in $list) {
            $output.Position = $entry.EntryOffset
            # Write size and position
            Write-BigEndianValue ([UInt32]($entry.FileOffset))
            Write-BigEndianValue ([UInt32]($entry.FileSize))
        }
    }
    catch {
        [Windows.Forms.MessageBox]::Show([string]::Format($saveFailedMessage, $_), $mainTitle)
        explorer.exe "/select,`"$($outputFile.FullName)`""
        return
    }
    finally {
        $output.Dispose()
    }

    # Display the "Save" dialog box, select a place to save the patch
    $finalBigName = $Null
    $saveFileDialog = New-Object Windows.Forms.SaveFileDialog
    try {
        $saveFileDialog.Title = $saveBigFileTitle
        $saveFileDialog.Filter = $bigFileFilter
        if ($saveFileDialog.ShowDialog($nativeWindow) -eq [Windows.Forms.DialogResult]::OK) {
            $finalBigName = $saveFileDialog.FileName
        }
    }
    finally {
        $saveFileDialog.Dispose()
    }

    if ($finalBigName -ne $Null) {
        Move-Item -Path $outputBigPath -Destination $finalBigName -Force
        if (-not $?) {
            [Windows.Forms.MessageBox]::Show([string]::Format($saveFailedMessage, $Error[0]), $mainTitle)
        }
    }
}

function Get-BigFileList($outList, $currentPrefix, $currentDirectory) {
    # Folders
    foreach ($child in $currentDirectory.GetDirectories()) {
        $name = $child.Name.ToLowerInvariant()
        $childPath = "$currentPrefix$name"
        Get-BigFileList $outList "$childPath\" $child
    }
    # Documents
    foreach ($child in $currentDirectory.GetFiles()) {
        $name = $child.Name.ToLowerInvariant()
        $childPath = "$currentPrefix$name"

        $outList.Add(@{
            File = $child
            PathBytes = [Text.Encoding]::UTF8.GetBytes($childPath)
        })
    }
}

Set-Location $PSScriptRoot
[Environment]::CurrentDirectory = $PSScriptRoot

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$jobSupport = @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

public class JobSupport
{
    public enum JOBOBJECTINFOCLASS
    {
        AssociateCompletionPortInformation = 7,
        BasicLimitInformation = 2,
        BasicUIRestrictions = 4,
        EndOfJobTimeInformation = 6,
        ExtendedLimitInformation = 9,
        SecurityLimitInformation = 5,
        GroupInformation = 11
    }

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_BASIC_LIMIT_INFORMATION
    {
        public Int64 PerProcessUserTimeLimit;
        public Int64 PerJobUserTimeLimit;
        public UInt32 LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public UInt32 ActiveProcessLimit;
        public Int64 Affinity;
        public UInt32 PriorityClass;
        public UInt32 SchedulingClass;
    }


    [StructLayout(LayoutKind.Sequential)]
    struct IO_COUNTERS
    {
        public UInt64 ReadOperationCount;
        public UInt64 WriteOperationCount;
        public UInt64 OtherOperationCount;
        public UInt64 ReadTransferCount;
        public UInt64 WriteTransferCount;
        public UInt64 OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    {
        public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string lpName);

    [DllImport("kernel32.dll")]
    public static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

    [DllImport("kernel32.dll")]
    public static extern bool SetInformationJobObject(IntPtr hJob, JOBOBJECTINFOCLASS JobObjectInfoClass, IntPtr lpJobObjectInfo, uint cbJobObjectInfoLength);

    [DllImport("kernel32.dll")]
    public static extern IntPtr GetCurrentProcess();

    public class TrackedProcess
    {
        private Process _process;
        private SynchronizationContext _synchronizationContext;
        private bool _succeeded;

        public bool Succeeded { get { return _succeeded; } }
        public string WorkingDirectory
        {
            get { return _process.StartInfo.WorkingDirectory; }
            set { _process.StartInfo.WorkingDirectory = value; }
        }
        public event EventHandler Exited;

        public TrackedProcess(Process process, SynchronizationContext synchronizationContext)
        {
            _process = process;
            _synchronizationContext = synchronizationContext;
            _process.EnableRaisingEvents = true;
            _process.Exited += ExitEventHandler;
        }

        public void Start()
        {
            _process.Start();
        }

        public void Kill()
        {
            _process.Kill();
        }

        private void ExitEventHandler(object sender, EventArgs e)
        {
            _succeeded = _process.ExitCode == 0;
            _process.Exited -= ExitEventHandler;
            _synchronizationContext.Post(ActualEventExecutor, null);
            _process.Dispose();
        }

        private void ActualEventExecutor(object state)
        {
            EventHandler handlers = Exited;
            Exited = null;
            if (handlers != null)
            {
                handlers(this, EventArgs.Empty);
            }
        }
    }

    private const UInt32 JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000;
    private static bool _initialized = false;

    public static TrackedProcess Prepare(string fileName, string arguments, SynchronizationContext context)
    {
        if (!_initialized)
        {
            Initialize();
        }

        Process process = new Process();
        process.StartInfo.FileName = fileName;
        process.StartInfo.Arguments = arguments;
        process.StartInfo.UseShellExecute = false;

        return new TrackedProcess(process, context);
    }

    private static bool Initialize()
    {
        IntPtr job = CreateJobObject(IntPtr.Zero, null);

        JOBOBJECT_BASIC_LIMIT_INFORMATION info = new JOBOBJECT_BASIC_LIMIT_INFORMATION();
        info.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;

        JOBOBJECT_EXTENDED_LIMIT_INFORMATION extendedInfo = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
        extendedInfo.BasicLimitInformation = info;

        int length = Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
        IntPtr extendedInfoPtr = Marshal.AllocHGlobal(length);
        try
        {
            Marshal.StructureToPtr(extendedInfo, extendedInfoPtr, false);

            SetInformationJobObject(job, JOBOBJECTINFOCLASS.ExtendedLimitInformation, extendedInfoPtr, (uint)length);

            IntPtr hProcess = GetCurrentProcess();
            return AssignProcessToJobObject(job, hProcess);
        }
        finally
        {
            Marshal.FreeHGlobal(extendedInfoPtr);
        }
    }
}
"@
Add-Type -TypeDefinition $jobSupport
$nativeWindow = New-Object Windows.Forms.NativeWindow
Initialize-Wpf $window $nativeWindow
$window.ShowDialog()
$nativeWindow.ReleaseHandle()