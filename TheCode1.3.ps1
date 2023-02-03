add-type -AssemblyName presentationframework
#getting dll assembly path automatically
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
{ 
   $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 
}
else
{ 
   $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
   if (!$ScriptPath){ $ScriptPath = "." } 
}
$dllpath=$scriptpath+"\buildreport.dll"
add-type -Path $dllpath

Function New-WPFMessageBox {

 
    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 20,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families.Name | Select -Skip 1 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segoe UI"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
            New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content) 
        }
        Catch
        {
            $_
        }        
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }
    

    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }
    

    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    }
}
function showabout {

$textblock=New-Object System.Windows.Controls.TextBlock
$textblock.Text = "A fully sick app developed by that`r`nSys Engineering legend P. Rijal!"
$textblock.Padding=10
$textblock.FontSize=18
$textblock.HorizontalAlignment="Left"

$stackpanel=new-object System.Windows.Controls.StackPanel
$stackpanel.Orientation="horizontal"
$stackpanel.AddChild($textblock)

New-WPFMessageBox -content $stackpanel -Title "About Buildreport" -titlebackground steelblue -titletextforeground white -cornerradius 0
}

function showresult {

$textblock=New-Object System.Windows.Controls.TextBlock
$textblock.Text = "Exported buildreport successfully to $exportpathname"
$textblock.Padding=10
$textblock.FontSize=18
$textblock.HorizontalAlignment="Left"

$stackpanel=new-object System.Windows.Controls.StackPanel
$stackpanel.Orientation="horizontal"
$stackpanel.AddChild($textblock)

New-WPFMessageBox -content $stackpanel -Title "Result" -titlebackground steelblue -titletextforeground white -cornerradius 0
}

[xml]$form = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      
        Title="Build Report" Height="450" Width="610" ResizeMode="CanResizeWithGrip">
    <TabControl Name="tabControl">
        <TabItem Header="General">
            <Grid Background="#FFE5E5E5">

                <Label Name="label" Content="Enter device names or keywords below to generate the build report." HorizontalAlignment="Left" Margin="21,10,0,0" VerticalAlignment="Top" Width="570" Grid.ColumnSpan="3" Height="26"/>
                <Label Name="label1" Content="Keyword 1" HorizontalAlignment="Left" Margin="49,60,0,0" VerticalAlignment="Top" Width="148" Height="26"/>
                <Label Name="label2" Content="Keyword 2" HorizontalAlignment="Left" Margin="49,99,0,0" VerticalAlignment="Top" Height="26" Width="66"/>
                <TextBox Name="keyword1" HorizontalAlignment="Left" Margin="186,64,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" Height="18"/>
                <TextBox Name="keyword2" HorizontalAlignment="Left" Margin="186,103,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" Height="18"/>
                <Button Name="buildreport" Content="Build Report" HorizontalAlignment="Left" Margin="186,298,0,0" VerticalAlignment="Top" Width="120" IsCancel="False" Height="20"/>
                <Label Name="label2_Copy" Content="Keyword 3" HorizontalAlignment="Left" Margin="49,138,0,0" VerticalAlignment="Top" Height="26" Width="66"/>
                <TextBox Name="keyword3" HorizontalAlignment="Left" Margin="186,142,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" RenderTransformOrigin="0.5,2.271" Height="18"/>
                <Label Name="label4" Content="Export Path &amp; filename" HorizontalAlignment="Left" Margin="49,0,0,0" VerticalAlignment="Center" Height="26" Width="132"/>
                <TextBox Name="exportpath" HorizontalAlignment="Left" Margin="186,0,0,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="381" Height="18"/>
                <Label Name="label5" Content="e.g.   C:\Users\prijal\Desktop\buildreport.pdf" HorizontalAlignment="Left" Margin="186,217,0,0" VerticalAlignment="Top" Width="299" Height="26"/>
                <Button Name="about" Content="About" HorizontalAlignment="Left" Height="23" Margin="537,0,0,0" VerticalAlignment="Top" Width="47" FontFamily="Bahnschrift Light" FontSize="11" RenderTransformOrigin="0.766,2.525"/>
                <TextBlock Name="textBlock" Margin="0,0,20,10" TextWrapping="Wrap" Text="SICE" Height="57" Width="96" FontSize="48" FontFamily="Haettenschweiler" VerticalAlignment="Bottom" HorizontalAlignment="Right" Foreground="#FFE7F3E9" Background="#FF050B6B" TextAlignment="Center"/>
              
                <TextBlock Name="textBlock1" HorizontalAlignment="Left" Margin="21,362,0,0" TextWrapping="Wrap" Text="© 2022 SICE Pty Ltd" VerticalAlignment="Top" Width="160" Height="27" FontSize="10" FontWeight="Bold"/>

            </Grid>
        </TabItem>
    </TabControl>
</Window>

"@

$nr=(new-object System.Xml.XmlNodeReader $form)
$win=[windows.markup.xamlreader]::Load( $nr )

$keyword1=$win.FindName("keyword1")
$keyword2=$win.FindName("keyword2")
$keyword3=$win.FindName("keyword3")
$exportpath=$win.FindName("exportpath")
$buildreport=$win.FindName("buildreport")
$button1=$win.FindName("button1")
$about=$win.FindName("about")

$about.add_click({

showabout

})

$buildreport.add_click({

# Set basic PDF settings for the document
Function Create-PDF([iTextSharp.text.Document]$Document, [string]$File, [int32]$TopMargin, [int32]$BottomMargin, [int32]$LeftMargin, [int32]$RightMargin, [string]$Author)
{
    $Document.SetPageSize([iTextSharp.text.PageSize]::A4)
    $Document.SetMargins($LeftMargin, $RightMargin, $TopMargin, $BottomMargin)
    [void][iTextSharp.text.pdf.PdfWriter]::GetInstance($Document, [System.IO.File]::Create($File))
    $Document.AddAuthor($Author)
    
}
# Add a text paragraph to the document, optionally with a font name, size and color
function Add-Text([iTextSharp.text.Document]$Document, [string]$Text, [string]$FontName = "Arial", [int32]$FontSize = 12, [string]$Color = "BLACK")
{
    $p = New-Object iTextSharp.text.Paragraph
    $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::NORMAL)
    $p.SpacingBefore = 2
    $p.SpacingAfter = 2
    $p.Add($Text)
    $Document.Add($p)
}
# Add a title to the document, optionally with a font name, size, color and centered
function Add-Title([iTextSharp.text.Document]$Document, [string]$Text, [Switch]$Centered, [string]$FontName = "Arial", [int32]$FontSize = 16, [string]$Color = "BLACK")
{
    $p = New-Object iTextSharp.text.Paragraph
    $p.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::BOLD)
    if($Centered) { $p.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER }
    $p.SpacingBefore = 5
    $p.SpacingAfter = 5
    $p.Add($Text)
    $Document.Add($p)
}

# Add a table to the document with an array as the data, a number of columns, and optionally centered
function Add-Tabletitle
{
  [CmdletBinding()]
  param
  (
    [iTextSharp.text.Document]$Document,
    [String[]]$Dataset,
    [int]$Cols = 3,
    [switch]$Centered,
    [switch] $UsegrayBG = $true,
    [switch] $UseConsoleFont = $false,
    [switch] $Noborder = $false,
    [int] $WidthPercentage = 0
    
  )
    
  $t = New-Object -TypeName iTextSharp.text.pdf.PDFPTable -ArgumentList ($Cols)
  if($WidthPercentage -ne 0){
    $t.WidthPercentage = $WidthPercentage
  }
  $Gray = new-object iTextSharp.text.Color 200, 200, 240
  $ConsoleFont = [iTextSharp.text.FontFactory]::GetFont("Courier", $ParagraphFontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.Color]::BLACK)
  
  $t.SpacingBefore = 5
  $t.SpacingAfter = 0
  if (!$Centered)
  {
    $t.HorizontalAlignment = 0
  }
  foreach ($data in $Dataset)
  {
    $p = $null 
    
    if($UseConsoleFont){
      $p = New-Object -TypeName iTextSharp.text.Phrase  $data, $ConsoleFont
    } else {
      $p = New-Object -TypeName iTextSharp.text.Phrase  $data
      
    }
    
    #$t.AddCell($data)
    $cell = New-Object iTextSharp.text.pdf.PdfPCell $p
    
    if($UsegrayBG){
      $cell.BackgroundColor = $Gray
    }

    if($Noborder){
      $cell.Border = $null
    }
    
    $t.AddCell($cell)
  }
  $Document.Add($t)
}
function Add-Table([iTextSharp.text.Document]$Document, [string[]]$Dataset, [int32]$Cols = 3, [Switch]$Centered, [switch] $UsegrayBG = $true)
{
    $t = New-Object iTextSharp.text.pdf.PDFPTable($Cols)
    $Gray = new-object iTextSharp.text.Color 200, 200, 240
    $t.SpacingBefore = 5
    $t.SpacingAfter = 20
    if(!$Centered) { $t.HorizontalAlignment = 0 }
    $cell = New-Object iTextSharp.text.pdf.PdfPCell
    if($UsegrayBG) { $cell.backgroundColor = $Gray }
    foreach($data in $Dataset)
    {
        $t.AddCell($data);
    }
    $Document.Add($t)
}

function Add-TableHeading([iTextSharp.text.Document]$Document, [string[]]$Dataset, [int32]$Cols = 3, [Switch]$Centered, [int32]$FontSize = 14, [string]$Color = "BLUE")
{
    $t = New-Object iTextSharp.text.pdf.PDFPTable($Cols)
    $t.SpacingBefore = 0
    $t.SpacingAfter = -5
  


    if(!$Centered) { $t.HorizontalAlignment = 0 }
    foreach($data in $Dataset)
    {
        $t.AddCell($data);
    }
    $Document.Add($t)
}

function Get-WindowsVersion{
[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true,
                ValueFromPipeline=$true
                )]
    [string[]]$ComputerName = $env:COMPUTERNAME
)


Begin
{
    $Table = New-Object System.Data.DataTable
    $Table.Columns.AddRange(@("ComputerName","WindowsEdition","Version","OSBuild"))
}
Process
{
    Foreach ($Computer in $ComputerName)
    {
        $Code = {
            $ProductName = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name ProductName).ProductName
            Try
            {
                $Version = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name ReleaseID –ErrorAction Stop).ReleaseID
            }
            Catch
            {
                $Version = "N/A"
            }
            $CurrentBuild = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name CurrentBuild).CurrentBuild
            $UBR = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name UBR).UBR
            $OSVersion = $CurrentBuild + "." + $UBR

            $TempTable = New-Object System.Data.DataTable
            $TempTable.Columns.AddRange(@("ComputerName","WindowsEdition","Version","OSBuild"))
            [void]$TempTable.Rows.Add($env:COMPUTERNAME,$ProductName,$Version,$OSVersion)
        
            Return $TempTable
        }

        If ($Computer -eq $env:COMPUTERNAME)
        {
            $Result = Invoke-Command –ScriptBlock $Code
            [void]$Table.Rows.Add($Result.Computername,$Result.'WindowsEdition',$Result.Version,$Result.'OSBuild')
        }
        Else
        {
            Try
            {
                $Result = Invoke-Command –ComputerName $Computer –ScriptBlock $Code –ErrorAction Stop
                [void]$Table.Rows.Add($Result.Computername,$Result.'WindowsEdition',$Result.Version,$Result.'OSBuild')
            }
            Catch
            {
                $_
            }
        }

    }

}
End
{
    Return $Table
}
}

Function Get-MSHotfix  
{  
    $outputs = Invoke-Expression "wmic qfe list"  
    $outputs = $outputs[1..($outputs.length)]  
      
      
    foreach ($output in $Outputs) {  
        if ($output) {  
            $output = $output -replace 'Security Update','Security-Update'  
            $output = $output -replace 'NT AUTHORITY','NT-AUTHORITY'  
            $output = $output -replace '\s+',' '  
            $parts = $output -split ' ' 
            if ($parts[5] -like "*/*/*") {  
                $Dateis = [datetime]::ParseExact($parts[5], '%M/%d/yyyy',[Globalization.cultureinfo]::GetCultureInfo("en-US").DateTimeFormat)  
            } else {  
                $Dateis = get-date([DateTime][Convert]::ToInt64("$parts[5]", 16)) -Format '%M/%d/yyyy'  
            }  
            New-Object -Type PSObject -Property @{  
                KBArticle = [string]$parts[0]  
                Computername = [string]$parts[1]  
                Description = [string]$parts[2]  
                FixComments = [string]$parts[6]  
                HotFixID = [string]$parts[3]  
                InstalledOn = Get-Date($Dateis)-format "dddd d MMMM yyyy"  
                InstalledBy = [string]$parts[4]  
                InstallDate = [string]$parts[7]  
                Name = [string]$parts[8]  
                ServicePackInEffect = [string]$parts[9]  
                Status = [string]$parts[10]  
            }  
        }  
    }  
} 
Function Get-processor {
[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$false,
                ValueFromPipelineByPropertyName=$true,
                ValueFromPipeline=$true
                )]
    [string[]]$ComputerName = $env:COMPUTERNAME
)
Begin
{
    $Table = New-Object System.Data.DataTable
    $Table.Columns.AddRange(@("ComputerName","name","numberofcores","capacity"))
}
Process
{
    Foreach ($Computer in $ComputerName)
    {
        $Code = {
            $processor = ((Get-WmiObject -Class Win32_Processor).name | Out-String).Trim()
            $numberofcores=((Get-WmiObject -Class Win32_Processor).numberofcores| Out-String).Trim()
            $memory = ((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum)).sum

            $TempTable = New-Object System.Data.DataTable
            $TempTable.Columns.AddRange(@("ComputerName","name","numberofcores","capacity"))
            [void]$TempTable.Rows.Add($env:COMPUTERNAME, $processor,$numberofcores,$memory)
        
            Return $TempTable
        }

        If ($Computer -eq $env:COMPUTERNAME)
        {
            $Result = Invoke-Command –ScriptBlock $Code
            [void]$Table.Rows.Add($Result.Computername,$Result.'name',$Result.numberofcores,$Result.'capacity')
        }
        Else
        {
            Try
            {
                $Result = Invoke-Command –ComputerName $Computer –ScriptBlock $Code –ErrorAction Stop
                [void]$Table.Rows.Add($Result.Computername,$Result.'name',$Result.numberofcores,$Result.'capacity')
            }
            Catch
            {
                $_
            }
        }

    }

}
End
{
    Return $Table
}
}

if($keyword1.Text -ne "$null" -and $keyword2.Text -ne "$null" -and $keyword3.Text -ne "$null") { $computers = get-adcomputer -Filter {Name -like $keyword1.Text  -or Name -like $keyword2.Text -or Name -like $keyword3.Text} | Select-Object name }
elseif ($keyword1.Text -ne "$null" -and $keyword2.Text -eq "$null" -and $keyword3.Text -eq "$null") { $computers = get-adcomputer -Filter {Name -like $keyword1.Text} | Select-Object name}
elseif ($keyword1.Text -ne "$null" -and $keyword2.Text -ne "$null" -and $keyword3.Text -eq "$null") { $computers = get-adcomputer -Filter {Name -like $keyword1.Text -or Name -like $keyword2.Text} | Select-Object name }
elseif ($keyword1.Text -ne "$null" -and $keyword2.Text -eq "$null" -and $keyword3.Text -ne "$null") { $computers = get-adcomputer -Filter {Name -like $keyword1.Text -or Name -like $keyword3.Text} | Select-Object name }
elseif ($keyword1.Text -eq "$null" -and $keyword2.Text -ne "$null" -and $keyword3.Text -ne "$null") { $computers = get-adcomputer -Filter {Name -like $keyword2.Text -or Name -like $keyword3.Text} | Select-Object name }
elseif ($keyword1.Text -eq "$null" -and $keyword2.Text -ne "$null" -and $keyword3.Text -eq "$null") { $computers = get-adcomputer -Filter {Name -like $keyword2.Text} | Select-Object name }
elseif ($keyword1.Text -eq "$null" -and $keyword2.Text -eq "$null" -and $keyword3.Text -ne "$null") { $computers = get-adcomputer -Filter {Name -like $keyword3.Text} | Select-Object name }

$i=1

$exportpathname=$exportpath.Text
$exportpathnameparent=split-path -path $exportpathname
$testpath=Test-Path -Path $exportpathnameparent
#creates pdf
$pdf = New-Object iTextSharp.text.Document
Create-PDF -Document $pdf -File $exportpath.text -TopMargin 20 -BottomMargin 20 -LeftMargin 15 -RightMargin 15 -Author "PawanRijal"

if ($testpath -eq $true -and $keyword1.Text -ne "$null" -or $keyword2.Text -ne "$null" -or $keyword3.Text -ne "$null"  ) {
foreach ($computer in $computers ){
$pdf.Open()
$computername=$computer.name

write-progress -id 1 -activity "Exporting Build Report to $exportpathname" -status "Working on $computername" -percentComplete ($i++ / $computers.count * 100)

$svrinfo=@()
Add-TableTitle -Document $pdf -Dataset @("Workstation Information") -Cols 1 -Centered -UsegrayBG
Add-TableHeading -Document $pdf -Dataset @("Name", "Description") -Cols 2 -Centered
Get-ADComputer -Identity $computer.name -Properties description | foreach { $svrinfo += $_.dnshostname; $svrinfo += "" + $_.description }
Add-Table -Document $pdf -Dataset $svrinfo -Cols 2 -Centered

$osinfo=@()
Add-TableTitle -Document $pdf -Dataset @("Operating System Information") -Cols 1 -Centered 
Add-TableHeading -Document $pdf -Dataset @("OS Edition", "OS Version", "OSBuild") -Cols 3 -Centered
Get-WindowsVersion -computername $computer.name | foreach { $osinfo += $_.WindowsEdition; $osinfo += "" + $_.Version ; $osinfo += "" + $_.OSBuild }
Add-Table -Document $pdf -Dataset $osinfo -Cols 3 -Centered

$RAMinfo=@()
Add-TableTitle -Document $pdf -Dataset @("CPU/RAM Information") -Cols 1 -Centered 
Add-TableHeading -Document $pdf -Dataset @("Name", "Cores", "RAM (GBs)") -Cols 3 -Centered

Get-processor -ComputerName $computer.name | foreach { $raminfo += $_.name; $raminfo += "" + $_.numberofcores; $raminfo += $([math]::round(($_.capacity /1gb),2)) }
Add-Table -Document $pdf -Dataset $RAMinfo -Cols 3 -Centered

$netinfo=@()
Add-TableTitle -Document $pdf -Dataset @("Network Information") -Cols 1 -Centered 
Add-TableHeading -Document $pdf -Dataset @("Interface Alias", "IP Address") -Cols 2 -Centered
Get-NetIPConfiguration -CimSession $computer.name | foreach { $netinfo += $_.interfacealias; $netinfo += "" + $_.ipv4address }
Add-Table -Document $pdf -Dataset $netinfo -Cols 2 -Centered

$diskinfo=@()
Add-TableTitle -Document $pdf -Dataset @("Hard Disk Information") -Cols 1 -Centered 
Add-TableHeading -Document $pdf -Dataset @("Letter", "Label", "Size (GB)", "Free Space (GB)") -Cols 4 -Centered
Get-WmiObject -Class win32_logicaldisk -ComputerName $computer.Name | foreach { $diskinfo += $_.deviceid; $diskinfo += "" + $_.volumename; $diskinfo += $([math]::round(($_.size /1gb),2)); $diskinfo += "" + $([math]::round(($_.FreeSpace /1gb),2)) }
Add-Table -Document $pdf -Dataset $diskinfo -Cols 4 -Centered


$Apps = @()
Add-TableTitle -Document $pdf -Dataset @("Installed Software as at $(Get-Date)") -Cols 1 -Centered 
Add-TableHeading -Document $pdf -Dataset @("Name", "Version") -Cols 2 -Centered
$Appslist = @()
$Appslist += Invoke-Command -ComputerName $computer.name {Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-object {$_.DisplayName -ne $null -and $_.SystemComponent -ne "1" } | Sort-Object displayname -Descending -Unique}  # 64 Bit
$Appslist += Invoke-Command -ComputerName $computer.name {Get-ItemProperty "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-object {$_.DisplayName -ne $null -and $_.SystemComponent -ne "1" } | Sort-Object displayname -Descending -Unique} # 32 Bit
$Appslist | where  {$_.displayname -notmatch  "feature pack" -and $_.displayname -notmatch  "Microsoft Edge Update"  -and $_.displayname -notmatch  "Update for" -and $_.displayname -notmatch  "service pack" -and!([string]::IsNullOrWhiteSpace($_.Displayname))} | Sort-Object displayname | foreach { $apps += $_.displayName; $apps += "" + $_.displayVersion }
Add-Table -Document $pdf -Dataset $apps -Cols 2 -Centered


$winupdates = @()
Add-TableTitle -Document $pdf -Dataset @("Installed Windows Patches as at $(Get-Date)") -Cols 1 -Centered 
Add-TableHeading -Document $pdf -Dataset @("ID", "Description", "URL") -Cols 3 -Centered
Get-HotFix -ComputerName $computer.name  | foreach { $winupdates += $_.hotfixid; $winupdates += "" + $_.description; $winupdates += $_.caption }
Add-Table -Document $pdf -Dataset $winupdates -Cols 3 -Centered 

$pdf.NewPage()

}

write-progress -id 1 -activity "Building Report" -status "Ready" -Completed
$pdf.Close()
showresult
}
})

$win.showdialog()



