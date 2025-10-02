#requires -version 7
#requires -modules ActiveDirectory

using namespace System.Threading
using namespace System.Reflection
using namespace System.Windows.Media
using namespace System.Windows.Markup
using namespace System.Windows.Shapes
using namespace System.Windows.Controls
using namespace System.Windows.Threading
using namespace Microsoft.ActiveDirectory.Management
using namespace System.Management.Automation.Runspaces

param(
   [Parameter()]
   [ValidateNotNull()]
   [String[]]$Domain = ('factory.acme.org', 'office.acme.org'),

   [Parameter()]
   [ValidateNotNullOrEmpty()]
   [String]$GroupName = 'Radmin',

   [Parameter(DontShow)]
   [ValidateNotNullOrEmpty()]
   [UInt16]$PingTimeout = 500,

   [Parameter(DontShow)]
   [ValidateNotNullOrEmpty()]
   [UInt16]$TimerInterval = 100,

   [Parameter(DontShow)]
   [ValidateNotNullOrEmpty()]
   [UInt16]$Threads = 20
)

Add-Type -AssemblyName PresentationFramework
$GetDNComponents = [DirectoryServices.SortOption].Assembly.GetType(
   'System.DirectoryServices.ActiveDirectory.Utils'
).GetMethod('GetDNComponents', [BindingFlags]'NonPublic, Static')
$pingJobs, $treeItems, $allComputers, $nodeComputers = @{}, @(), @{}, @{}

$window = [XamlReader]::Load([Xml.XmlNodeReader]::new([xml]@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"
        Title="ADPinger" Height="430" Width="600"
        WindowStartupLocation="CenterScreen">
   <Grid Background="#f8f8f8">
      <Grid.RowDefinitions>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <TextBox Name="FilterTextBox" Grid.Row="0" Margin="5" Padding="3"
               FontSize="14" VerticalContentAlignment="Center">
         <TextBox.Style>
            <Style TargetType="TextBox">
               <Setter Property="Foreground" Value="Black"/>
               <Setter Property="Template">
                  <Setter.Value>
                     <ControlTemplate TargetType="TextBox">
                        <Border x:Name="FilterTextBoxBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                SnapsToDevicePixels="True">
                           <Grid>
                              <ScrollViewer x:Name="PART_ContentHost"/>
                              <TextBlock x:Name="PlaceholderText"
                                         Text="Filter by host name..."
                                         Foreground="Gray" Margin="5,0,0,0"
                                         VerticalAlignment="Center"
                                         IsHitTestVisible="False"/>
                           </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                           <Trigger Property="Text" Value="{x:Null}">
                              <Setter TargetName="PlaceholderText" Property="Visibility" Value="Visible"/>
                           </Trigger>
                           <Trigger Property="IsFocused" Value="True">
                              <Setter TargetName="PlaceholderText" Property="Visibility" Value="Collapsed"/>
                           </Trigger>
                           <Trigger Property="Text" Value="{x:Static sys:String.Empty}">
                              <Setter TargetName="PlaceholderText" Property="Visibility" Value="Visible"/>
                           </Trigger>
                        </ControlTemplate.Triggers>
                     </ControlTemplate>
                  </Setter.Value>
               </Setter>
            </Style>
         </TextBox.Style>
      </TextBox>
      <TreeView Name="OuTree" Grid.Row="1" Margin="5" FontSize="14"
                Background="White" BorderThickness="1">
         <TreeView.ItemContainerStyle>
            <Style TargetType="TreeViewItem">
               <Setter Property="Background" Value="Transparent"/>
               <Setter Property="Padding" Value="2"/>
               <Style.Triggers>
                  <Trigger Property="IsMouseOver" Value="True">
                     <Setter Property="Background" Value="AliceBlue"/>
                  </Trigger>
                  <Trigger Property="IsSelected" Value="True">
                     <Setter Property="Background" Value="LightBlue"/>
                  </Trigger>
               </Style.Triggers>
            </Style>
         </TreeView.ItemContainerStyle>
         <TreeView.Resources>
            <DataTemplate x:Key="ComputerStatusTemplate">
               <StackPanel Orientation="Horizontal">
                  <Ellipse Height="12" Width="12" Margin="0,0,6,0" Fill="{Binding StatusBrush}"/>
                  <TextBlock Text="{Binding DisplayName}"/>
               </StackPanel>
            </DataTemplate>
         </TreeView.Resources>
      </TreeView>
   </Grid>
</Window>
'@))
$ouTree = $window.FindName('OuTree')
$filterTextBox = $window.FindName('FilterTextBox')

$runspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
$runspacePool.ApartmentState = [ApartmentState]::STA
$runspacePool.ThreadOptions = [PSThreadOptions]::ReuseThread
$runspacePool.Open()

#region functions
function Get-OuPath {
   param([String]$DN)
   $ous = $GetDNComponents.Invoke($null, @($DN)).Where{$_.Name -like 'ou*'}
   $ous.Count -gt 0 ? $(($ous[-1..0].Value -join '/') -replace '^(computer|server)s_') : '(OU)'
}

function New-StatusItem {
   param([String]$Text)
   New-Object TreeViewItem -Property @{
      Header = $Text
      IsEnabled = $false
   }
}

function Get-FilterText {
   $filterTextBox.Text.Trim()
}

function Start-PingJob {
   param([ADObject]$Computer, [TreeViewItem]$ParentNode)
   $powershell = [PowerShell]::Create()
   $powershell.RunspacePool = $runspacePool

   $asyncResult = $powershell.AddScript({
      param($HostName, $Timeout)
      try {
         $ping = [Net.NetworkInformation.Ping]::new()
         $result = $ping.Send($HostName, $Timeout)
         @{Success = $true; Online = ($result.Status -eq 'Success'); HostName = $HostName}
      }
      catch {
        @{Success = $false; Online = $false; HostName = $HostName}
      }
   }).AddArgument($Computer.DNSHostName).AddArgument($PingTimeout).BeginInvoke()

   $pingJobs[$asyncResult.AsyncWaitHandle] = @{
      PowerShell = $powershell
      AsyncResult = $asyncResult
      Node = $ParentNode
      Computer = $Computer
   }
}

function Ping-Node {
   param([TreeViewItem]$Node)
   if ($Node.Tag -isnot [ADComputer[]]) { return }
   $Node.Items.Clear()

   [void]$Node.Items.Add((New-StatusItem 'Ping node...'))

   foreach ($computer in $Node.Tag) {
      Start-PingJob -Computer $computer -ParentNode $Node
   }
}
#endregion functions

foreach ($group in $($Domain.ForEach{
   Get-ADObject -LDAPFilter "(&(objectClass=computer)(memberOf:1.2.840.113556.1.4.1941:= $(
      (Get-ADGroup -Filter {Name -eq $GroupName} -Server $_).DistinguishedName
   )))" -Properties DNSHostName, DistinguishedName -Server $_
} | Group-Object {Get-OuPath -DN $_.DistinguishedName})) {
   $node = New-Object TreeViewItem -Property @{
      Header = $group.Name
      Tag = $group.Group
      Background = [Brushes]::Transparent
   }

   $node.Add_Expanded({
      if ($this.Items.Count -eq 1 -and $this.Items[0] -eq 'Ping hosts...') {
         $this.Items.Clear()
         [void]$this.Items.Add((New-StatusItem 'Starting ping...'))
         $this.Tag.ForEach{Start-PingJob -Computer $_ -ParentNode $this}
      }
   })

   [void]$node.Items.Add('Ping hosts...')

   $nodeComputers[$node] = @()
   foreach ($computer in $group.Group) {
      $allComputers[$computer.DNSHostName] = @{
         Node =$node
         Computer = $computer
         Item = $null
         Online = $false
      }
      $nodeComputers[$node] += $computer.DNSHostName
   }

   [void]$ouTree.Items.Add($node)
   $treeItems += $node
}

$timer = New-Object DispatcherTimer -Property @{
   Interval = [TimeSpan]::FromMilliseconds($TimerInterval)
}
$timer.Add_Tick({
   $completedJobs = @()
   foreach ($handle in $pingJobs.Keys) {
      $job = $pingJobs[$handle]
      if (!$job.AsyncResult.IsCompleted) { continue }

      try {
         $result = $job.PowerShell.EndInvoke($job.AsyncResult)
         $allComputers[$job.Computer.DNSHostName].Online = $result.Online

         $job.Node.Dispatcher.Invoke([Action]{
            if ($job.Node.Items.Count -gt 0 -and $job.Node.Items[0].Header -like '*ping*') {
               $job.Node.Items.RemoveAt(0)
            }

            $computerItem = New-Object TreeViewItem -Property @{
               HeaderTemplate = $ouTree.Resources['ComputerStatusTemplate']
               Header = [PSCustomObject]@{
                  DisplayName = "$($job.Computer.Name) ($($job.Computer.DNSHostName))"
                  StatusBrush = $($result.Online ? [Brushes]::LimeGreen : [Brushes]::Crimson)
               }
               Tag = $job.Computer.DNSHostName
            }

            $allComputers[$job.Computer.DNSHostName].Item = $computerItem
            $filetrText = Get-FilterText
            if ([String]::IsNullOrEmpty($filterText) -or $job.Computer.DNSHostName `
               -like "*$filterText*" -or $job.Computer.Name -like "*$filterText*") {
               [void]$job.Node.Items.Add($computerItem)
            }
         })
      }
      catch {
         Write-Host "$($_.Exception.Message)" -ForegroundColor Red
      }
      finally {
         $job.PowerShell.Dispose()
         $completedJobs += $handle
      }
   }
   $completedJobs.ForEach{ $pingJobs.Remove($_) }
})
$timer.Start()

$filterTextBox.Add_TextChanged({
   $filterText = Get-FilterText
   $ouTree.Dispatcher.Invoke([Action]{
      foreach ($node in $treeItems) {
         $count = 0
         foreach ($computerDNS in $nodeComputers[$node]) {
            $computer = $allComputers[$computerDNS].Computer
            if ([String]::IsNullOrEmpty($filterText) -or $computer.DNSHostName `
               -like "*$filterText*" -or $computer.Name -like "*$filterText*") {
               $count++
            }
         }

         if ($count -gt 0) {
            if (!$ouTree.Items.Contains($node)) { [void]$ouTree.Items.Add($node) }
            foreach ($computerDNS in $nodeComputers[$node]) {
               $computerData = $allComputers[$computerDNS]
               if ($computerData.Item -eq $null) { continue }

               $show = [String]::IsNullOrEmpty($filterText) -or $computerData.Computer.DNSHostName -like `
                                     "*$filterText*" -or $computerData.Computer.Name -like "*$filterText*"
               if ($show -and !$node.Items.Contains($computerData.Item)) {
                  [void]$node.Items.Add($computerData.Item)
               }
               elseif (!$show -and $node.Items.Contains($computerData.Item)) {
                  [void]$node.Items.Remove($computerData.Item)
               }
            }
         }
         else {
            if ($ouTree.Items.Contains($node)) { [void]$ouTree.Items.Remove($node) }
         }
      }
   })
})

$window.Add_Loaded({
   $treeItems.ForEach{ $_.IsExpanded = $true }
   $filterTextBox.Focus()
})
$window.Add_Closed({
   $timer.Stop()
   $runspacePool.Dispose()
})
[void]$window.ShowDialog()
