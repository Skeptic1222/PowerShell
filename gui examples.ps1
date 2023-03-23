# Load the PresentationFramework assembly to use WPF
Add-Type -AssemblyName PresentationFramework

Function CreateButtonsSection { 
@"
    <!-- Section: Buttons -->
    <Label Grid.Column="0" Grid.Row="0" Content="Buttons" FontWeight="Bold" FontSize="18" />

    <!-- Basic Button -->
    <Button Grid.Column="1" Grid.Row="0" Content="Basic Button" Margin="5" />

    <!-- Button with custom style -->
    <Button Grid.Column="2" Grid.Row="0" Content="Styled Button" Margin="5">
        <Button.Style>
            <Style TargetType="{x:Type Button}">
                <Setter Property="Background" Value="LightBlue" />
                <Setter Property="BorderBrush" Value="Blue" />
                <Setter Property="BorderThickness" Value="2" />
                <Setter Property="Foreground" Value="White" />
                <Setter Property="FontSize" Value="14" />
                <Setter Property="Padding" Value="8" />
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type Button}">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}" />
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Button.Style>
    </Button>

    <!-- Button with mouse enter and leave effects -->
    <Button Grid.Column="1" Grid.Row="1" Content="Hover Button" Margin="5">
        <Button.Style>
            <Style TargetType="{x:Type Button}">
                <Setter Property="Background" Value="LightGray" />
                <Setter Property="FontSize" Value="12" />
                <Style.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                        <Setter Property="Background" Value="Orange" />
                        <Setter Property="Foreground" Value="White" />
                    </Trigger>
                </Style.Triggers>
            </Style>
        </Button.Style>
    </Button>

    <!-- Button with click event -->
    <Button Grid.Column="2" Grid.Row="1" Content="Click Me" Margin="5">
        <Button.Triggers>
            <EventTrigger RoutedEvent="Button.Click">
                <BeginStoryboard>
                    <Storyboard>
                        <ColorAnimation Storyboard.TargetProperty="(Button.Background).(SolidColorBrush.Color)" To="Green" Duration="0:0:0.3" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Button.Triggers>
    </Button>
"@
}


Function CreateDropDownsSection {
    @"
    <!-- Section: DropDowns -->
    <Label Grid.Column="0" Grid.Row="1" Content="DropDowns" FontWeight="Bold" FontSize="18" />

    <!-- Basic ComboBox -->
    <ComboBox Grid.Column="1" Grid.Row="1" Margin="5">
        <ComboBoxItem Content="Option 1" />
        <ComboBoxItem Content="Option 2" />
        <ComboBoxItem Content="Option 3" />
    </ComboBox>

    <!-- ComboBox with custom style -->
    <ComboBox Grid.Column="2" Grid.Row="1" Margin="5" >
        <ComboBox.ItemContainerStyle>
            <Style TargetType="{x:Type ComboBoxItem}">
                <Setter Property="Background" Value="LightGray" />
                <Setter Property="Margin" Value="2" />
            </Style>
        </ComboBox.ItemContainerStyle>
        <ComboBoxItem Content="Styled Option 1" />
        <ComboBoxItem Content="Styled Option 2" />
        <ComboBoxItem Content="Styled Option 3" />
    </ComboBox>

    <!-- ComboBox with selected item event -->
    <ComboBox Grid.Column="1" Grid.Row="2" Margin="5">
        <ComboBox.Triggers>
            <EventTrigger RoutedEvent="Selector.SelectionChanged">
                <BeginStoryboard>
                    <Storyboard>
                        <ColorAnimation Storyboard.TargetProperty="(ComboBox.Foreground).(SolidColorBrush.Color)" To="Green" Duration="0:0:0.3" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </ComboBox.Triggers>
        <ComboBoxItem Content="Select Me 1" />
        <ComboBoxItem Content="Select Me 2" />
        <ComboBoxItem Content="Select Me 3" />
    </ComboBox>
"@
}


Function CreateTogglesSection {
    @"
    <!-- Section: Toggles -->
    <Label Grid.Column="0" Grid.Row="2" Content="Toggles" FontWeight="Bold" FontSize="18" />

    <!-- Basic CheckBox -->
    <CheckBox Grid.Column="1" Grid.Row="2" Content="Basic CheckBox" Margin="5" />

    <!-- CheckBox with custom style -->
    <CheckBox Grid.Column="2" Grid.Row="2" Content="Styled CheckBox" Margin="5">
        <CheckBox.Style>
            <Style TargetType="{x:Type CheckBox}">
                <Setter Property="FontSize" Value="14" />
                <Setter Property="Foreground" Value="Blue" />
            </Style>
        </CheckBox.Style>
    </CheckBox>

    <!-- Basic RadioButton -->
    <RadioButton Grid.Column="1" Grid.Row="3" Content="Basic RadioButton" Margin="5" />

    <!-- RadioButton with custom style -->
    <RadioButton Grid.Column="2" Grid.Row="3" Content="Styled RadioButton" Margin="5">
        <RadioButton.Style>
            <Style TargetType="{x:Type RadioButton}">
                <Setter Property="FontSize" Value="14" />
                <Setter Property="Foreground" Value="Blue" />
            </Style>
        </RadioButton.Style>
    </RadioButton>
"@
}


Function CreateProgressBarsSection {
    @"
    <!-- Section: Progress Bars -->
    <Label Grid.Column="0" Grid.Row="3" Content="Progress Bars" FontWeight="Bold" FontSize="18" />

    <!-- Basic ProgressBar -->
    <ProgressBar Grid.Column="1" Grid.Row="3" Margin="5" Width="100" Height="20" Value="50" />

    <!-- ProgressBar with custom style -->
    <ProgressBar Grid.Column="2" Grid.Row="3" Margin="5" Width="100" Height="20">
        <ProgressBar.Style>
            <Style TargetType="{x:Type ProgressBar}">
                <Setter Property="Value" Value="50" />
                <Setter Property="Foreground" Value="Orange" />
            </Style>
        </ProgressBar.Style>
    </ProgressBar>

    <!-- Indeterminate ProgressBar -->
    <ProgressBar Grid.Column="1" Grid.Row="4" Margin="5" Width="100" Height="20" IsIndeterminate="True" />

    <!-- ProgressBar with custom style and animation -->
    <ProgressBar Grid.Column="2" Grid.Row="4" Margin="5" Width="100" Height="20" Name="AnimatedProgressBar">
        <ProgressBar.Style>
            <Style TargetType="{x:Type ProgressBar}">
                <Setter Property="Foreground" Value="Purple" />
            </Style>
        </ProgressBar.Style>
        <ProgressBar.Triggers>
            <EventTrigger RoutedEvent="FrameworkElement.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetProperty="Value" From="0" To="100" Duration="0:0:5" RepeatBehavior="Forever" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </ProgressBar.Triggers>
    </ProgressBar>
"@
}


Function CreateShapesSection {
    @"
    <!-- Section: Shapes -->
    <Label Grid.Column="0" Grid.Row="4" Content="Shapes" FontWeight="Bold" FontSize="18" />

    <!-- Basic Rectangle -->
    <Rectangle Grid.Column="1" Grid.Row="4" Margin="5" Width="50" Height="30" Fill="LightGray" />

    <!-- Rounded Rectangle -->
    <Rectangle Grid.Column="2" Grid.Row="4" Margin="5" Width="50" Height="30" Fill="LightBlue" RadiusX="10" RadiusY="10" />

    <!-- Basic Ellipse -->
    <Ellipse Grid.Column="1" Grid.Row="5" Margin="5" Width="50" Height="30" Fill="LightGreen" />

    <!-- Ellipse with custom stroke -->
    <Ellipse Grid.Column="2" Grid.Row="5" Margin="5" Width="50" Height="30" Fill="LightYellow" Stroke="Purple" StrokeThickness="2" />
"@
}


Function CreateNestedElementsSection {
    @"
    <!-- Section: Nested Elements -->
    <Label Grid.Column="0" Grid.Row="5" Content="Nested Elements" FontWeight="Bold" FontSize="18" />

    <!-- Basic Expander -->
    <Expander Grid.Column="1" Grid.Row="5" Header="Basic Expander" Margin="5">
        <StackPanel>
            <TextBlock Text="This is a basic expander." />
            <Button Content="Click Me!" />
        </StackPanel>
    </Expander>

    <!-- Customized Expander -->
    <Expander Grid.Column="2" Grid.Row="5" Header="Styled Expander" Margin="5">
        <Expander.Style>
            <Style TargetType="{x:Type Expander}">
                <Setter Property="HeaderTemplate">
                    <Setter.Value>
                        <DataTemplate>
                            <TextBlock Text="{Binding}" FontWeight="Bold" Foreground="Blue" />
                        </DataTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Expander.Style>
        <StackPanel>
            <TextBlock Text="This is a styled expander." />
            <Button Content="Click Me!" />
        </StackPanel>
    </Expander>

    <!-- Basic TabControl -->
    <TabControl Grid.Column="1" Grid.Row="6" Margin="5">
        <TabItem Header="Tab 1">
            <TextBlock Text="Content for Tab 1" />
        </TabItem>
        <TabItem Header="Tab 2">
            <TextBlock Text="Content for Tab 2" />
        </TabItem>
    </TabControl>

    <!-- Customized TabControl -->
    <TabControl Grid.Column="2" Grid.Row="6" Margin="5">
        <TabControl.Resources>
            <Style TargetType="{x:Type TabItem}">
                <Setter Property="HeaderTemplate">
                    <Setter.Value>
                        <DataTemplate>
                            <TextBlock Text="{Binding}" FontWeight="Bold" Foreground="Green" />
                        </DataTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </TabControl.Resources>
        <TabItem Header="Styled Tab 1">
            <TextBlock Text="Content for Styled Tab 1" />
        </TabItem>
        <TabItem Header="Styled Tab 2">
            <TextBlock Text="Content for Styled Tab 2" />
        </TabItem>
    </TabControl>
"@
}


[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Demo PowerShell GUI" Width="1200" Height="800" Background="White">
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <!-- Insert Sections Here -->
        $(CreateButtonsSection)
        $(CreateDropDownsSection)
        $(CreateTogglesSection)
        $(CreateProgressBarsSection)
        $(CreateShapesSection)
        $(CreateNestedElementsSection)
    </Grid>
</Window>
"@

Add-Type -AssemblyName PresentationFramework

$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$Window.ShowDialog() | Out-Null

