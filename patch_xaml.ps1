$path = "Passage\Passage.App\MainWindow.xaml"
$content = Get-Content $path -Raw

# Replace the IsEditing trigger first to include DescriptionPlaceholder
$oldIsEditing = '<DataTrigger Binding="{Binding Path=(app:MainWindow.IsEditing), RelativeSource={RelativeSource Self}}" Value="True">
          <Setter TargetName="HeadingTextBlock" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="InlineHeadingTextBox" Property="Visibility" Value="Visible" />
          <Setter TargetName="DescriptionTextBlock" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="InlineDescriptionTextBox" Property="Visibility" Value="Visible" />
          <Setter TargetName="EditIconPath" Property="Style" Value="{StaticResource BeatBoardCommitCheckIconStyle}" />
          <Setter TargetName="EditButton" Property="ToolTip" Value="Save changes" />
        </DataTrigger>'

$newIsEditing = '<DataTrigger Binding="{Binding Path=(app:MainWindow.IsEditing), RelativeSource={RelativeSource Self}}" Value="True">
          <Setter TargetName="HeadingTextBlock" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="InlineHeadingTextBox" Property="Visibility" Value="Visible" />
          <Setter TargetName="DescriptionTextBlock" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="DescriptionPlaceholder" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="InlineDescriptionTextBox" Property="Visibility" Value="Visible" />
          <Setter TargetName="EditIconPath" Property="Style" Value="{StaticResource BeatBoardCommitCheckIconStyle}" />
          <Setter TargetName="EditButton" Property="ToolTip" Value="Save changes" />
        </DataTrigger>'

$content = $content.Replace($oldIsEditing, $newIsEditing)

# Add the IsCollapsed and BoardDescription triggers
$oldTriggersEnd = '          <Setter TargetName="EditButton" Property="ToolTip" Value="Save changes" />
        </DataTrigger>
      </DataTemplate.Triggers>'

$newTriggersEnd = '          <Setter TargetName="EditButton" Property="ToolTip" Value="Save changes" />
        </DataTrigger>
        <DataTrigger Binding="{Binding IsCollapsed}" Value="True">
          <Setter TargetName="DescriptionTextBlock" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="DescriptionPlaceholder" Property="Visibility" Value="Collapsed" />
        </DataTrigger>
        <DataTrigger Binding="{Binding BoardDescription}" Value="">
          <Setter TargetName="DescriptionTextBlock" Property="Visibility" Value="Collapsed" />
          <Setter TargetName="DescriptionPlaceholder" Property="Visibility" Value="Visible" />
        </DataTrigger>
      </DataTemplate.Triggers>'

$content = $content.Replace($oldTriggersEnd, $newTriggersEnd)

Set-Content $path $content
Write-Host "Successfully patched MainWindow.xaml with triggers."
