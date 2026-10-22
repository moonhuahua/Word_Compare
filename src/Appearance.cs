using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Markup;

namespace Word_Compare {
 // Native caption colors are separate from WPF's client-area background.
 public static class Appearance {
  [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd,int attribute,ref int value,int size);
  public static int CaptionResult=-1,TextResult=-1,RequestedCaption;
  public static void ApplyCaption(Window window,bool dark){
   var handle=new WindowInteropHelper(window).Handle;if(handle==IntPtr.Zero)return;
   try{
    int enabled=dark?1:0;
    int caption=dark?0x11|(0x16<<8)|(0x1D<<16):0xF5|(0xF7<<8)|(0xFA<<16);
    int text=dark?0xE3|(0xE9<<8)|(0xF2<<16):0x26|(0x34<<8)|(0x45<<16);
    DwmSetWindowAttribute(handle,20,ref enabled,sizeof(int));
    RequestedCaption=caption;CaptionResult=DwmSetWindowAttribute(handle,35,ref caption,sizeof(int));
    TextResult=DwmSetWindowAttribute(handle,36,ref text,sizeof(int));
   }catch(DllNotFoundException){CaptionResult=-1;TextResult=-1;}catch(EntryPointNotFoundException){CaptionResult=-1;TextResult=-1;}
  }
  public static string RequestedCaptionColor {get{return string.Format("#{0:X2}{1:X2}{2:X2}",RequestedCaption&255,(RequestedCaption>>8)&255,(RequestedCaption>>16)&255);}}
  public static Style ToggleStyle(){return (Style)XamlReader.Parse(@"
   <Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='{x:Type ToggleButton}' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
    <Setter Property='Padding' Value='10,5'/><Setter Property='Margin' Value='3,1'/><Setter Property='MinHeight' Value='30'/><Setter Property='Cursor' Value='Hand'/>
    <Setter Property='Foreground' Value='{DynamicResource TextBrush}'/><Setter Property='Background' Value='{DynamicResource SurfaceBrush}'/><Setter Property='BorderBrush' Value='{DynamicResource BorderBrush}'/>
    <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='{x:Type ToggleButton}'>
     <Border x:Name='ToggleBorder' Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}' BorderThickness='1' CornerRadius='3' Padding='{TemplateBinding Padding}'>
      <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center' RecognizesAccessKey='True' TextElement.Foreground='{TemplateBinding Foreground}'/>
     </Border>
     <ControlTemplate.Triggers>
      <Trigger Property='IsChecked' Value='True'><Setter TargetName='ToggleBorder' Property='Background' Value='{DynamicResource AccentBrush}'/><Setter TargetName='ToggleBorder' Property='BorderBrush' Value='{DynamicResource AccentBrush}'/><Setter Property='Foreground' Value='White'/></Trigger>
      <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='ToggleBorder' Property='BorderBrush' Value='{DynamicResource RightBrush}'/></Trigger>
      <Trigger Property='IsKeyboardFocused' Value='True'><Setter TargetName='ToggleBorder' Property='BorderBrush' Value='{DynamicResource RightBrush}'/></Trigger>
      <Trigger Property='IsPressed' Value='True'><Setter TargetName='ToggleBorder' Property='Opacity' Value='0.8'/></Trigger>
      <Trigger Property='IsEnabled' Value='False'><Setter Property='Opacity' Value='0.5'/></Trigger>
     </ControlTemplate.Triggers>
    </ControlTemplate></Setter.Value></Setter>
   </Style>");}
  public static Style ZoomStyle(){return (Style)XamlReader.Parse(@"
   <Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='{x:Type ComboBox}' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
    <Setter Property='Foreground' Value='{DynamicResource TextBrush}'/><Setter Property='Background' Value='{DynamicResource SurfaceBrush}'/><Setter Property='BorderBrush' Value='{DynamicResource BorderBrush}'/><Setter Property='MinHeight' Value='30'/>
    <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='{x:Type ComboBox}'>
     <Grid>
      <ToggleButton Focusable='False' Margin='0' Padding='0' MinHeight='0' Background='{TemplateBinding Background}' Foreground='{TemplateBinding Foreground}' BorderBrush='{TemplateBinding BorderBrush}' IsChecked='{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}'>
       <ToggleButton.Template><ControlTemplate TargetType='{x:Type ToggleButton}'><Border Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}' BorderThickness='1' CornerRadius='3'><TextBlock Text='▾' HorizontalAlignment='Right' VerticalAlignment='Center' Margin='0,0,7,0' Foreground='{TemplateBinding Foreground}'/></Border></ControlTemplate></ToggleButton.Template>
      </ToggleButton>
      <ContentPresenter Margin='8,3,25,3' VerticalAlignment='Center' IsHitTestVisible='False' Content='{TemplateBinding SelectionBoxItem}' ContentTemplate='{TemplateBinding SelectionBoxItemTemplate}' TextElement.Foreground='{TemplateBinding Foreground}'/>
      <Popup x:Name='PART_Popup' Placement='Bottom' IsOpen='{TemplateBinding IsDropDownOpen}' AllowsTransparency='True' Focusable='False' PopupAnimation='None'>
       <Border MinWidth='{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}' MaxHeight='{TemplateBinding MaxDropDownHeight}' Background='{DynamicResource SurfaceBrush}' BorderBrush='{DynamicResource BorderBrush}' BorderThickness='1'><ScrollViewer><ItemsPresenter KeyboardNavigation.DirectionalNavigation='Contained'/></ScrollViewer></Border>
      </Popup>
     </Grid>
    </ControlTemplate></Setter.Value></Setter>
   </Style>");}
  public static Style ZoomItemStyle(){return (Style)XamlReader.Parse(@"
   <Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='{x:Type ComboBoxItem}' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
    <Setter Property='Foreground' Value='{DynamicResource TextBrush}'/><Setter Property='Padding' Value='8,5'/><Setter Property='HorizontalContentAlignment' Value='Stretch'/>
    <Setter Property='Template'><Setter.Value><ControlTemplate TargetType='{x:Type ComboBoxItem}'><Border x:Name='ItemBorder' Background='{DynamicResource SurfaceBrush}' Padding='{TemplateBinding Padding}'><ContentPresenter TextElement.Foreground='{TemplateBinding Foreground}'/></Border><ControlTemplate.Triggers><Trigger Property='IsHighlighted' Value='True'><Setter TargetName='ItemBorder' Property='Background' Value='{DynamicResource HoverBrush}'/></Trigger><Trigger Property='IsSelected' Value='True'><Setter TargetName='ItemBorder' Property='Background' Value='{DynamicResource AccentBrush}'/><Setter Property='Foreground' Value='White'/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
   </Style>");}
 }
}
