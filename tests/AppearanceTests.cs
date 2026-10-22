using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Word_Compare;

class AppearanceTests {
 static List<string> results=new List<string>();
 static T Field<T>(MainWindow w,string name){return (T)typeof(MainWindow).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(w);}
 static void Check(bool ok,string name){if(!ok)throw new Exception("FAILED "+name);results.Add("PASS "+name);Console.WriteLine("PASS "+name);}
 static void Pump(){var frame=new DispatcherFrame();Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,new Action(()=>frame.Continue=false));Dispatcher.PushFrame(frame);}
 static void Wait(Task task){var start=DateTime.UtcNow;while(!task.IsCompleted){Pump();Thread.Sleep(10);if((DateTime.UtcNow-start).TotalSeconds>20)throw new TimeoutException("UI comparison timeout");}task.GetAwaiter().GetResult();Pump();}
 static void Click(ToggleButton button,bool value){button.IsChecked=value;button.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));Pump();}
 static string ColorOf(System.Windows.Media.Brush b){var s=(SolidColorBrush)b;return "#"+s.Color.R.ToString("X2")+s.Color.G.ToString("X2")+s.Color.B.ToString("X2");}
 static string ToggleBackground(ToggleButton b){b.ApplyTemplate();return ColorOf(((Border)b.Template.FindName("ToggleBorder",b)).Background);}
 static void Capture(Window w,string file){w.UpdateLayout();var b=new RenderTargetBitmap((int)w.ActualWidth,(int)w.ActualHeight,96,96,PixelFormats.Pbgra32);b.Render(w);var png=new PngBitmapEncoder();png.Frames.Add(BitmapFrame.Create(b));using(var f=File.Create(file))png.Save(f);}
 static T Descendant<T>(DependencyObject parent) where T:DependencyObject {for(int i=0;i<VisualTreeHelper.GetChildrenCount(parent);i++){var child=VisualTreeHelper.GetChild(parent,i);if(child is T)return (T)child;var found=Descendant<T>(child);if(found!=null)return found;}return null;}
 [STAThread] static int Main(string[] args){MainWindow window=null;var app=new Application();SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext(Dispatcher.CurrentDispatcher));app.ShutdownMode=ShutdownMode.OnExplicitShutdown;string root=args[0],outdir=args[1];Directory.CreateDirectory(outdir);try{
  window=new MainWindow(new string[0]);window.Show();Pump();
  var dark=Field<ToggleButton>(window,"dark");var preserve=Field<ToggleButton>(window,"preserve");var panel=Field<ToggleButton>(window,"panelToggle");var only=Field<ToggleButton>(window,"only");var zoom=Field<ComboBox>(window,"zoom");
  Check(window.Title=="Word_Compare","窗口标题统一为Word_Compare");
  Check(new[]{dark,preserve,panel,only}.All(b=>b.GetType()==typeof(ToggleButton)),"四个工具栏开关为按钮而非复选框");
  Console.WriteLine("CAPTION="+Appearance.RequestedCaptionColor+" SET="+Appearance.CaptionResult+" TEXT="+Appearance.TextResult);
  Check(Appearance.RequestedCaptionColor=="#F5F7FA"&&Appearance.CaptionResult==0,"浅色原生标题栏颜色设置");
  Check(ToggleBackground(preserve)=="#087F72"&&ColorOf(preserve.Foreground)=="#FFFFFF","浅色选中按钮绿色底白字");
  Field<TextBox>(window,"leftPath").Text=Path.Combine(root,"示例","示例_原版.docx");Field<TextBox>(window,"rightPath").Text=Path.Combine(root,"示例","示例_修订版.docx");Wait(window.Run());int changes=window.Result.Changes;
  Check(changes==9,"对比结果保持9项差异");
  Click(preserve,false);Check(!window.Preserve&&ToggleBackground(preserve)=="#FFFFFF","保留格式按钮可关闭并取消选中底色");Click(preserve,true);
  Click(only,true);Check(Field<ListBox>(window,"list").Items.Count==changes,"仅显示差异按钮筛选正确");Click(only,false);Check(Field<ListBox>(window,"list").Items.Count==22,"关闭筛选恢复全部内容");
  Click(panel,false);Check(Field<ColumnDefinition>(window,"panelWidth").Width.Value==0&&Field<Border>(window,"sidePanel").Visibility==Visibility.Collapsed,"差异面板按钮完整隐藏面板");Click(panel,true);Check(Field<ColumnDefinition>(window,"panelWidth").Width.Value==270&&Field<Border>(window,"sidePanel").Visibility==Visibility.Visible,"差异面板按钮恢复面板");
  var type=Field<ComboBox>(window,"typeFilter");var state=Field<ComboBox>(window,"stateFilter");var order=Field<ComboBox>(window,"sortOrder");var highlight=Field<ComboBox>(window,"highlight");var list=Field<ListBox>(window,"list");var changeList=Field<ListBox>(window,"changes");
  window.WindowState=WindowState.Normal;window.Width=1100;window.Height=700;Pump();window.UpdateLayout();var rowView=Descendant<RowView>(list);Check(rowView!=null&&rowView.ActualWidth<=list.ActualWidth&&Math.Abs(rowView.ColumnDefinitions[0].ActualWidth-rowView.ColumnDefinitions[2].ActualWidth)<1,"最小窗口宽度下双栏等宽且不超出视口");
  window.Width=1450;Pump();window.UpdateLayout();Check(rowView.ActualWidth<=list.ActualWidth&&Math.Abs(rowView.ActualWidth-(list.ActualWidth-19))<5,"窗口变宽时行宽立即跟随视口");window.WindowState=WindowState.Maximized;Pump();
  type.SelectedIndex=3;Pump();Check(list.Items.Count==4&&changeList.Items.Count==4,"类型筛选同步双栏与差异面板");
  typeof(MainWindow).GetMethod("Navigate",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(window,new object[]{1});Pump();Check(((DisplayRow)list.SelectedItem).Row.Kind==ChangeKind.Modified,"差异导航遵守当前筛选");
  var reviewed=((DisplayRow)list.SelectedItem).Row;typeof(MainWindow).GetMethod("MarkReview",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(window,new object[]{ReviewState.Accepted});Pump();Check(reviewed.Review==ReviewState.Accepted&&Field<TextBlock>(window,"reviewSummary").Text.Contains("尚未保存"),"实际审核按钮动作更新状态及保存提醒");
  state.SelectedIndex=2;Pump();Check(list.Items.Count==1&&changeList.Items.Count==1,"按已采纳审核状态筛选");
  typeof(MainWindow).GetMethod("MarkReview",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(window,new object[]{ReviewState.Rejected});Pump();Check(list.Items.Count==0&&changeList.Items.Count==0,"更新状态后不匹配项目立即退出筛选视图");
  state.SelectedIndex=3;Pump();Check(list.Items.Count==1&&((DisplayRow)list.Items[0]).Row.Review==ReviewState.Rejected,"不采纳状态筛选可恢复定位");
  type.SelectedIndex=0;state.SelectedIndex=0;order.SelectedIndex=2;Pump();var sorted=changeList.Items.Cast<DiffRow>().ToList();Func<DiffRow,int> len=r=>Math.Max(r.Left==null?0:r.Left.Text.Length,r.Right==null?0:r.Right.Text.Length);Check(sorted.Select(len).SequenceEqual(sorted.Select(len).OrderByDescending(n=>n))&&list.Items.Cast<DisplayRow>().Select(d=>d.Row.Index).SequenceEqual(Enumerable.Range(0,22)),"长度排序仅调整差异面板且双栏保持文档顺序");
  highlight.SelectedIndex=2;Pump();Check(window.Result.Options.Highlight==HighlightLevel.Sentence&&window.Result.Changes==changes&&reviewed.Review==ReviewState.Rejected,"粒度选择即时更新且保留审核意见");highlight.SelectedIndex=0;
  ReviewStore.Save(window.Result,Path.Combine(outdir,"UI审核记录.json"));typeof(MainWindow).GetField("reviewDirty",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(window,false);window.Refresh();order.SelectedIndex=0;Pump();
  Capture(window,Path.Combine(outdir,"按钮版_浅色.png"));
  Click(dark,true);Check(window.Dark&&Appearance.RequestedCaptionColor=="#11161D"&&Appearance.CaptionResult==0&&Appearance.TextResult==0,"暗黑原生标题栏请求#11161D且DWM调用成功");
  Check(ToggleBackground(dark)=="#176F65"&&ToggleBackground(preserve)=="#176F65"&&ColorOf(dark.Foreground)=="#FFFFFF","暗黑选中按钮绿色底白字");
  Check(new[]{zoom,type,state,order,highlight}.All(c=>ColorOf(c.Background)=="#191E26"&&ColorOf(c.Foreground)=="#E3E9F2"),"暗黑缩放及新增筛选粒度控件配色正确");
  zoom.IsDropDownOpen=true;Pump();var popup=(Popup)zoom.Template.FindName("PART_Popup",zoom);Check(popup.IsOpen&&ColorOf(((Border)popup.Child).Background)=="#191E26","暗黑缩放弹出列表配色正确");zoom.IsDropDownOpen=false;zoom.SelectedIndex=2;Pump();Check(Math.Abs(window.Zoom-1.2)<0.001,"自定义缩放框仍可切换120%");zoom.SelectedIndex=1;
  Capture(window,Path.Combine(outdir,"按钮版_暗黑.png"));
  for(int i=0;i<3;i++){Click(dark,false);Check(Appearance.RequestedCaptionColor=="#F5F7FA","明暗往返浅色标题栏设置第"+(i+1)+"次");Click(dark,true);Check(Appearance.RequestedCaptionColor=="#11161D","明暗往返暗色标题栏设置第"+(i+1)+"次");}
  Check(window.Result.Changes==changes,"主题及显示按钮不改变对比结果");
  File.WriteAllLines(Path.Combine(outdir,"界面按钮与标题栏测试.txt"),results.Concat(new[]{"TOTAL "+results.Count+" PASSED","Caption="+Appearance.RequestedCaptionColor}),new UTF8Encoding(true));window.Close();app.Shutdown();return 0;
 }catch(Exception e){Console.WriteLine(e);File.WriteAllText(Path.Combine(outdir,"界面测试失败.txt"),e.ToString(),new UTF8Encoding(true));if(window!=null)window.Close();app.Shutdown();return 1;}}
}
