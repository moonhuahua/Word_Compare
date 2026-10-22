using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using Word_Compare;
class EngineTests {
 static int count; static List<string> results=new List<string>();
 static void Check(bool ok,string name){if(!ok)throw new Exception("FAILED: "+name);count++;results.Add("PASS "+name);Console.WriteLine("PASS "+name);}
 static WordDoc Text(params string[] lines){var d=new WordDoc();int n=0;foreach(var line in lines){var u=new Unit{Text=line,Location="正文 · 段落 "+(++n)};var cell=new Cell();cell.Spans.Add(new TextSpan{Text=line});u.Cells.Add(cell);d.Units.Add(u);}return d;}
 static Comparison C(WordDoc a,WordDoc b,CompareOptions o=null){return DiffEngine.Compare(a,b,o??new CompareOptions(),CancellationToken.None);}
 static int Main(string[] args){string dir=args[0],output=args[1],fixtures=Path.Combine(Path.GetDirectoryName(dir),"tests","fixtures");Directory.CreateDirectory(output);try{
  var a=DocxReader.Load(Path.Combine(dir,"示例_原版.docx"),CancellationToken.None);var b=DocxReader.Load(Path.Combine(dir,"示例_修订版.docx"),CancellationToken.None);var c=C(a,b);
  Check(c.Count(ChangeKind.Added)>0,"新增段落与表格行");Check(c.Count(ChangeKind.Removed)>0,"删除段落");Check(c.Count(ChangeKind.Modified)>0,"中文参数修改");Check(c.Count(ChangeKind.Formatting)>0,"纯字体格式变化");Check(c.Count(ChangeKind.Image)==1,"相同尺寸图片的内容替换");
  Check(c.Rows.Any(r=>r.Kind!=ChangeKind.Same&&(r.Left??r.Right).Part.StartsWith("页眉")),"页眉内容变化");Check(C(a,a).Changes==0,"同一文档比较无差异");
  var d1=DocxReader.Load(Path.Combine(fixtures,"测试_格式A.docx"),CancellationToken.None);var d2=DocxReader.Load(Path.Combine(fixtures,"测试_格式B.docx"),CancellationToken.None);Check(C(d1,d2).Count(ChangeKind.Formatting)==1,"纯格式修改不误报文本修改");Check(C(d1,d2,new CompareOptions{Formatting=false}).Changes==0,"关闭格式比较");
  Check(C(d1,DocxReader.Load(Path.Combine(fixtures,"测试_分段运行.docx"),CancellationToken.None)).Changes==0,"等价Run切分不产生伪差异");
  var s1=DocxReader.Load(Path.Combine(fixtures,"测试_空白A.docx"),CancellationToken.None);var s2=DocxReader.Load(Path.Combine(fixtures,"测试_空白B.docx"),CancellationToken.None);Check(C(s1,s2,new CompareOptions{IgnoreSpace=true,IgnoreCase=true}).Changes==0,"忽略大小写及空白不误报格式变化");
  Check(C(Text("第一段","相同","相同","末段"),Text("第一段","新增","相同","相同","末段")).Rows.Count==5,"重复段落不会丢失");
  Check(C(Text("开始","参数10吨","结尾"),Text("开始","参数12吨","新增设备","结尾")).Count(ChangeKind.Modified)==1,"插入段落不导致后文整体错位");
  bool[] lm,rm;DiffEngine.Inline("中文10吨😀","中文12吨😀",new CompareOptions(),out lm,out rm);Check(lm.Count(v=>v)==1&&rm.Count(v=>v)==1,"中文及Emoji字符高亮");
  Check(C(Text(),Text()).Changes==0,"两份空文档");Check(C(Text(""),Text()).Changes==0,"默认忽略空段落");
  var noextras=C(a,b,new CompareOptions{Extras=false});Check(noextras.Rows.All(r=>(r.Left??r.Right).Part=="正文"),"附属内容开关");
  string report=Path.Combine(output,"引擎对比报告.docx");Reports.Save(c,report,false);var parsed=DocxReader.Load(report,CancellationToken.None);Check(parsed.Units.Count>5,"导出DOCX可由真实文档读取器重开");
  Reports.Save(c,Path.Combine(output,"引擎对比报告.html"),false);Reports.Save(c,Path.Combine(output,"引擎差异清单.csv"),true);Check(File.ReadAllText(Path.Combine(output,"引擎对比报告.html")).Contains("data:image/png;base64,"),"HTML报告嵌入真实图片");
  bool blocked=false;try{Reports.Save(c,a.Path,false);}catch(IOException){blocked=true;}Check(blocked,"禁止报告覆盖输入原文档");
  blocked=false;try{DocxReader.Load(Path.Combine(fixtures,"测试_无效.docx"),CancellationToken.None);}catch(InvalidDataException){blocked=true;}Check(blocked,"损坏DOCX正确拒绝");
  var cancel=new CancellationTokenSource();cancel.Cancel();blocked=false;try{DiffEngine.Compare(a,b,new CompareOptions(),cancel.Token);}catch(OperationCanceledException){blocked=true;}Check(blocked,"取消令牌终止对比");
  var large=Enumerable.Range(0,3500).Select(i=>"大文档第"+i+"段技术参数").ToArray();var changed=large.ToArray();changed[1700]+="修订";var timer=System.Diagnostics.Stopwatch.StartNew();var lc=C(Text(large),Text(changed));Check(lc.Changes==1&&lc.Rows.Count==3500,"3500段大文档保留完整顺序");results.Add("大文档耗时 "+timer.ElapsedMilliseconds+" ms");
  File.WriteAllLines(Path.Combine(output,"引擎测试结果.txt"),results.Concat(new[]{"TOTAL "+count+" PASSED",Reports.Summary(c)}),new UTF8Encoding(true));return 0;
 }catch(Exception e){Console.WriteLine(e);File.WriteAllText(Path.Combine(output,"引擎测试失败.txt"),e.ToString());return 1;}}
}
