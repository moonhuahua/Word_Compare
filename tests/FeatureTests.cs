using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Word_Compare;

class FeatureTests {
 static List<string> results=new List<string>();static string output;
 static void Check(bool value,string name){if(!value)throw new Exception("FAILED "+name);results.Add("PASS "+name);Console.WriteLine("PASS "+name);}
 static WordDoc Doc(params string[] texts){var doc=new WordDoc{Path=Path.Combine(output,"virtual-"+Guid.NewGuid().ToString("N")+".docx")};foreach(string text in texts){var cell=new Cell();cell.Spans.Add(new TextSpan{Text=text});var u=new Unit{Text=text,Location="正文 · 段落 "+(doc.Units.Count+1)};u.Cells.Add(cell);doc.Units.Add(u);}return doc;}
 static Comparison Compare(WordDoc a,WordDoc b,CompareOptions o=null){return DiffEngine.Compare(a,b,o??new CompareOptions(),CancellationToken.None);}
 static bool Reject(Comparison c,string file){try{ReviewStore.Load(c,file);return false;}catch(InvalidDataException){return true;}}
 static string Masked(string text,bool[] mask){return string.Concat(text.Where((c,i)=>mask[i]));}
 static int Main(string[] args){string root=args[0];output=Path.Combine(root,"验证结果");Directory.CreateDirectory(output);string record=Path.Combine(output,"审核完整记录.json"),invalid=Path.Combine(output,"审核无效样例.json");try{
  bool[] a,b;var options=new CompareOptions();DiffEngine.Inline("flow 10 m3","flow 12 m3",options,out a,out b);Check(Masked("flow 10 m3",a)=="0","字符粒度仅标出变化字符");
  options.Highlight=HighlightLevel.Word;DiffEngine.Inline("flow 10 m3","flow 12 m3",options,out a,out b);Check(Masked("flow 10 m3",a)=="10"&&Masked("flow 12 m3",b)=="12","词与数字粒度完整高亮数值");
  DiffEngine.Inline("中文甲😀","中文乙😀",options,out a,out b);Check(Masked("中文甲😀",a)=="甲"&&Masked("中文乙😀",b)=="乙","词模式兼容中文及Emoji");
  options.Highlight=HighlightLevel.Sentence;string left="设备正常。产水10吨！请确认。",right="设备正常。产水12吨！请确认。";DiffEngine.Inline(left,right,options,out a,out b);Check(Masked(left,a)=="产水10吨！"&&Masked(right,b)=="产水12吨！","中文句子粒度仅标出变化句子");
  DiffEngine.Inline("Start. Flow 10. Done!","Start. Flow 12. Done!",options,out a,out b);Check(Masked("Start. Flow 10. Done!",a)==" Flow 10.","英文句号与中文标点均识别句界");
  options.Highlight=HighlightLevel.Paragraph;DiffEngine.Inline(left,right,options,out a,out b);Check(a.All(v=>v)&&b.All(v=>v),"整段粒度高亮完整变化段落");
  foreach(HighlightLevel level in Enum.GetValues(typeof(HighlightLevel))){options.Highlight=level;options.IgnoreCase=true;options.IgnoreSpace=true;DiffEngine.Inline("ABC 中 文 😀","abc 中文😀",options,out a,out b);Check(!a.Any(v=>v)&&!b.Any(v=>v),"忽略空白大小写与高亮粒度兼容 "+level);}
  var original=DocxReader.Load(Path.Combine(root,"示例","示例_原版.docx"),CancellationToken.None);var revised=DocxReader.Load(Path.Combine(root,"示例","示例_修订版.docx"),CancellationToken.None);var c=Compare(original,revised);int changes=c.Changes,rows=c.Rows.Count;var delta=c.Rows.Where(r=>r.Kind!=ChangeKind.Same).ToList();delta[0].Review=ReviewState.Accepted;delta[1].Review=ReviewState.Rejected;
  Check(original.Fingerprint==DocxReader.FileHash(original.Path)&&original.Fingerprint.Length==64,"读取时保存真实文件SHA256");
  DiffEngine.Rehighlight(c,HighlightLevel.Sentence);Check(c.Rows.Count==rows&&c.Changes==changes&&delta[0].Review==ReviewState.Accepted,"切换高亮不改变对齐差异及审核状态");
  Check(ReviewTools.Filter(c,(int)ChangeKind.Added,0,false).All(r=>r.Kind==ChangeKind.Added),"差异类型筛选");
  Check(ReviewTools.Filter(c,0,1,false).Count==changes-2,"待审筛选排除相同内容和已审项目");
  Check(ReviewTools.Filter(c,0,2,false).Single()==delta[0]&&ReviewTools.Filter(c,0,3,false).Single()==delta[1],"已采纳与不采纳分别筛选");
  var typed=ReviewTools.Sort(delta,1);Check(typed.Select(r=>(int)r.Kind).SequenceEqual(typed.Select(r=>(int)r.Kind).OrderBy(v=>v)),"按差异类型排序");
  var sized=ReviewTools.Sort(delta,2);Func<DiffRow,int> length=r=>Math.Max(r.Left==null?0:r.Left.Text.Length,r.Right==null?0:r.Right.Text.Length);Check(sized.Select(length).SequenceEqual(sized.Select(length).OrderByDescending(v=>v)),"按内容长度降序排序");
  ReviewStore.Save(c,record);ReviewStore.Save(c,record);Check(!Directory.GetFiles(output,"*.tmp").Any(),"审核记录可原子覆盖且清理临时文件");
  var restored=Compare(original,revised);Check(ReviewStore.Load(restored,record)==2&&restored.Rows[delta[0].Index].Review==ReviewState.Accepted&&restored.Rows[delta[1].Index].Review==ReviewState.Rejected,"重新比较后完整恢复审核状态");
  Check(Reject(Compare(revised,original),record),"拒绝左右顺序不匹配的审核记录");
  Check(Reject(Compare(original,revised,new CompareOptions{IgnoreCase=true}),record),"拒绝比较规则不同的审核记录");
  var other=DocxReader.Load(revised.Path,CancellationToken.None);other.Fingerprint=new string('F',64);Check(Reject(Compare(original,other),record),"拒绝源文件指纹不同的审核记录");
  var serializer=new JavaScriptSerializer{MaxJsonLength=128*1024*1024};var data=serializer.Deserialize<ReviewFile>(File.ReadAllText(record));data.rows[1].state="999";File.WriteAllText(invalid,serializer.Serialize(data));var clean=Compare(original,revised);Check(Reject(clean,invalid)&&clean.Rows.All(r=>r.Review==ReviewState.Pending),"无效状态整批拒绝且不部分应用");
  data=serializer.Deserialize<ReviewFile>(File.ReadAllText(record));data.rows[1]=data.rows[0];File.WriteAllText(invalid,serializer.Serialize(data));Check(Reject(clean,invalid)&&clean.Rows.All(r=>r.Review==ReviewState.Pending),"重复项目整批拒绝且不部分应用");File.Delete(invalid);
  string longText=new string('中',4000)+"\"\\\n<script>alert(1)</script>尾部";var many=Compare(Doc(Enumerable.Range(0,101).Select(i=>"原参数"+i+longText).ToArray()),Doc(Enumerable.Range(0,101).Select(i=>"新参数"+i+longText).ToArray()));string full=Path.Combine(output,"长文JSON验证.json");Reports.Save(many,full,true);var all=serializer.Deserialize<ReviewFile>(File.ReadAllText(full));var before=(Dictionary<string,object>)all.rows[0].before;Check(all.rows.Count==101&&(string)before["text"]=="原参数0"+longText,"JSON不截断超过50项及长中文内容并正确转义");File.Delete(full);
  bool blocked=false;try{ReviewStore.Save(c,original.Path);}catch(IOException){blocked=true;}Check(blocked,"JSON审核记录禁止覆盖原文档");
  string html=Path.Combine(output,"审核对比报告.html"),csv=Path.Combine(output,"审核差异清单.csv"),docx=Path.Combine(output,"审核对比报告.docx");Reports.Save(c,html,false);Reports.Save(c,csv,true);Reports.Save(c,docx,false);Check(File.ReadAllText(html).Contains("已采纳")&&File.ReadAllText(csv).Contains("不采纳")&&DocxReader.Load(docx,CancellationToken.None).Units.Any(u=>u.Text.Contains("已采纳")),"HTML CSV DOCX报告包含审核状态");
  var safe=Compare(Doc("<script>alert(1)</script>"),Doc("<script>alert(2)</script>"));Reports.Save(safe,html,false);Check(!File.ReadAllText(html).Contains("<script>"),"HTML正文转义阻止脚本注入");Reports.Save(c,html,false);
  Check(DocxReader.FileHash(original.Path)==original.Fingerprint&&DocxReader.FileHash(revised.Path)==revised.Fingerprint,"审核和报告导出均保持输入文档字节不变");
  File.WriteAllLines(Path.Combine(output,"新增功能测试.txt"),results.Concat(new[]{"TOTAL "+results.Count+" PASSED"}),new UTF8Encoding(true));return 0;
 }catch(Exception e){Console.WriteLine(e);File.WriteAllText(Path.Combine(output,"新增功能测试失败.txt"),e.ToString(),new UTF8Encoding(true));return 1;}}
}
