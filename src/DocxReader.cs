using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Globalization;
using System.Diagnostics;
using System.Threading;

namespace Word_Compare {
 public static class OfficeConvert {
  public static string FindLibreOffice(){string p=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),@"LibreOffice\program\soffice.com");if(File.Exists(p))return p;p=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),@"LibreOffice\program\soffice.com");return File.Exists(p)?p:null;}
  public static void Convert(string input,string outputDirectory,string format,string profile,CancellationToken ct){
   string exe=FindLibreOffice();if(exe==null)throw new InvalidOperationException("此格式需要 LibreOffice。请先安装 LibreOffice，或在 Word 中另存为 DOCX 后比较。");
   Directory.CreateDirectory(outputDirectory);Directory.CreateDirectory(profile);
   var start=new ProcessStartInfo(exe,"-env:UserInstallation="+new Uri(Path.GetFullPath(profile)+Path.DirectorySeparatorChar).AbsoluteUri+" --headless --invisible --norestore --convert-to "+format+" --outdir \""+outputDirectory+"\" \""+input+"\""){UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true};
   var log=new StringBuilder();using(var process=new Process{StartInfo=start}){
    process.OutputDataReceived+=(s,e)=>{if(e.Data!=null)lock(log)log.AppendLine(e.Data);};process.ErrorDataReceived+=(s,e)=>{if(e.Data!=null)lock(log)log.AppendLine(e.Data);};process.Start();process.BeginOutputReadLine();process.BeginErrorReadLine();
    var watch=Stopwatch.StartNew();while(!process.WaitForExit(200)){if(ct.IsCancellationRequested||watch.Elapsed.TotalSeconds>90){using(var kill=Process.Start(new ProcessStartInfo("taskkill.exe","/PID "+process.Id+" /T /F"){UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true})){kill.WaitForExit(10000);}ct.ThrowIfCancellationRequested();throw new TimeoutException("LibreOffice 转换超过 90 秒，本次转换进程树已请求终止。");}}
    process.WaitForExit();if(process.ExitCode!=0)throw new IOException("LibreOffice 转换失败："+log);
   }
   string expected=Path.Combine(outputDirectory,Path.GetFileNameWithoutExtension(input)+"."+format.Split(':')[0]);if(!File.Exists(expected)||new FileInfo(expected).Length<100)throw new IOException("LibreOffice 未生成有效文件："+log);
  }
  public static string TempRoot(){string p=Path.Combine(Path.GetTempPath(),"Word_Compare_"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(p);return p;}
  public static void Clean(string p){if(string.IsNullOrEmpty(p))return;string root=Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar;string full=Path.GetFullPath(p);if(!full.StartsWith(root,StringComparison.OrdinalIgnoreCase)||!Path.GetFileName(full).StartsWith("Word_Compare_",StringComparison.Ordinal))return;try{Directory.Delete(full,true);}catch(IOException){}catch(UnauthorizedAccessException){}}
 }
 public sealed class DocxReader {
  static readonly XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships",a="http://schemas.openxmlformats.org/drawingml/2006/main",wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  ZipArchive zip;WordDoc doc;Dictionary<string,XElement> styles=new Dictionary<string,XElement>();Dictionary<string,string> styleNames=new Dictionary<string,string>();StyleInfo defaults=new StyleInfo();XElement defaultParagraph;
  Dictionary<string,XElement> nums=new Dictionary<string,XElement>(),abstracts=new Dictionary<string,XElement>();Dictionary<string,int> counters=new Dictionary<string,int>();Dictionary<string,string> relationships;CancellationToken token;int tableNumber;
  public static string FileHash(string path){using(var s=File.OpenRead(path))using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(s)).Replace("-","");}
  public static WordDoc Load(string path,CancellationToken ct){
   ct.ThrowIfCancellationRequested();if(new FileInfo(path).Length>256L*1024*1024)throw new IOException("文档超过 256 MB，请拆分后比较。");
   string before=FileHash(path);var result=LoadCore(path,ct);ct.ThrowIfCancellationRequested();if(before!=FileHash(path))throw new IOException("文档在读取期间发生变化，请保存后重新对比。");result.Fingerprint=before;return result;
  }
  static WordDoc LoadCore(string path,CancellationToken ct){
   path=Path.GetFullPath(path);if(!File.Exists(path))throw new FileNotFoundException("找不到文档。",path);if(new FileInfo(path).Length>256L*1024*1024)throw new IOException("文档超过 256 MB，请拆分后比较。");
   string ext=Path.GetExtension(path).ToLowerInvariant();if(ext==".txt"){string t=File.ReadAllText(path,Encoding.UTF8);var d=new WordDoc{Path=path};int i=0;foreach(string line in t.Replace("\r","").Split('\n')){var u=new Unit{Text=line,Location="正文 · 段落 "+(++i),Number=i};var cell=new Cell();cell.Spans.Add(new TextSpan{Text=line});u.Cells.Add(cell);d.Units.Add(u);}return d;}
   if(ext==".doc"||ext==".rtf"||ext==".odt"){
    string temp=OfficeConvert.TempRoot();try{OfficeConvert.Convert(path,Path.Combine(temp,"out"),"docx",Path.Combine(temp,"profile"),ct);var result=Load(Path.Combine(temp,"out",Path.GetFileNameWithoutExtension(path)+".docx"),ct);result.Path=path;result.Notes.Insert(0,"旧格式通过 LibreOffice 临时转换后比较，转换可能影响复杂排版。");return result;}finally{OfficeConvert.Clean(temp);}}
   if(ext!=".docx"&&ext!=".docm")throw new NotSupportedException("支持 DOCX、DOCM、DOC、RTF、ODT 和 UTF-8 TXT。");
   using(var stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))using(var archive=new ZipArchive(stream,ZipArchiveMode.Read)){
    var reader=new DocxReader{zip=archive,doc=new WordDoc{Path=path},token=ct};reader.Read();return reader.doc;
   }
  }
  XDocument Xml(string name){var e=zip.GetEntry(name);if(e==null)return null;if(e.Length>100L*1024*1024)throw new IOException("文档 XML 部件过大："+name);using(var s=e.Open())using(var r=XmlReader.Create(s,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=100L*1024*1024}))return XDocument.Load(r);}
  static string V(XElement e,string name,string fallback){return e==null?fallback:(string)e.Attribute(w+name)??fallback;}
  static double D(string s,double fallback){double v;return double.TryParse(s,NumberStyles.Float,CultureInfo.InvariantCulture,out v)?v:fallback;}
  static bool On(XElement e){return e!=null&&!new[]{"0","false","off","none"}.Contains(V(e,"val","1"));}
  static void Apply(StyleInfo s,XElement rp,XElement pp){
   if(rp!=null){var f=rp.Element(w+"rFonts");if(f!=null)s.Font=V(f,"eastAsia",V(f,"ascii",s.Font));if(rp.Element(w+"sz")!=null)s.Size=Math.Max(4,Math.Min(96,D(V(rp.Element(w+"sz"),"val","22"),22)/2));
    if(rp.Element(w+"b")!=null)s.Bold=On(rp.Element(w+"b"));if(rp.Element(w+"i")!=null)s.Italic=On(rp.Element(w+"i"));if(rp.Element(w+"u")!=null)s.Underline=On(rp.Element(w+"u"));if(rp.Element(w+"strike")!=null)s.Strike=On(rp.Element(w+"strike"));if(rp.Element(w+"color")!=null)s.Color=V(rp.Element(w+"color"),"val","");
   }
   if(pp!=null){if(pp.Element(w+"jc")!=null)s.Align=V(pp.Element(w+"jc"),"val","left");foreach(string tag in new[]{"ind","spacing","keepNext","pageBreakBefore","pBdr","shd"}){var el=pp.Element(w+tag);if(el!=null)s.Extra+=tag+Canonical(el);}}
  }
  static string Canonical(XElement e){return e.Name.LocalName+"("+string.Join(",",e.Attributes().Where(x=>!x.IsNamespaceDeclaration).OrderBy(x=>x.Name.LocalName).Select(x=>x.Name.LocalName+"="+x.Value))+")"+string.Concat(e.Elements().Select(Canonical));}
  void Inherit(StyleInfo s,string id,HashSet<string> visited){XElement st;if(id==null||!styles.TryGetValue(id,out st)||!visited.Add(id))return;Inherit(s,V(st.Element(w+"basedOn"),"val",null),visited);Apply(s,st.Element(w+"rPr"),st.Element(w+"pPr"));}
  void Read(){
   var main=Xml("word/document.xml");if(main==null)throw new InvalidDataException("不是有效的 Word 文档：缺少 word/document.xml。");var sx=Xml("word/styles.xml");if(sx!=null){var dd=sx.Root.Element(w+"docDefaults");if(dd!=null)Apply(defaults,dd.Descendants(w+"rPr").FirstOrDefault(),dd.Descendants(w+"pPr").FirstOrDefault());foreach(var s in sx.Root.Elements(w+"style")){string id=V(s,"styleId","");styles[id]=s;styleNames[id]=V(s.Element(w+"name"),"val",id);if(V(s,"type","")=="paragraph"&&V(s,"default","")=="1")defaultParagraph=s;}}
   var nx=Xml("word/numbering.xml");if(nx!=null){foreach(var n in nx.Root.Elements(w+"num"))nums[V(n,"numId","")]=n;foreach(var n in nx.Root.Elements(w+"abstractNum"))abstracts[V(n,"abstractNumId","")]=n;}
   if(main.Descendants().Any(e=>e.Name==w+"ins"||e.Name==w+"del"))doc.Notes.Add("修订按接受后的可见内容比较；不比较修订历史。");
   ReadPart("word/document.xml","正文",main.Root.Element(w+"body"));
   foreach(string kind in new[]{"header","footer"}){int n=0;foreach(var entry in zip.Entries.Where(e=>RegexName(e.FullName,kind)).OrderBy(e=>e.FullName,StringComparer.Ordinal)){var x=Xml(entry.FullName);ReadPart(entry.FullName,(kind=="header"?"页眉":"页脚")+(++n),x.Root);}}
   foreach(string kind in new[]{"footnotes","endnotes","comments"}){var x=Xml("word/"+kind+".xml");if(x==null)continue;foreach(var item in x.Root.Elements()){string id=V(item,"id","");if(id.StartsWith("-",StringComparison.Ordinal)||V(item,"type","")!="")continue;ReadPart("word/"+kind+".xml",(kind=="footnotes"?"脚注":kind=="endnotes"?"尾注":"批注")+id,item);}}
   if(doc.Units.Count==0)doc.Notes.Add("文档没有可比较的文本或图片。");
  }
  static bool RegexName(string name,string kind){return System.Text.RegularExpressions.Regex.IsMatch(name,@"^word/"+kind+@"\d+\.xml$");}
  void ReadPart(string path,string title,XElement root){
   relationships=new Dictionary<string,string>();int slash=path.LastIndexOf('/');string relpath=path.Substring(0,slash+1)+"_rels/"+path.Substring(slash+1)+".rels";var rx=Xml(relpath);if(rx!=null)foreach(var r in rx.Root.Elements()){if((string)r.Attribute("TargetMode")=="External")continue;string target=(string)r.Attribute("Target");if(target!=null)relationships[(string)r.Attribute("Id")]=new Uri(new Uri("http://local/"+path),target).AbsolutePath.TrimStart('/');}
   tableNumber=0;int num=0;foreach(var node in BodyElements(root)){
    token.ThrowIfCancellationRequested();if(node.Name==w+"p"){var u=Paragraph(node,title);u.Number=++num;u.Location=title+" · 段落 "+num;doc.Units.Add(u);}
    else if(node.Name==w+"tbl"){int table=++tableNumber,row=0;foreach(var tr in node.Elements(w+"tr")){var u=new Unit{Part=title,Kind="表格行",Number=++num,Location=title+" · 表 "+table+" / 行 "+(++row)};u.Format="table:"+Canonical(node.Element(w+"tblPr")??new XElement(w+"tblPr"));
      foreach(var tc in tr.Elements(w+"tc")){var cell=new Cell();cell.Span=(int)D(V(tc.Element(w+"tcPr")==null?null:tc.Element(w+"tcPr").Element(w+"gridSpan"),"val","1"),1);bool first=true;
       foreach(var p in tc.Descendants(w+"p").Where(p=>!p.Ancestors(w+"del").Any())){var pu=Paragraph(p,title);if(!first)Add(cell,"\n",defaults);foreach(var span in pu.Cells[0].Spans)Add(cell,span.Text,span.Style);u.Pictures.AddRange(pu.Pictures);u.Format+="|"+pu.Format;first=false;}
       var cp=tc.Element(w+"tcPr");if(cp!=null)u.Format+="cell:"+Canonical(cp);u.Cells.Add(cell);
      }
      Complete(u);doc.Units.Add(u);
     }}
   }
  }
  IEnumerable<XElement> BodyElements(XElement root){if(root==null)yield break;foreach(var e in root.Elements()){if(e.Name==w+"p"||e.Name==w+"tbl")yield return e;else if(e.Name!=w+"del"&&e.Name!=w+"moveFrom")foreach(var n in BodyElements(e))yield return n;}}
  static void Add(Cell cell,string text,StyleInfo s){if(text.Length==0)return;if(cell.Spans.Count>0&&cell.Spans[cell.Spans.Count-1].Style.Key==s.Key)cell.Spans[cell.Spans.Count-1].Text+=text;else cell.Spans.Add(new TextSpan{Text=text,Style=s.Copy()});}
  Unit Paragraph(XElement p,string part){
   var u=new Unit{Part=part};var cell=new Cell();u.Cells.Add(cell);var pp=p.Element(w+"pPr");var ps=defaults.Copy();if(defaultParagraph!=null)Inherit(ps,V(defaultParagraph,"styleId",null),new HashSet<string>());string styleId=V(pp==null?null:pp.Element(w+"pStyle"),"val",null);Inherit(ps,styleId,new HashSet<string>());Apply(ps,null,pp);u.Format=ps.Key;
   XElement numpr=pp==null?null:pp.Element(w+"numPr");if(numpr==null&&styleId!=null&&styles.ContainsKey(styleId))numpr=styles[styleId].Descendants(w+"numPr").FirstOrDefault();string marker=ListMarker(numpr);if(marker!="")Add(cell,marker,ps);
   foreach(var run in p.Descendants(w+"r").Where(r=>!r.Ancestors().Any(e=>e.Name==w+"del"||e.Name==w+"moveFrom"))){var s=ps.Copy();var rp=run.Element(w+"rPr");string rid=V(rp==null?null:rp.Element(w+"rStyle"),"val",null);Inherit(s,rid,new HashSet<string>());Apply(s,rp,null);
    foreach(var el in run.Elements()){
     if(el.Name==w+"t")Add(cell,el.Value,s);else if(el.Name==w+"tab")Add(cell,"\t",s);else if(el.Name==w+"br"||el.Name==w+"cr")Add(cell,"\n",s);
     else if(el.Name==w+"drawing"||el.Name==w+"pict"||el.Name==w+"object"){
      var ids=el.Descendants().Attributes().Where(at=>at.Name==rel+"embed"||at.Name==rel+"id").Select(at=>at.Value).Distinct().ToList();bool found=false;
      foreach(string id in ids){string target;if(!relationships.TryGetValue(id,out target))continue;var entry=zip.GetEntry(Uri.UnescapeDataString(target));if(entry==null||entry.Length>25L*1024*1024)continue;byte[] bytes;using(var ms=new MemoryStream()){using(var es=entry.Open())es.CopyTo(ms);bytes=ms.ToArray();}string hash;using(var sha=SHA256.Create())hash=Convert.ToBase64String(sha.ComputeHash(bytes));var extent=el.Descendants(wp+"extent").FirstOrDefault();var pic=new Picture{Data=bytes,Hash=hash,Label="图片",Width=D((string)(extent==null?null:extent.Attribute("cx")),0)/9525,Height=D((string)(extent==null?null:extent.Attribute("cy")),0)/9525};u.Pictures.Add(pic);found=true;}
      Add(cell,found?"〔图片〕":"〔嵌入对象〕",s);if(!found)u.Media+="object:"+Canonical(el);
     }
    }
   }
   foreach(var math in p.Descendants().Where(e=>e.Name.LocalName=="oMath"&&!e.Ancestors().Any(x=>x.Name.LocalName=="oMath"))){Add(cell,"〔公式："+string.Concat(math.Descendants().Where(x=>x.Name.LocalName=="t").Select(x=>x.Value))+"〕",ps);}
   Complete(u);return u;
  }
  void Complete(Unit u){u.Text=string.Join("\n│\n",u.Cells.Select(c=>c.Text));u.Media+=string.Join("|",u.Pictures.Select(p=>p.Hash+":"+p.Width+":"+p.Height));}
  string ListMarker(XElement np){
   if(np==null)return "";string id=V(np.Element(w+"numId"),"val","0"),level=V(np.Element(w+"ilvl"),"val","0");XElement num,ab;if(id=="0"||!nums.TryGetValue(id,out num)||!abstracts.TryGetValue(V(num.Element(w+"abstractNumId"),"val",""),out ab))return "";
   var lvl=ab.Elements(w+"lvl").FirstOrDefault(e=>V(e,"ilvl","")==level);if(lvl==null)return "";string key=id+":"+level;int start=(int)D(V(lvl.Element(w+"start"),"val","1"),1);var ov=num.Elements(w+"lvlOverride").FirstOrDefault(e=>V(e,"ilvl","")==level);if(ov!=null)start=(int)D(V(ov.Element(w+"startOverride"),"val",start.ToString()),start);
   int value=counters.ContainsKey(key)?counters[key]+1:start;counters[key]=value;for(int k=(int)D(level,0)+1;k<9;k++)counters.Remove(id+":"+k);string fmt=V(lvl.Element(w+"numFmt"),"val","decimal");if(fmt=="bullet")return "• ";string text=V(lvl.Element(w+"lvlText"),"val","%1.");for(int k=0;k<9;k++){string ck=id+":"+k;int v=k.ToString()==level?value:counters.ContainsKey(ck)?counters[ck]:1;var parent=ab.Elements(w+"lvl").FirstOrDefault(e=>V(e,"ilvl","")==k.ToString());string pf=V(parent==null?null:parent.Element(w+"numFmt"),"val","decimal");text=text.Replace("%"+(k+1),Number(v,pf));}return text+" ";
  }
  static string Number(int v,string fmt){if(fmt=="lowerLetter"||fmt=="upperLetter"){string s="";int n=Math.Max(1,v);while(n>0){n--;s=(char)('A'+n%26)+s;n/=26;}return fmt=="lowerLetter"?s.ToLowerInvariant():s;}if((fmt=="lowerRoman"||fmt=="upperRoman")&&v>0&&v<4000){var s=new StringBuilder();int[] values={1000,900,500,400,100,90,50,40,10,9,5,4,1};string[] letters={"M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"};for(int i=0;i<values.Length;i++)while(v>=values[i]){s.Append(letters[i]);v-=values[i];}return fmt=="lowerRoman"?s.ToString().ToLowerInvariant():s.ToString();}return v.ToString(CultureInfo.InvariantCulture);}
 }
}
