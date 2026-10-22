using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Threading;

namespace Word_Compare {
 public static class Reports {
  static readonly XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  public static string H(string s){return System.Net.WebUtility.HtmlEncode(s??"");}
  public static void Save(Comparison c,string file,bool differencesOnly){
   string full=Path.GetFullPath(file);if(string.Equals(full,Path.GetFullPath(c.Left.Path),StringComparison.OrdinalIgnoreCase)||string.Equals(full,Path.GetFullPath(c.Right.Path),StringComparison.OrdinalIgnoreCase))throw new IOException("报告不能覆盖输入文档，请使用新的文件名。");
   string ext=Path.GetExtension(file).ToLowerInvariant();if(ext==".json")ReviewStore.Save(c,file);else if(ext==".html"||ext==".htm")Html(c,file,differencesOnly);else if(ext==".csv")Csv(c,file,differencesOnly);else if(ext==".docx")Docx(c,file,differencesOnly);else if(ext==".pdf"){
    string temp=OfficeConvert.TempRoot();try{string src=Path.Combine(temp,"Word_Compare报告.docx");Docx(c,src,differencesOnly);OfficeConvert.Convert(src,Path.Combine(temp,"pdf"),"pdf",Path.Combine(temp,"profile"),CancellationToken.None);File.Copy(Path.Combine(temp,"pdf","Word_Compare报告.pdf"),file,true);}finally{OfficeConvert.Clean(temp);}
   }else throw new NotSupportedException("报告支持 DOCX、HTML、CSV、PDF 和 JSON。");
  }
  static IEnumerable<DiffRow> Rows(Comparison c,bool only){return c.Rows.Where(r=>!only||r.Kind!=ChangeKind.Same);}
  static string TextHtml(Unit u,bool[] mask,ChangeKind kind,bool left){
   if(u==null)return "<span class='muted'>此处无内容</span>";var s=new StringBuilder();int offset=0;s.Append("<small>").Append(H(u.Location)).Append("</small>");if(u.Kind=="表格行")s.Append("<div class='cells'>");
   foreach(var cell in u.Cells){s.Append("<div class='cell'>");foreach(var span in cell.Spans){int i=0;while(i<span.Text.Length){bool changed=mask!=null&&offset+i<mask.Length&&mask[offset+i];int end=i+1;while(end<span.Text.Length&&((mask!=null&&offset+end<mask.Length&&mask[offset+end])==changed))end++;var st=span.Style;
     s.Append("<span style='white-space:pre-wrap;font-family:").Append(H(System.Text.RegularExpressions.Regex.Replace(st.Font,@"[^\p{L}\p{N} _.,-]",""))).Append(";font-size:").Append(st.Size.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append("pt;").Append(st.Bold?"font-weight:bold;":"").Append(st.Italic?"font-style:italic;":"").Append(st.Underline?"text-decoration:underline;":"").Append(changed?(left?"background:#ffd5d5;color:#8f2437;":"background:#c9f1db;color:#096443;"):"").Append("'>").Append(H(span.Text.Substring(i,end-i))).Append("</span>");i=end;}
     offset+=span.Text.Length;}s.Append("</div>");offset+=3;}
   if(u.Kind=="表格行")s.Append("</div>");foreach(var pic in u.Pictures){string mime=Mime(pic.Data);if(mime!="")s.Append("<img alt='图片' src='data:").Append(mime).Append(";base64,").Append(Convert.ToBase64String(pic.Data)).Append("'>");else s.Append("<p>图片指纹：").Append(H(pic.Hash.Substring(0,12))).Append("</p>");}return s.ToString();
  }
  static string Mime(byte[] b){if(b.Length>3&&b[0]==137&&b[1]==80)return "image/png";if(b.Length>3&&b[0]==255&&b[1]==216)return "image/jpeg";if(b.Length>3&&b[0]==71&&b[1]==73)return "image/gif";if(b.Length>3&&b[0]==66&&b[1]==77)return "image/bmp";return "";}
  static void Html(Comparison c,string file,bool only){
   var s=new StringBuilder("<!doctype html><html lang='zh-CN'><meta charset='utf-8'><title>Word_Compare 双栏对比报告</title><style>body{font:14px 'Microsoft YaHei',sans-serif;color:#263445;background:#f5f7fa;margin:24px}h1{font-size:23px}table{width:100%;border-collapse:collapse;table-layout:fixed;background:white}th,td{border:1px solid #cbd5e1;padding:12px;vertical-align:top;overflow-wrap:anywhere}th{background:#e6f2f0}small{display:block;color:#64748b;margin-bottom:10px}.added{background:#edf9f0}.removed{background:#fff0f1}.modified,.formatting,.image{background:#fff9e8}.muted{color:#94a3b8}.cells{display:flex;gap:0}.cells .cell{flex:1;border:1px solid #cbd5e1;padding:8px;min-width:0}img{max-width:100%;max-height:240px;margin:8px} @media print{body{background:white;margin:0}tr{break-inside:avoid}}</style><h1>Word_Compare 双栏对比报告</h1>");
   s.Append("<p>左侧：").Append(H(c.Left.Path)).Append("<br>右侧：").Append(H(c.Right.Path)).Append("</p><p>").Append(H(Summary(c))).Append("</p><p>范围：").Append(c.Options.Extras?"正文及页眉、页脚、脚注、尾注、批注":"仅正文").Append("；格式比较：").Append(c.Options.Formatting?"开启":"关闭").Append("；忽略空白：").Append(c.Options.IgnoreSpace?"是":"否").Append("；忽略大小写：").Append(c.Options.IgnoreCase?"是":"否").Append("</p><table><col style='width:85px'><col><col><thead><tr><th>类型</th><th>左侧 · 原文档</th><th>右侧 · 对比文档</th></tr></thead><tbody>");
   foreach(var r in Rows(c,only)){s.Append("<tr class='").Append(r.Kind.ToString().ToLowerInvariant()).Append("'><td>").Append(r.KindText).Append("<br><small>").Append(r.ReviewText).Append("</small></td><td>").Append(TextHtml(r.Left,r.LeftMask,r.Kind,true)).Append("</td><td>").Append(TextHtml(r.Right,r.RightMask,r.Kind,false)).Append("</td></tr>");}
   s.Append("</tbody></table><p>此报告为内容对齐视图，不复刻 Word 原分页；浮动对象与复杂域可能简化。</p></html>");File.WriteAllText(file,s.ToString(),new UTF8Encoding(true));
  }
  public static string Summary(Comparison c){return string.Format("差异 {0} 项  ·  修改 {1}  ·  新增 {2}  ·  删除 {3}  ·  格式 {4}  ·  图片 {5}",c.Changes,c.Count(ChangeKind.Modified),c.Count(ChangeKind.Added),c.Count(ChangeKind.Removed),c.Count(ChangeKind.Formatting),c.Count(ChangeKind.Image));}
  static string Q(string s){s=s??"";if(s.Length>0&&"=+-@".IndexOf(s[0])>=0)s="'"+s;return "\""+s.Replace("\"","\"\"")+"\"";}
  static void Csv(Comparison c,string file,bool only){var s=new StringBuilder("类型,审核状态,左侧位置,右侧位置,左侧内容,右侧内容\r\n");foreach(var r in Rows(c,only))s.AppendLine(string.Join(",",new[]{r.KindText,r.ReviewText,r.Left==null?"":r.Left.Location,r.Right==null?"":r.Right.Location,r.Left==null?"":r.Left.Text,r.Right==null?"":r.Right.Text}.Select(Q)));File.WriteAllText(file,s.ToString(),new UTF8Encoding(true));}
  static XElement P(string text,bool bold){return new XElement(w+"p",new XElement(w+"r",new XElement(w+"rPr",new XElement(w+"rFonts",new XAttribute(w+"eastAsia","宋体")),new XElement(w+"sz",new XAttribute(w+"val",20)),bold?new XElement(w+"b"):null),new XElement(w+"t",new XAttribute(XNamespace.Xml+"space","preserve"),text)));}
  static IEnumerable<XElement> UnitParagraphs(Unit u,bool[] mask,bool left){
   if(u==null){yield return P("此处无内容",false);yield break;}yield return P(u.Location,true);int offset=0;
   foreach(var cell in u.Cells){var p=new XElement(w+"p");foreach(var span in cell.Spans){int i=0;while(i<span.Text.Length){bool ch=mask!=null&&offset+i<mask.Length&&mask[offset+i];int end=i+1;while(end<span.Text.Length&&((mask!=null&&offset+end<mask.Length&&mask[offset+end])==ch))end++;
     var st=span.Style;var rp=new XElement(w+"rPr",new XElement(w+"rFonts",new XAttribute(w+"eastAsia",st.Font),new XAttribute(w+"ascii",st.Font)),new XElement(w+"sz",new XAttribute(w+"val",Math.Round(st.Size*2))),st.Bold?new XElement(w+"b"):null,st.Italic?new XElement(w+"i"):null,st.Underline?new XElement(w+"u",new XAttribute(w+"val","single")):null,ch?new XElement(w+"shd",new XAttribute(w+"fill",left?"FFD5D5":"C9F1DB")):null);
     string text=span.Text.Substring(i,end-i);var run=new XElement(w+"r",rp);string[] lines=text.Split('\n');for(int j=0;j<lines.Length;j++){if(j>0)run.Add(new XElement(w+"br"));run.Add(new XElement(w+"t",new XAttribute(XNamespace.Xml+"space","preserve"),lines[j]));}p.Add(run);i=end;}
     offset+=span.Text.Length;}if(!p.HasElements)p.Add(new XElement(w+"r",new XElement(w+"t","")));yield return p;offset+=3;}
   foreach(var image in u.Pictures)yield return P("〔图片指纹 "+image.Hash.Substring(0,12)+"；请在 HTML 报告中查看图片〕",false);
  }
  static void Entry(ZipArchive z,string name,string content){var e=z.CreateEntry(name);using(var s=new StreamWriter(e.Open(),new UTF8Encoding(false)))s.Write(content);}
  static void Docx(Comparison c,string file,bool only){
   var body=new XElement(w+"body",P("Word_Compare 双栏对比报告",true),P("左侧："+c.Left.Path,false),P("右侧："+c.Right.Path,false),P(Summary(c),true),P("对齐内容视图；图片详见 HTML 报告。原始文档未修改。",false));
   var table=new XElement(w+"tbl",new XElement(w+"tblPr",new XElement(w+"tblW",new XAttribute(w+"w",0),new XAttribute(w+"type","auto")),new XElement(w+"tblBorders",new[]{"top","left","bottom","right","insideH","insideV"}.Select(n=>new XElement(w+n,new XAttribute(w+"val","single"),new XAttribute(w+"sz",4),new XAttribute(w+"color","CBD5E1"))))),new XElement(w+"tblGrid",new[]{900,6900,6900}.Select(v=>new XElement(w+"gridCol",new XAttribute(w+"w",v)))));
   table.Add(new XElement(w+"tr",new XElement(w+"trPr",new XElement(w+"tblHeader")),new[]{"类型","左侧 · 原文档","右侧 · 对比文档"}.Select(s=>new XElement(w+"tc",P(s,true)))));
   foreach(var r in Rows(c,only)){string fill=r.Kind==ChangeKind.Added?"EDF9F0":r.Kind==ChangeKind.Removed?"FFF0F1":r.Kind==ChangeKind.Same?"FFFFFF":"FFF9E8";var tr=new XElement(w+"tr");tr.Add(new XElement(w+"tc",P(r.KindText,false),P(r.ReviewText,false)));tr.Add(new XElement(w+"tc",new XElement(w+"tcPr",new XElement(w+"shd",new XAttribute(w+"fill",fill))),UnitParagraphs(r.Left,r.LeftMask,true)));tr.Add(new XElement(w+"tc",new XElement(w+"tcPr",new XElement(w+"shd",new XAttribute(w+"fill",fill))),UnitParagraphs(r.Right,r.RightMask,false)));table.Add(tr);}body.Add(table);
   body.Add(new XElement(w+"sectPr",new XElement(w+"pgSz",new XAttribute(w+"w",16838),new XAttribute(w+"h",11906),new XAttribute(w+"orient","landscape")),new XElement(w+"pgMar",new XAttribute(w+"top",720),new XAttribute(w+"right",720),new XAttribute(w+"bottom",720),new XAttribute(w+"left",720),new XAttribute(w+"header",360),new XAttribute(w+"footer",360))));
   using(var fs=new FileStream(file,FileMode.Create,FileAccess.Write))using(var z=new ZipArchive(fs,ZipArchiveMode.Create)){
    Entry(z,"[Content_Types].xml","<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>");
    Entry(z,"_rels/.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");Entry(z,"word/document.xml",new XDocument(new XElement(w+"document",new XAttribute(XNamespace.Xmlns+"w",w),body)).ToString());
   }
  }
 }
}
