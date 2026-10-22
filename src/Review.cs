using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Web.Script.Serialization;

namespace Word_Compare {
 // Review decisions are annotations; they never merge or edit source documents.
 public static class ReviewTools {
  public static List<DiffRow> Filter(Comparison c, int type, int state, bool only){
   return c.Rows.Where(r=>(!only||r.Kind!=ChangeKind.Same)&&(type==0||(int)r.Kind==type)&&(state==0||(r.Kind!=ChangeKind.Same&&(int)r.Review==state-1))).ToList();
  }
  public static List<DiffRow> Sort(IEnumerable<DiffRow> rows,int order){
   if(order==1)return rows.OrderBy(r=>r.Kind).ThenBy(r=>r.Index).ToList();
   if(order==2)return rows.OrderByDescending(r=>Math.Max(r.Left==null?0:r.Left.Text.Length,r.Right==null?0:r.Right.Text.Length)).ThenBy(r=>r.Index).ToList();
   return rows.OrderBy(r=>r.Index).ToList();
  }
  public static string Summary(Comparison c){return string.Format("待审 {0} · 已采纳 {1} · 不采纳 {2}",c.Rows.Count(r=>r.Kind!=ChangeKind.Same&&r.Review==ReviewState.Pending),c.Rows.Count(r=>r.Kind!=ChangeKind.Same&&r.Review==ReviewState.Accepted),c.Rows.Count(r=>r.Kind!=ChangeKind.Same&&r.Review==ReviewState.Rejected));}
 }
 public sealed class ReviewEntry {public string id,kind,state;public int index;public object before,after;}
 public sealed class ReviewFile {public string format,binding,createdUtc,application;public int version;public object source,options,statistics;public List<ReviewEntry> rows;}
 public static class ReviewStore {
  static JavaScriptSerializer Serializer(){return new JavaScriptSerializer{MaxJsonLength=128*1024*1024,RecursionLimit=128};}
  static string Hash(string text){using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-","");}
  static string UnitKey(Unit u){if(u==null)return "";return Serializer().Serialize(new[]{u.Part,u.Location,u.Kind,u.Text,u.Signature,u.Media});}
  public static string RowId(DiffRow r){return Hash(Serializer().Serialize(new[]{r.Index.ToString(),r.Kind.ToString(),UnitKey(r.Left),UnitKey(r.Right)}));}
  static string DocKey(WordDoc d){return d.Fingerprint.Length>0?d.Fingerprint:Hash(string.Join("\n",d.Units.Select(UnitKey)));}
  static string Binding(Comparison c){return Hash(Serializer().Serialize(new[]{DocKey(c.Left),DocKey(c.Right),c.Options.Formatting.ToString(),c.Options.IgnoreSpace.ToString(),c.Options.IgnoreCase.ToString(),c.Options.IgnoreEmpty.ToString(),c.Options.Extras.ToString()}));}
  static object UnitData(Unit u){if(u==null)return null;return new{part=u.Part,location=u.Location,kind=u.Kind,text=u.Text,format=u.Format,media=u.Media,cells=u.Cells.Select(cell=>new{span=cell.Span,runs=cell.Spans.Select(s=>new{text=s.Text,style=s.Style}).ToArray()}).ToArray(),images=u.Pictures.Select(p=>new{fingerprint=p.Hash,width=p.Width,height=p.Height,label=p.Label}).ToArray()};}
  public static void Save(Comparison c,string file){
   string full=Path.GetFullPath(file);if(new[]{c.Left.Path,c.Right.Path}.Where(p=>!string.IsNullOrEmpty(p)).Any(p=>string.Equals(Path.GetFullPath(p),full,StringComparison.OrdinalIgnoreCase)))throw new IOException("审核记录不能覆盖输入文档。");
   var data=new ReviewFile{format="Word_Compare.Review",version=1,application="Word_Compare",createdUtc=DateTime.UtcNow.ToString("o"),binding=Binding(c),source=new{left=new{path=c.Left.Path,sha256=DocKey(c.Left)},right=new{path=c.Right.Path,sha256=DocKey(c.Right)}},options=c.Options,statistics=new{totalRows=c.Rows.Count,changes=c.Changes,pending=c.Rows.Count(r=>r.Kind!=ChangeKind.Same&&r.Review==ReviewState.Pending),accepted=c.Rows.Count(r=>r.Kind!=ChangeKind.Same&&r.Review==ReviewState.Accepted),rejected=c.Rows.Count(r=>r.Kind!=ChangeKind.Same&&r.Review==ReviewState.Rejected)},rows=c.Rows.Select(r=>new ReviewEntry{id=RowId(r),index=r.Index,kind=r.Kind.ToString(),state=r.Review.ToString(),before=UnitData(r.Left),after=UnitData(r.Right)}).ToList()};
   string json=Serializer().Serialize(data);string temp=full+"."+Guid.NewGuid().ToString("N")+".tmp";
   try{File.WriteAllText(temp,json,new UTF8Encoding(true));if(File.Exists(full))File.Replace(temp,full,null);else File.Move(temp,full);}finally{if(File.Exists(temp))File.Delete(temp);}
  }
  public static int Load(Comparison c,string file){
   if(new FileInfo(file).Length>128L*1024*1024)throw new IOException("审核记录超过 128 MB。");
   ReviewFile data;try{data=Serializer().Deserialize<ReviewFile>(File.ReadAllText(file,Encoding.UTF8));}catch(Exception e){throw new InvalidDataException("无效的审核 JSON："+e.Message);}
   if(data==null||data.format!="Word_Compare.Review"||data.version!=1||data.binding!=Binding(c))throw new InvalidDataException("审核记录的文档内容、左右顺序或比较规则与当前结果不同，请先打开匹配的文档并比较。");
   var expected=c.Rows.ToDictionary(RowId);if(data.rows==null||data.rows.Count!=expected.Count)throw new InvalidDataException("审核记录不完整，未载入任何状态。");
   var states=new Dictionary<DiffRow,ReviewState>();foreach(var entry in data.rows){DiffRow row;ReviewState state;
    if(entry==null||entry.id==null||!expected.TryGetValue(entry.id,out row)||entry.index!=row.Index||entry.kind!=row.Kind.ToString()||!Enum.TryParse<ReviewState>(entry.state,out state)||!Enum.IsDefined(typeof(ReviewState),state)||states.ContainsKey(row)||(row.Kind==ChangeKind.Same&&state!=ReviewState.Pending))throw new InvalidDataException("审核记录含无效或重复项目，未载入任何状态。");states.Add(row,state);
   }
   foreach(var pair in states)pair.Key.Review=pair.Value;return states.Count(p=>p.Key.Kind!=ChangeKind.Same&&p.Value!=ReviewState.Pending);
  }
 }
}
