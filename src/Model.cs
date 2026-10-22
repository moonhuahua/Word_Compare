using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace Word_Compare {
 public sealed class StyleInfo {
  public string Font="Microsoft YaHei", Color="", Align="left", Extra=""; public double Size=11; public bool Bold,Italic,Underline,Strike;
  public StyleInfo Copy(){return (StyleInfo)MemberwiseClone();}
  public string Key {get{return Font+"|"+Size.ToString(CultureInfo.InvariantCulture)+"|"+Color+"|"+Bold+Italic+Underline+Strike+"|"+Align+"|"+Extra;}}
 }
 public sealed class TextSpan {public string Text=""; public StyleInfo Style=new StyleInfo();}
 public sealed class Picture {public byte[] Data;public string Hash="",Label="图片";public double Width,Height;}
 public sealed class Cell {public List<TextSpan> Spans=new List<TextSpan>();public int Span=1;public string Text{get{return string.Concat(Spans.Select(s=>s.Text));}}}
 public sealed class Unit {
  public string Part="正文",Location="",Kind="段落",Text="",Format="",Media="";public int Number;
  public List<Cell> Cells=new List<Cell>();public List<Picture> Pictures=new List<Picture>();
  public string Signature {get{return Format+"|"+string.Join("|",Cells.Select(c=>c.Span+":"+string.Join("",c.Spans.Where(s=>s.Text.Length>0).Select(s=>s.Text+":"+s.Style.Key))));}}
 }
 public sealed class WordDoc {public string Path="",Fingerprint="";public List<Unit> Units=new List<Unit>();public List<string> Notes=new List<string>();}
 public enum HighlightLevel { Character,Word,Sentence,Paragraph }
 public sealed class CompareOptions {public bool Formatting=true,IgnoreSpace=false,IgnoreCase=false,IgnoreEmpty=true,Extras=true;public HighlightLevel Highlight=HighlightLevel.Character;}
 public enum ChangeKind { Same,Added,Removed,Modified,Formatting,Image }
 public enum ReviewState { Pending,Accepted,Rejected }
 public sealed class DiffRow {
  public Unit Left,Right; public ChangeKind Kind;public int Index;public bool[] LeftMask,RightMask;public ReviewState Review;
  public string ReviewText {get{return Kind==ChangeKind.Same?"":new[]{"待审","已采纳","不采纳"}[(int)Review];}}
  public string KindText {get{return new[]{"相同","新增","删除","修改","格式","图片"}[(int)Kind];}}
  public string Location {get{return (Left??Right).Location;}}
  public string Preview {get{var s=(Right??Left).Text.Replace('\n',' ');return s.Length>65?s.Substring(0,65)+"…":s;}}
  public string ListText {get{return KindText+" · "+ReviewText+" · "+Location+"\n"+Preview;}}
 }
 public sealed class Comparison {
  public WordDoc Left,Right;public CompareOptions Options;public List<DiffRow> Rows=new List<DiffRow>();
  public int Count(ChangeKind k){return Rows.Count(r=>r.Kind==k);}public int Changes{get{return Rows.Count(r=>r.Kind!=ChangeKind.Same);}}
 }
 public static class DiffEngine {
  static string Normalize(string s,CompareOptions o){if(o.IgnoreSpace)s=Regex.Replace(s,@"\s+","");if(o.IgnoreCase)s=s.ToUpperInvariant();return s;}
  static string Key(Unit u,CompareOptions o){return u.Kind+"|"+Normalize(u.Text,o);}
  static string FormatKey(Unit u,CompareOptions o){return u.Format+"|"+string.Join("|",u.Cells.Select(c=>c.Span+":"+string.Join("",c.Spans.Select(s=>new{Text=Normalize(s.Text,o),Style=s.Style.Key}).Where(s=>s.Text.Length>0).Select(s=>s.Text+":"+s.Style))));}
  // LCS uses a bounded matrix. Large documents use unique paragraph anchors.
  static List<Tuple<int,int>> Matches(string[] a,string[] b,CancellationToken ct){
   var result=new List<Tuple<int,int>>(); MatchRange(a,0,a.Length,b,0,b.Length,result,ct);return result;
  }
  static void MatchRange(string[] a,int a0,int a1,string[] b,int b0,int b1,List<Tuple<int,int>> result,CancellationToken ct){
   ct.ThrowIfCancellationRequested();while(a0<a1&&b0<b1&&a[a0]==b[b0]){result.Add(Tuple.Create(a0++,b0++));}
   int suffix=0;while(a0<a1-suffix&&b0<b1-suffix&&a[a1-1-suffix]==b[b1-1-suffix])suffix++;
   int ae=a1-suffix,be=b1-suffix,n=ae-a0,m=be-b0;
   if(n>0&&m>0){
    if((long)n*m<=6000000){
     var dp=new int[n+1,m+1];for(int i=n-1;i>=0;i--){ct.ThrowIfCancellationRequested();for(int j=m-1;j>=0;j--)dp[i,j]=a[a0+i]==b[b0+j]?dp[i+1,j+1]+1:Math.Max(dp[i+1,j],dp[i,j+1]);}
     int x=0,y=0;while(x<n&&y<m){if(a[a0+x]==b[b0+y]){result.Add(Tuple.Create(a0+x,b0+y));x++;y++;}else if(dp[x+1,y]>=dp[x,y+1])x++;else y++;}
    }else{
     var ac=new Dictionary<string,int>();var bc=new Dictionary<string,int>();
     for(int i=a0;i<ae;i++){if(ac.ContainsKey(a[i]))ac[a[i]]=-1;else ac[a[i]]=i;}
     for(int j=b0;j<be;j++){if(bc.ContainsKey(b[j]))bc[b[j]]=-1;else bc[b[j]]=j;}
     var candidates=ac.Where(k=>k.Value>=0&&bc.ContainsKey(k.Key)&&bc[k.Key]>=0).Select(k=>Tuple.Create(k.Value,bc[k.Key])).OrderBy(k=>k.Item1).ToList();
     // Longest increasing subsequence of unique anchors, O(n log n).
     var tails=new List<int>();var previous=new int[candidates.Count];
     for(int k=0;k<candidates.Count;k++){int lo=0,hi=tails.Count;while(lo<hi){int mid=(lo+hi)/2;if(candidates[tails[mid]].Item2<candidates[k].Item2)lo=mid+1;else hi=mid;}previous[k]=lo>0?tails[lo-1]:-1;if(lo==tails.Count)tails.Add(k);else tails[lo]=k;}
     if(tails.Count>0){var chain=new List<Tuple<int,int>>();for(int k=tails[tails.Count-1];k>=0;k=previous[k])chain.Add(candidates[k]);chain.Reverse();int x=a0,y=b0;foreach(var anchor in chain){MatchRange(a,x,anchor.Item1,b,y,anchor.Item2,result,ct);result.Add(anchor);x=anchor.Item1+1;y=anchor.Item2+1;}MatchRange(a,x,ae,b,y,be,result,ct);}
    }
   }
   for(int k=suffix;k>0;k--)result.Add(Tuple.Create(a1-k,b1-k));
  }
  static double Similarity(Unit a,Unit b,CompareOptions o){
   if(a.Kind!=b.Kind)return 0;string x=Normalize(a.Text,o),y=Normalize(b.Text,o);if(x==y)return 1;if(x.Length==0||y.Length==0)return 0;
   var xs=new HashSet<string>();var ys=new HashSet<string>();for(int i=0;i<x.Length;i++)xs.Add(x.Substring(i,Math.Min(2,x.Length-i)));for(int i=0;i<y.Length;i++)ys.Add(y.Substring(i,Math.Min(2,y.Length-i)));
   return 2.0*xs.Count(ys.Contains)/(xs.Count+ys.Count);
  }
  static DiffRow Row(Unit a,Unit b,CompareOptions o){
   var r=new DiffRow{Left=a,Right=b};if(a==null)r.Kind=ChangeKind.Added;else if(b==null)r.Kind=ChangeKind.Removed;
   else if(Normalize(a.Text,o)!=Normalize(b.Text,o)||a.Kind!=b.Kind||!a.Cells.Select(c=>c.Span).SequenceEqual(b.Cells.Select(c=>c.Span)))r.Kind=ChangeKind.Modified;
   else if(a.Media!=b.Media)r.Kind=ChangeKind.Image;
   else if(o.Formatting&&FormatKey(a,o)!=FormatKey(b,o))r.Kind=ChangeKind.Formatting;else r.Kind=ChangeKind.Same;
   if(a!=null&&b!=null&&r.Kind==ChangeKind.Modified)Inline(a.Text,b.Text,o,out r.LeftMask,out r.RightMask);return r;
  }
  static void Region(List<Unit> a,int a0,int ae,List<Unit> b,int b0,int be,CompareOptions o,List<DiffRow> rows,CancellationToken ct){
   int n=ae-a0,m=be-b0;if(n==0){for(int j=b0;j<be;j++)rows.Add(Row(null,b[j],o));return;}if(m==0){for(int i=a0;i<ae;i++)rows.Add(Row(a[i],null,o));return;}
   if((long)n*m>100000){for(int k=0;k<Math.Max(n,m);k++)rows.Add(Row(k<n?a[a0+k]:null,k<m?b[b0+k]:null,o));return;}
   var score=new double[n+1,m+1];var take=new byte[n,m];
   for(int i=n-1;i>=0;i--){ct.ThrowIfCancellationRequested();for(int j=m-1;j>=0;j--){double sim=Similarity(a[a0+i],b[b0+j],o);double paired=sim>=.24?score[i+1,j+1]+sim+0.2:-1;double del=score[i+1,j],ins=score[i,j+1];if(paired>=del&&paired>=ins){score[i,j]=paired;take[i,j]=1;}else if(del>=ins){score[i,j]=del;take[i,j]=2;}else{score[i,j]=ins;take[i,j]=3;}}}
   int x=0,y=0;while(x<n||y<m){if(x==n){rows.Add(Row(null,b[b0+y++],o));}else if(y==m){rows.Add(Row(a[a0+x++],null,o));}else if(take[x,y]==1)rows.Add(Row(a[a0+x++],b[b0+y++],o));else if(take[x,y]==2)rows.Add(Row(a[a0+x++],null,o));else rows.Add(Row(null,b[b0+y++],o));}
  }
  public static Comparison Compare(WordDoc left,WordDoc right,CompareOptions o,CancellationToken ct){
   var c=new Comparison{Left=left,Right=right,Options=o};var parts=left.Units.Select(u=>u.Part).Concat(right.Units.Select(u=>u.Part)).Distinct();
   foreach(string part in parts){if(!o.Extras&&part!="正文")continue;
    Func<Unit,bool> include=u=>u.Part==part&&!(o.IgnoreEmpty&&string.IsNullOrWhiteSpace(u.Text)&&u.Pictures.Count==0);
    var a=left.Units.Where(include).ToList();var b=right.Units.Where(include).ToList();var matches=Matches(a.Select(u=>Key(u,o)).ToArray(),b.Select(u=>Key(u,o)).ToArray(),ct);int x=0,y=0;
    foreach(var p in matches){Region(a,x,p.Item1,b,y,p.Item2,o,c.Rows,ct);c.Rows.Add(Row(a[p.Item1],b[p.Item2],o));x=p.Item1+1;y=p.Item2+1;}Region(a,x,a.Count,b,y,b.Count,o,c.Rows,ct);
   }
   for(int i=0;i<c.Rows.Count;i++)c.Rows[i].Index=i;return c;
  }
  public static void Inline(string a,string b,CompareOptions o,out bool[] am,out bool[] bm){
   am=Enumerable.Repeat(true,a.Length).ToArray();bm=Enumerable.Repeat(true,b.Length).ToArray();
   var aa=Tokens(a,o);var bb=Tokens(b,o);
   foreach(var p in Matches(aa.Select(t=>Normalize(t.Item3,o)).ToArray(),bb.Select(t=>Normalize(t.Item3,o)).ToArray(),CancellationToken.None)){
    for(int k=0;k<aa[p.Item1].Item2;k++)am[aa[p.Item1].Item1+k]=false;
    for(int k=0;k<bb[p.Item2].Item2;k++)bm[bb[p.Item2].Item1+k]=false;
   }
   if(o.IgnoreSpace){for(int i=0;i<a.Length;i++)if(char.IsWhiteSpace(a[i]))am[i]=false;for(int i=0;i<b.Length;i++)if(char.IsWhiteSpace(b[i]))bm[i]=false;}
  }
  static List<Tuple<int,int,string>> Tokens(string text,CompareOptions o){
   var result=new List<Tuple<int,int,string>>();int start=0;
   while(start<text.Length){int end=start+StringInfo.GetNextTextElement(text,start).Length;
    if(o.Highlight==HighlightLevel.Paragraph)end=text.Length;
    else if(o.Highlight==HighlightLevel.Word&&Regex.IsMatch(text.Substring(start,1),@"[A-Za-z0-9_]")){while(end<text.Length&&Regex.IsMatch(text.Substring(end,1),@"[A-Za-z0-9_]"))end++;}
    else if(o.Highlight==HighlightLevel.Sentence){while(end<text.Length){char previous=text[end-1];if("。！？!?；;\n".IndexOf(previous)>=0||(previous=='.'&&(end==text.Length||char.IsWhiteSpace(text[end]))))break;end+=StringInfo.GetNextTextElement(text,end).Length;}}
    string token=text.Substring(start,end-start);if(!o.IgnoreSpace||!string.IsNullOrWhiteSpace(token))result.Add(Tuple.Create(start,end-start,token));start=end;
   }return result;
  }
  public static void Rehighlight(Comparison c,HighlightLevel level){c.Options.Highlight=level;foreach(var r in c.Rows.Where(r=>r.Kind==ChangeKind.Modified))Inline(r.Left.Text,r.Right.Text,c.Options,out r.LeftMask,out r.RightMask);}
 }
}
