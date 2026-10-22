using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Word_Compare;
class ConversionTests {
 static int Main(string[] args){string root=args[0],output=Path.Combine(root,"验证结果"),temp=OfficeConvert.TempRoot();try{
  string source=Path.Combine(root,"示例","示例_原版.docx");string dest=Path.Combine(output,"旧格式样例");
  OfficeConvert.Convert(source,dest,"doc",Path.Combine(temp,"profile1"),CancellationToken.None);
  string doc=Path.Combine(dest,"示例_原版.doc");var old=DocxReader.Load(doc,CancellationToken.None);var original=DocxReader.Load(source,CancellationToken.None);
  if(!old.Units.Any(u=>u.Text.Contains("0.6～0.8"))||old.Units.Count(u=>u.Kind=="表格行")!=original.Units.Count(u=>u.Kind=="表格行"))throw new Exception("DOC conversion lost text or table rows");
  var c=DiffEngine.Compare(original,DocxReader.Load(Path.Combine(root,"示例","示例_修订版.docx"),CancellationToken.None),new CompareOptions(),CancellationToken.None);
  Reports.Save(c,Path.Combine(output,"PDF对比报告.pdf"),false);
  if(new FileInfo(Path.Combine(output,"PDF对比报告.pdf")).Length<1000)throw new Exception("PDF empty");
  File.WriteAllText(Path.Combine(output,"转换测试结果.txt"),"PASS DOC 旧格式读取：保留中文参数、表格行\r\nPASS PDF 报告导出\r\n",new UTF8Encoding(true));Console.WriteLine("PASS DOC conversion and PDF export");return 0;
 }catch(Exception e){Console.WriteLine(e);File.WriteAllText(Path.Combine(output,"转换测试失败.txt"),e.ToString());return 1;}finally{OfficeConvert.Clean(temp);}}
}
