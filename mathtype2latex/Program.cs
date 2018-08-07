using System;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Vml;

namespace mathtype2latex
{
    class Program
    {
        static void Main(string[] args)
        {
            String wordPath = @"C:\Users\dxw\Desktop\123";//word在哪里
            String jpgPathPro = @"C:\Users\dxw\Desktop\jpg\";//jpg放哪里
            var files = Directory.GetFiles(wordPath, "*.docx");
            //在当前文件夹中遍历所有文档
            foreach (var file in files)
            {
                WordprocessingDocument docx = WordprocessingDocument.Open(file, true);
                foreach(var image in docx.MainDocumentPart.Document.Body.Descendants<ImageData>())
                {
                    ImagePart p = docx.MainDocumentPart.GetPartById(image.RelationshipId) as ImagePart;
                    int hash = p.GetHashCode();
                    Stream stream = p.GetStream();
                    byte[] bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, bytes.Length);
                    stream.Seek(0, SeekOrigin.Begin);
                    string jpgPath = jpgPathPro + hash + ".wmf";
                    FileStream fs = new FileStream(jpgPath, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    bw.Write(bytes);
                    bw.Close();
                    fs.Close();
                }

            }
        }
        
    }
}
