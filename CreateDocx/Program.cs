using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.Diagnostics;
using System.Drawing;

namespace CreateDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"D:\DocXExample.doc";
            string title = "標題";
            string content = "內文。";


            //設定標題字體格式
            var headLineFormat = new Formatting();
            headLineFormat.Size = 18D;
            headLineFormat.Position = 12;

            //設定內文
            var paraFormat = new Formatting();
            paraFormat.Size = 10D;

            //產生文件於記憶體中
            var doc = DocX.Create(fileName);

            //插入標題及內文
            doc.InsertParagraph(title, false, headLineFormat);
            doc.InsertParagraph(content, false, paraFormat);

            //宣告一個兩列三欄的表格
            Table table = doc.AddTable(2, 3);

            // 將表格置中
            table.Alignment = Alignment.center;

            //設定表格樣式為普通樣式(無邊框)
            table.Design = TableDesign.TableNormal;

            // 添加內容至表格中
            table.Rows[0].Cells[0].Paragraphs.First().Append("A");
            table.Rows[0].Cells[1].Paragraphs.First().Append("B");
            table.Rows[0].Cells[2].Paragraphs.First().Append("C");
            table.Rows[1].Cells[0].Paragraphs.First().Append("D");
            table.Rows[1].Cells[1].Paragraphs.First().Append("E");
            table.Rows[1].Cells[2].Paragraphs.First().Append("F");

            //將表格塞入文件
            doc.InsertTable(table);

            //儲存文件
            doc.Save();

            //打開該文件
            Process.Start("WINWORD.EXE", fileName);

            Console.WriteLine("Word 文件產生完畢");
            Console.ReadLine();
        }
    }
}
