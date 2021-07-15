using System;
using Microsoft.Office.Interop.Word;


namespace MS_Office
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range r = doc.Range();
            r.Text = "Hello world";
            Table t = doc.Tables.Add(r, 3, 2);
            t.Borders.Enable = 1;
            foreach (Row row in t.Rows)
            {
                foreach (Cell cell in row.Cells)
                {    
                    cell.Range.Font.Name = "Times New Roman";
                    cell.Range.Font.Size = 26;
                    cell.Range.Bold = 1;
                    if (cell.RowIndex == 1 && cell.ColumnIndex == 1) 
                    {
                        cell.Range.Text = "СИГНАЛ1";
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    if (cell.RowIndex == 1 && cell.ColumnIndex == 2)
                    {
                        cell.Range.Text = "20П SMD";
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }
            }
            //app.Documents.Open(@"D:\Doc1.docx");
            try
            {
                doc.Save();
                doc.Close();
                app.Quit();
            }
            catch (Exception e)
            {
                Console.Write("Exception: ");
                Console.WriteLine(e.Message);
            }
            Console.WriteLine("End");
        }
    }
}
