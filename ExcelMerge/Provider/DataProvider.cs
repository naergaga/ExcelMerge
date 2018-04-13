using ExcelMerge.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.IO;
using OfficeOpenXml;

namespace ExcelMerge.Provider
{
    public class DataProvider
    {

        public List<Book> GetList(string dir)
        {
            List<Book> list = new List<Book>();
            foreach (var file in Directory.GetFiles(dir,"*.xls",SearchOption.AllDirectories))
            {
                list.Add(Get(file));
            }
            return list;
        }

        public Book Merge(List<Book> bookList)
        {
            var book1 = new Book { SheetList = new List<Sheet>() };
            var sheet1 = new Sheet { Rows = new List<List<object>>() };
            book1.SheetList.Add(sheet1);

            foreach (var book in bookList)
            {
                foreach (var row in book.SheetList[0].Rows)
                {
                    sheet1.Rows.Add(row);
                }
            }
            return book1;
        }

        public void Export(Book book, string path)
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("统计");
                var rowIndex = 1;
                foreach (var row in book.SheetList[0].Rows)
                {
                    var columnIndex = 1;
                    foreach (var value in row)
                    {
                        ws.Cells[rowIndex, columnIndex].Value = value;
                        columnIndex++;
                    }
                    rowIndex++;
                }
                p.SaveAs(new FileInfo(path));
            }
        }

        public Book Get(string path)
        {
            var book = new Book { Path = path, SheetList = new List<Sheet>() };
            using (var reader = ExcelReaderFactory.CreateReader(File.OpenRead(path)))
            {
                do
                {
                    var sheet = new Sheet { Rows = new List<List<object>>() };

                    while (reader.Read())
                    {
                        var row = new List<object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row.Add(reader.GetValue(i));
                        }

                        if (row[1] != null)
                            sheet.Rows.Add(row);
                    }
                    book.SheetList.Add(sheet);
                    break;
                } while (reader.NextResult());
            }
            return book;
        }

    }
}
