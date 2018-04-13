using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge.Data
{
    public class Book
    {
        public string Path { get; set; }
        public List<Sheet> SheetList { get; set; }
    }
}
