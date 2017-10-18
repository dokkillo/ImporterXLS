using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ImporterXLS.Utils;


namespace ImporterXLS
{
    public class ImporterXLS<T> where T : class, new()
    {
        public IEnumerable<T> Load(string PathFile, int SheetPage = 1, int HeaderRow = 1)
        {
            List<T> ObjList = new List<T>();
            LoadExcel loadExcel = new LoadExcel();
            var dicto = loadExcel.OpenExcelFile(PathFile, SheetPage, HeaderRow);           

            foreach (var item in dicto)
            {
                ObjList.Add(item.ToObject<T>());
            }

            return ObjList;
        }

    }
}
