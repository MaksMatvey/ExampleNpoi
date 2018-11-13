using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.IO;
using System.Windows;


namespace WpfApp1
{
    class ExcelManager
    {

        public static void CreateChangeExcelFile(ISheet sh, IWorkbook wb, string textFromRowBox, string textFromColumnBox, string textFromStringBox, string textPathBox)
        {
                int textRowBox = int.Parse(textFromRowBox);
                int textColumnBox = TextManager.LettersToNumber(textFromColumnBox);
                var currentRow = sh.CreateRow(textRowBox - 1);
                var currentCell = currentRow.CreateCell(textColumnBox - 1);
                currentCell.SetCellValue(textFromStringBox);
                sh.AutoSizeColumn(textColumnBox - 1);

                if (!File.Exists(textPathBox))
                {
                    File.Delete(textPathBox);
                }

                using (FileStream fs = new FileStream(textPathBox, FileMode.OpenOrCreate))
                {
                    wb.Write(fs);
                    //fs.Seek(0, SeekOrigin.End);
                }

        }
    }
}
