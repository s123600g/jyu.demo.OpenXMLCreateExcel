using ExcelGenerator;
using ExcelGenerator.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace GenExcelFile
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelProvider excelProvider = new ExcelProvider();

            // 產生Excel 內容
            byte[] fileContent = excelProvider.GenSpreadsheetDocumentEntity(
                argDatas: GetExcelContentSample()
            );

            // 產生Excel檔案
            GenSampleData(
                argData: fileContent
            );

            Console.WriteLine("Done!!");
        }

        public static void GenSampleData(
            byte[] argData
        )
        {
            string defaulPath = Path.Combine(
                 Directory.GetCurrentDirectory(),
                 "sample"
            );

            if (
                !Directory.Exists(
                    defaulPath
                )
            )
            {
                Directory.CreateDirectory(defaulPath);
            }

            using (
                FileStream file = new FileStream(
                    Path.Combine(
                        defaulPath,
                        "MyExcelTest.xlsx"
                    ),
                    FileMode.Create,
                    FileAccess.Write
                )
            )
            {
                file.Write(argData);
            }
        }

        private static ExcelContentEntity GetExcelContentSample()
        {
            ExcelContentEntity data = new ExcelContentEntity();

            List<ExcelSheetEntity> sheetDatas = new List<ExcelSheetEntity>();

            sheetDatas.Add(new ExcelSheetEntity
            {
                SheetName = "MySheet1",
                RowData = new List<List<dynamic>>
                {
                    new List<dynamic>{
                        "編號","姓名"
                    },
                    new List<dynamic>{
                        1,"Test1"
                    },
                    new List<dynamic>{
                        2,"Test2"
                    }
                }
            });

            sheetDatas.Add(new ExcelSheetEntity
            {
                SheetName = "MySheet2",
                RowData = new List<List<dynamic>>
                {
                    new List<dynamic>{
                        "編號","姓名"
                    },
                    new List<dynamic>{
                        3,"Test3"
                    },
                    new List<dynamic>{
                        4,"Test4"
                    }
                }
            });

            data.Sheets = sheetDatas;

            return data;
        }
    }
}
