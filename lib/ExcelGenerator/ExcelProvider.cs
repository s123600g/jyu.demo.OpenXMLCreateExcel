using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelGenerator.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ExcelGenerator
{
    // Create a spreadsheet document by providing a file name (Open XML SDK)
    // https://learn.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name

    public class ExcelProvider
    {
        public ExcelProvider()
        {
        }

        /// <summary>
        /// Number型態參考集合
        /// </summary>
        private List<Type> _numberTypeRefs => new List<Type>
        {
            typeof(int),
            typeof(decimal),
            typeof(double),
            typeof(float),
        };

        /// <summary>
        /// 產生Excel檔案流
        /// </summary>
        /// <param name="argDatas">
        /// 預計要產生內容
        /// </param>
        /// <returns>
        /// <see cref="MemoryStream"/>
        /// </returns>
        public byte[] GenSpreadsheetDocumentEntity(
            ExcelContentEntity argDatas
        )
        {
            byte[] result = new byte[] { };

            MemoryStream memoryStream = new MemoryStream();

            // 建立文件實體並開始建置內容，所有內容都放置在記憶體內
            using (
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(
                    memoryStream,
                    SpreadsheetDocumentType.Workbook
                )
            )
            {
                // 建置文件內容實體，就像是開一個全新Excel活頁簿
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // 建置活頁簿內工作表單實體區塊
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(
                    new Sheets()
                );

                #region 欄位資料解析

                uint SheetUnitId = 1;

                foreach (
                   ExcelSheetEntity item in argDatas.Sheets
                )
                {
                    // Step 1. 先初始化工作表單實體
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                    // Step 2. 初始化工作表單資料區實體
                    SheetData sheetData = new SheetData();

                    #region 資料內容產生

                    foreach (
                        var row in item.RowData
                    )
                    {
                        // Step 3. 初始化工作列實體
                        Row sheetRow = new Row();

                        // Step 4. 產生每個欄位資料內容清單
                        var cells = row.Select(value =>
                        {
                            Cell tempCell = new Cell();

                            CellValues dataType = GetCellValueDataType(
                                argData: value
                            );

                            string cellValue = string.Empty;

                            #region 內容處理轉換成字串

                            switch (
                                dataType
                            )
                            {
                                case CellValues.Number:

                                    cellValue = ConvertNumberToString(
                                        argValue: value
                                    );

                                    break;

                                case CellValues.String:

                                    cellValue = value;

                                    break;
                            }

                            #endregion

                            tempCell.CellValue = new CellValue(cellValue);

                            tempCell.DataType = CellValues.String;

                            return tempCell;

                        }).ToList();

                        // Step 4-1. 將欄位資料內容清單新增進工作列
                        sheetRow.Append(cells);

                        // Step 5. 更新工作列
                        sheetData.AppendChild(sheetRow);
                    }

                    #endregion

                    // Step 6. 將更新好工作表單內容注入完整工作表單實體
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    // Step 7. 將建置好工作表單註冊進文件實體
                    sheets.Append(new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), // 每一個工作表單都會有自己獨特識別Id，用此Id來註冊
                        SheetId = SheetUnitId,
                        Name = item.SheetName
                    });

                    SheetUnitId++;
                }

                #endregion

                // 儲存資料操作
                workbookPart.Workbook.Save();
            }

            // 將記憶體指標重新指到開始位置
            memoryStream.Seek(0, SeekOrigin.Begin);

            // 取出存放在記憶體內容
            result = memoryStream.ToArray();

            // 清空記憶體
            memoryStream.Close();

            return result;
        }

        /// <summary>
        /// 判斷欄位內容類型
        /// </summary>
        /// <remarks>
        /// 預設類型為string，當不符合Number，就會是預設類型。
        /// </remarks>
        /// <param name="argData">
        /// 欄位內容
        /// </param>
        /// <returns>
        /// <see cref="CellValues"/>
        /// </returns>
        private CellValues GetCellValueDataType(
            dynamic argData
        )
        {
            CellValues result = CellValues.String;

            Type propertyType = ((Object)argData).GetType();

            #region Number Type

            if (
                _numberTypeRefs.Contains(propertyType)
            )
            {
                result = CellValues.Number;
            }

            #endregion

            return result;
        }

        /// <summary>
        /// 轉換數值型態為字串
        /// </summary>
        /// <param name="argValue"></param>
        /// <returns>
        /// <see cref="string"/>
        /// </returns>
        private string ConvertNumberToString(
            dynamic argValue
        )
        {
            string result = Convert.ToString(
                argValue,
                CultureInfo.CurrentCulture
            );

            return result;
        }
    }
}
