using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using StarWarsUnlimitedDeckBrowser.Models;
using CsvHelper;
using System.Globalization;

namespace StarWarsUnlimitedDeckBrowser.CollectionParser
{
    public static class Parser
    {
        public static List<CollectionEntity> ParseCollection(string fileLocation) 
        {
            XSSFWorkbook wb = new(fileLocation);

            ISheet sheet = wb.GetSheetAt(0);

            List<CollectionEntity> dataList = [];

            for (int i = 3; i < 271; i++)
            {
                IRow row = sheet.GetRow(i);

                if (row != null) 
                {
                    dataList.Add(new CollectionEntity
                    {
                        Set = "SOR",
                        CardNumber = row.GetCell(0).StringCellValue,
                        Count = (int)row.GetCell(1).NumericCellValue + (int)row.GetCell(2).NumericCellValue + (int)row.GetCell(3).NumericCellValue + (int)row.GetCell(4).NumericCellValue,
                        IsFoil = "FALSE"
                    });
                }
            }

            return dataList;

        }

        public static void ConvertCollectionToCsv(string fileLocation, List<CollectionEntity> data)
        {
            using var writer = new StreamWriter(fileLocation);
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.WriteRecords(data);
        }
    }
}
