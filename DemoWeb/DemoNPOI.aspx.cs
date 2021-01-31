using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DemoWeb
{
    public partial class DemoNPOI : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void downloadExcel(object sender, EventArgs e)
        {
            // khởi tạo wb rỗng
            XSSFWorkbook wb = new XSSFWorkbook();

            // Tạo ra 1 sheet
            ISheet sheet = wb.CreateSheet();

            // Bắt đầu ghi lên sheet

            // Tạo row
            var row0 = sheet.CreateRow(0);
            // Merge lại row đầu 3 cột
            row0.CreateCell(0); // tạo ra cell trc khi merge
            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 2);
            sheet.AddMergedRegion(cellMerge);
            row0.GetCell(0).SetCellValue("Thông tin sinh viên");

            // Ghi tên cột ở row 1
            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue("MSSV");
            row1.CreateCell(1).SetCellValue("Tên");
            row1.CreateCell(2).SetCellValue("Phone");

            // bắt đầu duyệt mảng và ghi tiếp tục
            int rowIndex = 2;

            for (int rowIdx = 1; rowIdx < 20000; rowIdx++)
            {
                var newRow = sheet.CreateRow(rowIdx);
                for (int colInx = 0; colInx < 10; colInx++)
                {
                    newRow.CreateCell(colInx).SetCellValue("AAA");
                }
            }


            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Sample1.xlsx");
            wb.Write(Response.OutputStream);
            Response.End();

        }
    }
}