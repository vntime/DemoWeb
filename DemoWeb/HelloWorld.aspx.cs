using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DemoWeb
{
    public partial class HelloWorld : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void uploadFile(object sender, EventArgs e)
        {
            if (FileUpload1.PostedFile == null || FileUpload1.PostedFile.ContentLength == 0)
            {
                lb_statusUpload.Text = "Chua chon file!";
                return;
            }

            string fileName = System.IO.Path.GetFileName(FileUpload1.PostedFile.FileName);
            System.Diagnostics.Debug.WriteLine(FileUpload1.PostedFile.FileName);
            System.Diagnostics.Debug.WriteLine(fileName);

            string savePath = Server.MapPath(fileName);
            FileUpload1.PostedFile.SaveAs(savePath);

            List<List<String>> arr = ReadExcel(savePath);

            for (int i = 0; i < arr.Count; i++)
            {
                List<String> row = arr[i];
                for (int j = 0; j < row.Count; j++)
                {
                    System.Diagnostics.Debug.WriteLine(row[j]);
                }
            }

            File.Delete(savePath);


            //Response Excel
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Sample1");
            for (int i = 0; i < 200000; i++)
            {
                for (int j = 0; j < 1000; j++)
                {
                    //string c = Char.ConvertFromUtf32(j + 1 + 64);
                    //string addr = c + (i + 1);
                    try
                    {
                        ws.Cells[i+1,j+1].Value = "AAAAAAAAAAAAAAAAAAA";
                    }
                    catch (Exception)
                    {
                        System.Diagnostics.Debug.WriteLine("i="+i);
                        System.Diagnostics.Debug.WriteLine("j=" + j);
                        throw;
                    }
                    //ws.Cells[addr].Style.Font.Bold = true;
                }
            }


            //pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Sample1.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());  // send the file
            Response.End();
        }

        static List<List<string>> ReadExcel(string pathFile)
        {
            List<List<string>> rows = new List<List<string>>();
            using (var package = new ExcelPackage(new FileInfo(pathFile)))
            {
                //var firstSheet = package.Workbook.Worksheets["First Sheet"];
                var firstSheet = package.Workbook.Worksheets.First();
                int rowCnt = firstSheet.Dimension.Rows;
                int colCnt = firstSheet.Dimension.Columns;


                for (int i = 0; i < rowCnt; i++)
                {
                    List<string> row = new List<string>();
                    for (int j = 0; j < colCnt; j++)
                    {
                        var cell = firstSheet.Cells[i + 1, j + 1];
                        var cellValue = "";
                        if (cell != null && cell.Value != null)
                        {
                            cellValue = cell.Value.ToString();
                        }
                        row.Add(cellValue);
                    }
                    rows.Add(row);
                }
            }
            return rows;
        }
    }
}