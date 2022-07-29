using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Drawing;
using QRCoder;

namespace meteorCRMExport
{
    public partial class deliveryLabel : System.Web.UI.Page
    {
        public Bitmap code(string msg)
        {
            QRCode qrCode = new QRCode(new QRCodeGenerator().CreateQrCode(msg, QRCodeGenerator.ECCLevel.M));
            Bitmap bitmap = new Bitmap(this.Server.MapPath("~/").ToString().Trim() + "image\\icon.jpg");
            Color black = Color.Black;
            Color transparent = Color.Transparent;
            Bitmap icon = bitmap;
            int iconSizePercent = 20;
            int iconBorderWidth = 2;
            return qrCode.GetGraphic(2, black, transparent, icon, iconSizePercent, iconBorderWidth, false);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            var request = HttpContext.Current.Request;

            if (request.InputStream.Length == 0)
            {
                Response.ContentType = "text/html";
                Response.Write("数据错误!");
                Response.End();
            }

            byte[] requestData = new byte[request.InputStream.Length];

            request.InputStream.Read(requestData, 0, (int)request.InputStream.Length);

            var jsonData = Encoding.UTF8.GetString(requestData);

            dynamic m = JsonConvert.DeserializeObject<dynamic>(jsonData);

            var _id = m._id;
            var companyId = m.companyId;
            var customerName = m.customerName;
            var userName = m.userName;
            var departmentName = m.departmentName;
            var customerOrderNo = m.customerOrderNo;
            var salesOrderNo = m.salesOrderNo;
            var number = m.number;
            var products = m.products;

            OutputExcel(_id, companyId, customerName, userName, departmentName, customerOrderNo, salesOrderNo, number, products);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic _id, dynamic companyId, dynamic customerName, dynamic userName, dynamic departmentName, dynamic customerOrderNo, dynamic salesOrderNo, dynamic number, dynamic products)
        {
            GC.Collect();
            Application excel;// = new Application(); 
            int rowIndex = 0;

            _Workbook xBk;
            _Worksheet xSt;

            excel = new Application();

            excel.DisplayAlerts = false;

            xBk = excel.Workbooks.Add(true);

            xSt = (_Worksheet)xBk.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);

            xSt.Name = "送货标签";

            xSt.PageSetup.LeftMargin = 0.0;
            xSt.PageSetup.RightMargin = 0.0;
            xSt.PageSetup.HeaderMargin = 0.0;
            xSt.PageSetup.FooterMargin = 0.0;
            xSt.PageSetup.TopMargin = 0.0;
            xSt.PageSetup.BottomMargin = 0.0;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 10;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 0.85f;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 35.0f;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 0.85f;

            for (int i = 0; i < Convert.ToInt32(number); i++)
            {
                Bitmap bitmap1 = this.code("http://erp1.mro9.com/app?reurl=/Product/Product/updatePictureBySalesOrderProductId/" + products[i]["_id"].ToString());
                bitmap1.Save(this.Server.MapPath("~/").ToString().Trim() + "temp\\" + products[i]["_id"].ToString() + ".png");
            }

            for (int i = 0; i < Convert.ToInt32(number); i++) {
                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 7.5f;

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 10.5f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = "行号：" + products[i]["sortNum"].ToString();
                float num = (float)i * 171.75f + 8f;
                xSt.Shapes.AddPicture(this.Server.MapPath("~/").ToString().Trim() + "image\\" + companyId.ToString() + ".gif", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 99f, num, 125f, 10f);

                num = (float)i * 171.75f + 20.0f;
                xSt.Shapes.AddPicture(this.Server.MapPath("~/").ToString().Trim() + "temp\\" + products[i]["_id"].ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 184, num, 40, 40);

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 13.5f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = "包装：" + products[i]["tag"].ToString();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 15f;

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 13.5f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 12;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = customerName.ToString().Trim();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 13.5f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 12;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = customerOrderNo.ToString().Trim();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 13.5f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 12;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = userName.ToString().Trim();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 6f;

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 12f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 8;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = "型号：" + products[i]["proNo"].ToString();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 25.5f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].WrapText = true;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 8;
                excel.Cells[rowIndex, 2] = "描述：" + products[i]["proName"].ToString();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 12.75f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = "数量：" + products[i]["quantity"].ToString();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 11.25f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 8;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                excel.Cells[rowIndex, 2] = "备注：" + products[i]["storehouseName"].ToString();

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 9.75f;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].Font.Size = 7;
                excel.Cells[rowIndex, 2] = salesOrderNo.ToString();
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                rowIndex++;
                xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 3]].RowHeight = 7.5f;
            }

            xSt.PageSetup.PrintArea = "A1:C" + Convert.ToString(rowIndex);

            excel.Visible = true;

            string path = Server.MapPath("~/") + "temp\\" + _id.ToString() + "_deliveryLabel.xlsx";
            string filename = _id.ToString() + "_deliveryLabel.xlsx";

            //保存excel
            xBk.SaveCopyAs(path);

            //ds = null;
            xBk.Close(false, null, null);

            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xSt);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xBk);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            try
            {
                if (excel != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out lpdwProcessId);
                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
            }

            try
            {
                for (int i = 0; i < Convert.ToInt32(number); i++)
                {
                    if (File.Exists(this.Server.MapPath("~/").ToString().Trim() + "temp\\" + products[i]["_id"].ToString() + ".png"))
                    {
                        File.Delete(this.Server.MapPath("~/").ToString().Trim() + "temp\\" + products[i]["_id"].ToString() + ".png");
                    }
                }
            }
            catch (Exception ex)
            {

            }

            //

            xBk = null;
            excel = null;
            xSt = null;
            GC.Collect();


            Response.Clear();
            Response.ContentType = "text/html";
            Response.Write(filename);
            Response.End();
        }
    }
}