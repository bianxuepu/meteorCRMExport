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

namespace meteorCRMExport
{
    public partial class deliveryNote : System.Web.UI.Page
    {
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

            var filetype = Request.QueryString["type"];
            if (filetype == null || filetype == "")
            {
                filetype = "excel";
            }

            var orderModel = m.orderModel;
            var companyModel = m.companyModel;
            var customerModel = m.customerModel;
            var productModel = m.productModel;
            var userModel = m.userModel;

            OutputExcel(orderModel, companyModel, customerModel, productModel, userModel, filetype);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic companyModel, dynamic customerModel, dynamic productModel, dynamic userModel, string strType)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.Name = "送货单";
            xSt.PageSetup.LeftMargin = 360.0 / 7.0;
            xSt.PageSetup.RightMargin = 360.0 / 7.0;
            xSt.PageSetup.HeaderMargin = 0.0;
            xSt.PageSetup.FooterMargin = 0.0;
            xSt.PageSetup.TopMargin = 360.0 / 7.0;
            xSt.PageSetup.BottomMargin = 400.0 / 7.0;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 9;
            excel.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 5;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 25;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 9;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 13;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 8]].RowHeight = 24.75;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Merge(false);
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Font.Size = 14;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Font.Bold = true;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Value2 = "送货清单";
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            xSt.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "/image/" + companyModel["nid"].ToString().Trim() + ".gif", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 380, 4, 150, 13);

            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 8]].RowHeight = 7.5;

            xSt.Range[excel.Cells[3, 1], excel.Cells[9, 8]].RowHeight = 15.75;
            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Merge(false);
            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Value2 = "收货单位：";

            xSt.Range[excel.Cells[3, 3], excel.Cells[3, 4]].Merge(false);
            xSt.Range[excel.Cells[3, 3], excel.Cells[3, 4]].Value2 = customerModel["customerName"];

            xSt.Cells[3, 5] = "供货单位：";
            xSt.Range[excel.Cells[3, 6], excel.Cells[3, 8]].Merge(false);
            xSt.Range[excel.Cells[3, 6], excel.Cells[3, 8]].Value2 = companyModel.name;

            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Merge(false);
            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Value2 = "收货信息：";

            xSt.Range[excel.Cells[4, 3], excel.Cells[6, 4]].Merge(false);
            xSt.Range[excel.Cells[4, 3], excel.Cells[6, 4]].Value2 = orderModel["fhxx"];

            xSt.Cells[4, 5].Font.Bold = true;
            xSt.Cells[4, 5] = "联系人：";
            xSt.Range[excel.Cells[4, 6], excel.Cells[4, 8]].Merge(false);
            xSt.Range[excel.Cells[4, 6], excel.Cells[4, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[4, 6], excel.Cells[4, 8]].Font.Bold = true;
            xSt.Range[excel.Cells[4, 6], excel.Cells[4, 8]].Value2 = userModel["name"];

            xSt.Cells[5, 5].Font.Bold = true;
            xSt.Cells[5, 5] = "手机：";
            xSt.Range[excel.Cells[5, 6], excel.Cells[5, 8]].Merge(false);
            xSt.Range[excel.Cells[5, 6], excel.Cells[5, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[5, 6], excel.Cells[5, 8]].Font.Bold = true;
            xSt.Range[excel.Cells[5, 6], excel.Cells[5, 8]].Value2 = userModel["mobilePhone"];

            xSt.Cells[6, 5].Font.Bold = true;
            xSt.Cells[6, 5] = "服务专线：";
            xSt.Range[excel.Cells[6, 6], excel.Cells[6, 8]].Merge(false);
            xSt.Range[excel.Cells[6, 6], excel.Cells[6, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[6, 6], excel.Cells[6, 8]].Font.Bold = true;
            xSt.Range[excel.Cells[6, 6], excel.Cells[6, 8]].Value2 = "400-816-1658 转 " + userModel["extensionNum"];

            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 2]].Merge(false);
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 2]].Value2 = "客户订单号：";
            xSt.Range[excel.Cells[7, 3], excel.Cells[7, 8]].Merge(false);
            xSt.Range[excel.Cells[7, 3], excel.Cells[7, 8]].Value2 = orderModel["customerOrderNo"];

            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 2]].Merge(false);
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 2]].Value2 = "申请人：";
            xSt.Range[excel.Cells[8, 3], excel.Cells[8, 4]].Merge(false);
            xSt.Range[excel.Cells[8, 3], excel.Cells[8, 4]].Value2 = orderModel["userName"];
            xSt.Cells[8, 5] = "系统单号：";
            xSt.Range[excel.Cells[8, 6], excel.Cells[8, 8]].Merge(false);
            xSt.Range[excel.Cells[8, 6], excel.Cells[8, 8]].Value2 = orderModel["orderNo"];

            xSt.Range[excel.Cells[9, 1], excel.Cells[9, 2]].Merge(false);
            xSt.Range[excel.Cells[9, 1], excel.Cells[9, 2]].Value2 = "送货日期：";
            xSt.Range[excel.Cells[9, 3], excel.Cells[9, 4]].Merge(false);
            xSt.Range[excel.Cells[9, 3], excel.Cells[9, 4]].Value2 = "";
            xSt.Cells[9, 5] = "页码：";
            xSt.Range[excel.Cells[9, 6], excel.Cells[9, 8]].Merge(false);
            xSt.Range[excel.Cells[9, 6], excel.Cells[9, 8]].NumberFormat = "@";
            xSt.Range[excel.Cells[9, 6], excel.Cells[9, 8]].Value2 = "1 / 1";

            xSt.Range[excel.Cells[10, 1], excel.Cells[10, 8]].RowHeight = 22.5;
            if(orderModel["tbsm"].ToString() != "")
            {
                xSt.Range[excel.Cells[10, 1], excel.Cells[10, 8]].Merge(false);
                xSt.Range[excel.Cells[10, 6], excel.Cells[10, 8]].Font.Bold = true;
                xSt.Range[excel.Cells[10, 6], excel.Cells[10, 8]].Font.Italic = true;
                xSt.Range[excel.Cells[10, 6], excel.Cells[10, 8]].Font.Size = 12;
                xSt.Range[excel.Cells[10, 6], excel.Cells[10, 8]].Value2 = "特别说明：" + orderModel["tbsm"];
            }

            xSt.Range[excel.Cells[11, 1], excel.Cells[11, 8]].RowHeight = 15;
            xSt.Cells[11, 1] = "序号";
            xSt.Range[excel.Cells[11, 2], excel.Cells[11, 3]].Merge(false);
            xSt.Range[excel.Cells[11, 2], excel.Cells[11, 3]].Value2 = "产品货号";
            xSt.Range[excel.Cells[11, 4], excel.Cells[11, 5]].Merge(false);
            xSt.Range[excel.Cells[11, 4], excel.Cells[11, 5]].Value2 = "产品描述";
            xSt.Cells[11, 6] = "送货数量";
            xSt.Cells[11, 7] = "单位";
            xSt.Cells[11, 8] = "备注";

            xSt.Cells[11, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Cells[11, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Cells[11, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[11, 1], excel.Cells[11, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[11, 1], excel.Cells[11, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

            int j = 12;
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                xSt.Cells[j, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Cells[j, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Cells[j, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].RowHeight = 36;
                xSt.Cells[j, 1] = Convert.ToString(i+1);
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Merge(false);
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].WrapText = true;
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].NumberFormat = "@";
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
                xSt.Range[excel.Cells[j, 4], excel.Cells[j, 5]].Merge(false);
                xSt.Range[excel.Cells[j, 4], excel.Cells[j, 5]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";
                excel.Cells[j, 6].NumberFormat = "@";
                xSt.Cells[j, 6] = productModel[i]["Sales_OutboundOrderProduct"]["quantity"].ToString("N2");
                xSt.Cells[j, 7] = productModel[i]["Product_Product"]["unit"];
                xSt.Cells[j, 8] = productModel[i]["Sales_OutboundOrderProduct"]["remark"];

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                j++;
            }

            for (int i = 12 - Convert.ToInt32(orderModel["number"]); i > 0; i--)
            {
                xSt.Cells[j, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Cells[j, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Cells[j, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].RowHeight = 36;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                j++;
            }

            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].RowHeight = 30;
            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 8]].RowHeight = 27.75;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 8]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 4]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 4]].Value2 = "非常感谢您的订购、验货及签收！";
            excel.Cells[j + 1, 6] = "收货方签收：";
            xSt.Cells[j + 1, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j + 1, 6].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 8]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 4]].Merge(false);
            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 4]].Value2 = "百禧百地科技（天津）有限公司（www.mro9.com）";
            excel.Cells[j + 2, 6] = "日期：";
            xSt.Cells[j + 2, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j + 2, 6].Font.Bold = true;
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 8]].Merge(false);
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            string path = "";
            string filename = "";

            if (strType.ToLower() == "excel")
            {
                filename = orderModel["_id"] + "_deliveryNote.xlsx";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + "_deliveryNote.xlsx";

                //保存excel
                xBk.SaveCopyAs(path);
            }
            else
            {
                filename = orderModel["_id"] + "_deliveryNote.pdf";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + "_deliveryNote.pdf";

                //保存pdf
                xBk.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path.Replace(".xlsx", ".pdf"), XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //加密pdf
                //Encry(path, path.Replace(".pdf", "-1.pdf"));
                //path = path.Replace(".pdf", "-1.pdf");
            }



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