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
    public partial class salesInvoice : System.Web.UI.Page
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

            var orderModel = m.orderModel;
            var customerModel = m.customerModel;
            var productModel = m.productModel;

            OutputExcel(orderModel, customerModel, productModel);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic customerModel, dynamic productModel)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            xSt.PageSetup.LeftMargin = 0 / 0.035;
            xSt.PageSetup.RightMargin = 0 / 0.035;
            xSt.PageSetup.HeaderMargin = 0.8 / 0.035;
            xSt.PageSetup.FooterMargin = 0.8 / 0.035;
            xSt.PageSetup.TopMargin = 0.8 / 0.035;
            xSt.PageSetup.BottomMargin = 0.8 / 0.035;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 8;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 9], excel.Cells[1, 9]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 10], excel.Cells[1, 10]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 11], excel.Cells[1, 11]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 12], excel.Cells[1, 12]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 13], excel.Cells[1, 13]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 14], excel.Cells[1, 14]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 15], excel.Cells[1, 15]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 16], excel.Cells[1, 16]].ColumnWidth = 5;
            xSt.Range[excel.Cells[1, 17], excel.Cells[1, 17]].ColumnWidth = 7;
            xSt.Range[excel.Cells[1, 18], excel.Cells[1, 18]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 19], excel.Cells[1, 19]].ColumnWidth = 7;
            xSt.Range[excel.Cells[1, 20], excel.Cells[1, 20]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 21], excel.Cells[1, 21]].ColumnWidth = 13;
            xSt.Range[excel.Cells[1, 22], excel.Cells[1, 22]].ColumnWidth = 5;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 22]].RowHeight = 1;

            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 22]].RowHeight = 30;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 22]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 22]].Merge(false);
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 22]].Font.Size = 13;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 22]].Font.Bold = true;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 22]].Value2 = orderModel["orderNo"] + " 发票申请单";

            xSt.Range[excel.Cells[3, 1], excel.Cells[6, 22]].RowHeight = 24;

            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 3]].Merge(false);
            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 3]].Value2 = "开票公司：" + orderModel["companyName"].ToString().Trim();

            xSt.Range[excel.Cells[3, 6], excel.Cells[3, 11]].Merge(false);
            xSt.Range[excel.Cells[3, 6], excel.Cells[3, 11]].Value2 = "客户名称：" + customerModel["customerName"].ToString().Trim();

            xSt.Range[excel.Cells[3, 15], excel.Cells[3, 22]].Merge(false);
            //if (Convert.ToBoolean(orderModel["invoiceType"].ToString()))
            //{
            //    xSt.Range[excel.Cells[3, 14], excel.Cells[3, 21]].Value2 = "发票类型：增值税发票";
            //} 
            //else
            //{
            //    xSt.Range[excel.Cells[3, 14], excel.Cells[3, 21]].Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            //    xSt.Range[excel.Cells[3, 14], excel.Cells[3, 21]].Value2 = "发票类型：普通发票";
            //}
            xSt.Range[excel.Cells[3, 15], excel.Cells[3, 22]].Value2 = "发票类型：" + orderModel["invoiceType"];

            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 3]].Merge(false);
            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 3]].Value2 = "价税合计：" + Convert.ToDecimal(orderModel["total"].ToString()).ToString("N2");

            //xSt.PageSetup.PrintTitleRows = "$8:$8";  标题行 
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 22]].RowHeight = 22;
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 14]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[7, 1] = "订单号";
            excel.Cells[7, 2] = "询/报型号";
            excel.Cells[7, 3] = "询/报名称";
            excel.Cells[7, 4] = "物料号";
            excel.Cells[7, 5] = "订货号";
            excel.Cells[7, 6] = "描述";
            excel.Cells[7, 7] = "税收分类编码";
            excel.Cells[7, 8] = "税收分类名称";
            excel.Cells[7, 9] = "进项型号";
            excel.Cells[7, 10] = "进项名称";
            excel.Cells[7, 11] = "产品—税收分类编码";
            excel.Cells[7, 12] = "产品—税收分类名称";
            excel.Cells[7, 13] = "产品—开票型号";
            excel.Cells[7, 14] = "产品—开票名称";
            excel.Cells[7, 15] = "单位";
            excel.Cells[7, 16] = "订单数量";
            excel.Cells[7, 17] = "未税单价";
            excel.Cells[7, 18] = "未税总价";
            excel.Cells[7, 19] = "含税单价";
            excel.Cells[7, 20] = "含税总价";
            excel.Cells[7, 21] = "备注";
            excel.Cells[7, 22] = "采购人";

            int j = 7;
            decimal totalNoTax = 0;
            decimal total = 0;
            string customerOrderNo = "";
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                j++;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 22]].WrapText = true;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 14]].NumberFormat = "@";
                excel.Cells[j, 1] = productModel[i]["Sales_InvoiceProduct"]["customerOrderNo"].ToString().Trim();
                if (customerOrderNo.IndexOf(productModel[i]["Sales_InvoiceProduct"]["customerOrderNo"].ToString().Trim()) == -1)
                {
                    customerOrderNo += productModel[i]["Sales_InvoiceProduct"]["customerOrderNo"].ToString().Trim() + '；';
                }
                excel.Cells[j, 2] = productModel[i]["Sales_InvoiceProduct"]["quotationProductModel"].ToString().Trim();
                excel.Cells[j, 3] = productModel[i]["Sales_InvoiceProduct"]["quotationProductName"].ToString().Trim();
                excel.Cells[j, 4] = productModel[i]["Sales_OrderProduct"]["materialNo"]==null ? "" : productModel[i]["Sales_OrderProduct"]["materialNo"].ToString().Trim();
                excel.Cells[j, 5] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
                excel.Cells[j, 6] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim()));
                try
                {
                    excel.Cells[j, 7] = Convert.ToString(productModel[i]["Purchase_InvoiceProduct"]["taxEncodingNo"].ToString().Trim() + "00000000000000000000").Substring(0, 19).Replace("0000000000000000000", "");
                    excel.Cells[j, 8] = productModel[i]["Purchase_InvoiceProduct"]["taxEncoding"].ToString().Trim();
                    excel.Cells[j, 9] = productModel[i]["Purchase_InvoiceProduct"]["taxNo"].ToString().Trim();
                    excel.Cells[j, 10] = productModel[i]["Purchase_InvoiceProduct"]["taxName"].ToString().Trim();
                }
                catch { }
                excel.Cells[j, 11] = Convert.ToString(productModel[i]["Product_Product"]["taxEncodingNo"].ToString().Trim() + "00000000000000000000").Substring(0, 19).Replace("0000000000000000000", "");
                excel.Cells[j, 12] = productModel[i]["Product_Product"]["taxEncoding"].ToString().Trim();
                excel.Cells[j, 13] = productModel[i]["Product_Product"]["taxNo"].ToString().Trim();
                excel.Cells[j, 14] = productModel[i]["Product_Product"]["taxName"].ToString().Trim();

                xSt.Range[excel.Cells[j, 15], excel.Cells[j, 16]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j, 15] = productModel[i]["Product_Product"]["unit"].ToString().Trim();
                excel.Cells[j, 16] = (string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_InvoiceProduct"]["quantity"].ToString()).ToString("N2"));

                xSt.Range[excel.Cells[j, 17], excel.Cells[j, 20]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                excel.Cells[j, 17] = (string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["unitPriceNoTax"].ToString()) ? 0 : productModel[i]["Sales_InvoiceProduct"]["unitPriceNoTax"].ToString());
                excel.Cells[j, 18] = (string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["totalPriceNoTax"].ToString()) ? 0 : productModel[i]["Sales_InvoiceProduct"]["totalPriceNoTax"].ToString());
                if (!string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["totalPriceNoTax"].ToString()))
                {
                    totalNoTax += Convert.ToDecimal(productModel[i]["Sales_InvoiceProduct"]["totalPriceNoTax"].ToString());
                }
                excel.Cells[j, 19] = (string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["unitPrice"].ToString()) ? 0 :productModel[i]["Sales_InvoiceProduct"]["unitPrice"].ToString());
                excel.Cells[j, 20] = (string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["totalPrice"].ToString()) ? 0 : productModel[i]["Sales_InvoiceProduct"]["totalPrice"].ToString());
                if (!string.IsNullOrEmpty(productModel[i]["Sales_InvoiceProduct"]["totalPrice"].ToString()))
                {
                    total += Convert.ToDecimal(productModel[i]["Sales_InvoiceProduct"]["totalPrice"].ToString());
                }

                excel.Cells[j, 21] = productModel[i]["Sales_InvoiceProduct"]["remark"].ToString().Trim();

                xSt.Range[excel.Cells[j, 22], excel.Cells[j, 21]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                try
                {
                    excel.Cells[j, 22] = productModel[i]["Purchase_Order"]["personInChargeName"].ToString().Trim();
                }
                catch
                { }
                

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 22]].EntireRow.AutoFit();//行高根据内容自动调整

                //如果不换行，行高自适应后，行高小于23，则最低行高为23
                if (Convert.ToInt32(xSt.Range[excel.Cells[j, 1], excel.Cells[j, 22]].RowHeight) < 22)
                {
                    xSt.Range[excel.Cells[j, 1], excel.Cells[j, 22]].RowHeight = 22;
                }
            }

            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].HorizontalAlignment = XlHAlign.xlHAlignRight;//设置标题格式为居中对齐 
            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].Formula = "=SUM(R8:R" + j + ")";

            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].HorizontalAlignment = XlHAlign.xlHAlignRight;//设置标题格式为居中对齐 
            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].Formula = "=SUM(T8:T" + j + ")";

            xSt.Range[excel.Cells[7, 1], excel.Cells[j, 22]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[7, 1], excel.Cells[j, 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 22]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[7, 22], excel.Cells[j, 22]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 22]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j + 1, 18], excel.Cells[j + 1, 18]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j + 1, 20], excel.Cells[j + 1, 20]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 22]].RowHeight = 22;

            xSt.Range[excel.Cells[4, 6], excel.Cells[4, 11]].Merge(false);
            xSt.Range[excel.Cells[4, 6], excel.Cells[4, 11]].Value2 = "税前金额：" + totalNoTax.ToString();

            xSt.Range[excel.Cells[4, 15], excel.Cells[4, 22]].Merge(false);
            xSt.Range[excel.Cells[4, 15], excel.Cells[4, 22]].Value2 = "税　　额：" + (total - totalNoTax).ToString();

            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 22]].Merge(false);
            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 22]].Value2 = "摘　　要：" + orderModel["summary"];
            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 22]].Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 22]].Font.Bold = true;

            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 22]].Merge(false);
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 22]].Value2 = "客户订单号：" + customerOrderNo.TrimEnd('；');

            excel.Visible = true;

            string path = "";
            string filename = "";

            filename = orderModel["_id"] + ".xlsx";
            path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + ".xlsx";

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