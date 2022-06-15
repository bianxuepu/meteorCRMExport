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
    public partial class quotation : System.Web.UI.Page
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
            var version = Request.QueryString["version"];
            if (filetype == null || filetype == "")
            {
                filetype = "excel";
            }
            
            var quotationModel = m.quotationModel;
            var companyModel = m.companyModel;
            var customerModel = m.customerModel;
            var productModel = m.productModel;
            var userModel = m.userModel;

            if (version == "1")
            {
                OutputExcel(quotationModel, companyModel, customerModel, productModel, userModel, filetype);
            }
            else
            {
                OutputExcel1(quotationModel, companyModel, customerModel, productModel, userModel, filetype);
            }            
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic quotationModel, dynamic companyModel, dynamic customerModel, dynamic productModel, dynamic userModel, string strType)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.PageSetup.LeftMargin = 1.4 / 0.035;
            xSt.PageSetup.RightMargin = 0.5 / 0.035;
            xSt.PageSetup.HeaderMargin = 2 / 0.035;
            xSt.PageSetup.FooterMargin = 1 / 0.035;
            xSt.PageSetup.TopMargin = 3.3 / 0.035;
            xSt.PageSetup.BottomMargin = 1.8 / 0.035;
            xSt.PageSetup.RightHeaderPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"].ToString().Trim() + "logosno.jpg";
            xSt.PageSetup.RightHeader = "&G";
            xSt.PageSetup.LeftHeader = @"&""微软雅黑,Bold""&14" + "报价单";
            xSt.PageSetup.LeftFooterPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + "footline.jpg";
            xSt.PageSetup.LeftFooter = "&G";
            xSt.PageSetup.CenterFooter = @"&""微软雅黑""&8" + "共&N页，第&P页\n" + companyModel["name"].ToString().Trim() + "　　地址：" + companyModel["address"].ToString().Trim() + "　　电话：" + companyModel["phone"].ToString().Trim();

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 8;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 5;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 36;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 12;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 7]].RowHeight = 1;
            xSt.Range[excel.Cells[2, 1], excel.Cells[13, 7]].RowHeight = 18;

            excel.Cells[2, 1] = customerModel["customerName"].ToString().Trim();
            excel.Cells[2, 4] = "报价方：";
            excel.Cells[2, 5] = companyModel["name"].ToString().Trim();

            if (!string.IsNullOrEmpty(quotationModel["purchaserName"].ToString().Trim()))
            {
                excel.Cells[3, 1] = quotationModel["purchaserName"].ToString().Trim();
            }
            else
            {
                excel.Cells[3, 1] = quotationModel["userName"].ToString().Trim();
            }
            excel.Cells[3, 4] = "地址：";
            excel.Cells[3, 5] = companyModel["address"].ToString().Trim();
            excel.Cells[4, 1] = (customerModel["customerAddress"]["provinceName"].ToString().Trim() == "请选择省" ? "" : customerModel["customerAddress"]["provinceName"].ToString().Trim()) + (customerModel["customerAddress"]["cityName"].ToString().Trim() == "请选择市" || customerModel["customerAddress"]["cityName"].ToString().Trim() == "市辖区" ? "" : customerModel["customerAddress"]["cityName"].ToString().Trim()) + (customerModel["customerAddress"]["districtName"].ToString().Trim() == "请选择县/区" ? "" : customerModel["customerAddress"]["districtName"].ToString().Trim()) + customerModel["customerAddress"]["address"].ToString().Trim();
            excel.Cells[4, 4] = "联系人：";
            excel.Cells[4, 5] = quotationModel["bidderName"].ToString().Trim();
            excel.Cells[5, 4] = "电话：";
            excel.Cells[5, 5] = companyModel["phone"].ToString().Trim() + (string.IsNullOrEmpty(userModel["extensionNum"].ToString()) ? "" : "-" + userModel["extensionNum"].ToString().Trim());
            excel.Cells[6, 4] = "手机：";
            xSt.Range[excel.Cells[6, 5], excel.Cells[6, 5]].NumberFormat = "@";
            excel.Cells[6, 5] = userModel["mobilePhone"].ToString().Trim();
            excel.Cells[7, 1] = "报价单号：";
            excel.Cells[7, 3] = quotationModel["quotationNo"].ToString().Trim();
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 3]].Font.Bold = true;
            //excel.Cells[8, 4] = "客户询价单号：";
            //excel.Cells[8, 5] = Dr["CustomerQueryNo"].ToString().Trim();
            excel.Cells[7, 4] = "运费负担：";
            excel.Cells[7, 5] = quotationModel["quotationYFFD"].ToString().Trim();
            excel.Cells[8, 4] = "收款条件";
            excel.Cells[8, 5] = quotationModel["quotationJSFS"].ToString().Trim();
            excel.Cells[9, 4] = "报价有效期：";
            excel.Cells[9, 5] = quotationModel["quotationYXQ"].ToString().Trim();
            excel.Cells[10, 4] = "日期：";
            xSt.Range[excel.Cells[10, 5], excel.Cells[10, 6]].Merge(false);
            xSt.Range[excel.Cells[10, 5], excel.Cells[10, 6]].Value2 = Convert.ToDateTime(quotationModel["quotationTime"]).ToString("yyyy年MM月dd日");
            xSt.Range[excel.Cells[10, 5], excel.Cells[10, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excel.Cells[11, 1] = "您好：" + (string.IsNullOrEmpty(quotationModel["purchaserName"].ToString().Trim()) ? quotationModel["userName"].ToString().Trim() : quotationModel["purchaserName"].ToString().Trim());
            excel.Cells[12, 1] = "感谢您的询价，我们的报价如下：";

            xSt.Range[excel.Cells[13, 1], excel.Cells[13, 7]].RowHeight = 6;

            //xSt.PageSetup.PrintTitleRows = "$15:$15";  标题行 
            xSt.Range[excel.Cells[14, 1], excel.Cells[14, 7]].RowHeight = 20;
            excel.Cells[14, 1] = "行号";
            xSt.Range[excel.Cells[14, 2], excel.Cells[14, 3]].Merge(false);
            xSt.Range[excel.Cells[14, 2], excel.Cells[14, 3]].Value2 = "产品描述";
            excel.Cells[14, 4] = "数量";
            excel.Cells[14, 5] = "单位";
            excel.Cells[14, 6] = "单价";
            excel.Cells[14, 7] = "总价";
            xSt.Range[excel.Cells[14, 1], excel.Cells[14, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[14, 1], excel.Cells[14, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            int j = 14;
            for (int i = 0; i < Convert.ToInt32(quotationModel["number"]); i++)
            {
                j = j + 6;

                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 1]].RowHeight = 20;
                xSt.Range[excel.Cells[j - 4, 1], excel.Cells[j, 1]].RowHeight = 14;

                //if (!string.IsNullOrEmpty(Ds.Tables[1].Rows[i]["Parameter"].ToString().Trim()))
                //    xSt.Range[excel.Cells[j - 4, 1], excel.Cells[j - 1, 1]).RowHeight = 28;

                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                excel.Cells[j - 5, 1] = i + 1;
                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].NumberFormat = "@";
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_QuotationProduct"]["isInquiry"].ToString()) ? productModel[i]["Sales_QuotationProduct"]["inquiryProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim());
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Font.Bold = true;
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j - 5, 4], excel.Cells[j - 5, 4]].NumberFormat = "@";
                excel.Cells[j - 5, 4] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["quantity"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["quantity"]).ToString("N2"));
                xSt.Range[excel.Cells[j - 5, 4], excel.Cells[j - 5, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j - 5, 5] = productModel[i]["Product_Product"]["unit"].ToString().Trim();
                xSt.Range[excel.Cells[j - 5, 5], excel.Cells[j - 5, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j - 5, 6] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["unitPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["unitPrice"].ToString()).ToString("N2") + " ￥");
                xSt.Range[excel.Cells[j - 5, 6], excel.Cells[j - 5, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                excel.Cells[j - 5, 7] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["totalPrice"].ToString()).ToString("N2") + " ￥");
                xSt.Range[excel.Cells[j - 5, 7], excel.Cells[j - 5, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_QuotationProduct"]["isInquiry"].ToString()) ? productModel[i]["Sales_QuotationProduct"]["inquiryProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim()));
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].WrapText = true;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Merge(false);
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Value2 = productModel[i]["Sales_QuotationProduct"]["delivery"].ToString().Trim();
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            }

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 1]].RowHeight = 24;
            excel.Cells[j + 1, 6] = "总金额：";
            xSt.Range[excel.Cells[j + 1, 6], excel.Cells[j + 1, 6]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 6], excel.Cells[j + 1, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j + 1, 7] = (string.IsNullOrEmpty(quotationModel["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(quotationModel["totalPrice"].ToString()).ToString("N2") + " ￥");
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 1]].RowHeight = 24;
            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 1]].Font.Bold = true;
            if (Convert.ToBoolean(quotationModel["hasTaxNoInvoice"].ToString()))
            {
                excel.Cells[j + 2, 1] = "注：以上价格以人民币计算，不含税。";
            }
            else
            {
                excel.Cells[j + 2, 1] = "注：以上价格以人民币计算，含" + quotationModel["hasTax"]; 
            }

            xSt.Range[excel.Cells[j + 3, 1], excel.Cells[j + 3, 1]].RowHeight = 28;

            xSt.Range[excel.Cells[j + 4, 1], excel.Cells[j + 12, 1]].RowHeight = 16;
            excel.Cells[j + 4, 1] = "我们期待您的订单！";
            excel.Cells[j + 5, 1] = "祝一切好！";
            excel.Cells[j + 7, 1] = quotationModel["bidderName"].ToString().Trim();
            excel.Cells[j + 8, 1] = companyModel["name"].ToString().Trim();


            Range range = xSt.Range[excel.Cells[j + 5, 2], excel.Cells[j + 5, 2]];
            xSt.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "/image/" + companyModel["nid"].ToString().Trim() + "sno.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Convert.ToSingle(range.Left) - 10, Convert.ToSingle(range.Top) - 35, 106, 106);


            excel.Cells[j + 10, 1] = "********************************************************";
            excel.Cells[j + 11, 1] = "为提升我们的服务质量，作为重点客户，您可以直接联系我们的销售总监关于合作、建议、意见和投诉等事宜。";
            excel.Cells[j + 12, 1] = "邮箱：csr01@mro9.com";

            excel.Visible = true;

            string path = "";
            string filename = "";

            if (strType.ToLower() == "excel")
            {
                filename = quotationModel["_id"] + ".xlsx";
                path = Server.MapPath("~/") + "temp\\" + filename;

                //保存excel
                xBk.SaveCopyAs(path);
            }
            else
            {
                filename = quotationModel["_id"] + ".pdf";
                path = Server.MapPath("~/") + "temp\\" + filename;

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

        public void OutputExcel1(dynamic quotationModel, dynamic companyModel, dynamic customerModel, dynamic productModel, dynamic userModel, string strType)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            xSt.PageSetup.LeftMargin = 0.8 / 0.035;
            xSt.PageSetup.RightMargin = 0.5 / 0.035;
            xSt.PageSetup.HeaderMargin = 2 / 0.035;
            xSt.PageSetup.FooterMargin = 1 / 0.035;
            xSt.PageSetup.TopMargin = 3 / 0.035;
            xSt.PageSetup.BottomMargin = 1.8 / 0.035;
            xSt.PageSetup.LeftHeaderPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"] + "2logosno.jpg";
            xSt.PageSetup.LeftHeader = "&G";
            xSt.PageSetup.CenterHeader = @"&""微软雅黑,Bold""&16" + "报价单　QUOTATION";
            xSt.PageSetup.RightHeader = @"&""微软雅黑""&8" + "QUO NO.：" + quotationModel["quotationNo"].ToString().Trim() + "\nDATE：" + Convert.ToDateTime(quotationModel["quotationTime"]).ToString("yyyy年MM月dd日") + "\n　";
            xSt.PageSetup.LeftFooterPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + "footline2.jpg";
            xSt.PageSetup.LeftFooter = "&G";
            xSt.PageSetup.CenterFooter = @"&""微软雅黑""&8" + "共&N页，第&P页\n" + companyModel["name"].ToString().Trim() + "　　地址：" + companyModel["address"].ToString().Trim() + "　　电话：" + companyModel["phone"].ToString().Trim();

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 8;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 16;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 40;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 12.5;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 11;
            xSt.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 11;
            xSt.Range[excel.Cells[1, 9], excel.Cells[1, 9]].ColumnWidth = 11;
            xSt.Range[excel.Cells[1, 10], excel.Cells[1, 10]].ColumnWidth = 11;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 7]].RowHeight = 1;
            xSt.Range[excel.Cells[2, 1], excel.Cells[4, 7]].RowHeight = 14;

            excel.Cells[2, 1] = "TO：";
            xSt.Range[excel.Cells[2, 2], excel.Cells[2, 3]].Merge(false);
            xSt.Range[excel.Cells[2, 2], excel.Cells[2, 3]].Value2 = customerModel["customerName"].ToString().Trim();
            xSt.Range[excel.Cells[2, 2], excel.Cells[2, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excel.Cells[2, 4] = "FROM：";
            xSt.Range[excel.Cells[2, 5], excel.Cells[2, 10]].Merge(false);
            xSt.Range[excel.Cells[2, 5], excel.Cells[2, 10]].Value2 = companyModel["name"].ToString().Trim();
            xSt.Range[excel.Cells[2, 5], excel.Cells[2, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excel.Cells[3, 1] = "ATT：";
            xSt.Range[excel.Cells[3, 2], excel.Cells[3, 3]].Merge(false);
            if (!string.IsNullOrEmpty(quotationModel["purchaserName"].ToString()))
            {
                xSt.Range[excel.Cells[3, 2], excel.Cells[3, 3]].Value2 = quotationModel["purchaserName"].ToString().Trim();
            }
            else
            {
                xSt.Range[excel.Cells[3, 2], excel.Cells[3, 3]].Value2 = quotationModel["userName"].ToString().Trim();
            }
            xSt.Range[excel.Cells[3, 2], excel.Cells[3, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excel.Cells[3, 4] = "ATT：";
            xSt.Range[excel.Cells[3, 5], excel.Cells[3, 10]].Merge(false);
            xSt.Range[excel.Cells[3, 5], excel.Cells[3, 10]].Value2 = quotationModel["bidderName"].ToString().Trim() + "　" + userModel["mobilePhone"].ToString().Trim();
            xSt.Range[excel.Cells[3, 5], excel.Cells[3, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excel.Cells[4, 4] = "TEL：";
            xSt.Range[excel.Cells[4, 5], excel.Cells[4, 10]].Merge(false);
            xSt.Range[excel.Cells[4, 5], excel.Cells[4, 10]].Value2 = companyModel["phone"].ToString().Trim() + (string.IsNullOrEmpty(userModel["extensionNum"].ToString()) ? "" : "-" + userModel["extensionNum"].ToString().Trim());
            xSt.Range[excel.Cells[4, 5], excel.Cells[4, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 10]].RowHeight = 9;

            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 10]].RowHeight = 29;
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 10]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            excel.Cells[6, 1] = "NO.";
            excel.Cells[6, 2] = "编号\nSTOCK NO.";
            excel.Cells[6, 3] = "产品描述\nDESCRIPTION";
            excel.Cells[6, 4] = "货期\nDELIVERY";
            excel.Cells[6, 5] = "数量\nQUAN";
            excel.Cells[6, 6] = "单位\nUNIT";
            excel.Cells[6, 7] = "未税单价\n(N-VAT)";
            excel.Cells[6, 8] = "含税单价\n(VAT)";
            excel.Cells[6, 9] = "未税金额\nTOTAL (N-VAT)";
            excel.Cells[6, 10] = "含税金额\nTOTAL (VAT)";

            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 10]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            Color c = Color.FromArgb(242, 242, 242);
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 10]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c);
            //xSt.Range[excel.Cells[6, 1], excel.Cells[6, 10]).Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            int j = 6;
            for (int i = 0; i < Convert.ToInt32(quotationModel["number"]); i++)
            {
                j = j + 1;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 10]].RowHeight = 32;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 1]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 2]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 3], excel.Cells[j, 3]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 4], excel.Cells[j, 4]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 5], excel.Cells[j, 5]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 6], excel.Cells[j, 6]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 7], excel.Cells[j, 7]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 8], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 9], excel.Cells[j, 9]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 10], excel.Cells[j, 10]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 10], excel.Cells[j, 10]].Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 10]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                excel.Cells[j, 1] = i + 1;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 2]].NumberFormat = "@";
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 2]].WrapText = true;
                excel.Cells[j, 2] = (Convert.ToBoolean(productModel[i]["Sales_QuotationProduct"]["isInquiry"].ToString()) ? productModel[i]["Sales_QuotationProduct"]["inquiryProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
                xSt.Range[excel.Cells[j, 3], excel.Cells[j, 3]].WrapText = true;
                excel.Cells[j, 3] = (Convert.ToBoolean(productModel[i]["Sales_QuotationProduct"]["isInquiry"].ToString()) ? productModel[i]["Sales_QuotationProduct"]["inquiryProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";
                xSt.Range[excel.Cells[j, 4], excel.Cells[j, 4]].WrapText = true;
                excel.Cells[j, 4] = productModel[i]["Sales_QuotationProduct"]["delivery"].ToString().Trim() + " ";
                xSt.Range[excel.Cells[j, 5], excel.Cells[j, 5]].NumberFormat = "@";
                excel.Cells[j, 5] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["quantity"]).ToString("N2"));
                excel.Cells[j, 6] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
                xSt.Range[excel.Cells[j, 5], excel.Cells[j, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 7], excel.Cells[j, 7]].NumberFormat = "@";
                excel.Cells[j, 7] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["unitPriceNoTax"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["unitPriceNoTax"].ToString()).ToString("N2")) + " ";
                xSt.Range[excel.Cells[j, 8], excel.Cells[j, 8]].NumberFormat = "@";
                excel.Cells[j, 8] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["unitPrice"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["unitPrice"].ToString()).ToString("N2")) + " ";
                xSt.Range[excel.Cells[j, 9], excel.Cells[j, 9]].NumberFormat = "@";
                excel.Cells[j, 9] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["totalPriceNoTax"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["totalPriceNoTax"].ToString()).ToString("N2")) + " ";
                xSt.Range[excel.Cells[j, 10], excel.Cells[j, 10]].NumberFormat = "@";
                excel.Cells[j, 10] = (string.IsNullOrEmpty(productModel[i]["Sales_QuotationProduct"]["totalPrice"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_QuotationProduct"]["totalPrice"].ToString()).ToString("N2")) + " ";
                xSt.Range[excel.Cells[j, 7], excel.Cells[j, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            }

            if (Convert.ToInt32(quotationModel["number"]) < 8)
            {
                for (int i = 0; i < 8 - Convert.ToInt32(quotationModel["number"]); i++)
                {
                    j = j + 1;

                    xSt.Range[excel.Cells[j, 1], excel.Cells[j, 10]].RowHeight = 32;
                    xSt.Range[excel.Cells[j, 1], excel.Cells[j, 1]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 2], excel.Cells[j, 2]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 3], excel.Cells[j, 3]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 4], excel.Cells[j, 4]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 5], excel.Cells[j, 5]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 6], excel.Cells[j, 6]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 7], excel.Cells[j, 7]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 8], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 9], excel.Cells[j, 9]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 10], excel.Cells[j, 10]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 10], excel.Cells[j, 10]].Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xSt.Range[excel.Cells[j, 1], excel.Cells[j, 10]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                    excel.Cells[j, 1] = i + 1 + Convert.ToInt32(quotationModel["number"]);
                    xSt.Range[excel.Cells[j, 1], excel.Cells[j, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[j, 2] = " ";
                    excel.Cells[j, 3] = " ";
                    excel.Cells[j, 4] = " ";
                    excel.Cells[j, 5] = " ";
                    excel.Cells[j, 6] = " ";
                    xSt.Range[excel.Cells[j, 5], excel.Cells[j, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[j, 7] = " ";
                    excel.Cells[j, 8] = " ";
                    excel.Cells[j, 9] = " ";
                    excel.Cells[j, 10] = " ";
                    xSt.Range[excel.Cells[j, 7], excel.Cells[j, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                }
            }

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 10]].RowHeight = 22.5;

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 3]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 3]].Value2 = "■ 附加说明 INSTRUCTION：" + quotationModel["quotationFJSM"].ToString().Trim();
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 3]].WrapText = true;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 2, 3]].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            //xSt.Range[excel.Cells[j + 1, 4], excel.Cells[j + 1, 5]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 4], excel.Cells[j + 1, 4]].Value2 = "税率 VAT RATE";
            xSt.Range[excel.Cells[j + 1, 4], excel.Cells[j + 1, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 1, 4], excel.Cells[j + 1, 4]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 1, 4], excel.Cells[j + 1, 4]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 4], excel.Cells[j + 1, 4]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].Value2 = quotationModel["hasTax"];
            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 5], excel.Cells[j + 1, 6]].Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]].Value2 = "未税总额 AMOUNT(N-VAT)";
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 1, 10], excel.Cells[j + 1, 10]].NumberFormat = "@";
            excel.Cells[j + 1, 10] = (string.IsNullOrEmpty(quotationModel["totalPriceNoTax"].ToString()) ? "0.00" : Convert.ToDecimal(quotationModel["totalPriceNoTax"].ToString()).ToString("N2")) + "";
            xSt.Range[excel.Cells[j + 1, 10], excel.Cells[j + 1, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Range[excel.Cells[j + 1, 10], excel.Cells[j + 1, 10]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 10], excel.Cells[j + 1, 10]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 10], excel.Cells[j + 1, 10]].Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 2, 4], excel.Cells[j + 2, 5]].Merge(false);
            xSt.Range[excel.Cells[j + 2, 4], excel.Cells[j + 2, 5]].Value2 = "货币 CURRENCY";
            xSt.Range[excel.Cells[j + 2, 4], excel.Cells[j + 2, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 2, 4], excel.Cells[j + 2, 5]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 2, 4], excel.Cells[j + 2, 5]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 2, 4], excel.Cells[j + 2, 5]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            excel.Cells[j + 2, 6] = "RMB";
            xSt.Range[excel.Cells[j + 2, 6], excel.Cells[j + 2, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[j + 2, 6], excel.Cells[j + 2, 6]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 2, 6], excel.Cells[j + 2, 6]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 2, 6], excel.Cells[j + 2, 6]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 2, 6], excel.Cells[j + 2, 6]].Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 9]].Merge(false);
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 9]].Value2 = "含税总额 AMOUNT(VAT)";
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 9]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 9]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 2, 7], excel.Cells[j + 2, 9]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 2, 10], excel.Cells[j + 2, 10]].NumberFormat = "@";
            excel.Cells[j + 2, 10] = (string.IsNullOrEmpty(quotationModel["totalPrice"].ToString()) ? "0.00" : Convert.ToDecimal(quotationModel["totalPrice"].ToString()).ToString("N2")) + "";
            xSt.Range[excel.Cells[j + 2, 10], excel.Cells[j + 2, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Range[excel.Cells[j + 2, 10], excel.Cells[j + 2, 10]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 2, 10], excel.Cells[j + 2, 10]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 2, 10], excel.Cells[j + 2, 10]].Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            xSt.Range[excel.Cells[j + 3, 1], excel.Cells[j + 3, 10]].RowHeight = 7;

            excel.Cells[j + 4, 1] = "我们期待您的订单！";
            xSt.Range[excel.Cells[j + 4, 7], excel.Cells[j + 4, 10]].Merge(false);
            xSt.Range[excel.Cells[j + 4, 7], excel.Cells[j + 4, 10]].Value2 = "＊报价单有效期" + quotationModel["quotationYXQ"].ToString().Trim();
            xSt.Range[excel.Cells[j + 4, 7], excel.Cells[j + 4, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Range[excel.Cells[j + 4, 7], excel.Cells[j + 4, 10]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            excel.Cells[j + 5, 1] = "祝一切好！";

            xSt.Range[excel.Cells[j + 4, 1], excel.Cells[j + 5, 1]].Font.Size = 10;
            xSt.Range[excel.Cells[j + 4, 1], excel.Cells[j + 5, 1]].Font.Italic = true;

            Range range = xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 9]];
            xSt.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "/image/" + companyModel["nid"].ToString().Trim() + "sno.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Convert.ToSingle(range.Left) + 100, Convert.ToSingle(range.Top) + 5, 106, 106);

            excel.Visible = true;

            string path = "";
            string filename = "";

            if (strType.ToLower() == "excel")
            {
                filename = quotationModel["_id"] + ".xlsx";
                path = Server.MapPath("~/") + "temp\\" + filename;

                //保存excel
                xBk.SaveCopyAs(path);
            }
            else
            {
                filename = quotationModel["_id"] + ".pdf";
                path = Server.MapPath("~/") + "temp\\" + filename;

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