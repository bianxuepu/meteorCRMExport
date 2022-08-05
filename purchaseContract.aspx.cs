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
    public partial class purchaseContract : System.Web.UI.Page
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
            var supplierModel = m.supplierModel;
            var productModel = m.productModel;
            var userModel = m.userModel;

            OutputExcel(orderModel, companyModel, supplierModel, productModel, userModel, filetype);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic companyModel, dynamic supplierModel, dynamic productModel, dynamic userModel, string strType)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.PageSetup.LeftMargin = 0.9 / 0.035;
            xSt.PageSetup.RightMargin = 0.9 / 0.035;
            xSt.PageSetup.HeaderMargin = 2 / 0.035;
            xSt.PageSetup.FooterMargin = 1 / 0.035;
            xSt.PageSetup.TopMargin = 3.3 / 0.035;
            xSt.PageSetup.BottomMargin = 1.8 / 0.035;
            xSt.PageSetup.LeftHeaderPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"].ToString().Trim() + "logo.jpg";
            xSt.PageSetup.LeftHeader = "&G";
            xSt.PageSetup.CenterHeader = @"&""微软雅黑,Bold""&14" + "采购合同";
            xSt.PageSetup.RightHeader = @"&""微软雅黑""&9" + "合同单号：" + orderModel["orderNo"].ToString().Trim() + "　";
            xSt.PageSetup.LeftFooterPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + "footline.jpg";
            xSt.PageSetup.LeftFooter = @"&""微软雅黑""&8" + companyModel["name"].ToString().Trim();
            xSt.PageSetup.RightFooter = @"&""微软雅黑""&8" + "共&N页，第&P页　";

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
            xSt.Range[excel.Cells[2, 1], excel.Cells[4, 7]].RowHeight = 15;

            excel.Cells[2, 1] = "买方：" + companyModel["name"].ToString().Trim();
            excel.Cells[3, 1] = "卖方：" + (orderModel["supplierAnotherName"].ToString().Trim()!="" ? orderModel["supplierAnotherName"].ToString().Trim() : supplierModel["supplierName"].ToString().Trim());
            excel.Cells[4, 1] = "签订日期：" + Convert.ToDateTime(orderModel["recordDate"].ToString()).ToString("yyyy年MM月dd日");

            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 7]].RowHeight = 12;
            xSt.Range[excel.Cells[6, 1], excel.Cells[7, 7]].RowHeight = 16;

            excel.Cells[6, 1] = "根据《中华人民共和国合同法》及相关法律、法规，买、卖双方经友好协商一致同意按下列条款签订本合同：";
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 7]].Font.Bold = true;
            excel.Cells[7, 1] = "一、合同标的：";

            //xSt.PageSetup.PrintTitleRows = "$8:$8";  标题行 
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 7]].RowHeight = 20;
            excel.Cells[8, 1] = "行号";
            xSt.Range[excel.Cells[8, 2], excel.Cells[8, 3]].Merge(false);
            xSt.Range[excel.Cells[8, 2], excel.Cells[8, 3]].Value2 = "产品描述";
            excel.Cells[8, 4] = "数量";
            excel.Cells[8, 5] = "单位";
            excel.Cells[8, 6] = "单价";
            excel.Cells[8, 7] = "总价";
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            int j = 8;
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                j = j + 6;

                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 1]].RowHeight = 16;
                xSt.Range[excel.Cells[j - 4, 1], excel.Cells[j - 2, 1]].RowHeight = 13;
                xSt.Range[excel.Cells[j - 1, 1], excel.Cells[j, 1]].RowHeight = 14;

                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                excel.Cells[j - 5, 1] = i + 1;
                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].NumberFormat = "@";
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Value2 = productModel[i]["Product_Product"]["proNo"].ToString().Trim() + " ";
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Font.Bold = true;
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j - 5, 4], excel.Cells[j - 5, 4]].NumberFormat = "@";
                excel.Cells[j - 5, 4] = (string.IsNullOrEmpty(productModel[i]["Purchase_OrderProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Purchase_OrderProduct"]["quantity"]).ToString("N2"));
                xSt.Range[excel.Cells[j - 5, 4], excel.Cells[j - 5, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j - 5, 5] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
                xSt.Range[excel.Cells[j - 5, 5], excel.Cells[j - 5, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j - 5, 6] = (string.IsNullOrEmpty(productModel[i]["Purchase_OrderProduct"]["unitPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Purchase_OrderProduct"]["unitPrice"].ToString()).ToString("N2") + " ￥");
                xSt.Range[excel.Cells[j - 5, 6], excel.Cells[j - 5, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                excel.Cells[j - 5, 7] = (string.IsNullOrEmpty(productModel[i]["Purchase_OrderProduct"]["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Purchase_OrderProduct"]["totalPrice"].ToString()).ToString("N2") + " ￥");
                xSt.Range[excel.Cells[j - 5, 7], excel.Cells[j - 5, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 2, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 2, 3]].Value2 = productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim()) + " ";
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 2, 3]].WrapText = true;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 2, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 2, 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                xSt.Range[excel.Cells[j - 1, 2], excel.Cells[j - 1, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 1, 2], excel.Cells[j - 1, 3]].Value2 = "税收编码：" + productModel[i]["Product_Product"]["taxEncodingNo"].ToString().Trim();
                xSt.Range[excel.Cells[j - 1, 2], excel.Cells[j - 1, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Merge(false);
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Value2 = "交货时间：" + Convert.ToDateTime(productModel[i]["Purchase_OrderProduct"]["delivery"].ToString()).ToString("yyyy-MM-dd");
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            }

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 1]].RowHeight = 22;
            excel.Cells[j + 1, 2] = "大写总金额：";
            xSt.Range[excel.Cells[j + 1, 3], excel.Cells[j + 1, 4]].Merge(false);
            xSt.Range[excel.Cells[j + 1, 3], excel.Cells[j + 1, 4]].Value2 = MoneyToUpper(string.IsNullOrEmpty(orderModel["totalPrice"].ToString()) ? 0 : Convert.ToDecimal(orderModel["totalPrice"].ToString()).ToString("N2"));
            excel.Cells[j + 1, 6] = "总金额：";
            xSt.Range[excel.Cells[j + 1, 6], excel.Cells[j + 1, 6]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 6], excel.Cells[j + 1, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j + 1, 7] = (string.IsNullOrEmpty(orderModel["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(orderModel["totalPrice"].ToString()).ToString("N2") + " ￥");
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 1]].RowHeight = 22;
            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 1]].Font.Bold = true;
            if (Convert.ToBoolean(orderModel["hasTaxNoInvoice"].ToString()))
            {
                excel.Cells[j + 2, 1] = "注：以上价格以人民币计算，不含税。"; 
            }
            else
            {
                excel.Cells[j + 2, 1] = "注：以上价格以人民币计算，含" + orderModel["hasTax"];
            }

            xSt.Range[excel.Cells[j + 3, 1], excel.Cells[j + 3, 1]].RowHeight = 16;
            excel.Cells[j + 3, 1] = "二、附加说明：";
            xSt.Range[excel.Cells[j + 3, 1], excel.Cells[j + 3, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].Value2 = string.IsNullOrEmpty(orderModel["purchaseOrderFJSM"].ToString().Trim()) ? "无" : orderModel["purchaseOrderFJSM"].ToString().Trim();
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 4, 1], excel.Cells[j + 4, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 5, 1], excel.Cells[j + 5, 1]].RowHeight = 27;
            excel.Cells[j + 5, 1] = "三、交货日期：";
            xSt.Range[excel.Cells[j + 5, 1], excel.Cells[j + 5, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].Value2 = "本合同约定交货时间为" + Convert.ToDateTime(orderModel["deliveryDate"].ToString()).ToString("yyyy年MM月dd日") + "。分项交货时间在标的中注明。\n具体交货时间以买方书面通知为准。买方没有发出交货通知的，卖方按照上述约定的时间交货。";
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 6, 1], excel.Cells[j + 6, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 7, 1], excel.Cells[j + 7, 1]].RowHeight = 16;
            excel.Cells[j + 7, 1] = "四、运费负担：";
            xSt.Range[excel.Cells[j + 7, 1], excel.Cells[j + 7, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].Value2 = orderModel["purchaseOrderYFFD"].ToString().Trim();
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 8, 1], excel.Cells[j + 8, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 9, 1], excel.Cells[j + 9, 1]].RowHeight = 16;
            excel.Cells[j + 9, 1] = "五、付款方式：";
            xSt.Range[excel.Cells[j + 9, 1], excel.Cells[j + 9, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].Value2 = orderModel["purchaseOrderFKFS"].ToString().Trim();
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 10, 1], excel.Cells[j + 10, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 11, 1], excel.Cells[j + 11, 1]].RowHeight = 41;
            excel.Cells[j + 11, 1] = "六、延迟交付后果：";
            xSt.Range[excel.Cells[j + 11, 1], excel.Cells[j + 11, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].Value2 = "卖方未能按照合同约定交货的，每迟延交货一日应向买方支付合同价款2‰的违约金。迟延超过15日，买方有权单方面解除合同。买方因此解除合同的，有权要求卖方退还全部买方已支付款项，并支付合同价款20%的违约金。如因此给买方造成损失违约金不足以弥补的，卖方还应继续承担赔偿责任。";
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 12, 1], excel.Cells[j + 12, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 13, 1], excel.Cells[j + 13, 1]].RowHeight = 82;
            excel.Cells[j + 13, 1] = "七、缺陷产品处理：";
            xSt.Range[excel.Cells[j + 13, 1], excel.Cells[j + 13, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].Value2 = "卖方应保证其交付的标的物为全新的符合合同约定规格及质量标准的货物。标的物运至交货地点后3日内，买方应对产品外包装、数量、外观质量进行验收。如存在数量、规格、质量缺陷的，卖方应在得到买方通知后及时进行更换、补足，并承担相应费用。如产品存在隐性瑕疵，买方应在发现后的合理期限内通知卖方对该部分产品予以更换，卖方应在收到通知后及时更换并承担相应费用。如卖方更换产品累计发生三次或占全部标的物 10%以上的，买方有权解除合同，退还所有标的物，并要求卖方返还买方已支付全部款项，同时买方还有权要求卖方支付合同价款20%的违约金。如因此给买方造成损失违约金不足以弥补的，卖方还应继续承担赔偿责任。";
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 14, 1], excel.Cells[j + 14, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 15, 1], excel.Cells[j + 15, 1]].RowHeight = 16;
            excel.Cells[j + 15, 1] = "八、争议解决：";
            xSt.Range[excel.Cells[j + 15, 1], excel.Cells[j + 15, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].Value2 = "合同履行过程中，买卖双方如有争议，应通过友好协商解决。若协商不成，均可向" + orderModel["purchaseOrderZYJJ"].ToString().Trim() + "有管辖权的人民法院提出诉讼。";
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 16, 1], excel.Cells[j + 16, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 17, 1], excel.Cells[j + 17, 1]].RowHeight = 16;
            excel.Cells[j + 17, 1] = "九、合同签订：";
            xSt.Range[excel.Cells[j + 17, 1], excel.Cells[j + 17, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].Value2 = "本合同一式两份，具有同等的法律效力，自双方签字盖章之日起生效。";
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 18, 1], excel.Cells[j + 18, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 19, 1], excel.Cells[j + 19, 1]].RowHeight = 25;
            excel.Cells[j + 19, 1] = "十、收货信息：";
            xSt.Range[excel.Cells[j + 19, 1], excel.Cells[j + 19, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].Value2 = "卖方将把本合同中的货物发送到：" + orderModel["purchaseOrderSHXX"].ToString().Trim();
            xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 20, 1], excel.Cells[j + 20, 1]].RowHeight = 16;

            xSt.Range[excel.Cells[j + 21, 1], excel.Cells[j + 28, 1]].RowHeight = 18;

            excel.Cells[j + 21, 1] = "买方：" + companyModel["name"].ToString().Trim();
            excel.Cells[j + 21, 4] = "卖方：" + (orderModel["supplierAnotherName"].ToString().Trim() != "" ? orderModel["supplierAnotherName"].ToString().Trim() : supplierModel["supplierName"].ToString().Trim());
            excel.Cells[j + 22, 1] = "买方代表：" + orderModel["personInChargeName"].ToString().Trim() + "　" + userModel["mobilePhone"].ToString().Trim();
            excel.Cells[j + 22, 4] = "卖方代表：" + orderModel["contactName"].ToString().Trim() + "　" + supplierModel["mobilePhone"].ToString().Trim();
            xSt.Range[excel.Cells[j + 22, 5], excel.Cells[j + 22, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j + 23, 1] = "买方公章：";
            excel.Cells[j + 23, 4] = "卖方公章：";
            excel.Cells[j + 24, 1] = "社会信用代码：" + companyModel["license"]["code"];
            excel.Cells[j + 24, 4] = "社会信用代码：" + (orderModel["supplierAnotherName"].ToString().Trim() != "" ? "" : supplierModel["tin"].ToString().Trim());
            excel.Cells[j + 25, 1] = "开户行：" + companyModel["license"]["bank"];
            excel.Cells[j + 25, 4] = "开户行：";
            try
            {
                excel.Cells[j + 25, 4] = "开户行：" + supplierModel["bankName"];
            }
            catch { }
            excel.Cells[j + 26, 1] = "帐号：" + companyModel["license"]["account"];
            excel.Cells[j + 26, 4] = "帐号：";
            try
            {
                excel.Cells[j + 26, 4] = "帐号：" + supplierModel["bankAccount"];
            }
            catch { }
            xSt.Range[excel.Cells[j + 27, 1], excel.Cells[j + 27, 1]].RowHeight = 20;
            excel.Cells[j + 27, 1] = "地址：" + companyModel["address"].ToString().Trim();
            xSt.Range[excel.Cells[j + 27, 4], excel.Cells[j + 27, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 27, 4], excel.Cells[j + 27, 7]].Value2 = "地址：" + (supplierModel["supplierAddress"]["provinceName"].ToString().Trim() == "请选择省" ? "" : supplierModel["supplierAddress"]["provinceName"].ToString().Trim()) + (supplierModel["supplierAddress"]["cityName"].ToString().Trim() == "请选择市" || supplierModel["supplierAddress"]["cityName"].ToString().Trim() == "市辖区" ? "" : supplierModel["supplierAddress"]["cityName"].ToString().Trim()) + (supplierModel["supplierAddress"]["districtName"].ToString().Trim() == "请选择县/区" ? "" : supplierModel["supplierAddress"]["districtName"].ToString().Trim()) + supplierModel["supplierAddress"]["address"].ToString().Trim();
            xSt.Range[excel.Cells[j + 27, 4], excel.Cells[j + 27, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 27, 4], excel.Cells[j + 27, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excel.Cells[j + 28, 1] = "电话：" + companyModel["phone"];
            excel.Cells[j + 28, 4] = "电话：" + (supplierModel["supplierPhone"].ToString() != "[]" ? supplierModel["supplierPhone"][0]["phone"].ToString().Trim() : "");

            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 1]].RowHeight = 18;
            xSt.Range[excel.Cells[j + 30, 1], excel.Cells[j + 30, 1]].RowHeight = 14;
            xSt.Range[excel.Cells[j + 31, 1], excel.Cells[j + 31, 1]].RowHeight = 9;
            xSt.Range[excel.Cells[j + 31, 1], excel.Cells[j + 31, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 31, 1], excel.Cells[j + 31, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 46, 1]].RowHeight = 24;

            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 32, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 32, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 32, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 32, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 32, 7]].Font.Size = 12;
            xSt.Range[excel.Cells[j + 32, 1], excel.Cells[j + 32, 7]].Value2 = "发票及发票寄送";

            xSt.Range[excel.Cells[j + 33, 1], excel.Cells[j + 46, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 33, 1], excel.Cells[j + 46, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xSt.Range[excel.Cells[j + 33, 1], excel.Cells[j + 46, 7]].Font.Size = 10;

            xSt.Range[excel.Cells[j + 33, 1], excel.Cells[j + 33, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 33, 1], excel.Cells[j + 33, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 33, 1], excel.Cells[j + 33, 7]].Value2 = "发票信息：";

            xSt.Range[excel.Cells[j + 34, 1], excel.Cells[j + 34, 7]].Merge(false); 
            xSt.Range[excel.Cells[j + 34, 1], excel.Cells[j + 34, 7]].Value2 = "公司名称：" + companyModel["name"];

            xSt.Range[excel.Cells[j + 35, 1], excel.Cells[j + 35, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 35, 1], excel.Cells[j + 35, 7]].Value2 = "税号：" + companyModel["license"]["code"];

            xSt.Range[excel.Cells[j + 36, 1], excel.Cells[j + 36, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 36, 1], excel.Cells[j + 36, 7]].Value2 = "地址：" + companyModel["license"]["address"];

            xSt.Range[excel.Cells[j + 37, 1], excel.Cells[j + 37, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 37, 1], excel.Cells[j + 37, 7]].Value2 = "电话：" + companyModel["license"]["phone"];

            xSt.Range[excel.Cells[j + 38, 1], excel.Cells[j + 38, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 38, 1], excel.Cells[j + 38, 7]].Value2 = "开户行：" + companyModel["license"]["bank"];

            xSt.Range[excel.Cells[j + 39, 1], excel.Cells[j + 39, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 39, 1], excel.Cells[j + 39, 7]].Value2 = "账号：" + companyModel["license"]["account"];

            xSt.Range[excel.Cells[j + 41, 1], excel.Cells[j + 41, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 41, 1], excel.Cells[j + 41, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 41, 1], excel.Cells[j + 41, 7]].Font.Underline = true;
            xSt.Range[excel.Cells[j + 41, 1], excel.Cells[j + 41, 7]].Value2 = "请您将发票寄送到：";

            xSt.Range[excel.Cells[j + 42, 1], excel.Cells[j + 42, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 42, 1], excel.Cells[j + 42, 7]].Value2 = "发票寄送地址：" + companyModel["mailingAddress"];

            xSt.Range[excel.Cells[j + 43, 1], excel.Cells[j + 43, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 43, 1], excel.Cells[j + 43, 7]].Value2 = "收件人：" + orderModel["personInChargeName"].ToString().Trim();

            xSt.Range[excel.Cells[j + 44, 1], excel.Cells[j + 44, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 44, 1], excel.Cells[j + 44, 7]].Value2 = "电话：" + companyModel["telephone"];

            xSt.Range[excel.Cells[j + 46, 1], excel.Cells[j + 46, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 46, 1], excel.Cells[j + 46, 7]].Font.Size = 12;
            xSt.Range[excel.Cells[j + 46, 1], excel.Cells[j + 46, 7]].Font.Underline = true;
            xSt.Range[excel.Cells[j + 46, 1], excel.Cells[j + 46, 7]].Font.Italic = true;
            xSt.Range[excel.Cells[j + 46, 1], excel.Cells[j + 46, 7]].Value2 = "特别注意：请务必按照“发票寄送地址”寄送发票，不要按开票信息地址寄送！";

            xSt.Range[excel.Cells[j + 47, 1], excel.Cells[j + 47, 1]].RowHeight = 14;
            xSt.Range[excel.Cells[j + 47, 1], excel.Cells[j + 47, 7]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 47, 1], excel.Cells[j + 47, 7]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;


            excel.Visible = true;

            string path = "";
            string filename = "";

            if (strType.ToLower() == "excel")
            {
                filename = orderModel["_id"] + ".xlsx";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + ".xlsx";

                //保存excel
                xBk.SaveCopyAs(path);
            }
            else
            {
                filename = orderModel["_id"] + ".pdf";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + ".pdf";

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

        public string MoneyToUpper(string LowerMoney)
        {
            string functionReturnValue = null;
            bool IsNegative = false; // 是否是负数
            if (LowerMoney.Trim().Substring(0, 1) == "-")
            {
                // 是负数则先转为正数
                LowerMoney = LowerMoney.Trim().Remove(0, 1);
                IsNegative = true;
            }
            string strLower = null;
            string strUpart = null;
            string strUpper = null;
            int iTemp = 0;
            // 保留两位小数 123.489→123.49　　123.4→123.4
            LowerMoney = Math.Round(double.Parse(LowerMoney), 2).ToString();
            if (LowerMoney.IndexOf(".") > 0)
            {
                if (LowerMoney.IndexOf(".") == LowerMoney.Length - 2)
                {
                    LowerMoney = LowerMoney + "0";
                }
            }
            else
            {
                LowerMoney = LowerMoney + ".00";
            }
            strLower = LowerMoney;
            iTemp = 1;
            strUpper = "";
            while (iTemp <= strLower.Length)
            {
                switch (strLower.Substring(strLower.Length - iTemp, 1))
                {
                    case ".":
                        strUpart = "圆";
                        break;
                    case "0":
                        strUpart = "零";
                        break;
                    case "1":
                        strUpart = "壹";
                        break;
                    case "2":
                        strUpart = "贰";
                        break;
                    case "3":
                        strUpart = "叁";
                        break;
                    case "4":
                        strUpart = "肆";
                        break;
                    case "5":
                        strUpart = "伍";
                        break;
                    case "6":
                        strUpart = "陆";
                        break;
                    case "7":
                        strUpart = "柒";
                        break;
                    case "8":
                        strUpart = "捌";
                        break;
                    case "9":
                        strUpart = "玖";
                        break;
                }

                switch (iTemp)
                {
                    case 1:
                        strUpart = strUpart + "分";
                        break;
                    case 2:
                        strUpart = strUpart + "角";
                        break;
                    case 3:
                        strUpart = strUpart + "";
                        break;
                    case 4:
                        strUpart = strUpart + "";
                        break;
                    case 5:
                        strUpart = strUpart + "拾";
                        break;
                    case 6:
                        strUpart = strUpart + "佰";
                        break;
                    case 7:
                        strUpart = strUpart + "仟";
                        break;
                    case 8:
                        strUpart = strUpart + "万";
                        break;
                    case 9:
                        strUpart = strUpart + "拾";
                        break;
                    case 10:
                        strUpart = strUpart + "佰";
                        break;
                    case 11:
                        strUpart = strUpart + "仟";
                        break;
                    case 12:
                        strUpart = strUpart + "亿";
                        break;
                    case 13:
                        strUpart = strUpart + "拾";
                        break;
                    case 14:
                        strUpart = strUpart + "佰";
                        break;
                    case 15:
                        strUpart = strUpart + "仟";
                        break;
                    case 16:
                        strUpart = strUpart + "万";
                        break;
                    default:
                        strUpart = strUpart + "";
                        break;
                }

                strUpper = strUpart + strUpper;
                iTemp = iTemp + 1;
            }

            strUpper = strUpper.Replace("零拾", "零");
            strUpper = strUpper.Replace("零佰", "零");
            strUpper = strUpper.Replace("零仟", "零");
            strUpper = strUpper.Replace("零零零", "零");
            strUpper = strUpper.Replace("零零", "零");
            strUpper = strUpper.Replace("零角零分", "整");
            strUpper = strUpper.Replace("零分", "整");
            strUpper = strUpper.Replace("零角", "零");
            strUpper = strUpper.Replace("零亿零万零圆", "亿圆");
            strUpper = strUpper.Replace("亿零万零圆", "亿圆");
            strUpper = strUpper.Replace("零亿零万", "亿");
            strUpper = strUpper.Replace("零万零圆", "万圆");
            strUpper = strUpper.Replace("零亿", "亿");
            strUpper = strUpper.Replace("零万", "万");
            strUpper = strUpper.Replace("零圆", "圆");
            strUpper = strUpper.Replace("零零", "零");

            // 对壹圆以下的金额的处理
            if (strUpper.Substring(0, 1) == "圆")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "零")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "角")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "分")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "整")
            {
                strUpper = "零圆整";
            }
            functionReturnValue = strUpper;

            if (IsNegative == true)
            {
                return "负" + functionReturnValue;
            }
            else
            {
                return functionReturnValue;
            }
        }
    }
}