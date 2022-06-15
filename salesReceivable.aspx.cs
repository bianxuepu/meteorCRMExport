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
    public partial class salesReceivable : System.Web.UI.Page
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

            var data = m.data;
            var number = Convert.ToInt32(m.number);

            OutputExcel(data, number);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic data, int number)
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

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].Font.Bold = true;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].Merge(false);
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].RowHeight = 40;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].Font.Size = 16;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].Value2 = " 应收查询树型表";

            excel.Cells[2, 1] = "行号";
            excel.Cells[2, 2] = "结算单位名称";
            excel.Cells[2, 3] = "应收款";
            excel.Cells[2, 4] = "已开票金额";
            excel.Cells[2, 5] = "待审发票额";
            excel.Cells[2, 6] = "未提发票额";
            excel.Cells[2, 7] = "已开票未过账";
            excel.Cells[2, 8] = "有问题金额";
            excel.Cells[2, 9] = "备注";

            int j = 2;
            for (int i = 0; i < number; i++)
            {
                j++;

                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 2]].NumberFormat = "@";
                xSt.Range[excel.Cells[j, 9], excel.Cells[j, 9]].NumberFormat = "@";

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j, 3], excel.Cells[j, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j, 9], excel.Cells[j, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                excel.Cells[j, 1] = Convert.ToString(i+1);
                excel.Cells[j, 2] = data[i]["customerName"];
                excel.Cells[j, 3] = data[i]["total"];
                excel.Cells[j, 4] = data[i]["kaipiao"];
                excel.Cells[j, 5] = data[i]["daishen"];
                excel.Cells[j, 6] = data[i]["weiti"];
                excel.Cells[j, 7] = data[i]["weiguozhang"];
                excel.Cells[j, 8] = Convert.ToDecimal(data[i]["total"])- Convert.ToDecimal(data[i]["kaipiao"])- Convert.ToDecimal(data[i]["daishen"])- Convert.ToDecimal(data[i]["weiti"])+ Convert.ToDecimal(data[i]["weiguozhang"]);
            }

            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 10]].Columns.AutoFit();//行高根据内容自动调整

            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 10]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 10]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[1, 10], excel.Cells[j, 10]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 10]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 10]].RowHeight = 22;

            excel.Visible = true;

            string path = "";
            string filename = "";

            filename = "salesReceivable" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            path = Server.MapPath("~/") + "temp\\" + filename;

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