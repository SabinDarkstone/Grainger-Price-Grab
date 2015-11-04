using System;
using System.Windows.Forms;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Grainger_Price_Grab {
	public partial class ProofOfConcept : Form {

		public ProofOfConcept() {
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e) {

			openFileDialog1.ShowDialog();
			string fileName = openFileDialog1.FileName;

			Excel.Application xlApp;
			Excel.Workbook xlWorkbook;
			Excel.Worksheet xlWorksheet;
			Excel.Range range, used;

			string str;
			int rCnt = 0;

			xlApp = new Excel.Application();

			xlWorkbook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
			xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

			range = xlWorksheet.Range[xlWorksheet.Cells[1, 1], xlWorksheet.Cells[xlWorksheet.UsedRange.Rows.Count, 1]];
			used = xlWorksheet.UsedRange;

			if (xlWorksheet.Cells[1, 2].Value != "Nomenclature") {
				used = xlWorksheet.Range[xlWorksheet.Cells[1, 1], xlWorksheet.Cells[range.Rows.Count, 2]];
			}

			for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++) {
				str = (string)(range.Cells[rCnt, 1] as Excel.Range).Value;
				textBox3.Text = str;

				xlWorksheet.Cells[1, used.Columns.Count + 1].Value = DateTime.Today.ToShortDateString();
				xlWorksheet.Cells[1, 2].Value = "Nomenclature";

				if (RunQuery()) {
					xlWorksheet.Cells[rCnt, 2].Value = textBox2.Text;
					xlWorksheet.Cells[rCnt, used.Columns.Count + 1].Value = textBox1.Text;
				} else {
					xlWorksheet.Cells[rCnt, 2].Value = "Error";
					xlWorksheet.Cells[rCnt, used.Columns.Count + 1].Value = "Error";
				}
			}

			xlWorkbook.SaveAs(fileName);
			xlWorkbook.Close();
			xlApp.Quit();

			ReleaseObject(xlWorksheet);
			ReleaseObject(xlWorkbook);
			ReleaseObject(xlApp);
		}

		/// <summary>
		/// Executes a query for a single node from Grainger site
		/// </summary>
		private bool RunQuery() {
			string url = "http://www.grainger.com/product/ACUITY-LITHONIA-Emerg-Light-" + textBox3.Text;
			HtmlWeb web = new HtmlWeb();
			HtmlAgilityPack.HtmlDocument doc = null;

			try {
				doc = web.Load(url);
			} catch (System.Net.WebException web_esc) {
				MessageBox.Show(web_esc.ToString(), "Web Exception");
				return false;
			}

			try {
				HtmlNode productInfo = doc.DocumentNode.SelectSingleNode("//*[@id=\"productPage\"]/div[1]/div[2]");

				var info = productInfo.InnerText;
				var priceBegin = info.IndexOf("$");
				var price = info.Substring(priceBegin, info.IndexOf("\n", priceBegin) - priceBegin);
				var nameBegin = info.IndexOf("Mfr. Model #") + 13;
				var name = info.Substring(info.IndexOf("Mfr. Model #") + 13, (info.IndexOf("\n", nameBegin)) - nameBegin);

				textBox1.Text = price;
				textBox2.Text = name;
				return true;

			} catch (NullReferenceException null_exc) {
				// MessageBox.Show(null_exc.ToString(), "Null Reference");
				return false;
			} catch (Exception exc) {
				// MessageBox.Show(exc.ToString(), "Generic Exception");
				return false;
			}
		}

		private void ReleaseObject(object obj) {
			try {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			} catch (Exception exc) {
				obj = null;
			} finally {
				GC.Collect();
			}
		}
	}
}