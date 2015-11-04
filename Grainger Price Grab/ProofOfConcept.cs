using System;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace Grainger_Price_Grab {
	public partial class ProofOfConcept : Form {
		public ProofOfConcept() {
			InitializeComponent();
		}

		/// <summary>
		/// Execute single node scrape from given grainger item
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button1_Click(object sender, EventArgs e) {
			string url = "http://www.grainger.com/product/ACUITY-LITHONIA-Emerg-Light-" + textBox3.Text;
			HtmlWeb web = new HtmlWeb();
			HtmlAgilityPack.HtmlDocument doc = null;

			try {
				doc = web.Load(url);
			} catch (System.Net.WebException web_esc) {
				MessageBox.Show(web_esc.ToString(), "Web Exception");
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

			} catch (NullReferenceException null_exc) {
				MessageBox.Show(null_exc.ToString(), "Null Reference");
			} catch (Exception exc) {
				MessageBox.Show(exc.ToString(), "Generic Exception");
			}
		}

	}
}