using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using TAX.HTKK.Utilities.Core;

namespace HTKK.Forms;

public class frmHoTroKeKhai : Form
{
	private IContainer components = null;

	private WebBrowser webBrowser1;

	public frmHoTroKeKhai()
	{
		InitializeComponent();
	}

	private void FrmWebBrowse_Load(object sender, EventArgs e)
	{
		try
		{
			webBrowser1.Navigate("https://nhantokhai.gdt.gov.vn/ihtkk_nnt/loginForm.jsp");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
	{
		try
		{
			webBrowser1.Document.GetElementById("c2").InnerText = WinSettings.HTKKHeaderInfo.mst;
		}
		catch (Exception)
		{
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bf: Expected O, but got Unknown
		//IL_0112: Unknown result type (might be due to invalid IL or missing references)
		//IL_011c: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmHoTroKeKhai));
		webBrowser1 = new WebBrowser();
		((Control)this).SuspendLayout();
		((Control)webBrowser1).Dock = (DockStyle)5;
		((Control)webBrowser1).Location = new Point(0, 0);
		((Control)webBrowser1).MinimumSize = new Size(20, 20);
		((Control)webBrowser1).Name = "webBrowser1";
		((Control)webBrowser1).Size = new Size(1075, 661);
		((Control)webBrowser1).TabIndex = 0;
		webBrowser1.Url = new Uri("", UriKind.Relative);
		webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(1075, 661);
		((Control)this).Controls.Add((Control)(object)webBrowser1);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Control)this).Name = "frmHoTroKeKhai";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Hỗ trợ kê khai trực tuyến";
		((Form)this).Load += FrmWebBrowse_Load;
		((Control)this).ResumeLayout(false);
	}
}
