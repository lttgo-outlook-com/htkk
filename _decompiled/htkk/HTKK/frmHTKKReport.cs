using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using FarPoint.Win.Spread;

namespace HTKK;

public class frmHTKKReport : Form
{
	private IContainer components = null;

	private FpSpread winFpSpreadReport;

	private SheetView winFpSpreadReport_Sheet1;

	public frmHTKKReport()
	{
		InitializeComponent();
	}

	private void frmHTKKReport_Load(object sender, EventArgs e)
	{
		bool flag = WinFnProcess.WinLoadFromFile(winFpSpreadReport, WinSettings.ToKhaiInfo.ReportTemplates);
		WinSpreadProcess.fpSpDisplayReport(winFpSpreadReport, WinSettings.ToKhaiInfo.SheetInfos);
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
		//IL_0141: Unknown result type (might be due to invalid IL or missing references)
		//IL_014b: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmHTKKReport));
		winFpSpreadReport = new FpSpread();
		winFpSpreadReport_Sheet1 = new SheetView();
		((ISupportInitialize)winFpSpreadReport).BeginInit();
		((ISupportInitialize)winFpSpreadReport_Sheet1).BeginInit();
		((Control)this).SuspendLayout();
		((Control)winFpSpreadReport).AccessibleDescription = "";
		((Control)winFpSpreadReport).Dock = (DockStyle)5;
		((Control)winFpSpreadReport).Location = new Point(0, 0);
		winFpSpreadReport.Name = "winFpSpreadReport";
		winFpSpreadReport.Sheets.AddRange(new SheetView[1] { winFpSpreadReport_Sheet1 });
		((Control)winFpSpreadReport).Size = new Size(1013, 634);
		((Control)winFpSpreadReport).TabIndex = 0;
		winFpSpreadReport_Sheet1.Reset();
		winFpSpreadReport_Sheet1.SheetName = "Sheet1";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(1013, 634);
		((Control)this).Controls.Add((Control)(object)winFpSpreadReport);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Control)this).Name = "frmHTKKReport";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "frmHTKKReport";
		((Form)this).Load += frmHTKKReport_Load;
		((ISupportInitialize)winFpSpreadReport).EndInit();
		((ISupportInitialize)winFpSpreadReport_Sheet1).EndInit();
		((Control)this).ResumeLayout(false);
	}
}
