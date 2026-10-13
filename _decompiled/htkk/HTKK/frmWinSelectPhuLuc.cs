using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.Model;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmWinSelectPhuLuc : Form
{
	private FpSpread winFpSpread;

	private static frmHTKKSystems frmHTKKSystems;

	private IContainer components = null;

	private Panel pnVT3Button;

	private Button cmdClose;

	private Panel f;

	private Button btnDongY;

	private Panel pnVT2_PhuLucKeKhai;

	private CheckBox chkChonPhuLucKeKhaiAll;

	private FpSpread axfpsPhuLuc;

	private SheetView axfpsPhuLuc_Sheet1;

	private Panel panel1;

	private Label label3;

	public frmWinSelectPhuLuc(FpSpread winFpSpread)
	{
		this.winFpSpread = winFpSpread;
		InitializeComponent();
		WinSpreadProcess.WinSetupSpread(axfpsPhuLuc);
		GridLine horizontalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
		GridLine verticalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
		axfpsPhuLuc.Sheets[0].HorizontalGridLine = horizontalGridLine;
		axfpsPhuLuc.Sheets[0].VerticalGridLine = verticalGridLine;
		axfpsPhuLuc.Sheets[0].RowHeader.Visible = false;
		axfpsPhuLuc.Sheets[0].ColumnHeader.Visible = false;
		axfpsPhuLuc.BorderStyle = (BorderStyle)1;
		axfpsPhuLuc.RowSplitBoxPolicy = SplitBoxPolicy.Never;
		axfpsPhuLuc.HorizontalScrollBarPolicy = ScrollBarPolicy.Never;
		axfpsPhuLuc.Sheets[0].HorizontalGridLine = new GridLine();
		axfpsPhuLuc.FocusRenderer = null;
		if (!WinSpreadProcess.BindDataToPhuLucSpread(axfpsPhuLuc.Sheets[0], WinSettings.xmlMenuNode.FirstChild.ChildNodes, chkChonPhuLucKeKhaiAll, winFpSpread))
		{
			((Form)this).Close();
		}
	}

	private void btnDongY_Click(object sender, EventArgs e)
	{
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b7: Expected O, but got Unknown
		//IL_0087: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (WinSettings.MenuID.Equals("N8"))
			{
				int row = 20;
				if (Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) != 2021 && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[row, 0].Value))
				{
					frmMessage frmMessage = new frmMessage("0669", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
			}
			if (!WinDoWorkEvent.CheckChonPlBctc(WinSettings.TTPeriodInfo, axfpsPhuLuc))
			{
				return;
			}
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			int num = ((!(WinSettings.ToKhaiInfo.KieuKHBS == "3")) ? 1 : 2);
			for (int i = 1; i < val.Count - num; i++)
			{
				if (axfpsPhuLuc.Sheets[0].Cells[i - 1, 0].Value != null && (!WinFnProcess.IsWorking(((XmlNode)val[i].Attributes["SheetName"]).Value) & (bool)axfpsPhuLuc.Sheets[0].Cells[i - 1, 0].Value))
				{
					HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val, i, getTemplate: true, WinSettings.TTPeriodInfo);
					WinSettings.ToKhaiInfo.SheetInfos[i] = hTKKSheetInfo;
					WinSpreadProcess.ProcessSheet(i, winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
					winFpSpread.Sheets[i].Visible = true;
					WinSettings.DicSheet.Add(((XmlNode)val[i].Attributes["SheetName"]).Value, 2);
					WinDoWorkEvent.RefreshSheet(winFpSpread, i, 1);
				}
			}
			((Form)this).Close();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			((Form)this).Close();
		}
	}

	private void cmdClose_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void chkChonPhuLucKeKhaiAll_CheckedChanged(object sender, EventArgs e)
	{
	}

	private void CheckAll(SheetView sview, bool value)
	{
		for (int i = 0; i < axfpsPhuLuc.Sheets[0].RowCount; i++)
		{
			if (!sview.Cells[i, 0].Locked)
			{
				axfpsPhuLuc.Sheets[0].Cells[i, 0].Value = value;
			}
		}
		chkChonPhuLucKeKhaiAll.Checked = value;
	}

	private void chkChonPhuLucKeKhaiAll_Click(object sender, EventArgs e)
	{
		if (chkChonPhuLucKeKhaiAll.Checked)
		{
			CheckAll(axfpsPhuLuc.Sheets[0], value: true);
		}
		else
		{
			CheckAll(axfpsPhuLuc.Sheets[0], value: false);
		}
	}

	private void axfpsPhuLuc_ButtonClicked(object sender, EditorNotifyEventArgs e)
	{
		try
		{
			int num = -1;
			int num2 = 2;
			for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
			{
				if (axfpsPhuLuc.ActiveSheet.Cells[i, e.Column].Value.ToString().ToLower() == "false")
				{
					num2 = 0;
					break;
				}
				num2 = 1;
			}
			for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
			{
				if (axfpsPhuLuc.ActiveSheet.Cells[i, e.Column].Value.ToString().ToLower() == "true")
				{
					num = i;
				}
			}
			if (WinSettings.TTPeriodInfo.menuID.Equals("68") || WinSettings.TTPeriodInfo.menuID.Equals("14") || WinSettings.TTPeriodInfo.menuID.Equals("13") || WinSettings.TTPeriodInfo.menuID.Equals("18") || WinSettings.TTPeriodInfo.menuID.Equals("62"))
			{
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
				for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
				{
					if (i == num || num == -1)
					{
						axfpsPhuLuc.ActiveSheet.Cells[i, 0].Locked = false;
					}
					else
					{
						axfpsPhuLuc.ActiveSheet.Cells[i, 0].Locked = true;
					}
				}
			}
			else
			{
				switch (num2)
				{
				case 2:
					((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
					chkChonPhuLucKeKhaiAll.Checked = true;
					break;
				case 1:
					((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
					chkChonPhuLucKeKhaiAll.Checked = true;
					break;
				default:
					((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
					chkChonPhuLucKeKhaiAll.Checked = false;
					break;
				}
			}
			if (axfpsPhuLuc.ActiveSheet.Cells[e.Row, e.Column].Value.ToString().ToLower() == "false")
			{
				chkChonPhuLucKeKhaiAll.Checked = false;
				return;
			}
			for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
			{
				if (axfpsPhuLuc.ActiveSheet.Cells[i, e.Column].Value.ToString().ToLower() == "false")
				{
					chkChonPhuLucKeKhaiAll.Checked = false;
					return;
				}
			}
			chkChonPhuLucKeKhaiAll.Checked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmWinSelectPhuLuc_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
			FormProcess.ChildFormClosed((Form)(object)WinSettings.myWinToKhai, (Control)(object)WinSettings.myWinToKhai.panelChild);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_07be: Unknown result type (might be due to invalid IL or missing references)
		//IL_07c8: Expected O, but got Unknown
		//IL_0875: Unknown result type (might be due to invalid IL or missing references)
		//IL_087f: Expected O, but got Unknown
		//IL_0988: Unknown result type (might be due to invalid IL or missing references)
		//IL_0992: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmWinSelectPhuLuc));
		pnVT3Button = new Panel();
		btnDongY = new Button();
		cmdClose = new Button();
		f = new Panel();
		pnVT2_PhuLucKeKhai = new Panel();
		axfpsPhuLuc = new FpSpread();
		axfpsPhuLuc_Sheet1 = new SheetView();
		chkChonPhuLucKeKhaiAll = new CheckBox();
		panel1 = new Panel();
		label3 = new Label();
		((Control)pnVT3Button).SuspendLayout();
		((Control)f).SuspendLayout();
		((Control)pnVT2_PhuLucKeKhai).SuspendLayout();
		((ISupportInitialize)axfpsPhuLuc).BeginInit();
		((ISupportInitialize)axfpsPhuLuc_Sheet1).BeginInit();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)pnVT3Button).Controls.Add((Control)(object)btnDongY);
		((Control)pnVT3Button).Controls.Add((Control)(object)cmdClose);
		((Control)pnVT3Button).Location = new Point(3, 102);
		((Control)pnVT3Button).Name = "pnVT3Button";
		((Control)pnVT3Button).Size = new Size(338, 28);
		((Control)pnVT3Button).TabIndex = 1;
		((Control)btnDongY).Location = new Point(94, 2);
		((Control)btnDongY).Name = "btnDongY";
		((Control)btnDongY).Size = new Size(80, 25);
		((Control)btnDongY).TabIndex = 8;
		((Control)btnDongY).Text = "Đồn&g ý";
		((ButtonBase)btnDongY).UseVisualStyleBackColor = true;
		((Control)btnDongY).Click += btnDongY_Click;
		((Control)cmdClose).Location = new Point(178, 2);
		((Control)cmdClose).Name = "cmdClose";
		((Control)cmdClose).Size = new Size(80, 25);
		((Control)cmdClose).TabIndex = 6;
		((Control)cmdClose).Text = "Đó&ng";
		((ButtonBase)cmdClose).UseVisualStyleBackColor = true;
		((Control)cmdClose).Click += cmdClose_Click;
		((Control)f).Controls.Add((Control)(object)pnVT2_PhuLucKeKhai);
		((Control)f).Controls.Add((Control)(object)pnVT3Button);
		((Control)f).Location = new Point(3, 24);
		((Control)f).Name = "f";
		((Control)f).Size = new Size(345, 136);
		((Control)f).TabIndex = 54;
		pnVT2_PhuLucKeKhai.BorderStyle = (BorderStyle)1;
		((Control)pnVT2_PhuLucKeKhai).Controls.Add((Control)(object)axfpsPhuLuc);
		((Control)pnVT2_PhuLucKeKhai).Controls.Add((Control)(object)chkChonPhuLucKeKhaiAll);
		((Control)pnVT2_PhuLucKeKhai).Location = new Point(3, 4);
		((Control)pnVT2_PhuLucKeKhai).Name = "pnVT2_PhuLucKeKhai";
		((Control)pnVT2_PhuLucKeKhai).Size = new Size(338, 95);
		((Control)pnVT2_PhuLucKeKhai).TabIndex = 2;
		((Control)axfpsPhuLuc).AccessibleDescription = "axfpsPhuLuc, Sheet1, Row 0, Column 0, ";
		((Control)axfpsPhuLuc).Location = new Point(4, 24);
		axfpsPhuLuc.Name = "axfpsPhuLuc";
		axfpsPhuLuc.Sheets.AddRange(new SheetView[1] { axfpsPhuLuc_Sheet1 });
		((Control)axfpsPhuLuc).Size = new Size(327, 66);
		((Control)axfpsPhuLuc).TabIndex = 2;
		axfpsPhuLuc.ButtonClicked += axfpsPhuLuc_ButtonClicked;
		axfpsPhuLuc_Sheet1.Reset();
		axfpsPhuLuc_Sheet1.SheetName = "Sheet1";
		axfpsPhuLuc_Sheet1.ReferenceStyle = ReferenceStyle.R1C1;
		axfpsPhuLuc_Sheet1.ColumnCount = 1;
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerFooterDefaultEnhanced";
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.Columns.Get(0).Width = 308f;
		axfpsPhuLuc_Sheet1.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.DefaultStyle.Parent = "DataAreaDefault";
		axfpsPhuLuc_Sheet1.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefaultEnhanced";
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.Parent = "RowHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.RowHeader.Columns.Default.Resizable = false;
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.SheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.SheetCornerStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.SheetCornerStyle.Parent = "CornerDefaultEnhanced";
		axfpsPhuLuc_Sheet1.SheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.ReferenceStyle = ReferenceStyle.A1;
		((Control)chkChonPhuLucKeKhaiAll).AutoSize = true;
		((Control)chkChonPhuLucKeKhaiAll).Location = new Point(5, 3);
		((Control)chkChonPhuLucKeKhaiAll).Name = "chkChonPhuLucKeKhaiAll";
		((Control)chkChonPhuLucKeKhaiAll).Size = new Size(127, 17);
		((Control)chkChonPhuLucKeKhaiAll).TabIndex = 1;
		((Control)chkChonPhuLucKeKhaiAll).Text = "Chọn phụ lục kê khai";
		((ButtonBase)chkChonPhuLucKeKhaiAll).UseVisualStyleBackColor = true;
		((Control)chkChonPhuLucKeKhaiAll).Click += chkChonPhuLucKeKhaiAll_Click;
		((Control)panel1).BackgroundImage = (Image)componentResourceManager.GetObject("panel1.BackgroundImage");
		((Control)panel1).BackgroundImageLayout = (ImageLayout)3;
		((Control)panel1).Controls.Add((Control)(object)label3);
		((Control)panel1).Dock = (DockStyle)1;
		((Control)panel1).Location = new Point(0, 0);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(350, 21);
		((Control)panel1).TabIndex = 55;
		((Control)label3).AutoSize = true;
		((Control)label3).BackColor = Color.Transparent;
		((Control)label3).Font = new Font("Tahoma", 8.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label3).ForeColor = Color.White;
		((Control)label3).Location = new Point(1, 3);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(126, 13);
		((Control)label3).TabIndex = 0;
		((Control)label3).Text = "Thêm phụ lục kê khai";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(350, 163);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)f);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmWinSelectPhuLuc";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Thêm phụ lục kê khai";
		((Form)this).FormClosed += new FormClosedEventHandler(frmWinSelectPhuLuc_FormClosed);
		((Control)pnVT3Button).ResumeLayout(false);
		((Control)f).ResumeLayout(false);
		((Control)pnVT2_PhuLucKeKhai).ResumeLayout(false);
		((Control)pnVT2_PhuLucKeKhai).PerformLayout();
		((ISupportInitialize)axfpsPhuLuc).EndInit();
		((ISupportInitialize)axfpsPhuLuc_Sheet1).EndInit();
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
