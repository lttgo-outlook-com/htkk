using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.Model;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmHTKKPrintOption : Form
{
	private FpSpread winFpSpreadForm;

	private FpSpread winFpSpreadReport = new FpSpread();

	private List<HTKKCMCSheetInfo> TmpSheetInfo = new List<HTKKCMCSheetInfo>();

	private string[] arrBarcode;

	private Dictionary<string, string> dicBarcodeEachSheet = new Dictionary<string, string>();

	private Dictionary<string, int> dicTotalPagesASheet = new Dictionary<string, int>();

	private IContainer components = null;

	private ComboBox cbbChonMayIn;

	private Label lblChonMayIn;

	private Label lblTrangIn;

	private Label lblSoBanIn;

	private TextBox txtSoBanIn;

	private TextBox txtTrangIn;

	private Panel pnPrintOption;

	private Button btnIn;

	private Button btnXemTruoc;

	private Button btnDong;

	private Panel pnButton;

	private CheckBox ckInThongTinDieuChinh;

	private Panel panel1;

	private Label label3;

	public frmHTKKPrintOption(string[] _arrBarcode, FpSpread _winFpSpreadForm, Dictionary<string, string> dicBarcodeEachSheet)
	{
		InitializeComponent();
		((ISupportInitialize)winFpSpreadReport).BeginInit();
		winFpSpreadReport.Name = "winFpSpreadReport";
		((ISupportInitialize)winFpSpreadReport).EndInit();
		arrBarcode = _arrBarcode;
		winFpSpreadForm = _winFpSpreadForm;
		this.dicBarcodeEachSheet = dicBarcodeEachSheet;
		dicTotalPagesASheet = new Dictionary<string, int>();
	}

	private void btnIn_Click(object sender, EventArgs e)
	{
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Expected O, but got Unknown
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0558: Unknown result type (might be due to invalid IL or missing references)
		//IL_0257: Unknown result type (might be due to invalid IL or missing references)
		//IL_025e: Expected O, but got Unknown
		WinSettings.isActionPrint = true;
		List<HTKKCMCSheetInfo> list = new List<HTKKCMCSheetInfo>();
		list = WinSettings.ToKhaiInfo.SheetInfos;
		Dictionary<string, int> dictionary = new Dictionary<string, int>();
		dictionary = WinSettings.DicSheet;
		try
		{
			int num = 0;
			if (((Control)cbbChonMayIn).Text == "")
			{
				frmMessage frmMessage = new frmMessage("0026", "3", "1", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			bool flag = WinFnProcess.WinLoadFromFile(winFpSpreadReport, WinSettings.ToKhaiInfo.ReportTemplates);
			WinSpreadProcess.fpSpDisplayReport(winFpSpreadReport, WinSettings.ToKhaiInfo.SheetInfos);
			SpreadPrint2.inThongTinDL(winFpSpreadReport);
			SpreadPrint2.SetupPageBreak(winFpSpreadReport);
			dicTotalPagesASheet = new Dictionary<string, int>();
			int totalPage = 0;
			int num2 = 0;
			PrintDocument printDoc = new PrintDocument();
			if (WinSettings.MenuID == "70")
			{
				SpreadPrint2.setMlNs(winFpSpreadReport, printDoc);
			}
			for (int i = 0; i < winFpSpreadReport.Sheets.Count; i++)
			{
				if (i >= WinSettings.ToKhaiInfo.SheetInfos.Count || WinSettings.ToKhaiInfo.SheetInfos[i] == null || !WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName))
				{
					continue;
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu || i == 0 || ((WinSettings.xmlMenuNode.Attributes["ParentID"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["ParentID"]).Value) == "101_6" || WinSettings.MenuID == "57")
				{
					if (!ThayDoiThuePhaiNopKHBS(winFpSpreadReport, i) && WinSettings.ToKhaiInfo.SheetInfos[i].SheetName == "KHBS")
					{
						winFpSpreadReport.Sheets[i].Visible = false;
						continue;
					}
					PageSetup pageSetup = new PageSetup(winFpSpreadReport, i, showPrintButton: false, showPreviewButton: false);
					PrintDocument val = new PrintDocument();
					val.PrinterSettings.PrinterName = ((Control)cbbChonMayIn).Text;
					winFpSpreadReport.Sheets[i].PrintInfo.Margin = new PrintMargin(val.PrinterSettings.DefaultPageSettings.Margins.Left, val.PrinterSettings.DefaultPageSettings.Margins.Top, val.PrinterSettings.DefaultPageSettings.Margins.Right, val.PrinterSettings.DefaultPageSettings.Margins.Bottom, winFpSpreadReport.Sheets[i].PrintInfo.Margin.Header, winFpSpreadReport.Sheets[i].PrintInfo.Margin.Footer);
					SpreadPrint2.PortraitWidth = (SpreadPrint2.LandscapeHeight = val.PrinterSettings.DefaultPageSettings.Bounds.Width);
					SpreadPrint2.PortraitHeight = (SpreadPrint2.LandscapeWidth = val.PrinterSettings.DefaultPageSettings.Bounds.Height);
					winFpSpreadReport.Sheets[i].PrintInfo.Centering = Centering.Horizontal;
					Rectangle rect = ((winFpSpreadReport.Sheets[i].PrintInfo.Orientation != PrintOrientation.Landscape) ? SpreadPrint2.rectPortrait : SpreadPrint2.rectLandscape);
					Graphics g = ((Control)this).CreateGraphics();
					int ownerPrintPageCount = winFpSpreadReport.GetOwnerPrintPageCount(g, rect, i);
					dicTotalPagesASheet.Add(num2.ToString(), ownerPrintPageCount);
					num2++;
					num += ownerPrintPageCount;
					if (i == 0)
					{
						totalPage = ownerPrintPageCount;
					}
				}
				else
				{
					winFpSpreadReport.Sheets[i].Visible = false;
				}
			}
			string[] array = null;
			array = ((!(WinSettings.TTPeriodInfo.menuID == "33")) ? GetArrayBarcodeForPrint(num) : GetArrayBarcodeForPrint(totalPage));
			if (((Control)txtSoBanIn).Text != "")
			{
				for (int j = 0; j < int.Parse(((Control)txtSoBanIn).Text); j++)
				{
					if (((Control)txtTrangIn).Text != "")
					{
						List<int> list2 = new List<int>();
						list2 = StringToList.lstPages(((Control)txtTrangIn).Text, num);
						if (list2 != null)
						{
							SpreadPrint2 spreadPrint = new SpreadPrint2(winFpSpreadReport, printDoc, ((Control)cbbChonMayIn).Text, array, num, list2, _isPrint: true);
							spreadPrint.PrintPreview(this);
						}
						else
						{
							frmMessage frmMessage = new frmMessage("0456", "3", "1", "Cảnh báo");
							((Form)frmMessage).ShowDialog();
						}
					}
					else
					{
						SpreadPrint2 spreadPrint = new SpreadPrint2(winFpSpreadReport, printDoc, ((Control)cbbChonMayIn).Text, array, num, null, _isPrint: true);
						spreadPrint.PrintPreview(this);
					}
				}
				((Form)this).Close();
			}
			else
			{
				frmMessage frmMessage = new frmMessage("0033", "3", "1", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
			}
			WinSettings.ToKhaiInfo.SheetInfos = list;
			WinSettings.DicSheet = dictionary;
		}
		catch (Exception ex)
		{
			WinSettings.ToKhaiInfo.SheetInfos = list;
			WinSettings.DicSheet = dictionary;
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			((Form)this).Close();
		}
		finally
		{
			WinSettings.isActionPrint = false;
		}
	}

	private void btnDong_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void btnXemTruoc_Click(object sender, EventArgs e)
	{
		//IL_00cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d6: Expected O, but got Unknown
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_051c: Unknown result type (might be due to invalid IL or missing references)
		//IL_028a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0291: Expected O, but got Unknown
		WinSettings.isActionPrint = true;
		List<HTKKCMCSheetInfo> list = new List<HTKKCMCSheetInfo>();
		list = WinSettings.ToKhaiInfo.SheetInfos;
		try
		{
			int num = 0;
			if (((Control)cbbChonMayIn).Text == "")
			{
				frmMessage frmMessage = new frmMessage("0026", "3", "1", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			bool flag = WinFnProcess.WinLoadFromFile(winFpSpreadReport, WinSettings.ToKhaiInfo.ReportTemplates);
			WinSpreadProcess.fpSpDisplayReport(winFpSpreadReport, WinSettings.ToKhaiInfo.SheetInfos);
			winFpSpreadReport.Sheets[0].Recalculate();
			SpreadPrint2.inThongTinDL(winFpSpreadReport);
			SpreadPrint2.SetupPageBreak(winFpSpreadReport);
			int num2 = 0;
			int num3 = 0;
			dicTotalPagesASheet = new Dictionary<string, int>();
			PrintDocument printDoc = new PrintDocument();
			SpreadPrint2.setMlNs(winFpSpreadReport, printDoc);
			for (int i = 0; i < winFpSpreadReport.Sheets.Count; i++)
			{
				correctBorder1(winFpSpreadReport.Sheets[i]);
				if (i >= WinSettings.ToKhaiInfo.SheetInfos.Count || WinSettings.ToKhaiInfo.SheetInfos[i] == null || !WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName))
				{
					continue;
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu || i == 0 || WinSettings.ToKhaiInfo.SheetInfos[i].SheetName.ToUpper() == "KHBS" || ((WinSettings.xmlMenuNode.Attributes["ParentID"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["ParentID"]).Value) == "101_6" || WinSettings.MenuID == "57" || ThayDoiThuePhaiNopKHBS(winFpSpreadReport, i))
				{
					if (!ThayDoiThuePhaiNopKHBS(winFpSpreadReport, i) && WinSettings.ToKhaiInfo.SheetInfos[i].SheetName == "KHBS")
					{
						winFpSpreadReport.Sheets[i].Visible = false;
						continue;
					}
					PrintDocument val = new PrintDocument();
					val.PrinterSettings.PrinterName = ((Control)cbbChonMayIn).Text;
					winFpSpreadReport.Sheets[i].PrintInfo.Margin = new PrintMargin(val.PrinterSettings.DefaultPageSettings.Margins.Left, val.PrinterSettings.DefaultPageSettings.Margins.Top, val.PrinterSettings.DefaultPageSettings.Margins.Right, val.PrinterSettings.DefaultPageSettings.Margins.Bottom, winFpSpreadReport.Sheets[i].PrintInfo.Margin.Header, winFpSpreadReport.Sheets[i].PrintInfo.Margin.Footer);
					SpreadPrint2.PortraitWidth = (SpreadPrint2.LandscapeHeight = val.PrinterSettings.DefaultPageSettings.Bounds.Width);
					SpreadPrint2.PortraitHeight = (SpreadPrint2.LandscapeWidth = val.PrinterSettings.DefaultPageSettings.Bounds.Height);
					Rectangle rect = ((winFpSpreadReport.Sheets[i].PrintInfo.Orientation != PrintOrientation.Landscape) ? SpreadPrint2.rectPortrait : SpreadPrint2.rectLandscape);
					Graphics g = ((Control)this).CreateGraphics();
					int ownerPrintPageCount = winFpSpreadReport.GetOwnerPrintPageCount(g, rect, i);
					dicTotalPagesASheet.Add(num3.ToString(), ownerPrintPageCount);
					num3++;
					num += ownerPrintPageCount;
					if (i == 0)
					{
						num2 = ownerPrintPageCount;
					}
				}
				else
				{
					winFpSpreadReport.Sheets[i].Visible = false;
				}
			}
			string[] array = null;
			if (((Control)txtTrangIn).Text != "")
			{
				List<int> list2 = new List<int>();
				list2 = StringToList.lstPages(((Control)txtTrangIn).Text, num);
				if (list2 != null)
				{
					SpreadPrint2 spreadPrint = new SpreadPrint2(winFpSpreadReport, printDoc, ((Control)cbbChonMayIn).Text, array, num, list2, _isPrint: false);
					SpreadPrint2.inThongTinDL(winFpSpreadReport);
					spreadPrint.PrintPreview(this);
				}
				else
				{
					frmMessage frmMessage = new frmMessage("0456", "3", "1", "Cảnh báo");
					((Form)frmMessage).ShowDialog();
				}
			}
			else
			{
				SpreadPrint2 spreadPrint = new SpreadPrint2(winFpSpreadReport, printDoc, ((Control)cbbChonMayIn).Text, array, num, null, _isPrint: false);
				spreadPrint.PrintPreview(this);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			((Form)this).Close();
		}
		finally
		{
			WinSettings.ToKhaiInfo.SheetInfos = list;
			WinSettings.isActionPrint = false;
		}
	}

	private void correctBorder1(SheetView sv)
	{
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d6: Invalid comparison between Unknown and I4
		int nonEmptyRowCount = sv.NonEmptyRowCount;
		int nonEmptyColumnCount = sv.NonEmptyColumnCount;
		DefaultSheetStyleModel defaultSheetStyleModel = (DefaultSheetStyleModel)sv.Models.Style;
		StyleInfo styleInfo = null;
		for (int i = 0; i < nonEmptyRowCount; i++)
		{
			for (int j = 0; j < nonEmptyColumnCount; j++)
			{
				styleInfo = defaultSheetStyleModel.GetCompositeInfo(i, j, -1, null);
				if (styleInfo.Border is ComplexBorder)
				{
					ComplexBorder complexBorder = (ComplexBorder)styleInfo.Border;
					if (complexBorder.Bottom.Width == 1 && (int)complexBorder.Bottom.DashStyle == 0 && complexBorder.Left.Width == 1 && (int)complexBorder.Left.DashStyle == 0 && complexBorder.Top.Width == 1 && (int)complexBorder.Top.DashStyle == 0 && complexBorder.Right.Width == 1 && (int)complexBorder.Right.DashStyle == 0)
					{
						sv.Cells[i, j].Border = new LineBorder(complexBorder.Bottom.Color, 1);
					}
				}
			}
		}
	}

	private void frmHTKKPrintOption_Load(object sender, EventArgs e)
	{
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		foreach (string installedPrinter in PrinterSettings.InstalledPrinters)
		{
			cbbChonMayIn.Items.Add((object)installedPrinter);
			PrinterSettings val = new PrinterSettings();
			val.PrinterName = installedPrinter;
			if (val.IsDefaultPrinter)
			{
				cbbChonMayIn.SelectedText = installedPrinter;
			}
		}
		((Control)txtSoBanIn).Text = "1";
		string maToKhai = WinSettings.ToKhaiInfo.MaToKhai;
		if (maToKhai.Equals("33"))
		{
			((Control)ckInThongTinDieuChinh).Visible = true;
		}
		else
		{
			((Control)ckInThongTinDieuChinh).Visible = false;
		}
	}

	private void txtSoBanIn_TextChanged(object sender, EventArgs e)
	{
	}

	protected string[] GetArrayBarcodeForPrint(int totalPage)
	{
		try
		{
			string[] array;
			if (arrBarcode.Length == 1)
			{
				array = arrBarcode;
			}
			else
			{
				int num = 0;
				int num2 = 0;
				int length = arrBarcode[1].Length;
				int num3 = ((length % totalPage != 0) ? (length / totalPage + 1) : (length / totalPage));
				int num4 = 1;
				int num5;
				for (num5 = num3; num5 > 960; num5 = ((num3 % num4 != 0) ? (num3 / num4 + 1) : (num3 / num4)))
				{
					num4++;
				}
				num = totalPage * num4;
				array = new string[num];
				arrBarcode[0] = arrBarcode[0].Replace("$$$", num.ToString().PadLeft(3, '0'));
				if (dicBarcodeEachSheet != null && dicBarcodeEachSheet.Count == 1)
				{
					for (int i = 0; i < num; i++)
					{
						string text = arrBarcode[0].Replace("###", (i + 1).ToString().PadLeft(3, '0'));
						if (i < num - 1)
						{
							array[i] = text + arrBarcode[1].Substring(num2, num5) + "#";
						}
						else
						{
							array[i] = text + arrBarcode[1].Substring(num2) + "#";
						}
						num2 += num5;
					}
					return array;
				}
				Dictionary<string, int> dictionary = new Dictionary<string, int>();
				int num6 = 0;
				foreach (string key in dicBarcodeEachSheet.Keys)
				{
					int length2 = dicBarcodeEachSheet[key].Length;
					if (length2 <= num5)
					{
						dictionary.Add(key, num5);
						num6++;
						continue;
					}
					int value = 0;
					dicTotalPagesASheet.TryGetValue(key, out value);
					int num7 = ((length2 % num5 == 0) ? (length2 / num5) : (length2 / num5));
					if (value <= 0)
					{
						continue;
					}
					if (num7 == value)
					{
						num6 += num7;
						int num8 = length2 % num5;
						int num9 = num8 / num7 + 1;
						int num10 = num9 + num5;
						if (num10 > length2)
						{
							num10 = length2;
						}
						dictionary.Add(key, num10);
					}
					else
					{
						num6 += value;
						int num10 = ((length2 % value == 0) ? (length2 / value) : (length2 / value + 1));
						dictionary.Add(key, num10);
					}
				}
				int num11 = 0;
				bool flag = false;
				for (int i = 0; i < num; i++)
				{
					string text = arrBarcode[0].Replace("###", (i + 1).ToString().PadLeft(3, '0'));
					string text2 = (dicBarcodeEachSheet.ContainsKey(num11.ToString()) ? dicBarcodeEachSheet[num11.ToString()] : string.Empty);
					if (i < num - 1)
					{
						int num12 = (dictionary.ContainsKey(num11.ToString()) ? dictionary[num11.ToString()] : 0);
						if (num11 == dictionary.Count - 1 && !flag)
						{
							num12 = text2.Length / (num - i + 1) + 1;
							flag = true;
						}
						if (num2 + num12 <= text2.Length)
						{
							array[i] = text + text2.Substring(num2, num12) + "#";
							num2 += num12;
							if (num2 >= text2.Length)
							{
								num11++;
								num2 = 0;
							}
						}
						else
						{
							array[i] = text + text2.Substring(num2, text2.Length - num2) + "#";
							num11++;
							num2 = 0;
						}
					}
					else
					{
						array[i] = text + text2.Substring(num2) + "#";
					}
				}
			}
			return array;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void frmHTKKPrintOption_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
			winFpSpreadForm.ResumeLayout();
			FormProcess.ChildFormClosed((Form)(object)WinSettings.myWinToKhai, (Control)(object)WinSettings.myWinToKhai.panelChild);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool ThayDoiThuePhaiNopKHBS(FpSpread winFpSpread, int sheetIndex)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("BH");
			if (WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName != "KHBS" || !winFpSpread.Sheets[sheetIndex].Visible)
			{
				return false;
			}
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					if (ConvertHelpers.ConvertToDouble(sheetView.Cells[dataLocationTemplateRowStart + j, num - 1].Value) != 0.0)
					{
						return true;
					}
				}
			}
			if (WinSettings.MenuID == "01")
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("PathBH17"));
				if (ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) != 0.0)
				{
					return true;
				}
				hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("PathBH18"));
				if (ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) != 0.0)
				{
					return true;
				}
			}
			else
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct6"));
				if (ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) != 0.0)
				{
					return true;
				}
			}
			List<HTKKCMCRowValue> rowsValues = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables[0].RowsValues;
			List<HTKKCMCRowValue> rowsValues2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables[1].RowsValues;
			if (rowsValues.Count > 1 || rowsValues2.Count > 1)
			{
				return true;
			}
			if (rowsValues[0].RowValue[5] != "0")
			{
				return true;
			}
			if (rowsValues2[0].RowValue[5] != "0")
			{
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
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
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_069a: Unknown result type (might be due to invalid IL or missing references)
		//IL_06a4: Expected O, but got Unknown
		//IL_0751: Unknown result type (might be due to invalid IL or missing references)
		//IL_075b: Expected O, but got Unknown
		//IL_0889: Unknown result type (might be due to invalid IL or missing references)
		//IL_0893: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmHTKKPrintOption));
		cbbChonMayIn = new ComboBox();
		lblChonMayIn = new Label();
		lblTrangIn = new Label();
		lblSoBanIn = new Label();
		txtSoBanIn = new TextBox();
		txtTrangIn = new TextBox();
		pnPrintOption = new Panel();
		ckInThongTinDieuChinh = new CheckBox();
		btnIn = new Button();
		btnXemTruoc = new Button();
		btnDong = new Button();
		pnButton = new Panel();
		panel1 = new Panel();
		label3 = new Label();
		((Control)pnPrintOption).SuspendLayout();
		((Control)pnButton).SuspendLayout();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((ListControl)cbbChonMayIn).FormattingEnabled = true;
		((Control)cbbChonMayIn).Location = new Point(71, 4);
		((Control)cbbChonMayIn).Name = "cbbChonMayIn";
		((Control)cbbChonMayIn).Size = new Size(289, 21);
		((Control)cbbChonMayIn).TabIndex = 0;
		((Control)lblChonMayIn).AutoSize = true;
		((Control)lblChonMayIn).Location = new Point(3, 7);
		((Control)lblChonMayIn).Name = "lblChonMayIn";
		((Control)lblChonMayIn).Size = new Size(65, 13);
		((Control)lblChonMayIn).TabIndex = 1;
		((Control)lblChonMayIn).Text = "Chọn máy in";
		((Control)lblTrangIn).AutoSize = true;
		((Control)lblTrangIn).Location = new Point(3, 56);
		((Control)lblTrangIn).Name = "lblTrangIn";
		((Control)lblTrangIn).Size = new Size(46, 13);
		((Control)lblTrangIn).TabIndex = 2;
		((Control)lblTrangIn).Text = "Trang in";
		((Control)lblSoBanIn).AutoSize = true;
		((Control)lblSoBanIn).Location = new Point(3, 31);
		((Control)lblSoBanIn).Name = "lblSoBanIn";
		((Control)lblSoBanIn).Size = new Size(52, 13);
		((Control)lblSoBanIn).TabIndex = 3;
		((Control)lblSoBanIn).Text = "Số bản in";
		((Control)txtSoBanIn).Location = new Point(71, 29);
		((Control)txtSoBanIn).Name = "txtSoBanIn";
		((Control)txtSoBanIn).Size = new Size(289, 20);
		((Control)txtSoBanIn).TabIndex = 4;
		((Control)txtSoBanIn).TextChanged += txtSoBanIn_TextChanged;
		((Control)txtTrangIn).Location = new Point(72, 53);
		((Control)txtTrangIn).Name = "txtTrangIn";
		((Control)txtTrangIn).Size = new Size(288, 20);
		((Control)txtTrangIn).TabIndex = 5;
		pnPrintOption.BorderStyle = (BorderStyle)1;
		((Control)pnPrintOption).Controls.Add((Control)(object)ckInThongTinDieuChinh);
		((Control)pnPrintOption).Controls.Add((Control)(object)cbbChonMayIn);
		((Control)pnPrintOption).Controls.Add((Control)(object)txtTrangIn);
		((Control)pnPrintOption).Controls.Add((Control)(object)lblChonMayIn);
		((Control)pnPrintOption).Controls.Add((Control)(object)txtSoBanIn);
		((Control)pnPrintOption).Controls.Add((Control)(object)lblTrangIn);
		((Control)pnPrintOption).Controls.Add((Control)(object)lblSoBanIn);
		((Control)pnPrintOption).Location = new Point(6, 29);
		((Control)pnPrintOption).Name = "pnPrintOption";
		((Control)pnPrintOption).Size = new Size(369, 101);
		((Control)pnPrintOption).TabIndex = 6;
		((Control)ckInThongTinDieuChinh).AutoSize = true;
		((Control)ckInThongTinDieuChinh).Location = new Point(76, 79);
		((Control)ckInThongTinDieuChinh).Name = "ckInThongTinDieuChinh";
		((Control)ckInThongTinDieuChinh).Size = new Size(132, 17);
		((Control)ckInThongTinDieuChinh).TabIndex = 6;
		((Control)ckInThongTinDieuChinh).Text = "In thông tin điều chỉnh";
		((ButtonBase)ckInThongTinDieuChinh).UseVisualStyleBackColor = true;
		((Control)btnIn).Location = new Point(81, 4);
		((Control)btnIn).Name = "btnIn";
		((Control)btnIn).Size = new Size(80, 25);
		((Control)btnIn).TabIndex = 7;
		((Control)btnIn).Text = "&In";
		((ButtonBase)btnIn).UseVisualStyleBackColor = true;
		((Control)btnIn).Click += btnIn_Click;
		((Control)btnXemTruoc).Location = new Point(167, 4);
		((Control)btnXemTruoc).Name = "btnXemTruoc";
		((Control)btnXemTruoc).Size = new Size(80, 25);
		((Control)btnXemTruoc).TabIndex = 8;
		((Control)btnXemTruoc).Text = "Xem t&rước";
		((ButtonBase)btnXemTruoc).UseVisualStyleBackColor = true;
		((Control)btnXemTruoc).Click += btnXemTruoc_Click;
		((Control)btnDong).Location = new Point(252, 4);
		((Control)btnDong).Name = "btnDong";
		((Control)btnDong).Size = new Size(80, 25);
		((Control)btnDong).TabIndex = 9;
		((Control)btnDong).Text = "Đó&ng";
		((ButtonBase)btnDong).UseVisualStyleBackColor = true;
		((Control)btnDong).Click += btnDong_Click;
		((Control)pnButton).Controls.Add((Control)(object)btnIn);
		((Control)pnButton).Controls.Add((Control)(object)btnDong);
		((Control)pnButton).Controls.Add((Control)(object)btnXemTruoc);
		((Control)pnButton).Location = new Point(35, 133);
		((Control)pnButton).Name = "pnButton";
		((Control)pnButton).Size = new Size(342, 34);
		((Control)pnButton).TabIndex = 10;
		((Control)panel1).BackgroundImage = (Image)componentResourceManager.GetObject("panel1.BackgroundImage");
		((Control)panel1).BackgroundImageLayout = (ImageLayout)3;
		((Control)panel1).Controls.Add((Control)(object)label3);
		((Control)panel1).Dock = (DockStyle)1;
		((Control)panel1).Location = new Point(0, 0);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(381, 21);
		((Control)panel1).TabIndex = 26;
		((Control)label3).AutoSize = true;
		((Control)label3).BackColor = Color.Transparent;
		((Control)label3).Font = new Font("Tahoma", 8.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label3).ForeColor = Color.White;
		((Control)label3).Location = new Point(1, 3);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(61, 13);
		((Control)label3).TabIndex = 0;
		((Control)label3).Text = "In tờ khai";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(381, 172);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)pnButton);
		((Control)this).Controls.Add((Control)(object)pnPrintOption);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmHTKKPrintOption";
		((Form)this).StartPosition = (FormStartPosition)4;
		((Control)this).Text = "In tờ khai";
		((Form)this).Load += frmHTKKPrintOption_Load;
		((Form)this).FormClosed += new FormClosedEventHandler(frmHTKKPrintOption_FormClosed);
		((Control)pnPrintOption).ResumeLayout(false);
		((Control)pnPrintOption).PerformLayout();
		((Control)pnButton).ResumeLayout(false);
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
