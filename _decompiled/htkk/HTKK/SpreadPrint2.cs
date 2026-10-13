using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using CoolPrintPreview;
using FarPoint.Win.Spread;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;
using TBARCODE5Lib;

namespace HTKK;

public class SpreadPrint2
{
	public PrintDocument printDoc;

	private FpSpread winFpSpread;

	private int currentSheet;

	private int currentPageOfSheet;

	private Dictionary<int, int> dicSheet = new Dictionary<int, int>();

	private int totalPage;

	private int currentPage;

	private string[] arrBarcode;

	private TBarCode5Class barcode5;

	private int currentIdxBarcode;

	private int numBarcodeEachPage;

	private bool isPrint;

	private List<int> lstPagePrint = new List<int>();

	public static int LandscapeWidth { get; set; }

	public static int LandscapeHeight { get; set; }

	public static int PortraitWidth { get; set; }

	public static int PortraitHeight { get; set; }

	public static Rectangle rectPortrait => new Rectangle(30, 85, PortraitWidth - 60, PortraitHeight - 240);

	public static Rectangle rectLandscape => new Rectangle(30, 85, LandscapeWidth - 60, LandscapeHeight - (int)Math.Round(120.0, 0));

	public SpreadPrint2(FpSpread _winFpSpread, PrintDocument _printDoc, string _printerName, string[] _arrBarcode, int _totalPage, List<int> _lstPagePrint, bool _isPrint)
	{
		try
		{
			totalPage = _totalPage;
			currentPage = 1;
			printDoc = _printDoc;
			printDoc.PrinterSettings.PrinterName = _printerName;
			PortraitWidth = (LandscapeHeight = printDoc.PrinterSettings.DefaultPageSettings.Bounds.Width);
			PortraitHeight = (LandscapeWidth = printDoc.PrinterSettings.DefaultPageSettings.Bounds.Height);
			List<int> list = new List<int>();
			for (int num = _winFpSpread.Sheets.Count - 1; num >= 0; num--)
			{
				if (((num >= WinSettings.ToKhaiInfo.SheetInfos.Count || WinSettings.ToKhaiInfo.SheetInfos[num] == null || WinSettings.ToKhaiInfo.SheetInfos[num].IsEmptySheet) && _winFpSpread.Sheets[num].SheetName.ToLower() != "header") || !_winFpSpread.Sheets[num].Visible)
				{
					_winFpSpread.Sheets.RemoveAt(num);
				}
				else
				{
					_winFpSpread.Sheets[num].PrintInfo.Margin = new PrintMargin(printDoc.PrinterSettings.DefaultPageSettings.Margins.Left, printDoc.PrinterSettings.DefaultPageSettings.Margins.Top, printDoc.PrinterSettings.DefaultPageSettings.Margins.Right, printDoc.PrinterSettings.DefaultPageSettings.Margins.Bottom, _winFpSpread.Sheets[num].PrintInfo.Margin.Header, _winFpSpread.Sheets[num].PrintInfo.Margin.Footer);
					list.Add(num);
				}
			}
			for (int num = 0; num < list.Count; num++)
			{
				dicSheet.Add(num, list[list.Count - 1 - num]);
			}
			winFpSpread = _winFpSpread;
			currentSheet = 0;
			currentPageOfSheet = 1;
			arrBarcode = _arrBarcode;
			if (arrBarcode != null)
			{
				for (int num = 0; num < arrBarcode.Length; num++)
				{
					HTKKFile.LogRun("arrBarcode", arrBarcode[num]);
				}
			}
			currentPage = 1;
			lstPagePrint = _lstPagePrint;
			isPrint = _isPrint;
			try
			{
				barcode5 = new TBarCode5Class();
				barcode5.BarCode = tag_BarCType.eBC_PDF417;
				barcode5.LicenseMe("HCMTAX", tag_licKind.eLicKindDeveloper, 10000, "B48994B2", tag_licProduct.eLicProd2D);
			}
			catch (COMException)
			{
				HTKKFile.RegisterDll("TBarCode5.ocx");
				barcode5 = new TBarCode5Class();
				barcode5.BarCode = tag_BarCType.eBC_PDF417;
				barcode5.LicenseMe("HCMTAX", tag_licKind.eLicKindDeveloper, 10000, "B48994B2", tag_licProduct.eLicProd2D);
			}
			currentIdxBarcode = 0;
			if (arrBarcode == null)
			{
				numBarcodeEachPage = 0;
			}
			else
			{
				numBarcodeEachPage = ((arrBarcode.Length % _totalPage == 0) ? (arrBarcode.Length / _totalPage) : (arrBarcode.Length / _totalPage + 1));
			}
		}
		catch (Exception ex2)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex2);
		}
	}

	public void PrintPreview(frmHTKKPrintOption parent)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			printDoc.QueryPageSettings += new QueryPageSettingsEventHandler(printDoc_QueryPageSettings);
			printDoc.PrintPage += new PrintPageEventHandler(printDoc_PrintPage);
			if (isPrint)
			{
				printDoc.Print();
				return;
			}
			CoolPrintPreviewDialog coolPrintPreviewDialog = new CoolPrintPreviewDialog();
			coolPrintPreviewDialog.Document = printDoc;
			((Form)coolPrintPreviewDialog).WindowState = (FormWindowState)2;
			((Form)coolPrintPreviewDialog).ShowDialog((IWin32Window)(object)parent);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void printDoc_QueryPageSettings(object sender, QueryPageSettingsEventArgs e)
	{
		try
		{
			if (winFpSpread.Sheets[currentSheet].PrintInfo.Orientation == PrintOrientation.Landscape)
			{
				e.PageSettings.Landscape = true;
			}
			else
			{
				e.PageSettings.Landscape = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void printDoc_PrintPage(object sender, PrintPageEventArgs e)
	{
		//IL_0232: Unknown result type (might be due to invalid IL or missing references)
		//IL_0267: Expected O, but got Unknown
		//IL_05d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_0605: Expected O, but got Unknown
		//IL_0562: Unknown result type (might be due to invalid IL or missing references)
		//IL_0596: Expected O, but got Unknown
		try
		{
			int num = 20;
			int num2 = 16;
			Rectangle rect = (e.PageSettings.Landscape ? rectLandscape : rectPortrait);
			int ownerPrintPageCount = winFpSpread.GetOwnerPrintPageCount(e.Graphics, rect, currentSheet);
			if (lstPagePrint == null || lstPagePrint.Count == 0 || (lstPagePrint.Count > 0 && lstPagePrint.Contains(currentPage)))
			{
				int num3 = 3500;
				int num4 = 150;
				int num5 = 1000;
				if (e.PageSettings.Landscape)
				{
					num3 = 5000;
				}
				string text = string.Empty;
				for (int i = 0; i < numBarcodeEachPage; i++)
				{
					if ("~L1~L2~L3~L4~L5~L6~L7~L8~L9~T2~T3~T4~T5~T6~T7~T8~T9~H2~H4~D1~D2~D3~D4~D5~I1~I2~I3~I4~I7~I9~J1~J5~J6~J7~J8~J9~M8~M9~K2~K3~K4~K7~G5~G6~G7~G8~G9~99~N2~N8~B8~B9~C6~C7~I8~C8~C9~D6~J2~J3~J4~D7~D8~D9~E1~E2~E3~E9~E4~E5~E6~E7~E8~K5~N3~N4~N5~N6~N7~03~96~97~94~79~90~01~72~H5~H6~O1~O2~O3~O4~O5~O6~O7~B2~M6~P1~P5~P6~P7~Q4~Q5~Q6~Q7~P8~P9~Q1~Q2~Q3~Q8~Q9~R4~R9~R6~U2~U1~U3~U4~U5~U6~U7~U8~U9~X1~X2~X3~X4~X7~X8~X9~Y1~Y6~Y3~Y4~Z1~Z3~Y5~Z4~Z8~Z5~W1~Y2~Y8~Y7~".IndexOf("~" + WinSettings.TTPeriodInfo.menuID + "~") <= -1 && (currentSheet <= 0 || !(WinSettings.TTPeriodInfo.menuID == "33")))
					{
						text += ((currentIdxBarcode < arrBarcode.Length) ? arrBarcode[currentIdxBarcode] : string.Empty);
						int nLeft = ((i == 0) ? (num3 - (numBarcodeEachPage - 1 - i) * num5) : (num3 - (numBarcodeEachPage - 1 - i) * num5 + num4 * i));
						barcode5.Text = ((currentIdxBarcode < arrBarcode.Length) ? arrBarcode[currentIdxBarcode] : string.Empty);
						IntPtr hdc = e.Graphics.GetHdc();
						barcode5.BCDraw(hdc.ToInt32(), nLeft, 50, num5, 420);
						e.Graphics.ReleaseHdc();
						currentIdxBarcode++;
					}
				}
				e.Graphics.DrawString(text, new Font("Times New Roman", 3f, (FontStyle)0), Brushes.White, new RectangleF(1f, 1f, 28f, e.PageBounds.Height - 2));
				if (WinSettings.MenuID != "70")
				{
					string[] mucLucNganSach = WinDoWorkEvent.getMucLucNganSach();
					PrintInfo printInfo = winFpSpread.Sheets[0].PrintInfo;
					if (mucLucNganSach != null && mucLucNganSach.Length > 0)
					{
						string text2 = mucLucNganSach[0];
						string text3 = "/fz\"8\"/fn\"Times New Roman\"/fb1/fi1/fu1";
						text2 = text3 + text2 + "/n/fz\"8\"/fn\"Times New Roman\"/fb0/fi0/fu0";
						if (mucLucNganSach != null && mucLucNganSach.Length > 0)
						{
							for (int i = 1; i < mucLucNganSach.Length; i++)
							{
								text2 = ((i != mucLucNganSach.Length) ? (text2 + mucLucNganSach[i] + "/n/fz\"8\"/fn\"Times New Roman\"/fb0/fi0/fu0") : (text2 + mucLucNganSach[i]));
							}
						}
						printInfo.Footer = text2;
					}
					string kieuKHBS = WinSettings.ToKhaiInfo.KieuKHBS;
					if (currentSheet == 0 && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && (kieuKHBS.Equals("1") || kieuKHBS.Equals("2")))
					{
						bool flag = false;
						for (int num6 = winFpSpread.Sheets.Count - 1; num6 >= 0; num6--)
						{
							if (winFpSpread.Sheets[num6].SheetName.ToUpper() == "KHBS")
							{
								flag = true;
								break;
							}
						}
						if (!flag)
						{
							printInfo.Footer = printInfo.Footer + "/n/fz\"8\"/fn\"Times New Roman\"/fb0/fi0/fu0" + HTKKFile.GetMessage("0226");
						}
					}
					printInfo.ShowBorder = false;
					printInfo.Margin = new PrintMargin(printDoc.PrinterSettings.DefaultPageSettings.Margins.Left, printDoc.PrinterSettings.DefaultPageSettings.Margins.Top, printDoc.PrinterSettings.DefaultPageSettings.Margins.Right, printDoc.PrinterSettings.DefaultPageSettings.Margins.Bottom, printInfo.Margin.Header, printInfo.Margin.Footer);
					printInfo.Centering = Centering.Horizontal;
					winFpSpread.SetPrintInfo(printInfo, 0);
				}
				winFpSpread.OwnerPrintDraw(e.Graphics, rect, currentSheet, currentPageOfSheet);
				int num7 = 100;
				if (e.PageSettings.Landscape)
				{
					num7 = 80;
					e.Graphics.DrawString("Trang " + currentPage + "/" + totalPage, new Font("Times New Roman", 8f, (FontStyle)0), Brushes.Black, new RectangleF(rect.Width - 100, rect.Height + num7, 100f, 10f));
				}
				else
				{
					e.Graphics.DrawString("Trang " + currentPage + "/" + totalPage, new Font("Times New Roman", 8f, (FontStyle)0), Brushes.Black, new RectangleF(rect.Width - 50, rect.Height + num7, 100f, 10f));
				}
				if (currentPageOfSheet == ownerPrintPageCount)
				{
					currentSheet++;
					currentPageOfSheet = 1;
				}
				else
				{
					currentPageOfSheet++;
				}
				currentPage++;
				if (currentSheet >= winFpSpread.Sheets.Count - 1)
				{
					e.HasMorePages = false;
				}
				else
				{
					e.HasMorePages = true;
				}
			}
			else
			{
				if (currentPageOfSheet == ownerPrintPageCount)
				{
					currentSheet++;
					currentPageOfSheet = 1;
				}
				else
				{
					currentPageOfSheet++;
				}
				currentPage++;
				if (currentSheet >= winFpSpread.Sheets.Count - 1)
				{
					e.HasMorePages = false;
				}
				else
				{
					e.HasMorePages = true;
				}
				e.Cancel = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void inThongTinDL(FpSpread reportSpread)
	{
		try
		{
			string text = "";
			string text2 = "#B3#B4#B5#A2#A5#A6#A7#C1#C2#C3#C4#C5#";
			int num = -1;
			if (HTKKFormater.isNullString(WinSettings.HTKKHeaderInfo.mstDLyThue))
			{
				for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
				{
					num++;
					if (WinSettings.ToKhaiInfo.SheetInfos[i] == null || (!WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu && i != 0 && !text2.Contains("#" + WinSettings.MenuID + "#")))
					{
						continue;
					}
					text = WinSettings.ToKhaiInfo.SheetInfos[i].RowTTDLT;
					if (!HTKKFormater.isNullString(text) && !text.Equals("_"))
					{
						string[] array = text.Split(new char[1] { '_' });
						for (int j = int.Parse(array[0]) - 1; j < int.Parse(array[1]); j++)
						{
							reportSpread.Sheets[num].SetRowVisible(j, visible: false);
						}
					}
				}
			}
			if (WinSettings.MenuID == "A4" && reportSpread.Sheets[0].Visible)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/maHoSoKhaiDuAnDauTu"));
				string[] array2 = "28_39".Split(new char[1] { '_' });
				if (!hTKKCMCCellInfo.IsNullOrEmpty() && hTKKCMCCellInfo.Value.IsNullOrEmpty())
				{
					reportSpread.Sheets[0].Rows[ConvertHelpers.ConvertToInt(array2[0]) - 1, ConvertHelpers.ConvertToInt(array2[1]) - 1].Visible = false;
				}
			}
			int num2;
			if ((WinSettings.MenuID == "I4" || WinSettings.MenuID == "I7") && reportSpread.Sheets[0].Visible)
			{
				num2 = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam);
				int num3 = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang);
				if (num2 < 2023 || num2 > 2026 || (num2 == 2023 && num3 < 12))
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[2];
					int num4 = hTKKCMCTableInfo.Location.DataLocationReportRowStart - 1;
					reportSpread.Sheets[0].Rows[num4 - 2].Visible = false;
					for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
					{
						reportSpread.Sheets[0].Rows[num4 + i].Visible = false;
					}
				}
			}
			if (!(WinSettings.MenuID == "C7") || !reportSpread.Sheets[0].Visible)
			{
				return;
			}
			num2 = ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam);
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy");
			DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy");
			DateTime dateTime3 = ConvertHelpers.ConvertToDateTime("01/12/2023", "dd/MM/yyyy");
			DateTime dateTime4 = ConvertHelpers.ConvertToDateTime("31/12/2026", "dd/MM/yyyy");
			if ((!(dateTime >= dateTime3) || !(dateTime <= dateTime4)) && (!(dateTime2 >= dateTime3) || !(dateTime2 <= dateTime4)) && (!(dateTime < dateTime3) || !(dateTime2 > dateTime4)))
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[2];
				int num4 = hTKKCMCTableInfo.Location.DataLocationReportRowStart - 1;
				reportSpread.Sheets[0].Rows[num4 - 2].Visible = false;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					reportSpread.Sheets[0].Rows[num4 + i].Visible = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void SetupPageBreak(FpSpread reportSpread)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Expected O, but got Unknown
		try
		{
			int[] rowPageBreaks = reportSpread.GetRowPageBreaks(0, reCalc: true);
			List<int> list = new List<int>();
			int num = -1;
			int num2 = 0;
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile("ConfigPrinter.xml"));
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Root/Menu[@ID='" + WinSettings.MenuID + "']");
			if (val2.Count > 0 && val2[0].HasChildNodes)
			{
				foreach (XmlNode childNode in val2[0].ChildNodes)
				{
					XmlNode val3 = childNode;
					int num3 = ((val3.Attributes["Index"] != null) ? ConvertHelpers.ConvertToInt(((XmlNode)val3.Attributes["Index"]).Value) : 0);
					string pathStaticCell = ((val3.Attributes["PathStaticCell"] == null) ? "" : ((XmlNode)val3.Attributes["PathStaticCell"]).Value);
					int num4 = ((val3.Attributes["RowTren"] != null) ? ConvertHelpers.ConvertToInt(((XmlNode)val3.Attributes["RowTren"]).Value) : 0);
					int num5 = ((val3.Attributes["RowDuoi"] != null) ? ConvertHelpers.ConvertToInt(((XmlNode)val3.Attributes["RowDuoi"]).Value) : 0);
					int num6 = ((val3.Attributes["RowBreak"] != null) ? ConvertHelpers.ConvertToInt(((XmlNode)val3.Attributes["RowBreak"]).Value) : 0);
					rowPageBreaks = reportSpread.GetRowPageBreaks(num3, reCalc: true);
					if (rowPageBreaks.Length > 0 && WinSettings.ToKhaiInfo.SheetInfos[num3] != null)
					{
						int num7 = rowPageBreaks[^1] + 1;
						HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[num3].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals(pathStaticCell));
						if (hTKKCMCCellInfo != null && num7 >= hTKKCMCCellInfo.CellReport.RowIndex + num4 && num7 <= hTKKCMCCellInfo.CellReport.RowIndex + num5)
						{
							num = hTKKCMCCellInfo.CellReport.RowIndex + num6;
							reportSpread.Sheets[num3].SetRowPageBreak(num - 1, value: true);
						}
					}
				}
			}
			int num8 = -1;
			for (int num9 = 0; num9 < WinSettings.ToKhaiInfo.SheetInfos.Count; num9++)
			{
				num8++;
				if (WinSettings.ToKhaiInfo.SheetInfos[num9] == null || !reportSpread.Sheets[num8].Visible)
				{
					continue;
				}
				if (reportSpread.Sheets[num8].SheetName.ToLower() == "khbs")
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[num9].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("PL01_KHBS/NOIDUNGGIAITHICH/lyDoKhac"));
					int rowIndex = hTKKCMCCellInfo.CellReport.RowIndex;
					reportSpread.Sheets[num8].SetRowPageBreak(rowIndex - 7, value: true);
				}
				if (WinSettings.MenuID.Equals("57"))
				{
					reportSpread.Sheets[num8].SetRowPageBreak(64, value: true);
				}
				if (WinSettings.MenuID.Equals("58"))
				{
					if (WinSettings.TTPeriodInfo.DoanhNghiepHoatDongLienTuc == "LT")
					{
						reportSpread.Sheets[0].SetRowPageBreak(171, value: true);
					}
					else
					{
						reportSpread.Sheets[0].SetRowPageBreak(251, value: true);
					}
				}
			}
			int[] array;
			string text;
			if (WinSettings.MenuID.Equals("D1") || WinSettings.MenuID.Equals("Z5"))
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[6] != null)
				{
					array = GeneralFunctions.ArrayCellId("B_1");
					text = "";
					if (WinSettings.TTPeriodInfo.KieuKy == "M")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 6 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
						}
					}
					else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && WinSettings.TTPeriodInfo.Quy != "1")
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
						}
					}
					reportSpread.Sheets[6].Cells[array[1], array[0]].Text = text;
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[7] != null)
				{
					array = GeneralFunctions.ArrayCellId("B_4");
					text = "";
					if (WinSettings.TTPeriodInfo.KieuKy == "M")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2024"))
						{
							text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2025") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
						{
							text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
						}
						else if ((WinSettings.TTPeriodInfo.Nam.Equals("2025") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) > 6) || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
						{
							text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
						}
					}
					else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2024"))
						{
							text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2025") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 2)
						{
							text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
						}
						else if ((WinSettings.TTPeriodInfo.Nam.Equals("2025") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) > 2) || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
						{
							text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
						}
					}
					reportSpread.Sheets[7].Cells[array[1], array[0]].Text = text;
				}
			}
			if (WinSettings.MenuID.Equals("D4") || WinSettings.MenuID.Equals("Z3"))
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[2] != null)
				{
					array = GeneralFunctions.ArrayCellId("B_1");
					text = "";
					if (WinSettings.TTPeriodInfo.ThuHo != "TH")
					{
						if (WinSettings.TTPeriodInfo.KieuKy == "M")
						{
							if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 6 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
							{
								text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
							}
							else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
							{
								text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "D")
						{
							if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
							{
								text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
							}
							else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
							{
								text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
							}
						}
						else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
						{
							if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && WinSettings.TTPeriodInfo.Quy != "1")
							{
								text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
							}
							else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))
							{
								text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
							}
						}
					}
					else if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
						}
					}
					else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && WinSettings.TTPeriodInfo.Quy != "1" && WinSettings.TTPeriodInfo.Quy != "2")
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
						}
					}
					reportSpread.Sheets[2].Cells[array[1], array[0]].Text = text;
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[3] != null)
				{
					array = GeneralFunctions.ArrayCellId("B_4");
					text = "";
					if (WinSettings.TTPeriodInfo.Nam.Equals("2024"))
					{
						text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
					}
					else if (WinSettings.TTPeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 2 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Quy))))
					{
						text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
					}
					else if ((WinSettings.TTPeriodInfo.Nam.Equals("2025") && (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) > 6 || ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) > 2)) || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
					{
						text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
					}
					reportSpread.Sheets[3].Cells[array[1], array[0]].Text = text;
				}
			}
			if (WinSettings.MenuID.Equals("G5"))
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[4] != null)
				{
					array = GeneralFunctions.ArrayCellId("B_1");
					text = "";
					if (WinSettings.TTPeriodInfo.KieuKy == "D" || WinSettings.TTPeriodInfo.KieuKy == "M")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
						}
					}
					else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
					{
						if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4"))
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
						}
						else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))
						{
							text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
						}
					}
					reportSpread.Sheets[4].Cells[array[1], array[0]].Text = text;
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[5] != null)
				{
					array = GeneralFunctions.ArrayCellId("B_4");
					text = "";
					if (WinSettings.TTPeriodInfo.Nam.Equals("2024"))
					{
						text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
					}
					else if (WinSettings.TTPeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 2 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Quy))))
					{
						text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
					}
					else if ((WinSettings.TTPeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) > 6 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) > 2 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Quy)))) || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
					{
						text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
					}
					reportSpread.Sheets[5].Cells[array[1], array[0]].Text = text;
				}
			}
			if (!WinSettings.MenuID.Equals("G7"))
			{
				return;
			}
			if (WinSettings.ToKhaiInfo.SheetInfos[4] != null)
			{
				array = GeneralFunctions.ArrayCellId("B_1");
				text = "";
				if (WinSettings.TTPeriodInfo.KieuKy == "D_D")
				{
					DateTime dateTime = new DateTime(2023, 7, 1);
					DateTime dateTime2 = new DateTime(2023, 12, 31);
					DateTime dateTime3 = new DateTime(2024, 6, 30);
					DateTime dateTime4 = new DateTime(2024, 1, 1);
					DateTime dateTime5 = DateTime.Now;
					DateTime dateTime6 = DateTime.Now;
					if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) && WinSettings.TTPeriodInfo.TuNgay.Length > 8)
					{
						dateTime5 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy");
					}
					if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.DenNgay) && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
					{
						dateTime6 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy");
					}
					if (((DateTime.Compare(dateTime5, dateTime3) <= 0 && DateTime.Compare(dateTime, dateTime5) <= 0) || (DateTime.Compare(dateTime6, dateTime3) <= 0 && DateTime.Compare(dateTime, dateTime6) <= 0) || (DateTime.Compare(dateTime5, dateTime) <= 0 && DateTime.Compare(dateTime3, dateTime6) <= 0)) && WinSettings.TTPeriodInfo.TuNgay.Length > 8 && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
					{
						text = ((DateTime.Compare(dateTime6, dateTime3) <= 0 && DateTime.Compare(dateTime4, dateTime6) <= 0) ? ((DateTime.Compare(dateTime5, dateTime3) > 0 || DateTime.Compare(dateTime4, dateTime5) > 0) ? "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15  VÀ NGHỊ QUYẾT SỐ 110/2023/QH15" : "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15") : ((DateTime.Compare(dateTime5, dateTime4) < 0 && DateTime.Compare(dateTime3, dateTime6) < 0) ? "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15  VÀ NGHỊ QUYẾT SỐ 110/2023/QH15" : ((DateTime.Compare(dateTime5, dateTime4) < 0) ? "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15" : "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15")));
					}
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "M")
				{
					if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 7 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 12)
					{
						text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
					}
					else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) >= 1 && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6)
					{
						text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
					}
				}
				else if (WinSettings.TTPeriodInfo.KieuKy == "Q")
				{
					if (WinSettings.TTPeriodInfo.Nam.Equals("2023") && (WinSettings.TTPeriodInfo.Quy == "3" || WinSettings.TTPeriodInfo.Quy == "4"))
					{
						text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 101/2023/QH15";
					}
					else if (WinSettings.TTPeriodInfo.Nam.Equals("2024") && (WinSettings.TTPeriodInfo.Quy == "1" || WinSettings.TTPeriodInfo.Quy == "2"))
					{
						text = "Phụ lục \n GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 110/2023/QH15";
					}
				}
				reportSpread.Sheets[4].Cells[array[1], array[0]].Text = text;
			}
			if (WinSettings.ToKhaiInfo.SheetInfos[5] == null)
			{
				return;
			}
			array = GeneralFunctions.ArrayCellId("B_4");
			text = "";
			if (WinSettings.TTPeriodInfo.KieuKy == "D_D")
			{
				DateTime dateTime7 = new DateTime(2024, 7, 1);
				DateTime dateTime8 = new DateTime(2024, 12, 31);
				DateTime dateTime9 = new DateTime(2025, 6, 30);
				DateTime dateTime10 = new DateTime(2025, 1, 1);
				DateTime dateTime5 = DateTime.Now;
				DateTime dateTime6 = DateTime.Now;
				if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) && WinSettings.TTPeriodInfo.TuNgay.Length > 8)
				{
					dateTime5 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy");
				}
				if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.DenNgay) && WinSettings.TTPeriodInfo.DenNgay.Length > 8)
				{
					dateTime6 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy");
				}
				text = GeneralFunctions.GetTextByDateRange(dateTime5, dateTime6);
			}
			if (WinSettings.TTPeriodInfo.KieuKy == "M" || WinSettings.TTPeriodInfo.KieuKy == "Q")
			{
				if (WinSettings.TTPeriodInfo.Nam.Equals("2024"))
				{
					text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 142/2024/QH15";
				}
				else if (WinSettings.TTPeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) <= 6 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) <= 2 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Quy))))
				{
					text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 174/2024/QH15";
				}
				else if ((WinSettings.TTPeriodInfo.Nam.Equals("2025") && ((ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) > 6 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Thang)) || (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) > 2 && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Quy)))) || WinSettings.TTPeriodInfo.Nam.Equals("2026"))
				{
					text = "GIẢM THUẾ GIÁ TRỊ GIA TĂNG THEO NGHỊ QUYẾT SỐ 204/2025/QH15";
				}
			}
			reportSpread.Sheets[5].Cells[array[1], array[0]].Text = text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void setMlNs(FpSpread winFpSpreadReport, PrintDocument printDoc)
	{
		try
		{
			string[] mucLucNganSach = WinDoWorkEvent.getMucLucNganSach();
			if (mucLucNganSach == null)
			{
				return;
			}
			for (int i = 0; i < winFpSpreadReport.Sheets.Count; i++)
			{
				if (i >= WinSettings.ToKhaiInfo.SheetInfos.Count || WinSettings.ToKhaiInfo.SheetInfos[i] == null || !WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName))
				{
					continue;
				}
				PrintInfo printInfo = winFpSpreadReport.Sheets[i].PrintInfo;
				if (mucLucNganSach != null && mucLucNganSach.Length > 0)
				{
					string text = mucLucNganSach[0];
					string text2 = "/fz\"8\"/fn\"Times New Roman\"/fb1/fi1/fu1";
					text = text2 + text + "/n/fz\"8\"/fn\"Times New Roman\"/fb0/fi0/fu0";
					if (mucLucNganSach != null && mucLucNganSach.Length > 0)
					{
						for (int j = 1; j < mucLucNganSach.Length; j++)
						{
							text = ((j != mucLucNganSach.Length) ? (text + mucLucNganSach[j] + "/n/fz\"8\"/fn\"Times New Roman\"/fb0/fi0/fu0") : (text + mucLucNganSach[j]));
						}
					}
					printInfo.Footer = text;
				}
				string kieuKHBS = WinSettings.ToKhaiInfo.KieuKHBS;
				if (i == 0 && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && (kieuKHBS.Equals("1") || kieuKHBS.Equals("2")))
				{
					bool flag = false;
					for (int num = winFpSpreadReport.Sheets.Count - 1; num >= 0; num--)
					{
						if (winFpSpreadReport.Sheets[num].SheetName.ToUpper() == "KHBS")
						{
							flag = true;
							break;
						}
					}
					if (!flag)
					{
						printInfo.Footer = printInfo.Footer + "/n/fz\"8\"/fn\"Times New Roman\"/fb0/fi0/fu0" + HTKKFile.GetMessage("0226");
					}
				}
				printInfo.ShowBorder = false;
				printInfo.Margin = new PrintMargin(printDoc.PrinterSettings.DefaultPageSettings.Margins.Left, printDoc.PrinterSettings.DefaultPageSettings.Margins.Top, printDoc.PrinterSettings.DefaultPageSettings.Margins.Right, printDoc.PrinterSettings.DefaultPageSettings.Margins.Bottom, printInfo.Margin.Header, printInfo.Margin.Footer);
				printInfo.Centering = Centering.Horizontal;
				winFpSpreadReport.SetPrintInfo(printInfo, i);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
