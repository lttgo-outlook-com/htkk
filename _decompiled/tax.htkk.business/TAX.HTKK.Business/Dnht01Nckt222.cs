using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

internal class Dnht01Nckt222
{
	private string cellId_SoDeNghiHoan = "H_6";

	private string cellId_HTTruocKTSau = "S_8";

	private string cellId_KTTruocHTSau = "S_10";

	private string cellId_LoaiHT = "L_12";

	private string cellId_KeKhai_Thang = "E_17";

	private string cellId_KeKhai_Quy = "K_17";

	private string cellId_KeKhai_Tu = "G_19";

	private string cellId_KeKhai_Tu_Nam = "J_19";

	private string cellId_KeKhai_Den = "G_21";

	private string cellId_KeKhai_Den_Nam = "J_21";

	private string cellId_SoTienDeNghiHoan = "AA_70";

	private string cellId_TongTienNSNN_BangChu = "L_84";

	private string cellId_TongTienHoanTra_BangChu = "Z_86";

	private ComboBoxCellType comboQuy = new ComboBoxCellType();

	private ComboBoxCellType comboThang = new ComboBoxCellType();

	private int[] CellId_SoDeNghiHoan => GeneralFunctions.ArrayCellId(cellId_SoDeNghiHoan);

	private int[] CellId_HTTruocKTSau => GeneralFunctions.ArrayCellId(cellId_HTTruocKTSau);

	private int[] CellId_KTTruocHTSau => GeneralFunctions.ArrayCellId(cellId_KTTruocHTSau);

	private int[] CellId_LoaiHT => GeneralFunctions.ArrayCellId(cellId_LoaiHT);

	private int[] CellId_KeKhai_Thang => GeneralFunctions.ArrayCellId(cellId_KeKhai_Thang);

	private int[] CellId_KeKhai_Quy => GeneralFunctions.ArrayCellId(cellId_KeKhai_Quy);

	private int[] CellId_KeKhai_Tu => GeneralFunctions.ArrayCellId(cellId_KeKhai_Tu);

	private int[] CellId_KeKhai_Tu_Nam => GeneralFunctions.ArrayCellId(cellId_KeKhai_Tu_Nam);

	private int[] CellId_KeKhai_Den => GeneralFunctions.ArrayCellId(cellId_KeKhai_Den);

	private int[] CellId_KeKhai_Den_Nam => GeneralFunctions.ArrayCellId(cellId_KeKhai_Den_Nam);

	private int[] CellId_SoTienDeNghiHoan => GeneralFunctions.ArrayCellId(cellId_SoTienDeNghiHoan);

	private int[] CellId_TongTienNSNN_BangChu => GeneralFunctions.ArrayCellId(cellId_TongTienNSNN_BangChu);

	private int[] CellId_TongTienHoanTra_BangChu => GeneralFunctions.ArrayCellId(cellId_TongTienHoanTra_BangChu);

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[16, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = PeriodInfo.KieuKy;
			sheetView.Cells[16, HTKKFile.ColumnLettersToIndex("AB") - 1].Text = PeriodInfo.kyKKhai;
			sheetView.Cells[98, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = (PeriodInfo.KieuKy.Equals("M") ? "0" : "1");
			sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("G") - 1].Text = PeriodInfo.TuThang;
			sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("J") - 1].Text = PeriodInfo.DenThang;
			string text = sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
			if (text.Equals("Hoàn xuất khẩu"))
			{
				string[] items = GetDM("..\\InterfaceIni\\Catalogue_Ly_Do_Hoan_Thue.xml", 0).ToArray();
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = items;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].CellType = comboBoxCellType;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].Text = ToKhaiInfo.SheetInfos[0].StaticCells[16].Value;
			}
			else if (text.Equals("Hoàn dự án đầu tư"))
			{
				string[] items = GetDM("..\\InterfaceIni\\Catalogue_Ly_Do_Hoan_Thue.xml", 1).ToArray();
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = items;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].CellType = comboBoxCellType;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].Text = ToKhaiInfo.SheetInfos[0].StaticCells[16].Value;
			}
			if (PeriodInfo.KieuKy == "M")
			{
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = DateTime.ParseExact("01/" + PeriodInfo.TuThang, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AB") - 1].Text = DateTime.ParseExact("01/" + PeriodInfo.TuThang, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
					.ToString("dd/MM/yyyy");
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = PeriodInfo.TuThang;
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = PeriodInfo.TuThang;
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = DateTime.ParseExact("01/" + PeriodInfo.DenThang, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AB") - 1].Text = DateTime.ParseExact("01/" + PeriodInfo.DenThang, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
					.ToString("dd/MM/yyyy");
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = PeriodInfo.DenThang;
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = PeriodInfo.DenThang;
			}
			else
			{
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(PeriodInfo.Quy), int.Parse(PeriodInfo.Nam), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AB") - 1].Text = GeneralFunctions.GetNgayCuoiQuy(int.Parse(PeriodInfo.Quy), int.Parse(PeriodInfo.Nam), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = PeriodInfo.TuThang;
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = PeriodInfo.TuThang;
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]), int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AB") - 1].Text = GeneralFunctions.GetNgayCuoiQuy(int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]), int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = PeriodInfo.DenThang;
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = PeriodInfo.DenThang;
			}
			bool flag = PeriodInfo.KieuKy == "Q";
			if (PeriodInfo.KieuKy == "M")
			{
				initComboboxThang();
				sheetView.Cells[CellId_KeKhai_Tu[1], CellId_KeKhai_Tu[0]].CellType = comboThang;
				sheetView.Cells[CellId_KeKhai_Den[1], CellId_KeKhai_Den[0]].CellType = comboThang;
			}
			else if (flag)
			{
				initComboboxQuy();
				sheetView.Cells[CellId_KeKhai_Tu[1], CellId_KeKhai_Tu[0]].CellType = comboQuy;
				sheetView.Cells[CellId_KeKhai_Den[1], CellId_KeKhai_Den[0]].CellType = comboQuy;
			}
			string nameTable = "01_DNHT_I";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("AM");
			for (int num3 = dataLocationTemplateRowStart; num3 <= num; num3++)
			{
				int num4 = HTKKFile.ColumnLettersToIndex("AH");
				string text2 = sheetView.Cells[num3 - 1, num4 - 2].Text;
				string text3 = sheetView.Cells[num3 - 1, num2 - 1].Text;
				if (!string.IsNullOrEmpty(text3))
				{
					sheetView.Cells[num3 - 1, num2 - 1].Text = text3.Replace("Q", "").Replace("/", "");
				}
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.MaskChar = '.';
				switch (text2)
				{
				case "Y":
					maskCellType.Mask = "0000";
					break;
				case "Q":
					maskCellType.Mask = "0/0000";
					break;
				case "M":
					maskCellType.Mask = "00/0000";
					break;
				case "D":
					maskCellType.Mask = "00/00/0000";
					break;
				}
				sheetView.Cells[num3 - 1, num2 - 1].ResetCellType();
				sheetView.Cells[num3 - 1, num2 - 1].CellType = maskCellType;
			}
			int[] array = GeneralFunctions.ArrayCellId("L_14");
			int[] array2 = GeneralFunctions.ArrayCellId("T_14");
			int[] array3 = GeneralFunctions.ArrayCellId("AG_14");
			sheetView.Cells[array[1], array[0]].Locked = true;
			sheetView.Cells[array2[1], array2[0]].Locked = true;
			string text4 = TTinChung.tenCQTCapCuc;
			string text5 = TTinChung.tenCQTNoiNop;
			if (text4.IndexOf("-") == 4)
			{
				text4 = text4.Substring(6).Trim();
			}
			if ("10100,70100,71300,71100".Split(new char[1] { ',' }).Contains(TTinChung.maCQTCapCuc))
			{
				sheetView.Cells[array[1], array[0]].Text = text4;
				sheetView.Cells[array2[1], array2[0]].Text = TTinChung.tenCQTNoiNop;
				sheetView.Cells[array3[1], array3[0]].Text = TTinChung.maCQTNoiNop;
			}
			else
			{
				sheetView.Cells[array[1], array[0]].Text = text4;
				sheetView.Cells[array2[1], array2[0]].Text = text4;
				text5 = text4;
				sheetView.Cells[array3[1], array3[0]].Text = TTinChung.maCQTCapCuc;
			}
			Row row = sheetView.Rows[array[1]];
			float preferredHeight = row.GetPreferredHeight();
			string text6 = text4;
			if (text4.Length < text5.Length)
			{
				text6 = text5;
			}
			if (text6.Length > 30)
			{
				float num5 = preferredHeight / row.Height;
				preferredHeight += num5 * 5f;
				row.Height = preferredHeight;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_0a6d: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("BD");
			int num5 = HTKKFile.ColumnLettersToIndex("BH");
			int num6 = HTKKFile.ColumnLettersToIndex("AH");
			int num7 = HTKKFile.ColumnLettersToIndex("AM");
			int num8 = HTKKFile.ColumnLettersToIndex("AA");
			int num9 = HTKKFile.ColumnLettersToIndex("AX");
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == num2 && iRow == 12)
			{
				foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.LoaiHoanThue)
				{
					if (sheetView.Cells[num - 1, num2 - 1].Text == item.Value)
					{
						sheetView.Cells[num - 1, num4 - 1].Text = item.Key;
						break;
					}
				}
				return;
			}
			if (iCol == num3 && iRow == 23)
			{
				foreach (KeyValuePair<string, string> item2 in HTKKWinDictionaryDiaDanh.LyDoHoanThue)
				{
					if (sheetView.Cells[num - 1, num3 - 1].Text == item2.Value)
					{
						sheetView.Cells[num - 1, num5 - 1].Text = item2.Key;
						break;
					}
				}
				return;
			}
			if (iCol == num6)
			{
				string text = sheetView.Cells[num - 1, num6 - 2].Text;
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.MaskChar = '.';
				switch (text)
				{
				case "Y":
					maskCellType.Mask = "0000";
					break;
				case "Q":
					maskCellType.Mask = "0/0000";
					break;
				case "M":
					maskCellType.Mask = "00/0000";
					break;
				case "D":
					maskCellType.Mask = "00/00/0000";
					break;
				}
				sheetView.Cells[(int)iRow - 1, num7 - 1].Text = "";
				sheetView.Cells[(int)iRow - 1, num7 - 1].ResetCellType();
				sheetView.Cells[(int)iRow - 1, num7 - 1].CellType = maskCellType;
			}
			else if (iCol == num8 || iCol == num9)
			{
				string nameTable = "01_DNHT_I";
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable);
				int count = hTKKCMCTableInfo.RowsValues.Count;
				int num10 = 82 + count - 1;
				int num11 = 86 + count - 1;
				string value = GeneralFunctions.GetValue(sheetView, "AX_" + num10);
				string text2 = MoneyHelpers.MoneyToText(value);
				winFpSpread.Sheets[0].Cells[CellId_TongTienNSNN_BangChu[1] + count - 1, CellId_TongTienNSNN_BangChu[0]].Text = text2;
				value = GeneralFunctions.GetValue(sheetView, "L_" + num11);
				if (value != null && !string.IsNullOrEmpty(value))
				{
					text2 = MoneyHelpers.MoneyToText(value);
					winFpSpread.Sheets[0].Cells[CellId_TongTienHoanTra_BangChu[1] + count - 1, CellId_TongTienHoanTra_BangChu[0]].Text = text2;
				}
			}
			else
			{
				if (iCol != num7)
				{
					return;
				}
				string text = sheetView.Cells[num - 1, num6 - 2].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, num7 - 1].Text;
				frmMessage frmMessage = new frmMessage("0456", "3", "", "Thông báo");
				switch (text)
				{
				case "M":
					if (!(ConvertHelpers.ConvertToDateTime("01/" + text3, "dd/MM/yyyy") == DateTime.MinValue))
					{
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BI") - 1].Text = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BJ") - 1].Text = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
							.ToString("dd/MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BK") - 1].Text = text3;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BL") - 1].Text = text3;
					}
					break;
				case "D":
					if (!(ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy") == DateTime.MinValue))
					{
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BI") - 1].Text = text3;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BJ") - 1].Text = text3;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BK") - 1].Text = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy").ToString("MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BL") - 1].Text = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy").ToString("MM/yyyy");
					}
					break;
				case "Y":
					if (!(ConvertHelpers.ConvertToDateTime("01/01/" + text3, "dd/MM/yyyy") == DateTime.MinValue))
					{
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BI") - 1].Text = DateTime.ParseExact("01/01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BJ") - 1].Text = DateTime.ParseExact("01/12/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
							.ToString("dd/MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BK") - 1].Text = "01/" + text3;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BL") - 1].Text = "12/" + text3;
					}
					break;
				case "Q":
					if (!(ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text3.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue))
					{
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BI") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text3.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text3.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BJ") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text3.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text3.Split(new char[1] { '/' })[1]), 1, 1).AddMonths(3).AddDays(-1.0)
							.ToString("dd/MM/yyyy");
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BK") - 1].Text = text3;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BL") - 1].Text = text3;
					}
					break;
				default:
					sheetView.Cells[(int)iRow - 1, num7 - 1].Text = "";
					((Form)frmMessage).ShowDialog();
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("AH");
			int num3 = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("AX");
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			if ((int)iCol != HTKKFile.ColumnLettersToIndex("L") || (int)iRow - 1 != 11)
			{
				return;
			}
			Dictionary<string, string> loaiHoanThue = HTKKWinDictionaryDiaDanh.LoaiHoanThue;
			string text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
			foreach (KeyValuePair<string, string> item in loaiHoanThue)
			{
				if (item.Value == text)
				{
					sheetView.Cells[GeneralFunctions.ArrayCellId("BD_12")[1], GeneralFunctions.ArrayCellId("BD_12")[0]].Text = item.Key;
					break;
				}
			}
			if (text.Equals("Hoàn xuất khẩu"))
			{
				string[] items = GetDM("..\\InterfaceIni\\Catalogue_Ly_Do_Hoan_Thue.xml", 0).ToArray();
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = items;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].CellType = comboBoxCellType;
			}
			else if (text.Equals("Hoàn dự án đầu tư"))
			{
				string[] items = GetDM("..\\InterfaceIni\\Catalogue_Ly_Do_Hoan_Thue.xml", 1).ToArray();
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = items;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].CellType = comboBoxCellType;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_HTTruocKTSau[1], CellId_HTTruocKTSau[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_KTTruocHTSau[1], CellId_KTTruocHTSau[0]].Value);
			if (idxRow == CellId_HTTruocKTSau[1] && idxCol == CellId_HTTruocKTSau[0])
			{
				if (!flag)
				{
					sheetView.Cells[CellId_KTTruocHTSau[1], CellId_KTTruocHTSau[0]].Value = true;
				}
				else
				{
					sheetView.Cells[CellId_KTTruocHTSau[1], CellId_KTTruocHTSau[0]].Value = false;
				}
			}
			else if (idxRow == CellId_KTTruocHTSau[1] && idxCol == CellId_KTTruocHTSau[0])
			{
				if (!flag2)
				{
					sheetView.Cells[CellId_HTTruocKTSau[1], CellId_HTTruocKTSau[0]].Value = true;
				}
				else
				{
					sheetView.Cells[CellId_HTTruocKTSau[1], CellId_HTTruocKTSau[0]].Value = false;
				}
			}
			bool flag3 = sheetView.Cells[16, HTKKFile.ColumnLettersToIndex("AA") - 1].Text == "Q";
			bool flag4 = sheetView.Cells[16, HTKKFile.ColumnLettersToIndex("AA") - 1].Text == "M";
			if (idxRow == CellId_KeKhai_Quy[1] && idxCol == CellId_KeKhai_Quy[0])
			{
				if (!flag3)
				{
					sheetView.Cells[CellId_KeKhai_Thang[1], CellId_KeKhai_Thang[0]].Value = true;
					initComboboxThang();
					sheetView.Cells[CellId_KeKhai_Tu[1], CellId_KeKhai_Tu[0]].CellType = comboThang;
					sheetView.Cells[CellId_KeKhai_Den[1], CellId_KeKhai_Den[0]].CellType = comboThang;
				}
				else
				{
					sheetView.Cells[CellId_KeKhai_Thang[1], CellId_KeKhai_Thang[0]].Value = false;
					initComboboxQuy();
					sheetView.Cells[CellId_KeKhai_Tu[1], CellId_KeKhai_Tu[0]].CellType = comboQuy;
					sheetView.Cells[CellId_KeKhai_Den[1], CellId_KeKhai_Den[0]].CellType = comboQuy;
				}
			}
			else if (idxRow == CellId_KeKhai_Thang[1] && idxCol == CellId_KeKhai_Thang[0])
			{
				if (!flag4)
				{
					sheetView.Cells[CellId_KeKhai_Quy[1], CellId_KeKhai_Quy[0]].Value = true;
					initComboboxQuy();
					sheetView.Cells[CellId_KeKhai_Tu[1], CellId_KeKhai_Tu[0]].CellType = comboQuy;
					sheetView.Cells[CellId_KeKhai_Den[1], CellId_KeKhai_Den[0]].CellType = comboQuy;
				}
				else
				{
					sheetView.Cells[CellId_KeKhai_Quy[1], CellId_KeKhai_Quy[0]].Value = false;
					initComboboxThang();
					sheetView.Cells[CellId_KeKhai_Tu[1], CellId_KeKhai_Tu[0]].CellType = comboThang;
					sheetView.Cells[CellId_KeKhai_Den[1], CellId_KeKhai_Den[0]].CellType = comboThang;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void initComboboxQuy()
	{
		comboQuy = new ComboBoxCellType();
		string[] array = new string[4];
		for (int i = 1; i <= 4; i++)
		{
			array[i - 1] = "Q" + i;
		}
		comboQuy.ItemData = array;
		comboQuy.Items = array;
	}

	private void initComboboxThang()
	{
		comboThang = new ComboBoxCellType();
		string[] array = new string[12];
		for (int i = 1; i <= 12; i++)
		{
			array[i - 1] = i.ToString();
		}
		comboThang.ItemData = array;
		comboThang.Items = array;
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy");
			XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai");
			DateTime dateTime = DateTime.ParseExact(ToKhaiInfo.SheetInfos[0].StaticCells[63].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
			if (val.InnerText == "Q")
			{
				val2.InnerText = ((dateTime.Month - 1) / 3 + 1).ToString().ToString() + "/" + dateTime.ToString("yyyy");
			}
			else
			{
				val2.InnerText = dateTime.ToString("MM/yyyy");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/BuTruChoKhoanPNop/tienBangChu"));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/BuTruChoKhoanPNop/tongCot11"));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/tienBangChu"));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/soTienHoanTra"));
		winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = MoneyHelpers.MoneyToText(hTKKCMCCellInfo2.Value);
		winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = MoneyHelpers.MoneyToText(hTKKCMCCellInfo4.Value);
		string text = winFpSpread.Sheets[0].Cells[13, 11].Text;
		winFpSpread.Sheets[0].Cells[13, 11].Text = text;
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[8];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AS") - 1].Text;
				if (ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy") > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AS") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
					value = 2;
				}
				else if (!(ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy") == DateTime.MinValue))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AS") - 1);
				}
				if (!sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("D") - 1].Text.IsNulOrEmpty())
				{
					if (Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AX") - 1].Value) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AX") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AX") - 1);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AX") - 1);
				}
				if (Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AX") - 1].Value) != 0.0)
				{
					if (sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("D") - 1].Text.IsNulOrEmpty())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				}
				string text2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AM") - 1].Text;
				if (string.IsNullOrEmpty(text2))
				{
					continue;
				}
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AG") - 1].Text;
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy"));
				if (hTKKCMCCellInfo.Value.IsNulOrEmpty())
				{
					continue;
				}
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(hTKKCMCCellInfo.Value, "dd/MM/yyyy");
				switch (text3)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text2, "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime("01/" + text2, "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
					}
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
					}
					break;
				case "Y":
					if (ConvertHelpers.ConvertToDateTime("01/01/" + text2, "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime("01/01/" + text2, "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
					}
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text2.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text2.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text2.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text2.Split(new char[1] { '/' })[1], "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
					}
					break;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public List<string> GetDM(string fileDmThue, int loaiHT)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		try
		{
			List<string> list = new List<string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(fileDmThue))
			{
				return null;
			}
			val.Load(fileDmThue);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string obj = array[0];
				string text = "";
				if (obj.ToIntOrDefault() == loaiHT)
				{
					text = array[2];
					list.Add(text);
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}
}
