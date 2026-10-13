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

internal class Dnht01aNckt267
{
	private string cellId_SoDeNghiHoan = "H_6";

	private string cellId_HTTruocKTSau = "S_8";

	private string cellId_KTTruocHTSau = "S_10";

	private string cellId_NgayCap = "AE_29";

	private ComboBoxCellType comboQuy = new ComboBoxCellType();

	private ComboBoxCellType comboThang = new ComboBoxCellType();

	private int[] CellId_SoDeNghiHoan => GeneralFunctions.ArrayCellId(cellId_SoDeNghiHoan);

	private int[] CellId_HTTruocKTSau => GeneralFunctions.ArrayCellId(cellId_HTTruocKTSau);

	private int[] CellId_KTTruocHTSau => GeneralFunctions.ArrayCellId(cellId_KTTruocHTSau);

	private int[] CellId_NgayCap => GeneralFunctions.ArrayCellId(cellId_NgayCap);

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.AutoCalculation = true;
			int indexColKyTinh = HTKKFile.ColumnLettersToIndex("AN");
			int num = HTKKFile.ColumnLettersToIndex("BK");
			int[] array = GeneralFunctions.ArrayCellId("BJ_104");
			int[] array2 = GeneralFunctions.ArrayCellId("BK_104");
			int[] array3 = GeneralFunctions.ArrayCellId("BO_104");
			int[] array4 = GeneralFunctions.ArrayCellId("BP_104");
			int[] array5 = GeneralFunctions.ArrayCellId("BJ_105");
			int[] array6 = GeneralFunctions.ArrayCellId("BK_105");
			int[] array7 = GeneralFunctions.ArrayCellId("BO_105");
			int[] array8 = GeneralFunctions.ArrayCellId("BP_105");
			int[] array9 = GeneralFunctions.ArrayCellId("J_110");
			int[] array10 = GeneralFunctions.ArrayCellId("M_110");
			int[] array11 = GeneralFunctions.ArrayCellId("BJ_110");
			int[] array12 = GeneralFunctions.ArrayCellId("BK_110");
			int[] array13 = GeneralFunctions.ArrayCellId("BO_110");
			int[] array14 = GeneralFunctions.ArrayCellId("BP_110");
			if (PeriodInfo.KieuKy == "Q")
			{
				sheetView.Cells[array5[1], array5[0]].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(PeriodInfo.TuThang.Split(new char[1] { '/' })[0]), int.Parse(PeriodInfo.TuThang.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[array6[1], array6[0]].Text = GeneralFunctions.GetNgayCuoiQuy(int.Parse(PeriodInfo.TuThang.Split(new char[1] { '/' })[0]), int.Parse(PeriodInfo.TuThang.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[array7[1], array7[0]].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]), int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
				sheetView.Cells[array8[1], array8[0]].Text = GeneralFunctions.GetNgayCuoiQuy(int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]), int.Parse(PeriodInfo.DenThang.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
			}
			string text = sheetView.Cells[array[1], array[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None);
				string text2 = DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + dateTime.ToString("/MM/yyyy");
				sheetView.Cells[array2[1], array2[0]].Text = text2;
			}
			string text3 = sheetView.Cells[array3[1], array3[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text3))
			{
				DateTime dateTime = DateTime.ParseExact(text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None);
				string text2 = DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + dateTime.ToString("/MM/yyyy");
				sheetView.Cells[array4[1], array4[0]].Text = text2;
			}
			string text4 = sheetView.Cells[array11[1], array11[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text4))
			{
				DateTime dateTime = DateTime.ParseExact(text4, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None);
				string text2 = DateTime.DaysInMonth(dateTime.Year, dateTime.Month) + dateTime.ToString("/MM/yyyy");
				sheetView.Cells[array12[1], array12[0]].Text = text2;
			}
			string text5 = sheetView.Cells[array13[1], array13[0]].Text.Trim();
			if (!string.IsNullOrEmpty(text5))
			{
				DateTime dateTime2 = DateTime.ParseExact(text5, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None);
				string text6 = DateTime.DaysInMonth(dateTime2.Year, dateTime2.Month) + dateTime2.ToString("/MM/yyyy");
				sheetView.Cells[array14[1], array14[0]].Text = text6;
			}
			GetKyKKhai(sheetView, ToKhaiInfo);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text7 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text;
				FomartKyTinh(sheetView, text7, dataLocationTemplateRowStart + i, indexColKyTinh);
			}
			int[] array15 = GeneralFunctions.ArrayCellId("H_14");
			int[] array16 = GeneralFunctions.ArrayCellId("AA_14");
			int[] array17 = GeneralFunctions.ArrayCellId("BI_14");
			string text8 = TTinChung.tenCQTCapCuc;
			string tenCQTNoiNop = TTinChung.tenCQTNoiNop;
			if (text8.IndexOf("-") == 4)
			{
				text8 = text8.Substring(6).Trim();
			}
			if ("10100,70100,71300,71100".Split(new char[1] { ',' }).Contains(TTinChung.maCQTCapCuc))
			{
				sheetView.Cells[array15[1], array15[0]].Text = text8;
				sheetView.Cells[array16[1], array16[0]].Text = TTinChung.tenCQTNoiNop;
				sheetView.Cells[array17[1], array17[0]].Text = TTinChung.maCQTNoiNop;
			}
			else
			{
				sheetView.Cells[array15[1], array15[0]].Text = text8;
				sheetView.Cells[array16[1], array16[0]].Text = text8;
				tenCQTNoiNop = text8;
				sheetView.Cells[array17[1], array17[0]].Text = TTinChung.maCQTCapCuc;
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
		//IL_087d: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("J");
			int num3 = HTKKFile.ColumnLettersToIndex("AI");
			int num4 = HTKKFile.ColumnLettersToIndex("BK");
			int num5 = HTKKFile.ColumnLettersToIndex("AN");
			int num6 = HTKKFile.ColumnLettersToIndex("AE");
			int num7 = HTKKFile.ColumnLettersToIndex("AY");
			int num8 = HTKKFile.ColumnLettersToIndex("BL");
			int num9 = HTKKFile.ColumnLettersToIndex("BM");
			int num10 = HTKKFile.ColumnLettersToIndex("BN");
			int num11 = HTKKFile.ColumnLettersToIndex("BO");
			string path_TongTienNSNN_BangChu = "pathTongTienNSNN_BangChu";
			string path_TongCol11 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/BuTruChoKhoanPNop/tongCot11";
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == path_TongTienNSNN_BangChu);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == path_TongCol11);
			string path_TongTienHoanTra_BangChu = "pathTongTienHoanTra_BangChu";
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == path_TongTienHoanTra_BangChu);
			string path_TongTienHoanTra = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/soTienHoanTra";
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == path_TongTienHoanTra);
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == num3)
			{
				string text = sheetView.Cells[num - 1, num4 - 1].Text;
				FomartKyTinh(sheetView, text, (int)iRow, num5);
				sheetView.Cells[(int)iRow - 1, num5 - 1].Text = "";
			}
			else if (iCol == num6 || iCol == num7)
			{
				int num12 = Convert.ToInt32(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
				string text2 = MoneyHelpers.MoneyToText(num12.ToString());
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = ((num12 == 0) ? "" : text2);
				int num13 = Convert.ToInt32(sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
				string text3 = MoneyHelpers.MoneyToText(num13.ToString());
				sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = ((num13 == 0) ? "" : text3);
			}
			else if (iCol == num5)
			{
				string text4 = sheetView.Cells[num - 1, num4 - 1].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, num5 - 1].Text;
				frmMessage frmMessage = new frmMessage("0456", "3", "", "Thông báo");
				switch (text4)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text5, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = DateTime.ParseExact("01/" + text5, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num9 - 1].Text = DateTime.ParseExact("01/" + text5, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num10 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text5;
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text5, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num9 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num10 - 1].Text = ConvertHelpers.ConvertToDateTime(text5, "dd/MM/yyyy").ToString("MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = ConvertHelpers.ConvertToDateTime(text5, "dd/MM/yyyy").ToString("MM/yyyy");
					break;
				case "Y":
					if (ConvertHelpers.ConvertToDateTime("01/01/" + text5, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = DateTime.ParseExact("01/01/" + text5, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num9 - 1].Text = DateTime.ParseExact("01/12/" + text5, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num10 - 1].Text = "01/" + text5;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = "12/" + text5;
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text5.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text5.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text5.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text5.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num9 - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text5.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text5.Split(new char[1] { '/' })[1]), 1, 1).AddMonths(3).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, num10 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text5;
					break;
				default:
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = "";
					((Form)frmMessage).ShowDialog();
					break;
				}
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("AB"))
			{
				GetKyKKhai(sheetView, ToKhaiInfo);
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
			int num = HTKKFile.ColumnLettersToIndex("AW");
			int num2 = HTKKFile.ColumnLettersToIndex("AX");
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[0];
			int num3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			if ((idxCol == num - 1 || idxCol == num2 - 1) && idxRow >= num3 && idxRow <= num3 + hTKKCMCTableInfo.RowsValues.Count)
			{
				bool flag3 = Convert.ToBoolean(sheetView.Cells[idxRow, num - 1].Value);
				bool flag4 = Convert.ToBoolean(sheetView.Cells[idxRow, num2 - 1].Value);
				if (idxCol == num - 1 && flag3)
				{
					sheetView.Cells[idxRow, num2 - 1].Value = false;
					sheetView.Cells[idxRow, num - 1].Value = true;
				}
				else
				{
					sheetView.Cells[idxRow, num2 - 1].Value = true;
					sheetView.Cells[idxRow, num - 1].Value = false;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		SheetView sheetView = winFpSpread.Sheets[0];
		int indexColKyTinh = HTKKFile.ColumnLettersToIndex("AN");
		int num = HTKKFile.ColumnLettersToIndex("BK");
		HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathTongTienNSNN_BangChu"));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/BuTruChoKhoanPNop/tongCot11"));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("pathTongTienHoanTra_BangChu"));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/soTienHoanTra"));
		winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = MoneyHelpers.MoneyToText(hTKKCMCCellInfo2.Value);
		winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = MoneyHelpers.MoneyToText(hTKKCMCCellInfo4.Value);
		HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
		int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
		for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
		{
			string text = sheetView.Cells[dataLocationTemplateRowStart + num2 - 1, num - 1].Text;
			FomartKyTinh(sheetView, text, dataLocationTemplateRowStart + num2, indexColKyTinh);
		}
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
			string text = sheetView.Cells[CellId_NgayCap[1], CellId_NgayCap[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_NgayCap[1], CellId_NgayCap[0]);
			if (ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy") > DateTime.Now)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayCap[1], CellId_NgayCap[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
				value = 2;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string[] array = new string[8];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1);
				string text2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AT") - 1].Text;
				if (ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
					value = 2;
				}
				else if (!(ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") == DateTime.MinValue))
				{
				}
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AN") - 1].Text;
				if (string.IsNullOrEmpty(text3))
				{
					continue;
				}
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BK") - 1].Text;
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ngay"));
				if (hTKKCMCCellInfo.Value.IsNulOrEmpty())
				{
					continue;
				}
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(hTKKCMCCellInfo.Value, "dd/MM/yyyy");
				switch (text4)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text3, "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime("01/" + text3, "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					break;
				case "Y":
					if ((GeneralFunctions.IsNumeric(text3) && text3.Length == 4 && ConvertHelpers.ConvertToDateTime("01/01/" + text3, "dd/MM/yyyy") == DateTime.MinValue) || ConvertHelpers.ConvertToDateTime("01/01/" + text3, "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
					else if (!GeneralFunctions.IsNumeric(text3) || text3.Length != 4)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
						value = 2;
					}
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text3.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue || ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text3.Split(new char[1] { '/' })[1], "dd/MM/yyyy") > dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
						value = 2;
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

	public void FomartKyTinh(SheetView sViewTK, string maLanPhatSinh, int iRow, int indexColKyTinh)
	{
		try
		{
			MaskCellType maskCellType = new MaskCellType();
			maskCellType.MaskChar = '.';
			switch (maLanPhatSinh)
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
			sViewTK.Cells[iRow - 1, indexColKyTinh - 1].ResetCellType();
			sViewTK.Cells[iRow - 1, indexColKyTinh - 1].CellType = maskCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void GetKyKKhai(SheetView sViewTK, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
		int[] array = GeneralFunctions.ArrayCellId("BR_104");
		string pathNgay = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ngay";
		HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == pathNgay);
		string text = sViewTK.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
		if (!string.IsNullOrEmpty(text) && string.IsNullOrEmpty(sViewTK.Cells[array[1], array[0]].Text))
		{
			DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None);
			sViewTK.Cells[array[1], array[0]].Text = dateTime.ToString("MM/yyyy");
		}
	}
}
