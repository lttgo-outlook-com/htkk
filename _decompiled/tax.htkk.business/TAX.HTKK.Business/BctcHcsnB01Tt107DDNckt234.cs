using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

internal class BctcHcsnB01Tt107DDNckt234
{
	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_12");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = !Convert.ToBoolean(valueCustom);
			winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("D_20")[1], GeneralFunctions.ArrayCellId("D_20")[0]].Value = false;
			for (int i = 39; i <= 74; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_14");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (!Convert.ToBoolean(valueCustom2))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			for (int i = 39; i <= 48; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_10");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (!Convert.ToBoolean(valueCustom))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			for (int i = 39; i <= 67; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (!Convert.ToBoolean(valueCustom))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			for (int i = 39; i <= 67; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (!Convert.ToBoolean(valueCustom))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void NhapAmDuong(List<string> cell, FpSpread winFpSpread, long iRow, long iCol)
	{
		string item = iRow.ToString();
		if (cell.IndexOf(item) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
		{
			winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		List<string> list = new List<string>();
		list.Add("52");
		list.Add("55");
		List<string> cell = list;
		if (winFpSpread.ActiveSheetIndex == 0)
		{
			NhapAmDuong(cell, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		List<string> cell = new List<string>();
		if (winFpSpread.ActiveSheetIndex == 1)
		{
			NhapAmDuong(cell, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		List<string> list = new List<string>();
		list.Add("47");
		list.Add("48");
		list.Add("49");
		list.Add("54");
		list.Add("55");
		list.Add("60");
		list.Add("61");
		list.Add("62");
		List<string> cell = list;
		if (winFpSpread.ActiveSheetIndex == 2)
		{
			NhapAmDuong(cell, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		List<string> list = new List<string>();
		list.Add("49");
		list.Add("54");
		list.Add("55");
		list.Add("60");
		list.Add("61");
		list.Add("62");
		List<string> cell = list;
		if (winFpSpread.ActiveSheetIndex == 3)
		{
			NhapAmDuong(cell, winFpSpread, iRow, iCol);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		int[] array = GeneralFunctions.ArrayCellId("D_12");
		string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
		if (idxCol == array[0] && idxRow == array[1])
		{
			sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = !Convert.ToBoolean(valueCustom);
			if (!Convert.ToBoolean(valueCustom))
			{
				sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("R") - 1].Text = "";
			}
			if (!Convert.ToBoolean(valueCustom))
			{
				sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = "";
			}
		}
		int[] array2 = GeneralFunctions.ArrayCellId("D_20");
		string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
		if (idxCol != array2[0] || idxRow != array2[1])
		{
			return;
		}
		for (int i = 39; i < 74; i++)
		{
			List<int> list = new List<int>();
			list.Add(59);
			List<int> list2 = list;
			if (!list2.Contains(i + 1))
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom2);
			}
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sheetView = winFpSpread.Sheets[1];
		int[] array = GeneralFunctions.ArrayCellId("D_12");
		string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
		if (idxCol != array[0] || idxRow != array[1])
		{
			return;
		}
		for (int i = 22; i <= 48; i++)
		{
			List<int> list = new List<int>();
			list.Add(32);
			list.Add(40);
			List<int> list2 = list;
			if (!list2.Contains(i + 1))
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = !Convert.ToBoolean(valueCustom);
			}
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sheetView = winFpSpread.Sheets[2];
		int[] array = GeneralFunctions.ArrayCellId("D_35");
		string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
		if (idxCol != array[0] || idxRow != array[1])
		{
			return;
		}
		for (int i = 39; i <= 67; i++)
		{
			List<int> list = new List<int>();
			list.Add(51);
			list.Add(57);
			List<int> list2 = list;
			if (!list2.Contains(i + 1))
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom);
			}
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sheetView = winFpSpread.Sheets[3];
		int[] array = GeneralFunctions.ArrayCellId("D_35");
		string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
		if (idxCol != array[0] || idxRow != array[1])
		{
			return;
		}
		for (int i = 39; i <= 67; i++)
		{
			List<int> list = new List<int>();
			list.Add(41);
			list.Add(51);
			list.Add(57);
			List<int> list2 = list;
			if (!list2.Contains(i + 1))
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom);
			}
		}
	}

	private List<string> getPathDataFileOld(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex)
	{
		List<string> list = new List<string>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			string text2 = "";
			string text3 = "";
			string text4 = "";
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			text = text + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			list.Add(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}
}
