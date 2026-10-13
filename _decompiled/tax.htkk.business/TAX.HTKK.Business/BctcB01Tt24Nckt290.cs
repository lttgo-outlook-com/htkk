using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcB01Tt24Nckt290
{
	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_22");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_28");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol != array2[0] || idxRow != array2[1])
			{
				return;
			}
			List<int> list = new List<int>();
			list.Add(92);
			List<int> list2 = list;
			for (int i = 47; i < 114; i++)
			{
				if (!list2.Contains(i + 1))
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("D_15");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_17");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				for (int i = 26; i < 52; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("K") - 1].Locked = !Convert.ToBoolean(valueCustom2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("D_15");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_17");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol != array2[0] || idxRow != array2[1])
			{
				return;
			}
			List<int> list = new List<int>();
			list.Add(45);
			list.Add(53);
			List<int> list2 = list;
			for (int i = 27; i < 63; i++)
			{
				if (!list2.Contains(i + 1))
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = !Convert.ToBoolean(valueCustom2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheckLayDl)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			switch (sheetIndex)
			{
			case 0:
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				break;
			case 1:
				num = HTKKFile.ColumnLettersToIndex("N");
				num2 = HTKKFile.ColumnLettersToIndex("Q");
				break;
			default:
				num = HTKKFile.ColumnLettersToIndex("K");
				num2 = HTKKFile.ColumnLettersToIndex("M");
				break;
			}
			if (File.Exists(text) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num)
				{
					sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text = text + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			list.Add(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}
}
