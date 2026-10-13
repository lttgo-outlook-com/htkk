using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcB01CtqTt125Nckt305
{
	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
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

	public void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("K");
			int num3 = HTKKFile.ColumnLettersToIndex("M");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("Q");
			int num6 = HTKKFile.ColumnLettersToIndex("S");
			int num7 = HTKKFile.ColumnLettersToIndex("U");
			int num8 = HTKKFile.ColumnLettersToIndex("W");
			int num9 = HTKKFile.ColumnLettersToIndex("Y");
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[4];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[4];
			if (iCol == num2 || iCol == num4 || iCol == num5)
			{
				sheetView.Cells[(int)iRow - 1, num8 - 1].Value = Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num2 - 1].Value) + Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num4 - 1].Value) - Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num5 - 1].Value);
			}
			if (iCol == num3 || iCol == num6 || iCol == num7)
			{
				sheetView.Cells[(int)iRow - 1, num9 - 1].Value = Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num3 - 1].Value) + Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num6 - 1].Value) - Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num7 - 1].Value);
			}
			HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num8 - 1);
			HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num9 - 1);
			if (Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num8 - 1].Value) != Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num2 - 1].Value) + Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num4 - 1].Value) - Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num5 - 1].Value))
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, (int)iRow - 1, num8 - 1, "");
			}
			if (Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num9 - 1].Value) != Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num3 - 1].Value) + Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num6 - 1].Value) - Convert.ToInt64(sheetView.Cells[(int)iRow - 1, num7 - 1].Value))
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, (int)iRow - 1, num9 - 1, "");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, 0, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_36");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol != array2[0] || idxRow != array2[1])
			{
				return;
			}
			List<int> list = new List<int>();
			list.Add(97);
			List<int> list2 = list;
			for (int i = 47; i <= 134; i++)
			{
				if (!list2.Contains(i + 1))
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom2);
				}
			}
			for (int i = 145; i <= 170; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = !Convert.ToBoolean(valueCustom2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			int num = 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			int[] array = GeneralFunctions.ArrayCellId("D_15");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, num, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_17");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				for (int i = 26; i <= 43; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("K") - 1].Locked = !Convert.ToBoolean(valueCustom2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			int num = 2;
			SheetView sheetView = winFpSpread.Sheets[num];
			int[] array = GeneralFunctions.ArrayCellId("D_16");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, num, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_18");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol != array2[0] || idxRow != array2[1])
			{
				return;
			}
			List<int> list = new List<int>();
			list.Add(37);
			list.Add(46);
			List<int> list2 = list;
			for (int i = 28; i <= 57; i++)
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
			throw ex;
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			int num = 3;
			SheetView sheetView = winFpSpread.Sheets[num];
			int[] array = GeneralFunctions.ArrayCellId("D_16");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, num, Convert.ToBoolean(valueCustom));
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_18");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol != array2[0] || idxRow != array2[1])
			{
				return;
			}
			List<int> list = new List<int>();
			list.Add(30);
			list.Add(46);
			list.Add(55);
			List<int> list2 = list;
			for (int i = 28; i <= 57; i++)
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
			throw ex;
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex, bool isCheckLayDl)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, sheetIndex);
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

	private List<string> getPathDataFileOld(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int sheetIndex)
	{
		List<string> list = new List<string>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
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
	}
}
