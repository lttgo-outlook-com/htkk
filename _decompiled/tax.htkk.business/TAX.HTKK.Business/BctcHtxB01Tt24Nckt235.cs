using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcHtxB01Tt24Nckt235
{
	private static string formulaCol5 = "IF(J{0}+P{0}-S{0}-M{0}>0,J{0}+P{0}-S{0}-M{0},0)";

	private static string formulaCol6 = "IF(M{0}+S{0}-P{0}-J{0}>0,M{0}+S{0}-P{0}-J{0},0)";

	private static int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private static int idxCol_M => HTKKFile.ColumnLettersToIndex("M") - 1;

	private static int idxCol_P => HTKKFile.ColumnLettersToIndex("P") - 1;

	private static int idxCol_S => HTKKFile.ColumnLettersToIndex("S") - 1;

	private static int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private static int idxCol_Y => HTKKFile.ColumnLettersToIndex("Y") - 1;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_20");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			sheetView.Cells[19, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = !Convert.ToBoolean(valueCustom);
			winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("D_28")[1], GeneralFunctions.ArrayCellId("D_28")[0]].Value = false;
			for (int i = 47; i <= 73; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_22");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (Convert.ToBoolean(valueCustom2))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: true, isCheckLayDl: true, callIsCheck: true);
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
			winFpSpread.Sheets[1].Cells[GeneralFunctions.ArrayCellId("D_15")[1], GeneralFunctions.ArrayCellId("D_15")[0]].Value = false;
			for (int i = 24; i <= 39; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_13");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (Convert.ToBoolean(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: true, isCheckLayDl: true, callIsCheck: true);
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
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("D_37");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (Convert.ToBoolean(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: true, isCheckLayDl: true, callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void NhapAmDuong(string ctNhapAm, FpSpread winFpSpread, long iRow, long iCol)
	{
		string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
		if (ctNhapAm.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
		{
			winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string ctNhapAm = "R_55;W_55";
		if (winFpSpread.ActiveSheetIndex == 0)
		{
			NhapAmDuong(ctNhapAm, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			string text = "~53~62~79~86~87~88~89~90~91~103~";
			if (text.IndexOf("~" + iRow + "~") >= 0 && (iCol - 1 == idxCol_J || iCol - 1 == idxCol_M || iCol - 1 == idxCol_P || iCol - 1 == idxCol_S))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_V].Formula = string.Format(formulaCol5, iRow.ToString());
				sheetView.Cells[(int)iRow - 1, idxCol_Y].Formula = string.Format(formulaCol6, iRow.ToString());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			if (ToKhaiInfo.SheetInfos[0] != null)
			{
				Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if ((value != 2 || !(winFpSpread.TabStrip[key].BackColor == Color.Pink)) && (value != 1 || !(winFpSpread.TabStrip[key].BackColor == Color.Yellow)))
				{
					if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Pink;
					}
					else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					else if (value == 0)
					{
						winFpSpread.TabStrip[key].BackColor = Color.White;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 0;
			int num2 = 0;
			SheetView sheetView = winFpSpread.Sheets[0];
			int num3 = 20;
			int num4 = HTKKFile.ColumnLettersToIndex("R");
			int num5 = HTKKFile.ColumnLettersToIndex("D");
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[num3 - 1, num5 - 1].Value))
			{
				if (string.IsNullOrEmpty(sheetView.Cells[num3 - 1, num4 - 1].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num3 - 1, num4 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					num2 = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num3 - 1, num4 - 1);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num3 - 1, num4 - 1);
			}
			int[] array = GeneralFunctions.ArrayCellId("D_22");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, Convert.ToBoolean(valueCustom));
			if ((num2 != 2 || !(winFpSpread.TabStrip[num].BackColor == Color.Pink)) && (num2 != 1 || !(winFpSpread.TabStrip[num].BackColor == Color.Yellow)))
			{
				if (num2 == 2 && winFpSpread.TabStrip[num].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[num].BackColor = Color.Pink;
				}
				else if (num2 == 1 && winFpSpread.TabStrip[num].BackColor == Color.White)
				{
					winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				}
				else if (num2 == 0)
				{
					winFpSpread.TabStrip[num].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 1;
			int num2 = 0;
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("D_13");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, Convert.ToBoolean(valueCustom));
			if ((num2 != 2 || !(winFpSpread.TabStrip[num].BackColor == Color.Pink)) && (num2 != 1 || !(winFpSpread.TabStrip[num].BackColor == Color.Yellow)))
			{
				if (num2 == 2 && winFpSpread.TabStrip[num].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[num].BackColor = Color.Pink;
				}
				else if (num2 == 1 && winFpSpread.TabStrip[num].BackColor == Color.White)
				{
					winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				}
				else if (num2 == 0)
				{
					winFpSpread.TabStrip[num].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = 2;
			int num2 = 0;
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("D_37");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, num, Convert.ToBoolean(valueCustom));
			if ((num2 != 2 || !(winFpSpread.TabStrip[num].BackColor == Color.Pink)) && (num2 != 1 || !(winFpSpread.TabStrip[num].BackColor == Color.Yellow)))
			{
				if (num2 == 2 && winFpSpread.TabStrip[num].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[num].BackColor = Color.Pink;
				}
				else if (num2 == 1 && winFpSpread.TabStrip[num].BackColor == Color.White)
				{
					winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				}
				else if (num2 == 0)
				{
					winFpSpread.TabStrip[num].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		int[] array = GeneralFunctions.ArrayCellId("D_20");
		string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
		if (idxCol == array[0] && idxRow == array[1])
		{
			sheetView.Cells[19, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = !Convert.ToBoolean(valueCustom);
			if (!Convert.ToBoolean(valueCustom))
			{
				sheetView.Cells[19, HTKKFile.ColumnLettersToIndex("R") - 1].Text = "";
			}
			if (!Convert.ToBoolean(valueCustom))
			{
				sheetView.Cells[19, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = "";
			}
		}
		int[] array2 = GeneralFunctions.ArrayCellId("D_22");
		string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
		if (idxCol == array2[0] && idxRow == array2[1])
		{
			getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: false, Convert.ToBoolean(valueCustom2), callIsCheck: true);
		}
		int[] array3 = GeneralFunctions.ArrayCellId("D_28");
		string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, array3[1] + 1, array3[0] + 1, "2");
		if (idxCol != array3[0] || idxRow != array3[1])
		{
			return;
		}
		for (int i = 47; i < 73; i++)
		{
			List<int> list = new List<int>();
			list.Add(57);
			List<int> list2 = list;
			if (!list2.Contains(i + 1))
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom3);
			}
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sheetView = winFpSpread.Sheets[1];
		int[] array = GeneralFunctions.ArrayCellId("D_13");
		string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
		if (idxCol == array[0] && idxRow == array[1])
		{
			getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
		}
		int[] array2 = GeneralFunctions.ArrayCellId("D_15");
		string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
		if (idxCol == array2[0] && idxRow == array2[1])
		{
			for (int i = 24; i <= 39; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = !Convert.ToBoolean(valueCustom2);
			}
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		SheetView sView = winFpSpread.Sheets[2];
		int[] array = GeneralFunctions.ArrayCellId("D_37");
		string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
		if (idxCol == array[0] && idxRow == array[1])
		{
			getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isReadDataFile, bool isCheckLayDl, bool callIsCheck)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			switch (sheetIndex)
			{
			case 0:
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
				break;
			case 1:
				num = HTKKFile.ColumnLettersToIndex("J");
				num2 = HTKKFile.ColumnLettersToIndex("K");
				num3 = HTKKFile.ColumnLettersToIndex("M");
				break;
			default:
				num4 = HTKKFile.ColumnLettersToIndex("J");
				num5 = HTKKFile.ColumnLettersToIndex("M");
				num6 = HTKKFile.ColumnLettersToIndex("V");
				num7 = HTKKFile.ColumnLettersToIndex("Y");
				num8 = HTKKFile.ColumnLettersToIndex("AD");
				num9 = HTKKFile.ColumnLettersToIndex("AE");
				break;
			}
			if (File.Exists(text) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				if (sheetIndex == 2)
				{
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num7)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num6)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
					}
					return;
				}
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
				return;
			}
			if (sheetIndex == 2)
			{
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num7 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].Value = 0;
					}
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num6 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value = 0;
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
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

	private void checkSoNamTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheck)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			int num10 = 0;
			string msg = "";
			switch (sheetIndex)
			{
			case 0:
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AC");
				msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[70, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
				break;
			case 2:
				num5 = HTKKFile.ColumnLettersToIndex("J");
				num6 = HTKKFile.ColumnLettersToIndex("M");
				num7 = HTKKFile.ColumnLettersToIndex("V");
				num8 = HTKKFile.ColumnLettersToIndex("Y");
				num9 = HTKKFile.ColumnLettersToIndex("AD");
				num10 = HTKKFile.ColumnLettersToIndex("AE");
				msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[72, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
				break;
			case 1:
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("K");
				num4 = HTKKFile.ColumnLettersToIndex("M");
				msg = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[71, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
				break;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (!File.Exists(pathDataFileOld[0]))
				{
					continue;
				}
				if (sheetIndex == 2)
				{
					string valueCustom;
					object value;
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num5 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num5, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1, msg);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num6 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num6 - 1].Locked)
					{
						continue;
					}
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num6, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num6 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num6 - 1, msg);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num6 - 1);
					}
				}
				else
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
					{
						continue;
					}
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
					object value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, msg);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
					}
				}
			}
			if ((num != 2 || !(winFpSpread.TabStrip[sheetIndex].BackColor == Color.Pink)) && (num != 1 || !(winFpSpread.TabStrip[sheetIndex].BackColor == Color.Yellow)))
			{
				if (num == 2 && winFpSpread.TabStrip[sheetIndex].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.Pink;
				}
				else if (num == 1 && winFpSpread.TabStrip[sheetIndex].BackColor == Color.White)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.Yellow;
				}
				else if (num == 0)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
