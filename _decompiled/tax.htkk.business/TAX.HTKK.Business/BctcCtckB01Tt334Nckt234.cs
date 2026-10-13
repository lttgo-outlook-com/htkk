using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcCtckB01Tt334Nckt234
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
			for (int i = 39; i < 216; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_14");
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
			winFpSpread.Sheets[1].Cells[GeneralFunctions.ArrayCellId("D_12")[1], GeneralFunctions.ArrayCellId("D_12")[0]].Value = false;
			SheetView sheetView = winFpSpread.Sheets[1];
			for (int i = 22; i < 90; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_10");
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
			winFpSpread.Sheets[2].Cells[GeneralFunctions.ArrayCellId("D_35")[1], GeneralFunctions.ArrayCellId("D_35")[0]].Value = false;
			SheetView sheetView = winFpSpread.Sheets[2];
			for (int i = 39; i < 116; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
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

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			winFpSpread.Sheets[3].Cells[GeneralFunctions.ArrayCellId("D_35")[1], GeneralFunctions.ArrayCellId("D_35")[0]].Value = false;
			SheetView sheetView = winFpSpread.Sheets[3];
			for (int i = 39; i < 150; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (Convert.ToBoolean(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 3, isReadDataFile: true, isCheckLayDl: true, callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			winFpSpread.Sheets[4].Cells[GeneralFunctions.ArrayCellId("C_34")[1], GeneralFunctions.ArrayCellId("C_34")[0]].Value = false;
			SheetView sheetView = winFpSpread.Sheets[4];
			for (int i = 41; i < 63; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
			}
			int[] array = GeneralFunctions.ArrayCellId("C_32");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (Convert.ToBoolean(valueCustom))
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 4, isReadDataFile: true, isCheckLayDl: true, callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void NhapAmDuong(string soAmDuong, FpSpread winFpSpread, long iRow, long iCol)
	{
		string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
		if (soAmDuong.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
		{
			winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string soAmDuong = "R_49;W_49;R_60;W_60;R_70;W_70;R_79;W_79;R_83;W_83;R_87;W_87;R_91;W_91;R_95;W_95;R_104;W_104;R_157;W_157";
		if (winFpSpread.ActiveSheetIndex == 0)
		{
			NhapAmDuong(soAmDuong, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string text = "R_40;W_40;R_42;W_42;R_45;W_45;R_47;W_47;R_48;W_48;R_51;W_51;R_54;W_54;R_56;W_56;R_62;W_62;R_67;W_67;R_68;W_68;R_69;W_69;R_70;W_70;R_71;W_71;";
		text += "R_88;W_88;R_90;W_90;R_92;W_92;R_95;W_95;R_97;W_97;R_99;W_99";
		if (winFpSpread.ActiveSheetIndex == 2)
		{
			NhapAmDuong(text, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string soAmDuong = "R_47;W_47;R_74;W_74;R_75;W_75;R_83;W_83;R_86;W_86;R_88;W_88;R_94;W_94;R_99;W_99;R_100;W_100;R_101;W_101;R_102;W_102;R_103;W_103;R_120;W_120;R_122;W_122;R_124;W_124;R_126;W_126;R_128;W_128;R_129;W_129;R_131;W_131;R_133;W_133";
		if (winFpSpread.ActiveSheetIndex == 3)
		{
			NhapAmDuong(soAmDuong, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string soAmDuong = "P_48;T_48;X_48;AB_48;AF_48;AJ_48;AN_48;AS_48";
		if (winFpSpread.ActiveSheetIndex == 4)
		{
			NhapAmDuong(soAmDuong, winFpSpread, iRow, iCol);
		}
		SheetView sheetView = winFpSpread.Sheets[4];
		int num = HTKKFile.ColumnLettersToIndex("AN");
		int num2 = HTKKFile.ColumnLettersToIndex("AS");
		int num3 = HTKKFile.ColumnLettersToIndex("P");
		int num4 = HTKKFile.ColumnLettersToIndex("X");
		int num5 = HTKKFile.ColumnLettersToIndex("AB");
		int num6 = HTKKFile.ColumnLettersToIndex("T");
		int num7 = HTKKFile.ColumnLettersToIndex("AF");
		int num8 = HTKKFile.ColumnLettersToIndex("AJ");
		string text = "P###+X###-AB###";
		string text2 = "T###+AF###-AJ###";
		if (iCol == num3 || iCol == num4 || iCol == num5)
		{
			sheetView.Cells[(int)iRow - 1, num - 1].Formula = text.Replace("###", iRow.ToString());
		}
		if (iCol == num6 || iCol == num7 || iCol == num8)
		{
			sheetView.Cells[(int)iRow - 1, num2 - 1].Formula = text2.Replace("###", iRow.ToString());
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
			if (ToKhaiInfo.SheetInfos[3] != null)
			{
				Validate_3(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[4] != null)
			{
				Validate_4(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
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
			int num3 = 12;
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
			int[] array = GeneralFunctions.ArrayCellId("D_14");
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
			SheetView sView = winFpSpread.Sheets[1];
			int sheetIndex = 1;
			int[] array = GeneralFunctions.ArrayCellId("D_10");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex, Convert.ToBoolean(valueCustom));
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
			SheetView sView = winFpSpread.Sheets[2];
			int sheetIndex = 2;
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[3];
			int sheetIndex = 3;
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[4];
			int sheetIndex = 4;
			int[] array = GeneralFunctions.ArrayCellId("C_32");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex, Convert.ToBoolean(valueCustom));
			int num = HTKKFile.ColumnLettersToIndex("AN");
			int num2 = HTKKFile.ColumnLettersToIndex("AS");
			int num3 = HTKKFile.ColumnLettersToIndex("BD");
			int num4 = HTKKFile.ColumnLettersToIndex("BE");
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			for (int i = 43; i < 63; i++)
			{
				if (i == 54 || i == 57 || i == 58)
				{
					continue;
				}
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, i, num, "NumberCellType");
				object value = sheetView.Cells[i - 1, num3 - 1].Value;
				if (ConvertHelpers.ConvertToDouble(valueCustom2) != ConvertHelpers.ConvertToDouble(value))
				{
					if (sheetView.Cells[i - 1, num - 1].BackColor != Color.Pink)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, i - 1, num - 1, text);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i - 1, num - 1);
				}
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, i, num2, "NumberCellType");
				object value2 = sheetView.Cells[i - 1, num4 - 1].Value;
				if (ConvertHelpers.ConvertToDouble(valueCustom3) != ConvertHelpers.ConvertToDouble(value2))
				{
					if (sheetView.Cells[i - 1, num2 - 1].BackColor != Color.Pink)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, i - 1, num2 - 1, text2);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i - 1, num2 - 1);
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
		try
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
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				for (int i = 39; i < 216; i++)
				{
					List<int> list = new List<int>();
					list.Add(168);
					list.Add(170);
					list.Add(171);
					list.Add(172);
					list.Add(186);
					List<int> list2 = list;
					if (!list2.Contains(i + 1))
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom2);
					}
				}
			}
			int[] array3 = GeneralFunctions.ArrayCellId("D_14");
			string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, array3[1] + 1, array3[0] + 1, "2");
			if (idxCol == array3[0] && idxRow == array3[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: false, Convert.ToBoolean(valueCustom3), callIsCheck: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("D_12");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				for (int i = 22; i < 90; i++)
				{
					List<int> list = new List<int>();
					list.Add(38);
					list.Add(55);
					list.Add(61);
					list.Add(71);
					list.Add(88);
					List<int> list2 = list;
					if (!list2.Contains(i + 1))
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("I") - 1].Locked = !Convert.ToBoolean(valueCustom);
					}
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_10");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: false, Convert.ToBoolean(valueCustom2), callIsCheck: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("D_35");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				for (int i = 39; i < 116; i++)
				{
					List<int> list = new List<int>();
					list.Add(53);
					list.Add(60);
					list.Add(82);
					list.Add(84);
					list.Add(85);
					list.Add(86);
					List<int> list2 = list;
					if (!list2.Contains(i + 1))
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom);
					}
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: false, Convert.ToBoolean(valueCustom2), callIsCheck: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			int[] array = GeneralFunctions.ArrayCellId("D_35");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				for (int i = 39; i < 150; i++)
				{
					List<int> list = new List<int>();
					list.Add(85);
					list.Add(92);
					list.Add(114);
					list.Add(116);
					list.Add(117);
					list.Add(118);
					List<int> list2 = list;
					if (!list2.Contains(i + 1))
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom);
					}
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 3, isReadDataFile: false, Convert.ToBoolean(valueCustom2), callIsCheck: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[4];
			int[] array = GeneralFunctions.ArrayCellId("C_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				for (int i = 41; i < 63; i++)
				{
					List<int> list = new List<int>();
					if (!list.Contains(i + 1))
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = !Convert.ToBoolean(valueCustom);
					}
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("C_32");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 4, isReadDataFile: false, Convert.ToBoolean(valueCustom2), callIsCheck: false);
			}
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 4, isCheck: false);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			int num11 = 0;
			int num12 = 0;
			int num13 = 0;
			int num14 = 0;
			int num15 = 0;
			int num16 = 0;
			int num17 = 0;
			string text = "";
			if (sheetIndex == 0)
			{
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AC");
			}
			else if (sheetIndex == 2 || sheetIndex == 3)
			{
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AC");
			}
			else if (sheetIndex == 1)
			{
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("K");
				num4 = HTKKFile.ColumnLettersToIndex("M");
			}
			else
			{
				num10 = HTKKFile.ColumnLettersToIndex("T");
				num11 = HTKKFile.ColumnLettersToIndex("P");
				num15 = HTKKFile.ColumnLettersToIndex("AX");
				num13 = HTKKFile.ColumnLettersToIndex("AN");
				num10 = HTKKFile.ColumnLettersToIndex("T");
				num14 = HTKKFile.ColumnLettersToIndex("AY");
				num12 = HTKKFile.ColumnLettersToIndex("AF");
				num3 = HTKKFile.ColumnLettersToIndex("X");
				num4 = HTKKFile.ColumnLettersToIndex("BA");
				num2 = HTKKFile.ColumnLettersToIndex("AJ");
				num5 = HTKKFile.ColumnLettersToIndex("AB");
				num6 = HTKKFile.ColumnLettersToIndex("BB");
				num8 = HTKKFile.ColumnLettersToIndex("AS");
				num13 = HTKKFile.ColumnLettersToIndex("AN");
				num17 = HTKKFile.ColumnLettersToIndex("BC");
			}
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			if (sheetIndex == 4)
			{
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (!File.Exists(pathDataFileOld[0]))
					{
						continue;
					}
					string valueCustom;
					object value;
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num11 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num11 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num11, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num15 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num11 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num11 - 1, text2);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num11 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num10 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num10, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num14 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1, text2);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num3 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text2);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num5 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num5, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num6 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1, text2);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num13 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num13 - 1].Locked)
					{
						continue;
					}
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num13, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num17 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num13 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num13 - 1, text2);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num13 - 1);
					}
				}
			}
			else
			{
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (!File.Exists(pathDataFileOld[0]) || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
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
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text2);
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
			int num10 = 0;
			int num11 = 0;
			int num12 = 0;
			int num13 = 0;
			int num14 = 0;
			int num15 = 0;
			int num16 = 0;
			string text2 = "";
			if (sheetIndex == 0)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
			}
			else if (sheetIndex == 2 || sheetIndex == 3)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
			}
			else if (sheetIndex == 1)
			{
				num = HTKKFile.ColumnLettersToIndex("J");
				num2 = HTKKFile.ColumnLettersToIndex("K");
				num3 = HTKKFile.ColumnLettersToIndex("M");
			}
			else
			{
				num9 = HTKKFile.ColumnLettersToIndex("T");
				num10 = HTKKFile.ColumnLettersToIndex("P");
				num14 = HTKKFile.ColumnLettersToIndex("AX");
				num12 = HTKKFile.ColumnLettersToIndex("AN");
				num9 = HTKKFile.ColumnLettersToIndex("T");
				num13 = HTKKFile.ColumnLettersToIndex("AY");
				num11 = HTKKFile.ColumnLettersToIndex("AF");
				num2 = HTKKFile.ColumnLettersToIndex("X");
				num3 = HTKKFile.ColumnLettersToIndex("BA");
				num = HTKKFile.ColumnLettersToIndex("AJ");
				num4 = HTKKFile.ColumnLettersToIndex("AB");
				num5 = HTKKFile.ColumnLettersToIndex("BB");
				num7 = HTKKFile.ColumnLettersToIndex("AS");
				num12 = HTKKFile.ColumnLettersToIndex("AN");
				num16 = HTKKFile.ColumnLettersToIndex("BC");
			}
			if (File.Exists(text) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				if (sheetIndex == 4)
				{
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num9)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num14 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num12)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num13 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num11)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num7)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num16 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num12 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
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
						if (!isReadDataFile && !callIsCheck)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
				return;
			}
			if (sheetIndex == 4)
			{
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num9 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Value = 0;
					}
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num12 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value = 0;
					}
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num11 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
					}
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value = 0;
					}
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num7 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num12 - 1].Value = 0;
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
}
