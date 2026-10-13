using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcCtckB01Tt334HnNckt234
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
			for (int i = 39; i < 220; i++)
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
			for (int i = 22; i < 99; i++)
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
			for (int i = 39; i < 114; i++)
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
		string soAmDuong = "R_49;W_49;R_61;W_61;R_71;W_71;R_80;W_80;R_84;W_84;R_88;W_88;R_92;W_92;R_96;W_96;R_106;W_106;R_159;W_159";
		if (winFpSpread.ActiveSheetIndex == 0)
		{
			NhapAmDuong(soAmDuong, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string soAmDuong = "R_40;W_40;R_42;W_42;R_45;W_45;R_47;W_47;R_48;W_48;R_51;W_51;R_54;W_54;R_60;W_60;R_65;W_65;R_66;W_66;R_67;W_67;R_68;W_68;R_69;W_69;R_86;W_86;R_88;W_88;R_90;W_90;R_93;W_93;R_95;W_95;R_97;W_97";
		if (winFpSpread.ActiveSheetIndex == 2)
		{
			NhapAmDuong(soAmDuong, winFpSpread, iRow, iCol);
		}
	}

	public static void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		string text = "R_47;W_47;R_74;W_74;R_75;W_75;R_83;W_83;R_86;W_86;R_88;W_88;R_94;W_94;R_99;W_99;R_100;W_100;R_101;W_101;R_102;W_102;R_103;W_103";
		text += ";R_120;W_120;R_124;W_124;R_126;W_126;R_128;W_128;R_129;W_129;R_131;W_131;R_133;W_133";
		if (winFpSpread.ActiveSheetIndex == 3)
		{
			NhapAmDuong(text, winFpSpread, iRow, iCol);
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
				for (int i = 39; i < 220; i++)
				{
					List<int> list = new List<int>();
					list.Add(172);
					list.Add(173);
					list.Add(174);
					list.Add(175);
					list.Add(176);
					list.Add(170);
					list.Add(190);
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
				for (int i = 22; i < 99; i++)
				{
					List<int> list = new List<int>();
					list.Add(38);
					list.Add(55);
					list.Add(61);
					list.Add(72);
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
				for (int i = 39; i < 114; i++)
				{
					List<int> list = new List<int>();
					list.Add(53);
					list.Add(58);
					list.Add(80);
					list.Add(81);
					list.Add(82);
					list.Add(83);
					list.Add(84);
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
					list.Add(115);
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
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
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
				else if (sheetIndex == 0 && (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex == 130 || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex == 149))
				{
					if (ConvertHelpers.ConvertToDouble(valueCustom) < 0.0)
					{
						num = 1;
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex == 130)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text3);
						}
						else if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex == 149)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text4);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
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
				num9 = HTKKFile.ColumnLettersToIndex("P");
				num10 = HTKKFile.ColumnLettersToIndex("T");
				num11 = HTKKFile.ColumnLettersToIndex("AF");
				num12 = HTKKFile.ColumnLettersToIndex("AK");
				num13 = HTKKFile.ColumnLettersToIndex("AS");
				num14 = HTKKFile.ColumnLettersToIndex("AT");
			}
			if (File.Exists(text) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				if (sheetIndex == 4)
				{
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num12)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num14 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num11)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num13 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
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
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num12 && !isReadDataFile && !callIsCheck)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Value = 0;
					}
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num11 && !isReadDataFile && !callIsCheck)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value = 0;
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile && !callIsCheck)
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
