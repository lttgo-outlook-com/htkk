using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

internal class BctcCtckB01Tt334GNDNckt234
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
			for (int i = 24; i < 92; i++)
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
			for (int i = 40; i < 117; i++)
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
			for (int i = 40; i < 151; i++)
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
			for (int i = 40; i < 55; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("D_20");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			int[] array2 = GeneralFunctions.ArrayCellId("D_12");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = !Convert.ToBoolean(valueCustom2);
				if (!Convert.ToBoolean(valueCustom2))
				{
					sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("R") - 1].Text = "";
				}
				if (!Convert.ToBoolean(valueCustom2))
				{
					sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("AC") - 1].Text = "";
				}
			}
			if (idxCol == array[0] && idxRow == array[1])
			{
				for (int i = 39; i < 216; i++)
				{
					List<int> list = new List<int>();
					list.Add(168);
					list.Add(169);
					list.Add(170);
					list.Add(171);
					list.Add(172);
					list.Add(186);
					List<int> list2 = list;
					if (!list2.Contains(i + 1))
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = !Convert.ToBoolean(valueCustom);
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
				for (int i = 24; i < 92; i++)
				{
					List<int> list = new List<int>();
					list.Add(40);
					list.Add(57);
					list.Add(63);
					list.Add(73);
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
				for (int i = 40; i < 117; i++)
				{
					List<int> list = new List<int>();
					list.Add(54);
					list.Add(61);
					list.Add(83);
					list.Add(84);
					list.Add(85);
					list.Add(86);
					list.Add(87);
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
				for (int i = 40; i < 151; i++)
				{
					List<int> list = new List<int>();
					list.Add(86);
					list.Add(93);
					list.Add(115);
					list.Add(116);
					list.Add(117);
					list.Add(118);
					list.Add(119);
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
				for (int i = 40; i < 55; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = !Convert.ToBoolean(valueCustom);
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("C_32");
			string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, array2[1] + 1, array2[0] + 1, "2");
			if (idxCol == array2[0] && idxRow == array2[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 4, isReadDataFile: false, Convert.ToBoolean(valueCustom2), callIsCheck: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
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
			int index = 0;
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[0];
			int num2 = 12;
			int num3 = HTKKFile.ColumnLettersToIndex("R");
			int num4 = HTKKFile.ColumnLettersToIndex("D");
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[num2 - 1, num4 - 1].Value))
			{
				if (string.IsNullOrEmpty(sheetView.Cells[num2 - 1, num3 - 1].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 - 1, num3 - 1);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num2 - 1, num3 - 1);
			}
			int[] array = GeneralFunctions.ArrayCellId("D_14");
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, Convert.ToBoolean(valueCustom));
			if ((num != 2 || !(winFpSpread.TabStrip[index].BackColor == Color.Pink)) && (num != 1 || !(winFpSpread.TabStrip[index].BackColor == Color.Yellow)))
			{
				if (num == 2 && winFpSpread.TabStrip[index].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[index].BackColor = Color.Pink;
				}
				else if (num == 1 && winFpSpread.TabStrip[index].BackColor == Color.White)
				{
					winFpSpread.TabStrip[index].BackColor = Color.Yellow;
				}
				else if (num == 0)
				{
					winFpSpread.TabStrip[index].BackColor = Color.White;
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
			int[] array = GeneralFunctions.ArrayCellId("D_10");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, Convert.ToBoolean(valueCustom));
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
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, Convert.ToBoolean(valueCustom));
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
			int[] array = GeneralFunctions.ArrayCellId("D_34");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 3, Convert.ToBoolean(valueCustom));
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
			SheetView sView = winFpSpread.Sheets[4];
			int[] array = GeneralFunctions.ArrayCellId("C_32");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 4, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void checkSoNamTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheck)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string path = "";
			string path2 = "";
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			int num9 = 0;
			int num10 = 0;
			if (sheetIndex == 0)
			{
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AC");
				path = pathDataFileOld[2];
			}
			if (sheetIndex == 2 || sheetIndex == 3)
			{
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AC");
				path = pathDataFileOld[0];
			}
			if (sheetIndex == 1)
			{
				num7 = HTKKFile.ColumnLettersToIndex("M");
				num8 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("K");
				num4 = HTKKFile.ColumnLettersToIndex("O");
				num5 = HTKKFile.ColumnLettersToIndex("L");
				num6 = HTKKFile.ColumnLettersToIndex("Q");
				num2 = HTKKFile.ColumnLettersToIndex("J");
				path = pathDataFileOld[0];
				path2 = pathDataFileOld[1];
			}
			if (sheetIndex == 4)
			{
				num9 = HTKKFile.ColumnLettersToIndex("P");
				num10 = HTKKFile.ColumnLettersToIndex("AP");
				num3 = HTKKFile.ColumnLettersToIndex("T");
				num4 = HTKKFile.ColumnLettersToIndex("AN");
				num7 = HTKKFile.ColumnLettersToIndex("X");
				num8 = HTKKFile.ColumnLettersToIndex("AO");
				path = pathDataFileOld[0];
				path2 = pathDataFileOld[1];
			}
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			if (sheetIndex == 1)
			{
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (!File.Exists(path) && !File.Exists(path2))
					{
						continue;
					}
					string valueCustom;
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num3 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
						object value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num7 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num7, "NumberCellType");
						object value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1, text2);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num5 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].Locked)
					{
						continue;
					}
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num5, "NumberCellType");
					double num11 = Convert.ToDouble(sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num6 - 1].Value) + Convert.ToDouble(sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value);
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(num11) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1, text);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num5 - 1);
					}
				}
			}
			if (sheetIndex == 0 || sheetIndex == 2 || sheetIndex == 3)
			{
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (!File.Exists(path) || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
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
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
					}
				}
			}
			if (sheetIndex == 4)
			{
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (!File.Exists(path) && !File.Exists(path2))
					{
						continue;
					}
					string valueCustom;
					object value;
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num3 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text4);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num9 && !sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Locked)
					{
						valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num9, "NumberCellType");
						value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num10 - 1].Value;
						if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
						{
							if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].BackColor != Color.Pink)
							{
								num = 1;
								HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1, text);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1);
						}
					}
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num7 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Locked)
					{
						continue;
					}
					valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num7, "NumberCellType");
					value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value;
					if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
					{
						if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].BackColor != Color.Pink)
						{
							num = 1;
							HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1, text4);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1);
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
			string text2 = pathDataFileOld[1];
			string text3 = pathDataFileOld[2];
			string text4 = text3;
			string text5 = "";
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			string text6 = "";
			if (sheetIndex == 0)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
			}
			if (sheetIndex == 2 || sheetIndex == 3)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
			}
			if (sheetIndex == 1)
			{
				text4 = text;
				text5 = text2;
				num = HTKKFile.ColumnLettersToIndex("J");
				num2 = HTKKFile.ColumnLettersToIndex("K");
				num3 = HTKKFile.ColumnLettersToIndex("O");
				num4 = HTKKFile.ColumnLettersToIndex("L");
				num5 = HTKKFile.ColumnLettersToIndex("Q");
				num7 = HTKKFile.ColumnLettersToIndex("M");
				num8 = HTKKFile.ColumnLettersToIndex("R");
				text6 = "J###+Q###";
			}
			if (sheetIndex == 0 || sheetIndex == 1 || sheetIndex == 2 || sheetIndex == 3)
			{
				if (File.Exists(text4) && isCheckLayDl)
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
					hTKKCMCSheetInfo.GetDetailSheetInfo(text4);
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
						{
							int row = hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1;
							sheetView.Cells[row, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[row, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
					}
				}
				else
				{
					for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
					{
						if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
						{
							sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
						}
					}
				}
				if (File.Exists(text5))
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
					hTKKCMCSheetInfo.GetDetailSheetInfo(text5);
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num4)
						{
							continue;
						}
						int row = hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1;
						sheetView.Cells[row, num5 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile && !callIsCheck)
						{
							if (sheetIndex == 1)
							{
								sheetView.Cells[row, num4 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value) + ConvertHelpers.ConvertToDouble(sheetView.Cells[row, num4 - 3].Value);
							}
							else
							{
								sheetView.Cells[row, num4 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
					}
				}
				else
				{
					for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
					{
						if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num4 && !isReadDataFile)
						{
							sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value = ConvertHelpers.ConvertToDouble(sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num4 - 3].Value);
						}
					}
				}
				if (File.Exists(text))
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
					hTKKCMCSheetInfo.GetDetailSheetInfo(text);
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (sheetIndex == 1 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num4)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
					}
					return;
				}
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (sheetIndex == 1 && ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num4 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = 0;
					}
				}
			}
			else
			{
				if (sheetIndex != 4)
				{
					return;
				}
				text4 = text;
				text5 = text2;
				int num9 = HTKKFile.ColumnLettersToIndex("P");
				int num10 = HTKKFile.ColumnLettersToIndex("AJ");
				int num11 = HTKKFile.ColumnLettersToIndex("AP");
				num = HTKKFile.ColumnLettersToIndex("AB");
				num2 = HTKKFile.ColumnLettersToIndex("T");
				num3 = HTKKFile.ColumnLettersToIndex("AN");
				num4 = HTKKFile.ColumnLettersToIndex("AF");
				num7 = HTKKFile.ColumnLettersToIndex("X");
				num8 = HTKKFile.ColumnLettersToIndex("AO");
				if (File.Exists(text5) && isCheckLayDl)
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
					hTKKCMCSheetInfo.GetDetailSheetInfo(text5);
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
						{
							int row = hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1;
							sheetView.Cells[row, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[row, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num4)
						{
							int row = hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1;
							sheetView.Cells[row, num8 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile)
							{
								sheetView.Cells[row, num7 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
					}
				}
				else
				{
					for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
					{
						if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
						{
							sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
						}
						if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num4 && !isReadDataFile)
						{
							sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = 0;
						}
					}
				}
				if (File.Exists(text4) && isCheckLayDl)
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
					hTKKCMCSheetInfo.GetDetailSheetInfo(text4);
					for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
					{
						if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num10)
						{
							int row = hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1;
							sheetView.Cells[row, num11 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							if (!isReadDataFile && !callIsCheck)
							{
								sheetView.Cells[row, num9 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
							}
						}
					}
					return;
				}
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num10 && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num9 - 1].Value = 0;
					}
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
			string item = text;
			string item2 = text;
			string text2 = "";
			string text3 = "";
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			string text4 = HTKKConfig.DATAFILES + "\\" + HeaderInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text4);
			if (sheetIndex == 0)
			{
				text = text + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				string value = ToKhaiInfo.SheetInfos[sheetIndex].DataFile.Replace("_GND", "") + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_";
				if (directoryInfo.Exists)
				{
					FileInfo[] files = directoryInfo.GetFiles();
					if (files.Length > 0)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(value))
							{
								item2 = text4 + "\\" + fileInfo.Name;
								break;
							}
						}
					}
				}
			}
			if (sheetIndex == 1 || sheetIndex == 2 || sheetIndex == 3)
			{
				text = text + "_Q0" + PeriodInfo.Quy + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				int num = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
				int month = GeneralFunctions.GetNgayDauQuy(num, num2, 1, 1).Month;
				Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(month, num2, HeaderInfo.NgayBatDauNamTaiChinh);
				if (num == 1)
				{
					num = 4;
					num2--;
				}
				else
				{
					num--;
				}
				string value2 = "";
				quyNamByThangNam.TryGetValue(num.ToString(), out value2);
				if (num2 == ConvertHelpers.ConvertToInt(value2))
				{
					string value3 = ToKhaiInfo.SheetInfos[sheetIndex].DataFile + "_Q0" + num + value2;
					string value4 = ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
					if (directoryInfo.Exists)
					{
						FileInfo[] files = directoryInfo.GetFiles();
						if (files.Length > 0)
						{
							FileInfo[] array = files;
							foreach (FileInfo fileInfo in array)
							{
								if (fileInfo.Name.Contains(value3) && fileInfo.Name.Contains(value4))
								{
									item = text4 + "\\" + fileInfo.Name;
									break;
								}
							}
						}
					}
				}
			}
			if (sheetIndex == 4)
			{
				int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
				text = text + "_Q04" + (num2 - 1).ToString() + "_0110" + (num2 - 1) + "_3112" + (num2 - 1) + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				int num = ConvertHelpers.ConvertToInt(PeriodInfo.Quy) - 1;
				int month = GeneralFunctions.GetNgayDauQuy(num, num2, 1, 1).Month;
				Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(month, num2, HeaderInfo.NgayBatDauNamTaiChinh);
				string value2 = "";
				quyNamByThangNam.TryGetValue(num.ToString(), out value2);
				if (num2 == ConvertHelpers.ConvertToInt(value2))
				{
					string value3 = ToKhaiInfo.SheetInfos[sheetIndex].DataFile + "_Q0" + num + value2;
					string value4 = ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
					if (directoryInfo.Exists)
					{
						FileInfo[] files = directoryInfo.GetFiles();
						if (files.Length > 0)
						{
							FileInfo[] array = files;
							foreach (FileInfo fileInfo in array)
							{
								if (fileInfo.Name.Contains(value3) && fileInfo.Name.Contains(value4))
								{
									item = text4 + "\\" + fileInfo.Name;
									break;
								}
							}
						}
					}
				}
			}
			list.Add(text);
			list.Add(item);
			list.Add(item2);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "R_49;W_49;R_60;W_60;R_70;W_70;R_79;W_79;R_83;W_83;R_87;W_87;R_91;W_91;R_95;W_95;R_104;W_104;R_157;W_157;";
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("D_10");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (Convert.ToBoolean(valueCustom))
			{
				if ((int)iCol == HTKKFile.ColumnLettersToIndex("J"))
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("J") - 1].Value) + ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Value);
				}
			}
			else if ((int)iCol == HTKKFile.ColumnLettersToIndex("J"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("J") - 1].Value);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "R_41;W_41;R_43;W_43;;R_46;W_46;R_48;W_48;R_49;W_49;R_52;W_52;R_55;W_55;R_57;W_57;R_63;W_63;R_68;W_68;R_69;;W_69;R_70;W_70;R_71;W_71;R_72;W_72;R_89;W_89;R_91;W_91;R_93;W_93;R_96;W_96;R_98;W_98;R_100;W_100";
			if (winFpSpread.ActiveSheetIndex == 2)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "R_48;W_48;R_75;W_75;R_76;W_76;R_84;W_84;R_87;W_87;R_89;W_89;R_95;W_95;R_100;W_100;R_101;W_101;R_102;W_102;R_103;W_103;R_104;W_104;R_121;W_121;R_123;W_123;R_125;W_125;R_127;W_127;R_129;W_129;R_130;W_130;R_132;W_132;R_134;W_134";
			if (winFpSpread.ActiveSheetIndex == 3)
			{
				string value = HTKKFile.ColumnIndexToLetters((int)iCol) + "_" + iRow;
				if (text.IndexOf(value) >= 0 && ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) > 0.0)
				{
					winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value = ConvertHelpers.ConvertToDouble(winFpSpread.ActiveSheet.Cells[(int)iRow - 1, (int)iCol - 1].Value) * -1.0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
