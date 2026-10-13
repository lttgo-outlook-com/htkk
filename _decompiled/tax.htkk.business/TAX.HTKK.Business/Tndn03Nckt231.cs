using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn03Nckt231
{
	private static Color lockCellBackColor = Color.FromArgb(255, 192, 192, 192);

	private static Color editableCellBackColor = Color.FromArgb(0, 0, 0, 0);

	public bool ValidatePeriod_axfpsPhuLuc(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) < 2017)
			{
				if ((bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0482", "2", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return frmMessage.Yes;
				}
			}
			else if ((bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0483", "2", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return frmMessage.Yes;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static void Refresh_1(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].ResetFormula();
				winFpSpread.Sheets[1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Formula = "E37+E40";
			}
			string formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = formula;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_2(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[2].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].ResetFormula();
				winFpSpread.Sheets[2].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[2].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Formula = "E25+E28+E29+E30+E31+E34+E35-E36-E37";
			}
			string formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = formula;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_3(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[3].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].ResetFormula();
				winFpSpread.Sheets[3].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[3].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Formula = "E46+E49";
			}
			string formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = formula;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_4(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[0].Cells[50, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				winFpSpread.Sheets[0].Cells[50, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
				winFpSpread.Sheets[4].Cells[31, HTKKFile.ColumnLettersToIndex("M") - 1].ResetFormula();
				winFpSpread.Sheets[4].Cells[31, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			winFpSpread.Sheets[4].Cells[31, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "SUM(M25:M31)";
			if (winFpSpread.Sheets[4].Visible)
			{
				sheetView.Cells[50, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[50, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[4].SheetName + "'!M32";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_6(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[6].Cells[3, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
				winFpSpread.Sheets[6].Cells[3, HTKKFile.ColumnLettersToIndex("U") - 1].Value = 0;
				winFpSpread.Sheets[6].Cells[8, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
				winFpSpread.Sheets[6].Cells[8, HTKKFile.ColumnLettersToIndex("U") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[6].Cells[3, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U50";
				winFpSpread.Sheets[6].Cells[8, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U56";
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible)
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U4+'" + winFpSpread.Sheets[7].SheetName + "'!U7";
			}
			else
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible | winFpSpread.Sheets[8].Visible)
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
			}
			else
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void UpdateFormula_6(FpSpread winFpSpread)
	{
		try
		{
			winFpSpread.Sheets[6].Cells[3, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U50";
			winFpSpread.Sheets[6].Cells[8, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U56";
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible)
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U4+'" + winFpSpread.Sheets[7].SheetName + "'!U7";
			}
			else
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible | winFpSpread.Sheets[8].Visible)
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
			}
			else
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void UpdateFormula_7(FpSpread winFpSpread)
	{
		try
		{
			winFpSpread.Sheets[7].Cells[6, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U60";
			winFpSpread.Sheets[7].Cells[10, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U65";
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible)
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U4+'" + winFpSpread.Sheets[7].SheetName + "'!U7";
			}
			else
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible | winFpSpread.Sheets[8].Visible)
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
			}
			else
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_7(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[7].Cells[6, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
				winFpSpread.Sheets[7].Cells[6, HTKKFile.ColumnLettersToIndex("U") - 1].Value = 0;
				winFpSpread.Sheets[7].Cells[10, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
				winFpSpread.Sheets[7].Cells[10, HTKKFile.ColumnLettersToIndex("U") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[7].Cells[6, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U60";
				winFpSpread.Sheets[7].Cells[10, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = "U65";
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible)
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U4+'" + winFpSpread.Sheets[7].SheetName + "'!U7";
			}
			else
			{
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible | winFpSpread.Sheets[8].Visible)
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
			}
			else
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void UpdateFormula_8(FpSpread winFpSpread, int actionType)
	{
		try
		{
			winFpSpread.Sheets[8].Cells[3, HTKKFile.ColumnLettersToIndex("X") - 1].Formula = "X43";
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible | winFpSpread.Sheets[8].Visible)
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
			}
			else
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_8(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[8].Cells[3, HTKKFile.ColumnLettersToIndex("X") - 1].ResetFormula();
				winFpSpread.Sheets[8].Cells[3, HTKKFile.ColumnLettersToIndex("X") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[8].Cells[3, HTKKFile.ColumnLettersToIndex("X") - 1].Formula = "X43";
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (winFpSpread.Sheets[6].Visible | winFpSpread.Sheets[7].Visible | winFpSpread.Sheets[8].Visible)
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
			}
			else
			{
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_9(FpSpread winFpSpread, int actionType)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if (actionType == 2)
			{
				winFpSpread.Sheets[9].Cells[27, HTKKFile.ColumnLettersToIndex("AR") - 1].ResetFormula();
				winFpSpread.Sheets[9].Cells[27, HTKKFile.ColumnLettersToIndex("AR") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[9].Cells[27, HTKKFile.ColumnLettersToIndex("AR") - 1].Formula = "SUM(AR25:AR27)";
			}
			if (winFpSpread.Sheets[9].Visible)
			{
				sheetView.Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				sheetView.Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[9].SheetName + "'!AR19";
			}
			else
			{
				sheetView.Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				sheetView.Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_10(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
				winFpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("O") - 1].Value = 0;
				winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
				winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[10].SheetName + "'!E29";
				winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Refresh_11(FpSpread winFpSpread, int actionType)
	{
		try
		{
			if (actionType == 2)
			{
				winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
				winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].Value = 0;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[11].SheetName + "'!O18";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		winFpSpread.ResumeLayout();
		checkTKChinh(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, levelWarning);
		checkPL3_1A(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_1B(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_1C(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_2A(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_2B(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3A(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3B(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_3C(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_4(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_5(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_6(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_7(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_8(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_9(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		checkPL3_1_ND20(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	private void checkTKChinh(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_015c: Unknown result type (might be due to invalid IL or missing references)
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			if (PeriodInfo.LoaiToKhai == "C")
			{
				if ((bool)sheetView.Cells[9, 3].Value)
				{
					if (ToKhaiInfo.SheetInfos[13] == null)
					{
						frmMessage frmMessage = new frmMessage("0337", "3", "", "Thông báo");
						HTKKFile.SetNodeBackGroundPink(sheetView, 9, 3, HTKKFile.GetMessage("0337"));
						value = 2;
						((Form)frmMessage).ShowDialog();
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, 9, 3);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 9, 3);
				}
				if ((bool)sheetView.Cells[11, 3].Value)
				{
					if (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) < 2017)
					{
						if (ToKhaiInfo.SheetInfos[12] == null)
						{
							frmMessage frmMessage = new frmMessage("0338", "3", "", "Thông báo");
							HTKKFile.SetNodeBackGroundPink(sheetView, 11, 3, HTKKFile.GetMessage("0338"));
							value = 2;
							((Form)frmMessage).ShowDialog();
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, 11, 3);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, 11, 3);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 11, 3);
				}
				double num = ConvertHelpers.ConvertToDouble(sheetView.Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Value);
				if (num != 0.0)
				{
					if (ToKhaiInfo.SheetInfos[1] == null && ToKhaiInfo.SheetInfos[2] == null && ToKhaiInfo.SheetInfos[3] == null)
					{
						frmMessage frmMessage = new frmMessage("0347", "3", "", "Thông báo");
						HTKKFile.SetNodeBackGroundPink(sheetView, 30, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage("0347"));
						value = 2;
						((Form)frmMessage).ShowDialog();
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 30, HTKKFile.ColumnLettersToIndex("O") - 1);
				}
			}
			double num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_50"));
			double num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_51"));
			double num4 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_52"));
			double num5 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_48"));
			double num6 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_49"));
			if (num2 > num5 - num6 && num5 - num6 <= 0.0)
			{
				if (num3 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 50, 14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_120")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 50, 14);
				}
				if (num4 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 51, 14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_121")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 51, 14);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 50, 14);
				HTKKFile.RemoveValidate(sheetView, 51, 14);
			}
			double num7 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_53"));
			if (num7 > 0.0)
			{
				double num8 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_54"));
				if (100.0 * num8 > num7 * 10.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 53, 14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 53, 14);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 53, 14);
			}
			double num9 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_59"));
			int[] array = GeneralFunctions.ArrayCellId("O_59");
			if (num9 < 0.0 || num9 > 100.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			}
			double num10 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_60"));
			double num11 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_61"));
			double num12 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_62"));
			double num13 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_65"));
			double num14 = num10 - num11 - num12 - num13;
			double num15 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_63"));
			double num16 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "O_64"));
			if (num12 < num15 + num16 && (winFpSpread.Sheets[6].Visible || winFpSpread.Sheets[7].Visible))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 61, 14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 61, 14);
			}
			int row = 82;
			int column = HTKKFile.ColumnLettersToIndex("F") - 1;
			int num17 = HTKKFile.ColumnLettersToIndex("I") - 1;
			bool flag = false;
			if (sheetView.Cells[row, column] != null)
			{
				flag = (bool)sheetView.Cells[row, column].Value;
			}
			int[] array2 = new int[4] { 84, 86, 88, 90 };
			for (int i = 0; i < array2.Length; i++)
			{
				sheetView.Cells[array2[i], num17].Locked = !flag;
				if (array2[i] != 90)
				{
					if ((string.IsNullOrEmpty(ConvertHelpers.ConvertToString(sheetView.Cells[array2[i], num17].Value)) || (sheetView.Cells[array2[i], num17].CellType.ToString() == "NumberCellType" && ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[i], num17].Value) == 0.0)) && flag)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[i], num17, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, array2[i], num17);
					}
				}
			}
			int[] array3 = GeneralFunctions.ArrayCellId("I_87");
			string text = sheetView.Cells[array3[1], array3[0]].Text;
			bool flag2 = false;
			bool success = false;
			if (!string.IsNullOrEmpty(text))
			{
				DateTime t = HTKKFormater.ConvertToDateTime(text, out success);
				DateTime hanNopToKhai = GeneralFunctions.GetHanNopToKhai(PeriodInfo);
				if (success)
				{
					if (DateTime.Compare(t, hanNopToKhai) <= 0)
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, array3[1], array3[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						if (value == 0)
						{
							value = 1;
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
					}
				}
				else if (!success)
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, array3[1], array3[0], HTKKFile.GetMessage("0403"));
					value = 2;
				}
			}
			int[] array4 = GeneralFunctions.ArrayCellId("N_95");
			string text2 = sheetView.Cells[array4[1], array4[0]].Text;
			DateTime t2 = DateTime.MinValue;
			bool success2 = false;
			bool success3 = false;
			if (!string.IsNullOrEmpty(text2))
			{
				t2 = HTKKFormater.ConvertToDateTime(text2, out success2);
			}
			int[] array5 = GeneralFunctions.ArrayCellId("J_95");
			string text3 = sheetView.Cells[array5[1], array5[0]].Text;
			DateTime t3 = new DateTime(1900, 1, 1);
			if (!string.IsNullOrEmpty(text3))
			{
				t3 = HTKKFormater.ConvertToDateTime(text3, out success3);
			}
			if (success3 && success2)
			{
				if (DateTime.Compare(t3, t2) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array4[1], array4[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_92")));
					value = 2;
				}
				else
				{
					string denNgayNC = getDenNgayNC(HeaderInfo.NgayBatDauNamTaiChinh, ConvertHelpers.ConvertToInt(PeriodInfo.Nam));
					bool success4 = false;
					DateTime t4 = HTKKFormater.ConvertToDateTime(denNgayNC, out success4);
					if (success4)
					{
						if (DateTime.Compare(t4, t2) < 0)
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, array4[1], array4[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_91")));
							if (value == 0)
							{
								value = 1;
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, array4[1], array4[0]);
						}
					}
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
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_1A(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			SheetView sView = winFpSpread.Sheets[1];
			double num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sView, "E_24"));
			double num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sView, "E_23"));
			int[] array = GeneralFunctions.ArrayCellId("E_24");
			if (num > num2)
			{
				HTKKFile.SetNodeBackGroundPink(sView, array[1], array[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_12")));
				if (value == 0)
				{
					value = 2;
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, array[1], array[0]);
			}
			double num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sView, "E_36"));
			double num4 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sView, "E_35"));
			if (num3 > num4)
			{
				HTKKFile.SetNodeBackGroundYellow(sView, 35, 4, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_13")));
				if (value == 0)
				{
					value = 1;
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, 35, 4);
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_1B(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			SheetView sheetView = winFpSpread.Sheets[2];
			if (sheetView.Visible)
			{
				switch (value)
				{
				case 2:
					winFpSpread.TabStrip[2].BackColor = Color.Pink;
					break;
				case 1:
					winFpSpread.TabStrip[2].BackColor = Color.Yellow;
					break;
				default:
					winFpSpread.TabStrip[2].BackColor = Color.White;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_1C(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(3, out value);
			SheetView sheetView = winFpSpread.Sheets[3];
			if (sheetView.Visible)
			{
				double num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "E_50"));
				double num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_31"));
				int[] array = GeneralFunctions.ArrayCellId("E_50");
				if (num == num2)
				{
					HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
				}
				switch (value)
				{
				case 2:
					winFpSpread.TabStrip[3].BackColor = Color.Pink;
					break;
				case 1:
					winFpSpread.TabStrip[3].BackColor = Color.Yellow;
					break;
				default:
					winFpSpread.TabStrip[3].BackColor = Color.White;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_2A(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(4, out value);
			SheetView sheetView = winFpSpread.Sheets[4];
			if (!sheetView.Visible)
			{
				return;
			}
			int num = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_26"));
			int num2 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_27"));
			int num3 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_28"));
			int num4 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_29"));
			int num5 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_30"));
			int num6 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int[] array = new int[5] { num, num2, num3, num4, num5 };
			for (int i = 0; i < array.Length; i++)
			{
				if (array[i] > num6)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, i + 25, 2, "Năm phát sinh lỗ phải nhỏ hơn năm quyết toán thuế TNDN.");
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i + 25, 2);
				}
			}
			double num7 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_26"));
			double num8 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_27"));
			double num9 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_28"));
			double num10 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_29"));
			double num11 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_30"));
			double[] array2 = new double[5] { num7, num8, num9, num10, num11 };
			double num12 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_26"));
			double num13 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_27"));
			double num14 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_28"));
			double num15 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_29"));
			double num16 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_30"));
			double[] array3 = new double[5] { num12, num13, num14, num15, num16 };
			double num17 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_26"));
			double num18 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_27"));
			double num19 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_28"));
			double num20 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_29"));
			double num21 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_30"));
			double[] array4 = new double[5] { num17, num18, num19, num20, num16 };
			for (int i = 0; i < array2.Length; i++)
			{
				if (array2[i] < array3[i] + array4[i])
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, i + 25, 3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i + 25, 3);
				}
			}
			double num22 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_32"));
			double num23 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "R_32"));
			double num24 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_48"));
			double num25 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_49"));
			double num26 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_50"));
			if (num22 + num23 > num24 - num25 && num26 < num24 - num25)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 31, 17, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, 31, 17);
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[4].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[4].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[4].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_2B(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(5, out value);
			SheetView sheetView = winFpSpread.Sheets[5];
			if (!sheetView.Visible)
			{
				return;
			}
			int num = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_26").ToString());
			int num2 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_27").ToString());
			int num3 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_28").ToString());
			int num4 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_29").ToString());
			int num5 = ConvertHelpers.ConvertToInt(GeneralFunctions.GetObjectValue(sheetView, "C_30").ToString());
			int num6 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int[] array = new int[5] { num, num2, num3, num4, num5 };
			for (int i = 0; i < array.Length; i++)
			{
				if (array[i] > num6)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, i + 25, 2, "Năm phát sinh lỗ phải nhỏ hơn năm quyết toán thuế TNDN.");
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i + 25, 2);
				}
			}
			double num7 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_26"));
			double num8 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_27"));
			double num9 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_28"));
			double num10 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_29"));
			double num11 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "D_30"));
			double[] array2 = new double[5] { num7, num8, num9, num10, num11 };
			double num12 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_26"));
			double num13 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_27"));
			double num14 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_28"));
			double num15 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_29"));
			double num16 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "H_30"));
			double[] array3 = new double[5] { num12, num13, num14, num15, num16 };
			double num17 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_26"));
			double num18 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_27"));
			double num19 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_28"));
			double num20 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_29"));
			double num21 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_30"));
			double[] array4 = new double[5] { num17, num18, num19, num20, num21 };
			for (int i = 0; i < array2.Length; i++)
			{
				if (array2[i] < array3[i] + array4[i])
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, i + 25, 3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_14")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i + 25, 3);
				}
			}
			double num22 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "M_32"));
			double num23 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "R_32"));
			double num24 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_48"));
			double num25 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_49"));
			double num26 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_50"));
			if (num22 + num23 > num24 - num25 && num26 < num24 - num25)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 31, 17, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, 31, 17);
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[5].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[5].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[5].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_3A(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
	}

	private void checkPL3_3B(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
	}

	private void checkPL3_3C(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(8, out value);
			SheetView sheetView = winFpSpread.Sheets[8];
			if (!sheetView.Visible)
			{
				return;
			}
			for (int i = 0; i < ToKhaiInfo.SheetInfos[8].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[8].Tables[i];
				double num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[8], "X_" + hTKKCMCTableInfo.Header.StaticCells[13].CellTemplate.RowIndex));
				double num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[8], "X_" + hTKKCMCTableInfo.Header.StaticCells[12].CellTemplate.RowIndex));
				if (num > num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCTableInfo.Header.StaticCells[13].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCTableInfo.Header.StaticCells[13].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("X") - 1);
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[8].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[8].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[8].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(9, out value);
			SheetView sheetView = winFpSpread.Sheets[9];
			if (!sheetView.Visible)
			{
				return;
			}
			int num = ToKhaiInfo.SheetInfos[9].Tables[0].Location.DataLocationTemplateRowStart - 1;
			int count = ToKhaiInfo.SheetInfos[9].Tables[0].RowsValues.Count;
			for (int i = 0; i < count; i++)
			{
				double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[i + num, HTKKFile.ColumnLettersToIndex("AR") - 1].Value);
				double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[i + num, HTKKFile.ColumnLettersToIndex("AN") - 1].Value);
				double num4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[i + num, HTKKFile.ColumnLettersToIndex("Z") - 1].Value);
				if (num2 > num3 || num2 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, i + num, HTKKFile.ColumnLettersToIndex("AR") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, i + num, HTKKFile.ColumnLettersToIndex("AR") - 1);
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[9].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[9].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[9].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(10, out value);
			SheetView sheetView = winFpSpread.Sheets[10];
			if (!sheetView.Visible)
			{
				return;
			}
			double num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "E_29"));
			double num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "E_30"));
			double num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[5], "M_32"));
			if (num > 0.0)
			{
				if (num2 > num3 && winFpSpread.Sheets[5].Visible)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 29, 4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_93")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 29, 4);
				}
			}
			else if (num < 0.0)
			{
				sheetView.Cells[29, 4].Value = 0;
			}
			double num4 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "E_31"));
			double num5 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(sheetView, "E_33"));
			if (num4 >= 0.0)
			{
				if (num5 > num4 * 0.1)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 32, 4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_94")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 32, 4);
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[10].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[10].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[10].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
	}

	private void checkPL3_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(15, out value);
			bool flag = false;
			SheetView sheetView = winFpSpread.Sheets[15];
			if (!sheetView.Visible)
			{
				return;
			}
			int num = ToKhaiInfo.SheetInfos[15].Tables[0].RowsValues.Count - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[15];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string[] lstRowValue = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, lstRowValue))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, HTKKFile.ColumnLettersToIndex("T") - 1);
				}
				else
				{
					flag = true;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[15].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[15].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[15].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_7(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(12, out value);
			SheetView sheetView = winFpSpread.Sheets[12];
			if (!sheetView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[12];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int colunmIndexByColumName = HTKKFormater.GetColunmIndexByColumName("D");
			int colunmIndexByColumName2 = HTKKFormater.GetColunmIndexByColumName("I");
			int colunmIndexByColumName3 = HTKKFormater.GetColunmIndexByColumName("O");
			int colunmIndexByColumName4 = HTKKFormater.GetColunmIndexByColumName("U");
			int colunmIndexByColumName5 = HTKKFormater.GetColunmIndexByColumName("V");
			int colunmIndexByColumName6 = HTKKFormater.GetColunmIndexByColumName("W");
			int colunmIndexByColumName7 = HTKKFormater.GetColunmIndexByColumName("X");
			int colunmIndexByColumName8 = HTKKFormater.GetColunmIndexByColumName("Y");
			int colunmIndexByColumName9 = HTKKFormater.GetColunmIndexByColumName("Z");
			int colunmIndexByColumName10 = HTKKFormater.GetColunmIndexByColumName("AA");
			int colunmIndexByColumName11 = HTKKFormater.GetColunmIndexByColumName("AB");
			int colunmIndexByColumName12 = HTKKFormater.GetColunmIndexByColumName("AC");
			int colunmIndexByColumName13 = HTKKFormater.GetColunmIndexByColumName("AD");
			int colunmIndexByColumName14 = HTKKFormater.GetColunmIndexByColumName("AE");
			int colunmIndexByColumName15 = HTKKFormater.GetColunmIndexByColumName("AF");
			int colunmIndexByColumName16 = HTKKFormater.GetColunmIndexByColumName("AG");
			string[] array = new string[16];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName);
				array[1] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName2);
				array[2] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName3);
				array[3] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName4);
				array[4] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName5);
				array[5] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName6);
				array[6] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName7);
				array[7] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName8);
				array[8] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName9);
				array[9] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName10);
				array[10] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName11);
				array[11] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName12);
				array[12] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName13);
				array[13] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName14);
				array[14] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName15);
				array[15] = ConvertHelpers.ConvertCellValueToText(sheetView, num + i, colunmIndexByColumName16);
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName3);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName4);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName5);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName6);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName7);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName8);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName9);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName10);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName11);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName12);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName13);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName14);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName15);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName16);
					continue;
				}
				bool flag = false;
				for (int j = 3; j < array.Length; j++)
				{
					if (array[j].Equals("True"))
					{
						flag = true;
					}
				}
				if (!flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName5, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName6, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName7, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName8, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName9, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName10, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName11, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName12, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName13, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName14, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName15, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, colunmIndexByColumName16, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName4);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName5);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName6);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName7);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName8);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName9);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName10);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName11);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName12);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName13);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName14);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName15);
					HTKKFile.RemoveValidate(sheetView, num + i, colunmIndexByColumName16);
				}
			}
			double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[45, HTKKFile.ColumnLettersToIndex("M") - 1].Value);
			double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[45, HTKKFile.ColumnLettersToIndex("Y") - 1].Value);
			if (num2 < 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 45, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 45, HTKKFile.ColumnLettersToIndex("M") - 1);
			}
			if (num3 < 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 45, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 45, HTKKFile.ColumnLettersToIndex("Y") - 1);
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[12].Tables.Count; i++)
			{
				int num4 = ToKhaiInfo.SheetInfos[12].Tables[i].Location.DataLocationTemplateRowStart - 1;
				int count = ToKhaiInfo.SheetInfos[12].Tables[i].RowsValues.Count;
				for (int j = 0; j < count; j++)
				{
					double num5 = ConvertHelpers.ConvertToDouble(sheetView.Cells[j + num4, HTKKFile.ColumnLettersToIndex("M") - 1].Value);
					if (num5 < 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, j + num4, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, j + num4, HTKKFile.ColumnLettersToIndex("M") - 1);
					}
				}
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[12].Tables.Count; i++)
			{
				int num4 = ToKhaiInfo.SheetInfos[12].Tables[i].Location.DataLocationTemplateRowStart - 1;
				int count = ToKhaiInfo.SheetInfos[12].Tables[i].RowsValues.Count;
				for (int j = 0; j < count; j++)
				{
					double num5 = ConvertHelpers.ConvertToDouble(sheetView.Cells[j + num4, HTKKFile.ColumnLettersToIndex("Y") - 1].Value);
					if (num5 < 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, j + num4, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, j + num4, HTKKFile.ColumnLettersToIndex("Y") - 1);
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[12].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[12].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[12].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(13, out value);
			bool flag = false;
			SheetView sheetView = winFpSpread.Sheets[13];
			if (!sheetView.Visible)
			{
				return;
			}
			int num = ToKhaiInfo.SheetInfos[13].Tables[0].RowsValues.Count - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[13];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int colunmIndexByColumName = HTKKFormater.GetColunmIndexByColumName("D");
			int colunmIndexByColumName2 = HTKKFormater.GetColunmIndexByColumName("F");
			int colunmIndexByColumName3 = HTKKFormater.GetColunmIndexByColumName("I");
			int colunmIndexByColumName4 = HTKKFormater.GetColunmIndexByColumName("K");
			int colunmIndexByColumName5 = HTKKFormater.GetColunmIndexByColumName("L");
			int colunmIndexByColumName6 = HTKKFormater.GetColunmIndexByColumName("O");
			int colunmIndexByColumName7 = HTKKFormater.GetColunmIndexByColumName("P");
			int colunmIndexByColumName8 = HTKKFormater.GetColunmIndexByColumName("Q");
			int colunmIndexByColumName9 = HTKKFormater.GetColunmIndexByColumName("R");
			int colunmIndexByColumName10 = HTKKFormater.GetColunmIndexByColumName("S");
			int colunmIndexByColumName11 = HTKKFormater.GetColunmIndexByColumName("U");
			int colunmIndexByColumName12 = HTKKFormater.GetColunmIndexByColumName("W");
			int colunmIndexByColumName13 = HTKKFormater.GetColunmIndexByColumName("Y");
			int colunmIndexByColumName14 = HTKKFormater.GetColunmIndexByColumName("Z");
			string[] array = new string[14];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName);
				array[1] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName2);
				array[2] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName3);
				array[3] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName4);
				array[4] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName5);
				array[5] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName6);
				array[6] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName7);
				array[7] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName8);
				array[8] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName9);
				array[9] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName10);
				array[10] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName11);
				array[11] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName12);
				array[12] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName13);
				array[13] = ConvertHelpers.ConvertCellValueToText(sheetView, num2 + i, colunmIndexByColumName14);
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName2);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName3);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName4);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName5);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName6);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName7);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName8);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName9);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName10);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName11);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName12);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName13);
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName14);
					continue;
				}
				flag = true;
				if (!ValidateBase.CheckMaSoThue(array[1]) && !string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, colunmIndexByColumName2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName2);
				}
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, colunmIndexByColumName, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName);
				}
				if (string.IsNullOrEmpty(sheetView.Cells[num2 + i, colunmIndexByColumName3].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, colunmIndexByColumName3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName3);
				}
				if (string.IsNullOrEmpty(sheetView.Cells[num2 + i, colunmIndexByColumName4].Text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, colunmIndexByColumName4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, colunmIndexByColumName4);
				}
			}
			double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[28, HTKKFile.ColumnLettersToIndex("K") - 1].Value);
			if (num3 != 0.0 && !flag)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 28, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_104")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, 28, HTKKFile.ColumnLettersToIndex("K") - 1);
			}
			double num4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[42 + num, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
			if (num4 != num3 && num4 != 0.0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, 42 + num, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_87")));
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, 42 + num, HTKKFile.ColumnLettersToIndex("U") - 1);
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[13].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[13].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[13].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_9(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(14, out value);
			SheetView sheetView = winFpSpread.Sheets[14];
			if (!sheetView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[14];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			if (tables.Count > 0)
			{
				HTKKCMCStaticOfTable header = tables[0].Header;
				string[] array = new string[8];
				string defaultValue = tables[0].Header.StaticCells[0].DefaultValue;
				string defaultValue2 = tables[0].Header.StaticCells[3].DefaultValue;
				string defaultValue3 = tables[0].Header.StaticCells[6].DefaultValue;
				string defaultValue4 = tables[0].Header.StaticCells[7].DefaultValue;
				array[0] = defaultValue;
				array[3] = defaultValue2;
				array[6] = defaultValue3;
				array[7] = defaultValue4;
				string[] array2 = new string[8];
				string defaultValue5 = tables[0].RowsInfo.Cells[0].DefaultValue;
				string defaultValue6 = tables[0].RowsInfo.Cells[6].DefaultValue;
				string defaultValue7 = tables[0].RowsInfo.Cells[7].DefaultValue;
				array2[0] = defaultValue5;
				array2[6] = defaultValue6;
				array2[7] = defaultValue7;
				HTKKCMCStaticOfTable hTKKCMCStaticOfTable = new HTKKCMCStaticOfTable();
				HTKKCMCRowInfo hTKKCMCRowInfo = new HTKKCMCRowInfo();
				int num = 0;
				for (int i = 0; i < tables.Count; i++)
				{
					hTKKCMCStaticOfTable = tables[i].Header;
					hTKKCMCRowInfo = tables[i].RowsInfo;
					num = tables[i].Location.DataLocationTemplateRowStart - 1;
					array[1] = sheetView.Cells[hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1].Text;
					array[2] = sheetView.Cells[hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1].Text;
					array[4] = ConvertHelpers.ConvertCellValueToText(sheetView, hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[4].CellTemplate.ColIndex - 1);
					array[5] = ConvertHelpers.ConvertCellValueToText(sheetView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
					bool flag = GeneralFunctions.HaveToValidateHeaderRow(header, array);
					bool flag2 = false;
					double num2 = 0.0;
					for (int j = 0; j < tables[i].RowsValues.Count; j++)
					{
						array2[1] = sheetView.Cells[num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Text;
						array2[2] = sheetView.Cells[num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1].Text;
						array2[3] = sheetView.Cells[num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1].Text;
						array2[4] = ConvertHelpers.ConvertCellValueToText(sheetView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
						array2[5] = ConvertHelpers.ConvertCellValueToText(sheetView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						array2[6] = sheetView.Cells[num + j, hTKKCMCRowInfo.Cells[6].CellTemplate.ColIndex - 1].Text;
						bool visible = sheetView.Cells[num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1].Row.Visible;
						if (GeneralFunctions.HaveToValidateRow(hTKKCMCRowInfo, array2) && visible)
						{
							flag2 = true;
						}
						else
						{
							if (!flag)
							{
								HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
								HTKKFile.RemoveValidateBackGround(sheetView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
								continue;
							}
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[1].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[6].CellTemplate.ColIndex - 1);
						}
						if (string.IsNullOrEmpty(array2[3]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[3].CellTemplate.ColIndex - 1);
						}
						if (!ValidateBase.CheckMaSoThue(array2[2]) && !string.IsNullOrEmpty(array2[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[2].CellTemplate.ColIndex - 1);
						}
						num2 += ConvertHelpers.ConvertToDouble(array2[4]);
						if ((array2[5] == hTKKCMCRowInfo.Cells[5].DefaultValue) | string.IsNullOrEmpty(array2[5]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidateBackGround(sheetView, num + j, hTKKCMCRowInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
					}
					if (!flag2 && !flag)
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
						continue;
					}
					if (num2 != 100.0)
					{
						for (int j = 0; j < tables[i].RowsValues.Count; j++)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_107")));
							value = 2;
						}
					}
					else
					{
						for (int j = 0; j < tables[i].RowsValues.Count; j++)
						{
							HTKKFile.RemoveValidate(sheetView, num + j, hTKKCMCRowInfo.Cells[4].CellTemplate.ColIndex - 1);
						}
					}
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[1].CellTemplate.ColIndex - 1);
					}
					if (!ValidateBase.CheckMaSoThue(array[2]) && !string.IsNullOrEmpty(array[2]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[2].CellTemplate.ColIndex - 1);
					}
					if ((array[5] == hTKKCMCStaticOfTable.StaticCells[5].DefaultValue) | string.IsNullOrEmpty(array[5]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.RowIndex - 1, hTKKCMCStaticOfTable.StaticCells[5].CellTemplate.ColIndex - 1);
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[14].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[14].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[14].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void checkPL3_1_ND20(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			if (winFpSpread.Sheets[16] == null || !winFpSpread.Sheets[16].Visible)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(16, out value);
			bool flag = false;
			SheetView sheetView = winFpSpread.Sheets[16];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[16];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (!sheetView.Visible)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[16];
			int num12 = ToKhaiInfo.SheetInfos[16].Tables[0].RowsValues.Count - 1;
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
			int num13 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				if (string.IsNullOrEmpty(array[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				}
				if (string.IsNullOrEmpty(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("O") - 1);
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("I") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				else
				{
					if (string.IsNullOrEmpty(array[13]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("I") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_122")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("I") - 1);
					}
					if (array[1] == "Việt Nam" && !string.IsNullOrEmpty(array[2]))
					{
						if (!ValidateBase.CheckMaSoThue(array[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("O") - 1);
						}
					}
				}
				if (!ConvertHelpers.ConvertToBool(array[3]) && !ConvertHelpers.ConvertToBool(array[4]) && !ConvertHelpers.ConvertToBool(array[5]) && !ConvertHelpers.ConvertToBool(array[6]) && !ConvertHelpers.ConvertToBool(array[7]) && !ConvertHelpers.ConvertToBool(array[8]) && !ConvertHelpers.ConvertToBool(array[9]) && !ConvertHelpers.ConvertToBool(array[10]) && !ConvertHelpers.ConvertToBool(array[11]) && !ConvertHelpers.ConvertToBool(array[12]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("W") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AA") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("U") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("V") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("W") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("X") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AA") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AC") - 1);
					HTKKFile.RemoveValidate(sheetView, num13 + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
				{
					HTKKFile.RemoveValidate(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
					HTKKFile.RemoveValidate(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
					HTKKFile.RemoveValidate(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					HTKKFile.SetNodeBackGroundPink(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
					value = 2;
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, 43 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				HTKKFile.RemoveValidate(sheetView, 44 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
				HTKKFile.RemoveValidate(sheetView, 45 + hTKKCMCTableInfo.RowsValues.Count, HTKKFile.ColumnLettersToIndex("N") - 1);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[16].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_II/ct1"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[16].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_II/ct2a"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[16].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_II/ct2c"));
			if (!ConvertHelpers.ConvertToBool(hTKKCMCCellInfo.Value))
			{
				for (int i = 1; i < 9; i++)
				{
					for (int num14 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num14 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num14++)
					{
						string[] array = hTKKCMCSheetInfo.Tables[i].RowsValues[num14 - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + 1].RowValue.ToArray();
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[i].RowsInfo, array))
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("AF") - 1);
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("AG") - 1);
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("AH") - 1);
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("AI") - 1);
							continue;
						}
						if (array[3].ToIntOrDefault() < 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num14, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_117")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("M") - 1);
						}
						if (array[7].ToIntOrDefault() > 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num14, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_118")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("Y") - 1);
						}
						if (array[1].ToIntOrDefault() > 0 && array[4].IsNullOrEmpty())
						{
							if (!hTKKCMCCellInfo2.Value.ToBoolOrDefault() && !hTKKCMCCellInfo3.Value.ToBoolOrDefault())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num14, HTKKFile.ColumnLettersToIndex("N") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("N") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("N") - 1);
						}
						if (array[5].ToIntOrDefault() > 0 && array[8].IsNullOrEmpty())
						{
							if (!hTKKCMCCellInfo2.Value.ToBoolOrDefault() && !hTKKCMCCellInfo3.Value.ToBoolOrDefault())
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num14, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("AB") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("AB") - 1);
						}
					}
				}
				bool flag2 = hTKKCMCSheetInfo.StaticCells[1].Value.ToBoolOrDefault();
				bool flag3 = hTKKCMCSheetInfo.StaticCells[2].Value.ToBoolOrDefault();
				bool flag4 = hTKKCMCSheetInfo.StaticCells[3].Value.ToBoolOrDefault();
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_2/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_4/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_6/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_7/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_9/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/Dong_1/ct06"));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/Dong_2/ct06"));
				HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, testColor: true);
				if (flag2)
				{
					value = CheckTongCot6MucIIIP1P3(value, sheetView, hTKKCMCSheetInfo);
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1, testColor: true);
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1, testColor: true);
				}
				if (flag3)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/khong"));
					HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_2/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_2/khong"));
					HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/co"));
					HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/khong"));
					if (!hTKKCMCCellInfo14.Value.ToBoolOrDefault() && !hTKKCMCCellInfo15.Value.ToBoolOrDefault() && !hTKKCMCCellInfo16.Value.ToBoolOrDefault() && !hTKKCMCCellInfo17.Value.ToBoolOrDefault() && !hTKKCMCCellInfo18.Value.ToBoolOrDefault() && !hTKKCMCCellInfo19.Value.ToBoolOrDefault())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo14.CellTemplate.RowIndex - 1, hTKKCMCCellInfo14.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo15.CellTemplate.RowIndex - 1, hTKKCMCCellInfo15.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo16.CellTemplate.RowIndex - 1, hTKKCMCCellInfo16.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, hTKKCMCCellInfo17.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo18.CellTemplate.RowIndex - 1, hTKKCMCCellInfo18.CellTemplate.ColIndex - 1, testColor: true);
						HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo19.CellTemplate.RowIndex - 1, hTKKCMCCellInfo19.CellTemplate.ColIndex - 1, testColor: true);
					}
				}
				for (int i = 9; i < 12; i++)
				{
					for (int num14 = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; num14 < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num14++)
					{
						string[] array = hTKKCMCSheetInfo.Tables[i].RowsValues[num14 - hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + 1].RowValue.ToArray();
						if (!GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[i].RowsInfo, array) && (!flag4 || i != 9))
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("D") - 1);
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("V") - 1);
						}
						else if (string.IsNullOrEmpty(sheetView.Cells[num14, HTKKFile.ColumnLettersToIndex("D") - 1].Text) && !sheetView.Cells[num14, HTKKFile.ColumnLettersToIndex("D") - 1].Locked)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num14, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num14, HTKKFile.ColumnLettersToIndex("D") - 1);
						}
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[16].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[16].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[16].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static int CheckTongCot6MucIIIP1P3(int iwarningColor, SheetView sView, HTKKCMCSheetInfo sheetInfo)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_1/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo2 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/Dong_1/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo3 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_1/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo4 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_4/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo5 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_6/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo6 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_7/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo7 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_1/Dong_9/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo8 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/Dong_1/ct06"));
		HTKKCMCCellInfo hTKKCMCCellInfo9 = sheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_IV/IV_3/Dong_2/ct06"));
		bool flag = false;
		if (sView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m || sView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToNumber() > 0m)
		{
			flag = true;
		}
		if (!flag)
		{
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			HTKKFile.SetNodeBackGroundPink(sView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_88")));
			iwarningColor = 2;
		}
		else
		{
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1, testColor: true);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1, testColor: true);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1, testColor: true);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1, testColor: true);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1, testColor: true);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1, testColor: true);
			HTKKFile.RemoveValidate(sView, hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1, testColor: true);
		}
		return iwarningColor;
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year - 1;
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("//tieuMucHachToan");
			if ((val != null) & !string.IsNullOrEmpty(PeriodInfo.TieuMucHachToanVal))
			{
				val.InnerText = PeriodInfo.TieuMucHachToanVal;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_12(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 9 && idxCol == 1)
			{
				ShowBangKe("Bangke_03_7.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_16(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			winFpSpread.SuspendLayout();
			SheetView sheetView = winFpSpread.Sheets[16];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[16];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0])
			{
				AnMucIII_IV(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 38 + count && idxCol == GeneralFunctions.ArrayCellId("N_40")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
			}
			if (idxRow == 40 + count && idxCol == GeneralFunctions.ArrayCellId("N_42")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				AnMucIII_IV(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 41 + count && idxCol == GeneralFunctions.ArrayCellId("N_43")[0] && ConvertHelpers.ConvertToBool(sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = false;
				sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
				sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
				sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				AnMucIII_IV(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow == 42 + count && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
				{
					sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
					sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
					sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
					AnMucIII_IV(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
				}
				else
				{
					sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value = false;
					sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value = false;
					sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value = false;
				}
			}
			if ((idxRow == 43 + count || idxRow == 44 + count || idxRow == 45 + count) && idxCol == GeneralFunctions.ArrayCellId("N_44")[0] && (ConvertHelpers.ConvertToBool(sheetView.Cells[43 + count, GeneralFunctions.ArrayCellId("N_45")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[44 + count, GeneralFunctions.ArrayCellId("N_46")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[45 + count, GeneralFunctions.ArrayCellId("N_47")[0]].Value)))
			{
				sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value = false;
				sheetView.Cells[41 + count, GeneralFunctions.ArrayCellId("N_43")[0]].Value = false;
				sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value = false;
				sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value = true;
				AnMucIII_IV(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count);
			}
			if (idxRow >= 38 + count && idxRow <= 45 + count && idxCol == GeneralFunctions.ArrayCellId("N_44")[0])
			{
				XoaDLIII_IV(38 + count, GeneralFunctions.ArrayCellId("N_40")[0], sheetView, count, ToKhaiInfo);
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1/ct12a"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2/ct12a"));
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = $"AF{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AF{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AG") - 1].Formula = $"AG{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AG{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AH") - 1].Formula = $"AH{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AH{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AI") - 1].Formula = $"AI{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AI{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					for (int num12 = 1; num12 < 9; num12++)
					{
						for (int num13 = hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1; num13 < hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[num12].RowsValues.Count; num13++)
						{
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = true;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AE") - 1].Formula = "";
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AF") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = true;
						}
					}
					bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && idxRow == 40 + count)
					{
						for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							if (num12 == 109 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
							else if (num12 == 111 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
							else if (num12 >= 114 + num9 && num12 < 119 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							}
							else if (num12 == 119 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]+V[4]-V[5]".Replace("[1]", (num12 - 7).ToString()).Replace("[2]", (num12 - 6).ToString()).Replace("[3]", (num12 - 5).ToString())
									.Replace("[4]", (num12 - 4).ToString())
									.Replace("[5]", (num12 - 2).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
							else if (num12 == 120 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]".Replace("[1]", (num12 - 8).ToString()).Replace("[2]", (num12 - 7).ToString()).Replace("[3]", (num12 - 6).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
							else if (num12 == 121 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]+V[2]+V[3]".Replace("[1]", (num12 - 1).ToString()).Replace("[2]", (num12 - 2).ToString()).Replace("[3]", (num12 - 3).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
							else if (num12 == 122 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "IF(V[1]<>0,ROUND(V[2]/V[1]*100,3),0)".Replace("[1]", num12.ToString()).Replace("[2]", (num12 - 4).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
							else
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							}
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
					{
						for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							if (num12 == 109 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								continue;
							}
							if (num12 == 111 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								continue;
							}
							if (num12 >= 114 + num8 && num12 < 119 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
								continue;
							}
							if (num12 == 119 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]+V[4]-V[5]".Replace("[1]", (num12 - 7).ToString()).Replace("[2]", (num12 - 6).ToString()).Replace("[3]", (num12 - 5).ToString())
									.Replace("[4]", (num12 - 4).ToString())
									.Replace("[5]", (num12 - 2).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								continue;
							}
							if (num12 == 120 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]".Replace("[1]", (num12 - 8).ToString()).Replace("[2]", (num12 - 7).ToString()).Replace("[3]", (num12 - 6).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								continue;
							}
							if (num12 == 121 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]+V[2]+V[3]".Replace("[1]", (num12 - 1).ToString()).Replace("[2]", (num12 - 2).ToString()).Replace("[3]", (num12 - 3).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								continue;
							}
							if (num12 == 122 + num8)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "IF(V[1]<>0,ROUND(V[2]/V[1]*100,3),0)".Replace("[1]", num12.ToString()).Replace("[2]", (num12 - 4).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								continue;
							}
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
					}
					if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && idxRow == 40 + count)
					{
						for (int num12 = 166 + num10; num12 < 200 + num10; num12++)
						{
							if (num12 != 166 + num10 && num12 != 176 + num10 && num12 != 186 + num10 && num12 != 191 + num10 && num12 != 194 + num10 && num12 != 195 + num10 && num12 != 196 + num10 && num12 != 197 + num10 && num12 != 198 + num10 && num12 != 199 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							}
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
					}
					if (!ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
					{
						for (int num12 = 166 + num10; num12 < 200 + num10; num12++)
						{
							if (num12 != 166 + num10 && num12 != 176 + num10 && num12 != 186 + num10 && num12 != 191 + num10 && num12 != 194 + num10 && num12 != 195 + num10 && num12 != 196 + num10 && num12 != 197 + num10 && num12 != 198 + num10 && num12 != 199 + num10)
							{
								if (num12 == 179 + num10 || num12 == 182 + num10 || num12 == 187 + num10 || num12 == 188 + num10)
								{
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
								}
								else
								{
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (num12 + 1).ToString());
									sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
								}
							}
						}
						for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						}
					}
				}
				else
				{
					sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
					sheetView.Cells[59 + count, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
					HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1/ct12a"));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2/ct12a"));
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AF") - 1].Formula = $"AF{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AF{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AG") - 1].Formula = $"AG{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AG{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AH") - 1].Formula = $"AH{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AH{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					sheetView.Cells[60 + count, HTKKFile.ColumnLettersToIndex("AI") - 1].Formula = $"AI{hTKKCMCCellInfo.CellTemplate.RowIndex.ToStringOrDefault()}+AI{hTKKCMCCellInfo2.CellTemplate.RowIndex.ToStringOrDefault()}";
					for (int num12 = 1; num12 < 9; num12++)
					{
						for (int num13 = hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1; num13 < hTKKCMCSheetInfo.Tables[num12].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[num12].RowsValues.Count; num13++)
						{
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = string.Format("J{0}-H{0}", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = string.Format("V{0}-R{0}", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AE") - 1].Formula = string.Format("ABS(M{0})+ABS(Y{0})", (num13 + 1).ToString());
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AF") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
							sheetView.Cells[num13, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
						}
					}
					for (int num12 = 106 + num8; num12 < 123 + num8; num12++)
					{
						if (num12 == 109 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 2).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 == 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (num12 - 1).ToString()).Replace("@@@", num12.ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 >= 114 + num8 && num12 < 119 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						else if (num12 == 119 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]+V[4]-V[5]".Replace("[1]", (num12 - 7).ToString()).Replace("[2]", (num12 - 6).ToString()).Replace("[3]", (num12 - 5).ToString())
								.Replace("[4]", (num12 - 4).ToString())
								.Replace("[5]", (num12 - 2).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 == 120 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M[1]-M[2]-M[3]".Replace("[1]", (num12 - 8).ToString()).Replace("[2]", (num12 - 7).ToString()).Replace("[3]", (num12 - 6).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N[1]-N[2]-N[3]".Replace("[1]", (num12 - 8).ToString()).Replace("[2]", (num12 - 7).ToString()).Replace("[3]", (num12 - 6).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R[1]-R[2]-R[3]".Replace("[1]", (num12 - 8).ToString()).Replace("[2]", (num12 - 7).ToString()).Replace("[3]", (num12 - 6).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]".Replace("[1]", (num12 - 8).ToString()).Replace("[2]", (num12 - 7).ToString()).Replace("[3]", (num12 - 6).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 == 121 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]+V[2]+V[3]".Replace("[1]", (num12 - 1).ToString()).Replace("[2]", (num12 - 2).ToString()).Replace("[3]", (num12 - 3).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (num12 == 122 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "IF(V[1]<>0,ROUND(V[2]/V[1]*100,3),0)".Replace("[1]", num12.ToString()).Replace("[2]", (num12 - 4).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (num12 + 1).ToString());
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					for (int num12 = 166 + num10; num12 < 200 + num10; num12++)
					{
						if (num12 != 166 + num10 && num12 != 176 + num10 && num12 != 186 + num10 && num12 != 191 + num10 && num12 != 194 + num10 && num12 != 195 + num10 && num12 != 196 + num10 && num12 != 197 + num10 && num12 != 198 + num10 && num12 != 199 + num10)
						{
							if (num12 == 179 + num10 || num12 == 182 + num10 || num12 == 187 + num10 || num12 == 188 + num10)
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							}
							else
							{
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (num12 + 1).ToString());
								sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Value = "";
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
			}
			if (idxRow == 100 + num8 && (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] || idxCol == GeneralFunctions.ArrayCellId("U_102")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value = false;
				}
			}
			if (idxRow == 130 + num9 && (idxCol == GeneralFunctions.ArrayCellId("X_132")[0] || idxCol == GeneralFunctions.ArrayCellId("U_132")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_132")[0] && sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("U_132")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_132")[0] && sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("U_132")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value = false;
				}
			}
			if (idxRow == 160 + num10 && (idxCol == GeneralFunctions.ArrayCellId("X_162")[0] || idxCol == GeneralFunctions.ArrayCellId("U_162")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_162")[0] && sheetView.Cells[160 + num10, GeneralFunctions.ArrayCellId("X_162")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[160 + num10, GeneralFunctions.ArrayCellId("U_162")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_162")[0] && sheetView.Cells[160 + num10, GeneralFunctions.ArrayCellId("U_162")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[160 + num10, GeneralFunctions.ArrayCellId("X_162")[0]].Value = false;
				}
			}
			if (idxRow == 100 + num8 && (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] || idxCol == GeneralFunctions.ArrayCellId("U_102")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value = false;
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else
				{
					for (int num12 = 106 + num8; num12 < 114 + num8; num12++)
					{
						if (num12 != 109 + num8 && num12 != 111 + num8)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (idxRow == 130 + num9 && (idxCol == GeneralFunctions.ArrayCellId("X_132")[0] || idxCol == GeneralFunctions.ArrayCellId("U_132")[0]))
			{
				if (idxCol == GeneralFunctions.ArrayCellId("X_132")[0] && sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("U_132")[0]].Value = false;
				}
				if (idxCol == GeneralFunctions.ArrayCellId("U_132")[0] && sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("U_132")[0]].Value.ToBoolOrDefault())
				{
					sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value = false;
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value))
				{
					for (int num12 = 136 + num9; num12 < 151 + num9; num12++)
					{
						if (num12 != 138 + num9 && num12 != 141 + num9 && num12 != 147 + num9)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				else
				{
					for (int num12 = 136 + num9; num12 < 151 + num9; num12++)
					{
						if (num12 != 138 + num9 && num12 != 141 + num9 && num12 != 147 + num9)
						{
							sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						}
					}
					for (int num12 = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; num12++)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (idxRow != 160 + num10 || (idxCol != GeneralFunctions.ArrayCellId("X_162")[0] && idxCol != GeneralFunctions.ArrayCellId("U_162")[0]))
			{
				return;
			}
			if (idxCol == GeneralFunctions.ArrayCellId("X_162")[0] && sheetView.Cells[160 + num8, GeneralFunctions.ArrayCellId("X_162")[0]].Value.ToBoolOrDefault())
			{
				sheetView.Cells[160 + num8, GeneralFunctions.ArrayCellId("U_162")[0]].Value = false;
			}
			if (idxCol == GeneralFunctions.ArrayCellId("U_102")[0] && sheetView.Cells[160 + num8, GeneralFunctions.ArrayCellId("U_162")[0]].Value.ToBoolOrDefault())
			{
				sheetView.Cells[160 + num8, GeneralFunctions.ArrayCellId("X_162")[0]].Value = false;
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[160 + num10, GeneralFunctions.ArrayCellId("X_162")[0]].Value))
			{
				for (int num12 = 167 + num10; num12 < 194 + num10; num12++)
				{
					if (num12 != 176 + num10 && num12 != 179 + num10 && num12 != 182 + num10 && num12 != 186 + num10 && num12 != 187 + num10 && num12 != 188 + num10 && num12 != 191 + num10)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				return;
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int num12 = 167 + num10; num12 < 194 + num10; num12++)
				{
					if (num12 != 176 + num10 && num12 != 179 + num10 && num12 != 182 + num10 && num12 != 186 + num10 && num12 != 187 + num10 && num12 != 188 + num10 && num12 != 191 + num10)
					{
						sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				return;
			}
			for (int num12 = 167 + num10; num12 < 194 + num10; num12++)
			{
				if (num12 != 176 + num10 && num12 != 179 + num10 && num12 != 182 + num10 && num12 != 186 + num10 && num12 != 187 + num10 && num12 != 188 + num10 && num12 != 191 + num10)
				{
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			for (int num12 = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; num12 < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; num12++)
			{
				sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
				sheetView.Cells[num12, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		finally
		{
			winFpSpread.ResumeLayout();
		}
	}

	private static void AnMucIII_IV(int idxRow, int idxCol, SheetView sView, int rowCountTable1)
	{
		if (ConvertHelpers.ConvertToBool(sView.Cells[idxRow, idxCol].Value))
		{
			for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
			{
				if (!(sView.Cells[i, HTKKFile.ColumnLettersToIndex("AR") - 1].Text == "x"))
				{
					continue;
				}
				for (int j = 0; j < sView.NonEmptyColumnCount; j++)
				{
					Cell cell = sView.Cells[i, j];
					if (!cell.Locked && cell.CellType != null)
					{
						if (cell.CellType.ToString() == "NumberCellType")
						{
							cell.Value = 0;
						}
						else if (cell.CellType.ToString() == "CheckBoxCellType")
						{
							cell.Value = false;
						}
						else
						{
							cell.Value = "";
						}
					}
				}
				sView.Rows[i].Visible = false;
			}
			return;
		}
		for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
		{
			if (sView.Cells[i, HTKKFile.ColumnLettersToIndex("AR") - 1].Text == "x")
			{
				sView.Rows[i].Visible = true;
			}
		}
	}

	private static void XoaDLIII_IV(int idxRow, int idxCol, SheetView sView, int rowCountTable1, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		for (int i = 51 + rowCountTable1; i < sView.NonEmptyRowCount; i++)
		{
			if (!(sView.Cells[i, HTKKFile.ColumnLettersToIndex("AR") - 1].Text == "x"))
			{
				continue;
			}
			for (int j = 0; j < sView.NonEmptyColumnCount; j++)
			{
				Cell cell = sView.Cells[i, j];
				if (cell.Locked && !cell.Note.IsNulOrEmpty())
				{
					HTKKFile.RemoveValidateBackGround(sView, i, j);
				}
				else if (!cell.Locked && !cell.Note.IsNulOrEmpty())
				{
					HTKKFile.RemoveValidate(sView, i, j);
				}
				if (!cell.Locked && cell.CellType != null)
				{
					if (cell.CellType.ToString() == "NumberCellType")
					{
						cell.Value = 0;
					}
					else if (cell.CellType.ToString() == "CheckBoxCellType")
					{
						cell.Value = false;
					}
					else
					{
						cell.Value = "";
					}
				}
			}
		}
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[16];
		int num = rowCountTable1 + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
		int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
		int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
		int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
		int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
		int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
		int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
		int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
		int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
		int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
		int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
		if (ConvertHelpers.ConvertToBool(sView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
		{
			sView.Cells[100 + num8, GeneralFunctions.ArrayCellId("U_102")[0]].Value = false;
			for (int i = 106 + num8; i < 114 + num8; i++)
			{
				if (i != 109 + num8 && i != 111 + num8)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
			{
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
			}
		}
		else
		{
			if (!ConvertHelpers.ConvertToBool(sView.Cells[40 + rowCountTable1, GeneralFunctions.ArrayCellId("N_42")[0]].Value) && !ConvertHelpers.ConvertToBool(sView.Cells[42 + rowCountTable1, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
						sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
			}
			if (ConvertHelpers.ConvertToBool(sView.Cells[42 + rowCountTable1, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
			{
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			else
			{
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
		}
		if (ConvertHelpers.ConvertToBool(sView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value))
		{
			for (int i = 136 + num9; i < 151 + num9; i++)
			{
				if (i != 138 + num9 && i != 141 + num9 && i != 147 + num9)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
			{
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
			}
		}
		else
		{
			for (int i = 136 + num9; i < 151 + num9; i++)
			{
				if (i != 138 + num9 && i != 141 + num9 && i != 147 + num9)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
			{
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
			}
		}
		if (ConvertHelpers.ConvertToBool(sView.Cells[160 + num10, GeneralFunctions.ArrayCellId("X_162")[0]].Value))
		{
			for (int i = 167 + num10; i < 194 + num10; i++)
			{
				if (i != 176 + num10 && i != 179 + num10 && i != 182 + num10 && i != 186 + num10 && i != 187 + num10 && i != 188 + num10 && i != 191 + num10)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
			{
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
			}
			return;
		}
		if (ConvertHelpers.ConvertToBool(sView.Cells[40 + rowCountTable1, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
		{
			for (int i = 167 + num10; i < 194 + num10; i++)
			{
				if (i != 176 + num10 && i != 179 + num10 && i != 182 + num10 && i != 186 + num10 && i != 187 + num10 && i != 188 + num10 && i != 191 + num10)
				{
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
			{
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
			}
			return;
		}
		for (int i = 167 + num10; i < 194 + num10; i++)
		{
			if (i != 176 + num10 && i != 179 + num10 && i != 182 + num10 && i != 186 + num10 && i != 187 + num10 && i != 188 + num10 && i != 191 + num10)
			{
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
				sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
			}
		}
		for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
		{
			sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
			sView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string[] array = new string[14]
			{
				"N95", "J95", "O56", "O57", "O58", "O59", "O61", "O62", "O65", "O69",
				"O70", "O72", "O73", "O74"
			};
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[96, HTKKFile.ColumnLettersToIndex("E") - 1].Formula = "IF(O80>0,ROUND(O80*E95*0.03/100,0),0)";
					break;
				}
			}
			int[] array2 = GeneralFunctions.ArrayCellId("E_95");
			string value = GeneralFunctions.GetValue(winFpSpread.Sheets[0], "J_95");
			string value2 = GeneralFunctions.GetValue(winFpSpread.Sheets[0], "N_95");
			if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(value2))
			{
				TimeSpan timeSpan = DateTime.ParseExact(value2.Replace("/", ""), "ddMMyyyy", CultureInfo.CurrentCulture) - DateTime.ParseExact(value.Replace("/", ""), "ddMMyyyy", CultureInfo.CurrentCulture);
				winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Value = timeSpan.TotalDays + 1.0;
			}
			if (text == "O46")
			{
				winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
			}
			array = new string[10] { "O31", "O34", "O35", "O36", "O37", "O38", "O39", "O41", "O42", "O43" };
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
					break;
				}
			}
			if (text.Equals("O67") || text.Equals("I89"))
			{
				winFpSpread.Sheets[0].Cells[90, HTKKFile.ColumnLettersToIndex("I") - 1].ResetFormula();
				winFpSpread.Sheets[0].Cells[90, HTKKFile.ColumnLettersToIndex("I") - 1].Formula = "O67-I89";
			}
			double num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[0], "O_53"));
			if (num < 0.0)
			{
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_54")[1], GeneralFunctions.ArrayCellId("O_54")[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_54")[1], GeneralFunctions.ArrayCellId("O_54")[0]].Locked = true;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_56")[1], GeneralFunctions.ArrayCellId("O_56")[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_56")[1], GeneralFunctions.ArrayCellId("O_56")[0]].Locked = true;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_57")[1], GeneralFunctions.ArrayCellId("O_57")[0]].Value = 0;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_57")[1], GeneralFunctions.ArrayCellId("O_57")[0]].Locked = true;
			}
			else
			{
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_54")[1], GeneralFunctions.ArrayCellId("O_54")[0]].Locked = false;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_56")[1], GeneralFunctions.ArrayCellId("O_56")[0]].Locked = false;
				winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId("O_57")[1], GeneralFunctions.ArrayCellId("O_57")[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string[] array = new string[12]
			{
				"E38", "E39", "E35", "E30", "E32", "E33", "E34", "E23", "E26", "E27",
				"E28", "E29"
			};
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
					winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string[] array = new string[12]
			{
				"E35", "E36", "E37", "E29", "E30", "E31", "E32", "E33", "E26", "E27",
				"E23", "E24"
			};
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
					winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string[] array = new string[23]
			{
				"E47", "E48", "E34", "E35", "E36", "E37", "E38", "E39", "E40", "E41",
				"E42", "E43", "E44", "E45", "E24", "E25", "E26", "E27", "E28", "E29",
				"E30", "E31", "E32"
			};
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[30, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[1].SheetName + "'!E41+'" + winFpSpread.Sheets[2].SheetName + "'!E38+'" + winFpSpread.Sheets[3].SheetName + "'!E50";
					winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string[] array = new string[5] { "M30", "M29", "M28", "M27", "M26" };
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[0, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[4].SheetName + "'!M32";
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[6];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[6];
			string[] array = new string[4] { "U50", "U49", "U47", "Q37" };
			List<int> list = new List<int>();
			for (int i = 0; i < array.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			bool flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					double num = SumChiTieuCellTable(sheetView, hTKKCMCSheetInfo, list[0]);
					sheetView.Cells[3, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
					sheetView.Cells[3, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num;
					winFpSpread.Sheets[0].Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U4+'" + winFpSpread.Sheets[7].SheetName + "'!U7";
					break;
				}
			}
			string[] array2 = new string[5] { "U52", "U53", "U54", "U55", "U56" };
			list.Clear();
			for (int i = 0; i < array2.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array2[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					double num2 = SumChiTieuCellTable(sheetView, hTKKCMCSheetInfo, list[4]);
					sheetView.Cells[8, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
					sheetView.Cells[8, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num2;
					winFpSpread.Sheets[0].Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
					break;
				}
			}
			string[] array3 = new string[3] { "U52", "U53", "U54" };
			list.Clear();
			for (int i = 0; i < array3.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array3[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count - 1; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					int rowIndex = hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[list.Count - 1]].CellTemplate.RowIndex;
					int num3 = HTKKFile.ColumnLettersToIndex("U");
					winFpSpread.Sheets[6].Cells[rowIndex - 1, num3 - 1].Formula = "ROUND(U" + hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[0]].CellTemplate.RowIndex + "*U" + hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[1]].CellTemplate.RowIndex + "/100,0)";
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_7(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[7];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[7];
			string[] array = new string[9] { "U60", "U58", "U59", "U56", "Q43", "U53", "U54", "U53", "U55" };
			List<int> list = new List<int>();
			for (int i = 0; i < array.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			bool flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					double num = SumChiTieuCellTable(sheetView, hTKKCMCSheetInfo, list[0]);
					sheetView.Cells[6, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
					sheetView.Cells[6, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num;
					winFpSpread.Sheets[0].Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U4+'" + winFpSpread.Sheets[7].SheetName + "'!U7";
					break;
				}
			}
			string[] array2 = new string[9] { "U65", "U64", "U63", "U62", "U56", "U53", "U54", "U53", "U55" };
			list.Clear();
			for (int i = 0; i < array2.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array2[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					double num2 = SumChiTieuCellTable(sheetView, hTKKCMCSheetInfo, list[0]);
					sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
					sheetView.Cells[10, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num2;
					winFpSpread.Sheets[0].Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
					break;
				}
			}
			string text = "ROUND(IF(U{0} = 0, 0, (U{1} / U{0}) * U{2}), 0)";
			string[] array3 = new string[3] { "U53", "U54", "U55" };
			list.Clear();
			for (int i = 0; i < array3.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array3[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					winFpSpread.Sheets[7].Cells[hTKKCMCSheetInfo.Tables[i].Header.StaticCells[21].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("U") - 1].ResetFormula();
					winFpSpread.Sheets[7].Cells[hTKKCMCSheetInfo.Tables[i].Header.StaticCells[21].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = string.Format("ROUND(IF(U{0} = 0, 0, (U{1} / U{0}) *U{2}), 0)", hTKKCMCSheetInfo.Tables[i].Header.StaticCells[21].CellTemplate.RowIndex - 3, hTKKCMCSheetInfo.Tables[i].Header.StaticCells[21].CellTemplate.RowIndex - 2, hTKKCMCSheetInfo.Tables[i].Header.StaticCells[21].CellTemplate.RowIndex - 1);
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string[] array = new string[1] { "X43" };
			List<int> list = new List<int>();
			SheetView sheetView = winFpSpread.Sheets[8];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[8];
			for (int i = 0; i < array.Length; i++)
			{
				for (int j = 0; j <= hTKKCMCSheetInfo.Tables[0].Header.StaticCells.Count; j++)
				{
					if (array[i] == HTKKFile.ColumnIndexToLetters(hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.ColIndex) + hTKKCMCSheetInfo.Tables[0].Header.StaticCells[j].CellTemplate.RowIndex)
					{
						list.Add(j);
						break;
					}
				}
			}
			bool flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				for (int j = 0; j < list.Count; j++)
				{
					if (sheetView.ActiveCell.Column.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.ColIndex && sheetView.ActiveCell.Row.Index + 1 == hTKKCMCSheetInfo.Tables[i].Header.StaticCells[list[j]].CellTemplate.RowIndex)
					{
						flag = true;
						break;
					}
				}
				if (!flag)
				{
					continue;
				}
				double num = SumChiTieuCellTable(sheetView, hTKKCMCSheetInfo, list[0]);
				sheetView.Cells[3, HTKKFile.ColumnLettersToIndex("X") - 1].ResetFormula();
				sheetView.Cells[3, HTKKFile.ColumnLettersToIndex("X") - 1].Value = num;
				winFpSpread.Sheets[0].Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[6].SheetName + "'!U9+'" + winFpSpread.Sheets[7].SheetName + "'!U11+'" + winFpSpread.Sheets[8].SheetName + "'!X4";
				if (ToKhaiInfo.SheetInfos[13] != null)
				{
					List<HTKKCMCRowValue> rowsValues = ToKhaiInfo.SheetInfos[13].Tables[0].RowsValues;
					for (int k = 0; k < rowsValues.Count; k++)
					{
						int num2 = ToKhaiInfo.SheetInfos[13].Tables[0].Location.DataLocationTemplateRowStart + k;
						winFpSpread.Sheets[13].Cells[num2 - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = $"ROUND($K$29*L{num2}/100,0)";
					}
				}
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_9(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if (HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) == "Z")
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[9].Tables[0];
				long num = long.Parse(winFpSpread.ActiveSheet.ActiveCell.Row.Index.ToString()) + 1;
				if ((hTKKCMCTableInfo.Location.DataLocationTemplateRowStart <= num) & (num <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1))
				{
					winFpSpread.Sheets[0].Cells[36, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[9].SheetName + "'!Z19";
					winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[9].SheetName + "'!Z19+'" + winFpSpread.Sheets[11].SheetName + "'!R18";
				}
			}
			if (HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) == "AR")
			{
				winFpSpread.Sheets[0].Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[9].SheetName + "'!AR19";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_10(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			double num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetObjectValue(winFpSpread.Sheets[10], "E_29"));
			if (num < 0.0)
			{
				winFpSpread.Sheets[10].Cells[29, 4].Value = 0;
				winFpSpread.Sheets[10].Cells[29, 4].Locked = true;
			}
			else
			{
				winFpSpread.Sheets[10].Cells[29, 4].Locked = false;
			}
			double num2 = (double)GeneralFunctions.GetObjectValue(winFpSpread.Sheets[10], "E_31");
			if (num2 > 0.0)
			{
				winFpSpread.ActiveSheet.Protect = true;
				winFpSpread.Sheets[10].Cells[31, 4].Locked = true;
			}
			else if (num2 <= 0.0)
			{
				winFpSpread.ActiveSheet.Protect = true;
				winFpSpread.Sheets[10].Cells[31, 4].Locked = false;
			}
			if (num2 < 0.0)
			{
				winFpSpread.ActiveSheet.Protect = true;
				winFpSpread.Sheets[10].Cells[32, 4].Value = 0;
				winFpSpread.Sheets[10].Cells[32, 4].Locked = true;
			}
			else
			{
				winFpSpread.ActiveSheet.Protect = true;
				winFpSpread.Sheets[10].Cells[32, 4].Locked = false;
			}
			double num3 = (double)GeneralFunctions.GetObjectValue(winFpSpread.Sheets[10], "E_32");
			if (num2 < 0.0 && num3 > 0.0)
			{
				winFpSpread.ActiveSheet.Protect = true;
				winFpSpread.Sheets[10].Cells[33, 4].Value = 0;
				winFpSpread.Sheets[10].Cells[33, 4].Locked = true;
			}
			else
			{
				winFpSpread.ActiveSheet.Protect = true;
				winFpSpread.Sheets[10].Cells[33, 4].Locked = false;
			}
			string[] array = new string[8] { "E21", "E22", "E23", "E24", "E25", "E26", "E27", "E28" };
			string text = HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) + (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[10].SheetName + "'!E29";
					winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "O44-O46";
					break;
				}
			}
			string formula = "IF(ROUND(E34*E35/100,0)>=0,ROUND(E34*E35/100,0),0)";
			int[] array2 = GeneralFunctions.ArrayCellId("E_36");
			array = new string[15]
			{
				"E31", "E33", "E34", "E35", "E36", "E29", "E30", "E21", "E22", "E23",
				"E24", "E25", "E26", "E27", "E28"
			};
			for (int i = 0; i < array.Length; i++)
			{
				if (text == array[i])
				{
					winFpSpread.Sheets[0].Cells[68, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[10].SheetName + "'!E36";
					if (text != "E36")
					{
						winFpSpread.Sheets[10].Cells[array2[1], array2[0]].ResetFormula();
						winFpSpread.Sheets[10].Cells[array2[1], array2[0]].Formula = formula;
					}
					break;
				}
			}
			if (text == "E32")
			{
				winFpSpread.Sheets[0].Cells[51, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[10].SheetName + "'!E32";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_11(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if (HTKKFile.ColumnIndexToLetters(winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1) == "O")
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[11].Tables[0];
				long num = long.Parse(winFpSpread.ActiveSheet.ActiveCell.Row.Index.ToString()) + 1;
				if ((hTKKCMCTableInfo.Location.DataLocationTemplateRowStart <= num) & (num <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1))
				{
					winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[11].SheetName + "'!O18";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_12(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int num = (int)iRow;
			int num2 = (int)iCol;
			SheetView sheetView = winFpSpread.Sheets[12];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[12];
			int num3 = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart + -4;
			int num4 = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart + -3;
			int num5 = hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart + -2;
			int num6 = hTKKCMCSheetInfo.Tables[2].Location.DataLocationTemplateRowStart + -2;
			int num7 = hTKKCMCSheetInfo.Tables[3].Location.DataLocationTemplateRowStart + -3;
			int num8 = hTKKCMCSheetInfo.Tables[3].Location.DataLocationTemplateRowStart + -2;
			int num9 = hTKKCMCSheetInfo.Tables[4].Location.DataLocationTemplateRowStart + -2;
			int num10 = hTKKCMCSheetInfo.Tables[5].Location.DataLocationTemplateRowStart + -2;
			int num11 = hTKKCMCSheetInfo.Tables[6].Location.DataLocationTemplateRowStart + -3;
			int num12 = hTKKCMCSheetInfo.Tables[6].Location.DataLocationTemplateRowStart + -2;
			int num13 = hTKKCMCSheetInfo.Tables[7].Location.DataLocationTemplateRowStart + -2;
			int num14 = hTKKCMCSheetInfo.Tables[8].Location.DataLocationTemplateRowStart + -2;
			int num15 = HTKKFile.ColumnLettersToIndex("H");
			int num16 = HTKKFile.ColumnLettersToIndex("J");
			int num17 = HTKKFile.ColumnLettersToIndex("R");
			int num18 = HTKKFile.ColumnLettersToIndex("V");
			int tableIndex = hTKKCMCSheetInfo.GetTableIndex(num, num2);
			if (!(iCol == num15 || iCol == num16 || iCol == num17 || iCol == num18))
			{
				return;
			}
			if (tableIndex < 0)
			{
				if (iRow == num4 || iRow == num7)
				{
					sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				}
				if (iRow == num5 || iRow == num6)
				{
					sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
					sheetView.Cells[num4 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num5 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num6;
				}
				if (iRow == num8 || iRow == num9 || iRow == num10 || iRow == num11 || iRow == num14)
				{
					sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
					sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				}
				if (iRow == num12 || iRow == num13)
				{
					sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
					sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
					sheetView.Cells[num11 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num12 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num13;
				}
				return;
			}
			switch (tableIndex)
			{
			case 1:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num4 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num5 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num6;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num5 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				break;
			case 2:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num4 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num5 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num6;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num6 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				break;
			case 3:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num8 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				break;
			case 4:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num9 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				break;
			case 5:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num10 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				break;
			case 6:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num12 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				sheetView.Cells[num11 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num12 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num13;
				break;
			case 7:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num13 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				sheetView.Cells[num11 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num12 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num13;
				break;
			case 8:
				sheetView.Cells[num3 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num4 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num7;
				sheetView.Cells[num7 - 1, num2 - 1].Formula = HTKKFile.ColumnIndexToLetters(num2) + num8 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num9 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num10 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num11 + "+" + HTKKFile.ColumnIndexToLetters(num2) + num14;
				if (num >= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart && num <= hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count)
				{
					sheetView.Cells[num14 - 1, num2 - 1].Formula = "SUM(" + HTKKFile.ColumnIndexToLetters(num2) + hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + ":" + HTKKFile.ColumnIndexToLetters(num2) + (hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count) + ")";
				}
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_13(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iRow == 68 && iCol == HTKKFile.ColumnLettersToIndex("O"))
			{
				double num = (double)winFpSpread.Sheets[0].Cells[67, 14].Value;
				winFpSpread.Sheets[13].Cells[28, HTKKFile.ColumnLettersToIndex("K") - 1].Value = num;
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("L"))
			{
				winFpSpread.Sheets[13].Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Formula = $"ROUND($K$29*L{iRow}/100,0)";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_14(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("U"))
			{
				double num = 0.0;
				bool flag = false;
				int num2 = -1;
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					if (iRow == hTKKCMCTableInfo.Header.StaticCells[5].CellTemplate.RowIndex)
					{
						num2 = i;
						flag = true;
					}
					num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
				}
				if (flag)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_02_1_TD_TNDN/BangPBoThueTNDN/ChiTietBangPBoThueTNDN/soThueTNDNPNop"));
					winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2];
					int num3 = ((ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count - 1 > num2) ? (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2 + 1].Location.DataLocationTemplateRowStart - 1) : (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].RowsValues.Count));
					for (int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num4 < num3; num4++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num6 * (num5 / 100.0);
					}
				}
			}
			else
			{
				if (iCol != HTKKFile.ColumnLettersToIndex("R"))
				{
					return;
				}
				for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
				{
					double num7 = 0.0;
					int num2 = -1;
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
					int num3 = ((ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count - 1 > i) ? ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i + 1].Location.DataLocationTemplateRowStart : (ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i].RowsValues.Count));
					for (int num4 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart; num4 < num3; num4++)
					{
						double num5 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num4 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text.Replace(",", "."));
						double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
						num7 += num5;
						if (iRow == num4)
						{
							num2 = i;
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = num6 * (num5 / 100.0);
						}
					}
					if (num2 > -1)
					{
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num2].Header.StaticCells[4].CellTemplate.RowIndex - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Value = num7;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateFormula_14(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			double num = 0.0;
			for (int i = 0; i < ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[i];
				num += ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[5].Value.Replace(".", ""));
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_02_1_TD_TNDN/BangPBoThueTNDN/ChiTietBangPBoThueTNDN/soThueTNDNPNop"));
			winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_15(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (iCol == HTKKFile.ColumnLettersToIndex("O"))
			{
				winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].ResetFormula();
				winFpSpread.Sheets[0].Cells[53, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[15].SheetName + "'!O18";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_16(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[16];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[16];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo10 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo11 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo13 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo14 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo15 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo16 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_1_1.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo17 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo18 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo19 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo20 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo21 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo23 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo24 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo25 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo26 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo27 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo28 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo29 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.3/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo30 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.3/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo31 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.3/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo32 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.3/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo33 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo34 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo35 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo36 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo37 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.1/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo38 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.1/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo39 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.1/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo40 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.1/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo41 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.2/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo42 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.2/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo43 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.2/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo44 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.4_2.4.2/ct12d"));
			HTKKCMCCellInfo hTKKCMCCellInfo45 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.5/ct12a"));
			HTKKCMCCellInfo hTKKCMCCellInfo46 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.5/ct12b"));
			HTKKCMCCellInfo hTKKCMCCellInfo47 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.5/ct12c"));
			HTKKCMCCellInfo hTKKCMCCellInfo48 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL_01_ND20/Muc_III/Dong_II_2_2.5/ct12d"));
			if (iRow == hTKKCMCCellInfo5.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo17.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo5.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo6.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo7.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo8.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo9.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo13.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo9.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo13.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo10.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo10.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo14.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo11.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo11.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo15.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo12.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo12.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo16.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo21.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo25.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo29.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo33.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo45.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo21.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo22.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo23.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo24.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
				}
			}
			else if (iRow == hTKKCMCCellInfo37.CellTemplate.RowIndex || iRow == hTKKCMCCellInfo41.CellTemplate.RowIndex)
			{
				if (iCol == hTKKCMCCellInfo41.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo37.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo41.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo42.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo38.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo42.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo43.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo39.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo43.CellTemplate.RowIndex}";
				}
				if (iCol == hTKKCMCCellInfo44.CellTemplate.ColIndex)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
					sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo40.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo44.CellTemplate.RowIndex}";
				}
			}
			for (int num = 1; num < 9; num++)
			{
				if (iRow < hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart || iRow > hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count)
				{
					continue;
				}
				if (num == 1 || num == 2)
				{
					if (iCol == hTKKCMCCellInfo9.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo9.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo13.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AF{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AF{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo10.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo10.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo14.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo11.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo11.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo15.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo12.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo12.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo16.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
				else if (num == 3 || num == 4 || num == 5 || num == 8)
				{
					if (iCol == hTKKCMCCellInfo21.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AF{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AF{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo22.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo23.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo24.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
				else if (num == 6 || num == 7)
				{
					if (iCol == hTKKCMCCellInfo41.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo5.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo17.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo33.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AF{hTKKCMCCellInfo37.CellTemplate.RowIndex}+AF{hTKKCMCCellInfo41.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AF{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AF{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo42.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo6.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo18.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo34.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AG{hTKKCMCCellInfo38.CellTemplate.RowIndex}+AG{hTKKCMCCellInfo42.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AG{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo43.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo7.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo19.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo35.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AH{hTKKCMCCellInfo39.CellTemplate.RowIndex}+AH{hTKKCMCCellInfo43.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AH{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
					if (iCol == hTKKCMCCellInfo44.CellTemplate.ColIndex)
					{
						sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo8.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo20.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo17.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo21.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo25.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo29.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo33.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo45.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCCellInfo36.CellTemplate.RowIndex - 1, (int)iCol - 1].Formula = $"AI{hTKKCMCCellInfo40.CellTemplate.RowIndex}+AI{hTKKCMCCellInfo44.CellTemplate.RowIndex}";
						sheetView.Cells[hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 3, (int)iCol - 1].Formula = $"SUM(AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart - 1}:AI{hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[num].RowsValues.Count})";
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static double SumChiTieuCellTable(SheetView sView, HTKKCMCSheetInfo sInfo, int cellIndex)
	{
		double num = 0.0;
		try
		{
			for (int i = 0; i < sInfo.Tables.Count; i++)
			{
				num += double.Parse(GeneralFunctions.GetValue(sView, sInfo.Tables[i].Header.StaticCells[cellIndex].CellTemplate.RowIndex, sInfo.Tables[i].Header.StaticCells[cellIndex].CellTemplate.ColIndex));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num;
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		//IL_02a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0356: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1 == 83 && winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1 == HTKKFile.ColumnLettersToIndex("F"))
			{
				int row = 82;
				int column = HTKKFile.ColumnLettersToIndex("F") - 1;
				int column2 = HTKKFile.ColumnLettersToIndex("I") - 1;
				int column3 = HTKKFile.ColumnLettersToIndex("P") - 1;
				bool flag = (bool)winFpSpread.Sheets[0].Cells[row, column].Value;
				int[] array = new int[4] { 84, 86, 88, 90 };
				if (flag)
				{
					for (int i = 0; i < array.Length; i++)
					{
						winFpSpread.Sheets[0].Cells[array[i], column2].Locked = false;
					}
					return;
				}
				for (int i = 0; i < array.Length; i++)
				{
					winFpSpread.Sheets[0].Cells[array[i], column2].Locked = true;
				}
				winFpSpread.Sheets[0].Cells[array[0], column2].ResetValue();
				winFpSpread.Sheets[0].Cells[array[0], column3].Value = "";
				winFpSpread.Sheets[0].Cells[array[1], column2].ResetValue();
				winFpSpread.Sheets[0].Cells[array[2], column2].Value = 0;
				winFpSpread.Sheets[0].Cells[array[3], column2].ResetFormula();
				winFpSpread.Sheets[0].Cells[array[3], column2].Value = 0;
			}
			else if (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1 == 10 && winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1 == HTKKFile.ColumnLettersToIndex("D"))
			{
				if (ConvertHelpers.ConvertToBool(winFpSpread.ActiveSheet.ActiveCell.Value))
				{
					frmMessage frmMessage = new frmMessage("0337", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
			}
			else if (winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1 == 12 && winFpSpread.ActiveSheet.ActiveCell.Column.Index + 1 == HTKKFile.ColumnLettersToIndex("D") && ConvertHelpers.ConvertToBool(winFpSpread.ActiveSheet.ActiveCell.Value) && ConvertHelpers.ConvertToInt(PeriodInfo.Nam) < 2017)
			{
				frmMessage frmMessage = new frmMessage("0338", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
		}
		catch (Exception)
		{
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				int[] array = GeneralFunctions.ArrayCellId("J_95");
				DateTime dateTime = DateTime.ParseExact(DateTime.ParseExact(HeaderInfo.NgayBatDauNamTaiChinh + "/1900", "dd/MM/yyyy", CultureInfo.CurrentCulture).AddMonths(1).AddYears(1)
					.ToString("dd/MM/") + (double.Parse(PeriodInfo.Nam) + 1.0), "dd/MM/yyyy", CultureInfo.CurrentCulture);
				sView.Cells[array[1], array[0]].Value = dateTime.ToString("dd/MM/yyyy");
			}
			sView.Cells[GeneralFunctions.ArrayCellId("J_22")[1], GeneralFunctions.ArrayCellId("J_22")[0]].Text = PeriodInfo.TieuMucHachToanVal;
			int row = 82;
			int column = HTKKFile.ColumnLettersToIndex("F") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("I") - 1;
			bool flag = false;
			if (sView.Cells[row, column] != null)
			{
				flag = (bool)sView.Cells[row, column].Value;
			}
			int[] array2 = new int[4] { 84, 86, 88, 90 };
			if (flag)
			{
				for (int i = 0; i < array2.Length; i++)
				{
					sView.Cells[array2[i], column2].Locked = false;
				}
			}
			else
			{
				for (int i = 0; i < array2.Length; i++)
				{
					sView.Cells[array2[i], column2].Locked = true;
				}
			}
			int index = 7;
			int[] array3 = GeneralFunctions.ArrayCellId("O_31");
			int index2 = 12;
			int[] array4 = GeneralFunctions.ArrayCellId("O_37");
			int index3 = 21;
			int[] array5 = GeneralFunctions.ArrayCellId("O_46");
			int index4 = 25;
			int[] array6 = GeneralFunctions.ArrayCellId("O_51");
			int index5 = 28;
			int[] array7 = GeneralFunctions.ArrayCellId("O_54");
			int index6 = 35;
			int[] array8 = GeneralFunctions.ArrayCellId("O_61");
			int index7 = 36;
			int[] array9 = GeneralFunctions.ArrayCellId("O_62");
			int index8 = 39;
			int[] array10 = GeneralFunctions.ArrayCellId("O_65");
			int index9 = 43;
			int[] array11 = GeneralFunctions.ArrayCellId("O_69");
			if (PeriodInfo.LoaiToKhai == "B")
			{
				double num = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index].Value);
				sView.Cells[array3[1], array3[0]].ResetFormula();
				sView.Cells[array3[1], array3[0]].Value = num;
				double num2 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index2].Value);
				sView.Cells[array4[1], array4[0]].ResetFormula();
				sView.Cells[array4[1], array4[0]].Value = num2;
				double num3 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index3].Value);
				sView.Cells[array5[1], array5[0]].ResetFormula();
				sView.Cells[array5[1], array5[0]].Value = num3;
				double num4 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index4].Value);
				sView.Cells[array6[1], array6[0]].ResetFormula();
				sView.Cells[array6[1], array6[0]].Value = num4;
				double num5 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index5].Value);
				sView.Cells[array7[1], array7[0]].ResetFormula();
				sView.Cells[array7[1], array7[0]].Value = num5;
				double num6 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index6].Value);
				sView.Cells[array8[1], array8[0]].ResetFormula();
				sView.Cells[array8[1], array8[0]].Value = num6;
				double num7 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index7].Value);
				sView.Cells[array9[1], array9[0]].ResetFormula();
				sView.Cells[array9[1], array9[0]].Value = num7;
				double num8 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index8].Value);
				sView.Cells[array10[1], array10[0]].ResetFormula();
				sView.Cells[array10[1], array10[0]].Value = num8;
				double num9 = double.Parse(ToKhaiInfo.SheetInfos[0].StaticCells[index9].Value);
				sView.Cells[array11[1], array11[0]].ResetFormula();
				sView.Cells[array11[1], array11[0]].Value = num9;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (!ToKhaiInfo.SheetInfos[4].IsReadDataFile)
			{
				SheetView sheetView = winFpSpread.Sheets[4];
				int num = int.Parse(PeriodInfo.Nam);
				int num2 = num - 5;
				for (int i = 0; i < 5; i++)
				{
					sheetView.Cells[25 + i, 2].Value = num2 + i;
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.DecimalPlaces = 0;
					numberCellType.ShowSeparator = false;
					sheetView.Cells[25 + i, 2].CellType = numberCellType;
				}
			}
			winFpSpread.Sheets[0].Cells[50, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
			winFpSpread.Sheets[0].Cells[50, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[4].SheetName + "'!M32";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_5(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (!ToKhaiInfo.SheetInfos[5].IsReadDataFile)
			{
				int num = int.Parse(PeriodInfo.Nam);
				int num2 = num - 5;
				for (int i = 0; i < 5; i++)
				{
					sView.Cells[25 + i, 2].Value = num2 + i;
					NumberCellType numberCellType = new NumberCellType();
					numberCellType.Separator = ".";
					numberCellType.DecimalSeparator = ",";
					numberCellType.DecimalPlaces = 0;
					numberCellType.ShowSeparator = false;
					sView.Cells[25 + i, 2].CellType = numberCellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_6(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
			winFpSpread.Sheets[0].Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_7(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[60, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
			winFpSpread.Sheets[0].Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_8(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[61, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_9(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.Sheets[0].Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
			winFpSpread.Sheets[0].Cells[64, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "'" + winFpSpread.Sheets[9].SheetName + "'!AR19";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_16(FpSpread winFpSpread, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[16];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[16];
			int count = hTKKCMCSheetInfo.Tables[0].RowsValues.Count;
			int num = count + hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1;
			int num2 = num + hTKKCMCSheetInfo.Tables[2].RowsValues.Count - 1;
			int num3 = num2 + hTKKCMCSheetInfo.Tables[3].RowsValues.Count - 1;
			int num4 = num3 + hTKKCMCSheetInfo.Tables[4].RowsValues.Count - 1;
			int num5 = num4 + hTKKCMCSheetInfo.Tables[5].RowsValues.Count - 1;
			int num6 = num5 + hTKKCMCSheetInfo.Tables[6].RowsValues.Count - 1;
			int num7 = num6 + hTKKCMCSheetInfo.Tables[7].RowsValues.Count - 1;
			int num8 = num7 + hTKKCMCSheetInfo.Tables[8].RowsValues.Count - 1;
			int num9 = num8 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count - 1;
			int num10 = num9 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count - 1;
			int num11 = num10 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count - 1;
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[38 + count, GeneralFunctions.ArrayCellId("N_40")[0]].Value))
			{
				for (int i = 51 + count; i < sheetView.NonEmptyRowCount; i++)
				{
					if (sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AR") - 1].Text == "x")
					{
						sheetView.Rows[i].Visible = false;
					}
				}
			}
			else
			{
				for (int i = 51 + count; i < sheetView.NonEmptyRowCount; i++)
				{
					if (sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("AR") - 1].Text == "x")
					{
						sheetView.Rows[i].Visible = true;
					}
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int i = 1; i < 9; i++)
				{
					for (int j = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1; j < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; j++)
					{
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = true;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = "";
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = true;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AE") - 1].Formula = "";
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AF") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = true;
					}
				}
				bool flag = ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, HTKKFile.ColumnLettersToIndex("U") - 1].Value);
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 106 + num8; i < 123 + num8; i++)
					{
						if (i == 109 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 111 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i >= 114 + num8 && i < 119 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						else if (i == 119 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]+V[4]-V[5]".Replace("[1]", (i - 7).ToString()).Replace("[2]", (i - 6).ToString()).Replace("[3]", (i - 5).ToString())
								.Replace("[4]", (i - 4).ToString())
								.Replace("[5]", (i - 2).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 120 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M[1]-M[2]-M[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N[1]-N[2]-N[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R[1]-R[2]-R[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 121 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]+V[2]+V[3]".Replace("[1]", (i - 1).ToString()).Replace("[2]", (i - 2).ToString()).Replace("[3]", (i - 3).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 122 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "IF(V[1]<>0,ROUND(V[2]/V[1]*100,3),0)".Replace("[1]", i.ToString()).Replace("[2]", (i - 4).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
					}
					for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 106 + num8; i < 123 + num8; i++)
					{
						if (i == 109 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 111 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i >= 114 + num8 && i < 119 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						else if (i == 119 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]+V[4]-V[5]".Replace("[1]", (i - 7).ToString()).Replace("[2]", (i - 6).ToString()).Replace("[3]", (i - 5).ToString())
								.Replace("[4]", (i - 4).ToString())
								.Replace("[5]", (i - 2).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 120 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M[1]-M[2]-M[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N[1]-N[2]-N[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R[1]-R[2]-R[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 121 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]+V[2]+V[3]".Replace("[1]", (i - 1).ToString()).Replace("[2]", (i - 2).ToString()).Replace("[3]", (i - 3).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else if (i == 122 + num8)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "IF(V[1]<>0,ROUND(V[2]/V[1]*100,3),0)".Replace("[1]", i.ToString()).Replace("[2]", (i - 4).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
						else
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
					}
					for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
				}
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 166 + num10; i < 200 + num10; i++)
					{
						if (i != 166 + num10 && i != 176 + num10 && i != 186 + num10 && i != 191 + num10 && i != 194 + num10 && i != 195 + num10 && i != 196 + num10 && i != 197 + num10 && i != 198 + num10 && i != 199 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
					}
					for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
				if (!ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
				{
					for (int i = 166 + num10; i < 200 + num10; i++)
					{
						if (i != 166 + num10 && i != 176 + num10 && i != 186 + num10 && i != 191 + num10 && i != 194 + num10 && i != 195 + num10 && i != 196 + num10 && i != 197 + num10 && i != 198 + num10 && i != 199 + num10)
						{
							if (i == 179 + num10 || i == 182 + num10 || i == 187 + num10 || i == 188 + num10)
							{
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
							}
							else
							{
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (i + 1).ToString());
								sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
							}
						}
					}
					for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					}
				}
			}
			else
			{
				for (int i = 1; i < 9; i++)
				{
					for (int j = hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart; j < hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[i].RowsValues.Count; j++)
					{
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("H") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("J") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = string.Format("J{0}-H{0}", (j + 1).ToString());
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("Y") - 1].Formula = string.Format("V{0}-R{0}", (j + 1).ToString());
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AB") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AE") - 1].Formula = string.Format("ABS(M{0})+ABS(Y{0})", (j + 1).ToString());
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AF") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AG") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AH") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AI") - 1].Locked = false;
						sheetView.Cells[j, HTKKFile.ColumnLettersToIndex("AJ") - 1].Locked = false;
					}
				}
				for (int i = 106 + num8; i < 123 + num8; i++)
				{
					if (i == 109 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 2).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					else if (i == 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M###-M@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N###-N@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R###-R@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V###-V@@@".Replace("###", (i - 1).ToString()).Replace("@@@", i.ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					else if (i >= 114 + num8 && i < 119 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					}
					else if (i == 119 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]+V[4]-V[5]".Replace("[1]", (i - 7).ToString()).Replace("[2]", (i - 6).ToString()).Replace("[3]", (i - 5).ToString())
							.Replace("[4]", (i - 4).ToString())
							.Replace("[5]", (i - 2).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					else if (i == 120 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "M[1]-M[2]-M[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "N[1]-N[2]-N[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "R[1]-R[2]-R[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]-V[2]-V[3]".Replace("[1]", (i - 8).ToString()).Replace("[2]", (i - 7).ToString()).Replace("[3]", (i - 6).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					else if (i == 121 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "V[1]+V[2]+V[3]".Replace("[1]", (i - 1).ToString()).Replace("[2]", (i - 2).ToString()).Replace("[3]", (i - 3).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					else if (i == 122 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "IF(V[1]<>0,ROUND(V[2]/V[1]*100,3),0)".Replace("[1]", i.ToString()).Replace("[2]", (i - 4).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
					else
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (i + 1).ToString());
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
				}
				for (int i = 166 + num10; i < 200 + num10; i++)
				{
					if (i != 166 + num10 && i != 176 + num10 && i != 186 + num10 && i != 191 + num10 && i != 194 + num10 && i != 195 + num10 && i != 196 + num10 && i != 197 + num10 && i != 198 + num10 && i != 199 + num10)
					{
						if (i == 179 + num10 || i == 182 + num10 || i == 187 + num10 || i == 188 + num10)
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = true;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
						}
						else
						{
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = "";
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("R") - 1].Locked = false;
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Formula = "M###+N###+R###".Replace("###", (i + 1).ToString());
							sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
						}
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[100 + num8, GeneralFunctions.ArrayCellId("X_102")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value) || ConvertHelpers.ConvertToBool(sheetView.Cells[42 + count, GeneralFunctions.ArrayCellId("N_44")[0]].Value))
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 106 + num8; i < 114 + num8; i++)
				{
					if (i != 109 + num8 && i != 111 + num8)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[9].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[9].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[130 + num9, GeneralFunctions.ArrayCellId("X_132")[0]].Value))
			{
				for (int i = 136 + num9; i < 151 + num9; i++)
				{
					if (i != 138 + num9 && i != 141 + num9 && i != 147 + num9)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
			}
			else
			{
				for (int i = 136 + num9; i < 151 + num9; i++)
				{
					if (i != 138 + num9 && i != 141 + num9 && i != 147 + num9)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[10].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[10].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[160 + num10, GeneralFunctions.ArrayCellId("X_162")[0]].Value))
			{
				for (int i = 167 + num10; i < 194 + num10; i++)
				{
					if (i != 176 + num10 && i != 179 + num10 && i != 182 + num10 && i != 186 + num10 && i != 187 + num10 && i != 188 + num10 && i != 191 + num10)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				return;
			}
			if (ConvertHelpers.ConvertToBool(sheetView.Cells[40 + count, GeneralFunctions.ArrayCellId("N_42")[0]].Value))
			{
				for (int i = 167 + num10; i < 194 + num10; i++)
				{
					if (i != 176 + num10 && i != 179 + num10 && i != 182 + num10 && i != 186 + num10 && i != 187 + num10 && i != 188 + num10 && i != 191 + num10)
					{
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
						sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					}
				}
				for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Value = 0;
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
				}
				return;
			}
			for (int i = 167 + num10; i < 194 + num10; i++)
			{
				if (i != 176 + num10 && i != 179 + num10 && i != 182 + num10 && i != 186 + num10 && i != 187 + num10 && i != 188 + num10 && i != 191 + num10)
				{
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
					sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1; i < hTKKCMCSheetInfo.Tables[11].Location.DataLocationTemplateRowStart - 1 + hTKKCMCSheetInfo.Tables[11].RowsValues.Count; i++)
			{
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Formula = "";
				sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_15(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			if (PeriodInfo.NgayLapKHBS != "")
			{
				int row = 4;
				int column = HTKKFile.ColumnLettersToIndex("BD") - 1;
				double num = 0.0;
				num = GeneralFunctions.SoNgayNopCham(PeriodInfo);
				sView.Cells[row, column].Value = num;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private string getDenNgayNC(string ngayBdNamTc, int iNam)
	{
		string result = "";
		try
		{
			if (ngayBdNamTc.Equals("01/01"))
			{
				result = "31/03/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
			else if (ngayBdNamTc.Equals("01/04"))
			{
				result = "30/06/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
			else if (ngayBdNamTc.Equals("01/07"))
			{
				result = "30/09/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
			else if (ngayBdNamTc.Equals("01/10"))
			{
				result = "30/12/" + ConvertHelpers.ConvertToString(iNam + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}
}
