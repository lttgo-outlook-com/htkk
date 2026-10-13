using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnKk05Nckt216
{
	private static string cellId_ct21 = "U_38";

	private static string cellId_ct22 = "U_39";

	private static string cellId_ct23 = "U_40";

	private static string cellId_ct24 = "U_41";

	private static string cellId_ct25 = "U_42";

	private static string cellId_ct27 = "U_44";

	private static string cellId_ct28 = "U_45";

	private static string cellId_ct29 = "U_46";

	private static string cellId_ct30 = "U_47";

	private static string cellId_ct31 = "U_48";

	private static string cellId_ct32 = "U_49";

	private static string cellId_ct33 = "U_50";

	private static string cellId_ct34 = "U_51";

	private static int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private static int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private static int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private static int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private static int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private static int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private static int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private static int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private static int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private static int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private static int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private static int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private static int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
			{
				result = false;
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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
}
