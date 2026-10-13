using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class QtLncl01Nckt227
{
	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("U_39");
			int[] array2 = GeneralFunctions.ArrayCellId("U_40");
			int[] array3 = GeneralFunctions.ArrayCellId("U_41");
			int[] array4 = GeneralFunctions.ArrayCellId("U_42");
			int[] array5 = GeneralFunctions.ArrayCellId("U_43");
			int[] array6 = GeneralFunctions.ArrayCellId("U_44");
			int[] array7 = GeneralFunctions.ArrayCellId("U_45");
			int[] array8 = GeneralFunctions.ArrayCellId("U_46");
			int[] array9 = GeneralFunctions.ArrayCellId("U_47");
			int[] array10 = GeneralFunctions.ArrayCellId("U_48");
			int[] array11 = GeneralFunctions.ArrayCellId("U_49");
			int[] array12 = GeneralFunctions.ArrayCellId("U_50");
			int[] array13 = GeneralFunctions.ArrayCellId("U_51");
			int[] array14 = GeneralFunctions.ArrayCellId("F_53");
			int[] array15 = GeneralFunctions.ArrayCellId("L_53");
			int[] array16 = GeneralFunctions.ArrayCellId("R_53");
			int[] array17 = GeneralFunctions.ArrayCellId("F_55");
			if ((iRow - 1 == array15[1] && iCol - 1 == array15[0]) || (iRow - 1 == array16[1] && iCol - 1 == array16[0]))
			{
				string text = ConvertHelpers.ConvertCellValueToText(sheetView, array15[1], array15[0]);
				string text2 = ConvertHelpers.ConvertCellValueToText(sheetView, array16[1], array16[0]);
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
				{
					DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					double num = (dateTime2 - dateTime).TotalDays + 1.0;
					if (num < 0.0)
					{
						num = 0.0;
					}
					sheetView.Cells[array14[1], array14[0]].Value = num;
					sheetView.Cells[array17[1], array17[0]].Formula = "U51*F53*0.0003";
				}
				else
				{
					sheetView.Cells[array14[1], array14[0]].Value = 0;
				}
			}
			if ((iRow - 1 == array[1] && iCol - 1 == array[0]) || (iRow - 1 == array10[1] && iCol - 1 == array10[0]))
			{
				sheetView.Cells[array17[1], array17[0]].Formula = "U51*F53*0.0003";
			}
			if (iRow - 1 == array[1] && iCol - 1 == array[0])
			{
				if (ConvertHelpers.ConvertToDouble(sheetView.Cells[array[1], array[0]].Value) < 0.0)
				{
					sheetView.Cells[array3[1], array3[0]].Value = 0;
					sheetView.Cells[array4[1], array4[0]].Value = 0;
					sheetView.Cells[array5[1], array5[0]].Value = 0;
					sheetView.Cells[array6[1], array6[0]].Value = 0;
					sheetView.Cells[array7[1], array7[0]].Value = 0;
					sheetView.Cells[array8[1], array8[0]].Value = 0;
					sheetView.Cells[array10[1], array10[0]].Value = 0;
					sheetView.Cells[array3[1], array3[0]].Locked = true;
					sheetView.Cells[array4[1], array4[0]].Locked = true;
					sheetView.Cells[array5[1], array5[0]].Locked = true;
					sheetView.Cells[array6[1], array6[0]].Locked = true;
					sheetView.Cells[array7[1], array7[0]].Locked = true;
					sheetView.Cells[array8[1], array8[0]].Locked = true;
					sheetView.Cells[array10[1], array10[0]].Locked = true;
				}
				else
				{
					sheetView.Cells[array3[1], array3[0]].Locked = false;
					sheetView.Cells[array4[1], array4[0]].Locked = false;
					sheetView.Cells[array5[1], array5[0]].Locked = false;
					sheetView.Cells[array6[1], array6[0]].Locked = false;
					sheetView.Cells[array7[1], array7[0]].Locked = false;
					sheetView.Cells[array8[1], array8[0]].Locked = false;
					sheetView.Cells[array10[1], array10[0]].Locked = false;
				}
			}
			if (ConvertHelpers.ConvertToDouble(sheetView.Cells[array13[1], array13[0]].Value) < 0.0)
			{
				sheetView.Cells[array14[1], array14[0]].Value = 0;
				sheetView.Cells[array15[1], array15[0]].Value = "";
				sheetView.Cells[array16[1], array16[0]].Value = "";
				sheetView.Cells[array17[1], array17[0]].Value = 0;
				sheetView.Cells[array14[1], array14[0]].Locked = true;
				sheetView.Cells[array15[1], array15[0]].Locked = true;
				sheetView.Cells[array16[1], array16[0]].Locked = true;
				sheetView.Cells[array17[1], array17[0]].Locked = true;
			}
			else
			{
				sheetView.Cells[array14[1], array14[0]].Locked = false;
				sheetView.Cells[array15[1], array15[0]].Locked = false;
				sheetView.Cells[array16[1], array16[0]].Locked = false;
				sheetView.Cells[array17[1], array17[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int num = 0;
		if (levelWarning.Count > 0)
		{
			num = levelWarning[0];
		}
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("R_53");
			string value = GeneralFunctions.GetValue(sheetView, "R_53");
			DateTime dateTime = new DateTime(1900, 1, 1);
			if (!string.IsNullOrEmpty(value))
			{
				dateTime = DateTime.ParseExact(value.Replace("/", ""), "ddMMyyyy", CultureInfo.CurrentCulture);
			}
			int[] array2 = GeneralFunctions.ArrayCellId("L_53");
			string value2 = GeneralFunctions.GetValue(sheetView, "L_53");
			DateTime dateTime2 = new DateTime(1900, 1, 1);
			if (!string.IsNullOrEmpty(value2))
			{
				dateTime2 = DateTime.ParseExact(value2.Replace("/", ""), "ddMMyyyy", CultureInfo.CurrentCulture);
			}
			int[] array3 = GeneralFunctions.ArrayCellId("F_55");
			string value3 = GeneralFunctions.GetValue(sheetView, "F_55");
			int[] array4 = GeneralFunctions.ArrayCellId("F_53");
			string value4 = GeneralFunctions.GetValue(sheetView, "F_53");
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (ConvertHelpers.ConvertToDouble(value3) > 0.0 || ConvertHelpers.ConvertToDouble(value4) > 0.0)
			{
				if (string.IsNullOrEmpty(value2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
					num = 2;
				}
				if (string.IsNullOrEmpty(value))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")));
					num = 2;
				}
			}
			if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(value2) && dateTime < dateTime2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
				num = 2;
			}
			switch (num)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
