using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnKk02TT80Nckt251
{
	private bool iflag;

	private static string cellId_ct23 = "R_78";

	private static string cellId_ct28 = "R_83";

	private static string cellId_ct29 = "R_84";

	private int[] cellTemSoThang = GeneralFunctions.ArrayCellId("AD_78");

	private static int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private static int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private static int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	public bool IFlag()
	{
		return iflag;
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int? num = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
			int? num2 = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null));
			string quy = PeriodInfo.Quy;
			string nam = PeriodInfo.Nam;
			sView.Cells[cellTemSoThang[1], cellTemSoThang[0]].Value = ((!num2.HasValue || !num.HasValue) ? new int?(1) : (num2 - num + 1));
			if (ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				return;
			}
			if (PeriodInfo.KieuKy == "Q")
			{
				string text = sView.Cells[CellId_ct23[1], CellId_ct23[0]].Text;
				if (num.HasValue && num2.HasValue && (string.IsNullOrEmpty(text) || text.Equals("0")))
				{
					sView.Cells[cellTemSoThang[1], cellTemSoThang[0]].Value = num2 - num + 1;
					sView.Cells[CellId_ct23[1], CellId_ct23[0]].Value = (num2 - num + 1) * 11000000;
				}
			}
			else
			{
				sView.Cells[cellTemSoThang[1], cellTemSoThang[0]].Value = 1;
				sView.Cells[CellId_ct23[1], CellId_ct23[0]].Value = 11000000;
			}
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
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("R");
			if (num == iCol && iRow >= 75 && iRow <= 82)
			{
				winFpSpread.Sheets[0].RecalculateAll();
				int? num2 = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
				int? sothang = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null)) - num2 + 1;
				if (PeriodInfo.KieuKy == "M")
				{
					sothang = 1;
				}
				double valueCT = Convert.ToDouble(winFpSpread.Sheets[0].Cells[CellId_ct28[1], CellId_ct28[0]].Value);
				winFpSpread.Sheets[0].Cells[CellId_ct29[1], CellId_ct29[0]].Formula = GetFormula29(valueCT, sothang);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		return true;
	}

	private int? GetFirstNumber(string strInput)
	{
		try
		{
			if (string.IsNullOrEmpty(strInput))
			{
				return null;
			}
			string s = strInput.Substring(0, strInput.IndexOf('/'));
			return int.Parse(s);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private string GetFormula29(double valueCT28, int? sothang)
	{
		string text = "0";
		double num = 1000000.0;
		double num2 = valueCT28 / (double)sothang.Value;
		string newValue = valueCT28 + "/" + sothang;
		text = ((num2 <= 5.0 * num) ? "R83*5%" : ((num2 > 5.0 * num && num2 <= 10.0 * num) ? "5000000*5% + (R83-5000000)*10%" : ((num2 > 10.0 * num && num2 <= 18.0 * num) ? "5000000*5% + 5000000*10% + (R83-10000000)*15%" : ((num2 > 18.0 * num && num2 <= 32.0 * num) ? "5000000*5% + 5000000*10% + 8000000*15% + (R83-18000000)*20%" : ((num2 > 32.0 * num && num2 <= 52.0 * num) ? "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + (R83-32000000)*25%" : ((!(num2 > 52.0 * num) || !(num2 <= 80.0 * num)) ? "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + 20000000*25% + 28000000*30% + (R83-80000000)*35%" : "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + 20000000*25% + (R83-52000000)*30%"))))));
		return "ROUND((" + text.Replace("R83", newValue) + ")*" + sothang + ",0)";
	}
}
