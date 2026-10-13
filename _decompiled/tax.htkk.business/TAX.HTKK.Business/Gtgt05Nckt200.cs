using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Gtgt05Nckt200
{
	private string cellId_J54 = "J_54";

	private string cellId_M54 = "M_54";

	private string cellId_C54 = "C_54";

	private string cellId_I54 = "I_54";

	private int[] CellId_J54 => GeneralFunctions.ArrayCellId(cellId_J54);

	private int[] CellId_M54 => GeneralFunctions.ArrayCellId(cellId_M54);

	private int[] CellId_C54 => GeneralFunctions.ArrayCellId(cellId_C54);

	private int[] CellId_I54 => GeneralFunctions.ArrayCellId(cellId_I54);

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		bool result = true;
		foreach (KeyValuePair<int, int> item in levelWarning)
		{
			int key = item.Key;
			switch (item.Value)
			{
			case 2:
				winFpSpread.TabStrip[key].BackColor = Color.Pink;
				continue;
			case 1:
				if (winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					continue;
				}
				break;
			}
			winFpSpread.TabStrip[key].BackColor = Color.White;
		}
		return result;
	}

	public void AutoBindData_0(SheetView sView, HTKKPeriodInfo PeriodInfo, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public int GetDefaultYear()
	{
		return DateTime.Now.Year;
	}
}
