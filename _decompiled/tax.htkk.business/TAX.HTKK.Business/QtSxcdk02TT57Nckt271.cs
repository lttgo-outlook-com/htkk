using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class QtSxcdk02TT57Nckt271
{
	private int[] cellMst = GeneralFunctions.ArrayCellId("I_11");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellMst[1], cellMst[0]].Value = PeriodInfo.MaHSDuAnDT;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			bool result = true;
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
}
