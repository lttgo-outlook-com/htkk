using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Cph01TT57Nckt271
{
	private int[] cellMst = GeneralFunctions.ArrayCellId("I_9");

	private int[] cellDoanhNghiep = GeneralFunctions.ArrayCellId("L_5");

	private int[] cellDonViCongLap = GeneralFunctions.ArrayCellId("T_5");

	private HTKKCMCToKhaiInfo _ToKhaiInfo = null;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			_ToKhaiInfo = ToKhaiInfo;
			sView.Cells[cellMst[1], cellMst[0]].Value = PeriodInfo.MaHSDuAnDT;
			if (!(PeriodInfo.LoaiToKhai == "B"))
			{
				return;
			}
			string text = "48~49~50~51~52~53~55~56~57";
			string[] array = text.Split(new char[1] { '~' });
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("U_" + array[i]);
				double num = ConvertHelpers.ConvertToDouble(sView.Cells[array2[1], array2[0]].Value);
				if (!(num > 0.0))
				{
					continue;
				}
				for (int j = 0; j < array.Length; j++)
				{
					if (array[j] != array[i])
					{
						int[] array3 = GeneralFunctions.ArrayCellId("U_" + array[j]);
						sView.Cells[array3[1], array3[0]].Locked = true;
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

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[cellDoanhNghiep[1], cellDoanhNghiep[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[cellDonViCongLap[1], cellDonViCongLap[0]].Value);
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKFile.RemoveValidate(sheetView, cellDoanhNghiep[1], cellDoanhNghiep[0]);
			HTKKFile.RemoveValidate(sheetView, cellDonViCongLap[1], cellDonViCongLap[0]);
			if ((flag && flag2) || (!flag && !flag2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cellDoanhNghiep[1], cellDoanhNghiep[0], text);
				HTKKFile.SetNodeBackGroundPink(sheetView, cellDonViCongLap[1], cellDonViCongLap[0], text);
				value = 2;
			}
			switch (value)
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
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (idxRow == cellDoanhNghiep[1] && idxCol == cellDoanhNghiep[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[cellDoanhNghiep[1], cellDoanhNghiep[0]].Value))
				{
					sheetView.Cells[cellDonViCongLap[1], cellDonViCongLap[0]].Value = false;
				}
			}
			else if (idxRow == cellDonViCongLap[1] && idxCol == cellDonViCongLap[0] && Convert.ToBoolean(sheetView.Cells[cellDonViCongLap[1], cellDonViCongLap[0]].Value))
			{
				sheetView.Cells[cellDoanhNghiep[1], cellDoanhNghiep[0]].Value = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
