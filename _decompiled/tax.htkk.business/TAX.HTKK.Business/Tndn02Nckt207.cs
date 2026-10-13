using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn02Nckt207
{
	private string cellId_ct26 = "AW_62";

	private string cellId_ct33 = "AW_69";

	private string cellId_ct34 = "AW_72";

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			List<HTKKCMCCellInfo> staticCells = hTKKCMCSheetInfo.StaticCells;
			foreach (HTKKCMCCellInfo item in staticCells)
			{
				if ((item.CellTemplate.ColIndex - 1 == CellId_ct33[0]) & (item.CellTemplate.RowIndex - 1 == CellId_ct33[1]))
				{
					if (Convert.ToDouble(item.Value) <= 0.0)
					{
						sView.Cells[CellId_ct34[1], CellId_ct34[0]].Locked = true;
					}
					else
					{
						sView.Cells[CellId_ct34[1], CellId_ct34[0]].Locked = false;
					}
					break;
				}
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
			List<int> list = new List<int>();
			list.Add(61);
			list.Add(63);
			list.Add(64);
			list.Add(65);
			list.Add(66);
			list.Add(67);
			list.Add(68);
			List<int> list2 = list;
			if ((winFpSpread.ActiveSheetIndex == 0) & (iCol - 1 == CellId_ct33[0]) & list2.Contains(Convert.ToInt32(iRow)))
			{
				double num = Convert.ToDouble(winFpSpread.Sheets[0].Cells[CellId_ct33[1], CellId_ct33[0]].Value);
				if (num <= 0.0)
				{
					winFpSpread.Sheets[0].Cells[CellId_ct34[1], CellId_ct34[0]].Value = 0;
					winFpSpread.Sheets[0].Cells[CellId_ct34[1], CellId_ct34[0]].Locked = true;
				}
				else
				{
					winFpSpread.Sheets[0].Cells[CellId_ct34[1], CellId_ct34[0]].Locked = false;
				}
			}
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
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
				}
				else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
				{
					winFpSpread.TabStrip[key].BackColor = Color.Yellow;
				}
				else
				{
					winFpSpread.TabStrip[key].BackColor = Color.White;
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
