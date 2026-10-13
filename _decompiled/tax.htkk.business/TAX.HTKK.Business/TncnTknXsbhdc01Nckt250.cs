using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnTknXsbhdc01Nckt250
{
	private string cellId_ct10 = "N_11";

	private string cellId_ct11 = "N_13";

	private string cellId_ct22d = "V_57";

	private string cellId_ct23d = "V_58";

	private string cellId_ct24d = "V_59";

	private string cellId_ct25d = "V_60";

	private string cellId_ct22a = "K_57";

	private string cellId_ct23a = "K_58";

	private string cellId_ct24a = "K_59";

	private string cellId_ct22b = "O_57";

	private string cellId_ct23b = "O_58";

	private string cellId_ct24b = "O_59";

	private string cellId_ct26b = "O_61";

	private int[] CellId_ct10 => GeneralFunctions.ArrayCellId(cellId_ct10);

	private int[] CellId_ct11 => GeneralFunctions.ArrayCellId(cellId_ct11);

	private int[] CellId_ct22d => GeneralFunctions.ArrayCellId(cellId_ct22d);

	private int[] CellId_ct23d => GeneralFunctions.ArrayCellId(cellId_ct23d);

	private int[] CellId_ct24d => GeneralFunctions.ArrayCellId(cellId_ct24d);

	private int[] CellId_ct25d => GeneralFunctions.ArrayCellId(cellId_ct25d);

	private int[] CellId_ct22a => GeneralFunctions.ArrayCellId(cellId_ct22a);

	private int[] CellId_ct23a => GeneralFunctions.ArrayCellId(cellId_ct23a);

	private int[] CellId_ct24a => GeneralFunctions.ArrayCellId(cellId_ct24a);

	private int[] CellId_ct22b => GeneralFunctions.ArrayCellId(cellId_ct22b);

	private int[] CellId_ct23b => GeneralFunctions.ArrayCellId(cellId_ct23b);

	private int[] CellId_ct24b => GeneralFunctions.ArrayCellId(cellId_ct24b);

	private int[] CellId_ct26b => GeneralFunctions.ArrayCellId(cellId_ct26b);

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		int levelWarning2 = ((levelWarning.Count > 0) ? levelWarning[0] : 0);
		checkTKChinh(winFpSpread.Sheets[0], ToKhaiInfo, TTinChung, PeriodInfo, levelWarning2);
		return true;
	}

	private void checkTKChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int levelWarning)
	{
		try
		{
			int num = levelWarning;
			string text = (string)sView.Cells[CellId_ct10[1], CellId_ct10[0]].Value;
			string text2 = (string)sView.Cells[CellId_ct11[1], CellId_ct11[0]].Value;
			if (!string.IsNullOrEmpty(text))
			{
				if (text.Contains(" "))
				{
					HTKKFile.SetNodeBackGroundPink(sView, CellId_ct10[1], CellId_ct10[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_18")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, CellId_ct10[1], CellId_ct10[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, CellId_ct10[1], CellId_ct10[0]);
			}
			if (!string.IsNullOrEmpty(text2))
			{
				if (text2.Contains(" "))
				{
					HTKKFile.SetNodeBackGroundPink(sView, CellId_ct11[1], CellId_ct11[0], HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_19")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sView, CellId_ct11[1], CellId_ct11[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sView, CellId_ct11[1], CellId_ct11[0]);
			}
			switch (num)
			{
			case 2:
				sView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

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

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				sheetView.Cells[CellId_ct22b[1], CellId_ct22b[0]].Locked = false;
				sheetView.Cells[CellId_ct23b[1], CellId_ct23b[0]].Locked = false;
				sheetView.Cells[CellId_ct24b[1], CellId_ct24b[0]].Locked = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				sheetView.Cells[CellId_ct22b[1], CellId_ct22b[0]].Locked = false;
				sheetView.Cells[CellId_ct23b[1], CellId_ct23b[0]].Locked = false;
				sheetView.Cells[CellId_ct24b[1], CellId_ct24b[0]].Locked = false;
			}
			else
			{
				sheetView.Cells[CellId_ct22b[1], CellId_ct22b[0]].Locked = true;
				sheetView.Cells[CellId_ct23b[1], CellId_ct23b[0]].Locked = true;
				sheetView.Cells[CellId_ct24b[1], CellId_ct24b[0]].Locked = true;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo != null)
			{
				int num = HTKKFile.ColumnLettersToIndex("K");
				if (iCol == num && iRow == 57)
				{
					sheetView.Cells[CellId_ct22b[1], CellId_ct22b[0]].Formula = "ROUND(K57*0.05,0)";
				}
				if (iCol == num && iRow == 58)
				{
					sheetView.Cells[CellId_ct23b[1], CellId_ct23b[0]].Formula = "ROUND(K58*0.05,0)";
				}
				if (iCol == num && iRow == 59)
				{
					sheetView.Cells[CellId_ct24b[1], CellId_ct24b[0]].Formula = "ROUND(K59*0.05,0)";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
