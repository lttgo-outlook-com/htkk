using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkBctl01TT80Nckt284
{
	private string cellId_R68 = "R_68";

	private string cellId_ct13_donVi = "I_17";

	private string cellId_ct13_daiLy = "O_17";

	private string cellId_U46 = "U_46";

	private string cellId_L47 = "L_47";

	private string cellId_Q47 = "Q_47";

	private int[] CellId_R68 => GeneralFunctions.ArrayCellId(cellId_R68);

	private int[] CellId_ct13_donVi => GeneralFunctions.ArrayCellId(cellId_ct13_donVi);

	private int[] CellId_ct13_daiLy => GeneralFunctions.ArrayCellId(cellId_ct13_daiLy);

	private int[] CellId_U46 => GeneralFunctions.ArrayCellId(cellId_U46);

	private int[] CellId_L47 => GeneralFunctions.ArrayCellId(cellId_L47);

	private int[] CellId_Q47 => GeneralFunctions.ArrayCellId(cellId_Q47);

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E23 => GeneralFunctions.ArrayCellId("E_23");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_E27 => GeneralFunctions.ArrayCellId("E_27");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct13_donVi[1], CellId_ct13_donVi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct13_daiLy[1], CellId_ct13_daiLy[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				CheckBoxCellType cellType = new CheckBoxCellType();
				sheetView.Cells[CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]].CellType = cellType;
				sheetView.Cells[CellId_ct13_donVi[1], CellId_ct13_donVi[0]].CellType = cellType;
				sheetView.Cells[CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]].Value = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			if (idxRow == CellId_ct13_daiLy[1] && idxCol == CellId_ct13_daiLy[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]].Value))
				{
					sheetView.Cells[CellId_ct13_donVi[1], CellId_ct13_donVi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct13_donVi[1] && idxCol == CellId_ct13_donVi[0] && Convert.ToBoolean(sheetView.Cells[CellId_ct13_donVi[1], CellId_ct13_donVi[0]].Value))
			{
				sheetView.Cells[CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]].Value = false;
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
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
			{
				return false;
			}
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E22);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E23);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E25);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct13_donVi[1], CellId_ct13_donVi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct13_daiLy[1], CellId_ct13_daiLy[0], message);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct13_donVi[1], CellId_ct13_donVi[0], message);
					num = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct13_daiLy[1], CellId_ct13_daiLy[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct13_donVi[1], CellId_ct13_donVi[0], message2);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]);
					HTKKFile.RemoveValidate(sheetView, CellId_ct13_donVi[1], CellId_ct13_donVi[0]);
				}
			}
			else if (flag && flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct13_daiLy[1], CellId_ct13_daiLy[0], message3);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct13_donVi[1], CellId_ct13_donVi[0], message3);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct13_daiLy[1], CellId_ct13_daiLy[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct13_donVi[1], CellId_ct13_donVi[0]);
			}
			switch (num)
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
			throw ex;
		}
	}
}
