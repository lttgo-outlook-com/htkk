using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnNng04Nckt251
{
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("AD");
			int[] array = GeneralFunctions.ArrayCellId("U_96");
			int[] array2 = GeneralFunctions.ArrayCellId("U_97");
			int[] array3 = GeneralFunctions.ArrayCellId("U_101");
			int[] array4 = GeneralFunctions.ArrayCellId("U_102");
			if (iRow == array[1] + 1)
			{
				sheetView.Cells[array2[1], array2[0]].Value = sheetView.Cells[array2[1], num - 1].Value;
			}
			if (iRow == array3[1] + 1)
			{
				sheetView.Cells[array4[1], array4[0]].Value = sheetView.Cells[array4[1], num - 1].Value;
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
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("C_24");
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			if (hTKKCMCSheetInfo != null)
			{
				string message = HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_48"));
				if (!Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array[1], array[0]].Value))
				{
					winFpSpread.TabStrip[1].BackColor = Color.Pink;
					frmMessage frmMessage = new frmMessage("3", "Cảnh báo", message);
					((Form)frmMessage).ShowDialog();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
