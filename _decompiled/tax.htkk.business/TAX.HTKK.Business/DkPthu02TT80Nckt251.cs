using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkPthu02TT80Nckt251
{
	private string cellId_ct01Val = "O_30";

	private string cellId_ct03Val = "Q_30";

	private string cellId_ct16_donvi = "S_69";

	private string cellId_ct16_daily = "AC_69";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("H_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("I_30");

	private string cellFormat = "BC_124###BC_125###BC_126###BC_127###BC_128###BC_129###BC_130###BC_131###BC_132###BC_133###BC_134###BC_135###BC_136";

	private int[] CellId_ct01Val => GeneralFunctions.ArrayCellId(cellId_ct01Val);

	private int[] CellId_ct03Val => GeneralFunctions.ArrayCellId(cellId_ct03Val);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_78");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BN_78");

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId("BC_124");

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId("BC_125");

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId("BC_126");

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId("BC_127");

	private int[] CellId_ct36 => GeneralFunctions.ArrayCellId("BC_128");

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId("BC_129");

	private int[] CellId_ct41a => GeneralFunctions.ArrayCellId("BC_134");

	private int[] CellId_ct43 => GeneralFunctions.ArrayCellId("BC_136");

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_E27 => GeneralFunctions.ArrayCellId("E_27");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	private int[] CellId_E32 => GeneralFunctions.ArrayCellId("E_32");

	private int[] CellId_E33 => GeneralFunctions.ArrayCellId("E_33");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			switch (PeriodInfo.TieuMucHachToanVal)
			{
			case "1":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = true;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = false;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3756";
				break;
			case "3":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = false;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = true;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3956";
				break;
			}
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value = true;
			}
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sheetView.Cells[cellTenTHQT_TK[1], cellTenTHQT_TK[0]].Value = PeriodInfo.TenQuyetToan;
				sheetView.Cells[cellMaTHQT_TK[1], cellMaTHQT_TK[0]].Value = PeriodInfo.MaQuyetToan;
			}
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
			if (PeriodInfo.LoaiToKhai == "B")
			{
				sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string formula = "BN124";
				string formula2 = "BN125";
				string formula3 = "BN126";
				string formula4 = "BN127";
				string formula5 = "BN128";
				string formula6 = "BN129";
				string formula7 = "BN134";
				string formula8 = "BN136";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "BO124";
						formula2 = "BO125";
						formula3 = "BO126";
						formula4 = "BO127";
						formula5 = "BO128";
						formula6 = "BO129";
						formula7 = "BO134";
						formula8 = "BO136";
					}
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula;
					sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Formula = formula2;
					sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formula3;
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula4;
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula5;
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula6;
					sheetView.Cells[CellId_ct41a[1], CellId_ct41a[0]].Formula = formula7;
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula8;
				}
				if ((iRow == 113 || iRow == 114 || iRow == 120) && CellId_ct33[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "BO125";
					}
					sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Formula = formula2;
				}
				if ((iRow == 113 || iRow == 116 || iRow == 121) && CellId_ct34[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "BO126";
					}
					sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formula3;
				}
				if ((iRow == 113 || iRow == 117 || iRow == 122) && CellId_ct35[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula4 = "BO127";
					}
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula4;
				}
				if ((iRow == 113 || iRow == 118 || iRow == 123) && CellId_ct36[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "BO128";
					}
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula5;
				}
				if (((iRow >= 125 && iRow <= 128) || (iRow >= 113 && iRow <= 123)) && CellId_ct32[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "BO124";
					}
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula;
				}
				if (iRow >= 130 && iRow <= 133 && CellId_ct37[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "BO129";
					}
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula6;
				}
				if (iRow >= 113 && iRow <= 133 && CellId_ct41a[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "BO134";
					}
					sheetView.Cells[CellId_ct41a[1], CellId_ct41a[0]].Formula = formula7;
				}
				if (((iRow >= 113 && iRow <= 133) || iRow == 135) && CellId_ct43[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula8 = "BO136";
					}
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula8;
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0], 0);
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
			if (idxRow == CellId_ct16_daily[1] && idxCol == CellId_ct16_daily[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value))
				{
					sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct16_donvi[1] && idxCol == CellId_ct16_donvi[0] && Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value))
			{
				sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value = false;
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
			{
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E32);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E33);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message2);
					num = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0]);
					HTKKFile.RemoveValidate(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0]);
			}
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array = new int[8] { 124, 125, 126, 127, 128, 129, 134, 136 };
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("BC_" + array[i]);
				int[] array3 = GeneralFunctions.ArrayCellId("BN_" + array[i]);
				if (text == "VND")
				{
					array3 = GeneralFunctions.ArrayCellId("BO_" + array[i]);
				}
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
				double num2 = 0.0;
				HTKKFile.RemoveValidateBackGround(sheetView, array2[1], array2[0]);
				num2 = ((!(text != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					if (text != "VND")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message3);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message4);
					}
					num = 2;
				}
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
