using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkPthu01TT80Nckt251
{
	private string cellId_ct01Val = "O_30";

	private string cellId_ct03Val = "Q_30";

	private string cellId_ct16_donvi = "S_68";

	private string cellId_ct16_daily = "AC_68";

	private string cellFormat = "BC_124###BC_125###BC_126###BC_127";

	private int[] CellId_ct01Val => GeneralFunctions.ArrayCellId(cellId_ct01Val);

	private int[] CellId_ct03Val => GeneralFunctions.ArrayCellId(cellId_ct03Val);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_78");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BN_78");

	private int[] CellId_LoaiDuAn => GeneralFunctions.ArrayCellId("P_108");

	private int[] CellId_MaLoaiLoaiDuAn => GeneralFunctions.ArrayCellId("BN_108");

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId("BC_116");

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId("BC_119");

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId("BC_122");

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId("BC_123");

	private int[] CellId_ct30a => GeneralFunctions.ArrayCellId("BC_125");

	private int[] CellId_ct30b => GeneralFunctions.ArrayCellId("BC_126");

	private int[] CellId_ct30c => GeneralFunctions.ArrayCellId("BC_127");

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_E27 => GeneralFunctions.ArrayCellId("E_27");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId("E_29");

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
				string formula = "BN116";
				string formula2 = "BN119";
				string formula3 = "BN122";
				string formula4 = "BN123";
				string formula5 = "BN125";
				string formula6 = "BN126";
				string formula7 = "BN127";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula5 = "BO125";
						formula6 = "BO126";
						formula7 = "BO127";
					}
					sheetView.Cells[CellId_ct30a[1], CellId_ct30a[0]].Formula = formula5;
					sheetView.Cells[CellId_ct30b[1], CellId_ct30b[0]].Formula = formula6;
					sheetView.Cells[CellId_ct30c[1], CellId_ct30c[0]].Formula = formula7;
				}
				if (CellId_LoaiDuAn[1] + 1 == iRow && CellId_LoaiDuAn[0] + 1 == iCol)
				{
					formula5 = "BN125";
					formula6 = "BN126";
					formula7 = "BN127";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "BO125";
						formula6 = "BO126";
						formula7 = "BO127";
					}
					sheetView.Cells[CellId_ct30a[1], CellId_ct30a[0]].Formula = formula5;
					sheetView.Cells[CellId_ct30b[1], CellId_ct30b[0]].Formula = formula6;
					sheetView.Cells[CellId_ct30c[1], CellId_ct30c[0]].Formula = formula7;
				}
				if (iRow >= 112 && iRow <= 114 && CellId_ct26[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct26[1], CellId_ct26[0]].Formula = formula;
				}
				if (iRow >= 112 && iRow <= 114 && CellId_ct27[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct27[1], CellId_ct27[0]].Formula = formula2;
				}
				if (iRow >= 112 && iRow <= 114 && CellId_ct28[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Formula = formula3;
				}
				if (iRow >= 112 && iRow <= 114 && CellId_ct29[0] + 1 == iCol)
				{
					sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = formula4;
				}
				if (((iRow >= 112 && iRow <= 114) || iRow == 116) && CellId_ct30a[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "BO125";
					}
					sheetView.Cells[CellId_ct30a[1], CellId_ct30a[0]].Formula = formula5;
				}
				if (((iRow >= 112 && iRow <= 114) || iRow == 119) && CellId_ct30b[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "BO126";
					}
					sheetView.Cells[CellId_ct30b[1], CellId_ct30b[0]].Formula = formula6;
				}
				if (((iRow >= 112 && iRow <= 114) || iRow == 122 || iRow == 123) && CellId_ct30c[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "BO126";
					}
					sheetView.Cells[CellId_ct30c[1], CellId_ct30c[0]].Formula = formula7;
				}
				((Control)sheetView.FpSpread).Refresh();
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
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E26);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E29);
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
			int[] array = new int[4] { 116, 119, 122, 123 };
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("BN_" + array[i]);
				int[] array3 = GeneralFunctions.ArrayCellId("BC_" + array[i]);
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num2 = 0.0;
				num2 = ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0);
				HTKKFile.RemoveValidateBackGround(sheetView, array3[1], array3[0]);
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message3);
					num = 2;
				}
			}
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array4 = new int[3] { 125, 126, 127 };
			for (int i = 0; i < array4.Length; i++)
			{
				int[] array5 = GeneralFunctions.ArrayCellId("BC_" + array4[i]);
				int[] array6 = GeneralFunctions.ArrayCellId("BN_" + array4[i]);
				if (text == "VND")
				{
					array6 = GeneralFunctions.ArrayCellId("BO_" + array4[i]);
				}
				double value3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array6[1], array6[0]].Value);
				double value4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array5[1], array5[0]].Value);
				double num2 = 0.0;
				HTKKFile.RemoveValidateBackGround(sheetView, array5[1], array5[0]);
				num2 = ((!(text != "VND")) ? ((Math.Abs(value3) * 0.001 < 100000.0) ? (Math.Abs(value3) * 0.001) : 100000.0) : ((Math.Abs(value3) * 0.001 < 5.0) ? (Math.Abs(value3) * 0.001) : 5.0));
				if (Math.Abs(value4) > Math.Abs(value3) + num2 || Math.Abs(value4) < Math.Abs(value3) - num2)
				{
					if (text != "VND")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], message3);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array5[1], array5[0], message4);
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
