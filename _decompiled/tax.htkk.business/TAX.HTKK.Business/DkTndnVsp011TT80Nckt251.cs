using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTndnVsp011TT80Nckt251
{
	private string cellId_ct01Val = "O_30";

	private string cellId_ct02Val = "Q_30";

	private string cellId_ct03Val = "P_30";

	private string cellId_ct16_donvi = "Q_51";

	private string cellId_ct16_daily = "Y_51";

	private int[] CellId_ct01Val => GeneralFunctions.ArrayCellId(cellId_ct01Val);

	private int[] CellId_ct02Val => GeneralFunctions.ArrayCellId(cellId_ct02Val);

	private int[] CellId_ct03Val => GeneralFunctions.ArrayCellId(cellId_ct03Val);

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_ctTuNgay => GeneralFunctions.ArrayCellId("Y_23");

	private int[] CellId_ctDenNgay => GeneralFunctions.ArrayCellId("AM_23");

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId("AV_59");

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId("AV_60");

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId("AV_61");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_55");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_55");

	private int[] CellId_MaSoThueHDDK => GeneralFunctions.ArrayCellId("X_49");

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
				sheetView.Cells[CellId_ct02Val[1], CellId_ct02Val[0]].Value = false;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = false;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3752";
				break;
			case "2":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = false;
				sheetView.Cells[CellId_ct02Val[1], CellId_ct02Val[0]].Value = true;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = false;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3802";
				break;
			case "3":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = false;
				sheetView.Cells[CellId_ct02Val[1], CellId_ct02Val[0]].Value = false;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = true;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3952";
				break;
			}
			string cellFormat = "AV_59###AV_61";
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value = true;
			}
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
				string formula = "ROUND(AV59*AV60/100,2)";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_59###AV_61";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "ROUND(AV59*AV60/100,0)";
					}
					sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = formula;
				}
				if (iRow >= 59 && iRow <= 60 && CellId_ct29[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "ROUND(AV59*AV60/100,0)";
					}
					sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = formula;
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
			string message = HTKKFile.GetMessage(sheetView, CellId_E13);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E26);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E27);
			string message6 = HTKKFile.GetMessage(sheetView, CellId_E28);
			string message7 = HTKKFile.GetMessage(sheetView, CellId_E29);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct16_donvi[1], CellId_ct16_donvi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct16_daily[1], CellId_ct16_daily[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message3);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message3);
					num = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_daily[1], CellId_ct16_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct16_donvi[1], CellId_ct16_donvi[0], message2);
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
			string text = sheetView.Cells[CellId_ctTuNgay[1], CellId_ctTuNgay[0]].Text;
			string text2 = sheetView.Cells[CellId_ctDenNgay[1], CellId_ctDenNgay[0]].Text;
			bool flag3 = false;
			bool flag4 = false;
			HTKKFile.RemoveValidate(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0]);
			DateTime dtDate;
			if (!string.IsNullOrEmpty(text))
			{
				if (HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate) && dtDate >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
				{
					flag3 = true;
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0], message7);
					num = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0], message2);
				num = 2;
			}
			DateTime dtDate2;
			if (!string.IsNullOrEmpty(text2))
			{
				if (HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out dtDate2) && dtDate2 >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
				{
					flag4 = true;
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0], message7);
					num = 2;
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0], message2);
				num = 2;
			}
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text) && flag3 && flag4)
			{
				HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate);
				HTKKFormater.IsDateTime(text2, "dd/MM/yyyy", out dtDate2);
				if (dtDate > dtDate2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctTuNgay[1], CellId_ctTuNgay[0], message6);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctDenNgay[1], CellId_ctDenNgay[0], message6);
					num = 2;
				}
			}
			string text3 = sheetView.Cells[CellId_MaSoThueHDDK[1], CellId_MaSoThueHDDK[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_MaSoThueHDDK[1], CellId_MaSoThueHDDK[0]);
			if (!string.IsNullOrEmpty(text3) && !ValidateBase.CheckMaSoThue(text3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_MaSoThueHDDK[1], CellId_MaSoThueHDDK[0], message);
				num = 2;
			}
			string text4 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array = GeneralFunctions.ArrayCellId("BG_61");
			int[] array2 = GeneralFunctions.ArrayCellId("AV_61");
			double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array[1], array[0]].Value);
			double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
			double num2 = 0.0;
			num2 = ((!(text4 != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
			HTKKFile.RemoveValidateBackGround(sheetView, array2[1], array2[0]);
			if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
			{
				if (text4 != "VND")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message4);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], message5);
				}
				num = 2;
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
