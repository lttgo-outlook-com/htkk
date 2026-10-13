using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkLnclVsp01TT80Nckt284
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_O35 = "O_35";

	private string cellId_P35 = "P_35";

	private string cellId_Q35 = "Q_35";

	private string cellId_ct20_donvi = "S_57";

	private string cellId_ct20_daily = "AC_57";

	private string cellId_E13 = "E_13";

	private string cellId_E19 = "E_19";

	private string cellId_E25 = "E_25";

	private string cellId_E28 = "E_28";

	private string cellId_E29 = "E_29";

	private string cellId_E31 = "E_31";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_O35 => GeneralFunctions.ArrayCellId(cellId_O35);

	private int[] CellId_P35 => GeneralFunctions.ArrayCellId(cellId_P35);

	private int[] CellId_Q35 => GeneralFunctions.ArrayCellId(cellId_Q35);

	private int[] CellId_Y30 => GeneralFunctions.ArrayCellId("Y_30");

	private int[] CellId_Z30 => GeneralFunctions.ArrayCellId("Z_30");

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId("AV_77");

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId("AV_79");

	private int[] CellId_ct43 => GeneralFunctions.ArrayCellId("AV_80");

	private int[] CellId_ct45 => GeneralFunctions.ArrayCellId("AV_82");

	private int[] CellId_ct48 => GeneralFunctions.ArrayCellId("AV_85");

	private int[] CellId_ct49 => GeneralFunctions.ArrayCellId("AV_86");

	private int[] CellId_ct54 => GeneralFunctions.ArrayCellId("AV_94");

	private int[] CellId_ct56 => GeneralFunctions.ArrayCellId("AV_96");

	private int[] CellId_ct59 => GeneralFunctions.ArrayCellId("AV_99");

	private int[] CellId_ct60 => GeneralFunctions.ArrayCellId("AV_100");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("O_61");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_61");

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId(cellId_E13);

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId(cellId_E19);

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId(cellId_E25);

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId(cellId_E28);

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId(cellId_E29);

	private int[] CellId_E31 => GeneralFunctions.ArrayCellId(cellId_E31);

	private int[] CellId_E32 => GeneralFunctions.ArrayCellId("E_32");

	private int[] CellId_E33 => GeneralFunctions.ArrayCellId("E_33");

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.TieuMucHachToanVal != null)
			{
				switch (PeriodInfo.TieuMucHachToanVal)
				{
				case "1":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Value = true;
					sheetView.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3753";
					sheetView.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3753";
					break;
				case "3":
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sheetView.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3954";
					sheetView.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3954";
					break;
				case "2":
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sheetView.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3804";
					sheetView.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3804";
					break;
				case "4":
					sheetView.Cells[CellId_O35[1], CellId_O35[0]].Value = true;
					sheetView.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3753";
					sheetView.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3753";
					break;
				case "6":
					sheetView.Cells[CellId_P35[1], CellId_P35[0]].Value = true;
					sheetView.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3954";
					sheetView.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3954";
					break;
				case "5":
					sheetView.Cells[CellId_Q35[1], CellId_Q35[0]].Value = true;
					sheetView.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3804";
					sheetView.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3804";
					break;
				}
			}
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = true;
			}
			string text = "AV_70###AV_71###AV_72###AV_73###AV_74###AV_75###AV_76###AV_77###AV_79###AV_80###AV_82###AV_85###AV_86###AV_89###AV_90###AV_91";
			text += "###AV_92###AV_93###AV_94###AV_96###AV_99###AV_100";
			string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, text, text2);
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
				string formula = "IF(AV70-AV71-AV72-AV73-AV74-AV75-AV76 >0,AV70-AV71-AV72-AV73-AV74-AV75-AV76,0)";
				string formula2 = "ROUND(AV77*AV78/100,2)";
				string formula3 = "IF(AV77-AV79>0,AV77-AV79,0)";
				string formula4 = "ROUND(AV80*AV81/100,2)";
				string formula5 = "ROUND(AV82*AV83/100,2)";
				string formula6 = "ROUND(AV82*AV84/100,2)";
				string formula7 = "IF(AV89-AV92-AV93>0,AV89-AV92-AV93,0)";
				string formula8 = "ROUND(AV94*AV95/100,2)";
				string formula9 = "ROUND(AV96*AV97/100,2)";
				string formula10 = "ROUND(AV96*AV98/100,2)";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string text = "AV_70###AV_71###AV_72###AV_73###AV_74###AV_75###AV_76###AV_77###AV_79###AV_80###AV_82###AV_85###AV_86###AV_89###AV_90###AV_91";
					text += "###AV_92###AV_93###AV_94###AV_96###AV_99###AV_100";
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, text, text2);
					if (text2 == "VND")
					{
						formula = "IF(AV70-AV71-AV72-AV73-AV74-AV75-AV76 >0,AV70-AV71-AV72-AV73-AV74-AV75-AV76,0)";
						formula2 = "ROUND(AV77*AV78/100,0)";
						formula3 = "IF(AV77-AV79>0,AV77-AV79,0)";
						formula4 = "ROUND(AV80*AV81/100,0)";
						formula5 = "ROUND(AV82*AV83/100,0)";
						formula6 = "ROUND(AV82*AV84/100,0)";
						formula7 = "IF(AV89-AV92-AV93>0,AV89-AV92-AV93,0)";
						formula8 = "ROUND(AV94*AV95/100,0)";
						formula9 = "ROUND(AV96*AV97/100,0)";
						formula10 = "ROUND(AV96*AV98/100,0)";
					}
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula2;
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula3;
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula4;
					sheetView.Cells[CellId_ct48[1], CellId_ct48[0]].Formula = formula5;
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula6;
					sheetView.Cells[CellId_ct54[1], CellId_ct54[0]].Formula = formula7;
					sheetView.Cells[CellId_ct56[1], CellId_ct56[0]].Formula = formula8;
					sheetView.Cells[CellId_ct59[1], CellId_ct59[0]].Formula = formula9;
					sheetView.Cells[CellId_ct60[1], CellId_ct60[0]].Formula = formula10;
				}
				if (iRow >= 70 && iRow <= 76 && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula = "IF(AV70-AV71-AV72-AV73-AV74-AV75-AV76 >0,AV70-AV71-AV72-AV73-AV74-AV75-AV76,0)";
					}
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula;
				}
				if (iRow >= 70 && iRow <= 78 && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula2 = "ROUND(AV77*AV78/100,0)";
					}
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula2;
				}
				if (iRow >= 70 && iRow <= 79 && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula3 = "IF(AV77-AV79>0,AV77-AV79,0)";
					}
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula3;
				}
				if (iRow >= 70 && iRow <= 81 && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula4 = "ROUND(AV80*AV81/100,0)";
					}
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula4;
				}
				if (iRow >= 70 && iRow <= 83 && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula5 = "ROUND(AV82*AV83/100,0)";
					}
					sheetView.Cells[CellId_ct48[1], CellId_ct48[0]].Formula = formula5;
				}
				if (((iRow >= 70 && iRow <= 82) || iRow == 84) && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula6 = "ROUND(AV82*AV84/100,0)";
					}
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula6;
				}
				if (((iRow >= 92 && iRow <= 93) || iRow == 89) && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula7 = "IF(AV89-AV92-AV93>0,AV89-AV92-AV93,0)";
					}
					sheetView.Cells[CellId_ct54[1], CellId_ct54[0]].Formula = formula7;
				}
				if (((iRow >= 92 && iRow <= 95) || iRow == 89) && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula8 = "ROUND(AV94*AV95/100,0)";
					}
					sheetView.Cells[CellId_ct56[1], CellId_ct56[0]].Formula = formula8;
				}
				if (((iRow >= 92 && iRow <= 97) || iRow == 89) && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula9 = "ROUND(AV96*AV97/100,0)";
					}
					sheetView.Cells[CellId_ct59[1], CellId_ct59[0]].Formula = formula9;
				}
				if (((iRow >= 92 && iRow <= 96) || iRow == 89 || iRow == 98) && CellId_ct40[0] + 1 == iCol)
				{
					string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text2 == "VND")
					{
						formula10 = "ROUND(AV96*AV98/100,0)";
					}
					sheetView.Cells[CellId_ct60[1], CellId_ct60[0]].Formula = formula10;
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
			if (idxRow == CellId_ct20_daily[1] && idxCol == CellId_ct20_daily[0])
			{
				if (Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value))
				{
					sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value = false;
				}
			}
			else if (idxRow == CellId_ct20_donvi[1] && idxCol == CellId_ct20_donvi[0] && Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value))
			{
				sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = false;
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
			if (ToKhaiInfo.SheetInfos[0] != null)
			{
				Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			}
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink || winFpSpread.TabStrip[1].BackColor == Color.Pink)
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

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E31);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E32);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E33);
			string message6 = HTKKFile.GetMessage(sheetView, CellId_E13);
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value);
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag && flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message2);
					num = 2;
				}
				else if (!flag && !flag2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0]);
					HTKKFile.RemoveValidate(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0]);
			}
			int[] array = GeneralFunctions.ArrayCellId("V_49");
			string text = sheetView.Cells[array[1], array[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThue(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message6);
				num = 2;
			}
			string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array2 = new int[10] { 77, 79, 80, 82, 85, 86, 94, 96, 99, 100 };
			for (int i = 0; i < array2.Length; i++)
			{
				int[] array3 = GeneralFunctions.ArrayCellId("AV_" + array2[i]);
				int[] array4 = GeneralFunctions.ArrayCellId("BG_" + array2[i]);
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num2 = 0.0;
				HTKKFile.RemoveValidateBackGround(sheetView, array3[1], array3[0]);
				num2 = ((!(text2 != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					if (text2 != "VND")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message4);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message5);
					}
					num = 2;
				}
			}
			int[] array5 = new int[4] { 66, 67, 68, 69 };
			for (int i = 0; i < array5.Length; i++)
			{
				int[] array6 = GeneralFunctions.ArrayCellId("AN_" + array5[i]);
				HTKKFile.RemoveValidate(sheetView, array6[1], array6[0]);
				string text3 = sheetView.Cells[array6[1], array6[0]].Text;
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array6[1], array6[0], message);
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
