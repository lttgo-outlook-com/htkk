using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkPthu03TT80Nckt251
{
	private string cellId_ct01Val = "O_30";

	private string cellId_ct03Val = "Q_30";

	private string cellId_ct16_donvi = "Q_47";

	private string cellId_ct16_daily = "Y_47";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("H_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("I_30");

	private string cellFormat = "AV_82###AV_83###AV_84###AV_87###AV_88###AV_89###AV_90###AV_91###AV_94###AV_95###AV_96###AV_97###AV_98###AV_101###AV_102###AV_103###AV_104###AV_105###AV_108###AV_109###AV_110###AV_111###AV_112###AV_113###AV_114###AV_115###AV_116###AV_117###AV_118";

	private int[] CellId_ct01Val => GeneralFunctions.ArrayCellId(cellId_ct01Val);

	private int[] CellId_ct03Val => GeneralFunctions.ArrayCellId(cellId_ct03Val);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_58");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_58");

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId("AV_82");

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId("AV_83");

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId("AV_84");

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId("AV_87");

	private int[] CellId_ct36 => GeneralFunctions.ArrayCellId("AV_88");

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId("AV_89");

	private int[] CellId_ct38 => GeneralFunctions.ArrayCellId("AV_90");

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId("AV_91");

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId("AV_94");

	private int[] CellId_ct41 => GeneralFunctions.ArrayCellId("AV_95");

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId("AV_96");

	private int[] CellId_ct43 => GeneralFunctions.ArrayCellId("AV_97");

	private int[] CellId_ct44 => GeneralFunctions.ArrayCellId("AV_98");

	private int[] CellId_ct45 => GeneralFunctions.ArrayCellId("AV_101");

	private int[] CellId_ct46 => GeneralFunctions.ArrayCellId("AV_102");

	private int[] CellId_ct47 => GeneralFunctions.ArrayCellId("AV_103");

	private int[] CellId_ct48 => GeneralFunctions.ArrayCellId("AV_104");

	private int[] CellId_ct49 => GeneralFunctions.ArrayCellId("AV_105");

	private int[] CellId_ct50 => GeneralFunctions.ArrayCellId("AV_108");

	private int[] CellId_ct51 => GeneralFunctions.ArrayCellId("AV_109");

	private int[] CellId_ct52 => GeneralFunctions.ArrayCellId("AV_110");

	private int[] CellId_ct53 => GeneralFunctions.ArrayCellId("AV_111");

	private int[] CellId_ct58a => GeneralFunctions.ArrayCellId("AV_117");

	private int[] CellId_ct59 => GeneralFunctions.ArrayCellId("AV_118");

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
				string formula = "BG82";
				string formula2 = "BG83";
				string formula3 = "BG84";
				string formula4 = "BG87";
				string formula5 = "BG88";
				string formula6 = "BG89";
				string formula7 = "BG90";
				string formula8 = "BG91";
				string formula9 = "BG94";
				string formula10 = "BG95";
				string formula11 = "BG96";
				string formula12 = "BG97";
				string formula13 = "BG98";
				string formula14 = "BG101";
				string formula15 = "BG102";
				string formula16 = "BG103";
				string formula17 = "BG104";
				string formula18 = "BG105";
				string formula19 = "BG108";
				string formula20 = "BG109";
				string formula21 = "BG110";
				string formula22 = "BG111";
				string formula23 = "BG117";
				string formula24 = "BG118";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "BH82";
						formula2 = "BH83";
						formula3 = "BH84";
						formula4 = "BH87";
						formula5 = "BH88";
						formula6 = "BH89";
						formula7 = "BH90";
						formula8 = "BH91";
						formula9 = "BH94";
						formula10 = "BH95";
						formula11 = "BH96";
						formula12 = "BH97";
						formula13 = "BH98";
						formula14 = "BH101";
						formula15 = "BH102";
						formula16 = "BH103";
						formula17 = "BH104";
						formula18 = "BH105";
						formula19 = "BH108";
						formula20 = "BH109";
						formula21 = "BH110";
						formula22 = "BH111";
						formula23 = "BH117";
						formula24 = "BH118";
					}
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula;
					sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Formula = formula2;
					sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formula3;
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula4;
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula5;
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula6;
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula7;
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula8;
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula9;
					sheetView.Cells[CellId_ct41[1], CellId_ct41[0]].Formula = formula10;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula11;
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula12;
					sheetView.Cells[CellId_ct44[1], CellId_ct44[0]].Formula = formula13;
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula14;
					sheetView.Cells[CellId_ct46[1], CellId_ct46[0]].Formula = formula15;
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula16;
					sheetView.Cells[CellId_ct48[1], CellId_ct48[0]].Formula = formula17;
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula18;
					sheetView.Cells[CellId_ct50[1], CellId_ct50[0]].Formula = formula19;
					sheetView.Cells[CellId_ct51[1], CellId_ct51[0]].Formula = formula20;
					sheetView.Cells[CellId_ct52[1], CellId_ct52[0]].Formula = formula21;
					sheetView.Cells[CellId_ct53[1], CellId_ct53[0]].Formula = formula22;
					sheetView.Cells[CellId_ct58a[1], CellId_ct58a[0]].Formula = formula23;
					sheetView.Cells[CellId_ct59[1], CellId_ct59[0]].Formula = formula24;
				}
				if ((iRow == 71 || iRow == 73 || iRow == 78) && CellId_ct34[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "BH84";
						formula4 = "BH87";
						formula5 = "BH88";
					}
					sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formula3;
					sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = formula4;
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula5;
				}
				if ((iRow == 71 || iRow == 73 || iRow == 78 || iRow == 87 || iRow == 88) && CellId_ct37[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "BH89";
					}
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula6;
				}
				if ((iRow == 71 || iRow == 73 || iRow == 78 || iRow == 83 || iRow == 87 || iRow == 88 || iRow == 89) && CellId_ct33[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "BH83";
					}
					sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Formula = formula2;
				}
				if ((iRow == 71 || iRow == 74 || iRow == 79) && CellId_ct39[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula8 = "BH91";
						formula9 = "BH94";
						formula10 = "BH95";
					}
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula8;
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula9;
					sheetView.Cells[CellId_ct41[1], CellId_ct41[0]].Formula = formula10;
				}
				if ((iRow == 71 || iRow == 74 || iRow == 79 || iRow == 94 || iRow == 95) && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula11 = "BH96";
					}
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula11;
				}
				if ((iRow == 71 || iRow == 74 || iRow == 79 || iRow == 91 || iRow == 94 || iRow == 95 || iRow == 96) && CellId_ct38[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "BH90";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula7;
				}
				if ((iRow == 71 || iRow == 75 || iRow == 80) && CellId_ct44[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula13 = "BH98";
						formula14 = "BH101";
						formula15 = "BH102";
					}
					sheetView.Cells[CellId_ct44[1], CellId_ct44[0]].Formula = formula13;
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula14;
					sheetView.Cells[CellId_ct46[1], CellId_ct46[0]].Formula = formula15;
				}
				if ((iRow == 71 || iRow == 75 || iRow == 80 || iRow == 101 || iRow == 102) && CellId_ct47[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula16 = "BH103";
					}
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula16;
				}
				if ((iRow == 71 || iRow == 75 || iRow == 80 || iRow == 98 || iRow == 101 || iRow == 102 || iRow == 103) && CellId_ct38[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "BH90";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula7;
				}
				if ((iRow == 71 || iRow == 76 || iRow == 81) && CellId_ct49[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula18 = "BH105";
						formula19 = "BH108";
						formula20 = "BH109";
					}
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula18;
					sheetView.Cells[CellId_ct50[1], CellId_ct50[0]].Formula = formula19;
					sheetView.Cells[CellId_ct51[1], CellId_ct51[0]].Formula = formula20;
				}
				if ((iRow == 71 || iRow == 76 || iRow == 81 || iRow == 108 || iRow == 109) && CellId_ct52[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula21 = "BH110";
					}
					sheetView.Cells[CellId_ct52[1], CellId_ct52[0]].Formula = formula21;
				}
				if ((iRow == 71 || iRow == 76 || iRow == 81 || iRow == 105 || iRow == 108 || iRow == 109 || iRow == 110) && CellId_ct52[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula17 = "BH104";
					}
					sheetView.Cells[CellId_ct48[1], CellId_ct48[0]].Formula = formula17;
				}
				if ((iRow == 71 || iRow == 76 || iRow == 81 || iRow == 105 || iRow == 108 || iRow == 109 || iRow == 110) && CellId_ct43[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula12 = "BH97";
					}
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula12;
				}
				if ((iRow == 71 || iRow == 76 || iRow == 81 || iRow == 83 || iRow == 90 || iRow == 97 || iRow == 104 || iRow == 105 || iRow == 108 || iRow == 109 || iRow == 110) && CellId_ct32[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "BH82";
					}
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula;
				}
				if (iRow >= 112 && iRow <= 115 && CellId_ct53[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula22 = "BH111";
					}
					sheetView.Cells[CellId_ct53[1], CellId_ct53[0]].Formula = formula22;
				}
				if ((iRow == 71 || iRow == 76 || (iRow >= 81 && iRow <= 83) || iRow == 90 || iRow == 97 || iRow == 104 || iRow == 105 || (iRow >= 108 && iRow <= 115)) && CellId_ct58a[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula23 = "BH117";
					}
					sheetView.Cells[CellId_ct58a[1], CellId_ct58a[0]].Formula = formula23;
				}
				if ((iRow == 71 || iRow == 76 || (iRow >= 81 && iRow <= 83) || iRow == 90 || iRow == 97 || iRow == 104 || iRow == 105 || (iRow >= 108 && iRow <= 116)) && CellId_ct59[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula24 = "BH118";
					}
					sheetView.Cells[CellId_ct59[1], CellId_ct59[0]].Formula = formula24;
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
			int[] array = new int[24]
			{
				82, 83, 84, 87, 88, 89, 90, 91, 94, 95,
				96, 97, 98, 101, 102, 103, 104, 105, 108, 109,
				110, 111, 117, 118
			};
			for (int i = 0; i < array.Length; i++)
			{
				int[] array2 = GeneralFunctions.ArrayCellId("AV_" + array[i]);
				int[] array3 = GeneralFunctions.ArrayCellId("BG_" + array[i]);
				if (text == "VND")
				{
					array3 = GeneralFunctions.ArrayCellId("BH_" + array[i]);
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
