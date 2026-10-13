using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkLnclNd373Nckt293
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

	private int[] CellId_psc => GeneralFunctions.ArrayCellId("O_71");

	private int[] CellId_vsp => GeneralFunctions.ArrayCellId("O_73");

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId("AV_86");

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId("AV_88");

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId("AV_91");

	private int[] CellId_ct44 => GeneralFunctions.ArrayCellId("AV_93");

	private int[] CellId_ct45 => GeneralFunctions.ArrayCellId("AV_94");

	private int[] CellId_ct47 => GeneralFunctions.ArrayCellId("AV_96");

	private int[] CellId_ct49 => GeneralFunctions.ArrayCellId("AV_98");

	private int[] CellId_ct50 => GeneralFunctions.ArrayCellId("AV_99");

	private int[] CellId_ct53 => GeneralFunctions.ArrayCellId("AV_102");

	private int[] CellId_ct54 => GeneralFunctions.ArrayCellId("AV_103");

	private int[] CellId_ct61 => GeneralFunctions.ArrayCellId("AV_111");

	private int[] CellId_ct63 => GeneralFunctions.ArrayCellId("AV_113");

	private int[] CellId_ct66 => GeneralFunctions.ArrayCellId("AV_116");

	private int[] CellId_ct67 => GeneralFunctions.ArrayCellId("AV_117");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("O_69");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_69");

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
			sheetView.Cells[CellId_psc[1], CellId_psc[0]].Locked = true;
			sheetView.Cells[CellId_vsp[1], CellId_vsp[0]].Locked = true;
			sheetView.Cells[CellId_psc[1], CellId_psc[0]].Value = PeriodInfo.ToChucCNDNVonChuyenNhuong == 1;
			sheetView.Cells[CellId_vsp[1], CellId_vsp[0]].Value = PeriodInfo.ToChucCNDNVonChuyenNhuong == 2;
			BindDMNN(sheetView, PeriodInfo);
			LockUnlockCellsPSCVSP(sheetView, hTKKCMCSheetInfo);
			if (string.IsNullOrEmpty(TTinChung.tenDLyThue) || string.IsNullOrEmpty(TTinChung.mstDLyThue))
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], locked: true, "", false);
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], locked: true, "", false);
			}
			else if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value = true;
			}
			string cellFormat = "AV_82###AV_83###AV_84###AV_85###AV_86###AV_87###AV_88###AV_89###AV_90###AV_91###AV_93###AV_94###AV_96###AV_97###AV_98###AV_99###AV_102###AV_103###AV_106###AV_107###AV_108###AV_109###AV_110###AV_111###AV_113###AV_116###AV_117";
			string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
			sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text = PeriodInfo.DonViTinhMa;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Text = PeriodInfo.DonViTinhTen;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
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

	private void BindDMNN(SheetView sViewTK, HTKKPeriodInfo periodInfo)
	{
		try
		{
			if (periodInfo.TieuMucHachToanVal != null)
			{
				switch (periodInfo.TieuMucHachToanVal)
				{
				case "1":
					sViewTK.Cells[CellId_O30[1], CellId_O30[0]].Value = true;
					sViewTK.Cells[CellId_Y30[1], CellId_Y30[0]].Text = (((bool)sViewTK.Cells[CellId_vsp[1], CellId_vsp[0]].Value) ? "3753" : "3754");
					sViewTK.Cells[CellId_Z30[1], CellId_Z30[0]].Text = (((bool)sViewTK.Cells[CellId_vsp[1], CellId_vsp[0]].Value) ? "3753" : "3754");
					break;
				case "3":
					sViewTK.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sViewTK.Cells[CellId_Y30[1], CellId_Y30[0]].Text = (((bool)sViewTK.Cells[CellId_vsp[1], CellId_vsp[0]].Value) ? "3954" : "3953");
					sViewTK.Cells[CellId_Z30[1], CellId_Z30[0]].Text = (((bool)sViewTK.Cells[CellId_vsp[1], CellId_vsp[0]].Value) ? "3954" : "3953");
					break;
				case "2":
					sViewTK.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sViewTK.Cells[CellId_Y30[1], CellId_Y30[0]].Text = (((bool)sViewTK.Cells[CellId_vsp[1], CellId_vsp[0]].Value) ? "3804" : "3803");
					sViewTK.Cells[CellId_Z30[1], CellId_Z30[0]].Text = (((bool)sViewTK.Cells[CellId_vsp[1], CellId_vsp[0]].Value) ? "3804" : "3803");
					break;
				case "4":
					sViewTK.Cells[CellId_O35[1], CellId_O35[0]].Value = true;
					sViewTK.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3753";
					sViewTK.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3753";
					break;
				case "6":
					sViewTK.Cells[CellId_P35[1], CellId_P35[0]].Value = true;
					sViewTK.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3954";
					sViewTK.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3954";
					break;
				case "5":
					sViewTK.Cells[CellId_Q35[1], CellId_Q35[0]].Value = true;
					sViewTK.Cells[CellId_Y30[1], CellId_Y30[0]].Text = "3804";
					sViewTK.Cells[CellId_Z30[1], CellId_Z30[0]].Text = "3804";
					break;
				}
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
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string formula = "IF(AV82-AV83-AV84-AV85-AV86-AV87-AV89-AV90 >0,AV82-AV83-AV84-AV85-AV86-AV87-AV89-AV90,0)";
				if ((bool)sheetView.Cells[CellId_psc[1], CellId_psc[0]].Value)
				{
					formula = "IF(AV82-AV83-AV84-AV87-AV88-AV90 >0,AV82-AV83-AV84-AV87-AV88-AV90,0)";
				}
				string formula2 = "ROUND(AV91*AV92/100,2)";
				string formula3 = "IF(AV91-AV93>0,AV91-AV93,0)";
				string formula4 = "ROUND(AV94*AV95/100,2)";
				string formula5 = "ROUND(AV96*1.5/100,2)";
				string formula6 = "AV96+AV97-AV98";
				string formula7 = "ROUND(AV100*(AV96-AV98)/100,2)";
				string formula8 = "ROUND(AV101*(AV96-AV98)/100,2)";
				string formula9 = "IF(AV106-AV109-AV110>0,AV106-AV109-AV110,0)";
				string formula10 = "ROUND(AV111*AV112/100,2)";
				string formula11 = "ROUND(AV113 * AV114 / 100, 2)";
				string formula12 = "ROUND(AV113 * AV115 / 100, 2)";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_82###AV_83###AV_84###AV_85###AV_86###AV_87###AV_88###AV_89###AV_90###AV_91###AV_93###AV_94###AV_96###AV_97###AV_98###AV_99###AV_102###AV_103###AV_106###AV_107###AV_108###AV_109###AV_110###AV_111###AV_113###AV_116###AV_117";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					ReCalculate(sheetView, text);
				}
				if (iRow >= 82 && iRow <= 90 && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula;
				}
				if (iRow >= 82 && iRow <= 92 && CellId_ct44[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "ROUND(AV91*AV92/100,0)";
					}
					sheetView.Cells[CellId_ct44[1], CellId_ct44[0]].Formula = formula2;
				}
				if (iRow >= 82 && iRow <= 93 && CellId_ct45[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "IF(AV91-AV93>0,AV91-AV93,0)";
					}
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula3;
				}
				if (iRow >= 82 && iRow <= 95 && CellId_ct47[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula4 = "ROUND(AV94*AV95/100,0)";
					}
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula4;
				}
				if (iRow >= 82 && iRow <= 97 && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "ROUND(AV96*1.5/100,0)";
					}
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula5;
				}
				if (iRow >= 82 && iRow <= 98 && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "AV96+AV97-AV98";
					}
					sheetView.Cells[CellId_ct50[1], CellId_ct50[0]].Formula = formula6;
				}
				if (iRow >= 82 && iRow <= 100 && iRow != 99 && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "ROUND(AV100*(AV96-AV98)/100,0)";
					}
					sheetView.Cells[CellId_ct53[1], CellId_ct53[0]].Formula = formula7;
				}
				if (((iRow >= 82 && iRow <= 98) || iRow == 101) && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula8 = "ROUND(AV101*(AV96-AV98)/100,0)";
					}
					sheetView.Cells[CellId_ct54[1], CellId_ct54[0]].Formula = formula8;
				}
				if ((iRow == 106 || iRow == 109 || iRow == 110) && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula9 = "IF(AV106-AV109-AV110>0,AV106-AV109-AV110,0)";
					}
					sheetView.Cells[CellId_ct61[1], CellId_ct61[0]].Formula = formula9;
				}
				if ((iRow == 106 || iRow == 109 || iRow == 110 || iRow == 111 || iRow == 112) && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula10 = "ROUND(AV111*AV112/100,0)";
					}
					sheetView.Cells[CellId_ct63[1], CellId_ct63[0]].Formula = formula10;
				}
				if ((iRow == 106 || iRow == 109 || iRow == 110 || iRow == 111 || iRow == 112 || iRow == 113 || iRow == 114) && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula11 = "ROUND(AV113 * AV114 / 100, 0)";
					}
					sheetView.Cells[CellId_ct66[1], CellId_ct66[0]].Formula = formula11;
				}
				if ((iRow == 106 || iRow == 109 || iRow == 110 || iRow == 111 || iRow == 112 || iRow == 113 || iRow == 115) && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula12 = "ROUND(AV113 * AV115 / 100, 0)";
					}
					sheetView.Cells[CellId_ct67[1], CellId_ct67[0]].Formula = formula12;
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

	private void ReCalculate(SheetView sViewTK, string maLoaiTien)
	{
		try
		{
			string formula = "IF(AV82-AV83-AV84-AV85-AV86-AV87-AV89-AV90 >0,AV82-AV83-AV84-AV85-AV86-AV87-AV89-AV90,0)";
			if ((bool)sViewTK.Cells[CellId_psc[1], CellId_psc[0]].Value)
			{
				formula = "IF(AV82-AV83-AV84-AV87-AV88-AV90 >0,AV82-AV83-AV84-AV87-AV88-AV90,0)";
			}
			string formula2 = "ROUND(AV91*AV92/100,2)";
			string formula3 = "IF(AV91-AV93>0,AV91-AV93,0)";
			string formula4 = "ROUND(AV94*AV95/100,2)";
			string formula5 = "ROUND(AV96*1.5/100,2)";
			string formula6 = "AV96+AV97-AV98";
			string formula7 = "ROUND(AV100*(AV96-AV98)/100,2)";
			string formula8 = "ROUND(AV101*(AV96-AV98)/100,2)";
			string formula9 = "IF(AV106-AV109-AV110>0,AV106-AV109-AV110,0)";
			string formula10 = "ROUND(AV111*AV112/100,2)";
			string formula11 = "ROUND(AV113 * AV114 / 100, 2)";
			string formula12 = "ROUND(AV113 * AV115 / 100, 2)";
			if (maLoaiTien == "VND")
			{
				formula2 = "ROUND(AV91*AV92/100,0)";
				formula3 = "IF(AV91-AV93>0,AV91-AV93,0)";
				formula4 = "ROUND(AV94*AV95/100,0)";
				formula5 = "ROUND(AV96*1.5/100,0)";
				formula6 = "AV96+AV97-AV98";
				formula7 = "ROUND(AV100*(AV96-AV98)/100,0)";
				formula8 = "ROUND(AV101*(AV96-AV98)/100,0)";
				formula9 = "IF(AV106-AV109-AV110>0,AV106-AV109-AV110,0)";
				formula10 = "ROUND(AV111*AV112/100,0)";
				formula11 = "ROUND(AV113 * AV114 / 100, 0)";
				formula12 = "ROUND(AV113 * AV115 / 100, 0)";
			}
			sViewTK.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula;
			sViewTK.Cells[CellId_ct44[1], CellId_ct44[0]].Formula = formula2;
			sViewTK.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula3;
			sViewTK.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula4;
			sViewTK.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula5;
			sViewTK.Cells[CellId_ct50[1], CellId_ct50[0]].Formula = formula6;
			sViewTK.Cells[CellId_ct53[1], CellId_ct53[0]].Formula = formula7;
			sViewTK.Cells[CellId_ct54[1], CellId_ct54[0]].Formula = formula8;
			sViewTK.Cells[CellId_ct61[1], CellId_ct61[0]].Formula = formula9;
			sViewTK.Cells[CellId_ct63[1], CellId_ct63[0]].Formula = formula10;
			sViewTK.Cells[CellId_ct66[1], CellId_ct66[0]].Formula = formula11;
			sViewTK.Cells[CellId_ct67[1], CellId_ct67[0]].Formula = formula12;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void LockUnlockCellsPSCVSP(SheetView sView, HTKKCMCSheetInfo sInfoTK)
	{
		try
		{
			bool flag = (bool)sView.Cells[CellId_psc[1], CellId_psc[0]].Value;
			bool flag2 = (bool)sView.Cells[CellId_vsp[1], CellId_vsp[0]].Value;
			string text = "AV_85###AV_86###AV_89###AV_106###AV_107###AV_108###AV_109###AV_110###AV_111###AV_113###AV_116###AV_117###AV_118";
			string text2 = "###";
			string text3 = "";
			if (text3.Length > 0)
			{
				string[] array = text3.Split(new string[1] { text2 }, StringSplitOptions.None);
				for (int i = 0; i < array.Length; i++)
				{
					int[] CellId_format = GeneralFunctions.ArrayCellId(array[i]);
					sView.Cells[CellId_format[1], CellId_format[0]].Locked = flag2;
					if (flag2)
					{
						HTKKCMCCellInfo hTKKCMCCellInfo = sInfoTK.StaticCells.Find((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == CellId_format[1] + 1 && x.CellTemplate.ColIndex == CellId_format[0] + 1);
						sView.Cells[CellId_format[1], CellId_format[0]].Value = hTKKCMCCellInfo.DefaultValue;
					}
				}
			}
			if (text.Length <= 0)
			{
				return;
			}
			string[] array2 = text.Split(new string[1] { text2 }, StringSplitOptions.None);
			for (int i = 0; i < array2.Length; i++)
			{
				int[] CellId_format2 = GeneralFunctions.ArrayCellId(array2[i]);
				sView.Cells[CellId_format2[1], CellId_format2[0]].Locked = flag;
				if (flag)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = sInfoTK.StaticCells.Find((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == CellId_format2[1] + 1 && x.CellTemplate.ColIndex == CellId_format2[0] + 1);
					sView.Cells[CellId_format2[1], CellId_format2[0]].Value = hTKKCMCCellInfo.DefaultValue;
				}
			}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
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
			bool flag3 = (bool)sheetView.Cells[CellId_vsp[1], CellId_vsp[0]].Value;
			bool flag4 = (bool)sheetView.Cells[CellId_psc[1], CellId_psc[0]].Value;
			if ((flag3 && flag4) || (!flag3 && !flag4))
			{
				string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_vsp[1], CellId_vsp[0], message7);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_psc[1], CellId_psc[0], message7);
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_vsp[1], CellId_vsp[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_psc[1], CellId_psc[0]);
			}
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
			int[] array2 = new int[4] { 78, 79, 80, 81 };
			for (int i = 0; i < array2.Length; i++)
			{
				int[] array3 = GeneralFunctions.ArrayCellId("AN_" + array2[i]);
				HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
				string text2 = sheetView.Cells[array3[1], array3[0]].Text;
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message);
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
