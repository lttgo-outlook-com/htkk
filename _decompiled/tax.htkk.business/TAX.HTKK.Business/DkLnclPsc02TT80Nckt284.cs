using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkLnclPsc02TT80Nckt284
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_ct20_donvi = "S_57";

	private string cellId_ct20_daily = "AC_57";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("H_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("I_30");

	private string cellId_E13 = "E_13";

	private string cellId_E19 = "E_19";

	private string cellId_E25 = "E_25";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_X30 => GeneralFunctions.ArrayCellId("X_30");

	private int[] CellId_AF2 => GeneralFunctions.ArrayCellId("AF_2");

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_ct36 => GeneralFunctions.ArrayCellId("AV_75");

	private int[] CellId_ct38 => GeneralFunctions.ArrayCellId("AV_77");

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId("AV_78");

	private int[] CellId_ct41 => GeneralFunctions.ArrayCellId("AV_80");

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId("AV_81");

	private int[] CellId_ct43 => GeneralFunctions.ArrayCellId("AV_82");

	private int[] CellId_ct46 => GeneralFunctions.ArrayCellId("AV_85");

	private int[] CellId_ct47 => GeneralFunctions.ArrayCellId("AV_86");

	private int[] CellId_ct49 => GeneralFunctions.ArrayCellId("AV_88");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("O_61");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_61");

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId(cellId_E13);

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId(cellId_E19);

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId(cellId_E25);

	private int[] CellId_E32 => GeneralFunctions.ArrayCellId("E_32");

	private int[] CellId_E33 => GeneralFunctions.ArrayCellId("E_33");

	private int[] CellId_E34 => GeneralFunctions.ArrayCellId("E_34");

	private int[] CellId_E35 => GeneralFunctions.ArrayCellId("E_35");

	private int[] CellId_E36 => GeneralFunctions.ArrayCellId("E_36");

	private int[] CellId_E37 => GeneralFunctions.ArrayCellId("E_37");

	private int[] CellId_E38 => GeneralFunctions.ArrayCellId("E_38");

	private int[] CellId_E39 => GeneralFunctions.ArrayCellId("E_39");

	private int[] CellId_E40 => GeneralFunctions.ArrayCellId("E_40");

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
					sheetView.Cells[CellId_X30[1], CellId_X30[0]].Text = "3754";
					break;
				case "3":
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sheetView.Cells[CellId_X30[1], CellId_X30[0]].Text = "3953";
					break;
				case "2":
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sheetView.Cells[CellId_X30[1], CellId_X30[0]].Text = "3803";
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
			if (!string.IsNullOrEmpty(PeriodInfo.MaQuyetToan))
			{
				sheetView.Cells[cellTenTHQT_TK[1], cellTenTHQT_TK[0]].Value = PeriodInfo.TenQuyetToan;
				sheetView.Cells[cellMaTHQT_TK[1], cellMaTHQT_TK[0]].Value = PeriodInfo.MaQuyetToan;
			}
			sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text = PeriodInfo.DonViTinhMa;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Text = PeriodInfo.DonViTinhTen;
			sheetView.Cells[CellId_LoaiTien[1], CellId_LoaiTien[0]].Locked = true;
			string cellFormat = "AV_69###AV_70###AV_71###AV_72###AV_73###AV_74###AV_75###AV_77###AV_78###AV_80###AV_81###AV_82###AV_85###AV_86";
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

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			string text = "";
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					text = text + "Y_" + (dataLocationTemplateRowStart + j) + "###AE_" + (dataLocationTemplateRowStart + j) + "###AJ_" + (dataLocationTemplateRowStart + j) + "###AO_" + (dataLocationTemplateRowStart + j) + "###AT_" + (dataLocationTemplateRowStart + j) + "###AU_" + (dataLocationTemplateRowStart + j) + "###AX_" + (dataLocationTemplateRowStart + j) + "###AZ_" + (dataLocationTemplateRowStart + j) + "###BA_" + (dataLocationTemplateRowStart + j) + "###BC_" + (dataLocationTemplateRowStart + j) + "###";
				}
			}
			text = text.Substring(0, text.Length - 3);
			string donViTinhMa = PeriodInfo.DonViTinhMa;
			GeneralFunctions.changFormatCell(sView, text, donViTinhMa);
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], 1);
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
				string formula = "IF(AV69-AV70-AV71-AV72-AV73-AV74>0,AV69-AV70-AV71-AV72-AV73-AV74,0)";
				string formula2 = "ROUND(AV75*AV76/100,2)";
				string formula3 = "IF(AV75-AV77>0,AV75-AV77,0)";
				string formula4 = "ROUND(AV78*AV79/100,2)";
				string formula5 = "ROUND(AV80*1.5/100,2)";
				string formula6 = "IF(AV80-AV81>0,AV80-AV81,0)";
				string formula7 = "ROUND(AV82*AV83/100,2)";
				string formula8 = "ROUND(AV82*AV84/100,2)";
				string formula9 = "AV85-AV87";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_69###AV_70###AV_71###AV_72###AV_73###AV_74###AV_75###AV_77###AV_78###AV_80###AV_81###AV_82###AV_85###AV_86";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "IF(AV69-AV70-AV71-AV72-AV73-AV74>0,AV69-AV70-AV71-AV72-AV73-AV74,0)";
						formula2 = "ROUND(AV75*AV76/100,2)";
						formula3 = "IF(AV75-AV77>0,AV75-AV77,0)";
						formula4 = "ROUND(AV78*AV79/100,2)";
						formula5 = "ROUND(AV80*1.5/100,2)";
						formula6 = "IF(AV80-AV81>0,AV80-AV81,0)";
						formula7 = "ROUND(AV82*AV83/100,2)";
						formula8 = "ROUND(AV82*AV84/100,2)";
						formula9 = "AV85-AV87";
					}
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula;
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula2;
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula3;
					sheetView.Cells[CellId_ct41[1], CellId_ct41[0]].Formula = formula4;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula5;
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula6;
					sheetView.Cells[CellId_ct46[1], CellId_ct46[0]].Formula = formula7;
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula8;
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula9;
					SheetView sView = winFpSpread.Sheets[1];
					HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
					if (hTKKCMCSheetInfo2 != null)
					{
						string text2 = "";
						for (int i = 0; i < hTKKCMCSheetInfo2.Tables.Count; i++)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[i];
							int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
							for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
							{
								text2 = text2 + "Y_" + (dataLocationTemplateRowStart + j) + "###AE_" + (dataLocationTemplateRowStart + j) + "###AJ_" + (dataLocationTemplateRowStart + j) + "###AO_" + (dataLocationTemplateRowStart + j) + "###AT_" + (dataLocationTemplateRowStart + j) + "###AU_" + (dataLocationTemplateRowStart + j) + "###AX_" + (dataLocationTemplateRowStart + j) + "###AZ_" + (dataLocationTemplateRowStart + j) + "###BA_" + (dataLocationTemplateRowStart + j) + "###BC_" + (dataLocationTemplateRowStart + j) + "###";
							}
						}
						text2 = text2.Substring(0, text2.Length - 3);
						GeneralFunctions.changFormatCell(sView, text2, text);
						GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], 1);
					}
				}
				if (iRow >= 69 && iRow <= 74 && CellId_ct36[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "IF(AV69-AV70-AV71-AV72-AV73-AV74>0,AV69-AV70-AV71-AV72-AV73-AV74,0)";
					}
					sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = formula;
				}
				if (iRow >= 69 && iRow <= 76 && CellId_ct38[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "ROUND(AV75*AV76/100,0)";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula2;
				}
				if (iRow >= 69 && iRow <= 77 && CellId_ct39[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "IF(AV75-AV77>0,AV75-AV77,0)";
					}
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula3;
				}
				if (iRow >= 69 && iRow <= 79 && CellId_ct41[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula4 = "ROUND(AV78*AV79/100,0)";
					}
					sheetView.Cells[CellId_ct41[1], CellId_ct41[0]].Formula = formula4;
				}
				if (iRow >= 69 && iRow <= 80 && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "ROUND(AV80*1.5/100,0)";
					}
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula5;
				}
				if (iRow >= 69 && iRow <= 81 && CellId_ct43[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "IF(AV80-AV81>0,AV80-AV81,0)";
					}
					sheetView.Cells[CellId_ct43[1], CellId_ct43[0]].Formula = formula6;
				}
				if (iRow >= 69 && iRow <= 83 && CellId_ct46[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "ROUND(AV82*AV83/100,0)";
					}
					sheetView.Cells[CellId_ct46[1], CellId_ct46[0]].Formula = formula7;
				}
				if (((iRow >= 69 && iRow <= 82) || iRow == 84) && CellId_ct47[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula8 = "ROUND(AV82*AV84/100,0)";
					}
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula8;
				}
				if (((iRow >= 69 && iRow <= 83) || iRow == 87) && CellId_ct49[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula9 = "AV85-AV87";
					}
					sheetView.Cells[CellId_ct49[1], CellId_ct49[0]].Formula = formula9;
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
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E32);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E33);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E13);
			string message6 = HTKKFile.GetMessage(sheetView, CellId_E34);
			string message7 = HTKKFile.GetMessage(sheetView, CellId_E35);
			string message8 = HTKKFile.GetMessage(sheetView, CellId_E36);
			string message9 = HTKKFile.GetMessage(sheetView, CellId_E37);
			string message10 = HTKKFile.GetMessage(sheetView, CellId_E38);
			string message11 = HTKKFile.GetMessage(sheetView, CellId_E39);
			string message12 = HTKKFile.GetMessage(sheetView, CellId_E40);
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
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message5);
				num = 2;
			}
			string text2 = sheetView.Cells[CellId_AF2[1], CellId_AF2[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_AF2[1], CellId_AF2[0]);
			if (string.IsNullOrEmpty(text2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_AF2[1], CellId_AF2[0], message);
				num = 2;
			}
			string text3 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array2 = new int[9] { 75, 77, 78, 80, 81, 82, 85, 86, 88 };
			for (int i = 0; i < array2.Length; i++)
			{
				int[] array3 = GeneralFunctions.ArrayCellId("AV_" + array2[i]);
				int[] array4 = GeneralFunctions.ArrayCellId("BG_" + array2[i]);
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num2 = 0.0;
				HTKKFile.RemoveValidateBackGround(sheetView, array3[1], array3[0]);
				num2 = ((!(text3 != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					if (text3 != "VND")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message3);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message4);
					}
					num = 2;
				}
				if (hTKKCMCSheetInfo2 == null)
				{
					continue;
				}
				int[] array5 = GeneralFunctions.ArrayCellId("BI_" + array2[i]);
				string text4 = sheetView.Cells[array5[1], array5[0]].Text;
				if (text4 == "0")
				{
					if (array2[i] == 77)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message9);
					}
					if (array2[i] == 80)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message10);
					}
					if (array2[i] == 85)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message11);
					}
					if (array2[i] == 86)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message12);
					}
				}
			}
			if (hTKKCMCSheetInfo2 != null)
			{
				int[] array5 = GeneralFunctions.ArrayCellId("BI_70");
				int[] array3 = GeneralFunctions.ArrayCellId("AV_70");
				string text4 = sheetView.Cells[array5[1], array5[0]].Text;
				HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
				if (text4 == "0")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message6);
				}
				array5 = GeneralFunctions.ArrayCellId("BI_71");
				array3 = GeneralFunctions.ArrayCellId("AV_71");
				text4 = sheetView.Cells[array5[1], array5[0]].Text;
				HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
				if (text4 == "0")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message7);
				}
				array5 = GeneralFunctions.ArrayCellId("BI_74");
				array3 = GeneralFunctions.ArrayCellId("AV_74");
				text4 = sheetView.Cells[array5[1], array5[0]].Text;
				HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
				if (text4 == "0")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message8);
				}
			}
			int[] array6 = new int[4] { 65, 66, 67, 68 };
			for (int i = 0; i < array6.Length; i++)
			{
				int[] array7 = GeneralFunctions.ArrayCellId("AN_" + array6[i]);
				HTKKFile.RemoveValidate(sheetView, array7[1], array7[0]);
				string text5 = sheetView.Cells[array7[1], array7[0]].Text;
				if (string.IsNullOrEmpty(text5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, array7[1], array7[0], message);
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 0;
			string message = HTKKFile.GetMessage(sView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sView, CellId_E13);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
			}
			switch (num)
			{
			case 2:
				winFpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[1].BackColor = Color.White;
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
