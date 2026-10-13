using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkLnclVsp02TT80Nckt251
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("H_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("I_30");

	private string cellId_ct20_donvi = "Q_54";

	private string cellId_ct20_daily = "Y_54";

	private string cellId_E13 = "E_13";

	private string cellId_E19 = "E_19";

	private string cellId_E25 = "E_25";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_X30 => GeneralFunctions.ArrayCellId("X_30");

	private int[] CellId_AF2 => GeneralFunctions.ArrayCellId("V_41");

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId("AV_84");

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId("AV_86");

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId("AV_87");

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId("AV_89");

	private int[] CellId_ct45 => GeneralFunctions.ArrayCellId("AV_92");

	private int[] CellId_ct46 => GeneralFunctions.ArrayCellId("AV_93");

	private int[] CellId_ct47 => GeneralFunctions.ArrayCellId("AV_95");

	private int[] CellId_ct50 => GeneralFunctions.ArrayCellId("AV_100");

	private int[] CellId_ct52 => GeneralFunctions.ArrayCellId("AV_102");

	private int[] CellId_ct55 => GeneralFunctions.ArrayCellId("AV_105");

	private int[] CellId_ct56 => GeneralFunctions.ArrayCellId("AV_106");

	private int[] CellId_ct57 => GeneralFunctions.ArrayCellId("AV_107");

	private int[] CellId_ct59 => GeneralFunctions.ArrayCellId("AV_109");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_68");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_68");

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId(cellId_E13);

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId(cellId_E19);

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId(cellId_E25);

	private int[] CellId_E32 => GeneralFunctions.ArrayCellId("E_32");

	private int[] CellId_E33 => GeneralFunctions.ArrayCellId("E_33");

	private int[] CellId_E36 => GeneralFunctions.ArrayCellId("E_36");

	private int idxPL1_ColD => HTKKFile.ColumnLettersToIndex("D") - 1;

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
					sheetView.Cells[CellId_X30[1], CellId_X30[0]].Text = "3753";
					break;
				case "3":
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sheetView.Cells[CellId_X30[1], CellId_X30[0]].Text = "3954";
					break;
				case "2":
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sheetView.Cells[CellId_X30[1], CellId_X30[0]].Text = "3804";
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
			string cellFormat = "AV_77###AV_78###AV_79###AV_80###AV_81###AV_82###AV_83###AV_84###AV_86###AV_87###AV_89###AV_92###AV_93###AV_95###AV_96###AV_97###AV_98###AV_99###AV_100###AV_102###AV_105###AV_106###AV_107###AV_108###AV_109";
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
				string formula = "BG84";
				string formula2 = "BG86";
				string formula3 = "BG87";
				string formula4 = "BG89";
				string formula5 = "BG92";
				string formula6 = "BG93";
				string formula7 = "BG95";
				string formula8 = "BG100";
				string formula9 = "BG102";
				string formula10 = "BG105";
				string formula11 = "BG106";
				string formula12 = "BG107";
				string formula13 = "BG109";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_77###AV_78###AV_79###AV_80###AV_81###AV_82###AV_83###AV_84###AV_86###AV_87###AV_89###AV_92###AV_93###AV_95###AV_96###AV_97###AV_98###AV_99###AV_100###AV_102###AV_105###AV_106###AV_107###AV_108###AV_109";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "BH84";
						formula2 = "BH86";
						formula3 = "BH87";
						formula4 = "BH89";
						formula5 = "BH92";
						formula6 = "BH93";
						formula7 = "BH95";
						formula8 = "BH100";
						formula9 = "BH102";
						formula10 = "BH105";
						formula11 = "BH106";
						formula12 = "BH107";
						formula13 = "BH109";
					}
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula;
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula2;
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula3;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula4;
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula5;
					sheetView.Cells[CellId_ct46[1], CellId_ct46[0]].Formula = formula6;
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula7;
					sheetView.Cells[CellId_ct50[1], CellId_ct50[0]].Formula = formula8;
					sheetView.Cells[CellId_ct52[1], CellId_ct52[0]].Formula = formula9;
					sheetView.Cells[CellId_ct55[1], CellId_ct55[0]].Formula = formula10;
					sheetView.Cells[CellId_ct56[1], CellId_ct56[0]].Formula = formula11;
					sheetView.Cells[CellId_ct57[1], CellId_ct57[0]].Formula = formula12;
					sheetView.Cells[CellId_ct59[1], CellId_ct59[0]].Formula = formula13;
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
				if (iRow >= 77 && iRow <= 83 && CellId_ct37[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "BH84";
					}
					sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formula;
				}
				if (iRow >= 77 && iRow <= 85 && CellId_ct39[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "BH86";
					}
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula2;
				}
				if (iRow >= 77 && iRow <= 86 && CellId_ct40[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "BH87";
					}
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula3;
				}
				if (iRow >= 77 && iRow <= 88 && CellId_ct42[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula4 = "BH89";
					}
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula4;
				}
				if (iRow >= 77 && iRow <= 90 && CellId_ct45[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula5 = "BH92";
					}
					sheetView.Cells[CellId_ct45[1], CellId_ct45[0]].Formula = formula5;
				}
				if (((iRow >= 77 && iRow <= 89) || iRow == 91) && CellId_ct46[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula6 = "BH93";
					}
					sheetView.Cells[CellId_ct46[1], CellId_ct46[0]].Formula = formula6;
				}
				if (iRow >= 96 && iRow <= 97 && CellId_ct47[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula7 = "BH95";
					}
					sheetView.Cells[CellId_ct47[1], CellId_ct47[0]].Formula = formula7;
				}
				if (iRow >= 95 && iRow <= 99 && CellId_ct50[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula8 = "BH100";
					}
					sheetView.Cells[CellId_ct50[1], CellId_ct50[0]].Formula = formula8;
				}
				if (iRow >= 95 && iRow <= 101 && CellId_ct52[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula9 = "BH102";
					}
					sheetView.Cells[CellId_ct52[1], CellId_ct52[0]].Formula = formula9;
				}
				if (iRow >= 95 && iRow <= 103 && CellId_ct55[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula10 = "BH105";
					}
					sheetView.Cells[CellId_ct55[1], CellId_ct55[0]].Formula = formula10;
				}
				if (((iRow >= 95 && iRow <= 102) || iRow == 104) && CellId_ct56[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula11 = "BH106";
					}
					sheetView.Cells[CellId_ct56[1], CellId_ct56[0]].Formula = formula11;
				}
				if (((iRow >= 77 && iRow <= 90) || (iRow >= 95 && iRow <= 103) || iRow == 92 || iRow == 105) && CellId_ct57[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula12 = "BH107";
					}
					sheetView.Cells[CellId_ct57[1], CellId_ct57[0]].Formula = formula12;
				}
				if (((iRow >= 77 && iRow <= 90) || (iRow >= 95 && iRow <= 103) || iRow == 92 || iRow == 105 || iRow == 107 || iRow == 108) && CellId_ct59[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula13 = "BH109";
					}
					sheetView.Cells[CellId_ct59[1], CellId_ct59[0]].Formula = formula13;
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
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E32);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E33);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E13);
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
			int[] array = GeneralFunctions.ArrayCellId("V_43");
			string text = sheetView.Cells[array[1], array[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThue(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message5);
				num = 2;
			}
			string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array2 = new int[13]
			{
				84, 86, 87, 89, 92, 93, 95, 100, 102, 105,
				106, 107, 109
			};
			for (int i = 0; i < array2.Length; i++)
			{
				int[] array3 = GeneralFunctions.ArrayCellId("AV_" + array2[i]);
				int[] array4 = GeneralFunctions.ArrayCellId("BG_" + array2[i]);
				if (text2 == "VND")
				{
					array4 = GeneralFunctions.ArrayCellId("BH_" + array2[i]);
				}
				double value = ConvertHelpers.ConvertToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				double value2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				double num2 = 0.0;
				HTKKFile.RemoveValidateBackGround(sheetView, array3[1], array3[0]);
				num2 = ((!(text2 != "VND")) ? ((Math.Abs(value) * 0.001 < 100000.0) ? (Math.Abs(value) * 0.001) : 100000.0) : ((Math.Abs(value) * 0.001 < 5.0) ? (Math.Abs(value) * 0.001) : 5.0));
				if (Math.Abs(value2) > Math.Abs(value) + num2 || Math.Abs(value2) < Math.Abs(value) - num2)
				{
					if (text2 != "VND")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message3);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, array3[1], array3[0], message4);
					}
					num = 2;
				}
			}
			int[] array5 = new int[4] { 73, 74, 75, 76 };
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, CellId_E36);
			HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.Tables != null && hTKKCMCSheetInfo.Tables.Count > 0) ? hTKKCMCSheetInfo.Tables[0] : null);
			string message2 = HTKKFile.GetMessage("0071");
			if (hTKKCMCTableInfo != null)
			{
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string text = sheetView.Cells[num2 + i, idxPL1_ColD].Text;
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxPL1_ColD);
					if (!string.IsNullOrEmpty(text))
					{
						if (!HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out var dtDate))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxPL1_ColD, message2);
							num = 2;
						}
						else if (dtDate.CompareTo(DateTime.Now) > 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxPL1_ColD, message);
							num = 2;
						}
					}
				}
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
