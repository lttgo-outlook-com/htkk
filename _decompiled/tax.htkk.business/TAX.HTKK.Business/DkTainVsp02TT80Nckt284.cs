using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTainVsp02TT80Nckt284
{
	private string cellId_ct01Val = "O_30";

	private string cellId_ct02Val = "P_30";

	private string cellId_ct03Val = "Q_30";

	private string cellId_ct16_donvi = "Q_49";

	private string cellId_ct16_daily = "Y_49";

	private int[] cellTenTHQT_TK = GeneralFunctions.ArrayCellId("H_30");

	private int[] cellMaTHQT_TK = GeneralFunctions.ArrayCellId("I_30");

	private string cellFormat = "AV_74###AV_75###AV_77###AV_78###AV_79";

	private int[] CellId_ct01Val => GeneralFunctions.ArrayCellId(cellId_ct01Val);

	private int[] CellId_ct02Val => GeneralFunctions.ArrayCellId(cellId_ct02Val);

	private int[] CellId_ct03Val => GeneralFunctions.ArrayCellId(cellId_ct03Val);

	private int[] CellId_ct16_donvi => GeneralFunctions.ArrayCellId(cellId_ct16_donvi);

	private int[] CellId_ct16_daily => GeneralFunctions.ArrayCellId(cellId_ct16_daily);

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("Q_61");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_61");

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId("AV_75");

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId("AV_77");

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId("AV_79");

	private int idxPL1_ColAO => HTKKFile.ColumnLettersToIndex("AO") - 1;

	private int idxPL1_ColY => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxPL1_ColAF => HTKKFile.ColumnLettersToIndex("AF") - 1;

	private int idxPL1_ColBG => HTKKFile.ColumnLettersToIndex("BG") - 1;

	private int idxPL1_ColBI => HTKKFile.ColumnLettersToIndex("BI") - 1;

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_E27 => GeneralFunctions.ArrayCellId("E_27");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	private int[] CellId_E31 => GeneralFunctions.ArrayCellId("E_31");

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
				sheetView.Cells[CellId_ct02Val[1], CellId_ct02Val[0]].Value = false;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = false;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3751";
				break;
			case "2":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = false;
				sheetView.Cells[CellId_ct02Val[1], CellId_ct02Val[0]].Value = false;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = true;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3801";
				break;
			case "3":
				sheetView.Cells[CellId_ct01Val[1], CellId_ct01Val[0]].Value = false;
				sheetView.Cells[CellId_ct02Val[1], CellId_ct02Val[0]].Value = true;
				sheetView.Cells[CellId_ct03Val[1], CellId_ct03Val[0]].Value = false;
				sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3951";
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
			int num = 0;
			num = (text.Equals("USD") ? 6 : 0);
			int[] array = GeneralFunctions.ArrayCellId("AV_75");
			NumberCellType numberCellType = (NumberCellType)sheetView.Cells[array[1], array[0]].CellType;
			numberCellType.DecimalPlaces = num;
			sheetView.Cells[array[1], array[0]].CellType = numberCellType;
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
			string text2 = "";
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					text = text + "AO_" + (dataLocationTemplateRowStart + j) + "###";
					text2 = text2 + "AF_" + (dataLocationTemplateRowStart + j) + "###";
				}
			}
			text = text.Substring(0, text.Length - 3);
			text2 = text2.Substring(0, text2.Length - 3);
			string text3 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			GeneralFunctions.changFormatCell(sView, text, text3);
			GeneralFunctions.changFormatCell(sView, text2, text3);
			text = "";
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL02_1_TAIN_VSP/BangSLuongVaDThu/tongCot7");
			text = "AO_" + hTKKCMCCellInfo.CellTemplate.RowIndex;
			GeneralFunctions.changFormatCell(sView, text, text3);
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
				string formula = "BG75";
				string formula2 = "BG77";
				string formula3 = "BG79";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					int num = 0;
					num = (text.Equals("USD") ? 6 : 0);
					int[] array = GeneralFunctions.ArrayCellId("AV_75");
					NumberCellType numberCellType = (NumberCellType)sheetView.Cells[array[1], array[0]].CellType;
					numberCellType.DecimalPlaces = num;
					sheetView.Cells[array[1], array[0]].CellType = numberCellType;
					if (text == "VND")
					{
						formula = "BH75";
						formula2 = "BH77";
						formula3 = "BG79";
					}
					sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Formula = formula;
					sheetView.Cells[CellId_ct30[1], CellId_ct30[0]].Formula = formula2;
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula3;
					SheetView sheetView2 = winFpSpread.Sheets[1];
					HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[1];
					if (hTKKCMCSheetInfo2 != null)
					{
						string text2 = "";
						string text3 = "";
						for (int i = 0; i < hTKKCMCSheetInfo2.Tables.Count; i++)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[i];
							int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
							for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
							{
								text2 = text2 + "AO_" + (dataLocationTemplateRowStart + j) + "###";
								text3 = text3 + "AF_" + (dataLocationTemplateRowStart + j) + "###";
							}
						}
						text2 = text2.Substring(0, text2.Length - 3);
						text3 = text3.Substring(0, text3.Length - 3);
						GeneralFunctions.changFormatCell(sheetView2, text2, text);
						GeneralFunctions.changFormatCell(sheetView2, text3, text);
						for (int i = 0; i < hTKKCMCSheetInfo2.Tables.Count; i++)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[i];
							int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
							for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
							{
								if (text == "VND")
								{
									sheetView2.Cells[dataLocationTemplateRowStart + j - 1, idxPL1_ColAO].Formula = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, idxPL1_ColBI].Formula;
								}
								else
								{
									sheetView2.Cells[dataLocationTemplateRowStart + j - 1, idxPL1_ColAO].Formula = sheetView2.Cells[dataLocationTemplateRowStart + j - 1, idxPL1_ColBG].Formula;
								}
							}
						}
						text2 = "";
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL02_1_TAIN_VSP/BangSLuongVaDThu/tongCot7");
						text2 = "AO_" + hTKKCMCCellInfo.CellTemplate.RowIndex;
						GeneralFunctions.changFormatCell(sheetView2, text2, text);
						GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], 1);
					}
				}
				if (iRow >= 72 && iRow <= 74 && CellId_ct28[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "BH75";
					}
					sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Formula = formula;
				}
				if (iRow >= 72 && iRow <= 76 && CellId_ct30[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "BH77";
					}
					sheetView.Cells[CellId_ct30[1], CellId_ct30[0]].Formula = formula2;
				}
				if (iRow >= 72 && iRow <= 78 && CellId_ct32[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "BG79";
					}
					sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Formula = formula3;
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

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (activeSheetIndex == 1)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				SheetView sheetView2 = winFpSpread.Sheets[1];
				string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
				if (idxPL1_ColY + 1 == iCol || idxPL1_ColAF + 1 == iCol)
				{
					if (text == "VND")
					{
						sheetView2.Cells[(int)iRow - 1, idxPL1_ColAO].Formula = sheetView2.Cells[(int)iRow - 1, idxPL1_ColBI].Formula;
					}
					else
					{
						sheetView2.Cells[(int)iRow - 1, idxPL1_ColAO].Formula = sheetView2.Cells[(int)iRow - 1, idxPL1_ColBG].Formula;
					}
				}
			}
			GeneralFunctions.SaveAllCellTypeOfSheet(winFpSpread.Sheets[1], ToKhaiInfo.SheetInfos[1], 1);
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
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E31);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E32);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E13);
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
			int[] array = GeneralFunctions.ArrayCellId("Y_47");
			string text = sheetView.Cells[array[1], array[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text) && !ValidateBase.CheckMaSoThue(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message5);
				num = 2;
			}
			string text2 = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
			int[] array2 = new int[3] { 75, 77, 79 };
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
			int[] array5 = new int[5] { 71, 72, 73, 75, 76 };
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
			SheetView sView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = 0;
			string message = HTKKFile.GetMessage(sView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sView, CellId_E13);
			bool flag = false;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
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
