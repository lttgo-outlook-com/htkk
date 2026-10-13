using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTkVsp01TT80Nckt251
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_AF23 = "AF_23";

	private string cellId_ct08 = "Y_28";

	private string cellId_ct09 = "AS_28";

	private string cellId_ct20_donvi = "Q_51";

	private string cellId_ct20_daily = "Y_51";

	private string cellId_ct10 = "AM_29";

	private string cellId_E13 = "E_13";

	private string cellId_E19 = "E_19";

	private string cellId_E25 = "E_25";

	private string cellId_E28 = "E_28";

	private string cellId_E29 = "E_29";

	private string cellId_E31 = "E_31";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_X31 => GeneralFunctions.ArrayCellId("X_31");

	private int[] CellId_Y31 => GeneralFunctions.ArrayCellId("Y_31");

	private int[] CellId_Z31 => GeneralFunctions.ArrayCellId("Z_31");

	private int[] CellId_AA31 => GeneralFunctions.ArrayCellId("AA_31");

	private int[] CellId_AF23 => GeneralFunctions.ArrayCellId(cellId_AF23);

	private int[] CellId_ct08 => GeneralFunctions.ArrayCellId(cellId_ct08);

	private int[] CellId_ct09 => GeneralFunctions.ArrayCellId(cellId_ct09);

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_ct10 => GeneralFunctions.ArrayCellId(cellId_ct10);

	private int[] CellId_ct38 => GeneralFunctions.ArrayCellId("AV_69");

	private int[] CellId_ct38_CThuc => GeneralFunctions.ArrayCellId("BG_69");

	private int[] CellId_ct39 => GeneralFunctions.ArrayCellId("AV_70");

	private int[] CellId_ct39_CThuc => GeneralFunctions.ArrayCellId("BG_70");

	private int[] CellId_ct40 => GeneralFunctions.ArrayCellId("AV_71");

	private int[] CellId_ct40_CThuc => GeneralFunctions.ArrayCellId("BG_71");

	private int[] CellId_ct41 => GeneralFunctions.ArrayCellId("AV_72");

	private int[] CellId_ct41_CThuc => GeneralFunctions.ArrayCellId("BG_72");

	private int[] CellId_ct42 => GeneralFunctions.ArrayCellId("AV_73");

	private int[] CellId_ct42_CThuc => GeneralFunctions.ArrayCellId("BG_73");

	private int[] CellId_LoaiTien => GeneralFunctions.ArrayCellId("P_55");

	private int[] CellId_MaLoaiTien => GeneralFunctions.ArrayCellId("BG_55");

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
			sheetView.Cells[CellId_AF23[1], CellId_AF23[0]].Text = PeriodInfo.LanXuatBan;
			if (PeriodInfo.TieuMucHachToanVal != null)
			{
				switch (PeriodInfo.TieuMucHachToanVal)
				{
				case "1":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3751";
					sheetView.Cells[CellId_Y31[1], CellId_Y31[0]].Text = "3755";
					sheetView.Cells[CellId_Z31[1], CellId_Z31[0]].Text = "3756";
					sheetView.Cells[CellId_AA31[1], CellId_AA31[0]].Text = "3752";
					break;
				case "3":
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3951";
					sheetView.Cells[CellId_Y31[1], CellId_Y31[0]].Text = "3955";
					sheetView.Cells[CellId_Z31[1], CellId_Z31[0]].Text = "3956";
					sheetView.Cells[CellId_AA31[1], CellId_AA31[0]].Text = "3952";
					break;
				case "2":
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Value = true;
					sheetView.Cells[CellId_X31[1], CellId_X31[0]].Text = "3801";
					sheetView.Cells[CellId_Y31[1], CellId_Y31[0]].Text = "3805";
					sheetView.Cells[CellId_Z31[1], CellId_Z31[0]].Text = "3806";
					sheetView.Cells[CellId_AA31[1], CellId_AA31[0]].Text = "3802";
					break;
				}
			}
			if (Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value))
			{
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].ResetFormula();
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
			}
			else
			{
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
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
			string cellFormat = "AV_66###AV_67###AV_68###AV_69###AV_70###AV_71###AV_72###AV_73###AV_74";
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
				string formula = "IF(AV70>0,AV70,AV71)";
				string formula2 = "IF(AND(1.2*AV65<AV64,1.5*AV65>=AV64),ROUND(0.5*(AV64-1.2*AV65)*AV63,2),0)";
				string formula3 = "IF(AV64>AV65*1.5,AV72+AV73,0)";
				string formula4 = "IF(AV64>AV65*1.5,ROUND(0.5*(1.5*AV65-1.2*AV65)*AV63,2),0)";
				string formula5 = "IF(AV64>AV65*1.5,ROUND(0.6*(AV64-1.5*AV65)*AV63,2),0)";
				if (CellId_LoaiTien[1] + 1 == iRow && CellId_LoaiTien[0] + 1 == iCol)
				{
					string cellFormat = "AV_66###AV_67###AV_68###AV_69###AV_70###AV_71###AV_72###AV_73###AV_74";
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					GeneralFunctions.changFormatCell(sheetView, cellFormat, text);
					if (text == "VND")
					{
						formula = "IF(AV70>0,AV70,AV71)";
						formula2 = "IF(AND(1.2*AV65<AV64,1.5*AV65>=AV64),ROUND(0.5*(AV64-1.2*AV65)*AV63,0),0)";
						formula3 = "IF(AV64>AV65*1.5,AV72+AV73,0)";
						formula4 = "IF(AV64>AV65*1.5,ROUND(0.5*(1.5*AV65-1.2*AV65)*AV63,0),0)";
						formula5 = "IF(AV64>AV65*1.5,ROUND(0.6*(AV64-1.5*AV65)*AV63,0),0)";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula;
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula2;
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula3;
					sheetView.Cells[CellId_ct41[1], CellId_ct41[0]].Formula = formula4;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula5;
				}
				if (iRow >= 63 && iRow <= 65 && CellId_ct39[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula2 = "IF(AND(1.2*AV65<AV64,1.5*AV65>=AV64),ROUND(0.5*(AV64-1.2*AV65)*AV63,0),0)";
						formula5 = "IF(AV64>AV65*1.5,ROUND(0.6*(AV64-1.5*AV65)*AV63,0),0)";
					}
					sheetView.Cells[CellId_ct39[1], CellId_ct39[0]].Formula = formula2;
					sheetView.Cells[CellId_ct42[1], CellId_ct42[0]].Formula = formula5;
				}
				if ((iRow == 63 || iRow == 65) && CellId_ct41[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula4 = "IF(AV64>AV65*1.5,ROUND(0.5*(1.5*AV65-1.2*AV65)*AV63,0),0)";
					}
					sheetView.Cells[CellId_ct41[1], CellId_ct41[0]].Formula = formula4;
				}
				if (((iRow >= 63 && iRow <= 65) || (iRow >= 72 && iRow <= 73)) && CellId_ct39[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula3 = "IF(AV64>AV65*1.5,AV72+AV73,0)";
					}
					sheetView.Cells[CellId_ct40[1], CellId_ct40[0]].Formula = formula3;
				}
				if (((iRow >= 63 && iRow <= 65) || (iRow >= 71 && iRow <= 73)) && CellId_ct39[0] + 1 == iCol)
				{
					string text = sheetView.Cells[CellId_MaLoaiTien[1], CellId_MaLoaiTien[0]].Text;
					if (text == "VND")
					{
						formula = "IF(AV70>0,AV70,AV71)";
					}
					sheetView.Cells[CellId_ct38[1], CellId_ct38[0]].Formula = formula;
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
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value);
			bool flag3 = flag || flag2;
			if (idxRow == CellId_ct08[1] && idxCol == CellId_ct08[0])
			{
				if (!flag)
				{
					sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value = true;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
				}
				else
				{
					sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value = false;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Value = "";
				}
			}
			else if (idxRow == CellId_ct09[1] && idxCol == CellId_ct09[0])
			{
				if (!flag2)
				{
					sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value = true;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
				}
				else
				{
					sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value = false;
					sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
				}
			}
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
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value);
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value);
			bool flag3 = Convert.ToBoolean(sheetView.Cells[CellId_ct20_donvi[1], CellId_ct20_donvi[0]].Value);
			bool flag4 = Convert.ToBoolean(sheetView.Cells[CellId_ct20_daily[1], CellId_ct20_daily[0]].Value);
			if (!flag && !flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct08[1], CellId_ct08[0], message);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09[1], CellId_ct09[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct08[1], CellId_ct08[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct09[1], CellId_ct09[0]);
			}
			HTKKFile.RemoveValidate(sheetView, CellId_ct10[1], CellId_ct10[0]);
			if (flag2)
			{
				string text = sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Text;
				DateTime dtDate;
				if (string.IsNullOrEmpty(text) || text.Equals("../../...."))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct10[1], CellId_ct10[0], message);
					num = 2;
				}
				else if (!HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct10[1], CellId_ct10[0], HTKKFile.GetMessage("0071"));
					num = 2;
				}
			}
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag3 && flag4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message2);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message2);
					num = 2;
				}
				else if (!flag3 && !flag4)
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
			int[] array = GeneralFunctions.ArrayCellId("X_49");
			string text2 = sheetView.Cells[array[1], array[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			if (!string.IsNullOrEmpty(text2) && !ValidateBase.CheckMaSoThue(text2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array[1], array[0], message6);
				num = 2;
			}
			int[] array2 = new int[4] { 59, 60, 61, 64 };
			for (int i = 0; i < array2.Length; i++)
			{
				int[] array3 = GeneralFunctions.ArrayCellId("AN_" + array2[i]);
				HTKKFile.RemoveValidate(sheetView, array3[1], array3[0]);
				string text3 = sheetView.Cells[array3[1], array3[0]].Text;
				if (string.IsNullOrEmpty(text3))
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
