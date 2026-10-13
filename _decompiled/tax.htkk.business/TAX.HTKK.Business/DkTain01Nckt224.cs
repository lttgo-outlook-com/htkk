using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class DkTain01Nckt224
{
	private string cellId_O30 = "O_30";

	private string cellId_P30 = "P_30";

	private string cellId_Q30 = "Q_30";

	private string cellId_W31 = "W_31";

	private string cellId_O31 = "O_31";

	private string cellId_AF23 = "AF_23";

	private string cellId_R30 = "R_30";

	private string cellId_W30 = "W_30";

	private string cellId_ct08 = "Y_28";

	private string cellId_ct09 = "AS_28";

	private string cellId_ct20_donvi = "Q_49";

	private string cellId_ct20_daily = "Y_49";

	private string cellId_ct10 = "AM_29";

	private string cellId_ct35 = "AV_62";

	private string cellId_ct32 = "AV_59";

	private string cellId_PL1_ct9 = "K_19";

	private int[] CellId_O30 => GeneralFunctions.ArrayCellId(cellId_O30);

	private int[] CellId_P30 => GeneralFunctions.ArrayCellId(cellId_P30);

	private int[] CellId_Q30 => GeneralFunctions.ArrayCellId(cellId_Q30);

	private int[] CellId_W31 => GeneralFunctions.ArrayCellId(cellId_W31);

	private int[] CellId_O31 => GeneralFunctions.ArrayCellId(cellId_O31);

	private int[] CellId_AF23 => GeneralFunctions.ArrayCellId(cellId_AF23);

	private int[] CellId_R30 => GeneralFunctions.ArrayCellId(cellId_R30);

	private int[] CellId_W30 => GeneralFunctions.ArrayCellId(cellId_W30);

	private int[] CellId_ct08 => GeneralFunctions.ArrayCellId(cellId_ct08);

	private int[] CellId_ct09 => GeneralFunctions.ArrayCellId(cellId_ct09);

	private int[] CellId_ct20_donvi => GeneralFunctions.ArrayCellId(cellId_ct20_donvi);

	private int[] CellId_ct20_daily => GeneralFunctions.ArrayCellId(cellId_ct20_daily);

	private int[] CellId_ct10 => GeneralFunctions.ArrayCellId(cellId_ct10);

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_PL1_ct9 => GeneralFunctions.ArrayCellId(cellId_PL1_ct9);

	private int idxPL1_ColD => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxPL1_ColL => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxPL1_ColU => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxPL1_ColAB => HTKKFile.ColumnLettersToIndex("AB") - 1;

	private int idxPL1_ColAL => HTKKFile.ColumnLettersToIndex("AL") - 1;

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E28 => GeneralFunctions.ArrayCellId("E_28");

	private int[] CellId_E29 => GeneralFunctions.ArrayCellId("E_29");

	private int[] CellId_E30 => GeneralFunctions.ArrayCellId("E_30");

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
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Text = "1";
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Text = "0";
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Text = "0";
					break;
				case "3":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Text = "0";
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Text = "1";
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Text = "0";
					break;
				case "2":
					sheetView.Cells[CellId_O30[1], CellId_O30[0]].Text = "0";
					sheetView.Cells[CellId_P30[1], CellId_P30[0]].Text = "0";
					sheetView.Cells[CellId_Q30[1], CellId_Q30[0]].Text = "1";
					break;
				}
			}
			if (PeriodInfo.KieuKy == "D")
			{
				sheetView.Cells[CellId_AF23[1], CellId_AF23[0]].Text = PeriodInfo.LanXuatBan;
			}
			bool flag = false;
			bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_ct08[1], CellId_ct08[0]].Value);
			flag = Convert.ToBoolean(sheetView.Cells[CellId_ct09[1], CellId_ct09[0]].Value);
			if (!flag2 && !flag)
			{
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
			}
			else if (flag2)
			{
				GeneralFunctions.ChangeValueCell(sheetView, CellId_ct10[1], CellId_ct10[0], locked: true, "", "");
			}
			else
			{
				sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
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
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
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
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E22);
			string message3 = HTKKFile.GetMessage("0071");
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E29);
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
			else if (flag && flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct08[1], CellId_ct08[0], message2);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct09[1], CellId_ct09[0], message2);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct08[1], CellId_ct08[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct09[1], CellId_ct09[0]);
			}
			if (flag2)
			{
				string text = sheetView.Cells[CellId_ct10[1], CellId_ct10[0]].Text;
				DateTime dtDate;
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct10[1], CellId_ct10[0], message);
					num = 2;
				}
				else if (!HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct10[1], CellId_ct10[0], message3);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct10[1], CellId_ct10[0]);
				}
				double num2 = Convert.ToDouble(sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Value);
				if (num2 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct35[1], CellId_ct35[0], message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct35[1], CellId_ct35[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct10[1], CellId_ct10[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct35[1], CellId_ct35[0]);
			}
			if (!HTKKFormater.isNullString(TTinChung.mstDLyThue))
			{
				if (flag3 && flag4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message4);
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message4);
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
			else if (flag3 && flag4)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0], message5);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0], message5);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct20_daily[1], CellId_ct20_daily[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct20_donvi[1], CellId_ct20_donvi[0]);
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
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E13);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E28);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E30);
			HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.Tables != null && hTKKCMCSheetInfo.Tables.Count > 0) ? hTKKCMCSheetInfo.Tables[0] : null);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PathU29");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PLuc_01_DK/tongCong");
			bool flag = false;
			if (hTKKCMCTableInfo != null)
			{
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
				for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsValues.Count; num3++)
				{
					array[0] = sheetView.Cells[num2 + num3, idxPL1_ColD].Text;
					array[1] = sheetView.Cells[num2 + num3, idxPL1_ColL].Text;
					array[2] = sheetView.Cells[num2 + num3, idxPL1_ColU].Text;
					array[4] = sheetView.Cells[num2 + num3, idxPL1_ColAL].Text;
					string text = array[0] + array[1] + array[2] + array[4];
					if (string.IsNullOrEmpty(text) || text.Equals("0,0000"))
					{
						HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColD);
						HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColL);
						HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
						return;
					}
					bool flag2 = GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array);
					if (flag2)
					{
						flag = true;
					}
					if (!flag2)
					{
						HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColD);
						HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColL);
						HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
						continue;
					}
					if (string.IsNullOrEmpty(array[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxPL1_ColD, message);
						num = 2;
					}
					else if (!ValidateBase.CheckMaSoThue(array[0]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxPL1_ColD, message2);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColD);
					}
					if (string.IsNullOrEmpty(array[1]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + num3, idxPL1_ColL, message);
						num = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num2 + num3, idxPL1_ColL);
					}
				}
				double num4 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
				if (num4 != 100.0 && flag)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message3);
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message4);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
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

	private void XuLyCT22_ct5(SheetView sViewTK)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
