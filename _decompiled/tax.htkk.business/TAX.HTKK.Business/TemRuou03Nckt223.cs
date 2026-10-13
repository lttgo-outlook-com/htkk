using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TemRuou03Nckt223
{
	private string cellId_ngayMat = "K_9";

	private string cellId_ngayLap = "Y_56";

	private string cellId_lyDoMatTr = "I_47";

	private string cellId_nguoiKy = "Y_54";

	private string fmFColX = "IF(AND(V{0}<>0,S{0}<>0),(V{0}-S{0}+1), 0)";

	private string cellId_E14 = "E_14";

	private string cellId_E15 = "E_15";

	private string cellId_E16 = "E_16";

	private string cellId_E17 = "E_17";

	private string cellId_E18 = "E_18";

	private string cellId_E19 = "E_19";

	private string cellId_E21 = "E_21";

	private string cellId_E22 = "E_22";

	private string cellId_E23 = "E_23";

	private string cellId_E24 = "E_24";

	private string cellId_E25 = "E_25";

	private int[] CellId_ngayMat => GeneralFunctions.ArrayCellId(cellId_ngayMat);

	private int[] CellId_ngayLap => GeneralFunctions.ArrayCellId(cellId_ngayLap);

	private int[] CellId_lyDoMatTr => GeneralFunctions.ArrayCellId(cellId_lyDoMatTr);

	private int[] CellId_nguoiKy => GeneralFunctions.ArrayCellId(cellId_nguoiKy);

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_H => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_O => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_S => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_Y => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCol_X => HTKKFile.ColumnLettersToIndex("X") - 1;

	private int[] CellId_E14 => GeneralFunctions.ArrayCellId(cellId_E14);

	private int[] CellId_E15 => GeneralFunctions.ArrayCellId(cellId_E15);

	private int[] CellId_E16 => GeneralFunctions.ArrayCellId(cellId_E16);

	private int[] CellId_E17 => GeneralFunctions.ArrayCellId(cellId_E17);

	private int[] CellId_E18 => GeneralFunctions.ArrayCellId(cellId_E18);

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId(cellId_E19);

	private int[] CellId_E21 => GeneralFunctions.ArrayCellId(cellId_E21);

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId(cellId_E22);

	private int[] CellId_E23 => GeneralFunctions.ArrayCellId(cellId_E23);

	private int[] CellId_E24 => GeneralFunctions.ArrayCellId(cellId_E24);

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId(cellId_E25);

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("K");
			int num2 = 9;
			string text = sView.Cells[num2 - 1, num - 1].Text;
			if (string.IsNullOrEmpty(text))
			{
				sView.Cells[num2 - 1, num - 1].Text = DateTime.Now.Date.ToString("dd/MM/yyyy");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num && idxRow <= num2 && (idxCol == idxCol_S || idxCol == idxCol_V))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LeaveCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, long iNewRow, long iNewCol)
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_O)
			{
				string str = sheetView.Cells[(int)iRow - 1, idxCol_O].Text.Trim();
				sheetView.Cells[(int)iRow - 1, idxCol_O].Text = GeneralFunctions.FormatKyHieu(str);
			}
			if (iCol - 1 == idxCol_S || iCol - 1 == idxCol_V)
			{
				string text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				if (!HTKKFormater.isNullString(text))
				{
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				}
			}
			if (iCol - 1 != idxCol_V && iCol - 1 != idxCol_Y)
			{
				return;
			}
			string text2 = sheetView.Cells[(int)iRow - 1, idxCol_S].Text;
			string text3 = sheetView.Cells[(int)iRow - 1, idxCol_V].Text;
			int num3 = ConvertHelpers.ConvertToInt(text2);
			int num4 = ConvertHelpers.ConvertToInt(text3);
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text3))
			{
				if (num3 == 0 && num4 == 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = 0;
				}
				else if (num3 == 0 && num4 != 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num4 - num3;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num4 - num3 + 1;
				}
			}
			else
			{
				sheetView.Cells[(int)iRow - 1, idxCol_X].Value = 0;
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
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string message = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E15);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E17);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E18);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E16);
			string message6 = HTKKFile.GetMessage("0071");
			string message7 = HTKKFile.GetMessage(sheetView, CellId_E23);
			string message8 = HTKKFile.GetMessage(sheetView, CellId_E22);
			bool flag = true;
			string text = sheetView.Cells[CellId_ngayMat[1], CellId_ngayMat[0]].Text.Trim();
			DateTime dtDate;
			bool flag2 = HTKKFormater.IsDateTime(text, "dd/MM/yyyy", out dtDate);
			if (string.IsNullOrEmpty(text))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ngayMat[1], CellId_ngayMat[0], message);
				num = 2;
			}
			else if (!flag2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ngayMat[1], CellId_ngayMat[0], message6);
				num = 2;
			}
			else if (dtDate < new DateTime(2014, 1, 1))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ngayMat[1], CellId_ngayMat[0], HTKKFile.GetMessage(sheetView, CellId_E24));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ngayMat[1], CellId_ngayMat[0]);
				flag = false;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string[] array = null;
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				array = GeneralFunctions.checkTrung(winFpSpread, num2, "S", "V", "B", "H", "E", "O", "", isMST: false, 1);
			}
			string[] array2 = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array2[0] = sheetView.Cells[num2 + i, idxCol_C].Text;
				array2[2] = sheetView.Cells[num2 + i, idxCol_H].Text;
				array2[3] = sheetView.Cells[num2 + i, idxCol_O].Text;
				array2[4] = sheetView.Cells[num2 + i, idxCol_S].Text;
				array2[5] = sheetView.Cells[num2 + i, idxCol_V].Text;
				array2[7] = sheetView.Cells[num2 + i, idxCol_Y].Text;
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array2))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_C);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_H);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_O);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_S);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_V);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_Y);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_C);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_H);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_O);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_S);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_V);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_Y);
				if (HTKKFormater.isNullString(array2[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_C, message);
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_H, message);
					num = 2;
				}
				if (HTKKFormater.isNullString(array2[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_O, message);
					num = 2;
				}
				else
				{
					string text2 = GeneralFunctions.CheckSoTemRuou(array2[3]);
					if (text2.Equals("1") || text2.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_O, message2);
						num = 2;
					}
				}
				if (HTKKFormater.isNullString(array2[4]) && HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_S, message);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_V, message);
					num = 2;
				}
				else if (HTKKFormater.isNullString(array2[4]) && !HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_S, message);
					num = 2;
				}
				else if (!HTKKFormater.isNullString(array2[4]) && HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_V, message);
					num = 2;
				}
				else
				{
					double num3 = Convert.ToDouble(array2[4]);
					double num4 = Convert.ToDouble(array2[5]);
					if ((!(num4 - num3 >= 0.0) || !(num4 > 0.0)) && (num3 != 0.0 || num4 != 0.0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_S, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_V, message4);
						num = 2;
					}
				}
				if (HTKKFormater.isNullString(array2[7]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_Y, message7);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_Y);
				}
				if (array == null || array.Length <= 1)
				{
					continue;
				}
				for (int j = 0; j < array.Length; j++)
				{
					if (num2 + i == Convert.ToInt32(array[j]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_S, message5);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_V, message5);
						num = 2;
						break;
					}
				}
			}
			int num5 = CellId_lyDoMatTr[1] + hTKKCMCTableInfo.RowsValues.Count - 1;
			string str = sheetView.Cells[num5, CellId_lyDoMatTr[0]].Text.Trim();
			if (HTKKFormater.isNullString(str))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num5, CellId_lyDoMatTr[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num5, CellId_lyDoMatTr[0]);
			}
			int num6 = CellId_nguoiKy[1] + hTKKCMCTableInfo.RowsValues.Count - 1;
			string str2 = sheetView.Cells[num6, CellId_nguoiKy[0]].Text.Trim();
			if (HTKKFormater.isNullString(str2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num6, CellId_nguoiKy[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num6, CellId_nguoiKy[0]);
			}
			int num7 = CellId_ngayLap[1] + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text3 = sheetView.Cells[num7, CellId_ngayLap[0]].Text.Trim();
			DateTime dtDate2;
			bool flag3 = HTKKFormater.IsDateTime(text3, "dd/MM/yyyy", out dtDate2);
			if (string.IsNullOrEmpty(text3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num7, CellId_ngayLap[0], message);
				num = 2;
			}
			else if (!flag3)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num7, CellId_ngayLap[0], message6);
				num = 2;
			}
			else if (dtDate2 > DateTime.Now)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num7, CellId_ngayLap[0], HTKKFile.GetMessage(sheetView, CellId_E21));
				num = 2;
			}
			else if (!flag)
			{
				if (DateTime.Compare(dtDate2, dtDate) < 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ngayMat[1], CellId_ngayMat[0], HTKKFile.GetMessage(sheetView, CellId_E22));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ngayMat[1], CellId_ngayMat[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num7, CellId_ngayLap[0]);
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
			if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num3 - 1].Text;
				sheetView.Cells[num + i, num2 - 1].Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
