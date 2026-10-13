using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TemRuou06Nckt223
{
	private List<Dictionary<string, string>> rowValues;

	private string cellId_nguoiKy = "N_33";

	private string cellId_ngayLap = "N_35";

	private string cellId_nguoiLap = "D_33";

	private string pathCt2 = "ct2";

	private string pathCt3 = "ct3";

	private string pathCt4 = "ct4";

	private string pathCt18 = "ct18";

	private string pathCt19 = "ct19";

	private string pathCt20 = "ct20";

	private int[] CellId_nguoiKy => GeneralFunctions.ArrayCellId(cellId_nguoiKy);

	private int[] CellId_ngayLap => GeneralFunctions.ArrayCellId(cellId_ngayLap);

	private int[] CellId_nguoiLap => GeneralFunctions.ArrayCellId(cellId_nguoiLap);

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_D => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_E => HTKKFile.ColumnLettersToIndex("E") - 1;

	private int idxCol_F => HTKKFile.ColumnLettersToIndex("F") - 1;

	private int idxCol_G => HTKKFile.ColumnLettersToIndex("G") - 1;

	private List<int> lstRefCol5
	{
		get
		{
			List<int> list = new List<int>();
			list.Add(idxCol_H);
			list.Add(idxCol_I);
			list.Add(idxCol_J);
			list.Add(idxCol_K);
			return list;
		}
	}

	private int idxCol_H => HTKKFile.ColumnLettersToIndex("H") - 1;

	private int idxCol_I => HTKKFile.ColumnLettersToIndex("I") - 1;

	private int idxCol_J => HTKKFile.ColumnLettersToIndex("J") - 1;

	private int idxCol_K => HTKKFile.ColumnLettersToIndex("K") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_M => HTKKFile.ColumnLettersToIndex("M") - 1;

	private int idxCol_N => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_P => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_Q => HTKKFile.ColumnLettersToIndex("Q") - 1;

	private int idxCol_R => HTKKFile.ColumnLettersToIndex("R") - 1;

	private int idxCol_S => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_T => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_W => HTKKFile.ColumnLettersToIndex("W") - 1;

	private int idxCol_X => HTKKFile.ColumnLettersToIndex("X") - 1;

	private int idxCol_AA => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int[] CellId_E13 => GeneralFunctions.ArrayCellId("E_13");

	private int[] CellId_E14 => GeneralFunctions.ArrayCellId("E_14");

	private int[] CellId_E15 => GeneralFunctions.ArrayCellId("E_15");

	private int[] CellId_E16 => GeneralFunctions.ArrayCellId("E_16");

	private int[] CellId_E17 => GeneralFunctions.ArrayCellId("E_17");

	private int[] CellId_E18 => GeneralFunctions.ArrayCellId("E_18");

	private int[] CellId_E19 => GeneralFunctions.ArrayCellId("E_19");

	private int[] CellId_E20 => GeneralFunctions.ArrayCellId("E_20");

	private int[] CellId_E21 => GeneralFunctions.ArrayCellId("E_21");

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E23 => GeneralFunctions.ArrayCellId("E_23");

	private int[] CellId_E24 => GeneralFunctions.ArrayCellId("E_24");

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E26 => GeneralFunctions.ArrayCellId("E_26");

	private int[] CellId_E35 => GeneralFunctions.ArrayCellId("E_35");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.Tables.Count > 0) ? hTKKCMCSheetInfo.Tables[0] : null);
			string value = "";
			GeneralFunctions.dicVar.TryGetValue("key_action", out value);
			if (!hTKKCMCSheetInfo.IsReadDataFile)
			{
				LayGiaTriTonKyTruoc(sView, PeriodInfo, HeaderInfo, ToKhaiInfo);
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
			if (idxSheet != 0)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (idxRow >= num && idxRow <= num2)
			{
				if (idxCol == idxCol_R || idxCol == idxCol_T)
				{
					GeneralFunctions.CompareKeycode(e, 0);
				}
				else if (idxCol == idxCol_H || idxCol == idxCol_I || idxCol == idxCol_J || idxCol == idxCol_K)
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
			string text = "";
			int num3 = 0;
			string[] array = new string[0];
			if (iCol - 1 == idxCol_F)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (lstRefCol5.Contains((int)iCol - 1) || iCol - 1 == idxCol_P || iCol - 1 == idxCol_R || iCol - 1 == idxCol_T)
			{
				if (lstRefCol5.Contains((int)iCol - 1))
				{
					text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
				}
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_P + 1, "16");
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_Q + 1, "16");
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_S + 1, "16");
				int num4 = ConvertHelpers.ConvertToInt(valueCustom);
				int num5 = ConvertHelpers.ConvertToInt(valueCustom2);
				int num6 = ConvertHelpers.ConvertToInt(valueCustom3);
				string text2 = sheetView.Cells[(int)iRow - 1, idxCol_H].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, idxCol_I].Text;
				string text4 = sheetView.Cells[(int)iRow - 1, idxCol_J].Text;
				string text5 = sheetView.Cells[(int)iRow - 1, idxCol_K].Text;
				array = GeneralFunctions.getTusoDenso(text2, text3, text4, text5);
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_N + 1, "16");
				int num7;
				if (!array[0].Equals("-1"))
				{
					if (num4 + num5 + num6 > 0)
					{
						sheetView.Cells[(int)iRow - 1, idxCol_L].Text = GeneralFunctions.FormatSoHD(array[0]);
						num7 = ConvertHelpers.ConvertToInt(valueCustom4);
						int num8 = ConvertHelpers.ConvertToInt(array[0]);
						sheetView.Cells[(int)iRow - 1, idxCol_M].Text = GeneralFunctions.FormatSoHD((num8 + num7 - 1).ToString());
					}
					else
					{
						sheetView.Cells[(int)iRow - 1, idxCol_L].Text = "";
					}
					sheetView.Cells[(int)iRow - 1, idxCol_G].Value = array[2];
				}
				string text6 = sheetView.Cells[(int)iRow - 1, idxCol_M].Text;
				valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_N + 1, "16");
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_G + 1, "16");
				int num9 = ConvertHelpers.ConvertToInt(text6);
				num7 = ConvertHelpers.ConvertToInt(valueCustom4);
				int num10 = ConvertHelpers.ConvertToInt(valueCustom5);
				if (num10 == num7 && num4 + num6 + num5 > 0)
				{
					num3 = num10 - num7;
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
				else if (num10 > num7 && num4 + num6 + num5 > 0)
				{
					num3 = num10 - num7;
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = GeneralFunctions.FormatSoHD((num9 + 1).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = GeneralFunctions.FormatSoHD((num9 + num3).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
				else if (num10 > num7 && num4 + num6 + num5 == 0 && text2.Equals("") && text3.Equals(""))
				{
					num3 = num10;
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = GeneralFunctions.FormatSoHD(text4);
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = GeneralFunctions.FormatSoHD(text5);
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
				else if (num10 > num7 && num4 + num6 + num5 == 0 && text4.Equals("") && text5.Equals(""))
				{
					num3 = num10;
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = GeneralFunctions.FormatSoHD(text2);
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = GeneralFunctions.FormatSoHD(text3);
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
				else
				{
					num3 = num10;
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = GeneralFunctions.FormatSoHD(text2);
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = GeneralFunctions.FormatSoHD(text5);
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
			}
			if (iCol - 1 == idxCol_H || iCol - 1 == idxCol_I || iCol - 1 == idxCol_J || iCol - 1 == idxCol_K || iCol - 1 == idxCol_P || iCol - 1 == idxCol_R || iCol - 1 == idxCol_T)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				string soLuongHD = GeneralFunctions.GetSoLuongHD(text);
				if (iCol - 1 == idxCol_R)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_Q].Value = soLuongHD;
				}
				if (iCol - 1 == idxCol_T)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_S].Value = soLuongHD;
				}
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_P + 1, "16");
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_Q + 1, "16");
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_S + 1, "16");
				int num4 = ConvertHelpers.ConvertToInt(valueCustom);
				int num5 = ConvertHelpers.ConvertToInt(valueCustom2);
				int num6 = ConvertHelpers.ConvertToInt(valueCustom3);
				string text7 = sheetView.Cells[(int)iRow - 1, idxCol_L].Text;
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_N + 1, "16");
				int num7 = ConvertHelpers.ConvertToInt(valueCustom4);
				int num8 = ConvertHelpers.ConvertToInt(text7);
				int num9 = num8 + num7 - 1;
				if (num6 + num5 + num4 > 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_M].Text = GeneralFunctions.FormatSoHD(num9.ToString());
					if (!array[0].Equals("-1"))
					{
						sheetView.Cells[(int)iRow - 1, idxCol_L].Text = GeneralFunctions.FormatSoHD(array[0]);
					}
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_M].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_L].Text = "";
				}
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_G + 1, "16");
				int num10 = ConvertHelpers.ConvertToInt(valueCustom5);
				if (num10 > num7 && num4 + num5 + num6 > 0)
				{
					num3 = num10 - num7;
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = GeneralFunctions.FormatSoHD((num9 + 1).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = GeneralFunctions.FormatSoHD((num9 + num3).ToString());
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
			}
			if (lstRefCol5.Contains((int)iCol - 1) || iCol - 1 == idxCol_P || iCol - 1 == idxCol_R || iCol - 1 == idxCol_T)
			{
				string valueCustom5 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_G + 1, "16");
				string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_N + 1, "16");
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_P + 1, "16");
				string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_Q + 1, "16");
				string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, idxCol_S + 1, "16");
				int num10 = ConvertHelpers.ConvertToInt(valueCustom5);
				int num7 = ConvertHelpers.ConvertToInt(valueCustom4);
				int num4 = ConvertHelpers.ConvertToInt(valueCustom);
				int num5 = ConvertHelpers.ConvertToInt(valueCustom2);
				int num6 = ConvertHelpers.ConvertToInt(valueCustom3);
				if (num10 + num7 + num4 + num5 + num6 == 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
				else if (num10 == num7 && num6 + num5 + num4 > 0)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = num3.ToString();
				}
				else if (num10 < num7)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_V].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_W].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_X].Value = "0";
				}
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
			string message = HTKKFile.GetMessage(sheetView, CellId_E13);
			string message2 = HTKKFile.GetMessage(sheetView, CellId_E14);
			string message3 = HTKKFile.GetMessage(sheetView, CellId_E15);
			string message4 = HTKKFile.GetMessage(sheetView, CellId_E16);
			string message5 = HTKKFile.GetMessage(sheetView, CellId_E18);
			string message6 = HTKKFile.GetMessage(sheetView, CellId_E19);
			string message7 = HTKKFile.GetMessage(sheetView, CellId_E20);
			string message8 = HTKKFile.GetMessage(sheetView, CellId_E21);
			string message9 = HTKKFile.GetMessage(sheetView, CellId_E22);
			string message10 = HTKKFile.GetMessage(sheetView, CellId_E23);
			string message11 = HTKKFile.GetMessage(sheetView, CellId_E24);
			string message12 = HTKKFile.GetMessage(sheetView, CellId_E25);
			string message13 = HTKKFile.GetMessage(sheetView, CellId_E26);
			string message14 = HTKKFile.GetMessage(sheetView, CellId_E35);
			string message15 = HTKKFile.GetMessage("0071");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			double num3 = 0.0;
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num2, "H", "I", "B", "E", "D", "F", "", isMST: false, 1);
			string[] array2 = GeneralFunctions.checkTrung(winFpSpread, num2, "J", "K", "B", "E", "D", "F", "", isMST: false, 1);
			string[] array3 = GeneralFunctions.checkTrung(winFpSpread, num2, "V", "W", "B", "E", "D", "F", "", isMST: false, 1);
			string[] array4 = new string[hTKKCMCTableInfo.RowsInfo.Cells.Count];
			int num4 = -1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num2 + i, idxCol_AA].Text.Trim();
				array4[0] = sheetView.Cells[num2 + i, idxCol_C].Text.Trim();
				array4[2] = sheetView.Cells[num2 + i, idxCol_E].Text.Trim();
				array4[3] = sheetView.Cells[num2 + i, idxCol_F].Text.Trim();
				array4[4] = GeneralFunctions.GetValueCustom(sheetView, num2 + i + 1, idxCol_G + 1, "16");
				array4[5] = sheetView.Cells[num2 + i, idxCol_H].Text.Trim();
				array4[6] = sheetView.Cells[num2 + i, idxCol_I].Text.Trim();
				array4[7] = sheetView.Cells[num2 + i, idxCol_J].Text.Trim();
				array4[8] = sheetView.Cells[num2 + i, idxCol_K].Text.Trim();
				array4[9] = sheetView.Cells[num2 + i, idxCol_L].Text.Trim();
				array4[10] = sheetView.Cells[num2 + i, idxCol_M].Text.Trim();
				array4[11] = GeneralFunctions.GetValueCustom(sheetView, num2 + i + 1, idxCol_N + 1, "16");
				array4[12] = GeneralFunctions.GetValueCustom(sheetView, num2 + i + 1, idxCol_P + 1, "16");
				array4[14] = sheetView.Cells[num2 + i, idxCol_R].Text.Trim();
				array4[16] = sheetView.Cells[num2 + i, idxCol_T].Text.Trim();
				array4[17] = sheetView.Cells[num2 + i, idxCol_V].Text.Trim();
				array4[18] = sheetView.Cells[num2 + i, idxCol_W].Text.Trim();
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, array4))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_C);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_E);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_F);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_C);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_E);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_F);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_H);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_I);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_J);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_K);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_P);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_R);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_T);
				Color backColor = sheetView.Cells[num2 + i, idxCol_X].BackColor;
				sheetView.Cells[num2 + i, idxCol_V].BackColor = backColor;
				sheetView.Cells[num2 + i, idxCol_V].Note = "";
				sheetView.Cells[num2 + i, idxCol_W].BackColor = backColor;
				sheetView.Cells[num2 + i, idxCol_W].Note = "";
				if (text.Equals("1"))
				{
					num4++;
					if (rowValues != null && rowValues.Count > 0)
					{
						Dictionary<string, string> dictionary = rowValues[num4];
						if (!dictionary[pathCt18].Equals(array4[5]) || !dictionary[pathCt19].Equals(array4[6]))
						{
							HTKKFile.SetNodeBackGroundYellow(sheetView, num2 + i, idxCol_H, message14);
							HTKKFile.SetNodeBackGroundYellow(sheetView, num2 + i, idxCol_I, message14);
							if (num != 2)
							{
								num = 1;
							}
						}
					}
				}
				if (string.IsNullOrEmpty(array4[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_C, message);
					num = 2;
				}
				if (string.IsNullOrEmpty(array4[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_E, message);
					num = 2;
				}
				if (string.IsNullOrEmpty(array4[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_F, message);
					num = 2;
				}
				else
				{
					string text2 = GeneralFunctions.CheckSoTemRuou(array4[3]);
					if (text2 == "1")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_F, message3);
						num = 2;
					}
					else if (text2 == "2")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_F, message6);
						num = 2;
					}
				}
				string[] array5 = GeneralFunctions.checkSoTonSoPH(array4[5], array4[6], array4[7], array4[8]);
				if (array5[0].Equals("1"))
				{
					if (array5[1].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_H, message);
						num = 2;
					}
					else if (array5[1].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_I, message);
						num = 2;
					}
					else if (array5[1].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_H, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_I, message4);
						num = 2;
					}
					else if (array5[1].Equals("4"))
					{
						flag = GeneralFunctions.isDuLieuHL(array4[14]);
						if (!HTKKFormater.isNullString(array4[14]))
						{
							if (flag)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[14]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message8);
										num = 2;
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message9);
								num = 2;
							}
						}
						flag = GeneralFunctions.isDuLieuHL(array4[16]);
						if (!HTKKFormater.isNullString(array4[16]))
						{
							if (flag)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[16]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message8);
										num = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message9);
								num = 2;
							}
						}
						if ((!HTKKFormater.isNullString(array4[14]) || !HTKKFormater.isNullString(array4[16])) && !flag3 && !flag2)
						{
							string text3 = GeneralFunctions.checkTrungXoa("", array4[14], array4[16]);
							if (text3.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message8);
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message8);
								num = 2;
							}
						}
					}
				}
				else if (array5[0].Equals("2"))
				{
					if (array5[2].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_J, message);
						num = 2;
					}
					else if (array5[2].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_K, message);
						num = 2;
					}
					else if (array5[2].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_J, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_K, message4);
						num = 2;
					}
					else if (array5[2].Equals("4"))
					{
						flag = GeneralFunctions.isDuLieuHL(array4[14]);
						if (!HTKKFormater.isNullString(array4[14]))
						{
							if (flag)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[14]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message8);
										num = 2;
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message9);
								num = 2;
							}
						}
						flag = GeneralFunctions.isDuLieuHL(array4[16]);
						if (!HTKKFormater.isNullString(array4[16]))
						{
							if (flag)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[16]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message8);
										num = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message9);
								num = 2;
							}
						}
						if ((!HTKKFormater.isNullString(array4[14]) || !HTKKFormater.isNullString(array4[16])) && !flag3 && !flag2)
						{
							string text3 = GeneralFunctions.checkTrungXoa("", array4[14], array4[16]);
							if (text3.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message8);
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message8);
								num = 2;
							}
						}
					}
				}
				else if (array5[0].Equals("3"))
				{
					if (array5[1].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_H, message);
						num = 2;
					}
					else if (array5[1].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_I, message);
						num = 2;
					}
					else if (array5[1].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_H, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_I, message4);
						num = 2;
					}
					if (array5[2].Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_J, message);
						num = 2;
					}
					else if (array5[2].Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_K, message);
						num = 2;
					}
					else if (array5[2].Equals("3"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_J, message4);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_K, message4);
						num = 2;
					}
					if (array5[3].Equals("1") && !array4[6].Equals("0000000"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_I, message7);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_J, message7);
						num = 2;
					}
					if (array5[1].Equals("4") && array5[2].Equals("4"))
					{
						flag = GeneralFunctions.isDuLieuHL(array4[14]);
						if (!HTKKFormater.isNullString(array4[14]))
						{
							if (flag)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[14]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message8);
										num = 2;
										flag2 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message9);
								num = 2;
							}
						}
						flag = GeneralFunctions.isDuLieuHL(array4[16]);
						if (!HTKKFormater.isNullString(array4[16]))
						{
							if (flag)
							{
								if (!HTKKFormater.isNullString(array4[9]) && !HTKKFormater.isNullString(array4[10]))
								{
									string text3 = GeneralFunctions.checkHDXoa(array4[9], array4[10], array4[16]);
									if (text3.Equals("-1"))
									{
										HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message8);
										num = 2;
										flag3 = true;
									}
								}
							}
							else
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message9);
								num = 2;
							}
						}
						if ((!HTKKFormater.isNullString(array4[14]) || !HTKKFormater.isNullString(array4[16])) && !flag3 && !flag2)
						{
							string text3 = GeneralFunctions.checkTrungXoa("", array4[14], array4[16]);
							if (text3.Equals("3"))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message8);
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message8);
								num = 2;
							}
						}
					}
				}
				double num5 = ConvertHelpers.ConvertToDouble(array4[4]);
				double num6 = ConvertHelpers.ConvertToDouble(array4[12]);
				if (num6 > num5 && !HTKKFormater.isNullString(array4[12]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_P, message10);
					num = 2;
				}
				if (!HTKKFormater.isNullString(array4[14]))
				{
					num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(array4[14]));
					if (num3 > num5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message12);
						num = 2;
					}
				}
				if (!HTKKFormater.isNullString(array4[16]))
				{
					num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetSoLuongHD(array4[16]));
					if (num3 > num5)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message13);
						num = 2;
					}
				}
				double num7 = ConvertHelpers.ConvertToDouble(array4[11]);
				if (num7 > num5)
				{
					if (!HTKKFormater.isNullString(array4[12]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_P, message11);
					}
					if (!HTKKFormater.isNullString(array4[14]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_R, message11);
					}
					if (!HTKKFormater.isNullString(array4[16]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message11);
					}
					num = 2;
				}
				if (!HTKKFormater.isNullString(array4[5]) && !HTKKFormater.isNullString(array4[6]) && array.Length > 1)
				{
					for (int j = 0; j < array.Length; j++)
					{
						if (num2 + i == Convert.ToInt32(array[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_H, message2);
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_I, message2);
							num = 2;
							break;
						}
					}
				}
				if (!HTKKFormater.isNullString(array4[7]) && !HTKKFormater.isNullString(array4[8]) && array2.Length > 1)
				{
					for (int j = 0; j < array2.Length; j++)
					{
						if (num2 + i == Convert.ToInt32(array2[j]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_J, message2);
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_K, message2);
							num = 2;
							break;
						}
					}
				}
				if (HTKKFormater.isNullString(array4[17]) || HTKKFormater.isNullString(array4[18]) || array3.Length <= 1)
				{
					continue;
				}
				for (int j = 0; j < array3.Length; j++)
				{
					if (num2 + i == Convert.ToInt32(array3[j]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_V, message2);
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_W, message2);
						num = 2;
						break;
					}
				}
			}
			int num8 = CellId_nguoiKy[1] + hTKKCMCTableInfo.RowsValues.Count - 1;
			string value2 = sheetView.Cells[num8, CellId_nguoiKy[0]].Text.Trim();
			if (string.IsNullOrEmpty(value2))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num8, CellId_nguoiKy[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num8, CellId_nguoiKy[0]);
			}
			int num9 = CellId_nguoiLap[1] + hTKKCMCTableInfo.RowsValues.Count - 1;
			string value3 = sheetView.Cells[num9, CellId_nguoiLap[0]].Text.Trim();
			if (string.IsNullOrEmpty(value3))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num9, CellId_nguoiLap[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num9, CellId_nguoiLap[0]);
			}
			int num10 = CellId_ngayLap[1] + hTKKCMCTableInfo.RowsValues.Count - 1;
			string text4 = sheetView.Cells[num10, CellId_ngayLap[0]].Text.Trim();
			DateTime dtDate;
			bool flag4 = HTKKFormater.IsDateTime(text4, "dd/MM/yyyy", out dtDate);
			if (string.IsNullOrEmpty(text4))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num10, CellId_ngayLap[0], message);
				num = 2;
			}
			else if (!flag4)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num10, CellId_ngayLap[0], message15);
				num = 2;
			}
			else if (DateTime.Compare(dtDate, DateTime.Now) > 0)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num10, CellId_ngayLap[0], message5);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num10, CellId_ngayLap[0]);
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
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text = sheetView.Cells[num + i, num2 - 1].Text;
				sheetView.Cells[num + i, num3 - 1].Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void LayGiaTriTonKyTruoc(SheetView sViewTK, HTKKPeriodInfo periodInfo, HTKKHeaderInfo headerInfo, HTKKCMCToKhaiInfo tokhaiInfo)
	{
		try
		{
			string preperiodFileName_Quy = GeneralFunctions.GetPreperiodFileName_Quy(periodInfo, tokhaiInfo);
			if (string.IsNullOrEmpty(preperiodFileName_Quy))
			{
				return;
			}
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			bool flag = false;
			FileInfo[] files = directoryInfo.GetFiles();
			if (files.Length > 0)
			{
				FileInfo[] array = files;
				foreach (FileInfo fileInfo in array)
				{
					if (fileInfo.Name.Contains(preperiodFileName_Quy))
					{
						flag = true;
						break;
					}
				}
			}
			if (!flag)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = tokhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			List<string> list = new List<string>();
			list.Add(pathCt2);
			list.Add(pathCt3);
			list.Add(pathCt4);
			list.Add(pathCt18);
			list.Add(pathCt19);
			list.Add(pathCt20);
			List<string> lstCellPath = list;
			rowValues = GetRowDataValueInXml($"{text}\\{preperiodFileName_Quy}.xml", hTKKCMCTableInfo, lstCellPath);
			int count = rowValues.Count;
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int j = 0; j < count; j++)
			{
				sViewTK.Rows.Add(dataLocationTemplateRowStart, 1);
				sViewTK.CopyRange(dataLocationTemplateRowStart - 1, -1, dataLocationTemplateRowStart, -1, 1, -1, dataOnly: false);
				hTKKCMCTableInfo.addNewDefaultRow(j + 1);
			}
			hTKKCMCSheetInfo.IncrementStaticCell(dataLocationTemplateRowStart, count);
			int num = dataLocationTemplateRowStart - 1;
			for (int j = 0; j < rowValues.Count; j++)
			{
				sViewTK.Cells[num + j, idxCol_C].Text = rowValues[j][pathCt2];
				sViewTK.Cells[num + j, idxCol_D].Text = rowValues[j][pathCt2];
				sViewTK.Cells[num + j, idxCol_E].Text = rowValues[j][pathCt3];
				sViewTK.Cells[num + j, idxCol_F].Text = rowValues[j][pathCt4];
				sViewTK.Cells[num + j, idxCol_G].Value = rowValues[j][pathCt20];
				sViewTK.Cells[num + j, idxCol_H].Text = rowValues[j][pathCt18];
				sViewTK.Cells[num + j, idxCol_I].Text = rowValues[j][pathCt19];
				sViewTK.Cells[num + j, idxCol_V].Text = rowValues[j][pathCt18];
				sViewTK.Cells[num + j, idxCol_W].Text = rowValues[j][pathCt19];
				sViewTK.Cells[num + j, idxCol_X].Value = rowValues[j][pathCt20];
				sViewTK.Cells[num + j, idxCol_AA].Value = 1;
				Row row = sViewTK.Rows[num + j];
				float preferredHeight = row.GetPreferredHeight();
				if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
				{
					row.Height = preferredHeight;
				}
			}
			int num2 = 1;
			for (int j = num; j < num + hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				sViewTK.Cells[j, idxCol_C - 1].Text = num2.ToString();
				num2++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private List<Dictionary<string, string>> GetRowDataValueInXml(string filename, HTKKCMCTableInfo tableInfo, List<string> lstCellPath)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load(filename);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("Sections/Section[@TableName='" + tableInfo.TableName + "']");
			XmlNodeList val3 = val2.SelectNodes("RowInfo/Cell");
			Dictionary<int, string> dictionary = new Dictionary<int, string>();
			for (int i = 0; i < val3.Count; i++)
			{
				string value = ((XmlNode)val3[i].Attributes["Path"]).Value;
				if (lstCellPath.Contains(value))
				{
					dictionary.Add(i, value);
				}
				lstCellPath.Remove(value);
				if (lstCellPath.Count == 0)
				{
					break;
				}
			}
			if (dictionary.Count == 0)
			{
				return list;
			}
			XmlNodeList val4 = val2.SelectNodes("RowData");
			for (int j = 0; j < val4.Count; j++)
			{
				Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
				string text = "###";
				string[] array = ConvertHelpers.ConvertToString(val4[j].InnerXml).Split(new string[1] { text }, StringSplitOptions.None);
				for (int k = 0; k < array.Length; k++)
				{
					if (dictionary.ContainsKey(k))
					{
						dictionary2.Add(dictionary[k], array[k]);
					}
				}
				if (!dictionary2["ct18"].Equals("") && !dictionary2["ct19"].Equals(""))
				{
					list.Add(dictionary2);
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return list;
		}
	}
}
