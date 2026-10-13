using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class GhntNd64Nckt266
{
	private bool iflag;

	private static string cellId_ctLanDau = "G_2";

	private static string cellId_ctThayThe = "K_2";

	private static string cellId_ctLanThayThe = "T_2";

	private static string cellId_ct7a = "D_47";

	private static string cellId_ct7b = "D_49";

	private static string cellId_ct7c = "D_51";

	private static string cellId_ct7d = "D_53";

	private string pathCtDoanhNghiepNho = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CT08/CT_I/ct_a";

	private string pathCtDoanhSieuNghiepNho = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CT08/CT_I/ct_b";

	private string pathCTII_a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CT08/CT_II/ct_1";

	private string maLinhVuNN = "II.1";

	private string maLinhVuCN = "I.a1";

	private string maLinhVuTM = "I.a2";

	private static int[] CellId_ctLanDau => GeneralFunctions.ArrayCellId(cellId_ctLanDau);

	private static int[] CellId_ctThayThe => GeneralFunctions.ArrayCellId(cellId_ctThayThe);

	private static int[] CellId_ctLanThayThe => GeneralFunctions.ArrayCellId(cellId_ctLanThayThe);

	private static int[] CellId_ct7a => GeneralFunctions.ArrayCellId(cellId_ct7a);

	private static int[] CellId_ct7b => GeneralFunctions.ArrayCellId(cellId_ct7b);

	private static int[] CellId_ct7c => GeneralFunctions.ArrayCellId(cellId_ct7c);

	private static int[] CellId_ct7d => GeneralFunctions.ArrayCellId(cellId_ct7d);

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_D => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_G => HTKKFile.ColumnLettersToIndex("G") - 1;

	private int idxCol_L => HTKKFile.ColumnLettersToIndex("L") - 1;

	private int idxCol_P => HTKKFile.ColumnLettersToIndex("P") - 1;

	private int idxCol_W => HTKKFile.ColumnLettersToIndex("W") - 1;

	private int idxCol_AE => HTKKFile.ColumnLettersToIndex("AE") - 1;

	private int idxCol_AL => HTKKFile.ColumnLettersToIndex("AL") - 1;

	private int idxCol_AN => HTKKFile.ColumnLettersToIndex("AN") - 1;

	private int idxCol_AO => HTKKFile.ColumnLettersToIndex("AO") - 1;

	private int idxCol_AP => HTKKFile.ColumnLettersToIndex("AP") - 1;

	private int idxCol_AQ => HTKKFile.ColumnLettersToIndex("AQ") - 1;

	private int idxCol_AR => HTKKFile.ColumnLettersToIndex("AR") - 1;

	private int idxCol_N => HTKKFile.ColumnLettersToIndex("N") - 1;

	private int idxCol_T => HTKKFile.ColumnLettersToIndex("T") - 1;

	private int idxCol_AS => HTKKFile.ColumnLettersToIndex("AS") - 1;

	private int idxCol_AH => HTKKFile.ColumnLettersToIndex("AH") - 1;

	private int idxCol_AT => HTKKFile.ColumnLettersToIndex("AT") - 1;

	private int idxCol_AI => HTKKFile.ColumnLettersToIndex("AI") - 1;

	private int idxCol_AK => HTKKFile.ColumnLettersToIndex("AK") - 1;

	public bool IFlag()
	{
		return iflag;
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			if (PeriodInfo.LoaiToKhai.Equals("C"))
			{
				sView.Cells[CellId_ctLanDau[1], CellId_ctLanDau[0]].Text = "X";
				sView.Cells[CellId_ctThayThe[1], CellId_ctThayThe[0]].Text = "";
				sView.Cells[CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]].Locked = true;
			}
			else
			{
				sView.Cells[CellId_ctLanDau[1], CellId_ctLanDau[0]].Text = "";
				sView.Cells[CellId_ctThayThe[1], CellId_ctThayThe[0]].Text = "X";
				sView.Cells[CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]].Locked = false;
			}
			bool flag = (bool)sView.Cells[CellId_ct7d[1], CellId_ct7d[0]].Value;
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sView.Cells[num + i, idxCol_D].Locked = !flag;
				sView.Cells[num + i, idxCol_G].Locked = !flag;
				sView.Cells[num + i, idxCol_L].Locked = !flag;
				sView.Cells[num + i, idxCol_P].Locked = !flag;
				sView.Cells[num + i, idxCol_AE].Locked = !flag;
				sView.Cells[num + i, idxCol_AL].Locked = !flag;
				sView.Cells[num + i, idxCol_AI].Locked = !flag;
				sView.Cells[num + i, idxCol_AK].Locked = !flag;
				sView.Cells[num + i, idxCol_AN].Locked = !flag;
				sView.Cells[num + i, idxCol_AO].Locked = !flag;
				string text = sView.Cells[num + i, idxCol_AT].Text;
				sView.Cells[num + i, idxCol_L].Locked = false;
				sView.Cells[num + i, idxCol_P].Locked = false;
				if (text.Length >= 5)
				{
					if (text.Substring(3, 2).Equals("00"))
					{
						sView.Cells[num + i, idxCol_L].Locked = true;
						continue;
					}
					sView.Cells[num + i, idxCol_L].Locked = true;
					sView.Cells[num + i, idxCol_P].Locked = true;
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhNghiepNho);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhSieuNghiepNho);
			bool flag2 = (bool)sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
			bool flag3 = (bool)sView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			int num2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				sView.Cells[num2 + i, idxCol_C].Locked = !flag2;
				sView.Cells[num2 + i, idxCol_N].Locked = !flag2;
				sView.Cells[num2 + i, idxCol_T].Locked = !flag2;
				sView.Cells[num2 + i, idxCol_AH].Locked = !flag2;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[2];
			int num3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
			{
				sView.Cells[num3 + i, idxCol_C].Locked = !flag3;
				sView.Cells[num3 + i, idxCol_N].Locked = !flag3;
				sView.Cells[num3 + i, idxCol_T].Locked = !flag3;
				sView.Cells[num3 + i, idxCol_AH].Locked = !flag3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = (bool)sheetView.Cells[CellId_ct7d[1], CellId_ct7d[0]].Value;
			if (idxRow == CellId_ct7d[1] && idxCol == CellId_ct7d[0])
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int count = hTKKCMCTableInfo.RowsValues.Count;
				if (!flag)
				{
					for (int i = 1; i < count; i++)
					{
						hTKKCMCSheetInfo.Tables[0].RowsValues.RemoveAt(1);
						RemoveDataRows(sheetView, num + 1, 1, hTKKCMCSheetInfo);
					}
					sheetView.Cells[num, idxCol_D].Text = "";
					sheetView.Cells[num, idxCol_G].Text = "";
					sheetView.Cells[num, idxCol_L].Value = "";
					sheetView.Cells[num, idxCol_P].Value = "";
					sheetView.Cells[num, idxCol_W].Value = "";
					sheetView.Cells[num, idxCol_AE].Value = "";
					sheetView.Cells[num, idxCol_AL].Value = "";
					sheetView.Cells[num, idxCol_AN].Value = "";
					sheetView.Cells[num, idxCol_AO].Value = "";
					sheetView.Cells[num, idxCol_AP].Value = "";
					sheetView.Cells[num, idxCol_AQ].Value = "";
					sheetView.Cells[num, idxCol_AR].Value = "";
					sheetView.Cells[num, idxCol_AS].Value = "";
					sheetView.Cells[num, idxCol_AT].Value = "";
					sheetView.Cells[num, idxCol_AI].Value = "";
					sheetView.Cells[num, idxCol_AK].Value = "";
				}
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					sheetView.Cells[num + i, idxCol_D].Locked = !flag;
					sheetView.Cells[num + i, idxCol_G].Locked = !flag;
					sheetView.Cells[num + i, idxCol_L].Locked = !flag;
					sheetView.Cells[num + i, idxCol_P].Locked = !flag;
					sheetView.Cells[num + i, idxCol_AE].Locked = !flag;
					sheetView.Cells[num + i, idxCol_AL].Locked = !flag;
					sheetView.Cells[num + i, idxCol_AN].Locked = !flag;
					sheetView.Cells[num + i, idxCol_AO].Locked = !flag;
					sheetView.Cells[num + i, idxCol_AI].Locked = !flag;
					sheetView.Cells[num + i, idxCol_AK].Locked = !flag;
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhNghiepNho);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhSieuNghiepNho);
			bool flag2 = (bool)winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
			bool flag3 = (bool)winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			int count2;
			if (idxRow == hTKKCMCCellInfo.CellTemplate.RowIndex - 1 && idxCol == hTKKCMCCellInfo.CellTemplate.ColIndex - 1)
			{
				if (flag2)
				{
					winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = false;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
				int num2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
				count2 = hTKKCMCTableInfo2.RowsValues.Count;
				if (!flag2)
				{
					for (int i = 1; i < count2; i++)
					{
						hTKKCMCSheetInfo.Tables[1].RowsValues.RemoveAt(1);
						RemoveDataRows(sheetView, num2 + 1, 1, hTKKCMCSheetInfo);
					}
					sheetView.Cells[num2, idxCol_C].Text = "";
					sheetView.Cells[num2, idxCol_N].Value = "0";
					sheetView.Cells[num2, idxCol_T].Value = "0";
					sheetView.Cells[num2, idxCol_AH].Value = "0";
					sheetView.Cells[num2, idxCol_AP].Value = "";
				}
				for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
				{
					sheetView.Cells[num2 + i, idxCol_C].Locked = !flag2;
					sheetView.Cells[num2 + i, idxCol_N].Locked = !flag2;
					sheetView.Cells[num2 + i, idxCol_T].Locked = !flag2;
					sheetView.Cells[num2 + i, idxCol_AH].Locked = !flag2;
				}
			}
			if (idxRow != hTKKCMCCellInfo2.CellTemplate.RowIndex - 1 || idxCol != hTKKCMCCellInfo2.CellTemplate.ColIndex - 1)
			{
				return;
			}
			if (flag3)
			{
				winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = false;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[2];
			int num3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
			count2 = hTKKCMCTableInfo3.RowsValues.Count;
			if (!flag3)
			{
				for (int i = 1; i < count2; i++)
				{
					hTKKCMCSheetInfo.Tables[2].RowsValues.RemoveAt(1);
					RemoveDataRows(sheetView, num3 + 1, 1, hTKKCMCSheetInfo);
				}
				sheetView.Cells[num3, idxCol_C].Text = "";
				sheetView.Cells[num3, idxCol_N].Value = "0";
				sheetView.Cells[num3, idxCol_T].Value = "0";
				sheetView.Cells[num3, idxCol_AH].Value = "0";
				sheetView.Cells[num3, idxCol_AP].Value = "";
			}
			for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
			{
				sheetView.Cells[num3 + i, idxCol_C].Locked = !flag3;
				sheetView.Cells[num3 + i, idxCol_N].Locked = !flag3;
				sheetView.Cells[num3 + i, idxCol_T].Locked = !flag3;
				sheetView.Cells[num3 + i, idxCol_AH].Locked = !flag3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			if (winFpSpread.ActiveSheetIndex != 0)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhNghiepNho);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhSieuNghiepNho);
			bool flag = (bool)winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
			bool flag2 = (bool)winFpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			if ((iRow == hTKKCMCCellInfo.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo.CellTemplate.ColIndex) || (iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex && iCol == hTKKCMCCellInfo2.CellTemplate.ColIndex))
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int count = hTKKCMCTableInfo.RowsValues.Count;
				if (!flag)
				{
					for (int num2 = 1; num2 < count; num2++)
					{
						hTKKCMCSheetInfo.Tables[1].RowsValues.RemoveAt(1);
						RemoveDataRows(sheetView, num + 1, 1, hTKKCMCSheetInfo);
					}
					sheetView.Cells[num, idxCol_C].Text = "";
					sheetView.Cells[num, idxCol_N].Value = "0";
					sheetView.Cells[num, idxCol_T].Value = "0";
					sheetView.Cells[num, idxCol_AH].Value = "0";
					sheetView.Cells[num, idxCol_AP].Value = "";
				}
				for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
				{
					sheetView.Cells[num + num2, idxCol_C].Locked = !flag;
					sheetView.Cells[num + num2, idxCol_N].Locked = !flag;
					sheetView.Cells[num + num2, idxCol_T].Locked = !flag;
					sheetView.Cells[num + num2, idxCol_AH].Locked = !flag;
				}
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[2];
				int num3 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
				int count2 = hTKKCMCTableInfo2.RowsValues.Count;
				if (!flag2)
				{
					for (int num2 = 1; num2 < count2; num2++)
					{
						hTKKCMCSheetInfo.Tables[2].RowsValues.RemoveAt(1);
						RemoveDataRows(sheetView, num3 + 1, 1, hTKKCMCSheetInfo);
					}
					sheetView.Cells[num3, idxCol_C].Text = "";
					sheetView.Cells[num3, idxCol_N].Value = "0";
					sheetView.Cells[num3, idxCol_T].Value = "0";
					sheetView.Cells[num3, idxCol_AH].Value = "0";
					sheetView.Cells[num3, idxCol_AP].Value = "";
				}
				for (int num2 = 0; num2 < hTKKCMCTableInfo2.RowsValues.Count; num2++)
				{
					sheetView.Cells[num3 + num2, idxCol_C].Locked = !flag2;
					sheetView.Cells[num3 + num2, idxCol_N].Locked = !flag2;
					sheetView.Cells[num3 + num2, idxCol_T].Locked = !flag2;
					sheetView.Cells[num3 + num2, idxCol_AH].Locked = !flag2;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[0];
			int num4 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
			int num5 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1 + hTKKCMCTableInfo3.RowsValues.Count;
			if (iRow >= num4 && iRow <= num5)
			{
				if (iCol == idxCol_D + 1)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_L].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_AP].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_P].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_AQ].Text = "";
				}
				if (iCol == idxCol_G + 1)
				{
					string text = sheetView.Cells[(int)iRow - 1, idxCol_AT].Text;
					string text2 = "";
					string text3 = "";
					sheetView.Cells[(int)iRow - 1, idxCol_L].Locked = false;
					sheetView.Cells[(int)iRow - 1, idxCol_P].Locked = false;
					sheetView.Cells[(int)iRow - 1, idxCol_L].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_AP].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_P].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_AQ].Text = "";
					if (text.Length >= 5)
					{
						if (text.Substring(3, 2).Equals("00"))
						{
							text2 = text.Substring(0, 3);
							Dictionary<string, string> dicTinhTP = HTKKWinDictionaryDiaDanh.dicTinhTP;
							foreach (KeyValuePair<string, string> item in dicTinhTP)
							{
								if (item.Value.Equals(text2))
								{
									sheetView.Cells[(int)iRow - 1, idxCol_L].Text = item.Key;
									sheetView.Cells[(int)iRow - 1, idxCol_L].Locked = true;
									sheetView.Cells[(int)iRow - 1, idxCol_AP].Text = item.Value;
									break;
								}
							}
						}
						else
						{
							text2 = text.Substring(0, 3);
							text3 = text;
							Dictionary<string, string> dicTinhTP = HTKKWinDictionaryDiaDanh.dicTinhTP;
							foreach (KeyValuePair<string, string> item2 in dicTinhTP)
							{
								if (item2.Value.Equals(text2))
								{
									sheetView.Cells[(int)iRow - 1, idxCol_L].Text = item2.Key;
									sheetView.Cells[(int)iRow - 1, idxCol_L].Locked = true;
									sheetView.Cells[(int)iRow - 1, idxCol_AP].Text = item2.Value;
									break;
								}
							}
							Dictionary<string, string[]> dicQuanHuyenByTinhTP = HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP;
							ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
							comboBoxCellType.Items = HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[text2];
							sheetView.Cells[(int)iRow - 1, idxCol_P].CellType = comboBoxCellType;
							sheetView.Cells[(int)iRow - 1, idxCol_P].Text = "";
							Dictionary<string, string> dicQuanHuyen = HTKKWinDictionaryDiaDanh.dicQuanHuyen;
							foreach (KeyValuePair<string, string> item3 in dicQuanHuyen)
							{
								if (item3.Key.Equals(text3))
								{
									sheetView.Cells[(int)iRow - 1, idxCol_P].Text = item3.Value;
									sheetView.Cells[(int)iRow - 1, idxCol_P].Locked = true;
									sheetView.Cells[(int)iRow - 1, idxCol_AQ].Text = item3.Key;
									break;
								}
							}
						}
					}
				}
				if (iCol == idxCol_L + 1)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_P].Text = "";
					sheetView.Cells[(int)iRow - 1, idxCol_AQ].Text = "";
				}
				if (iCol == idxCol_AL + 1)
				{
					string text4 = sheetView.Cells[(int)iRow - 1, idxCol_AL].Text;
					if (!HTKKFormater.IsNumber(text4))
					{
						sheetView.Cells[(int)iRow - 1, idxCol_AL].Text = "";
					}
				}
			}
			if (iRow == CellId_ctLanThayThe[1] + 1 && iCol == CellId_ctLanThayThe[0] + 1 && PeriodInfo.LoaiToKhai.Equals("T"))
			{
				string text5 = sheetView.Cells[CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]].Text;
				if (ConvertHelpers.ConvertToInt(text5) == 0)
				{
					sheetView.Cells[CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]].Text = "";
				}
				else
				{
					sheetView.Cells[CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]].Text = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToInt(text5));
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		bool result = true;
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		if (winFpSpread.TabStrip[0].BackColor == Color.Pink)
		{
			result = false;
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
		return result;
	}

	private void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25"));
			string message12 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message13 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message14 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_28"));
			string message15 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
			string message16 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_29"));
			string message17 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_30"));
			string message18 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31"));
			string message19 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32"));
			string message20 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33"));
			string message21 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35"));
			if (PeriodInfo.LoaiToKhai.Equals("T"))
			{
				string text = sheetView.Cells[CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]].Text;
				HTKKFile.RemoveValidate(sheetView, CellId_ctLanThayThe[1], CellId_ctLanThayThe[0]);
				if (text.Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ctLanThayThe[1], CellId_ctLanThayThe[0], message16);
					num = 2;
				}
			}
			bool flag = (bool)sheetView.Cells[CellId_ct7a[1], CellId_ct7a[0]].Value;
			bool flag2 = (bool)sheetView.Cells[CellId_ct7b[1], CellId_ct7b[0]].Value;
			bool flag3 = (bool)sheetView.Cells[CellId_ct7c[1], CellId_ct7c[0]].Value;
			bool flag4 = (bool)sheetView.Cells[CellId_ct7d[1], CellId_ct7d[0]].Value;
			HTKKFile.RemoveValidate(sheetView, CellId_ct7a[1], CellId_ct7a[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct7b[1], CellId_ct7b[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct7c[1], CellId_ct7c[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_ct7d[1], CellId_ct7d[0]);
			if (flag || flag2 || flag3 || flag4)
			{
				if ((!flag || !flag2) && (flag || flag2))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct7a[1], CellId_ct7a[0], message3);
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct7b[1], CellId_ct7b[0], message3);
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct7a[1], CellId_ct7a[0], message);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct7b[1], CellId_ct7b[0], message);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct7c[1], CellId_ct7c[0], message);
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct7d[1], CellId_ct7d[0], message);
				num = 2;
			}
			if (flag4)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string text2 = sheetView.Cells[num2 + i, idxCol_D].Text;
					string text3 = sheetView.Cells[num2 + i, idxCol_G].Text;
					string text4 = sheetView.Cells[num2 + i, idxCol_L].Text;
					string text5 = sheetView.Cells[num2 + i, idxCol_P].Text;
					string text6 = sheetView.Cells[num2 + i, idxCol_AE].Text;
					string text7 = sheetView.Cells[num2 + i, idxCol_AL].Text;
					string text8 = sheetView.Cells[num2 + i, idxCol_AN].Text;
					string text9 = sheetView.Cells[num2 + i, idxCol_AO].Text;
					string text10 = sheetView.Cells[num2 + i, idxCol_AI].Text;
					string text11 = sheetView.Cells[num2 + i, idxCol_AK].Text;
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_D);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_G);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_L);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_P);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AE);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AL);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AN);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AO);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AI);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AK);
					if (text2.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_D, message16);
						num = 2;
					}
					if (text3.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_G, message16);
						num = 2;
					}
					if (text4.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_L, message16);
						num = 2;
					}
					if (text5.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_P, message16);
						num = 2;
					}
					if (text6.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AE, message16);
						num = 2;
					}
					if (text7.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AL, message16);
						num = 2;
					}
					else if (!HTKKFormater.IsNumber(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AL, message17);
						num = 2;
					}
					else if (ConvertHelpers.ConvertToDouble(text7) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AL, message18);
						num = 2;
					}
					if (text8.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AN, message16);
						num = 2;
					}
					if (text10.Equals(""))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num2 + i, idxCol_AI, message21);
						num = 1;
					}
					if (text11.Equals(""))
					{
						HTKKFile.SetNodeBackGroundYellow(sheetView, num2 + i, idxCol_AK, message21);
						num = 1;
					}
					DateTime dtDate;
					if (text9.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AO, message16);
						num = 2;
					}
					else if (HTKKFormater.IsDateTime(text9, "dd/MM/yyyy", out dtDate) && dtDate >= DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", CultureInfo.InvariantCulture))
					{
						DateTime dateTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
						if (dtDate > dateTime)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AO, message20);
							num = 2;
						}
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AO, message19);
						num = 2;
					}
				}
			}
			else
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_D);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_G);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_L);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_P);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_AE);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_AL);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_AN);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_AO);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_AI);
				HTKKFile.RemoveValidate(sheetView, num2, idxCol_AK);
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhNghiepNho);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCtDoanhSieuNghiepNho);
			bool flag5 = (bool)sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
			bool flag6 = (bool)sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			if (flag5 && flag6)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message2);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message2);
				num = 2;
			}
			for (int num3 = 1; num3 < hTKKCMCSheetInfo.Tables.Count; num3++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[num3];
				int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string text12 = sheetView.Cells[num2 + i, idxCol_C].Text;
					string text13 = sheetView.Cells[num2 + i, idxCol_AP].Text;
					double num4 = Convert.ToDouble(sheetView.Cells[num2 + i, idxCol_N].Value);
					double num5 = Convert.ToDouble(sheetView.Cells[num2 + i, idxCol_T].Value);
					double num6 = Convert.ToDouble(sheetView.Cells[num2 + i, idxCol_AH].Value);
					bool flag7 = false;
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_C);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_N);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_T);
					HTKKFile.RemoveValidate(sheetView, num2 + i, idxCol_AH);
					if (hTKKCMCTableInfo.TableName.Equals("Doanh_nghiep_nho"))
					{
						if (flag5)
						{
							if (text12.Equals(""))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_C, message16);
								num = 2;
							}
							if (num4 == 0.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_N, message16);
								num = 2;
							}
							if (num6 == 0.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AH, message16);
								num = 2;
							}
						}
						if (text13.Equals(maLinhVuNN) || text13.Equals(maLinhVuCN))
						{
							if (num4 > 100.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_N, message4);
								num = 2;
							}
							if (num5 <= 50000000000.0)
							{
								flag7 = true;
							}
							if (num6 <= 20000000000.0)
							{
								flag7 = true;
							}
							if (!flag7)
							{
								if (num5 > 50000000000.0)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message6);
									num = 2;
								}
								if (num6 > 20000000000.0)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AH, message8);
									num = 2;
								}
							}
							continue;
						}
						if (num4 > 50.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_N, message5);
							num = 2;
						}
						if (num5 <= 100000000000.0)
						{
							flag7 = true;
						}
						if (num6 <= 50000000000.0)
						{
							flag7 = true;
						}
						if (!flag7)
						{
							if (num5 > 100000000000.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message7);
								num = 2;
							}
							if (num6 > 50000000000.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AH, message9);
								num = 2;
							}
						}
					}
					else
					{
						if (!hTKKCMCTableInfo.TableName.Equals("Doanh_nghiep_sieu_nho"))
						{
							continue;
						}
						if (flag6)
						{
							if (text12.Equals(""))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_C, message16);
								num = 2;
							}
							if (num4 == 0.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_N, message16);
								num = 2;
							}
							if (num6 == 0.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AH, message16);
								num = 2;
							}
						}
						if (num4 > 10.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_N, message10);
							num = 2;
						}
						if (text13.Equals(maLinhVuNN) || text13.Equals(maLinhVuCN))
						{
							if (num5 <= 3000000000.0)
							{
								flag7 = true;
							}
							if (num6 <= 3000000000.0)
							{
								flag7 = true;
							}
							if (!flag7)
							{
								if (num5 > 3000000000.0)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message11);
									num = 2;
								}
								if (num6 > 3000000000.0)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AH, message13);
									num = 2;
								}
							}
							continue;
						}
						if (num5 <= 10000000000.0)
						{
							flag7 = true;
						}
						if (num6 <= 3000000000.0)
						{
							flag7 = true;
						}
						if (!flag7)
						{
							if (num5 > 10000000000.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_T, message12);
								num = 2;
							}
							if (num6 > 3000000000.0)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxCol_AH, message13);
								num = 2;
							}
						}
					}
				}
			}
			bool flag8 = false;
			bool flag9 = false;
			if (flag5 || flag6)
			{
				flag9 = true;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCTII_a);
			int num7 = hTKKCMCCellInfo3.CellTemplate.RowIndex - 82;
			string text14 = "82~84~86~88~90~92~94~96~98~100~102~104~106~108~110~112~114~116~118~120~122~124~126~128~130~132~134~136~140~142~144~146~148~150~152~154~156~158~160~162~164~166~168";
			string[] array = text14.Split(new char[1] { '~' });
			for (int i = 0; i < array.Length; i++)
			{
				if ((bool)sheetView.Cells[ConvertHelpers.ConvertToInt(array[i]) - 1 + num7, idxCol_D].Value)
				{
					flag8 = true;
					break;
				}
			}
			if (flag8 && flag9)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message15);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message15);
				for (int i = 0; i < array.Length; i++)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, ConvertHelpers.ConvertToInt(array[i]) - 1 + num7, idxCol_D, message15);
				}
				num = 2;
			}
			else if (flag8 || flag9)
			{
				for (int i = 0; i < array.Length; i++)
				{
					HTKKFile.RemoveValidate(sheetView, ConvertHelpers.ConvertToInt(array[i]) - 1 + num7, idxCol_D);
				}
			}
			else
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message14);
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message14);
				for (int i = 0; i < array.Length; i++)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, ConvertHelpers.ConvertToInt(array[i]) - 1 + num7, idxCol_D, message14);
				}
				num = 2;
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
			winFpSpread.TabStrip[0].BackColor = Color.Pink;
		}
	}

	public static void RemoveDataRows(SheetView sView, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (rowCount <= 0)
			{
				return;
			}
			sView.Rows.Remove(frRow, rowCount);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table.Location.DataLocationTemplateRowStart > frRow + 1)
				{
					table.Increment(rowCount * -1);
				}
			}
			sInfo.IncrementStaticCell(frRow, rowCount * -1);
		}
		catch (Exception)
		{
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
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
				if (idxRow >= num && idxRow <= num2 && idxCol == idxCol_AL)
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
				if (idxRow == CellId_ctLanThayThe[1] && idxCol == CellId_ctLanThayThe[0])
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
}
