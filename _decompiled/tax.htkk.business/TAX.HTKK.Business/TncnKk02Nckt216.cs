using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnKk02Nckt216
{
	private bool iflag;

	private static string cellId_ct21 = "V_41";

	private static string cellId_ct22 = "R_43";

	private static string cellId_ct23 = "R_44";

	private static string cellId_ct25 = "R_46";

	private static string cellId_ct26 = "R_47";

	private static string cellId_ct27 = "R_48";

	private static string cellId_ct28 = "R_49";

	private static string cellId_ct29 = "R_50";

	private string pathCt30 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanCuTru/ct30";

	private static string cellId_ct30 = "R_51";

	private string pathCt31 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanCuTru/ct31";

	private static string cellId_ct31 = "R_52";

	private static string cellId_ct32 = "R_53";

	private static string cellId_ct35 = "R_57";

	private static string cellId_ct36 = "R_58";

	private static string cellId_ct37 = "R_59";

	private string formulaCt37 = "ROUND(R57*R58/100,0)";

	private string pathCt37 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/CaNhanKhongCuTru/ct37";

	private static int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private static int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private static int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private static int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private static int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private static int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private static int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private static int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private static int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private static int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private List<int> lstIdxRowsRelate_Ct31
	{
		get
		{
			List<int> list = new List<int>();
			list.Add(CellId_ct22[1]);
			list.Add(CellId_ct23[1]);
			list.Add(CellId_ct25[1]);
			list.Add(CellId_ct26[1]);
			list.Add(CellId_ct27[1]);
			list.Add(CellId_ct28[1]);
			list.Add(CellId_ct29[1]);
			return list;
		}
	}

	private static int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private static int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	private static int[] CellId_ct36 => GeneralFunctions.ArrayCellId(cellId_ct36);

	private static int[] CellId_ct37 => GeneralFunctions.ArrayCellId(cellId_ct37);

	private List<int> lstIdxRowsRelate_Ct37
	{
		get
		{
			List<int> list = new List<int>();
			list.Add(CellId_ct35[1]);
			list.Add(CellId_ct36[1]);
			return list;
		}
	}

	public bool IFlag()
	{
		return iflag;
	}

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = ((winFpSpread != null && winFpSpread.Sheets.Count > 0) ? winFpSpread.Sheets[0] : null);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int? num = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
			int? num2 = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null));
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile && sheetView != null)
			{
				sheetView.Cells[CellId_ct25[1], CellId_ct25[0]].Value = 0;
				sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Value = 20;
			}
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
			int? num = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
			int? num2 = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null));
			string quy = PeriodInfo.Quy;
			string nam = PeriodInfo.Nam;
			if (!ToKhaiInfo.SheetInfos[0].IsReadDataFile)
			{
				if ((nam.Equals("2020") && ConvertHelpers.ConvertToInt(quy) >= 3) || ConvertHelpers.ConvertToInt(nam) > 2020)
				{
					string text = sView.Cells[CellId_ct25[1], CellId_ct25[0]].Text;
					if (num.HasValue && num2.HasValue && (string.IsNullOrEmpty(text) || text.Equals("0")))
					{
						sView.Cells[CellId_ct25[1], CellId_ct25[0]].Value = (num2 - num + 1) * 11000000;
					}
				}
				else
				{
					string text = sView.Cells[CellId_ct25[1], CellId_ct25[0]].Text;
					if (num.HasValue && num2.HasValue && (string.IsNullOrEmpty(text) || text.Equals("0")))
					{
						sView.Cells[CellId_ct25[1], CellId_ct25[0]].Value = (num2 - num + 1) * 9000000;
					}
				}
				sView.Cells[CellId_ct36[1], CellId_ct36[0]].Value = 20;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCt37);
			if (hTKKCMCCellInfo != null)
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct37[1], CellId_ct37[0], locked: false, "", Convert.ToDouble(hTKKCMCCellInfo.Value));
			}
			string nam2 = PeriodInfo.Nam;
			string quy2 = PeriodInfo.Quy;
			int column = GeneralFunctions.ArrayCellId("R_53")[0];
			int row = GeneralFunctions.ArrayCellId("R_53")[1];
			int column2 = GeneralFunctions.ArrayCellId("R_54")[0];
			int row2 = GeneralFunctions.ArrayCellId("R_54")[1];
			if (ConvertHelpers.ConvertToInt(nam2) > 2018 || (ConvertHelpers.ConvertToInt(nam2) == 2018 && ConvertHelpers.ConvertToInt(quy2) == 4))
			{
				sView.Cells[row, column].Value = "0";
				sView.Cells[row, column].Locked = true;
				sView.Cells[row2, column2].Value = "0";
				sView.Cells[row2, column2].Locked = true;
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
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				int num = HTKKFile.ColumnLettersToIndex("R");
				if (num == iCol && lstIdxRowsRelate_Ct37.Contains(Convert.ToInt32(iRow) - 1))
				{
					winFpSpread.Sheets[0].Cells[CellId_ct37[1], CellId_ct37[0]].Formula = formulaCt37;
				}
				if (num == iCol && lstIdxRowsRelate_Ct31.Contains(Convert.ToInt32(iRow) - 1))
				{
					winFpSpread.Sheets[0].RecalculateAll();
					int? num2 = (GetFirstNumber(PeriodInfo.TuThang).HasValue ? GetFirstNumber(PeriodInfo.TuThang) : ((int?)null));
					int? sothang = (GetFirstNumber(PeriodInfo.DenThang).HasValue ? GetFirstNumber(PeriodInfo.DenThang) : ((int?)null)) - num2 + 1;
					string text = winFpSpread.Sheets[0].Cells[CellId_ct30[1], CellId_ct30[0]].Text;
					double valueCT = Convert.ToDouble(text.Replace(".", ""));
					winFpSpread.Sheets[0].Cells[CellId_ct31[1], CellId_ct31[0]].Formula = GetFormula31(valueCT, sothang);
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
		//IL_027f: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = 0;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_19"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20"));
			double num2 = Convert.ToDouble(sheetView.Cells[CellId_ct22[1], CellId_ct22[0]].Value);
			double num3 = Convert.ToDouble(sheetView.Cells[CellId_ct23[1], CellId_ct23[0]].Value);
			double num4 = Convert.ToDouble(sheetView.Cells[CellId_ct25[1], CellId_ct25[0]].Value);
			double num5 = Convert.ToDouble(sheetView.Cells[CellId_ct26[1], CellId_ct26[0]].Value);
			double num6 = Convert.ToDouble(sheetView.Cells[CellId_ct27[1], CellId_ct27[0]].Value);
			double num7 = Convert.ToDouble(sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Value);
			double num8 = Convert.ToDouble(sheetView.Cells[CellId_ct29[1], CellId_ct29[0]].Value);
			double num9 = Convert.ToDouble(sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Value);
			double num10 = Convert.ToDouble(sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Value);
			string nam = PeriodInfo.Nam;
			string quy = PeriodInfo.Quy;
			if ((num2 != 0.0 || num3 != 0.0 || num5 != 0.0 || num6 != 0.0 || num7 != 0.0 || num8 != 0.0 || num9 != 0.0) && num10 != 0.0 && num2 >= num3 && ((Control)winFpSpread).AccessibleDefaultActionDescription != "NotCheck")
			{
				winFpSpread.ResumeLayout();
				iflag = true;
				frmMessage frmMessage = new frmMessage("0225", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				num = 1;
			}
			else
			{
				iflag = false;
			}
			if ((nam.Equals("2020") && ConvertHelpers.ConvertToInt(quy) >= 3) || ConvertHelpers.ConvertToInt(nam) > 2020)
			{
				if (num4 > 33000000.0 || num4 % 11000000.0 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 45, HTKKFile.ColumnLettersToIndex("R") - 1, message3);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 45, HTKKFile.ColumnLettersToIndex("R") - 1);
				}
				if (num5 % 4400000.0 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 46, HTKKFile.ColumnLettersToIndex("R") - 1, message5);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 46, HTKKFile.ColumnLettersToIndex("R") - 1);
				}
			}
			else
			{
				if (num4 > 27000000.0 || num4 % 9000000.0 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 45, HTKKFile.ColumnLettersToIndex("R") - 1, message2);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 45, HTKKFile.ColumnLettersToIndex("R") - 1);
				}
				if (num5 % 3600000.0 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 46, HTKKFile.ColumnLettersToIndex("R") - 1, message4);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 46, HTKKFile.ColumnLettersToIndex("R") - 1);
				}
			}
			if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(quy) == 4))
			{
				if (num9 > 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, 52, HTKKFile.ColumnLettersToIndex("R") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26")));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, 52, HTKKFile.ColumnLettersToIndex("R") - 1);
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
			winFpSpread.TabStrip[0].BackColor = Color.Pink;
		}
	}

	private int? GetFirstNumber(string strInput)
	{
		try
		{
			string s = strInput.Substring(0, strInput.IndexOf('/'));
			return int.Parse(s);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private string GetFormula31(double valueCT30, int? sothang)
	{
		string text = "0";
		double num = 1000000.0;
		double num2 = valueCT30 / (double)sothang.Value;
		text = ((num2 <= 5.0 * num) ? "R51*5%" : ((num2 > 5.0 * num && num2 <= 10.0 * num) ? "5000000*5% + (R51-5000000)*10%" : ((num2 > 10.0 * num && num2 <= 18.0 * num) ? "5000000*5% + 5000000*10% + (R51-10000000)*15%" : ((num2 > 18.0 * num && num2 <= 32.0 * num) ? "5000000*5% + 5000000*10% + 8000000*15% + (R51-18000000)*20%" : ((num2 > 32.0 * num && num2 <= 52.0 * num) ? "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + (R51-32000000)*25%" : ((!(num2 > 52.0 * num) || !(num2 <= 80.0 * num)) ? "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + 20000000*25% + 28000000*30% + (R51-80000000)*35%" : "5000000*5% + 5000000*10% + 8000000*15% + 14000000*20% + 20000000*25% + (R51-52000000)*30%"))))));
		return "ROUND((" + text.Replace("R51", num2.ToString()) + ")*" + sothang + ",0)";
	}
}
