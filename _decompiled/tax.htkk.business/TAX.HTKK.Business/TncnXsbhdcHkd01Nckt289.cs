using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class TncnXsbhdcHkd01Nckt289
{
	private string fomulaColW = "R###+AD###";

	private string MA_COLOR_LOCKED = "#e8f5ff";

	private string MA_COLOR_UNLOCK = "#FFFFFF";

	private static string cellId_ct08 = "W_51";

	private Dictionary<string, string> dicFomula = new Dictionary<string, string>();

	private static string cellId_ct09 = "W_52";

	private static string cellId_ct10 = "W_53";

	private static string cellId_ct11 = "W_54";

	private static string cellId_ct12 = "W_56";

	private static string cellId_ct13 = "W_57";

	private static string cellId_ct14 = "W_58";

	private static string cellId_ct15 = "W_59";

	private static string cellId_ct16 = "W_61";

	private static string cellId_ct17 = "W_62";

	private static string cellId_ct20 = "W_66";

	private static string cellId_ct21 = "W_67";

	private static string cellId_ct22 = "W_68";

	private static string fmlCt08 = "AC51";

	private static string fmlCt09 = "AC52";

	private static string fmlCt10 = "AC53";

	private static string fmlCt11 = "AC54";

	private static string fmlCt12 = "AC56";

	private static string fmlCt13 = "AC57";

	private static string fmlCt14 = "AC58";

	private static string fmlCt15 = "AC59";

	private static string fmlCt16 = "AC61";

	private static string fmlCt17 = "AC62";

	private static string fmlCt18 = "AC63";

	private static string fmlCt19 = "AC65";

	private static string fmlCt20 = "AC66";

	private static string fmlCt21 = "AC67";

	private static string fmlCt22 = "AC68";

	private static int[] CellId_ct08 => GeneralFunctions.ArrayCellId(cellId_ct08);

	private static int[] CellId_ct09 => GeneralFunctions.ArrayCellId(cellId_ct09);

	private static int[] CellId_ct10 => GeneralFunctions.ArrayCellId(cellId_ct10);

	private static int[] CellId_ct11 => GeneralFunctions.ArrayCellId(cellId_ct11);

	private static int[] CellId_ct12 => GeneralFunctions.ArrayCellId(cellId_ct12);

	private static int[] CellId_ct13 => GeneralFunctions.ArrayCellId(cellId_ct13);

	private static int[] CellId_ct14 => GeneralFunctions.ArrayCellId(cellId_ct14);

	private static int[] CellId_ct15 => GeneralFunctions.ArrayCellId(cellId_ct15);

	private static int[] CellId_ct16 => GeneralFunctions.ArrayCellId(cellId_ct16);

	private static int[] CellId_ct17 => GeneralFunctions.ArrayCellId(cellId_ct17);

	private static int[] CellId_ct18 => GeneralFunctions.ArrayCellId("W_63");

	private static int[] CellId_ct19 => GeneralFunctions.ArrayCellId("W_65");

	private static int[] CellId_ct20 => GeneralFunctions.ArrayCellId(cellId_ct20);

	private static int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private static int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 11 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_01_1_XSBHDC_HKD.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public static bool LoadPhuLucSpread(FpSpread axfpsPhuLuc, List<CheckBox> listCheckBox, List<MaskedTextBox> listMaskedTextBox, List<ComboBox> listCombobox, List<RadioButton> listRadioButton)
	{
		bool result = false;
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static bool CheckThang12Quy4(FpSpread axfpsPhuLuc, List<CheckBox> listCheckBox, List<MaskedTextBox> listMaskedTextBox, List<ComboBox> listCombobox, List<RadioButton> listRadioButton)
	{
		bool result = false;
		try
		{
			CheckBox checkBoxByName = GeneralFunctions.getCheckBoxByName(listCheckBox, "chkChonPhuLucKeKhaiAll");
			MaskedTextBox maskedTextBoxByName = GeneralFunctions.getMaskedTextBoxByName(listMaskedTextBox, "mskMonth");
			CheckBox checkBoxByName2 = GeneralFunctions.getCheckBoxByName(listCheckBox, "chkTkhaiThang");
			CheckBox checkBoxByName3 = GeneralFunctions.getCheckBoxByName(listCheckBox, "chkTKQuy");
			ComboBox comboBoxByName = GeneralFunctions.getComboBoxByName(listCombobox, "cbbQuy");
			if (checkBoxByName != null && maskedTextBoxByName != null && checkBoxByName2 != null && checkBoxByName3 != null && comboBoxByName != null)
			{
				CheckThang12Quy4(axfpsPhuLuc, checkBoxByName, maskedTextBoxByName, checkBoxByName2, checkBoxByName3, comboBoxByName);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public static void CheckThang12Quy4(FpSpread axfpsPhuLuc, CheckBox chkChonPhuLucKeKhaiAll, MaskedTextBox mskMonth, CheckBox chkTkhaiThang, CheckBox chkTKQuy, ComboBox cbbQuy)
	{
		try
		{
			if (chkTkhaiThang.Checked && !HTKKFormater.isNullString(((Control)mskMonth).Text) && ((Control)mskMonth).Text.Trim() == "12")
			{
				GeneralFunctions.axfpsPhuLucSetAllValues(axfpsPhuLuc, chkChonPhuLucKeKhaiAll, value: true);
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
			}
			else if (chkTKQuy.Checked && cbbQuy.SelectedItem != null && cbbQuy.SelectedItem.ToString() == "4")
			{
				GeneralFunctions.axfpsPhuLucSetAllValues(axfpsPhuLuc, chkChonPhuLucKeKhaiAll, value: true);
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			LockChiTieuToKhai(sView, ToKhaiInfo);
			int[] array = GeneralFunctions.ArrayCellId("R_52");
			int[] array2 = GeneralFunctions.ArrayCellId("R_57");
			int[] array3 = GeneralFunctions.ArrayCellId("R_66");
			if (PeriodInfo.KieuKy == "M")
			{
				if ((PeriodInfo.Nam.Equals("2026") && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) >= 3) || ConvertHelpers.ConvertToInt(PeriodInfo.Nam) > 2026)
				{
					sView.Cells[array[1], array[0]].Formula = "R51";
					sView.Cells[array[1], array[0]].Locked = true;
					sView.Cells[array[1], array[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
					sView.Cells[array2[1], array2[0]].Formula = "R56";
					sView.Cells[array2[1], array2[0]].Locked = true;
					sView.Cells[array2[1], array2[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
					sView.Cells[array3[1], array3[0]].Formula = "R65";
					sView.Cells[array3[1], array3[0]].Locked = true;
					sView.Cells[array3[1], array3[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				}
			}
			else if (PeriodInfo.KieuKy == "Q" && ((PeriodInfo.Nam.Equals("2026") && ConvertHelpers.ConvertToInt(PeriodInfo.Quy) >= 2) || ConvertHelpers.ConvertToInt(PeriodInfo.Quy) > 2026))
			{
				sView.Cells[array[1], array[0]].Formula = "R51";
				sView.Cells[array[1], array[0]].Locked = true;
				sView.Cells[array[1], array[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				sView.Cells[array2[1], array2[0]].Formula = "R56";
				sView.Cells[array2[1], array2[0]].Locked = true;
				sView.Cells[array2[1], array2[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				sView.Cells[array3[1], array3[0]].Formula = "R65";
				sView.Cells[array3[1], array3[0]].Locked = true;
				sView.Cells[array3[1], array3[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
	}

	private void SetFormula(SheetView sView)
	{
		try
		{
			sView.Cells[CellId_ct08[1], CellId_ct08[0]].Formula = fmlCt08;
			sView.Cells[CellId_ct09[1], CellId_ct09[0]].Formula = fmlCt09;
			sView.Cells[CellId_ct10[1], CellId_ct10[0]].Formula = fmlCt10;
			sView.Cells[CellId_ct11[1], CellId_ct11[0]].Formula = fmlCt11;
			sView.Cells[CellId_ct12[1], CellId_ct12[0]].Formula = fmlCt12;
			sView.Cells[CellId_ct13[1], CellId_ct13[0]].Formula = fmlCt13;
			sView.Cells[CellId_ct14[1], CellId_ct14[0]].Formula = fmlCt14;
			sView.Cells[CellId_ct15[1], CellId_ct15[0]].Formula = fmlCt15;
			sView.Cells[CellId_ct16[1], CellId_ct16[0]].Formula = fmlCt16;
			sView.Cells[CellId_ct17[1], CellId_ct17[0]].Formula = fmlCt17;
			sView.Cells[CellId_ct18[1], CellId_ct18[0]].Formula = fmlCt18;
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].Formula = fmlCt19;
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].Formula = fmlCt20;
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].Formula = fmlCt21;
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].Formula = fmlCt22;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void SetFormulaPL(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo)
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

	private void ResetFormula(SheetView sViewTK)
	{
		try
		{
			sViewTK.Cells[CellId_ct08[1], CellId_ct08[0]].ResetFormula();
			sViewTK.Cells[CellId_ct09[1], CellId_ct09[0]].ResetFormula();
			sViewTK.Cells[CellId_ct10[1], CellId_ct10[0]].ResetFormula();
			sViewTK.Cells[CellId_ct11[1], CellId_ct11[0]].ResetFormula();
			sViewTK.Cells[CellId_ct12[1], CellId_ct12[0]].ResetFormula();
			sViewTK.Cells[CellId_ct13[1], CellId_ct13[0]].ResetFormula();
			sViewTK.Cells[CellId_ct14[1], CellId_ct14[0]].ResetFormula();
			sViewTK.Cells[CellId_ct15[1], CellId_ct15[0]].ResetFormula();
			sViewTK.Cells[CellId_ct16[1], CellId_ct16[0]].ResetFormula();
			sViewTK.Cells[CellId_ct17[1], CellId_ct17[0]].ResetFormula();
			sViewTK.Cells[CellId_ct18[1], CellId_ct18[0]].ResetFormula();
			sViewTK.Cells[CellId_ct19[1], CellId_ct19[0]].ResetFormula();
			sViewTK.Cells[CellId_ct20[1], CellId_ct20[0]].ResetFormula();
			sViewTK.Cells[CellId_ct21[1], CellId_ct21[0]].ResetFormula();
			sViewTK.Cells[CellId_ct22[1], CellId_ct22[0]].ResetFormula();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		LockChiTieuToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			int index = 1;
			SheetView sheetView = winFpSpread.Sheets[index];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("G");
			if (iCol == num && sheetView.Cells[(int)iRow - 1, num - 1].Value.ToIntOrDefault() < 500000000)
			{
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
			if (PeriodInfo.KieuKy == "M")
			{
				if (((PeriodInfo.Nam.Equals("2026") && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) >= 3) || ConvertHelpers.ConvertToInt(PeriodInfo.Nam) > 2026) && iCol == num)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
				}
			}
			else if (PeriodInfo.KieuKy == "Q" && ((PeriodInfo.Nam.Equals("2026") && ConvertHelpers.ConvertToInt(PeriodInfo.Quy) >= 2) || ConvertHelpers.ConvertToInt(PeriodInfo.Quy) > 2026) && iCol == num)
			{
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = sheetView.Cells[(int)iRow - 1, num - 1].Value;
			}
			SetFormula(winFpSpread.Sheets[0]);
			LockChiTieuToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
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
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			if (!string.IsNullOrEmpty(HTKKFile.ErrorMessage) && (ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || ToKhaiInfo.SheetInfos[1].Tables[1].RowsValues.Count > 500 || ToKhaiInfo.SheetInfos[1].Tables[2].RowsValues.Count > 500))
			{
				winFpSpread.ResumeLayout();
				Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
				foreach (Process process in processesByName)
				{
					process.Kill();
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			switch (value)
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			if (ToKhaiInfo.SheetInfos[1] == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int num = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("E") - 1;
			int num3 = HTKKFile.ColumnLettersToIndex("F") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("G") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("H") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("I") - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			string message = HTKKFile.GetMessage("0047");
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54"));
			string message12 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55"));
			string message13 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56"));
			string message14 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_79"));
			string message15 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_80"));
			string message16 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_81"));
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < tables[i].RowsValues.Count; j++)
				{
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num4);
					HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num5);
				}
			}
			int num7 = -1;
			int[] array = GeneralFunctions.ArrayCellId("N_4");
			string text = sheetView.Cells[array[1], array[0]].Text.Trim();
			for (int i = 0; i < tables.Count; i++)
			{
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				int num8 = tables[i].Location.DataLocationTemplateRowStart - 1;
				for (int k = 0; k < tables[i].RowsValues.Count; k++)
				{
					List<string> rowValue = tables[i].RowsValues[k].RowValue;
					string text2 = sheetView.Cells[num8 + k, array[0]].Text.Trim();
					if (text2 == "1")
					{
						num7 = i;
						break;
					}
				}
				if (num7 >= 0)
				{
					break;
				}
			}
			if (num7 == -1)
			{
				for (int i = 0; i < tables.Count; i++)
				{
					int num9 = tables[i].Location.DataLocationTemplateRowStart - 1;
					for (int k = 0; k < tables[i].RowsValues.Count; k++)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num9 + k, num, message7, k + 1);
						value = 2;
					}
				}
			}
			else
			{
				int num8 = 0;
				StringBuilder stringBuilder = new StringBuilder();
				for (int i = 0; i < tables.Count; i++)
				{
					num8 = tables[i].Location.DataLocationTemplateRowStart - 1;
					HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
					List<HTKKCMCRowValue> rowsValue = tables[i].RowsValues;
					List<string> list = new List<string>();
					List<string> list2 = new List<string>();
					int r;
					for (r = 0; r < rowsValue.Count; r++)
					{
						List<string> rowValue = tables[i].RowsValues[r].RowValue;
						string text2 = sheetView.Cells[num8 + r, array[0]].Text.Trim();
						if (!(text2 == "1"))
						{
							continue;
						}
						string text3 = sheetView.Cells[num8 + r, num].Text;
						string text4 = sheetView.Cells[num8 + r, num2].Text.Trim();
						double num10 = Convert.ToDouble(sheetView.Cells[num8 + r, num3].Value);
						double num11 = Convert.ToDouble(sheetView.Cells[num8 + r, num4].Value);
						double num12 = Convert.ToDouble(sheetView.Cells[num8 + r, num5].Value);
						if (string.IsNullOrEmpty(text3))
						{
							stringBuilder.Append("Lỗi Họ tên tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")), r + 1);
							value = 2;
						}
						if (string.IsNullOrEmpty(text4))
						{
							stringBuilder.Append("Lỗi thiếu MST và CMT tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")), r + 1);
							value = 2;
						}
						if (!string.IsNullOrEmpty(text4))
						{
							if (!ValidateBase.CheckMaSoThueCCCD(text4))
							{
								stringBuilder.Append("Lỗi sai cấu trúc tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num2, message);
								value = 2;
							}
							else
							{
								int num13 = tables[i].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(rowsValue[r].RowValue[2]));
								if (num13 > 1)
								{
									stringBuilder.Append("Lỗi MST trùng nhau tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
									HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num2, message11, r + 1);
									value = 2;
								}
							}
						}
						if (i == 2)
						{
							if (num10 > 500000000.0 && num10 != num11)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message14, r + 1);
								value = 2;
							}
							else if (num10 <= 500000000.0 && num11 != 0.0 && num11 != num10)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message15, r + 1);
								value = 2;
							}
						}
						else if (PeriodInfo.KieuKy == "M")
						{
							if ((PeriodInfo.Nam.Equals("2026") && ConvertHelpers.ConvertToInt(PeriodInfo.Thang) >= 3) || ConvertHelpers.ConvertToInt(PeriodInfo.Nam) > 2026)
							{
								if (num10 < num11)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message16, r + 1);
									value = 2;
								}
							}
							else if (num10 > 500000000.0 && num10 != num11)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message14, r + 1);
								value = 2;
							}
							else if (num10 <= 500000000.0 && num11 != 0.0 && num11 != num10)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message15, r + 1);
								value = 2;
							}
						}
						else if (PeriodInfo.KieuKy == "Q")
						{
							if ((PeriodInfo.Nam.Equals("2026") && ConvertHelpers.ConvertToInt(PeriodInfo.Quy) >= 2) || ConvertHelpers.ConvertToInt(PeriodInfo.Quy) > 2026)
							{
								if (num10 < num11)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message16, r + 1);
									value = 2;
								}
							}
							else if (num10 > 500000000.0 && num10 != num11)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message14, r + 1);
								value = 2;
							}
							else if (num10 <= 500000000.0 && num11 != 0.0 && num11 != num10)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num4, message15, r + 1);
								value = 2;
							}
						}
						if (num10 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num8 + r, num3, message8, r + 1);
							value = 2;
						}
					}
				}
			}
			switch (value)
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

	private Dictionary<string, string> GetPreperiodValue(HTKKPeriodInfo periodInfo, HTKKCMCToKhaiInfo toKhaiInfo, HTKKHeaderInfo headerInfo, List<string> lstCellId)
	{
		try
		{
			Dictionary<string, string> result = new Dictionary<string, string>();
			List<string> preperiodFileName_ThangQuy = GeneralFunctions.GetPreperiodFileName_ThangQuy(periodInfo, toKhaiInfo);
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text);
			if (directoryInfo.Exists && periodInfo.LoaiToKhai == "C" && preperiodFileName_ThangQuy.Count > 0)
			{
				FileInfo[] files = directoryInfo.GetFiles();
				if (files.Length > 0)
				{
					bool flag = false;
					foreach (string item in preperiodFileName_ThangQuy)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(item))
							{
								result = GeneralFunctions.GetValueInXml(text + "\\" + fileInfo.Name, lstCellId);
								flag = true;
								break;
							}
						}
						if (flag)
						{
							break;
						}
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LockChiTieuToKhai(SheetView sView, HTKKCMCToKhaiInfo toKhaiInfo)
	{
		if (toKhaiInfo.SheetInfos[1] != null)
		{
			sView.Cells[CellId_ct08[1], CellId_ct08[0]].Locked = true;
			sView.Cells[CellId_ct09[1], CellId_ct09[0]].Locked = true;
			sView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = true;
			sView.Cells[CellId_ct11[1], CellId_ct11[0]].Locked = true;
			sView.Cells[CellId_ct12[1], CellId_ct12[0]].Locked = true;
			sView.Cells[CellId_ct13[1], CellId_ct13[0]].Locked = true;
			sView.Cells[CellId_ct14[1], CellId_ct14[0]].Locked = true;
			sView.Cells[CellId_ct15[1], CellId_ct15[0]].Locked = true;
			sView.Cells[CellId_ct16[1], CellId_ct16[0]].Locked = true;
			sView.Cells[CellId_ct17[1], CellId_ct17[0]].Locked = true;
			sView.Cells[CellId_ct18[1], CellId_ct18[0]].Locked = true;
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].Locked = true;
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].Locked = true;
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].Locked = true;
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].Locked = true;
			sView.Cells[CellId_ct08[1], CellId_ct08[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct09[1], CellId_ct09[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct10[1], CellId_ct10[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct11[1], CellId_ct11[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct12[1], CellId_ct12[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct13[1], CellId_ct13[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct14[1], CellId_ct14[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct15[1], CellId_ct15[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct16[1], CellId_ct16[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct17[1], CellId_ct17[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct18[1], CellId_ct18[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
		}
		else
		{
			sView.Cells[CellId_ct08[1], CellId_ct08[0]].Locked = false;
			sView.Cells[CellId_ct09[1], CellId_ct09[0]].Locked = false;
			sView.Cells[CellId_ct10[1], CellId_ct10[0]].Locked = false;
			sView.Cells[CellId_ct11[1], CellId_ct11[0]].Locked = false;
			sView.Cells[CellId_ct12[1], CellId_ct12[0]].Locked = false;
			sView.Cells[CellId_ct13[1], CellId_ct13[0]].Locked = false;
			sView.Cells[CellId_ct14[1], CellId_ct14[0]].Locked = false;
			sView.Cells[CellId_ct15[1], CellId_ct15[0]].Locked = false;
			sView.Cells[CellId_ct16[1], CellId_ct16[0]].Locked = false;
			sView.Cells[CellId_ct17[1], CellId_ct17[0]].Locked = false;
			sView.Cells[CellId_ct18[1], CellId_ct18[0]].Locked = false;
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].Locked = false;
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].Locked = false;
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].Locked = false;
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].Locked = false;
			sView.Cells[CellId_ct08[1], CellId_ct08[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct09[1], CellId_ct09[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct10[1], CellId_ct10[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct11[1], CellId_ct11[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct12[1], CellId_ct12[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct13[1], CellId_ct13[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct14[1], CellId_ct14[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct15[1], CellId_ct15[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct16[1], CellId_ct16[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct17[1], CellId_ct17[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct18[1], CellId_ct18[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		SetFormula(winFpSpread.Sheets[0]);
		LockChiTieuToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
		return true;
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
