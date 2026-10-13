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

public class TncnXsbhdc01Nckt216
{
	private string fomulaColW = "R###+AD###";

	private static string cellId_ct19 = "W_40";

	private Dictionary<string, string> dicFomula = new Dictionary<string, string>();

	private string cellId_ppCt19 = "AD_40";

	private static string cellId_ct20 = "W_41";

	private string cellId_ppCt20 = "AD_41";

	private static string cellId_ct21 = "W_42";

	private string cellId_ppCt21 = "AD_42";

	private static string cellId_ct22 = "W_43";

	private string cellId_ppCt22 = "AD_43";

	private static string cellId_ct23 = "W_44";

	private string cellId_ppCt23 = "AD_44";

	private static string cellId_ct24 = "W_46";

	private string cellId_ppCt24 = "AD_46";

	private static string cellId_ct25 = "W_47";

	private string cellId_ppCt25 = "AD_47";

	private static string cellId_ct26 = "W_48";

	private string cellId_ppCt26 = "AD_48";

	private static string cellId_ct27 = "W_49";

	private string cellId_ppCt27 = "AD_49";

	private static string cellId_ct28 = "W_50";

	private string cellId_ppCt28 = "AD_50";

	private static string cellId_ct29 = "W_52";

	private string cellId_ppCt29 = "AD_52";

	private static string cellId_ct31 = "W_54";

	private string cellId_ppCt31 = "AD_54";

	private static string cellId_ct33 = "W_57";

	private string cellId_ppCt33 = "AD_57";

	private static string cellId_ct34 = "W_58";

	private string cellId_ppCt34 = "AD_58";

	private static string cellId_ct35 = "W_59";

	private string cellId_ppCt35 = "AD_59";

	private static string cellId_ct36 = "W_60";

	private string cellId_ppCt36 = "AD_60";

	private static string cellId_ct37 = "W_61";

	private string cellId_ppCt37 = "AD_61";

	private static string cellId_ct19ps = "R_40";

	private static string cellId_ct20ps = "R_41";

	private static string cellId_ct21ps = "R_42";

	private static string cellId_ct22ps = "R_43";

	private static string cellId_ct23ps = "R_44";

	private static string cellId_ct24ps = "R_46";

	private static string cellId_ct25ps = "R_47";

	private static string cellId_ct26ps = "R_48";

	private static string cellId_ct27ps = "R_49";

	private static string cellId_ct28ps = "R_50";

	private static string cellId_ct29ps = "R_52";

	private static string cellId_ct31ps = "R_54";

	private static string cellId_ct33ps = "R_57";

	private static string cellId_ct34ps = "R_58";

	private static string cellId_ct35ps = "R_59";

	private static string cellId_ct36ps = "R_60";

	private static string cellId_ct37ps = "R_61";

	private static string fmlCt19 = "R40+AD40";

	private static string fmlCt20 = "R41+AD41";

	private static string fmlCt21 = "R42+AD42";

	private static string fmlCt22 = "R43+AD43";

	private static string fmlCt23 = "R44+AD44";

	private static string fmlCt24 = "R46+AD46";

	private static string fmlCt25 = "R47+AD47";

	private static string fmlCt26 = "R48+AD48";

	private static string fmlCt27 = "R49+AD49";

	private static string fmlCt28 = "R50+AD50";

	private static string fmlCt29 = "R52+AD52";

	private static string fmlCt31 = "R54+AD54";

	private static string fmlCt33 = "R57+AD57";

	private static string fmlCt34 = "R58+AD58";

	private static string fmlCt35 = "R59+AD59";

	private static string fmlCt36 = "R60+AD60";

	private static string fmlCt37 = "R61+AD61";

	private static int[] CellId_ct19 => GeneralFunctions.ArrayCellId(cellId_ct19);

	private int[] CellId_ppCt19 => GeneralFunctions.ArrayCellId(cellId_ppCt19);

	private static int[] CellId_ct20 => GeneralFunctions.ArrayCellId(cellId_ct20);

	private int[] CellId_ppCt20 => GeneralFunctions.ArrayCellId(cellId_ppCt20);

	private static int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private int[] CellId_ppCt21 => GeneralFunctions.ArrayCellId(cellId_ppCt21);

	private static int[] CellId_ct22 => GeneralFunctions.ArrayCellId(cellId_ct22);

	private int[] CellId_ppCt22 => GeneralFunctions.ArrayCellId(cellId_ppCt22);

	private static int[] CellId_ct23 => GeneralFunctions.ArrayCellId(cellId_ct23);

	private int[] CellId_ppCt23 => GeneralFunctions.ArrayCellId(cellId_ppCt23);

	private static int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private int[] CellId_ppCt24 => GeneralFunctions.ArrayCellId(cellId_ppCt24);

	private static int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private int[] CellId_ppCt25 => GeneralFunctions.ArrayCellId(cellId_ppCt25);

	private static int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ppCt26 => GeneralFunctions.ArrayCellId(cellId_ppCt26);

	private static int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private int[] CellId_ppCt27 => GeneralFunctions.ArrayCellId(cellId_ppCt27);

	private static int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ppCt28 => GeneralFunctions.ArrayCellId(cellId_ppCt28);

	private static int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ppCt29 => GeneralFunctions.ArrayCellId(cellId_ppCt29);

	private static int[] CellId_ct30 => GeneralFunctions.ArrayCellId("W_53");

	private static int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ppCt31 => GeneralFunctions.ArrayCellId(cellId_ppCt31);

	private static int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private int[] CellId_ppCt33 => GeneralFunctions.ArrayCellId(cellId_ppCt33);

	private static int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	private int[] CellId_ppCt34 => GeneralFunctions.ArrayCellId(cellId_ppCt34);

	private static int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	private int[] CellId_ppCt35 => GeneralFunctions.ArrayCellId(cellId_ppCt35);

	private static int[] CellId_ct36 => GeneralFunctions.ArrayCellId(cellId_ct36);

	private int[] CellId_ppCt36 => GeneralFunctions.ArrayCellId(cellId_ppCt36);

	private static int[] CellId_ct37 => GeneralFunctions.ArrayCellId(cellId_ct37);

	private int[] CellId_ppCt37 => GeneralFunctions.ArrayCellId(cellId_ppCt37);

	private static int[] CellId_ct19ps => GeneralFunctions.ArrayCellId(cellId_ct19ps);

	private static int[] CellId_ct20ps => GeneralFunctions.ArrayCellId(cellId_ct20ps);

	private static int[] CellId_ct21ps => GeneralFunctions.ArrayCellId(cellId_ct21ps);

	private static int[] CellId_ct22ps => GeneralFunctions.ArrayCellId(cellId_ct22ps);

	private static int[] CellId_ct23ps => GeneralFunctions.ArrayCellId(cellId_ct23ps);

	private static int[] CellId_ct24ps => GeneralFunctions.ArrayCellId(cellId_ct24ps);

	private static int[] CellId_ct25ps => GeneralFunctions.ArrayCellId(cellId_ct25ps);

	private static int[] CellId_ct26ps => GeneralFunctions.ArrayCellId(cellId_ct26ps);

	private static int[] CellId_ct27ps => GeneralFunctions.ArrayCellId(cellId_ct27ps);

	private static int[] CellId_ct28ps => GeneralFunctions.ArrayCellId(cellId_ct28ps);

	private static int[] CellId_ct29ps => GeneralFunctions.ArrayCellId(cellId_ct29ps);

	private static int[] CellId_ct30ps => GeneralFunctions.ArrayCellId("R_53");

	private static int[] CellId_ct31ps => GeneralFunctions.ArrayCellId(cellId_ct31ps);

	private static int[] CellId_ct33ps => GeneralFunctions.ArrayCellId(cellId_ct33ps);

	private static int[] CellId_ct34ps => GeneralFunctions.ArrayCellId(cellId_ct34ps);

	private static int[] CellId_ct35ps => GeneralFunctions.ArrayCellId(cellId_ct35ps);

	private static int[] CellId_ct36ps => GeneralFunctions.ArrayCellId(cellId_ct36ps);

	private static int[] CellId_ct37ps => GeneralFunctions.ArrayCellId(cellId_ct37ps);

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 3 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_01_1_XSBHDC.xls");
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
			CheckThang12Quy4(axfpsPhuLuc, listCheckBox, listMaskedTextBox, listCombobox, listRadioButton);
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
			List<string> list = new List<string>();
			list.Add(cellId_ct19);
			list.Add(cellId_ct20);
			list.Add(cellId_ct21);
			list.Add(cellId_ct22);
			list.Add(cellId_ct23);
			list.Add(cellId_ct24);
			list.Add(cellId_ct25);
			list.Add(cellId_ct26);
			list.Add(cellId_ct27);
			list.Add(cellId_ct28);
			list.Add(cellId_ct29);
			list.Add(cellId_ct31);
			list.Add(cellId_ct33);
			list.Add(cellId_ct34);
			list.Add(cellId_ct35);
			list.Add(cellId_ct36);
			list.Add(cellId_ct37);
			List<string> lstCellId = list;
			if (dicFomula.Count == 0)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				List<HTKKCMCCellInfo> staticCells = hTKKCMCSheetInfo.StaticCells;
				for (int i = 0; i < staticCells.Count; i++)
				{
					if (!dicFomula.ContainsKey(staticCells[i].CellTemplate.RowIndex.ToString()) && staticCells[i].CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex("W"))
					{
						dicFomula.Add("W_" + staticCells[i].CellTemplate.RowIndex, fomulaColW.Replace("###", staticCells[i].CellTemplate.RowIndex.ToString()));
					}
				}
			}
			Dictionary<string, string> preperiodValue = GetPreperiodValue(PeriodInfo, ToKhaiInfo, HeaderInfo, lstCellId);
			if (preperiodValue.Count <= 0)
			{
				return;
			}
			foreach (KeyValuePair<string, string> item in preperiodValue)
			{
				string[] array = item.Key.ToString().Split(new char[1] { '_' });
				if (array.Length == 2)
				{
					int num = ConvertHelpers.ConvertToInt(array[1]);
					int num2 = HTKKFile.ColumnLettersToIndex(array[0]);
					int num3 = HTKKFile.ColumnLettersToIndex("AD");
					if (ConvertHelpers.ConvertToDouble1(sView.Cells[num - 1, num2 - 1].Text) == 0.0)
					{
						sView.Cells[num - 1, num2 - 1].Formula = "";
						sView.Cells[num - 1, num2 - 1].Value = item.Value;
					}
					sView.Cells[num - 1, num3 - 1].Value = item.Value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void SetFormula(SheetView sView)
	{
		try
		{
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].Formula = fmlCt19;
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].Formula = fmlCt20;
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].Formula = fmlCt21;
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].Formula = fmlCt22;
			sView.Cells[CellId_ct23[1], CellId_ct23[0]].Formula = fmlCt23;
			sView.Cells[CellId_ct24[1], CellId_ct24[0]].Formula = fmlCt24;
			sView.Cells[CellId_ct25[1], CellId_ct25[0]].Formula = fmlCt25;
			sView.Cells[CellId_ct26[1], CellId_ct26[0]].Formula = fmlCt26;
			sView.Cells[CellId_ct27[1], CellId_ct27[0]].Formula = fmlCt27;
			sView.Cells[CellId_ct28[1], CellId_ct28[0]].Formula = fmlCt28;
			sView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = fmlCt29;
			sView.Cells[CellId_ct31[1], CellId_ct31[0]].Formula = fmlCt31;
			sView.Cells[CellId_ct33[1], CellId_ct33[0]].Formula = fmlCt33;
			sView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = fmlCt34;
			sView.Cells[CellId_ct35[1], CellId_ct35[0]].Formula = fmlCt35;
			sView.Cells[CellId_ct36[1], CellId_ct36[0]].Formula = fmlCt36;
			sView.Cells[CellId_ct37[1], CellId_ct37[0]].Formula = fmlCt37;
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
			sViewTK.Cells[CellId_ct19[1], CellId_ct19[0]].ResetFormula();
			sViewTK.Cells[CellId_ct20[1], CellId_ct20[0]].ResetFormula();
			sViewTK.Cells[CellId_ct21[1], CellId_ct21[0]].ResetFormula();
			sViewTK.Cells[CellId_ct22[1], CellId_ct22[0]].ResetFormula();
			sViewTK.Cells[CellId_ct23[1], CellId_ct23[0]].ResetFormula();
			sViewTK.Cells[CellId_ct24[1], CellId_ct24[0]].ResetFormula();
			sViewTK.Cells[CellId_ct25[1], CellId_ct25[0]].ResetFormula();
			sViewTK.Cells[CellId_ct26[1], CellId_ct26[0]].ResetFormula();
			sViewTK.Cells[CellId_ct27[1], CellId_ct27[0]].ResetFormula();
			sViewTK.Cells[CellId_ct28[1], CellId_ct28[0]].ResetFormula();
			sViewTK.Cells[CellId_ct29[1], CellId_ct29[0]].ResetFormula();
			sViewTK.Cells[CellId_ct31[1], CellId_ct31[0]].ResetFormula();
			sViewTK.Cells[CellId_ct33[1], CellId_ct33[0]].ResetFormula();
			sViewTK.Cells[CellId_ct34[1], CellId_ct34[0]].ResetFormula();
			sViewTK.Cells[CellId_ct35[1], CellId_ct35[0]].ResetFormula();
			sViewTK.Cells[CellId_ct36[1], CellId_ct36[0]].ResetFormula();
			sViewTK.Cells[CellId_ct37[1], CellId_ct37[0]].ResetFormula();
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
			int num = HTKKFile.ColumnLettersToIndex("W");
			if (HTKKFile.ColumnLettersToIndex("R") != (int)iCol)
			{
				return;
			}
			string text = "";
			if (dicFomula.Count <= 0)
			{
				return;
			}
			foreach (KeyValuePair<string, string> item in dicFomula)
			{
				string[] array = item.Key.Split(new char[1] { '_' });
				string value = item.Value;
				int num2 = ConvertHelpers.ConvertToInt(array[1]);
				int num3 = HTKKFile.ColumnLettersToIndex(array[0]);
				string formula = sheetView.Cells[num2 - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Formula;
				if (!HTKKFormater.isNullString(formula))
				{
					text = text + ";" + formula;
				}
				if (!HTKKFormater.isNullString(value) && (!HTKKFormater.isNullString(formula) || num2 == iRow) && (HTKKFormater.isNullString(formula) || text.IndexOf(iRow.ToString()) != -1))
				{
					if ((int)iRow >= 39 && (int)iRow <= 44 && num2 >= 39 && num2 <= 44 && !sheetView.Cells[num2 - 1, num3 - 1].Locked)
					{
						sheetView.Cells[num2 - 1, num3 - 1].Formula = value;
					}
					if ((int)iRow >= 45 && (int)iRow <= 50 && num2 >= 45 && num2 <= 50 && !sheetView.Cells[num2 - 1, num3 - 1].Locked)
					{
						sheetView.Cells[num2 - 1, num3 - 1].Formula = value;
					}
					if ((int)iRow >= 51 && (int)iRow <= 55 && num2 >= 51 && num2 <= 55 && !sheetView.Cells[num2 - 1, num3 - 1].Locked)
					{
						sheetView.Cells[num2 - 1, num3 - 1].Formula = value;
					}
					if ((int)iRow >= 56 && (int)iRow <= 62 && num2 >= 56 && num2 <= 62 && !sheetView.Cells[num2 - 1, num3 - 1].Locked)
					{
						sheetView.Cells[num2 - 1, num3 - 1].Formula = value;
					}
					if (num2 == 44 || num2 == 50)
					{
						text = "";
					}
				}
			}
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
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("H");
			if (num == (int)iCol && iRow >= ToKhaiInfo.SheetInfos[1].Tables[2].Location.DataLocationTemplateRowStart)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_1_BK_XSBHDC/BanHangDC/ct27"));
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Formula = "M" + hTKKCMCCellInfo.CellTemplate.RowIndex;
			}
			else if (num == (int)iCol && iRow >= ToKhaiInfo.SheetInfos[1].Tables[1].Location.DataLocationTemplateRowStart)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_1_BK_XSBHDC/BaoHiem/ct22"));
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Formula = "M" + hTKKCMCCellInfo.CellTemplate.RowIndex;
			}
			else
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_1_BK_XSBHDC/XoSo/ct17"));
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Formula = "M" + hTKKCMCCellInfo.CellTemplate.RowIndex;
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
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
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
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				double num = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct19[1], CellId_ct19[0]].Value);
				int[] array = GeneralFunctions.ArrayCellId("AC_40");
				double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array[1], array[0]].Value);
				if ((num != num2) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct19[1], CellId_ct19[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_31")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct19[1], CellId_ct19[0]);
				}
				double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct20[1], CellId_ct20[0]].Value);
				int[] array2 = GeneralFunctions.ArrayCellId("AC_41");
				double num4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
				if ((num3 != num4) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct20[1], CellId_ct20[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct20[1], CellId_ct20[0]);
				}
				double num5 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct21[1], CellId_ct21[0]].Value);
				int[] array3 = GeneralFunctions.ArrayCellId("AC_42");
				double num6 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array3[1], array3[0]].Value);
				if ((num5 != num6) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct21[1], CellId_ct21[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct21[1], CellId_ct21[0]);
				}
				double num7 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct22[1], CellId_ct22[0]].Value);
				int[] array4 = GeneralFunctions.ArrayCellId("AC_43");
				double num8 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array4[1], array4[0]].Value);
				if ((num7 != num8) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct22[1], CellId_ct22[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct22[1], CellId_ct22[0]);
				}
				double num9 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct23[1], CellId_ct23[0]].Value);
				int[] array5 = GeneralFunctions.ArrayCellId("AC_44");
				double num10 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array5[1], array5[0]].Value);
				if ((num9 != num10) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct23[1], CellId_ct23[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct23[1], CellId_ct23[0]);
				}
				double num11 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct24[1], CellId_ct24[0]].Value);
				int[] array6 = GeneralFunctions.ArrayCellId("AC_46");
				double num12 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array6[1], array6[0]].Value);
				if ((num11 != num12) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct24[1], CellId_ct24[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct24[1], CellId_ct24[0]);
				}
				double num13 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct25[1], CellId_ct25[0]].Value);
				int[] array7 = GeneralFunctions.ArrayCellId("AC_47");
				double num14 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array7[1], array7[0]].Value);
				if ((num13 != num14) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct25[1], CellId_ct25[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct25[1], CellId_ct25[0]);
				}
				double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct26[1], CellId_ct26[0]].Value);
				int[] array8 = GeneralFunctions.ArrayCellId("AC_48");
				double num16 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array8[1], array8[0]].Value);
				if ((num15 != num16) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct26[1], CellId_ct26[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct26[1], CellId_ct26[0]);
				}
				double num17 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct27[1], CellId_ct27[0]].Value);
				int[] array9 = GeneralFunctions.ArrayCellId("AC_49");
				double num18 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array9[1], array9[0]].Value);
				if ((num17 != num18) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct27[1], CellId_ct27[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_39")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct27[1], CellId_ct27[0]);
				}
				double num19 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct28[1], CellId_ct28[0]].Value);
				int[] array10 = GeneralFunctions.ArrayCellId("AC_50");
				double num20 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array10[1], array10[0]].Value);
				if ((num19 != num20) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct28[1], CellId_ct28[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct28[1], CellId_ct28[0]);
				}
				double num21 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct33[1], CellId_ct33[0]].Value);
				int[] array11 = GeneralFunctions.ArrayCellId("AC_57");
				double num22 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array11[1], array11[0]].Value);
				if ((num21 != num22) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct33[1], CellId_ct33[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct33[1], CellId_ct33[0]);
				}
				double num23 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct34[1], CellId_ct34[0]].Value);
				int[] array12 = GeneralFunctions.ArrayCellId("AC_58");
				double num24 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array12[1], array12[0]].Value);
				if ((num23 != num24) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct34[1], CellId_ct34[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct34[1], CellId_ct34[0]);
				}
				double num25 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Value);
				int[] array13 = GeneralFunctions.ArrayCellId("AC_59");
				double num26 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array13[1], array13[0]].Value);
				if ((num25 != num26) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct35[1], CellId_ct35[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct35[1], CellId_ct35[0]);
				}
				double num27 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct36[1], CellId_ct36[0]].Value);
				int[] array14 = GeneralFunctions.ArrayCellId("AC_60");
				double num28 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array14[1], array14[0]].Value);
				if ((num27 != num28) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct36[1], CellId_ct36[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct36[1], CellId_ct36[0]);
				}
				double num29 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Value);
				int[] array15 = GeneralFunctions.ArrayCellId("AC_61");
				double num30 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array15[1], array15[0]].Value);
				if ((num29 != num30) & (PeriodInfo.LoaiToKhai != "B"))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct37[1], CellId_ct37[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_45")));
					if (value == 0)
					{
						value = 1;
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct37[1], CellId_ct37[0]);
				}
			}
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
			int column = HTKKFile.ColumnLettersToIndex("G") - 1;
			int num4 = HTKKFile.ColumnLettersToIndex("H") - 1;
			int column2 = HTKKFile.ColumnLettersToIndex("I") - 1;
			int num5 = HTKKFile.ColumnLettersToIndex("J") - 1;
			int num6 = HTKKFile.ColumnLettersToIndex("K") - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			string message = HTKKFile.GetMessage("0047");
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_24"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_26"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27"));
			int num7 = 0;
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = 0; i < tables.Count; i++)
			{
				num7 = tables[i].Location.DataLocationTemplateRowStart - 1;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				List<HTKKCMCRowValue> rowsValue = tables[i].RowsValues;
				List<string> list = new List<string>();
				List<string> list2 = new List<string>();
				int r;
				for (r = 0; r < rowsValue.Count; r++)
				{
					List<string> rowValue = tables[i].RowsValues[r].RowValue;
					bool flag = GeneralFunctions.HaveToValidateRow(rowsInfo, rowValue.ToArray());
					HTKKFile.RemoveValidate(sheetView, num7 + r, num);
					HTKKFile.RemoveValidate(sheetView, num7 + r, num2);
					HTKKFile.RemoveValidate(sheetView, num7 + r, num3);
					HTKKFile.RemoveValidate(sheetView, num7 + r, num4);
					HTKKFile.RemoveValidate(sheetView, num7 + r, num5);
					if (!flag)
					{
						continue;
					}
					string text = sheetView.Cells[num7 + r, num].Text;
					if (string.IsNullOrEmpty(text))
					{
						stringBuilder.Append("Lỗi Họ tên tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")), r + 1);
						value = 2;
					}
					string text2 = sheetView.Cells[num7 + r, num2].Text.Trim();
					string text3 = sheetView.Cells[num7 + r, num3].Text.Trim();
					if (string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text3))
					{
						stringBuilder.Append("Lỗi thiếu MST và CMT tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")), r + 1);
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
						value = 2;
					}
					else
					{
						if (!string.IsNullOrEmpty(text2))
						{
							if (!ValidateBase.CheckMaSoThue(text2))
							{
								stringBuilder.Append("Lỗi sai cấu trúc tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num2, message);
								value = 2;
							}
							else
							{
								int num8 = 0;
								int num9 = 0;
								int num10 = tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(rowsValue[r].RowValue[1]));
								num10 += tables[1].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(rowsValue[r].RowValue[1]));
								if (num10 > 1)
								{
									stringBuilder.Append("Lỗi MST trùng nhau tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
									HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_21")), r + 1);
									value = 2;
								}
							}
						}
						if (!string.IsNullOrEmpty(text3))
						{
							if (text3.IndexOf(" ") > -1)
							{
								stringBuilder.Append("Lỗi CMT có dấu cách tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")), r + 1);
								value = 2;
							}
							else
							{
								int num10 = tables[0].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(rowsValue[r].RowValue[2]));
								num10 += tables[1].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(rowsValue[r].RowValue[2]));
								if (num10 > 1)
								{
									stringBuilder.Append("Lỗi CMT trùng nhau tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
									HTKKFile.SetNodeBackGroundYellow(sheetView, num7 + r, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_22")));
									if (value != 2)
									{
										value = 1;
									}
								}
							}
						}
					}
					double num11 = double.Parse(sheetView.Cells[num7 + r, column].Value.ToString());
					double num12 = double.Parse(sheetView.Cells[num7 + r, num4].Value.ToString());
					if (num12 > num11)
					{
						stringBuilder.Append("Lỗi Chỉ tiêu [11] phải <= [10] tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num4, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_23")), r + 1);
						value = 2;
					}
					double num13 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num7 + r, column2].Value);
					double num14 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num7 + r, num5].Value);
					if (num14 > Math.Round(num13 * 0.5, 0))
					{
						stringBuilder.Append("Lỗi Chỉ tiêu [13] phải <= [12] * 50% tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
						HTKKFile.SetNodeBackGroundPink(sheetView, num7 + r, num5, message4, r + 1);
						value = 2;
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
}
