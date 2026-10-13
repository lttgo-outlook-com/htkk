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

public class TncnXsbhdc01TT40Nckt249
{
	private string fomulaColW = "R###+AD###";

	private string MA_COLOR_LOCKED = "#e8f5ff";

	private string MA_COLOR_UNLOCK = "#FFFFFF";

	private static string cellId_ct13 = "W_40";

	private Dictionary<string, string> dicFomula = new Dictionary<string, string>();

	private string cellId_ppCt13 = "AD_40";

	private static string cellId_ct14 = "W_41";

	private string cellId_ppCt14 = "AD_41";

	private static string cellId_ct15 = "W_42";

	private string cellId_ppCt15 = "AD_42";

	private static string cellId_ct16 = "W_44";

	private string cellId_ppCt16 = "AD_44";

	private static string cellId_ct17 = "W_46";

	private string cellId_ppCt17 = "AD_46";

	private static string cellId_ct18 = "W_47";

	private string cellId_ppCt18 = "AD_47";

	private static string cellId_ct19 = "W_48";

	private string cellId_ppCt19 = "AD_48";

	private static string cellId_ct20 = "W_50";

	private string cellId_ppCt20 = "AD_50";

	private static string cellId_ct21 = "W_52";

	private string cellId_ppCt21 = "AD_52";

	private static string cellId_ct24 = "W_57";

	private string cellId_ppCt24 = "AD_57";

	private static string cellId_ct25 = "W_58";

	private string cellId_ppCt25 = "AD_58";

	private static string cellId_ct26 = "W_59";

	private string cellId_ppCt26 = "AD_59";

	private static string cellId_ct36 = "W_60";

	private static string cellId_ct27 = "W_61";

	private string cellId_ppCt27 = "AD_61";

	private static string cellId_ct13ps = "R_40";

	private static string cellId_ct14ps = "R_41";

	private static string cellId_ct15ps = "R_42";

	private static string cellId_ct16ps = "R_44";

	private static string cellId_ct17ps = "R_46";

	private static string cellId_ct18ps = "R_47";

	private static string cellId_ct19ps = "R_48";

	private static string cellId_ct20ps = "R_50";

	private static string cellId_ct21ps = "R_52";

	private static string cellId_ct24ps = "R_57";

	private static string cellId_ct25ps = "R_58";

	private static string cellId_ct26ps = "R_59";

	private static string cellId_ct27ps = "R_61";

	private static string fmlCt13 = "AC40";

	private static string fmlCt14 = "AC41";

	private static string fmlCt15 = "AC42";

	private static string fmlCt16 = "AC44";

	private static string fmlCt17 = "AC46";

	private static string fmlCt18 = "AC47";

	private static string fmlCt19 = "AC48";

	private static string fmlCt20 = "AC50";

	private static string fmlCt21 = "AC52";

	private static string fmlCt22 = "AC53";

	private static string fmlCt23 = "AC55";

	private static string fmlCt24 = "AC57";

	private static string fmlCt25 = "AC58";

	private static string fmlCt26 = "AC59";

	private static string fmlCt27 = "AC61";

	private static int[] CellId_ct13 => GeneralFunctions.ArrayCellId(cellId_ct13);

	private int[] CellId_ppCt13 => GeneralFunctions.ArrayCellId(cellId_ppCt13);

	private static int[] CellId_ct14 => GeneralFunctions.ArrayCellId(cellId_ct14);

	private int[] CellId_ppCt14 => GeneralFunctions.ArrayCellId(cellId_ppCt14);

	private static int[] CellId_ct15 => GeneralFunctions.ArrayCellId(cellId_ct15);

	private int[] CellId_ppCt15 => GeneralFunctions.ArrayCellId(cellId_ppCt15);

	private static int[] CellId_ct16 => GeneralFunctions.ArrayCellId(cellId_ct16);

	private int[] CellId_ppCt16 => GeneralFunctions.ArrayCellId(cellId_ppCt16);

	private static int[] CellId_ct17 => GeneralFunctions.ArrayCellId(cellId_ct17);

	private int[] CellId_ppCt17 => GeneralFunctions.ArrayCellId(cellId_ppCt17);

	private static int[] CellId_ct18 => GeneralFunctions.ArrayCellId(cellId_ct18);

	private int[] CellId_ppCt18 => GeneralFunctions.ArrayCellId(cellId_ppCt18);

	private static int[] CellId_ct19 => GeneralFunctions.ArrayCellId(cellId_ct19);

	private int[] CellId_ppCt19 => GeneralFunctions.ArrayCellId(cellId_ppCt19);

	private static int[] CellId_ct20 => GeneralFunctions.ArrayCellId(cellId_ct20);

	private int[] CellId_ppCt20 => GeneralFunctions.ArrayCellId(cellId_ppCt20);

	private static int[] CellId_ct21 => GeneralFunctions.ArrayCellId(cellId_ct21);

	private int[] CellId_ppCt21 => GeneralFunctions.ArrayCellId(cellId_ppCt21);

	private static int[] CellId_ct22 => GeneralFunctions.ArrayCellId("W_53");

	private static int[] CellId_ct23 => GeneralFunctions.ArrayCellId("W_55");

	private static int[] CellId_ct24 => GeneralFunctions.ArrayCellId(cellId_ct24);

	private int[] CellId_ppCt24 => GeneralFunctions.ArrayCellId(cellId_ppCt24);

	private static int[] CellId_ct25 => GeneralFunctions.ArrayCellId(cellId_ct25);

	private int[] CellId_ppCt25 => GeneralFunctions.ArrayCellId(cellId_ppCt25);

	private static int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ppCt26 => GeneralFunctions.ArrayCellId(cellId_ppCt26);

	private static int[] CellId_ct36 => GeneralFunctions.ArrayCellId(cellId_ct36);

	private static int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private int[] CellId_ppCt27 => GeneralFunctions.ArrayCellId(cellId_ppCt27);

	private static int[] CellId_ct13ps => GeneralFunctions.ArrayCellId(cellId_ct13ps);

	private static int[] CellId_ct14ps => GeneralFunctions.ArrayCellId(cellId_ct14ps);

	private static int[] CellId_ct15ps => GeneralFunctions.ArrayCellId(cellId_ct15ps);

	private static int[] CellId_ct16ps => GeneralFunctions.ArrayCellId(cellId_ct16ps);

	private static int[] CellId_ct17ps => GeneralFunctions.ArrayCellId(cellId_ct17ps);

	private static int[] CellId_ct18ps => GeneralFunctions.ArrayCellId(cellId_ct18ps);

	private static int[] CellId_ct19ps => GeneralFunctions.ArrayCellId(cellId_ct19ps);

	private static int[] CellId_ct20ps => GeneralFunctions.ArrayCellId(cellId_ct20ps);

	private static int[] CellId_ct21ps => GeneralFunctions.ArrayCellId(cellId_ct21ps);

	private static int[] CellId_ct22ps => GeneralFunctions.ArrayCellId("R_53");

	private static int[] CellId_ct24ps => GeneralFunctions.ArrayCellId(cellId_ct24ps);

	private static int[] CellId_ct25ps => GeneralFunctions.ArrayCellId(cellId_ct25ps);

	private static int[] CellId_ct26ps => GeneralFunctions.ArrayCellId(cellId_ct26ps);

	private static int[] CellId_ct27ps => GeneralFunctions.ArrayCellId(cellId_ct27ps);

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 3 && idxCol == 1)
			{
				ShowBangKe("Bang_Ke_01_1_XSBHDC_TT40.xls");
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
			list.Add(cellId_ct13);
			list.Add(cellId_ct14);
			list.Add(cellId_ct15);
			list.Add(cellId_ct16);
			list.Add(cellId_ct17);
			list.Add(cellId_ct18);
			list.Add(cellId_ct19);
			list.Add(cellId_ct27);
			list.Add(cellId_ct20);
			list.Add(cellId_ct21);
			list.Add(cellId_ct24);
			list.Add(cellId_ct25);
			list.Add(cellId_ct26);
			list.Add(cellId_ct36);
			list.Add(cellId_ct27);
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
			if (preperiodValue.Count > 0)
			{
				foreach (KeyValuePair<string, string> item in preperiodValue)
				{
					string[] array = item.Key.ToString().Split(new char[1] { '_' });
					if (array.Length == 2)
					{
						int num = ConvertHelpers.ConvertToInt(array[1]);
						int num2 = HTKKFile.ColumnLettersToIndex(array[0]);
						int num3 = HTKKFile.ColumnLettersToIndex("AD");
						if (Convert.ToDouble(sView.Cells[num - 1, num2 - 1].Value) == 0.0)
						{
							sView.Cells[num - 1, num2 - 1].Formula = "";
							sView.Cells[num - 1, num2 - 1].Value = item.Value;
						}
						sView.Cells[num - 1, num3 - 1].Value = item.Value;
					}
				}
			}
			LockChiTieuToKhai(sView, ToKhaiInfo);
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
			sView.Cells[CellId_ct13[1], CellId_ct13[0]].Formula = fmlCt13;
			sView.Cells[CellId_ct14[1], CellId_ct14[0]].Formula = fmlCt14;
			sView.Cells[CellId_ct15[1], CellId_ct15[0]].Formula = fmlCt15;
			sView.Cells[CellId_ct16[1], CellId_ct16[0]].Formula = fmlCt16;
			sView.Cells[CellId_ct17[1], CellId_ct17[0]].Formula = fmlCt17;
			sView.Cells[CellId_ct18[1], CellId_ct18[0]].Formula = fmlCt18;
			sView.Cells[CellId_ct19[1], CellId_ct19[0]].Formula = fmlCt19;
			sView.Cells[CellId_ct27[1], CellId_ct27[0]].Formula = fmlCt27;
			sView.Cells[CellId_ct20[1], CellId_ct20[0]].Formula = fmlCt20;
			sView.Cells[CellId_ct21[1], CellId_ct21[0]].Formula = fmlCt21;
			sView.Cells[CellId_ct22[1], CellId_ct22[0]].Formula = fmlCt22;
			sView.Cells[CellId_ct23[1], CellId_ct23[0]].Formula = fmlCt23;
			sView.Cells[CellId_ct24[1], CellId_ct24[0]].Formula = fmlCt24;
			sView.Cells[CellId_ct25[1], CellId_ct25[0]].Formula = fmlCt25;
			sView.Cells[CellId_ct26[1], CellId_ct26[0]].Formula = fmlCt26;
			sView.Cells[CellId_ct27[1], CellId_ct27[0]].Formula = fmlCt27;
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
			int num = HTKKFile.ColumnLettersToIndex("H");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			for (int i = 0; i < tables.Count; i++)
			{
				int num2 = tables[i].Location.DataLocationTemplateRowStart - 1;
				for (int j = 0; j < tables[i].RowsValues.Count; j++)
				{
					int num3 = num2 + j + 1;
					string formula = "IF(G" + num3 + ">100000000,G" + num3 + ",0)";
					string formula2 = sView.Cells[num3 - 1, num - 1].Formula;
					sView.Cells[num3 - 1, num - 1].Formula = formula;
				}
			}
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
			sViewTK.Cells[CellId_ct13[1], CellId_ct13[0]].ResetFormula();
			sViewTK.Cells[CellId_ct14[1], CellId_ct14[0]].ResetFormula();
			sViewTK.Cells[CellId_ct15[1], CellId_ct15[0]].ResetFormula();
			sViewTK.Cells[CellId_ct22[1], CellId_ct22[0]].ResetFormula();
			sViewTK.Cells[CellId_ct16[1], CellId_ct16[0]].ResetFormula();
			sViewTK.Cells[CellId_ct17[1], CellId_ct17[0]].ResetFormula();
			sViewTK.Cells[CellId_ct18[1], CellId_ct18[0]].ResetFormula();
			sViewTK.Cells[CellId_ct19[1], CellId_ct19[0]].ResetFormula();
			sViewTK.Cells[CellId_ct27[1], CellId_ct27[0]].ResetFormula();
			sViewTK.Cells[CellId_ct20[1], CellId_ct20[0]].ResetFormula();
			sViewTK.Cells[CellId_ct21[1], CellId_ct21[0]].ResetFormula();
			sViewTK.Cells[CellId_ct24[1], CellId_ct24[0]].ResetFormula();
			sViewTK.Cells[CellId_ct25[1], CellId_ct25[0]].ResetFormula();
			sViewTK.Cells[CellId_ct26[1], CellId_ct26[0]].ResetFormula();
			sViewTK.Cells[CellId_ct36[1], CellId_ct36[0]].ResetFormula();
			sViewTK.Cells[CellId_ct27[1], CellId_ct27[0]].ResetFormula();
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
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("H");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			if (iCol == num)
			{
				sheetView.Cells[(int)iRow - 1, num2 - 1].Value = sheetView.Cells[(int)iRow - 1, num3 - 1].Value;
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
			int[] array = GeneralFunctions.ArrayCellId("O_4");
			string text = sheetView.Cells[array[1], array[0]].Text.Trim();
			if (text == "0")
			{
				for (int i = 0; i < tables.Count; i++)
				{
					int num8 = tables[i].Location.DataLocationTemplateRowStart - 1;
					for (int k = 0; k < tables[i].RowsValues.Count; k++)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num8 + k, num, message7, k + 1);
						value = 2;
					}
				}
			}
			else
			{
				int num9 = 0;
				StringBuilder stringBuilder = new StringBuilder();
				for (int i = 0; i < tables.Count; i++)
				{
					num9 = tables[i].Location.DataLocationTemplateRowStart - 1;
					HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
					List<HTKKCMCRowValue> rowsValue = tables[i].RowsValues;
					List<string> list = new List<string>();
					List<string> list2 = new List<string>();
					int r;
					for (r = 0; r < rowsValue.Count; r++)
					{
						List<string> rowValue = tables[i].RowsValues[r].RowValue;
						string text2 = sheetView.Cells[num9 + r, array[0]].Text.Trim();
						if (!(text2 == "1"))
						{
							continue;
						}
						string text3 = sheetView.Cells[num9 + r, num].Text;
						string text4 = sheetView.Cells[num9 + r, num2].Text.Trim();
						string text5 = sheetView.Cells[num9 + r, num3].Text.Trim();
						double num10 = Convert.ToDouble(sheetView.Cells[num9 + r, num4].Value);
						double num11 = Convert.ToDouble(sheetView.Cells[num9 + r, num5].Value);
						if (string.IsNullOrEmpty(text3))
						{
							stringBuilder.Append("Lỗi Họ tên tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")), r + 1);
							value = 2;
						}
						if (string.IsNullOrEmpty(text4) && string.IsNullOrEmpty(text5))
						{
							stringBuilder.Append("Lỗi thiếu MST và CMT tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num2, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")), r + 1);
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
							value = 2;
						}
						else
						{
							if (!string.IsNullOrEmpty(text4))
							{
								if (!ValidateBase.CheckMaSoThueCCCD(text4))
								{
									stringBuilder.Append("Lỗi sai cấu trúc tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
									HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num2, message);
									value = 2;
								}
								else
								{
									int num12 = tables[i].RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(rowsValue[r].RowValue[2]));
									if (num12 > 1)
									{
										stringBuilder.Append("Lỗi MST trùng nhau tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
										HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num2, message11, r + 1);
										value = 2;
									}
								}
							}
							if (!string.IsNullOrEmpty(text5) && text5.IndexOf(" ") > -1)
							{
								stringBuilder.Append("Lỗi CMT có dấu cách tại dòng số " + (r + 1) + "\n" + Environment.NewLine);
								HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num3, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")), r + 1);
								value = 2;
							}
						}
						if (num10 == 0.0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num4, message8, r + 1);
							value = 2;
						}
						if (i == 2)
						{
							if (num11 != num10)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num5, message13, r + 1);
								value = 2;
							}
						}
						else if (num10 > 100000000.0 && num11 != num10)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num5, message12, r + 1);
							value = 2;
						}
						else if (num11 != 0.0 && num11 != num10)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num9 + r, num5, message9, r + 1);
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
			if (!sView.Cells[CellId_ct13[1], CellId_ct13[0]].Locked)
			{
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
				sView.Cells[CellId_ct23[1], CellId_ct23[0]].Locked = true;
				sView.Cells[CellId_ct24[1], CellId_ct24[0]].Locked = true;
				sView.Cells[CellId_ct25[1], CellId_ct25[0]].Locked = true;
				sView.Cells[CellId_ct26[1], CellId_ct26[0]].Locked = true;
				sView.Cells[CellId_ct27[1], CellId_ct27[0]].Locked = true;
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
				sView.Cells[CellId_ct23[1], CellId_ct23[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				sView.Cells[CellId_ct24[1], CellId_ct24[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				sView.Cells[CellId_ct25[1], CellId_ct25[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				sView.Cells[CellId_ct26[1], CellId_ct26[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
				sView.Cells[CellId_ct27[1], CellId_ct27[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_LOCKED);
			}
		}
		else if (sView.Cells[CellId_ct13[1], CellId_ct13[0]].Locked)
		{
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
			sView.Cells[CellId_ct23[1], CellId_ct23[0]].Locked = false;
			sView.Cells[CellId_ct24[1], CellId_ct24[0]].Locked = false;
			sView.Cells[CellId_ct25[1], CellId_ct25[0]].Locked = false;
			sView.Cells[CellId_ct26[1], CellId_ct26[0]].Locked = false;
			sView.Cells[CellId_ct27[1], CellId_ct27[0]].Locked = false;
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
			sView.Cells[CellId_ct23[1], CellId_ct23[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct24[1], CellId_ct24[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct25[1], CellId_ct25[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct26[1], CellId_ct26[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
			sView.Cells[CellId_ct27[1], CellId_ct27[0]].BackColor = ColorTranslator.FromHtml(MA_COLOR_UNLOCK);
		}
	}

	public bool UpdateFormula_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		SetFormula(winFpSpread.Sheets[0]);
		SetFormulaPL(winFpSpread.Sheets[0], ToKhaiInfo);
		LockChiTieuToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
		return true;
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SetFormula(winFpSpread.Sheets[0]);
			LockChiTieuToKhai(winFpSpread.Sheets[0], ToKhaiInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
