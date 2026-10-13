using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn01BNckt207
{
	private Regex regexCt26_mulVal = new Regex("^0$|^([1-9][0-9]?)(;[1-9][0-9]?)*$");

	private bool loadedValueCbxUuDai;

	private string cellId_ckbUuDai = "H_47";

	private string cellId_ct26 = "K_47";

	private string pathCt26 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct26";

	private NumberCellType numberCellType_ct26 = new NumberCellType();

	private string cellId_ct29 = "K_50";

	private string formulaCt29 = "ROUND(K41*K42/100*K47/100,0)";

	private string pathCt29 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct29";

	private string cellId_ct30 = "K_51";

	private string cellId_ct30a = "K_52";

	private string cellId_ct30b = "K_53";

	private string cellId_ct31 = "K_54";

	private string pathCt31 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct31";

	private string cellId_ct32 = "F_56";

	private string cellId_ct32a = "H_58";

	private string cellIdValue_ct32a = "P_58";

	private string cellId_ct32b = "H_60";

	private string cellId_ct32c = "H_62";

	private string cellId_ct32d = "H_64";

	private string cellId_ngayKy = "J_69";

	private int[] CellId_ckbUuDai => GeneralFunctions.ArrayCellId(cellId_ckbUuDai);

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private NumberCellType NumberCellType_ct26
	{
		get
		{
			numberCellType_ct26.DecimalPlaces = 3;
			numberCellType_ct26.DecimalSeparator = ".";
			numberCellType_ct26.MaximumValue = 100.0;
			return numberCellType_ct26;
		}
	}

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct30a => GeneralFunctions.ArrayCellId(cellId_ct30a);

	private int[] CellId_ct30b => GeneralFunctions.ArrayCellId(cellId_ct30b);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_ct32a => GeneralFunctions.ArrayCellId(cellId_ct32a);

	private int[] CellIdValue_ct32a => GeneralFunctions.ArrayCellId(cellIdValue_ct32a);

	private int[] CellId_ct32b => GeneralFunctions.ArrayCellId(cellId_ct32b);

	private int[] CellId_ct32c => GeneralFunctions.ArrayCellId(cellId_ct32c);

	private int[] CellId_ct32d => GeneralFunctions.ArrayCellId(cellId_ct32d);

	private int[] CellId_ngayKy => GeneralFunctions.ArrayCellId(cellId_ngayKy);

	private int[] CellId_E25 => GeneralFunctions.ArrayCellId("E_25");

	private int[] CellId_E30 => GeneralFunctions.ArrayCellId("E_30");

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E38 => GeneralFunctions.ArrayCellId("E_38");

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!GeneralFunctions.CheckNamTaiChinh(HeaderInfo.NgayBatDauNamTaiChinh))
			{
				return false;
			}
			if (!ValidatePeriod_mskQuy(PeriodInfo))
			{
				return false;
			}
			int num = int.Parse(PeriodInfo.Nam);
			int num2 = int.Parse(PeriodInfo.Quy);
			if (num <= 1999)
			{
				frmMessage frmMessage = new frmMessage("0067", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			if (num > 2014 || (num == 2014 && num2 == 4))
			{
				frmMessage frmMessage = new frmMessage("0341", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool ValidatePeriod_mskQuy(HTKKPeriodInfo PeriodInfo)
	{
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			DateTime dateTime = GeneralFunctions.NgayDauKyTinhThue(PeriodInfo);
			DateTime dateTime2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
			if (dateTime > dateTime2)
			{
				frmMessage frmMessage = new frmMessage("0045", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			loadedValueCbxUuDai = Convert.ToBoolean(sView.Cells[CellId_ckbUuDai[1], CellId_ckbUuDai[0]].Value);
			XuLyCheckBox_ThueSuatUuDai(sView, ToKhaiInfo, isBindData: true, loadedValueCbxUuDai);
			XuLyCheckBox_Ct32(sView);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
				if (idxRow == CellId_ct26[1] && idxCol == CellId_ct26[0] && Convert.ToBoolean(sheetView.Cells[CellId_ckbUuDai[1], CellId_ckbUuDai[0]].Value))
				{
					GeneralFunctions.CompareKeycode(e, 2);
				}
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeColumnIndex == CellId_ckbUuDai[0] && activeRowIndex == CellId_ckbUuDai[1])
			{
				XuLyCheckBox_ThueSuatUuDai(winFpSpread.ActiveSheet, ToKhaiInfo, isBindData: false, loadedValueCbxUuDai);
			}
			else if (activeColumnIndex == CellId_ct32[0] && activeRowIndex == CellId_ct32[1])
			{
				XuLyCheckBox_Ct32(winFpSpread.ActiveSheet);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo);
		if (winFpSpread.TabStrip[0].BackColor != Color.White || winFpSpread.TabStrip[1].BackColor != Color.White || winFpSpread.TabStrip[2].BackColor != Color.White)
		{
			return false;
		}
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = 0;
			string text = sheetView.Cells[CellId_ct26[1], CellId_ct26[0]].Text;
			if (Convert.ToBoolean(sheetView.Cells[CellId_ckbUuDai[1], CellId_ckbUuDai[0]].Value))
			{
				if (!regexCt26_mulVal.IsMatch(text))
				{
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct26[1], CellId_ct26[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct26[1], CellId_ct26[0]);
			}
			if (Convert.ToBoolean(sheetView.Cells[CellId_ct32[1], CellId_ct32[0]].Value))
			{
				string text2 = sheetView.Cells[CellIdValue_ct32a[1], CellIdValue_ct32a[0]].Text;
				if (string.IsNullOrEmpty(text2) || int.Parse(text2) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct32a[1], CellId_ct32a[0], HTKKFile.GetMessage(winFpSpread, CellId_E30[1], CellId_E30[0]));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct32a[1], CellId_ct32a[0]);
				}
				string text3 = sheetView.Cells[CellId_ct32b[1], CellId_ct32b[0]].Text;
				DateTime result;
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct32b[1], CellId_ct32b[0], HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]));
					num = 2;
				}
				else if (!DateTime.TryParseExact(text3, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct32b[1], CellId_ct32b[0], "Ngày không hợp lệ");
					num = 2;
				}
				else if (result <= GeneralFunctions.GetHanNopToKhai(PeriodInfo))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct32b[1], CellId_ct32b[0], HTKKFile.GetMessage(winFpSpread, CellId_E38[1], CellId_E38[0]));
					num = 1;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct32b[1], CellId_ct32b[0]);
				}
				object value = sheetView.Cells[CellId_ct32c[1], CellId_ct32c[0]].Value;
				double num2 = Convert.ToDouble(sheetView.Cells[CellId_ct31[1], CellId_ct31[0]].Value);
				if (value == null || (double)value == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct32c[1], CellId_ct32c[0], HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct32c[1], CellId_ct32c[0]);
				}
				object value2 = sheetView.Cells[CellId_ct32d[1], CellId_ct32d[0]].Value;
				if (value2 == null || Convert.ToDouble(value2) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct32d[1], CellId_ct32d[0], HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct32d[1], CellId_ct32d[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct32a[1], CellId_ct32a[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct32b[1], CellId_ct32b[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct32c[1], CellId_ct32c[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct32d[1], CellId_ct32d[0]);
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
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = 0;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int column = rowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			int num3 = rowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int column2 = rowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int column3 = rowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int column4 = rowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int column5 = rowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			string[] array = new string[8];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num2 + i, column].Text;
				array[1] = sheetView.Cells[num2 + i, num3].Text;
				array[2] = sheetView.Cells[num2 + i, column2].Text;
				array[3] = sheetView.Cells[num2 + i, column3].Text;
				array[4] = sheetView.Cells[num2 + i, column4].Text;
				array[5] = sheetView.Cells[num2 + i, column5].Text;
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, num3);
				}
				else if (!string.IsNullOrEmpty(array[1]) & !ValidateBase.CheckMaSoThue(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, num3, HTKKFile.GetMessage("0047"));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, num3);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_1_TNDN/BangThueTNDNTamTinh/tongCot12");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_1_TNDN/ct6");
			int column6 = hTKKCMCCellInfo2.CellTemplate.ColIndex - 1;
			int row = hTKKCMCCellInfo2.CellTemplate.RowIndex - 1;
			int num4 = hTKKCMCCellInfo.CellTemplate.ColIndex - 1;
			int num5 = hTKKCMCCellInfo.CellTemplate.RowIndex - 1;
			double num6 = Convert.ToDouble(sheetView.Cells[row, column6].Value);
			double num7 = Convert.ToDouble(sheetView.Cells[num5, num4].Value);
			if (num7 != num6)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num5, num4, HTKKFile.GetMessage("01B-TNDN009"));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num5, num4);
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
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int num = 0;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			SheetView sheetView = winFpSpread.Sheets[2];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			HTKKCMCStaticOfTable header = tables[0].Header;
			string[] array = new string[8];
			string[] array2 = new string[8];
			for (int i = 0; i < tables.Count; i++)
			{
				HTKKCMCStaticOfTable header2 = tables[i].Header;
				HTKKCMCRowInfo rowsInfo = tables[i].RowsInfo;
				int num2 = tables[i].Location.DataLocationTemplateRowStart - 1;
				array[1] = sheetView.Cells[header2.StaticCells[1].CellTemplate.RowIndex - 1, header2.StaticCells[1].CellTemplate.ColIndex - 1].Text;
				array[2] = sheetView.Cells[header2.StaticCells[2].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1].Text;
				array[4] = sheetView.Cells[header2.StaticCells[4].CellTemplate.RowIndex - 1, header2.StaticCells[4].CellTemplate.ColIndex - 1].Value.ToString();
				array[5] = sheetView.Cells[header2.StaticCells[5].CellTemplate.RowIndex - 1, header2.StaticCells[5].CellTemplate.ColIndex - 1].Value.ToString();
				bool flag = GeneralFunctions.HaveToValidateHeaderRow(sheetView, header, tables[i].iColumnSTT, array);
				bool flag2 = false;
				for (int j = 0; j < tables[i].RowsValues.Count; j++)
				{
					array2[1] = sheetView.Cells[num2 + j, rowsInfo.Cells[1].CellTemplate.ColIndex - 1].Text;
					array2[2] = sheetView.Cells[num2 + j, rowsInfo.Cells[2].CellTemplate.ColIndex - 1].Text;
					array2[3] = sheetView.Cells[num2 + j, rowsInfo.Cells[3].CellTemplate.ColIndex - 1].Text;
					array2[4] = sheetView.Cells[num2 + j, rowsInfo.Cells[4].CellTemplate.ColIndex - 1].Value.ToString();
					array2[5] = sheetView.Cells[num2 + j, rowsInfo.Cells[5].CellTemplate.ColIndex - 1].Value.ToString();
					array2[7] = sheetView.Cells[num2 + j, rowsInfo.Cells[7].CellTemplate.ColIndex - 1].Text;
					bool visible = sheetView.Cells[num2 + j, rowsInfo.Cells[1].CellTemplate.ColIndex - 1].Row.Visible;
					if (GeneralFunctions.HaveToValidateRow(rowsInfo, array2) && visible)
					{
						flag2 = true;
						if (string.IsNullOrEmpty(array2[1]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + j, rowsInfo.Cells[1].CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0017"));
							num = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num2 + j, rowsInfo.Cells[1].CellTemplate.ColIndex - 1);
						}
						if (!string.IsNullOrEmpty(array2[2]) & !ValidateBase.CheckMaSoThue(array2[2]))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + j, rowsInfo.Cells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0047"));
							num = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num2 + j, rowsInfo.Cells[2].CellTemplate.ColIndex - 1);
						}
						if (string.IsNullOrEmpty(array2[5]) || int.Parse(array2[5]) == 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + j, rowsInfo.Cells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0017"));
							num = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num2 + j, rowsInfo.Cells[5].CellTemplate.ColIndex - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num2 + j, rowsInfo.Cells[1].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sheetView, num2 + j, rowsInfo.Cells[2].CellTemplate.ColIndex - 1);
						HTKKFile.RemoveValidate(sheetView, num2 + j, rowsInfo.Cells[5].CellTemplate.ColIndex - 1);
					}
				}
				if (!flag2 && !flag)
				{
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[1].CellTemplate.RowIndex - 1, header2.StaticCells[1].CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[2].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1);
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[5].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1);
					continue;
				}
				if (string.IsNullOrEmpty(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, header2.StaticCells[1].CellTemplate.RowIndex - 1, header2.StaticCells[1].CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0017"));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[1].CellTemplate.RowIndex - 1, header2.StaticCells[1].CellTemplate.ColIndex - 1);
				}
				if (!string.IsNullOrEmpty(array[2]) & !ValidateBase.CheckMaSoThue(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, header2.StaticCells[2].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0047"));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[2].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1);
				}
				if (string.IsNullOrEmpty(array[5]) || Convert.ToDouble(array[5]) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, header2.StaticCells[5].CellTemplate.RowIndex - 1, header2.StaticCells[5].CellTemplate.ColIndex - 1, HTKKFile.GetMessage("0017"));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[5].CellTemplate.RowIndex - 1, header2.StaticCells[5].CellTemplate.ColIndex - 1);
				}
			}
			switch (num)
			{
			case 2:
				winFpSpread.TabStrip[2].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[2].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[2].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckKHBS(string fullNameDataFileToKhaiChinh, string strFileNameExtend, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		//IL_00e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0174: Unknown result type (might be due to invalid IL or missing references)
		//IL_0223: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0292: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string text = fullNameDataFileToKhaiChinh.Substring(fullNameDataFileToKhaiChinh.LastIndexOf('\\') + 1);
			string text2 = text.Substring(0, text.LastIndexOf('_') - 8);
			string text3 = fullNameDataFileToKhaiChinh.Substring(0, fullNameDataFileToKhaiChinh.LastIndexOf('\\'));
			int startIndex = text.LastIndexOf("_") + 1;
			int num = int.Parse(text.Substring(startIndex, 2)) - 1;
			string text4 = string.Concat(num);
			if (num < 10)
			{
				text4 = num.ToString().PadLeft(2, '0');
			}
			DateTime dateTime = DateTime.ParseExact(PeriodInfo.NgayLapKHBS, "d/M/yyyy", null);
			DateTime now = DateTime.Now;
			if ((now - dateTime).TotalDays < 0.0)
			{
				frmMessage frmMessage = new frmMessage("0408", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			string path = text3 + "\\" + text2 + "00000000_" + text4 + ".xml";
			if (text4 == "00")
			{
				if (!File.Exists(path))
				{
					frmMessage frmMessage = new frmMessage("0405", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				DateTime ngayBatDauDuocNopKHBS = GeneralFunctions.GetNgayBatDauDuocNopKHBS(PeriodInfo);
				if ((ngayBatDauDuocNopKHBS != DateTime.MinValue) & (dateTime < ngayBatDauDuocNopKHBS))
				{
					frmMessage frmMessage = new frmMessage("0406", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			else if (text4 != "00")
			{
				if (!GeneralFunctions.CheckKHBSLanTruoc(text2, HeaderInfo, num))
				{
					frmMessage frmMessage = new frmMessage("0404", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				DateTime dateKHBS = GeneralFunctions.getDateKHBS(text2, HeaderInfo, num);
				DateTime dateTime2 = DateTime.ParseExact(PeriodInfo.NgayLapKHBS, "d/M/yyyy", null);
				if ((dateTime2 - dateKHBS).TotalDays <= 0.0)
				{
					frmMessage frmMessage = new frmMessage("0407", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public void TongHopDieuChinhSoThuePhaiNopTang_3(FpSpread winFpSpread, HTKKCMCSheetInfo KHBS_LanTruoc, HTKKCMCSheetInfo sKHBSInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			SheetView sheetView2 = winFpSpread.Sheets[0];
			double? num = null;
			double? num2 = null;
			HTKKCMCCellInfo hTKKCMCCellInfo = KHBS_LanTruoc.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == pathCt31);
			if (hTKKCMCCellInfo == null)
			{
				return;
			}
			num = Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			num2 = Convert.ToDouble(hTKKCMCCellInfo.Value);
			if (num.HasValue)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo2 = sKHBSInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct4");
				HTKKCMCCellInfo hTKKCMCCellInfo3 = sKHBSInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct5");
				HTKKCMCCellInfo hTKKCMCCellInfo4 = sKHBSInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct6");
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].ResetFormula();
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
				sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].ResetFormula();
				sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = num;
				sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].ResetFormula();
				sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = num - num2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		//IL_0062: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("//tieuMucHachToan");
			if ((val != null) & !string.IsNullOrEmpty(PeriodInfo.TieuMucHachToanVal))
			{
				val.InnerText = PeriodInfo.TieuMucHachToanVal;
			}
			XmlNode val2 = xmlNode.SelectSingleNode("//CTieuTKhaiChinh/ct32b");
			if (val2 != null && string.IsNullOrEmpty(val2.InnerText))
			{
				((XmlElement)val2).SetAttribute("nil", "http://www.w3.org/2001/XMLSchema-instance", "true");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public string[] MucLucNganSach_0(HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			if (!string.IsNullOrEmpty(PeriodInfo.TieuMucHachToanVal))
			{
				return new string[2]
				{
					"Hô\u0303 trơ\u0323 ha\u0323ch toa\u0301n MLNSNN:",
					"Hạch toán chỉ tiêu [31] vào mục 1050, tiểu mục " + PeriodInfo.TieuMucHachToanVal
				};
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void XuLyCheckBox_ThueSuatUuDai(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, bool isBindData, bool cbxLoadedUuDai)
	{
		//IL_014b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0150: Unknown result type (might be due to invalid IL or missing references)
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_0153: Invalid comparison between Unknown and I4
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			bool flag = Convert.ToBoolean(sView.Cells[CellId_ckbUuDai[1], CellId_ckbUuDai[0]].Value);
			if (isBindData)
			{
				if (flag)
				{
					sView.Cells[CellId_ct26[1], CellId_ct26[0]].ResetCellType();
					GeneralFunctions.ChangeValueCell(sView, CellId_ct26[1], CellId_ct26[0], locked: false, "", "");
					GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: false, "", 0);
				}
				else
				{
					sView.Cells[CellId_ct26[1], CellId_ct26[0]].CellType = NumberCellType_ct26;
					sView.Cells[CellId_ct26[1], CellId_ct26[0]].Value = 0;
					GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: true, formulaCt29, 0);
				}
			}
			else if (flag)
			{
				DialogResult val = MessageBox.Show("Các chỉ tiêu liên quan đến mức thuế suất ưu đãi [26], [29] sẽ bị xóa bằng 0. Bạn có đồng ý không?", "Thông báo", (MessageBoxButtons)4);
				if ((int)val == 6)
				{
					sView.Cells[CellId_ct26[1], CellId_ct26[0]].ResetCellType();
					GeneralFunctions.ChangeValueCell(sView, CellId_ct26[1], CellId_ct26[0], locked: false, "", "");
					GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: false, "", 0);
				}
				else
				{
					sView.Cells[CellId_ckbUuDai[1], CellId_ckbUuDai[0]].Value = false;
				}
			}
			else
			{
				sView.Cells[CellId_ct26[1], CellId_ct26[0]].CellType = NumberCellType_ct26;
				sView.Cells[CellId_ct26[1], CellId_ct26[0]].Value = 0;
				GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: true, formulaCt29, 0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void XuLyCheckBox_Ct32(SheetView sView)
	{
		try
		{
			if (Convert.ToBoolean(sView.Cells[CellId_ct32[1], CellId_ct32[0]].Value))
			{
				sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].Locked = false;
				sView.Cells[CellId_ct32b[1], CellId_ct32b[0]].Locked = false;
				sView.Cells[CellId_ct32c[1], CellId_ct32c[0]].Locked = false;
				sView.Cells[CellId_ct32d[1], CellId_ct32d[0]].Locked = false;
				return;
			}
			sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].ResetValue();
			sView.Cells[CellIdValue_ct32a[1], CellIdValue_ct32a[0]].ResetValue();
			sView.Cells[CellId_ct32b[1], CellId_ct32b[0]].ResetValue();
			sView.Cells[CellId_ct32c[1], CellId_ct32c[0]].Value = 0;
			sView.Cells[CellId_ct32d[1], CellId_ct32d[0]].Value = 0;
			sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].Locked = true;
			sView.Cells[CellId_ct32b[1], CellId_ct32b[0]].Locked = true;
			sView.Cells[CellId_ct32c[1], CellId_ct32c[0]].Locked = true;
			sView.Cells[CellId_ct32d[1], CellId_ct32d[0]].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
