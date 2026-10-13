using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tndn01ANckt207
{
	private Regex regexCt32a_mulVal = new Regex("^0$|^([1-9][0-9]?)(;[1-9][0-9]?)*$");

	private string path_ct26 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct26";

	private string cellId_ct26 = "K_27";

	private string cellId_ct27 = "K_28";

	private string cellId_ct28 = "K_29";

	private string cellId_ct29 = "K_30";

	private string formulaCt29 = "IF(K27-K28-K29>=0,K27-K28-K29,0)";

	private string pathCt29 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct29";

	private string cellId_ct30 = "K_31";

	private string cellId_ct31 = "K_32";

	private string cellId_ct32 = "K_33";

	private string cellId_ct33 = "K_35";

	private string cellId_ct33a = "K_36";

	private string cellId_uuDai = "F_34";

	private bool loadedValueUuDai;

	private string path_ct32a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct32a";

	private string cellId_ct32a = "K_34";

	private NumberCellType numberCellType_ct32 = new NumberCellType();

	private string cellId_ct34 = "K_37";

	private string pathCt34 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ct34";

	private string formulaCt34 = "ROUND(K31*20/100,0)+ROUND(K32*22/100,0)+ROUND(K33*K34/100,0)+ROUND(K35*K36/100,0)";

	private string cellId_ct35 = "K_38";

	private string cellId_ct35a = "K_39";

	private string cellId_ct35b = "K_40";

	private string cellId_ct37 = "F_43";

	private string cellId_ct37a = "H_45";

	private string cellIdValue_ct37a = "P_45";

	private string cellId_ct37b = "H_47";

	private string cellId_ct37c = "H_49";

	private string cellId_ct37d = "H_51";

	private string cellId_ngayKy = "J_56";

	private int[] CellId_ct26 => GeneralFunctions.ArrayCellId(cellId_ct26);

	private int[] CellId_ct27 => GeneralFunctions.ArrayCellId(cellId_ct27);

	private int[] CellId_ct28 => GeneralFunctions.ArrayCellId(cellId_ct28);

	private int[] CellId_ct29 => GeneralFunctions.ArrayCellId(cellId_ct29);

	private int[] CellId_ct30 => GeneralFunctions.ArrayCellId(cellId_ct30);

	private int[] CellId_ct31 => GeneralFunctions.ArrayCellId(cellId_ct31);

	private int[] CellId_ct32 => GeneralFunctions.ArrayCellId(cellId_ct32);

	private int[] CellId_ct33 => GeneralFunctions.ArrayCellId(cellId_ct33);

	private int[] CellId_ct33a => GeneralFunctions.ArrayCellId(cellId_ct33a);

	private int[] CellId_UuDai => GeneralFunctions.ArrayCellId(cellId_uuDai);

	private int[] CellId_ct32a => GeneralFunctions.ArrayCellId(cellId_ct32a);

	private NumberCellType NumberCellType_ct32
	{
		get
		{
			numberCellType_ct32.DecimalPlaces = 3;
			numberCellType_ct32.DecimalSeparator = ",";
			numberCellType_ct32.MinimumValue = 0.0;
			numberCellType_ct32.MaximumValue = 100.0;
			return numberCellType_ct32;
		}
	}

	private TextCellType TextCellType_ct32a
	{
		get
		{
			TextCellType textCellType = new TextCellType();
			textCellType.WordWrap = true;
			textCellType.MaxLength = 14;
			return textCellType;
		}
	}

	private int[] CellId_ct34 => GeneralFunctions.ArrayCellId(cellId_ct34);

	private int[] CellId_ct35 => GeneralFunctions.ArrayCellId(cellId_ct35);

	private int[] CellId_ct35a => GeneralFunctions.ArrayCellId(cellId_ct35a);

	private int[] CellId_ct35b => GeneralFunctions.ArrayCellId(cellId_ct35b);

	private int[] CellId_ct37 => GeneralFunctions.ArrayCellId(cellId_ct37);

	private int[] CellId_ct37a => GeneralFunctions.ArrayCellId(cellId_ct37a);

	private int[] CellIdValue_ct37a => GeneralFunctions.ArrayCellId(cellIdValue_ct37a);

	private int[] CellId_ct37b => GeneralFunctions.ArrayCellId(cellId_ct37b);

	private int[] CellId_ct37c => GeneralFunctions.ArrayCellId(cellId_ct37c);

	private int[] CellId_ct37d => GeneralFunctions.ArrayCellId(cellId_ct37d);

	private int[] CellId_ngayKy => GeneralFunctions.ArrayCellId(cellId_ngayKy);

	private int[] CellId_E22 => GeneralFunctions.ArrayCellId("E_22");

	private int[] CellId_E33 => GeneralFunctions.ArrayCellId("E_33");

	private int[] CellId_E34 => GeneralFunctions.ArrayCellId("E_34");

	private int[] CellId_E40 => GeneralFunctions.ArrayCellId("E_40");

	private int[] CellId_E41 => GeneralFunctions.ArrayCellId("E_41");

	private int[] CellId_E43 => GeneralFunctions.ArrayCellId("E_43");

	private int[] CellId_E44 => GeneralFunctions.ArrayCellId("E_44");

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			loadedValueUuDai = Convert.ToBoolean(sView.Cells[CellId_UuDai[1], CellId_UuDai[0]].Value);
			XuLyCheckBox_ThueSuatUuDai(sView, ToKhaiInfo, isBindData: true, loadedValueUuDai);
			XuLyCheckBox_Ct37(sView);
			XuLyChung_ToKhaiChinh(sView, ToKhaiInfo, isBindData: true);
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
				if (idxRow == CellId_ct32a[1] && idxCol == CellId_ct32a[0] && Convert.ToBoolean(sheetView.Cells[CellId_UuDai[1], CellId_UuDai[0]].Value))
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
			if (winFpSpread.ActiveSheetIndex == 0)
			{
				XuLyChung_ToKhaiChinh(winFpSpread.Sheets[0], ToKhaiInfo, isBindData: false);
			}
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
			if (activeColumnIndex == CellId_UuDai[0] && activeRowIndex == CellId_UuDai[1])
			{
				XuLyCheckBox_ThueSuatUuDai(winFpSpread.ActiveSheet, ToKhaiInfo, isBindData: false, loadedValueUuDai);
			}
			if (activeColumnIndex == CellId_ct37[0] && activeRowIndex == CellId_ct37[1])
			{
				XuLyCheckBox_Ct37(winFpSpread.ActiveSheet);
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
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36"));
			string text = sheetView.Cells[CellId_ct32a[1], CellId_ct32a[0]].Text;
			if (Convert.ToBoolean(sheetView.Cells[CellId_UuDai[1], CellId_UuDai[0]].Value))
			{
				if (!regexCt32a_mulVal.IsMatch(text))
				{
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct32a[1], CellId_ct32a[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct32a[1], CellId_ct32a[0]);
			}
			double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct35[1], CellId_ct35[0]].Value);
			double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct35a[1], CellId_ct35a[0]].Value);
			double num4 = ConvertHelpers.ConvertToDouble(sheetView.Cells[CellId_ct35b[1], CellId_ct35b[0]].Value);
			if (num2 < num3 + num4)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct35[1], CellId_ct35[0], message);
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct35[1], CellId_ct35[0]);
			}
			if (Convert.ToBoolean(sheetView.Cells[CellId_ct37[1], CellId_ct37[0]].Value))
			{
				string text2 = sheetView.Cells[CellIdValue_ct37a[1], CellIdValue_ct37a[0]].Value.ToString();
				if (string.IsNullOrEmpty(text2) || int.Parse(text2) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37a[1], CellId_ct37a[0], HTKKFile.GetMessage(winFpSpread, CellId_E33[1], CellId_E33[0]));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct37a[1], CellId_ct37a[0]);
				}
				string text3 = sheetView.Cells[CellId_ct37b[1], CellId_ct37b[0]].Text;
				DateTime result;
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37b[1], CellId_ct37b[0], HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]));
					num = 2;
				}
				else if (!DateTime.TryParseExact(text3, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37b[1], CellId_ct37b[0], "Ngày không hợp lệ");
					num = 2;
				}
				else if (result <= GeneralFunctions.GetHanNopToKhai(PeriodInfo))
				{
					HTKKFile.SetNodeBackGroundYellow(sheetView, CellId_ct37b[1], CellId_ct37b[0], HTKKFile.GetMessage(winFpSpread, CellId_E44[1], CellId_E44[0]));
					num = 1;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct37b[1], CellId_ct37b[0]);
				}
				object value = sheetView.Cells[CellId_ct37c[1], CellId_ct37c[0]].Value;
				if (value == null || (double)value == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37c[1], CellId_ct37c[0], HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct37c[1], CellId_ct37c[0]);
				}
				object value2 = sheetView.Cells[CellId_ct37d[1], CellId_ct37d[0]].Value;
				if (value2 == null || Convert.ToDouble(value2) == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_ct37d[1], CellId_ct37d[0], HTKKFile.GetMessage(winFpSpread, CellId_E22[1], CellId_E22[0]));
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, CellId_ct37d[1], CellId_ct37d[0]);
				}
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, CellId_ct37a[1], CellId_ct37a[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct37b[1], CellId_ct37b[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct37c[1], CellId_ct37c[0]);
				HTKKFile.RemoveValidate(sheetView, CellId_ct37d[1], CellId_ct37d[0]);
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
			string message = HTKKFile.GetMessage("0047");
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			int column = rowsInfo.Cells[0].CellTemplate.ColIndex - 1;
			int num3 = rowsInfo.Cells[1].CellTemplate.ColIndex - 1;
			int num4 = rowsInfo.Cells[2].CellTemplate.ColIndex - 1;
			int num5 = rowsInfo.Cells[3].CellTemplate.ColIndex - 1;
			int num6 = rowsInfo.Cells[4].CellTemplate.ColIndex - 1;
			int num7 = rowsInfo.Cells[5].CellTemplate.ColIndex - 1;
			string[] array = new string[8];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array[0] = sheetView.Cells[num2 + i, column].Text;
				array[1] = sheetView.Cells[num2 + i, num3].Text;
				array[2] = sheetView.Cells[num2 + i, num4].Text;
				array[3] = sheetView.Cells[num2 + i, num5].Text;
				array[4] = sheetView.Cells[num2 + i, num6].Text;
				array[5] = sheetView.Cells[num2 + i, num7].Text;
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array))
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, num3);
					HTKKFile.RemoveValidate(sheetView, num2 + i, num4);
					HTKKFile.RemoveValidate(sheetView, num2 + i, num5);
					HTKKFile.RemoveValidate(sheetView, num2 + i, num6);
					HTKKFile.RemoveValidate(sheetView, num2 + i, num7);
				}
				else if (!string.IsNullOrEmpty(array[1]) & !ValidateBase.CheckMaSoThue(array[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, num3, message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num2 + i, num3);
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_1_TNDN/BangThueTNDNTamTinh/tongCot12");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == "PL01_1_TNDN/ct6");
			int column2 = hTKKCMCCellInfo2.CellTemplate.ColIndex - 1;
			int row = hTKKCMCCellInfo2.CellTemplate.RowIndex - 1;
			int num8 = hTKKCMCCellInfo.CellTemplate.ColIndex - 1;
			int num9 = hTKKCMCCellInfo.CellTemplate.RowIndex - 1;
			double num10 = Convert.ToDouble(sheetView.Cells[row, column2].Value);
			double num11 = Convert.ToDouble(sheetView.Cells[num9, num8].Value);
			if (num11 != num10)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, num9, num8, HTKKFile.GetMessage(sheetView.FpSpread, CellId_E34[1], CellId_E34[0]));
				num = 2;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, num9, num8);
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
			string message = HTKKFile.GetMessage(sheetView, CellId_E22);
			string message2 = HTKKFile.GetMessage("0047");
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			HTKKCMCStaticOfTable header = tables[0].Header;
			string[] array = new string[7];
			string[] array2 = new string[7];
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
					array2[6] = sheetView.Cells[num2 + j, rowsInfo.Cells[6].CellTemplate.ColIndex - 1].Text;
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
						if (string.IsNullOrEmpty(array2[5]) || Convert.ToDouble(array2[5]) == 0.0)
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
					HTKKFile.SetNodeBackGroundPink(sheetView, header2.StaticCells[1].CellTemplate.RowIndex - 1, header2.StaticCells[1].CellTemplate.ColIndex - 1, message);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[1].CellTemplate.RowIndex - 1, header2.StaticCells[1].CellTemplate.ColIndex - 1);
				}
				if (!string.IsNullOrEmpty(array[2]) & !ValidateBase.CheckMaSoThue(array[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, header2.StaticCells[2].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1, message2);
					num = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, header2.StaticCells[2].CellTemplate.RowIndex - 1, header2.StaticCells[2].CellTemplate.ColIndex - 1);
				}
				if (string.IsNullOrEmpty(array[5]) || int.Parse(array[5]) == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, header2.StaticCells[5].CellTemplate.RowIndex - 1, header2.StaticCells[5].CellTemplate.ColIndex - 1, message);
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

	private void XuLyCheckBox_ThueSuatUuDai(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, bool isBindData, bool cbxLoadedUuDai)
	{
		//IL_01f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0202: Invalid comparison between Unknown and I4
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt29));
			double num = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value);
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt34));
			double num2 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo2.Value);
			bool flag = Convert.ToBoolean(sView.Cells[CellId_UuDai[1], CellId_UuDai[0]].Value);
			if (isBindData)
			{
				if (flag)
				{
					sView.Cells[CellId_ct29[1], CellId_ct29[0]].Locked = false;
					sView.Cells[CellId_ct29[1], CellId_ct29[0]].ResetFormula();
					sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].CellType = TextCellType_ct32a;
					sView.Cells[CellId_ct34[1], CellId_ct34[0]].Locked = false;
					sView.Cells[CellId_ct34[1], CellId_ct34[0]].ResetFormula();
				}
				else
				{
					GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: true, formulaCt29, 0);
					sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].CellType = NumberCellType_ct32;
					GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: true, formulaCt34, 0);
				}
			}
			else if (flag)
			{
				DialogResult val = MessageBox.Show("Các chỉ tiêu liên quan đến mức thuế suất ưu đãi [32a], [34] sẽ bị xóa bằng 0. Bạn có đồng ý không?", "Thông báo", (MessageBoxButtons)4);
				if ((int)val == 6)
				{
					GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: false, "", 0);
					sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].CellType = TextCellType_ct32a;
					GeneralFunctions.ChangeValueCell(sView, CellId_ct32a[1], CellId_ct32a[0], locked: false, "", "");
					GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: false, "", 0);
				}
				else
				{
					sView.Cells[CellId_UuDai[1], CellId_UuDai[0]].Value = false;
				}
			}
			else
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: true, formulaCt29, 0);
				sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].CellType = NumberCellType_ct32;
				sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].Value = 0;
				GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: true, formulaCt34, 0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void XuLyCheckBox_Ct37(SheetView sView)
	{
		try
		{
			if (Convert.ToBoolean(sView.Cells[CellId_ct37[1], CellId_ct37[0]].Value))
			{
				sView.Cells[CellId_ct37a[1], CellId_ct37a[0]].Locked = false;
				sView.Cells[CellId_ct37b[1], CellId_ct37b[0]].Locked = false;
				sView.Cells[CellId_ct37c[1], CellId_ct37c[0]].Locked = false;
				sView.Cells[CellId_ct37d[1], CellId_ct37d[0]].Locked = false;
				return;
			}
			sView.Cells[CellId_ct37a[1], CellId_ct37a[0]].ResetValue();
			sView.Cells[CellIdValue_ct37a[1], CellIdValue_ct37a[0]].ResetValue();
			sView.Cells[CellId_ct37b[1], CellId_ct37b[0]].ResetValue();
			sView.Cells[CellId_ct37c[1], CellId_ct37c[0]].Value = 0;
			sView.Cells[CellId_ct37d[1], CellId_ct37d[0]].Value = 0;
			sView.Cells[CellId_ct37a[1], CellId_ct37a[0]].Locked = true;
			sView.Cells[CellId_ct37b[1], CellId_ct37b[0]].Locked = true;
			sView.Cells[CellId_ct37c[1], CellId_ct37c[0]].Locked = true;
			sView.Cells[CellId_ct37d[1], CellId_ct37d[0]].Locked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void XuLyChung_ToKhaiChinh(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, bool isBindData)
	{
		try
		{
			double num = 0.0;
			if (isBindData)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path == path_ct26);
				num = Convert.ToDouble(hTKKCMCCellInfo.Value);
			}
			else
			{
				num = Convert.ToDouble(sView.Cells[CellId_ct26[1], CellId_ct26[0]].Value);
			}
			double num2 = Convert.ToDouble(sView.Cells[CellId_ct27[1], CellId_ct27[0]].Value);
			double num3 = Convert.ToDouble(sView.Cells[CellId_ct29[1], CellId_ct29[0]].Value);
			bool flag = Convert.ToBoolean(sView.Cells[CellId_UuDai[1], CellId_UuDai[0]].Value);
			if (num - num2 < 0.0)
			{
				sView.Cells[CellId_ct28[1], CellId_ct28[0]].Value = 0;
				sView.Cells[CellId_ct28[1], CellId_ct28[0]].Locked = true;
			}
			else
			{
				sView.Cells[CellId_ct28[1], CellId_ct28[0]].Locked = false;
			}
			if (num < 0.0)
			{
				GeneralFunctions.ChangeValueCell(sView, CellId_ct28[1], CellId_ct28[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct29[1], CellId_ct29[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct30[1], CellId_ct30[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct31[1], CellId_ct31[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct32[1], CellId_ct32[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct32a[1], CellId_ct32a[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct33[1], CellId_ct33[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct33a[1], CellId_ct33a[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct35[1], CellId_ct35[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct35a[1], CellId_ct35a[0], locked: true, "", 0);
				GeneralFunctions.ChangeValueCell(sView, CellId_ct35b[1], CellId_ct35b[0], locked: true, "", 0);
			}
			else
			{
				if (sView.Cells[CellId_ct28[1], CellId_ct28[0]].Locked && num - num2 > 0.0)
				{
					sView.Cells[CellId_ct28[1], CellId_ct28[0]].Locked = false;
				}
				if (string.IsNullOrEmpty(sView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula) && !flag)
				{
					sView.Cells[CellId_ct29[1], CellId_ct29[0]].Formula = formulaCt29;
				}
				if (sView.Cells[CellId_ct30[1], CellId_ct30[0]].Locked)
				{
					sView.Cells[CellId_ct30[1], CellId_ct30[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct31[1], CellId_ct31[0]].Locked)
				{
					sView.Cells[CellId_ct31[1], CellId_ct31[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct32[1], CellId_ct32[0]].Locked)
				{
					sView.Cells[CellId_ct32[1], CellId_ct32[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].Locked)
				{
					sView.Cells[CellId_ct32a[1], CellId_ct32a[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct33[1], CellId_ct33[0]].Locked)
				{
					sView.Cells[CellId_ct33[1], CellId_ct33[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct33a[1], CellId_ct33a[0]].Locked)
				{
					sView.Cells[CellId_ct33a[1], CellId_ct33a[0]].Locked = false;
				}
				if (string.IsNullOrEmpty(sView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula) && !flag)
				{
					sView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formulaCt34;
				}
				if (sView.Cells[CellId_ct35[1], CellId_ct35[0]].Locked)
				{
					sView.Cells[CellId_ct35[1], CellId_ct35[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct35a[1], CellId_ct35a[0]].Locked)
				{
					sView.Cells[CellId_ct35a[1], CellId_ct35a[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct35b[1], CellId_ct35b[0]].Locked)
				{
					sView.Cells[CellId_ct35b[1], CellId_ct35b[0]].Locked = false;
				}
			}
			if (num3 < 0.0)
			{
				if (num > 0.0)
				{
					GeneralFunctions.ChangeValueCell(sView, CellId_ct30[1], CellId_ct30[0], locked: true, "", 0);
					GeneralFunctions.ChangeValueCell(sView, CellId_ct31[1], CellId_ct31[0], locked: true, "", 0);
					GeneralFunctions.ChangeValueCell(sView, CellId_ct32[1], CellId_ct32[0], locked: true, "", 0);
					GeneralFunctions.ChangeValueCell(sView, CellId_ct34[1], CellId_ct34[0], locked: true, "", 0);
					GeneralFunctions.ChangeValueCell(sView, CellId_ct35[1], CellId_ct35[0], locked: true, "", 0);
					GeneralFunctions.ChangeValueCell(sView, CellId_ct35a[1], CellId_ct35a[0], locked: true, "", 0);
					GeneralFunctions.ChangeValueCell(sView, CellId_ct35b[1], CellId_ct35b[0], locked: true, "", 0);
				}
			}
			else if (num > 0.0)
			{
				if (sView.Cells[CellId_ct30[1], CellId_ct30[0]].Locked)
				{
					sView.Cells[CellId_ct30[1], CellId_ct30[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct31[1], CellId_ct31[0]].Locked)
				{
					sView.Cells[CellId_ct31[1], CellId_ct31[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct32[1], CellId_ct32[0]].Locked)
				{
					sView.Cells[CellId_ct32[1], CellId_ct32[0]].Locked = false;
				}
				if (string.IsNullOrEmpty(sView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula) && !flag)
				{
					sView.Cells[CellId_ct34[1], CellId_ct34[0]].Formula = formulaCt34;
				}
				if (sView.Cells[CellId_ct35[1], CellId_ct35[0]].Locked)
				{
					sView.Cells[CellId_ct35[1], CellId_ct35[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct35a[1], CellId_ct35a[0]].Locked)
				{
					sView.Cells[CellId_ct35a[1], CellId_ct35a[0]].Locked = false;
				}
				if (sView.Cells[CellId_ct35b[1], CellId_ct35b[0]].Locked)
				{
					sView.Cells[CellId_ct35b[1], CellId_ct35b[0]].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
