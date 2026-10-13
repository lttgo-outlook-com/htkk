using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Cnkd01Nckt216
{
	private Dictionary<string, string[]> dicDMKhoangSan = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMDonVi = new Dictionary<string, string>();

	private Dictionary<string, string> dicDMKhoangSanTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMThueBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueBVMTTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMNganhNgheKD = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMNganhNgheKDTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMNhomNganhNgheKD = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMNhomNganhNgheKDTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMThueDB = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMThueDBTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMKSBVMT = new Dictionary<string, string[]>();

	private Dictionary<string, string> dicDMKSBVMTTmp = new Dictionary<string, string>();

	private List<item> _dmKSData = new List<item>();

	private List<item> _dmBVMTData = new List<item>();

	private bool _isDmMoi = false;

	private string namKK = "";

	private string kieuKyKK = "";

	private List<string> pathDonGiaTinh = new List<string>();

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

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			namKK = PeriodInfo.Nam;
			kieuKyKK = PeriodInfo.KieuKy;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AA");
			int[] array = GeneralFunctions.ArrayCellId("D_22");
			int[] array2 = GeneralFunctions.ArrayCellId("D_24");
			int[] array3 = GeneralFunctions.ArrayCellId("D_26");
			int[] array4 = GeneralFunctions.ArrayCellId("D_30");
			int[] array5 = GeneralFunctions.ArrayCellId("D_32");
			int[] array6 = GeneralFunctions.ArrayCellId("D_34");
			int[] array7 = GeneralFunctions.ArrayCellId("J_36");
			int[] array8 = GeneralFunctions.ArrayCellId("F_40");
			int[] array9 = GeneralFunctions.ArrayCellId("F_43");
			int[] array10 = GeneralFunctions.ArrayCellId("T_43");
			int[] array11 = GeneralFunctions.ArrayCellId("J_45");
			int[] array12 = GeneralFunctions.ArrayCellId("H_47");
			int[] array13 = GeneralFunctions.ArrayCellId("J_47");
			int[] array14 = GeneralFunctions.ArrayCellId("N_47");
			int[] array15 = GeneralFunctions.ArrayCellId("P_47");
			int[] array16 = GeneralFunctions.ArrayCellId("E_51");
			int[] array17 = GeneralFunctions.ArrayCellId("E_53");
			int[] array18 = GeneralFunctions.ArrayCellId("E_55");
			int[] array19 = GeneralFunctions.ArrayCellId("E_57");
			int[] array20 = GeneralFunctions.ArrayCellId("J_59");
			int[] array21 = GeneralFunctions.ArrayCellId("W_59");
			if (PeriodInfo.ToChucKhaiThay == "1")
			{
				sView.Cells[10, num - 1].Text = PeriodInfo.ToChucKhaiThay;
			}
			if (!PeriodInfo.KieuKy.Equals("Y"))
			{
				sView.Rows[12].Visible = false;
			}
			string nam = PeriodInfo.Nam;
			string thang = PeriodInfo.Thang;
			string quy = PeriodInfo.Quy;
			string kieuKy = PeriodInfo.KieuKy;
			int column = GeneralFunctions.ArrayCellId("U_105")[0];
			int row = GeneralFunctions.ArrayCellId("U_105")[1];
			int column2 = GeneralFunctions.ArrayCellId("U_106")[0];
			int row2 = GeneralFunctions.ArrayCellId("U_106")[1];
			if (kieuKy.Equals("Y"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018)
				{
					sView.Cells[row, column].Value = "0";
					sView.Cells[row, column].Locked = true;
					sView.Cells[row2, column2].Value = "0";
				}
			}
			else if (kieuKy.Equals("Q"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(quy) == 4))
				{
					sView.Cells[row, column].Value = "0";
					sView.Cells[row, column].Locked = true;
					sView.Cells[row2, column2].Value = "0";
				}
			}
			else if ((kieuKy.Equals("M") || kieuKy.Equals("D")) && (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(thang) >= 8)))
			{
				sView.Cells[row, column].Value = "0";
				sView.Cells[row, column].Locked = true;
				sView.Cells[row2, column2].Value = "0";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetFormula(SheetView sView)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ResetFormula(SheetView sViewTK)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("Q");
			int num5 = HTKKFile.ColumnLettersToIndex("AC");
			int num6 = HTKKFile.ColumnLettersToIndex("AB");
			int num7 = HTKKFile.ColumnLettersToIndex("K");
			int num8 = HTKKFile.ColumnLettersToIndex("M");
			int num9 = HTKKFile.ColumnLettersToIndex("U");
			int num10 = HTKKFile.ColumnLettersToIndex("H");
			int num11 = HTKKFile.ColumnLettersToIndex("J");
			int num12 = HTKKFile.ColumnLettersToIndex("W");
			int num13 = HTKKFile.ColumnLettersToIndex("AD");
			if (iCol == num)
			{
				string[] array = dicDMNganhNgheKD[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 1)
				{
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = array[0];
				}
			}
			if (iCol == num2)
			{
				if (sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim() == "matinh")
				{
					string text = HTKKWinDictionaryDiaDanh.dicTinhTP[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
					sheetView.Cells[(int)iRow - 1, num5 - 1].Text = text;
				}
				if (sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim() == "mahuyen")
				{
					foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.dicQuanHuyen)
					{
						if (item.Value == sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim())
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Text = item.Key;
						}
					}
				}
				if (sheetView.Cells[(int)iRow - 1, num6 - 1].Text.Trim() == "maxa")
				{
					foreach (KeyValuePair<string, string> item2 in HTKKWinDictionaryDiaDanh.dicPhuongXa)
					{
						if (item2.Value == sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim())
						{
							sheetView.Cells[(int)iRow - 1, num5 - 1].Text = item2.Key;
						}
					}
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "ttdb")
			{
				string[] array = dicDMThueDB[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 3)
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = array[1];
					sheetView.Cells[(int)iRow - 1, num8 - 1].Value = array[3];
					sheetView.Cells[(int)iRow - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(array[4]);
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = array[1];
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "tain1")
			{
				winFpSpread.TabStrip[0].BackColor = Color.White;
				HTKKFile.RemoveValidate(sheetView, (int)iRow - 1, num3 - 1);
				string text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				foreach (KeyValuePair<string, string[]> item3 in dicDMKhoangSan)
				{
					string[] value = item3.Value;
					if (value[2].Trim().Equals(text2) && value.Length > 3 && value.Length > 2)
					{
						if (ConvertHelpers.ConvertToInt(value[4]) == 0)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, (int)iRow - 1, num3 - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
							winFpSpread.TabStrip[0].BackColor = Color.Pink;
						}
						sheetView.Cells[(int)iRow - 1, num10 - 1].Value = value[1];
						sheetView.Cells[(int)iRow - 1, num11 - 1].Value = value[3];
						sheetView.Cells[(int)iRow - 1, num9 - 1].Value = ConvertHelpers.ConvertToDouble(value[4]);
						sheetView.Cells[(int)iRow - 1, num6 - 1].Value = value[0];
						sheetView.Cells[(int)iRow - 1, num13 - 1].Value = value[1];
						double donGia = getDonGia(item3.Key);
						sheetView.Cells[(int)iRow - 1, num4 - 1].Value = donGia;
					}
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "bvmt")
			{
				string[] array = dicDMThueBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 1)
				{
					sheetView.Cells[(int)iRow - 1, num10 - 1].Value = array[0];
					sheetView.Cells[(int)iRow - 1, num11 - 1].Value = array[2];
					sheetView.Cells[(int)iRow - 1, num4 - 1].Text = array[3];
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = array[0];
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = array[0];
				}
			}
			if (iCol == num3 && sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "tain2")
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[4];
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = rowsInfo.Cells.Find((HTKKCMCTableCellInfo obj) => obj.Path.Equals("ct2_ten"));
				bool flag = false;
				if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count() > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
				{
					string text3 = PeriodInfo.dicVar["IS_DM_MOI"].ToString();
					if (text3 == "1")
					{
						flag = true;
					}
				}
				string[] array = ((!flag) ? dicDMKhoangSan[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()] : dicDMKSBVMT[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()]);
				if (array.Length > 2)
				{
					sheetView.Cells[(int)iRow - 1, num10 - 1].Value = array[1];
					sheetView.Cells[(int)iRow - 1, num11 - 1].Value = array[3];
					sheetView.Cells[(int)iRow - 1, num6 - 1].Value = array[1];
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = array[1];
				}
			}
			if (iCol == num3 && iRow >= ToKhaiInfo.SheetInfos[0].Tables[2].Location.DataLocationTemplateRowStart && iRow <= ToKhaiInfo.SheetInfos[0].Tables[4].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[4].RowsValues.Count)
			{
				if (HTKKFormater.isNullString(sheetView.Cells[(int)iRow - 1, num3 - 1].Text))
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = false;
				}
			}
			if (iCol == num12)
			{
				object value2 = sheetView.Cells[(int)iRow - 1, num12 - 1].Value;
				sheetView.Cells[(int)iRow - 1, num12 - 1].Formula = "";
				sheetView.Cells[(int)iRow - 1, num12 - 1].Value = value2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellEditchange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("U");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("H");
			int num6 = HTKKFile.ColumnLettersToIndex("K");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("O");
			int num10 = HTKKFile.ColumnLettersToIndex("P");
			int num11 = HTKKFile.ColumnLettersToIndex("R");
			int num12 = HTKKFile.ColumnLettersToIndex("T");
			int num13 = HTKKFile.ColumnLettersToIndex("U");
			int num14 = HTKKFile.ColumnLettersToIndex("V");
			if (iCol == num && iRow == 18)
			{
				string[] array = dicDMNhomNganhNgheKD[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
				if (array.Length > 2)
				{
					sheetView.Cells[(int)iRow - 1, num2 - 1].Text = array[0];
					sheetView.Cells[(int)iRow + 1 - 1, num2 - 1].Text = array[2];
					sheetView.Cells[(int)iRow + 2 - 1, num2 - 1].Text = array[3];
					HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
					for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
					{
						double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Value);
						sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Value = ConvertHelpers.ConvertToDouble(array[2]) / 100.0 * num15;
						sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Value = Math.Round(ConvertHelpers.ConvertToDouble(array[3]) / 100.0 * num15, 0);
					}
				}
			}
			string nam = PeriodInfo.Nam;
			string thang = PeriodInfo.Thang;
			string quy = PeriodInfo.Quy;
			string kieuKy = PeriodInfo.KieuKy;
			if (!(PeriodInfo.KieuKy == "Y") || iCol != num5)
			{
				return;
			}
			string valueCustom = GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num5, "NumberCellType");
			string text = "";
			double num16 = (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
			if (ConvertHelpers.ConvertToDouble(valueCustom) * num16 < 100000000.0)
			{
				sheetView.Cells[(int)iRow - 1, num6 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = true;
				sheetView.Cells[(int)iRow - 1, num7 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = true;
				sheetView.Cells[(int)iRow - 1, num8 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = true;
				sheetView.Cells[(int)iRow - 1, num9 - 1].Value = 0;
				sheetView.Cells[(int)iRow - 1, num9 - 1].Locked = true;
				return;
			}
			sheetView.Cells[(int)iRow - 1, num6 - 1].Locked = false;
			sheetView.Cells[(int)iRow - 1, num8 - 1].Locked = false;
			sheetView.Cells[(int)iRow - 1, num9 - 1].Locked = false;
			if (kieuKy.Equals("Y"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018)
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "0";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = false;
				}
			}
			else if (kieuKy.Equals("Q"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(quy) == 4))
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "0";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = false;
				}
			}
			else if (kieuKy.Equals("M") || kieuKy.Equals("D"))
			{
				if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(thang) >= 8))
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Value = "0";
					sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num7 - 1].Locked = false;
				}
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
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("AC");
			int num5 = HTKKFile.ColumnLettersToIndex("AB");
			int num6 = HTKKFile.ColumnLettersToIndex("K");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("U");
			int num9 = HTKKFile.ColumnLettersToIndex("H");
			int num10 = HTKKFile.ColumnLettersToIndex("J");
			int num11 = HTKKFile.ColumnLettersToIndex("Q");
			int num12 = HTKKFile.ColumnLettersToIndex("W");
			if (iCol == num2)
			{
				if (sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "matinh")
				{
					string text = HTKKWinDictionaryDiaDanh.dicTinhTP[sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()];
					text = text.Substring(0, 3);
					string[] array = new string[1];
					int num13 = 1;
					foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.dicQuanHuyen)
					{
						if (item.Key.IndexOf(text) == 0)
						{
							string[] array3;
							string[] array2 = (array3 = array);
							int num14 = num13 - 1;
							nint num15 = num14;
							array2[num14] = array3[num15] + item.Value;
							num13++;
							Array.Resize(ref array, num13);
						}
					}
					ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
					comboBoxCellType.Items = array;
					sheetView.Cells[(int)iRow + 2 - 1, num2 - 1].CellType = comboBoxCellType;
					sheetView.Cells[(int)iRow + 4 - 1, num2 - 1].Text = "";
					sheetView.Cells[(int)iRow + 4 - 1, num4 - 1].Text = "";
				}
				if (sheetView.Cells[(int)iRow - 1, num5 - 1].Text.Trim() == "mahuyen")
				{
					foreach (KeyValuePair<string, string> item2 in HTKKWinDictionaryDiaDanh.dicQuanHuyen)
					{
						if (!(item2.Value == sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text.Trim()))
						{
							continue;
						}
						string value = item2.Key.Substring(0, 5);
						string[] array = new string[1];
						int num13 = 1;
						foreach (KeyValuePair<string, string> item3 in HTKKWinDictionaryDiaDanh.dicPhuongXa)
						{
							if (item3.Key.IndexOf(value) == 0)
							{
								string[] array3;
								string[] array4 = (array3 = array);
								int num16 = num13 - 1;
								nint num15 = num16;
								array4[num16] = array3[num15] + item3.Value;
								num13++;
								Array.Resize(ref array, num13);
							}
						}
						ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
						comboBoxCellType.Items = array;
						sheetView.Cells[(int)iRow + 2 - 1, num2 - 1].CellType = comboBoxCellType;
					}
				}
			}
			string text2 = "ROUND(M###*Q###,0)";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[4];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
			{
				if (iCol == num12)
				{
					object value2 = sheetView.Cells[(int)iRow - 1, num12 - 1].Value;
					sheetView.Cells[(int)iRow - 1, num12 - 1].Formula = "";
					sheetView.Cells[(int)iRow - 1, num12 - 1].Value = value2;
				}
				else if (iCol == num7 || iCol == num11)
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Formula = text2.Replace("###", iRow.ToString());
				}
			}
			if (iCol == num11 && sheetView.Cells[(int)iRow - 1, num4 - 1].Text.Trim() == "tain1")
			{
				validateTaiNguyen(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("E");
			int num3 = HTKKFile.ColumnLettersToIndex("C");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("AC");
			int num6 = HTKKFile.ColumnLettersToIndex("AB");
			int num7 = HTKKFile.ColumnLettersToIndex("K");
			int num8 = HTKKFile.ColumnLettersToIndex("M");
			int num9 = HTKKFile.ColumnLettersToIndex("U");
			int num10 = HTKKFile.ColumnLettersToIndex("H");
			int num11 = HTKKFile.ColumnLettersToIndex("J");
			int num12 = HTKKFile.ColumnLettersToIndex("N");
			int num13 = HTKKFile.ColumnLettersToIndex("O");
			if (iCol == num7 || (iCol == num4 && iRow == 18))
			{
				string pathTong11 = "PLuc_01_1_BK_CNKD/ct19";
				string path30a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct30";
				string path30c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct30";
				string path31a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct31";
				string path31c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct31";
				string path32a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct32";
				string path32c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct32";
				string path33a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct33";
				string path33c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct33";
				string path34b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/ThueGTGTPhatSinh/ct34";
				string path34d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/ThueTNCNPhatSinh/ct34";
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTong11));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path30a));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path30c));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path31a));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path31c));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path32a));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path32c));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path33a));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path33c));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path34b));
				HTKKCMCCellInfo hTKKCMCCellInfo11 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path34d));
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Formula = "SUM(P92:P99)";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Formula = "SUM(X92:X99)";
				double num14 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[1].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
				switch (winFpSpread.Sheets[1].Cells[17, num9 - 1].Text)
				{
				case "01":
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = 0;
					break;
				case "02":
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = 0;
					break;
				case "03":
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = 0;
					break;
				case "04":
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = 0;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = num14;
					sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = num14;
					break;
				}
				if ((int)iRow > 20)
				{
					int[] array = GeneralFunctions.ArrayCellId("U_19");
					int[] array2 = GeneralFunctions.ArrayCellId("U_20");
					double num15 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array[1], array[0]].Value);
					double num16 = ConvertHelpers.ConvertToDouble(sheetView.Cells[array2[1], array2[0]].Value);
					double num17 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, (int)iRow, num7, "NumberCellType"));
					sheetView.Cells[(int)iRow - 1, num12 - 1].Value = Math.Round(num17 * (num15 / 100.0), 0);
					sheetView.Cells[(int)iRow - 1, num13 - 1].Value = Math.Round(num17 * (num16 / 100.0), 0);
				}
			}
			if (iCol == num8)
			{
				string pathTong12 = "PLuc_01_1_BK_CNKD/ct20";
				string pathTK35 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/XDinhNVuThue/ct35";
				HTKKCMCCellInfo hTKKCMCCellInfo12 = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTong12));
				HTKKCMCCellInfo hTKKCMCCellInfo13 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTK35));
				double num18 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[1].Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Value);
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo13.CellTemplate.RowIndex - 1, hTKKCMCCellInfo13.CellTemplate.ColIndex - 1].Value = num18;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("F");
			int num3 = HTKKFile.ColumnLettersToIndex("G");
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			int num5 = HTKKFile.ColumnLettersToIndex("M");
			if (iCol == num5)
			{
				if (HTKKFormater.isNullString(sheetView.Cells[(int)iRow - 1, num - 1].Text))
				{
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Locked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		//IL_0137: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0176: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = (DateTime.Now.Year - 1) * 12 + DateTime.Now.Month;
			int num2 = (DateTime.Now.Year - 1) * 4 + DateTime.Now.Month / 3;
			int num3 = int.Parse(PeriodInfo.Nam);
			if (PeriodInfo.KieuKy == "M")
			{
				int num4 = (num3 - 1) * 12 + int.Parse(PeriodInfo.Thang);
				if (num4 <= 0)
				{
					return false;
				}
				if (num4 > num)
				{
					frmMessage frmMessage = new frmMessage("0062", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				int num5 = (num3 - 1) * 4 + int.Parse(PeriodInfo.Quy);
				if (num5 > num2)
				{
					frmMessage frmMessage = new frmMessage("0045", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				int num6 = 8059;
				if (num5 < num6)
				{
					frmMessage frmMessage = new frmMessage("01XSBHDC0007", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
		return true;
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(0, out value);
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			SheetView sheetView2 = winFpSpread.Sheets[1];
			double num = 0.0;
			double value2 = 0.0;
			double num2 = 0.0;
			double num3 = 0.0;
			double num4 = 0.0;
			double num5 = 0.0;
			double num6 = 0.0;
			double num7 = 0.0;
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			Cell cell = null;
			Cell cell2 = null;
			Cell cell3 = null;
			Cell cell4 = null;
			List<HTKKCMCCellInfo> staticCells = ToKhaiInfo.SheetInfos[0].StaticCells;
			for (int i = 0; i < staticCells.Count; i++)
			{
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/kyDieuChinh")
				{
					string text7 = sheetView.Cells[staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1].Text;
					DateTime result = default(DateTime);
					if (!string.IsNullOrEmpty(text7))
					{
						DateTime.TryParseExact("01/" + text7, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
						if (result > DateTime.MinValue)
						{
							if (result < DateTime.ParseExact(PeriodInfo.TuThang, "MM/yyyy", CultureInfo.CurrentCulture) || result > DateTime.ParseExact(PeriodInfo.DenThang, "MM/yyyy", CultureInfo.CurrentCulture) || result > DateTime.Now)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1, text);
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1);
							}
						}
					}
				}
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct34")
				{
					num2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, staticCells[i].CellTemplate.ColIndex, "NumberCellType"));
				}
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/ThueGTGTPhatSinh/ct34")
				{
					num = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, staticCells[i].CellTemplate.ColIndex, "NumberCellType"));
					value2 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, HTKKFile.ColumnLettersToIndex("AB"), "NumberCellType"));
					cell = sheetView.Cells[staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1];
					HTKKFile.RemoveValidateBackGround(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1);
				}
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/ThueGTGTPhatSinh/ct34")
				{
					num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, staticCells[i].CellTemplate.ColIndex, "NumberCellType"));
					cell2 = sheetView.Cells[staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1];
					HTKKFile.RemoveValidate(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1);
				}
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/XDinhNVuThue/ct35")
				{
					num3 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, staticCells[i].CellTemplate.ColIndex, "NumberCellType"));
					cell2 = sheetView.Cells[staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1];
					HTKKFile.RemoveValidate(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1);
				}
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/XDinhNVuThue/ct36")
				{
					num4 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, staticCells[i].CellTemplate.ColIndex, "NumberCellType"));
					num5 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, HTKKFile.ColumnLettersToIndex("AB"), "NumberCellType"));
					cell3 = sheetView.Cells[staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1];
					HTKKFile.RemoveValidateBackGround(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1);
				}
				if (staticCells[i].Path == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/XDinhNVuThue/ct37")
				{
					num6 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, staticCells[i].CellTemplate.ColIndex, "NumberCellType"));
					num7 = ConvertHelpers.ConvertToDouble(GeneralFunctions.GetValueCustom(sheetView, staticCells[i].CellTemplate.RowIndex, HTKKFile.ColumnLettersToIndex("AB"), "NumberCellType"));
					cell4 = sheetView.Cells[staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1];
					HTKKFile.RemoveValidateBackGround(sheetView, staticCells[i].CellTemplate.RowIndex - 1, staticCells[i].CellTemplate.ColIndex - 1);
				}
			}
			if (num3 > num2)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell2.Row.Index, cell2.Column.Index, text3);
				value = 2;
			}
			else
			{
				string nam = PeriodInfo.Nam;
				string thang = PeriodInfo.Thang;
				string quy = PeriodInfo.Quy;
				string kieuKy = PeriodInfo.KieuKy;
				if (num3 > 0.0)
				{
					if (kieuKy.Equals("Y"))
					{
						if (ConvertHelpers.ConvertToInt(nam) > 2018)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, cell2.Row.Index, cell2.Column.Index, text4);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, cell2.Row.Index, cell2.Column.Index);
						}
					}
					else if (kieuKy.Equals("Q"))
					{
						if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(quy) == 4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, cell2.Row.Index, cell2.Column.Index, text4);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, cell2.Row.Index, cell2.Column.Index);
						}
					}
					else if (kieuKy.Equals("M") || kieuKy.Equals("D"))
					{
						if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(thang) >= 8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, cell2.Row.Index, cell2.Column.Index, text4);
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, cell2.Row.Index, cell2.Column.Index);
						}
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, cell2.Row.Index, cell2.Column.Index);
				}
			}
			if (num != Math.Round(value2, 0) && sheetView2.Visible)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell.Row.Index, cell.Column.Index, text2);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, cell.Row.Index, cell.Column.Index);
			}
			if (num4 != num5 && sheetView2.Visible)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell3.Row.Index, cell3.Column.Index, text5);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, cell3.Row.Index, cell3.Column.Index);
			}
			if (num6 != num7 && sheetView2.Visible)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, cell4.Row.Index, cell4.Column.Index, text6);
				value = 2;
			}
			else
			{
				HTKKFile.RemoveValidateBackGround(sheetView, cell4.Row.Index, cell4.Column.Index);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[0].Tables[2];
			int num8 = ConvertHelpers.ConvertToInt(hTKKCMCTableInfo.Location.DataLocationTemplateRowStart);
			for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
			{
				string valColMa = GeneralFunctions.GetValueCustom(sheetView, num8 + j, HTKKFile.ColumnLettersToIndex("AB"), "TextCellType");
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, num8 + j, HTKKFile.ColumnLettersToIndex("Q"), "NumberCellType");
				HTKKFile.RemoveValidate(sheetView, num8 + j - 1, HTKKFile.ColumnLettersToIndex("C") - 1);
				item item;
				if (!HTKKFormater.isNullString(valColMa) && _dmKSData != null && _dmKSData.Count > 0)
				{
					item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
					if (item != null && (HTKKFormater.isNullString(item.thuesuat) || ConvertHelpers.ConvertToInt(item.thuesuat) == 0))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num8 + j - 1, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36")));
						value = 2;
					}
				}
				HTKKFile.RemoveValidate(sheetView, num8 + j - 1, HTKKFile.ColumnLettersToIndex("Q") - 1);
				if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || !(valueCustom.ToNumber() != 0m))
				{
					continue;
				}
				bool flag = true;
				item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
				if (item == null)
				{
					continue;
				}
				if (item.cqtgia.ToNumber() > 0m)
				{
					flag = false;
					if (valueCustom.ToNumber() < item.cqtgia.ToNumber())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num8 + j - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
						value = 2;
					}
				}
				if (flag && valueCustom.ToNumber() < item.giatoithieu.ToNumber())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num8 + j - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		int value = 0;
		levelWarning.TryGetValue(1, out value);
		string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
		int num = 0;
		if (PeriodInfo.KieuKy.Equals("Y"))
		{
			num = (ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Substring(3, 4)) - ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Substring(3, 4))) * 12 + (ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Substring(0, 2)) - ConvertHelpers.ConvertToInt(PeriodInfo.TuThang.Substring(0, 2))) + 1;
		}
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			List<int> list = new List<int>();
			int num2 = HTKKFile.ColumnLettersToIndex("D");
			int num3 = HTKKFile.ColumnLettersToIndex("F");
			int num4 = HTKKFile.ColumnLettersToIndex("G");
			int num5 = HTKKFile.ColumnLettersToIndex("H");
			int num6 = HTKKFile.ColumnLettersToIndex("K");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("O");
			int num10 = HTKKFile.ColumnLettersToIndex("P");
			int num11 = HTKKFile.ColumnLettersToIndex("R");
			int num12 = HTKKFile.ColumnLettersToIndex("T");
			int num13 = HTKKFile.ColumnLettersToIndex("U");
			int num14 = HTKKFile.ColumnLettersToIndex("V");
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string text9 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text;
					string text10 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text;
					string text11 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text;
					string text12 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num5 - 1].Text;
					string valueCustom = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num6, "NumberCellType");
					string valueCustom2 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num7, "NumberCellType");
					string valueCustom3 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num8, "NumberCellType");
					string valueCustom4 = GeneralFunctions.GetValueCustom(sheetView, dataLocationTemplateRowStart + j, num9, "NumberCellType");
					string text13 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num10 - 1].Text;
					string text14 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num11 - 1].Text;
					string text15 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num12 - 1].Text;
					object value2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num13 - 1].Value;
					object value3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num14 - 1].Value;
					string text16 = "";
					if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, rowsValues[j].RowValue.ToArray()))
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						continue;
					}
					if (HTKKFormater.isNullString(text9))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1, text3);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num2 - 1);
						list.Add(0);
					}
					if (HTKKFormater.isNullString(text10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text3);
						list.Add(1);
					}
					else if (!ValidateBase.CheckMaSoThue(text10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text);
						list.Add(1);
					}
					else if (dictionary.Count > 0)
					{
						if (dictionary.ContainsKey(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1, text2);
							list.Add(1);
						}
						else
						{
							dictionary.Add(text10, "0");
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
							list.Add(0);
						}
					}
					else
					{
						dictionary.Add(text10, "0");
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num3 - 1);
						list.Add(0);
					}
					if (ConvertHelpers.ConvertToDouble(valueCustom) > ConvertHelpers.ConvertToDouble1(text12) * (double)num && PeriodInfo.KieuKy.Equals("Y"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1, text4);
						list.Add(1);
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num6 - 1);
						list.Add(0);
					}
					if (ConvertHelpers.ConvertToDouble(valueCustom2) > ConvertHelpers.ConvertToDouble(valueCustom))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text6);
						list.Add(1);
					}
					else
					{
						string nam = PeriodInfo.Nam;
						string thang = PeriodInfo.Thang;
						string quy = PeriodInfo.Quy;
						string kieuKy = PeriodInfo.KieuKy;
						if (ConvertHelpers.ConvertToDouble(valueCustom2) > 0.0)
						{
							if (kieuKy.Equals("Y"))
							{
								if (ConvertHelpers.ConvertToInt(nam) > 2018)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text5);
									list.Add(1);
								}
								else
								{
									HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
									list.Add(0);
								}
							}
							else if (kieuKy.Equals("Q"))
							{
								if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(quy) == 4))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text5);
									list.Add(1);
								}
								else
								{
									HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
									list.Add(0);
								}
							}
							else if (kieuKy.Equals("M") || kieuKy.Equals("D"))
							{
								if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(thang) >= 8))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1, text5);
									list.Add(1);
								}
								else
								{
									HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
									list.Add(0);
								}
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num7 - 1);
							list.Add(0);
						}
					}
					if (Math.Abs(ConvertHelpers.ConvertToDouble(valueCustom3) - ConvertHelpers.ConvertToDouble(value2)) < 1.0 || ConvertHelpers.ConvertToDouble(valueCustom3) == 0.0 || ConvertHelpers.ConvertToDouble(value2) == 0.0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1);
						list.Add(0);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num8 - 1, text7);
						list.Add(1);
					}
					if (Math.Abs(ConvertHelpers.ConvertToDouble(valueCustom4) - ConvertHelpers.ConvertToDouble(value3)) < 1.0 || ConvertHelpers.ConvertToDouble1(valueCustom4) == 0.0 || ConvertHelpers.ConvertToDouble(value3) == 0.0)
					{
						HTKKFile.RemoveValidate(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1);
						list.Add(0);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, dataLocationTemplateRowStart + j - 1, num9 - 1, text8);
						list.Add(1);
					}
				}
			}
			for (int j = 0; j < list.Count; j++)
			{
				if (list[j] == 1)
				{
					value = 2;
					break;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[1].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMKhoangSan(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string text = array[2];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKhoangSanNew(List<PathDm> pathDms, HTKKPeriodInfo PeriodInfo)
	{
		//IL_0383: Unknown result type (might be due to invalid IL or missing references)
		//IL_038a: Expected O, but got Unknown
		//IL_03b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bf: Expected O, but got Unknown
		string text = "###0103###020101###020302###020303###0301###0302###0303###0304###0401###0402###0403###0404###0405###0702###07030101###07030102###07030201###07030202###090201###090202###090204###090205###";
		try
		{
			_isDmMoi = true;
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			_dmKSData = new List<item>();
			List<item> list = new List<item>();
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (pathDm.pathDonGiaTinh != null && pathDm.pathDonGiaTinh.Count > 0)
					{
						foreach (string p in pathDm.pathDonGiaTinh)
						{
							List<string> list2 = pathDonGiaTinh;
							Predicate<string> match = (string o) => o.Equals(p);
							if (HTKKFormater.isNullString(list2.Find(match)))
							{
								pathDonGiaTinh.Add(p);
							}
						}
					}
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int num = 0; num < items.Count; num++)
					{
						item item = items[num];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							list.Add(item);
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							list.Add(item);
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item7 in elementsByTagName)
				{
					XmlNode val2 = item7;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array[^1] == "1")
					{
						item item3 = new item();
						item3.id = array[0];
						item3.ma = array[1];
						item3.ten = array[2];
						item3.dvt = array[3];
						item3.thuesuat = array[4];
						if (!namKK.Equals("2019") || !kieuKyKK.Equals("Y"))
						{
							list.Add(item3);
						}
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				List<item> dmTaiNguyenForTinh = GeneralFunctions.getDmTaiNguyenForTinh(pathDonGiaTinh, PeriodInfo);
				List<item> list3 = Serializer.DeepCopy(list);
				foreach (item iT in list)
				{
					Predicate<item> match2 = (item obj) => obj.ma.Equals(iT.ma);
					List<item> list4 = dmTaiNguyenForTinh.FindAll(match2);
					if (list4 != null && list4.Count > 0)
					{
						item item4 = list3.Find((item obj) => obj.id == iT.id);
						int num2 = list3.FindIndex((item obj) => obj.id == iT.id);
						if (item4 != null && PeriodInfo.KieuKy == "M")
						{
							list3.Remove(item4);
						}
						bool flag = false;
						for (int num3 = 0; num3 < list4.Count; num3++)
						{
							item item5 = Serializer.DeepCopy(iT);
							item5.cqtgia = list4[num3].cqtgia;
							item5.isCqt = true;
							list3.Insert(num2 + num3, item5);
						}
						list3.Remove(item4);
					}
				}
				list = Serializer.DeepCopy(list3);
				list3 = new List<item>();
				foreach (item item8 in list)
				{
					item item6 = Serializer.DeepCopy(item8);
					string ten = item6.ten;
					List<item> list5 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list5 != null && list5.Count > 1)
					{
						item6 = GeneralFunctions.getTaiNguyenForCombobox(item6, list5);
						if (item6 == null && item8.id == list5[0].id)
						{
							list3.Add(Serializer.DeepCopy(item8));
						}
					}
					if (item6 != null)
					{
						if (!HTKKFormater.isNullString(item6.level) && HTKKFormater.IsNumber(item6.level))
						{
							item6.ten = HTKKFile.addSpace(ConvertHelpers.ConvertToInt(item6.level)) + item6.ten;
						}
						list3.Add(item6);
					}
				}
				for (int num4 = 0; num4 < list3.Count; num4++)
				{
					item item6 = Serializer.DeepCopy(list3[num4]);
					string text2 = item6.id + "_" + item6.ma + "_" + item6.thuesuat + "_" + item6.giatoithieu + "_" + item6.cqtgia;
					string[] array = new string[7] { text2, item6.ma, item6.ten, item6.dvt, item6.thuesuat, "0", "1" };
					item6.id = text2;
					if (!dictionary.ContainsKey(text2))
					{
						string value2 = array[2];
						dictionary2.Add(text2, value2);
						dictionary.Add(text2, array);
						_dmKSData.Add(item6);
					}
				}
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<item> GetItemDm_GetDMKhoangSan()
	{
		return _dmKSData;
	}

	private bool kyKeKhaiTrongHL(item itemTinh, HTKKPeriodInfo PeriodInfo)
	{
		bool result = false;
		try
		{
			DateTime[] array = new DateTime[2];
			DateTime[] array2 = new DateTime[2];
			if (HTKKFormater.isNullString(itemTinh.cqthltu))
			{
				return false;
			}
			ref DateTime reference = ref array2[0];
			reference = DateTime.ParseExact(itemTinh.cqthltu, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			if (HTKKFormater.isNullString(itemTinh.cqthlden))
			{
				return false;
			}
			ref DateTime reference2 = ref array2[1];
			reference2 = DateTime.ParseExact(itemTinh.cqthlden, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			string nam = PeriodInfo.Nam;
			string tuNgay = PeriodInfo.TuNgay;
			DateTime dateTime;
			if (!HTKKFormater.isNullString(tuNgay))
			{
				dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			}
			else
			{
				tuNgay = PeriodInfo.TuThang;
				dateTime = (HTKKFormater.isNullString(tuNgay) ? DateTime.ParseExact("01/01" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture) : DateTime.ParseExact("01/" + tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture));
			}
			array[0] = dateTime;
			tuNgay = PeriodInfo.DenNgay;
			if (!HTKKFormater.isNullString(tuNgay))
			{
				dateTime = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			}
			else
			{
				tuNgay = PeriodInfo.DenThang;
				if (!HTKKFormater.isNullString(tuNgay))
				{
					dateTime = DateTime.ParseExact("01/" + tuNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				}
				else
				{
					dateTime = DateTime.ParseExact("01/12" + nam, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					dateTime = new DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(1).AddDays(-1.0);
				}
			}
			array[1] = dateTime;
			if (array[0] >= array2[0] && array[1] <= array2[1])
			{
				return true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return result;
	}

	public Dictionary<string, string> GetDMKSBVMT(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Expected O, but got Unknown
		try
		{
			if (dicDMKhoangSanTmp.Count > 0)
			{
				return dicDMKhoangSanTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string text = array[2];
				dictionary2.Add(key, text);
				dictionary.Add(text, array);
			}
			dicDMKhoangSan = dictionary;
			dicDMKhoangSanTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMKSBVMTNew(List<PathDm> pathDms)
	{
		//IL_02ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f4: Expected O, but got Unknown
		//IL_0322: Unknown result type (might be due to invalid IL or missing references)
		//IL_0329: Expected O, but got Unknown
		try
		{
			_isDmMoi = true;
			if (dicDMKSBVMTTmp.Count > 0)
			{
				return dicDMKSBVMTTmp;
			}
			_dmBVMTData = new List<item>();
			List<item> list = new List<item>();
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
							{
								list.Add(item);
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
						{
							list.Add(item);
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item4 in elementsByTagName)
				{
					XmlNode val2 = item4;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array[^1] == "1")
					{
						item item3 = new item();
						item3.id = array[0];
						item3.ma = array[1];
						item3.ten = array[2];
						item3.dvt = array[3];
						item3.thuesuat = array[4];
						if (!namKK.Equals("2019") || !kieuKyKK.Equals("Y"))
						{
							list.Add(item3);
						}
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (item item5 in list)
				{
					string ten = item5.ten;
					List<item> list2 = list.FindAll((item obj) => obj.ten.Equals(ten));
					if (list2 != null && list2.Count > 1)
					{
						ten = ten + "( " + item5.ma + " )";
					}
					string text = item5.id + "_" + item5.ma + "_" + item5.thuesuat + "_" + item5.giatoithieu + "_" + item5.cqtgia;
					string[] array = new string[7] { text, item5.ma, ten, item5.dvt, item5.thuesuat, "0", "1" };
					item5.id = text;
					if (!dictionary.ContainsKey(text))
					{
						string text2 = array[2];
						dictionary2.Add(text, text2);
						if (!dictionary.ContainsKey(text2))
						{
							dictionary.Add(text2, array);
						}
						_dmBVMTData.Add(item5);
					}
				}
			}
			dicDMKSBVMT = dictionary;
			dicDMKSBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public List<item> GetItemDm_GetDMKSBVMT()
	{
		return _dmBVMTData;
	}

	public Dictionary<string, string> GetDMDonVi(string pathFile)
	{
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		try
		{
			if (dicDMDonVi.Count > 0)
			{
				return dicDMDonVi;
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			int num = 0;
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["value"]).Value;
				num++;
				dictionary.Add(num.ToString(), value);
			}
			dicDMDonVi = dictionary;
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void RemoveDataRows(SheetView sView, int frRow, int rowCount, HTKKCMCSheetInfo sInfo)
	{
		if (rowCount == 0)
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

	public static void fpSpReindexingTableTemplate(SheetView sView, HTKKCMCTableInfo tableInfo)
	{
		try
		{
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				_ = tableInfo.iColumnSTT;
				bool flag = 1 == 0;
				if (tableInfo.CellTableSTT != null)
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = tableInfo.CellTableSTT.Value + "." + (i + 1);
				}
				else
				{
					sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.iColumnSTT - 1].Text = (i + 1).ToString();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMThueBVMT(string pathFile, HTKKPeriodInfo PeriodInfo)
	{
		//IL_03ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f6: Expected O, but got Unknown
		//IL_0436: Unknown result type (might be due to invalid IL or missing references)
		//IL_043d: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "112020#092020#012020#012019#012018#052017#052015#092014#092013#022013#011900";
			List<string> list = new List<string>();
			string iNTERFACE_INI = HTKKConfig.INTERFACE_INI;
			DirectoryInfo directoryInfo = new DirectoryInfo(iNTERFACE_INI);
			if (directoryInfo.Exists)
			{
				FileInfo[] files = directoryInfo.GetFiles();
				if (files.Length > 0)
				{
					FileInfo[] array = files;
					foreach (FileInfo fileInfo in array)
					{
						if (fileInfo.Name.Contains("Catalogue_DM_thue_BVMT_"))
						{
							list.Add(fileInfo.Name.Replace("Catalogue_DM_thue_BVMT_", "").Replace(".xml", ""));
						}
					}
				}
			}
			DateTime t = new DateTime(1900, 1, 1);
			int num = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
			int month = ConvertHelpers.ConvertToInt(PeriodInfo.Thang);
			if (dicDMThueBVMTTmp.Count > 0)
			{
				return dicDMThueBVMTTmp;
			}
			string text3 = "";
			if (PeriodInfo.KieuKy == "M" || PeriodInfo.KieuKy == "D")
			{
				DateTime t2 = new DateTime(num, month, 1).AddMonths(1).AddDays(-1.0);
				for (int j = 0; j < list.Count; j++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(list[j].Substring(2)), ConvertHelpers.ConvertToInt(list[j].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = list[j];
					}
				}
			}
			else if (PeriodInfo.KieuKy == "Y")
			{
				DateTime t2 = new DateTime(num, 12, 1).AddMonths(1).AddDays(-1.0);
				if (PeriodInfo.Nam == "2020")
				{
					text = "012020";
				}
				else
				{
					for (int j = 0; j < list.Count; j++)
					{
						DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(list[j].Substring(2)), ConvertHelpers.ConvertToInt(list[j].Substring(0, 2)), 1);
						if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
						{
							t = dateTime;
							text = list[j];
						}
					}
				}
			}
			else if (PeriodInfo.KieuKy == "Q")
			{
				int iQuy = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				DateTime t2 = GeneralFunctions.GetNgayDauQuy(iQuy, num, 1, 1).AddMonths(1).AddDays(-1.0);
				for (int j = 0; j < list.Count; j++)
				{
					DateTime dateTime = new DateTime(ConvertHelpers.ConvertToInt(list[j].Substring(2)), ConvertHelpers.ConvertToInt(list[j].Substring(0, 2)), 1);
					if (DateTime.Compare(t2, dateTime) >= 0 && DateTime.Compare(dateTime, t) >= 0)
					{
						t = dateTime;
						text = list[j];
					}
				}
			}
			text3 = pathFile + "\\Catalogue_DM_thue_BVMT_" + text + ".xml";
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text3))
			{
				return null;
			}
			val.Load(text3);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array2 = Regex.Split(value, "###");
				string key = array2[0];
				string value2 = array2[1];
				dictionary2.Add(key, value2);
				dictionary.Add(array2[1], array2);
			}
			dicDMThueBVMT = dictionary;
			dicDMThueBVMTTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMNganhNgheKD(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (dicDMNganhNgheKDTmp.Count > 0)
			{
				return dicDMNganhNgheKDTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 2 && array[2] == "1")
				{
					string key = array[0];
					string value2 = array[1];
					dictionary2.Add(key, value2);
					dictionary.Add(array[1], array);
				}
			}
			dicDMNganhNgheKD = dictionary;
			dicDMNganhNgheKDTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMNhomNganhNgheKD(string pathFile)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			string text = pathFile + "\\DM_NHOMNGANHNGHE.xml";
			if (dicDMNhomNganhNgheKDTmp.Count > 0)
			{
				return dicDMNhomNganhNgheKDTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return null;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 1)
				{
					string key = array[0];
					string value2 = array[1];
					dictionary2.Add(key, value2);
					dictionary.Add(array[1], array);
				}
			}
			dicDMNhomNganhNgheKD = dictionary;
			dicDMNhomNganhNgheKDTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDB(string pathFile)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(pathFile))
			{
				return null;
			}
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				if (array.Length > 4)
				{
					string key = array[1];
					string text = array[2].Trim() + " (" + array[4].Trim() + "%)";
					if (!dictionary2.ContainsKey(key))
					{
						dictionary2.Add(key, text);
					}
					if (!dictionary.ContainsKey(text))
					{
						dictionary.Add(text, array);
					}
				}
			}
			dicDMThueDB = dictionary;
			dicDMThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMThueTieuThuDBNew(List<PathDm> pathDms)
	{
		//IL_04a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_04aa: Expected O, but got Unknown
		//IL_04d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_04df: Expected O, but got Unknown
		try
		{
			if (dicDMThueDBTmp.Count > 0)
			{
				return dicDMThueDBTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			Serializer serializer = new Serializer();
			if (pathDms == null || pathDms.Count == 0)
			{
				return null;
			}
			foreach (PathDm pathDm in pathDms)
			{
				if (pathDm == null)
				{
					continue;
				}
				string path;
				if (pathDm.dmNew)
				{
					if (HTKKFormater.isNullString(pathDm.pathDataDm) || HTKKFormater.isNullString(pathDm.pathSaveInt))
					{
						continue;
					}
					path = pathDm.pathSaveDataFile;
					if (!File.Exists(path))
					{
						path = pathDm.pathSaveInt;
					}
					string input = File.ReadAllText(path);
					dmconfig dmconfig = serializer.Deserialize<dmconfig>(input);
					input = File.ReadAllText(pathDm.pathDataDm);
					danhmuc danhmuc = serializer.Deserialize<danhmuc>(input);
					if (danhmuc == null || danhmuc.items.Count <= 0 || dmconfig == null || dmconfig.dmcfs.Count <= 0)
					{
						continue;
					}
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count == 0)
					{
						return null;
					}
					for (int i = 0; i < items.Count; i++)
					{
						item item = items[i];
						item item2 = item;
						if (HTKKFormater.isNullString(dmcf.active))
						{
							if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check) && item.ma.StartsWith("1"))
							{
								string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
								string key = array[1];
								string text = array[2];
								if (!string.IsNullOrEmpty(item.thuesuat))
								{
									text = text + " (" + item.thuesuat + " %)";
								}
								if (!dictionary2.ContainsKey(key))
								{
									dictionary2.Add(key, text);
								}
								if (!dictionary.ContainsKey(text))
								{
									dictionary.Add(text, array);
								}
							}
						}
						else if (dmcf.active.IndexOf("~" + item.id + "~") > -1 && item.ma.StartsWith("1") && item.check.Equals("true"))
						{
							string[] array = new string[7] { item.id, item.ma, item.ten, item.dvt, item.thuesuat, "0", "1" };
							string key = array[1];
							string text = array[2];
							if (!string.IsNullOrEmpty(item.thuesuat))
							{
								text = text + " (" + item.thuesuat + " %)";
							}
							if (!dictionary2.ContainsKey(key))
							{
								dictionary2.Add(key, text);
							}
							if (!dictionary.ContainsKey(text))
							{
								dictionary.Add(text, array);
							}
						}
					}
					continue;
				}
				path = pathDm.pathSaveDataFileOld;
				if (!File.Exists(path))
				{
					path = pathDm.pathSaveIntOld;
				}
				XmlDocument val = new XmlDocument();
				val.Load(path);
				XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
				foreach (XmlNode item3 in elementsByTagName)
				{
					XmlNode val2 = item3;
					string value = ((XmlNode)val2.Attributes["Value"]).Value;
					string[] array = Regex.Split(value, "###");
					if (array.Length > 2 && array[1].StartsWith("1") && array[^1] == "1")
					{
						string key = array[1];
						string text = array[2].Trim();
						string text2 = array[4].Trim();
						if (!string.IsNullOrEmpty(text2))
						{
							text = text + " (" + text2 + " %)";
						}
						if (!dictionary2.ContainsKey(key))
						{
							dictionary2.Add(key, text);
						}
						if (!dictionary.ContainsKey(text))
						{
							dictionary.Add(text, array);
						}
					}
				}
			}
			dicDMThueDB = dictionary;
			dicDMThueDBTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void AutoBindData_1(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("U");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int iCol = HTKKFile.ColumnLettersToIndex("H");
			int num5 = HTKKFile.ColumnLettersToIndex("K");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("N");
			int num8 = HTKKFile.ColumnLettersToIndex("O");
			int num9 = HTKKFile.ColumnLettersToIndex("P");
			int num10 = HTKKFile.ColumnLettersToIndex("R");
			int num11 = HTKKFile.ColumnLettersToIndex("T");
			int num12 = HTKKFile.ColumnLettersToIndex("U");
			int num13 = HTKKFile.ColumnLettersToIndex("V");
			if (PeriodInfo.KieuKy == "Y")
			{
				double num14 = (ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiDenThang, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(TTinChung.kyKKhaiTuThang, "MM/yyyy").Month + 1;
				HTKKCMCTableInfo hTKKCMCTableInfo = ToKhaiInfo.SheetInfos[1].Tables[0];
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					string valueCustom = GeneralFunctions.GetValueCustom(sView, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, iCol, "NumberCellType");
					if (ConvertHelpers.ConvertToDouble(valueCustom) * num14 < 100000000.0)
					{
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num5 - 1].Value = 0;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num5 - 1].Locked = true;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num6 - 1].Value = 0;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num6 - 1].Locked = true;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num7 - 1].Value = 0;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num7 - 1].Locked = true;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num8 - 1].Value = 0;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num8 - 1].Locked = true;
					}
					else
					{
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num5 - 1].Locked = false;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num6 - 1].Locked = false;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num7 - 1].Locked = false;
						sView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num8 - 1].Locked = false;
					}
				}
			}
			string nam = PeriodInfo.Nam;
			string thang = PeriodInfo.Thang;
			string quy = PeriodInfo.Quy;
			string kieuKy = PeriodInfo.KieuKy;
			int count = ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count;
			for (int i = 0; i < count; i++)
			{
				int row = ToKhaiInfo.SheetInfos[1].Tables[0].Location.DataLocationTemplateRowStart + i - 1;
				if (kieuKy.Equals("Y"))
				{
					if (ConvertHelpers.ConvertToInt(nam) > 2018)
					{
						sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Value = "0";
						sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sView.Cells[row, HTKKFile.ColumnLettersToIndex("P") - 1].Value = "0";
					}
				}
				else if (kieuKy.Equals("Q"))
				{
					if (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(quy) == 4))
					{
						sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Value = "0";
						sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
						sView.Cells[row, HTKKFile.ColumnLettersToIndex("P") - 1].Value = "0";
					}
				}
				else if ((kieuKy.Equals("M") || kieuKy.Equals("D")) && (ConvertHelpers.ConvertToInt(nam) > 2018 || (ConvertHelpers.ConvertToInt(nam) == 2018 && ConvertHelpers.ConvertToInt(thang) >= 8)))
				{
					sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Value = "0";
					sView.Cells[row, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
					sView.Cells[row, HTKKFile.ColumnLettersToIndex("P") - 1].Value = "0";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Refresh_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int actionType)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			switch (actionType)
			{
			case 2:
			{
				string pathTong11 = "PLuc_01_1_BK_CNKD/ct19";
				string path30a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct30";
				string path30c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct30";
				string path31a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct31";
				string path31c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct31";
				string path32a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct32";
				string path32c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct32";
				string path33a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct33";
				string path33c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct33";
				string pathTK35 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/XDinhNVuThue/ct35";
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTong11));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path30a));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path30c));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path31a));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path31c));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path32a));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path32c));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path33a));
				HTKKCMCCellInfo hTKKCMCCellInfo9 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path33c));
				HTKKCMCCellInfo hTKKCMCCellInfo10 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTK35));
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Formula = "";
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = 0;
				sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value = 0;
				break;
			}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			string pathTong11 = "PLuc_01_1_BK_CNKD/ct19";
			string path30a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct30";
			string path30c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct30";
			string path31a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct31";
			string path31c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct31";
			string path32a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct32";
			string path32c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct32";
			string path33a = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueGTGT/ct33";
			string path33c = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/DoanhThuThueTNCN/ct33";
			string path34b = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/ThueGTGTPhatSinh/ct34";
			string path34d = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/KKThueGTGT_TNCN/ThueTNCNPhatSinh/ct34";
			string pathTK35 = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/XDinhNVuThue/ct35";
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTong11));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path30a));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path30c));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path31a));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path31c));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path32a));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path32c));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path33a));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path33c));
			HTKKCMCCellInfo hTKKCMCCellInfo10 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path34b));
			HTKKCMCCellInfo hTKKCMCCellInfo11 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(path34d));
			HTKKCMCCellInfo hTKKCMCCellInfo12 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTK35));
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Formula = "";
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value = 0;
			sheetView.FpSpread.Sheets[0].Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Value = 0;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			int num = HTKKFile.ColumnLettersToIndex("G");
			int num2 = HTKKFile.ColumnLettersToIndex("U");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("H");
			int num6 = HTKKFile.ColumnLettersToIndex("K");
			int num7 = HTKKFile.ColumnLettersToIndex("M");
			int num8 = HTKKFile.ColumnLettersToIndex("N");
			int num9 = HTKKFile.ColumnLettersToIndex("O");
			int num10 = HTKKFile.ColumnLettersToIndex("P");
			int num11 = HTKKFile.ColumnLettersToIndex("R");
			int num12 = HTKKFile.ColumnLettersToIndex("T");
			int num13 = HTKKFile.ColumnLettersToIndex("U");
			int num14 = HTKKFile.ColumnLettersToIndex("V");
			string[] array = dicDMNhomNganhNgheKD[sheetView.Cells[17, num - 1].Text.Trim()];
			if (array.Length > 2)
			{
				sheetView.Cells[17, num2 - 1].Text = array[0];
				sheetView.Cells[18, num2 - 1].Text = array[2];
				sheetView.Cells[19, num2 - 1].Text = array[3];
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10/ct10.1_ma"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10/ct10.1_ten"));
			if (hTKKCMCCellInfo != null && hTKKCMCCellInfo2 != null && HTKKFormater.isNullString(hTKKCMCCellInfo.Value))
			{
				string value = "";
				HTKKWinDictionaryDiaDanh.dicTinhTP.TryGetValue(hTKKCMCCellInfo.Value, out value);
				winFpSpread.Sheets[1].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = value;
			}
			hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10/ct10.2_ma"));
			hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10/ct10.2_ten"));
			if (hTKKCMCCellInfo != null && hTKKCMCCellInfo2 != null && HTKKFormater.isNullString(hTKKCMCCellInfo.Value))
			{
				string value = "";
				HTKKWinDictionaryDiaDanh.dicQuanHuyen.TryGetValue(hTKKCMCCellInfo.Value, out value);
				winFpSpread.Sheets[1].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = value;
			}
			hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10/ct10.2_ma"));
			hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/CT10/ct10.2_ten"));
			if (hTKKCMCCellInfo != null && hTKKCMCCellInfo2 != null && HTKKFormater.isNullString(hTKKCMCCellInfo.Value))
			{
				string value = "";
				HTKKWinDictionaryDiaDanh.dicPhuongXa.TryGetValue(hTKKCMCCellInfo.Value, out value);
				winFpSpread.Sheets[1].Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public double getDonGia(string ma)
	{
		double result = 0.0;
		try
		{
			if (!HTKKFormater.isNullString(ma) && _dmKSData != null && _dmKSData.Count > 0)
			{
				item item = _dmKSData.Find((item x) => x.id.Contains(ma));
				if (item != null && item.cqtgia.ToNumber() > 0m)
				{
					result = ConvertHelpers.ConvertToDouble1(item.cqtgia.ToNumber());
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}

	public void convertId2Ma(XmlNode xmlNode, string path, int type)
	{
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (!_isDmMoi || xmlNode == null)
			{
				return;
			}
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(path);
			if (val == null || val.Count <= 0 || _dmKSData == null || _dmKSData.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item2 in val)
			{
				XmlNode val2 = item2;
				string idTN = val2.InnerText;
				if (!HTKKFormater.isNullString(idTN))
				{
					item item = _dmKSData.Find((item obj) => obj.id.Equals(idTN));
					item = ((type != 1) ? _dmBVMTData.Find((item obj) => obj.id.Equals(idTN)) : _dmKSData.Find((item obj) => obj.id.Equals(idTN)));
					if (item != null)
					{
						val2.InnerText = item.ma;
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

	public void convertMa2Id(XmlNode xmlNode, string path, HTKKPeriodInfo PeriodInfo, int type)
	{
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		try
		{
			if (!_isDmMoi || xmlNode == null)
			{
				return;
			}
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(path);
			if (val == null || val.Count <= 0 || _dmKSData == null || _dmKSData.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				string maTN = val2.InnerText;
				string[] array = new string[2];
				if (!HTKKFormater.isNullString(maTN))
				{
					List<item> list = ((type != 1) ? _dmBVMTData.FindAll((item obj) => obj.ma.Equals(maTN)) : _dmKSData.FindAll((item obj) => obj.ma.Equals(maTN)));
					if (list != null && list.Count > 0 && !HTKKFormater.isNullString(list[0].id))
					{
						array[0] = list[0].id;
					}
				}
				if (!HTKKFormater.isNullString(array[1]))
				{
					val2.InnerText = array[1];
				}
				else
				{
					val2.InnerText = array[0];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void StartImportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
			{
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KKhaiTBVMT_TN/ThueTaiNguyen/CTietThueTaiNguyen/ct2_ma", PeriodInfo, 1);
				convertMa2Id(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KKhaiTBVMT_TN/PhiBVMT/CTietPhiBVMT/ct2_ma", PeriodInfo, 2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			if (PeriodInfo.dicVar != null && PeriodInfo.dicVar.Count > 0 && PeriodInfo.dicVar.ContainsKey("IS_DM_MOI"))
			{
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KKhaiTBVMT_TN/ThueTaiNguyen/CTietThueTaiNguyen/ct2_ma", 1);
				convertId2Ma(xmlNode, "HSoKhaiThue/CTieuTKhaiChinh/KKhaiTBVMT_TN/PhiBVMT/CTietPhiBVMT/ct2_ma", 2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void validateTaiNguyen(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			if (iCol != HTKKFile.ColumnLettersToIndex("Q"))
			{
				return;
			}
			string valColMa = GeneralFunctions.GetValueCustom(sView, (int)iRow, HTKKFile.ColumnLettersToIndex("AB"), "TextCellType");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, (int)iRow, HTKKFile.ColumnLettersToIndex("Q"), "NumberCellType");
			HTKKFile.RemoveValidate(sView, (int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1);
			if (HTKKFormater.isNullString(valColMa) || _dmKSData == null || _dmKSData.Count <= 0 || !(valueCustom.ToNumber() != 0m))
			{
				return;
			}
			bool flag = true;
			item item = _dmKSData.Find((item x) => x.id.Contains(valColMa));
			if (item == null)
			{
				return;
			}
			if (item.cqtgia.ToNumber() > 0m)
			{
				flag = false;
				if (valueCustom.ToNumber() < item.cqtgia.ToNumber())
				{
					HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_35")));
				}
			}
			if (flag && valueCustom.ToNumber() < item.giatoithieu.ToNumber())
			{
				HTKKFile.SetNodeBackGroundPink(sView, (int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sView, GeneralFunctions.ArrayCellId("E_34")));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
