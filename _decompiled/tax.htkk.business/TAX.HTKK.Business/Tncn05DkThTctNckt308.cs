using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tncn05DkThTctNckt308
{
	private const string VIETNAM = "VN";

	private const string CMND = "1010";

	private const string CCCD = "2080";

	private const string ho_chieu = "1020";

	private const string TEN_VN = "Việt Nam";

	private const string KHAC = "Khác";

	private string CellID_G8 = "G_8";

	private bool isChangeByForm = false;

	private string CellID_P8 = "P_8";

	private string CellID_G35 = "G_35";

	private string CellID_T35 = "T_35";

	private string CellID_G10 = "G_10";

	private string CellID_G12 = "G_12";

	private int[] CellId_G8 => GeneralFunctions.ArrayCellId(CellID_G8);

	private int[] CellId_P8 => GeneralFunctions.ArrayCellId(CellID_P8);

	private int[] CellId_G35 => GeneralFunctions.ArrayCellId(CellID_G35);

	private int[] CellId_T35 => GeneralFunctions.ArrayCellId(CellID_T35);

	private int[] CellId_G10 => GeneralFunctions.ArrayCellId(CellID_G10);

	private int[] CellId_G12 => GeneralFunctions.ArrayCellId(CellID_G12);

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int[] array = GeneralFunctions.ArrayCellId("G_8");
			int[] array2 = GeneralFunctions.ArrayCellId("P_8");
			int[] array3 = GeneralFunctions.ArrayCellId("G_35");
			int[] array4 = GeneralFunctions.ArrayCellId("U_35");
			bool flag = sheetView.Cells[GeneralFunctions.ArrayCellId("G_8")[1], GeneralFunctions.ArrayCellId("G_8")[0]].Value.ToBoolOrDefault();
			string mst = TTinChung.mst;
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int num = location.DataLocationTemplateRowStart - 1;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			for (int i = 0; i < rowsValues.Count; i++)
			{
				string[] array5 = rowsValues[i].RowValue.ToArray();
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				string text;
				DateTime dateTime;
				string text5;
				string value4;
				string value5;
				string text2;
				string text3;
				bool flag2;
				string value2;
				string text4;
				DateTime dateTime2;
				string value3;
				string text6;
				string text7;
				string text8;
				string maHuyen;
				string text9;
				string text10;
				string text11;
				string value6;
				string text12;
				string text13;
				string text14;
				string text15;
				string maHuyen2;
				string text16;
				string text17;
				string text18;
				if (!flag)
				{
					text = array5[2];
					dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
					if (string.IsNullOrEmpty(text))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
						value = 2;
					}
					else if (dateTime == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_39")));
						value = 2;
					}
					else if (dateTime > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
						value = 2;
					}
					text2 = array5[5];
					text3 = array5[8];
					flag2 = array5[35].ToBoolOrDefault();
					if (string.IsNullOrEmpty(text3) && text2.Equals("VN"))
					{
						if (!flag2)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
							value = 2;
						}
					}
					else if (text2.Equals("VN") && text3.Length != 12)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_46")));
						value = 2;
					}
					else if (text2.Equals("VN") && !HTKKFormater.IsNumber(text3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65")));
						value = 2;
					}
					value2 = array5[10];
					if (string.IsNullOrEmpty(value2) && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
						value = 2;
					}
					text4 = array5[11];
					dateTime2 = ConvertHelpers.ConvertToDateTime(text4, "dd/MM/yyyy");
					if (string.IsNullOrEmpty(text4) && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
						value = 2;
					}
					else if (dateTime2 == DateTime.MinValue && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
						value = 2;
					}
					else if (dateTime2 > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48")));
						value = 2;
					}
					else if (!string.IsNullOrEmpty(text) && dateTime2 != DateTime.MinValue && dateTime != DateTime.MinValue && dateTime >= dateTime2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
						value = 2;
					}
					value3 = array5[12];
					if (string.IsNullOrEmpty(value3) && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50")));
						value = 2;
					}
					string mst2 = array5[1];
					if (string.IsNullOrEmpty(mst2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThueCCCD(mst2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
						value = 2;
					}
					else
					{
						int num2 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[1].Equals(mst2));
						if (mst2.Equals(TTinChung.mst))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
							value = 2;
						}
						else if (num2 > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
							value = 2;
						}
					}
					text5 = array5[3];
					value4 = array5[4];
					if (string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_67")));
						value = 2;
					}
					else if (!text5.Equals("1") && !text5.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
						value = 2;
					}
					value5 = array5[6];
					if (string.IsNullOrEmpty(value5))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
						value = 2;
					}
					text6 = array5[13];
					text7 = array5[14];
					text8 = array5[15];
					maHuyen = array5[16];
					text9 = array5[17];
					text10 = array5[18];
					text11 = array5[19];
					value6 = array5[20];
					text12 = array5[21];
					if (!string.IsNullOrEmpty(text6 + text8 + text11 + text12) && text12 == "Việt Nam")
					{
						if (string.IsNullOrEmpty(text6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text8) && !text7.Equals("5050307"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
							value = 2;
						}
						else if (text8.Trim() != "Khác" && string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
							value = 2;
						}
						else if (text11.Trim() != "Khác" && string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text12))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(value6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
							value = 2;
						}
						if (!GeneralFunctions.ValidateDiaBan2Cap(text10, maHuyen, text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
					}
					else if (!string.IsNullOrEmpty(text6 + text8 + text11 + text12) && text12 != "Việt Nam")
					{
						if (string.IsNullOrEmpty(text6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text12))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(value6))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
							value = 2;
						}
					}
					else
					{
						if (text12 == "Việt Nam" && text8.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						if (text12 == "Việt Nam" && text11.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
					}
					text13 = array5[22];
					text14 = array5[23];
					text15 = array5[24];
					maHuyen2 = array5[25];
					text16 = array5[26];
					text17 = array5[27];
					text18 = array5[28];
					if (!string.IsNullOrEmpty(text13 + text15 + text18))
					{
						if (string.IsNullOrEmpty(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text15) && !text14.Equals("5050307"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
							value = 2;
						}
						else if (text15.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else if (text15.Trim() != "Khác" && string.IsNullOrEmpty(text14))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text18))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
							value = 2;
						}
						else if (text18.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else if (text18.Trim() != "Khác" && string.IsNullOrEmpty(text17))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
							value = 2;
						}
						if (!GeneralFunctions.ValidateDiaBan2Cap(text17, maHuyen2, text14))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
					}
				}
				if (!flag)
				{
					continue;
				}
				string value7 = array5[0];
				if (string.IsNullOrEmpty(value7))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				string text19 = array5[1];
				if (!string.IsNullOrEmpty(text19) && !ValidateBase.CheckMaSoThueCCCD(text19))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
					value = 2;
				}
				text = array5[2];
				dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38")));
					value = 2;
				}
				else if (dateTime == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_39")));
					value = 2;
				}
				else if (dateTime > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40")));
					value = 2;
				}
				text5 = array5[3];
				value4 = array5[4];
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41")));
					value = 2;
				}
				else if (!text5.Equals("1") && !text5.Equals("2"))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_25")));
					value = 2;
				}
				value5 = array5[6];
				if (string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
					value = 2;
				}
				text2 = array5[5];
				text3 = array5[8];
				flag2 = array5[35].ToBoolOrDefault();
				if (string.IsNullOrEmpty(text3) && text2.Equals("VN"))
				{
					if (!flag2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
						value = 2;
					}
				}
				else if (text2.Equals("VN") && text3.Length != 12)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_46")));
					value = 2;
				}
				else if (text2.Equals("VN") && !HTKKFormater.IsNumber(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65")));
					value = 2;
				}
				value2 = array5[10];
				if (string.IsNullOrEmpty(value2) && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				text4 = array5[11];
				dateTime2 = ConvertHelpers.ConvertToDateTime(text4, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text4) && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
					value = 2;
				}
				else if (dateTime2 == DateTime.MinValue && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
					value = 2;
				}
				else if (dateTime2 > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48")));
					value = 2;
				}
				else if (!string.IsNullOrEmpty(text) && dateTime2 != DateTime.MinValue && dateTime != DateTime.MinValue && dateTime >= dateTime2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
					value = 2;
				}
				value3 = array5[12];
				if (string.IsNullOrEmpty(value3) && (!text2.Equals("VN") || (text2.Equals("VN") && flag2)))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50")));
					value = 2;
				}
				text6 = array5[13];
				text7 = array5[14];
				text8 = array5[15];
				maHuyen = array5[16];
				text9 = array5[17];
				text10 = array5[18];
				text11 = array5[19];
				value6 = array5[20];
				text12 = array5[21];
				if (text12 == "Việt Nam")
				{
					if (string.IsNullOrEmpty(text6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text8) && !text7.Equals("5050307"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
						value = 2;
					}
					else if (text8.Trim() != "Khác" && string.IsNullOrEmpty(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text11))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						value = 2;
					}
					else if (text11.Equals("Khác"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
					else if (text11.Trim() != "Khác" && string.IsNullOrEmpty(text10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text12))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
						value = 2;
					}
					if (!GeneralFunctions.ValidateDiaBan2Cap(text10, maHuyen, text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
				}
				else if (text12 != "Việt Nam")
				{
					if (string.IsNullOrEmpty(text6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text12))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35")));
						value = 2;
					}
				}
				text13 = array5[22];
				text14 = array5[23];
				text15 = array5[24];
				maHuyen2 = array5[25];
				text16 = array5[26];
				text17 = array5[27];
				text18 = array5[28];
				if (!string.IsNullOrEmpty(text13 + text15 + text18))
				{
					if (string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text15) && !text14.Equals("5050307"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
						value = 2;
					}
					else if (text15.Trim() != "Khác" && string.IsNullOrEmpty(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
						value = 2;
					}
					else if (text18.Trim() != "Khác" && string.IsNullOrEmpty(text17))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
						value = 2;
					}
					if (!GeneralFunctions.ValidateDiaBan2Cap(text17, maHuyen2, text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
					value = 2;
				}
			}
			removeMaKhac(sheetView, hTKKCMCSheetInfo);
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

	public void removeMaKhac(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		foreach (HTKKCMCTableInfo table in sInfo.Tables)
		{
			if (table == null)
			{
				continue;
			}
			List<HTKKCMCTableCellInfo> list = table.RowsInfo.Cells.FindAll((HTKKCMCTableCellInfo x) => x.Controltype == "100");
			if (list != null && list.Count > 0)
			{
				foreach (HTKKCMCTableCellInfo cellTenTinh in list)
				{
					List<HTKKCMCTableCellInfo> cells = table.RowsInfo.Cells;
					Predicate<HTKKCMCTableCellInfo> match = (HTKKCMCTableCellInfo x) => x.Controltype == "100" && x.CellTemplate.ColIndex == cellTenTinh.CellTemplate.ColIndex;
					int index = cells.FindIndex(match);
					int index2 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(cellTenTinh.indexColumnSelectedValue));
					int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
					for (int num = 0; num < table.RowsValues.Count; num++)
					{
						string text = table.RowsValues[num].RowValue[index];
						if (text == "Khác")
						{
							sView.Cells[dataLocationTemplateRowStart + num - 1, HTKKFile.ColumnLettersToIndex(cellTenTinh.indexColumnSelectedValue) - 1].Text = "";
							table.RowsValues[num].RowValue[index2] = "";
						}
					}
				}
			}
			List<HTKKCMCTableCellInfo> list2 = table.RowsInfo.Cells.FindAll((HTKKCMCTableCellInfo x) => x.Controltype == "101");
			if (list2 == null || list2.Count <= 0)
			{
				continue;
			}
			foreach (HTKKCMCTableCellInfo cellTenXa in list2)
			{
				List<HTKKCMCTableCellInfo> cells2 = table.RowsInfo.Cells;
				Predicate<HTKKCMCTableCellInfo> match2 = (HTKKCMCTableCellInfo x) => x.Controltype == "101" && x.CellTemplate.ColIndex == cellTenXa.CellTemplate.ColIndex;
				int index3 = cells2.FindIndex(match2);
				int index4 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(cellTenXa.indexColumnSelectedValue));
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				for (int num = 0; num < table.RowsValues.Count; num++)
				{
					string text2 = table.RowsValues[num].RowValue[index3];
					if (text2 == "Khác")
					{
						table.RowsValues[num].RowValue[index4] = "";
						sView.Cells[dataLocationTemplateRowStart + num - 1, HTKKFile.ColumnLettersToIndex(cellTenXa.indexColumnSelectedValue) - 1].Text = "";
					}
				}
			}
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 3 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_05DK_TH.xls");
			}
			else
			{
				if (winFpSpread.ActiveSheetIndex != 0)
				{
					return;
				}
				SheetView sheetView = winFpSpread.Sheets[0];
				bool flag = Convert.ToBoolean(sheetView.Cells[CellId_G8[1], CellId_G8[0]].Value);
				bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_P8[1], CellId_P8[0]].Value);
				bool flag3 = flag || flag2;
				if (idxRow == CellId_G8[1] && idxCol == CellId_G8[0])
				{
					if (!flag)
					{
						sheetView.Cells[CellId_P8[1], CellId_P8[0]].Value = true;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
						}
					}
					else
					{
						sheetView.Cells[CellId_P8[1], CellId_P8[0]].Value = false;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = false;
						}
					}
				}
				else
				{
					if (idxRow != CellId_P8[1] || idxCol != CellId_P8[0])
					{
						return;
					}
					if (!flag2)
					{
						sheetView.Cells[CellId_G8[1], CellId_G8[0]].Value = true;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = false;
						}
					}
					else
					{
						sheetView.Cells[CellId_G8[1], CellId_G8[0]].Value = false;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
						}
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

	public void PreBindData(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			sheetView.Cells[GeneralFunctions.ArrayCellId("G_35")[1], GeneralFunctions.ArrayCellId("G_35")[0]].Value = ((TTinChung.tenCQTCapCuc.IndexOf("-") == 4) ? TTinChung.tenCQTCapCuc.Substring(6) : TTinChung.tenCQTCapCuc);
			sheetView.Cells[GeneralFunctions.ArrayCellId("F_35")[1], GeneralFunctions.ArrayCellId("F_35")[0]].Value = TTinChung.maCQTCapCuc;
			sheetView.Cells[GeneralFunctions.ArrayCellId("T_35")[1], GeneralFunctions.ArrayCellId("T_35")[0]].Value = TTinChung.tenCQTQuanLy;
			sheetView.Cells[GeneralFunctions.ArrayCellId("S_35")[1], GeneralFunctions.ArrayCellId("S_35")[0]].Value = TTinChung.maCQTCapCuc;
			SheetView sheetView2 = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView2.Cells[CellId_G8[1], CellId_G8[0]].Value);
			bool flag2 = !flag;
			bool flag3 = flag || flag2;
			if (!flag)
			{
				sheetView2.Cells[CellId_P8[1], CellId_P8[0]].Value = true;
				for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
				{
					sheetView2.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
					sheetView2.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
				}
			}
			else
			{
				sheetView2.Cells[CellId_P8[1], CellId_P8[0]].Value = false;
				for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
				{
					sheetView2.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
				}
			}
			if (!flag2)
			{
				sheetView2.Cells[CellId_G8[1], CellId_G8[0]].Value = true;
				for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
				{
				}
			}
			else
			{
				sheetView2.Cells[CellId_G8[1], CellId_G8[0]].Value = false;
				for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
				{
				}
			}
			for (int i = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart; i < hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[0].RowsValues.Count; i++)
			{
				string value = hTKKCMCSheetInfo.Tables[0].RowsValues[i - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[5];
				if ("VN".Equals(value))
				{
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				}
				bool flag4 = hTKKCMCSheetInfo.Tables[0].RowsValues[i - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[35].ToBoolOrDefault();
				if ("VN".Equals(value) && !flag4)
				{
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = false;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = true;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = true;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = true;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = false;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = false;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Locked = false;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = 1;
			num = (int)iRow;
			reloadGiayToTuyThan(winFpSpread, iCol, iRow);
			if (isChangeByForm && iRow >= 43 && (iCol == HTKKFile.ColumnLettersToIndex("M") || iCol == HTKKFile.ColumnLettersToIndex("AF")))
			{
				reloadDiaChi(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("M"))
			{
				string text = sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
				if ("VN".Equals(text))
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Value = 0;
				}
				else
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Value = 0;
				}
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("M") || iCol == HTKKFile.ColumnLettersToIndex("O"))
			{
				string text = sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
				bool flag = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Value.ToBoolOrDefault();
				if ("VN".Equals(text) && !flag)
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Value = "";
				}
				else
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("T") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("U") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Value = "";
				}
			}
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
			isChangeByForm = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void reloadDiaChi(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == HTKKFile.ColumnLettersToIndex("M"))
			{
				string text = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
				if (!HTKKFormater.isNullString(text) && text != "VN")
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("X") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Y") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Z") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AA") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AB") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AC") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AD") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AE") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AF") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AG") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AH") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AI") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AK") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AL") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AM") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AO") - 1].Value = "";
				}
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("AF"))
			{
				string text = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AE") - 1].Text;
				if (!HTKKFormater.isNullString(text) && text != "VN")
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("X") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Y") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Z") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AA") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AB") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AC") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AD") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AG") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AH") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AI") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AJ") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AK") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AL") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AM") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AN") - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AO") - 1].Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		isChangeByForm = false;
	}

	public void reloadGiayToTuyThan(FpSpread winFpSpread, long iCol, long iRow)
	{
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_P8[1], CellId_P8[0]].Value);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				if (flag)
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
				}
				else
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
				}
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[5];
				if ("VN".Equals(value))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				}
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				int num = HTKKFile.ColumnLettersToIndex("AC");
				int num2 = HTKKFile.ColumnLettersToIndex("AA");
				int num3 = HTKKFile.ColumnLettersToIndex("Y");
				int num4 = HTKKFile.ColumnLettersToIndex("AL");
				int num5 = HTKKFile.ColumnLettersToIndex("AJ");
				int num6 = HTKKFile.ColumnLettersToIndex("AH");
				string text = array[15];
				string text2 = array[17];
				string text3 = array[19];
				if (text3.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num - 1].Text = "";
				}
				if (text2.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num2 - 1].Text = "";
				}
				if (text.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num3 - 1].Text = "";
				}
				string text4 = array[24];
				string text5 = array[26];
				string text6 = array[28];
				if (text6.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num4 - 1].Text = "";
				}
				if (text5.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num5 - 1].Text = "";
				}
				if (text4.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, num6 - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		//IL_004d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Expected O, but got Unknown
		try
		{
			string text = "//CTieuTKhaiChinh/DSachCNDangkyMST";
			string text2 = "ct13";
			string text3 = "soDinhDanh";
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(text);
			if (val == null || val.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				if (val2.SelectSingleNode(text2) != null && val2.SelectSingleNode(text3) != null)
				{
					string innerText = val2.SelectSingleNode(text3).InnerText;
					string innerText2 = val2.SelectSingleNode(text2).InnerText;
					if (string.IsNullOrEmpty(innerText2) && !string.IsNullOrEmpty(innerText))
					{
						val2.SelectSingleNode(text2).InnerText = innerText;
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

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Q") - 1].Text;
				if (!string.IsNullOrEmpty(text))
				{
					sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("T") - 1].Text = "";
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
