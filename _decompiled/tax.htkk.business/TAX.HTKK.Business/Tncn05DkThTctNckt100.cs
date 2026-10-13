using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tncn05DkThTctNckt100
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
				string text;
				DateTime dateTime;
				string text6;
				string value3;
				string text2;
				string text5;
				string text3;
				DateTime dateTime2;
				string text4;
				string text7;
				string text8;
				string text9;
				string value4;
				string text10;
				string value5;
				string text11;
				string value6;
				string text12;
				string text13;
				string text14;
				string text15;
				string value7;
				string text16;
				string value8;
				string text17;
				if (!flag)
				{
					text = array5[2];
					dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
					if (string.IsNullOrEmpty(text))
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
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
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
					}
					string value2 = array5[7];
					text2 = array5[9];
					text3 = array5[11];
					text4 = array5[12];
					text5 = array5[10];
					if (bool.Parse(value2) && (string.IsNullOrEmpty(text2) || !text2.Equals("1010")))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(text2))
					{
						if (!string.IsNullOrEmpty(text2 + text5 + text3 + text4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
							value = 2;
						}
						else if (text2.Equals("1010") || text2.Equals("2080") || text2.Equals("1020"))
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1);
						}
						else
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
							value = 2;
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1);
					}
					if (!string.IsNullOrEmpty(text5))
					{
						if (text2.Equals("1010"))
						{
							if (text5.Length != 9 && text5.Length != 12)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_45")));
								value = 2;
							}
							else if (!HTKKFormater.IsNumber(text5))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
							}
						}
						else if (text2.Equals("2080"))
						{
							if (text5.Length != 12)
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_46")));
								value = 2;
							}
							else if (!HTKKFormater.IsNumber(text5))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65")));
								value = 2;
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
							}
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
						}
					}
					else if (!string.IsNullOrEmpty(text2 + text5 + text3 + text4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
					}
					string mst2 = array5[0];
					if (string.IsNullOrEmpty(mst2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThue(mst2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
						value = 2;
					}
					else
					{
						int num2 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[0].Equals(mst2));
						if (mst2.Equals(TTinChung.mst))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
							value = 2;
						}
						else if (num2 > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1);
						}
					}
					text6 = array5[3];
					value3 = array5[4];
					if (string.IsNullOrEmpty(value3))
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1);
					}
					else if (text6.Equals("1") || text6.Equals("2"))
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
						value = 2;
					}
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("O") - 1);
					dateTime2 = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy");
					if (string.IsNullOrEmpty(text3))
					{
						if (!string.IsNullOrEmpty(text2 + text5 + text3 + text4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1);
						}
					}
					else if (dateTime2 == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
						value = 2;
					}
					else if (dateTime2 > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48")));
						value = 2;
					}
					else if (!string.IsNullOrEmpty(text))
					{
						if (dateTime != DateTime.MinValue && dateTime >= dateTime2)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1);
					}
					if (string.IsNullOrEmpty(text4))
					{
						if (!string.IsNullOrEmpty(text2 + text5 + text3 + text4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
					}
					text7 = array5[13];
					text8 = array5[14];
					text9 = array5[15];
					value4 = array5[16];
					text10 = array5[17];
					value5 = array5[18];
					text11 = array5[19];
					value6 = array5[20];
					text12 = array5[21];
					if (!string.IsNullOrEmpty(text7 + text9 + text10 + text11 + text12) && text12 == "Việt Nam")
					{
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1);
						}
						if (string.IsNullOrEmpty(text9) && !text8.Equals("5050307"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(text8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
						}
						if (string.IsNullOrEmpty(text10))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(value4))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
						}
						if (string.IsNullOrEmpty(text11))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(value5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
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
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
						}
					}
					else if (!string.IsNullOrEmpty(text7 + text9 + text10 + text11 + text12) && text12 != "Việt Nam")
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
						if (string.IsNullOrEmpty(text7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
							value = 2;
						}
						if (string.IsNullOrEmpty(text12))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55")));
							value = 2;
						}
						else if (text2.Equals("1010") || text2.Equals("2080"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")));
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
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1);
						if (text12 == "Việt Nam" && text9.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
						}
						if (text12 == "Việt Nam" && text10.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
						}
						if (text12 == "Việt Nam" && text11.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
						}
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
					}
					text13 = array5[22];
					text14 = array5[23];
					text15 = array5[24];
					value7 = array5[25];
					text16 = array5[26];
					value8 = array5[27];
					text17 = array5[28];
					if (!string.IsNullOrEmpty(text13 + text15 + text16 + text17))
					{
						if (string.IsNullOrEmpty(text13))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1);
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
						else if (string.IsNullOrEmpty(text14))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1);
						}
						if (string.IsNullOrEmpty(text16))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_58")));
							value = 2;
						}
						else if (text16.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(value7))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
						}
						if (string.IsNullOrEmpty(text17))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
							value = 2;
						}
						else if (text17.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
							value = 2;
						}
						else if (string.IsNullOrEmpty(value8))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
					}
				}
				if (!flag)
				{
					continue;
				}
				string value9 = array5[1];
				if (string.IsNullOrEmpty(value9))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				}
				string text18 = array5[0];
				if (!string.IsNullOrEmpty(text18) && !ValidateBase.CheckMaSoThue(text18))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("C") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_16")));
					value = 2;
				}
				string text19 = array5[7];
				if ("TRUE".Equals(text19.ToUpper()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("O") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_64")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("O") - 1);
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
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				}
				text6 = array5[3];
				value3 = array5[4];
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41")));
					value = 2;
				}
				else if (text6.Equals("1") || text6.Equals("2"))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("K") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
					value = 2;
				}
				string value10 = array5[6];
				if (string.IsNullOrEmpty(value10))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_42")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				}
				text2 = array5[9];
				if (string.IsNullOrEmpty(text2))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_43")));
					value = 2;
				}
				else if (text2.Equals("1010") || text2.Equals("2080") || text2.Equals("1020"))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1);
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_27")));
					value = 2;
				}
				text5 = array5[10];
				if (!string.IsNullOrEmpty(text5))
				{
					if (text2.Equals("1010"))
					{
						if (text5.Length != 9 && text5.Length != 12)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_45")));
							value = 2;
						}
						else if (!HTKKFormater.IsNumber(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
						}
					}
					else if (text2.Equals("2080"))
					{
						if (text5.Length != 12)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_46")));
							value = 2;
						}
						else if (!HTKKFormater.IsNumber(text5))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_65")));
							value = 2;
						}
						else
						{
							HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
						}
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1);
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("T") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				text3 = array5[11];
				dateTime2 = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_47")));
					value = 2;
				}
				else if (dateTime2 == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
					value = 2;
				}
				else if (dateTime2 > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48")));
					value = 2;
				}
				else if (!string.IsNullOrEmpty(text))
				{
					if (dateTime != DateTime.MinValue && dateTime >= dateTime2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("U") - 1);
				}
				text4 = array5[12];
				if (string.IsNullOrEmpty(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				}
				text7 = array5[13];
				text8 = array5[14];
				text9 = array5[15];
				value4 = array5[16];
				text10 = array5[17];
				value5 = array5[18];
				text11 = array5[19];
				value6 = array5[20];
				text12 = array5[21];
				if (text12 == "Việt Nam")
				{
					if (string.IsNullOrEmpty(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1);
					}
					if (string.IsNullOrEmpty(text9) && !text8.Equals("5050307"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
						value = 2;
					}
					else if (text9.Equals("Khác"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(text8))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
					}
					if (string.IsNullOrEmpty(text10))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_53")));
						value = 2;
					}
					else if (text10.Equals("Khác"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
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
					else if (string.IsNullOrEmpty(value5))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
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
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
					}
				}
				else if (text12 != "Việt Nam")
				{
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Z") - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AD") - 1);
					HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1);
					if (string.IsNullOrEmpty(text7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("X") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text12))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55")));
						value = 2;
					}
					else if (text2.Equals("1010") || text2.Equals("2080"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AF") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")));
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
				value7 = array5[25];
				text16 = array5[26];
				value8 = array5[27];
				text17 = array5[28];
				if (!string.IsNullOrEmpty(text13 + text15 + text16 + text17))
				{
					if (string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1);
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
					else if (string.IsNullOrEmpty(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1);
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_58")));
						value = 2;
					}
					else if (text16.Equals("Khác"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(value7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
					}
					if (string.IsNullOrEmpty(text17))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
						value = 2;
					}
					else if (text17.Equals("Khác"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_20")));
						value = 2;
					}
					else if (string.IsNullOrEmpty(value8))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AG") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_56")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AI") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_59")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_58")));
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_57")));
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
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = false;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
						}
					}
					else
					{
						sheetView.Cells[CellId_P8[1], CellId_P8[0]].Value = false;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						}
					}
				}
				else if (idxRow == CellId_P8[1] && idxCol == CellId_P8[0])
				{
					if (!flag2)
					{
						sheetView.Cells[CellId_G8[1], CellId_G8[0]].Value = true;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = false;
						}
					}
					else
					{
						sheetView.Cells[CellId_G8[1], CellId_G8[0]].Value = false;
						for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
						{
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = false;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Locked = true;
							sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
						}
					}
				}
				if (idxRow == 3 && idxCol == 2)
				{
					ShowBangKe("Bang_Ke_05DK_TH.xls");
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
				}
			}
			else
			{
				sheetView2.Cells[CellId_P8[1], CellId_P8[0]].Value = false;
				for (int i = ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart; i < ToKhaiInfo.SheetInfos[0].Tables[0].Location.DataLocationTemplateRowStart + ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count; i++)
				{
					sheetView2.Cells[i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Locked = true;
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
				string text = hTKKCMCSheetInfo.Tables[0].RowsValues[i - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[9];
				sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
				Dictionary<string, string> noiCap;
				ComboBoxCellType comboBoxCellType;
				if (text.Equals("1010"))
				{
					noiCap = HTKKWinDictionaryDiaDanh.NoiCap;
					comboBoxCellType = new ComboBoxCellType();
					comboBoxCellType.Items = noiCap.Select((KeyValuePair<string, string> s) => s.Key).ToArray();
					comboBoxCellType.ItemData = noiCap.Select((KeyValuePair<string, string> s) => s.Value).ToArray();
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Text = hTKKCMCSheetInfo.Tables[0].RowsValues[i - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[12];
					continue;
				}
				if (text.Equals("2080"))
				{
					string[] noiCapCCCD = HTKKWinDictionaryDiaDanh.NoiCapCCCD;
					comboBoxCellType = new ComboBoxCellType();
					comboBoxCellType.Items = noiCapCCCD;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Value = "";
					continue;
				}
				noiCap = HTKKWinDictionaryDiaDanh.QuocGia;
				string[] array = new string[noiCap.Count];
				string[] array2 = new string[noiCap.Count];
				int num = 0;
				foreach (KeyValuePair<string, string> item in noiCap)
				{
					string[] array4;
					string[] array3 = (array4 = array);
					int num2 = num;
					nint num3 = num2;
					array3[num2] = array4[num3] + item.Value;
					string[] array5 = (array4 = array2);
					int num4 = num;
					num3 = num4;
					array5[num4] = array4[num3] + item.Key;
					num++;
				}
				comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = array;
				comboBoxCellType.ItemData = array2;
				sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
				sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Text = hTKKCMCSheetInfo.Tables[0].RowsValues[i - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[12];
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
			if (isChangeByForm && iRow >= 43 && iCol == HTKKFile.ColumnLettersToIndex("M"))
			{
				reloadDiaChi(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
			Dictionary<string, string> noiCap;
			if (iCol == HTKKFile.ColumnLettersToIndex("Q"))
			{
				string tenLoaiGiayTo = sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text;
				string text = sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text;
				if (HTKKFormater.isNullString(tenLoaiGiayTo))
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = "";
				}
				else
				{
					Dictionary<string, string> loaiGiayToTuyThan = HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan;
					if (HTKKFormater.isNullString(text))
					{
						text = loaiGiayToTuyThan.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == tenLoaiGiayTo).Key;
						sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = text;
					}
				}
				sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
				if (text == "1010")
				{
					noiCap = HTKKWinDictionaryDiaDanh.NoiCap;
					noiCap.Remove("Cả nước");
					ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
					comboBoxCellType.Items = noiCap.Select((KeyValuePair<string, string> s) => s.Key).ToArray();
					comboBoxCellType.ItemData = noiCap.Select((KeyValuePair<string, string> s) => s.Value).ToArray();
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
				}
				else if (text == "2080")
				{
					string[] noiCapCCCD = HTKKWinDictionaryDiaDanh.NoiCapCCCD;
					ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
					comboBoxCellType.Items = noiCapCCCD;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Value = "";
				}
				else if (!HTKKFormater.isNullString(tenLoaiGiayTo))
				{
					noiCap = HTKKWinDictionaryDiaDanh.QuocGia;
					string[] array = new string[noiCap.Count];
					string[] array2 = new string[noiCap.Count];
					int num2 = 0;
					foreach (KeyValuePair<string, string> item in noiCap)
					{
						string[] array4;
						string[] array3 = (array4 = array);
						int num3 = num2;
						nint num4 = num3;
						array3[num3] = array4[num4] + item.Value;
						string[] array5 = (array4 = array2);
						int num5 = num2;
						num4 = num5;
						array5[num5] = array4[num4] + item.Key;
						num2++;
					}
					ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
					comboBoxCellType.Items = array;
					comboBoxCellType.ItemData = array2;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
				}
			}
			int num6 = HTKKFile.ColumnLettersToIndex("AF");
			int num7 = HTKKFile.ColumnLettersToIndex("AD");
			int num8 = HTKKFile.ColumnLettersToIndex("AB");
			int num9 = HTKKFile.ColumnLettersToIndex("Z");
			int num10 = HTKKFile.ColumnLettersToIndex("X");
			int num11 = HTKKFile.ColumnLettersToIndex("AE");
			int num12 = HTKKFile.ColumnLettersToIndex("AC");
			int num13 = HTKKFile.ColumnLettersToIndex("AA");
			int num14 = HTKKFile.ColumnLettersToIndex("Y");
			int num15 = HTKKFile.ColumnLettersToIndex("AM");
			int num16 = HTKKFile.ColumnLettersToIndex("AK");
			int num17 = HTKKFile.ColumnLettersToIndex("AI");
			int num18 = HTKKFile.ColumnLettersToIndex("AG");
			int num19 = HTKKFile.ColumnLettersToIndex("AL");
			int num20 = HTKKFile.ColumnLettersToIndex("AJ");
			int num21 = HTKKFile.ColumnLettersToIndex("AH");
			int num22 = HTKKFile.ColumnLettersToIndex("V");
			int num23 = HTKKFile.ColumnLettersToIndex("AP");
			int num24 = HTKKFile.ColumnLettersToIndex("R");
			string strValueQT = "";
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			string text6 = "";
			string text7 = "";
			string text8 = "";
			string strNoiCap = "";
			string text9 = "";
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_G8[1], CellId_G8[0]].Value);
			strValueQT = sheetView.Cells[(int)iRow - 1, num6 - 1].Text;
			text2 = sheetView.Cells[(int)iRow - 1, num11 - 1].Text;
			if (iCol == num7)
			{
				Dictionary<string, string> quocGia = HTKKWinDictionaryDiaDanh.QuocGia;
				if (HTKKFormater.isNullString(text2))
				{
					text2 = quocGia.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == strValueQT).Key;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text2;
				}
				text3 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text4 = sheetView.Cells[(int)iRow - 1, num12 - 1].Text;
				text5 = sheetView.Cells[(int)iRow - 1, num8 - 1].Text;
				text6 = sheetView.Cells[(int)iRow - 1, num13 - 1].Text;
				text7 = sheetView.Cells[(int)iRow - 1, num9 - 1].Text;
				text8 = sheetView.Cells[(int)iRow - 1, num14 - 1].Text;
				if (text2 == "VN" && flag)
				{
					sheetView.Cells[(int)iRow - 1, num15 - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, num19 - 1].Text = text4;
					sheetView.Cells[(int)iRow - 1, num16 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num20 - 1].Text = text6;
					sheetView.Cells[(int)iRow - 1, num17 - 1].Text = text7;
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = text8;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num15 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num19 - 1].Text = "";
				}
			}
			if (iCol == num8)
			{
				Dictionary<string, string> quocGia = HTKKWinDictionaryDiaDanh.QuocGia;
				if (HTKKFormater.isNullString(text2))
				{
					text2 = quocGia.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == strValueQT).Key;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text2;
				}
				text5 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text6 = sheetView.Cells[(int)iRow - 1, num13 - 1].Text;
				text7 = sheetView.Cells[(int)iRow - 1, num9 - 1].Text;
				text8 = sheetView.Cells[(int)iRow - 1, num14 - 1].Text;
				if (text2 == "VN" && flag)
				{
					sheetView.Cells[(int)iRow - 1, num16 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num20 - 1].Text = text6;
					sheetView.Cells[(int)iRow - 1, num17 - 1].Text = text7;
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = text8;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num16 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num20 - 1].Text = "";
				}
			}
			if (iCol == num9)
			{
				Dictionary<string, string> quocGia = HTKKWinDictionaryDiaDanh.QuocGia;
				if (HTKKFormater.isNullString(text2))
				{
					text2 = quocGia.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == strValueQT).Key;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text2;
				}
				text7 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text8 = sheetView.Cells[(int)iRow - 1, num14 - 1].Text;
				if (text2 == "VN" && flag)
				{
					sheetView.Cells[(int)iRow - 1, num17 - 1].Text = text7;
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = text8;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num17 - 1].Text = "";
					sheetView.Cells[(int)iRow - 1, num21 - 1].Text = "";
				}
			}
			if (iCol == num10)
			{
				Dictionary<string, string> quocGia = HTKKWinDictionaryDiaDanh.QuocGia;
				if (HTKKFormater.isNullString(text2))
				{
					text2 = quocGia.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == strValueQT).Key;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text2;
				}
				text3 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				if (text2 == "VN" && flag)
				{
					sheetView.Cells[(int)iRow - 1, num18 - 1].Text = text3;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, num18 - 1].Text = "";
				}
			}
			if (iCol != num22)
			{
				return;
			}
			strNoiCap = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
			text9 = sheetView.Cells[(int)iRow - 1, num24 - 1].Text;
			Dictionary<string, string> noiCap2 = HTKKWinDictionaryDiaDanh.NoiCap;
			if (HTKKFormater.isNullString(text9))
			{
				text9 = noiCap2.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == strNoiCap).Key;
				sheetView.Cells[(int)iRow - 1, num24 - 1].Text = text9;
			}
			noiCap = null;
			string text10 = "";
			if (text9.Equals("1010"))
			{
				noiCap = HTKKWinDictionaryDiaDanh.NoiCap;
			}
			else if (text9.Equals("2080"))
			{
				text10 = "";
			}
			else
			{
				noiCap = HTKKWinDictionaryDiaDanh.QuocGia;
			}
			if (noiCap != null && noiCap.Count() > 0)
			{
				foreach (KeyValuePair<string, string> item2 in noiCap)
				{
					if (text9.Equals("1010"))
					{
						if (item2.Key == strNoiCap)
						{
							text10 = item2.Value;
							break;
						}
					}
					else if (item2.Value == strNoiCap)
					{
						text10 = item2.Key;
						break;
					}
				}
			}
			sheetView.Cells[(int)iRow - 1, num23 - 1].Text = text10;
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
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		isChangeByForm = false;
	}

	public void reloadGiayToTuyThan(FpSpread winFpSpread, long iCol, long iRow)
	{
		SheetView sheetView = winFpSpread.Sheets[0];
		string text = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text;
		if (iCol != HTKKFile.ColumnLettersToIndex("M"))
		{
			return;
		}
		Dictionary<string, string> loaiGiayToTuyThan = HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan;
		ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
		Dictionary<string, string> dictionary = loaiGiayToTuyThan;
		string text2 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
		string tenQuocGia = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Text;
		Dictionary<string, string> quocGia = HTKKWinDictionaryDiaDanh.QuocGia;
		if (HTKKFormater.isNullString(text2))
		{
			text2 = quocGia.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == tenQuocGia).Key;
			sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("L") - 1].Text = text2;
		}
		if (!HTKKFormater.isNullString(text2) && text2 == "VN")
		{
			dictionary = loaiGiayToTuyThan.Where((KeyValuePair<string, string> p) => p.Key.Equals("1010") || p.Key.Equals("2080")).ToDictionary((KeyValuePair<string, string> p) => p.Value, (KeyValuePair<string, string> p) => p.Key);
			comboBoxCellType.Items = dictionary.Select((KeyValuePair<string, string> s) => s.Key).ToArray();
			comboBoxCellType.ItemData = dictionary.Select((KeyValuePair<string, string> s) => s.Value).ToArray();
			sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].CellType = comboBoxCellType;
			sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = new ComboBoxCellType();
			if (text.Equals("2080"))
			{
				string[] noiCapCCCD = HTKKWinDictionaryDiaDanh.NoiCapCCCD;
				ComboBoxCellType comboBoxCellType2 = new ComboBoxCellType();
				comboBoxCellType2.Items = noiCapCCCD;
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType2;
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AP") - 1].Value = "";
			}
		}
		else
		{
			if (HTKKFormater.isNullString(text2) || !(text2 != "VN"))
			{
				return;
			}
			dictionary = loaiGiayToTuyThan.Where((KeyValuePair<string, string> p) => p.Key.Equals("1020")).ToDictionary((KeyValuePair<string, string> p) => p.Value, (KeyValuePair<string, string> p) => p.Key);
			if (dictionary != null)
			{
				comboBoxCellType.Items = dictionary.Select((KeyValuePair<string, string> s) => s.Key).ToArray();
				comboBoxCellType.ItemData = dictionary.Select((KeyValuePair<string, string> s) => s.Value).ToArray();
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].CellType = comboBoxCellType;
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Value = dictionary.First().Key;
				sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Text = dictionary.First().Value;
			}
			loaiGiayToTuyThan = HTKKWinDictionaryDiaDanh.QuocGia;
			string[] array = new string[loaiGiayToTuyThan.Count];
			string[] array2 = new string[loaiGiayToTuyThan.Count];
			int num = 0;
			foreach (KeyValuePair<string, string> item in loaiGiayToTuyThan)
			{
				string[] array4;
				string[] array3 = (array4 = array);
				int num2 = num;
				nint num3 = num2;
				array3[num2] = array4[num3] + item.Value;
				string[] array5 = (array4 = array2);
				int num4 = num;
				num3 = num4;
				array5[num4] = array4[num3] + item.Key;
				num++;
			}
			comboBoxCellType = new ComboBoxCellType();
			comboBoxCellType.Items = array;
			comboBoxCellType.ItemData = array2;
			sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
			sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("V") - 1].CellType = comboBoxCellType;
		}
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
				reloadGiayToTuyThan(winFpSpread, HTKKFile.ColumnLettersToIndex("M"), hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i);
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, HTKKFile.ColumnLettersToIndex("Q"));
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, HTKKFile.ColumnLettersToIndex("V"));
				if (flag)
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("D") - 1].Text = "";
				}
				else
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
