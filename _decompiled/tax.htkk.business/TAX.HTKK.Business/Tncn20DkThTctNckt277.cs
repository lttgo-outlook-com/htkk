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
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class Tncn20DkThTctNckt277
{
	private const string VIETNAM = "VN";

	private const string CMND = "1010";

	private const string CCCD = "2080";

	private const string ho_chieu = "1020";

	private const string TEN_VN = "Việt Nam";

	private const string KHAC = "Khác";

	private bool isChangeByForm = false;

	private string CellID_K4 = "K_4";

	private string CellID_O4 = "O_4";

	private string CellID_G8 = "G_8";

	private string CellID_G10 = "G_10";

	private int idxColC = HTKKFile.ColumnLettersToIndex("C");

	private int idxColD = HTKKFile.ColumnLettersToIndex("D");

	private int idxColG = HTKKFile.ColumnLettersToIndex("G");

	private int idxColH = HTKKFile.ColumnLettersToIndex("H");

	private int idxColI = HTKKFile.ColumnLettersToIndex("I");

	private int idxColJ = HTKKFile.ColumnLettersToIndex("J");

	private int idxColK = HTKKFile.ColumnLettersToIndex("K");

	private int idxColM = HTKKFile.ColumnLettersToIndex("M");

	private int idxColL = HTKKFile.ColumnLettersToIndex("L");

	private int idxColN = HTKKFile.ColumnLettersToIndex("N");

	private int idxColO = HTKKFile.ColumnLettersToIndex("O");

	private int idxColQ = HTKKFile.ColumnLettersToIndex("Q");

	private int idxColR = HTKKFile.ColumnLettersToIndex("R");

	private int idxColS = HTKKFile.ColumnLettersToIndex("S");

	private int idxColT = HTKKFile.ColumnLettersToIndex("T");

	private int idxColU = HTKKFile.ColumnLettersToIndex("U");

	private int idxColV = HTKKFile.ColumnLettersToIndex("V");

	private int idxColW = HTKKFile.ColumnLettersToIndex("W");

	private int idxColX = HTKKFile.ColumnLettersToIndex("X");

	private int idxColY = HTKKFile.ColumnLettersToIndex("Y");

	private int idxColZ = HTKKFile.ColumnLettersToIndex("Z");

	private int idxColAA = HTKKFile.ColumnLettersToIndex("AA");

	private int idxColAB = HTKKFile.ColumnLettersToIndex("AB");

	private int idxColAC = HTKKFile.ColumnLettersToIndex("AC");

	private int idxColAD = HTKKFile.ColumnLettersToIndex("AD");

	private int idxColAE = HTKKFile.ColumnLettersToIndex("AE");

	private int idxColAF = HTKKFile.ColumnLettersToIndex("AF");

	private int idxColAG = HTKKFile.ColumnLettersToIndex("AG");

	private int idxColAH = HTKKFile.ColumnLettersToIndex("AH");

	private int idxColAI = HTKKFile.ColumnLettersToIndex("AI");

	private int idxColAJ = HTKKFile.ColumnLettersToIndex("AJ");

	private int idxColAL = HTKKFile.ColumnLettersToIndex("AL");

	private int idxColAM = HTKKFile.ColumnLettersToIndex("AM");

	private int idxTableColB_I = 0;

	private int idxTableColC_I = 1;

	private int idxTableColD_I = 2;

	private int idxTableColG_I = 3;

	private int idxTableColH_I = 4;

	private int idxTableColI_I = 5;

	private int idxTableColJ_I = 6;

	private int idxTableColK_I = 7;

	private int idxTableColM_I = 8;

	private int idxTableColN_I = 9;

	private int idxTableColO_I = 10;

	private int idxTableColQ_I = 11;

	private int idxTableColR_I = 12;

	private int idxTableColS_I = 13;

	private int idxTableColT_I = 14;

	private int idxTableColU_I = 15;

	private int idxTableColV_I = 16;

	private int idxTableColW_I = 17;

	private int idxTableColX_I = 18;

	private int idxTableColY_I = 19;

	private int idxTableColZ_I = 20;

	private int idxTableColAA_I = 21;

	private int idxTableColAB_I = 22;

	private int idxTableColAC_I = 23;

	private int idxTableColAD_I = 24;

	private int idxTableColAE_I = 25;

	private int idxTableColAF_I = 26;

	private int idxTableColAG_I = 27;

	private int idxTableColAH_I = 28;

	private int idxTableColAI_I = 29;

	private int idxTableColAJ_I = 30;

	private int idxTableColAK_I = 31;

	private int[] CellId_K4 => GeneralFunctions.ArrayCellId(CellID_K4);

	private int[] CellId_O4 => GeneralFunctions.ArrayCellId(CellID_O4);

	private int[] CellId_G8 => GeneralFunctions.ArrayCellId(CellID_G8);

	private int[] CellId_G10 => GeneralFunctions.ArrayCellId(CellID_G10);

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[18, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[19, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[27, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[28, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[29, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[30, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text18 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text19 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text20 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[33, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text21 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[34, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text22 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text23 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text24 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text25 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text26 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text27 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[40, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text28 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[41, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text29 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[42, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text30 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[43, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text31 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[44, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text32 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[45, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text33 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text34 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text35 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[48, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text36 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[49, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text37 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[50, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text38 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[51, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text39 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text40 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[53, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text41 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[54, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text42 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[55, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text43 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[56, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text44 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[57, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text45 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[58, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text46 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[59, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text47 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[60, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text48 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[61, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text49 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[62, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text50 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[63, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text51 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[64, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text52 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[65, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text53 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[66, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text54 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[67, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text55 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[68, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text56 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[69, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text57 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[70, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text58 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[71, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			bool flag = sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value.ToBoolOrDefault();
			string mst = TTinChung.mst;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count && i != 1; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int num = location.DataLocationTemplateRowStart - 1;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int j = 0; j < rowsValues.Count; j++)
				{
					string[] array = rowsValues[j].RowValue.ToArray();
					HTKKFile.RemoveValidate(sheetView, num + j, idxColC - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColD - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColG - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColH - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColL - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColM - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColN - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColQ - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColR - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColS - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColT - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColU - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColV - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColW - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColAL - 1);
					HTKKFile.RemoveValidate(sheetView, num + j, idxColAM - 1);
					if (i == 0)
					{
						HTKKFile.RemoveValidate(sheetView, num + j, idxColI - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColJ - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColK - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColO - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColX - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColY - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColZ - 1);
						HTKKFile.RemoveValidate(sheetView, num + j, idxColAA - 1);
					}
					bool flag2 = true;
					for (int k = 1; k < array.Length; k++)
					{
						if (!array[k].Equals(rowsInfo.Cells[k].DefaultValue))
						{
							flag2 = false;
							break;
						}
					}
					if (flag2 || i != 0)
					{
						continue;
					}
					string value2 = array[1];
					if (flag)
					{
						if (string.IsNullOrEmpty(value2))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColC - 1, text);
							value = 2;
						}
					}
					else if (!string.IsNullOrEmpty(value2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColC - 1, text43);
						value = 2;
					}
					string mst2 = array[2];
					if (flag)
					{
						if (!string.IsNullOrEmpty(mst2))
						{
							if (!ValidateBase.CheckMaSoThueCCCD(mst2))
							{
								HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text5);
								value = 2;
							}
							else
							{
								int num2 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[0].Equals(mst2));
								if (mst2.Replace("-", "").Equals(TTinChung.mst))
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text51);
									value = 2;
								}
								else if (num2 > 1)
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text52);
									value = 2;
								}
							}
						}
					}
					else if (string.IsNullOrEmpty(mst2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text44);
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThueCCCD(mst2))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text45);
						value = 2;
					}
					else
					{
						int num2 = hTKKCMCTableInfo.RowsValues.Count((HTKKCMCRowValue p) => p.RowValue[2].Equals(mst2));
						if (mst2.Replace("-", "").Equals(TTinChung.mst))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text51);
							value = 2;
						}
						else if (num2 > 1)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColD - 1, text52);
							value = 2;
						}
					}
					string text59 = array[3];
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(text59, "dd/MM/yyyy");
					if (string.IsNullOrEmpty(text59))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColG - 1, text2);
						value = 2;
					}
					else if (dateTime == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColG - 1, text3);
						value = 2;
					}
					else if (dateTime > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColG - 1, text4);
						value = 2;
					}
					string value3 = array[4];
					if (string.IsNullOrEmpty(value3))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColH - 1, text6);
						value = 2;
					}
					string obj = array[5];
					string text60 = array[7];
					string text61 = array[22];
					if (string.IsNullOrEmpty(text60))
					{
						if (!obj.ToBoolOrDefault() && text61.Equals("VN"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColK - 1, text8);
							value = 2;
						}
					}
					else if (text60.Length != 12)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColK - 1, text47);
						value = 2;
					}
					else if (!HTKKFormater.IsNumber(text60))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColK - 1, text48);
						value = 2;
					}
					string value4 = array[8];
					if ((text61 != "VN" || (text61 == "VN" && obj.ToBoolOrDefault())) && string.IsNullOrEmpty(value4))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColM - 1, text54);
						value = 2;
					}
					string text62 = array[9];
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text62, "dd/MM/yyyy");
					if (string.IsNullOrEmpty(text62) && (!text61.Equals("VN") || (text61 == "VN" && obj.ToBoolOrDefault())))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColN - 1, text9);
						value = 2;
					}
					else if (dateTime2 == DateTime.MinValue && (!text61.Equals("VN") || (text61 == "VN" && obj.ToBoolOrDefault())))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColN - 1, text10);
						value = 2;
					}
					else if (dateTime2 > DateTime.Now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColN - 1, text11);
						value = 2;
					}
					else if (!string.IsNullOrEmpty(text59) && dateTime2 != DateTime.MinValue && dateTime != DateTime.MinValue && dateTime >= dateTime2)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColN - 1, text50);
						value = 2;
					}
					string value5 = array[10];
					if (string.IsNullOrEmpty(value5) && (!text61.Equals("VN") || (text61 == "VN" && obj.ToBoolOrDefault())))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColO - 1, text12);
						value = 2;
					}
					string text63 = array[14];
					string text64 = array[27];
					string text65 = array[13];
					string maHuyen = array[26];
					string text66 = array[12];
					string text67 = array[25];
					string text68 = array[11];
					if (flag || !string.IsNullOrEmpty(text63 + text65 + text68))
					{
						if (string.IsNullOrEmpty(text63))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColT - 1, text13);
							value = 2;
						}
						if (string.IsNullOrEmpty(text65))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColS - 1, text16);
							value = 2;
						}
						else if (string.IsNullOrEmpty(text64) && !text65.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColS - 1, text17);
							value = 2;
						}
						if (string.IsNullOrEmpty(text68))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColQ - 1, text14);
							value = 2;
						}
						else if (string.IsNullOrEmpty(text67) && !text68.Equals("Khác"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColQ - 1, text19);
							value = 2;
						}
						if (!GeneralFunctions.ValidateDiaBan2Cap(text67, maHuyen, text64))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColQ - 1, text53);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColS - 1, text53);
							value = 2;
						}
					}
					string text69 = array[18];
					string text70 = array[30];
					string text71 = array[17];
					string maHuyen2 = array[29];
					string text72 = array[16];
					string text73 = array[28];
					string text74 = array[15];
					if (flag || !string.IsNullOrEmpty(text69 + text71 + text74))
					{
						if (string.IsNullOrEmpty(text69))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColX - 1, text20);
							value = 2;
						}
						if (string.IsNullOrEmpty(text71))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColW - 1, text23);
							value = 2;
						}
						else if (text71.Trim() != "Khác" && string.IsNullOrEmpty(text70))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColW - 1, text17);
							value = 2;
						}
						if (string.IsNullOrEmpty(text74))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColU - 1, text21);
							value = 2;
						}
						else if (text74.Trim() != "Khác" && string.IsNullOrEmpty(text73))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColU - 1, text19);
							value = 2;
						}
						if (!GeneralFunctions.ValidateDiaBan2Cap(text73, maHuyen2, text70))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColU - 1, text53);
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColW - 1, text53);
							value = 2;
						}
					}
					string value6 = array[19];
					if (flag && string.IsNullOrEmpty(value6))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColY - 1, text24);
						value = 2;
					}
					string value7 = array[20];
					if (flag && string.IsNullOrEmpty(value7))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColZ - 1, text25);
						value = 2;
					}
					string text75 = array[21];
					if (string.IsNullOrEmpty(text75))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAA - 1, text26);
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThueCCCD(text75))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAA - 1, text27);
						value = 2;
					}
					else if (text75.Replace("-", "").Equals(TTinChung.mst))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAA - 1, text51);
						value = 2;
					}
					string text76 = array[32];
					DateTime dtDate;
					if (string.IsNullOrEmpty(text76))
					{
						if (flag)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAL - 1, text55);
							value = 2;
						}
					}
					else if (!HTKKFormater.IsDateTime("01/" + text76, "dd/MM/yyyy", out dtDate))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAL - 1, text56);
						value = 2;
					}
					string text77 = array[33];
					if (!string.IsNullOrEmpty(text77) && !HTKKFormater.IsDateTime("01/" + text77, "dd/MM/yyyy", out dtDate))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAM - 1, text57);
						value = 2;
					}
					if (!string.IsNullOrEmpty(text76) && !string.IsNullOrEmpty(text77) && HTKKFormater.IsDateTime("01/" + text76, "dd/MM/yyyy", out var dtDate2) && HTKKFormater.IsDateTime("01/" + text77, "dd/MM/yyyy", out var dtDate3) && (dtDate2.Year > dtDate3.Year || (dtDate2.Year == dtDate3.Year && dtDate2.Month > dtDate3.Month)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAL - 1, text58);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + j, idxColAM - 1, text58);
						value = 2;
					}
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
				ShowBangKe("Bang_Ke_20TH_DangKyCapMa.xls");
			}
			else
			{
				if (winFpSpread.ActiveSheetIndex != 0)
				{
					return;
				}
				SheetView sheetView = winFpSpread.Sheets[0];
				bool flag = Convert.ToBoolean(sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value);
				bool flag2 = Convert.ToBoolean(sheetView.Cells[CellId_O4[1], CellId_O4[0]].Value);
				bool flag3 = flag || flag2;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					for (int j = 0; j < rowsValues.Count; j++)
					{
						if (idxRow == CellId_K4[1] && idxCol == CellId_K4[0])
						{
							if (!flag)
							{
								sheetView.Cells[CellId_O4[1], CellId_O4[0]].Value = true;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Text = "";
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Locked = false;
							}
							else
							{
								sheetView.Cells[CellId_O4[1], CellId_O4[0]].Value = false;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColC - 1].Text = "";
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColC - 1].Locked = false;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Text = "";
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Locked = true;
							}
						}
						else if (idxRow == CellId_O4[1] && idxCol == CellId_O4[0])
						{
							if (!flag2)
							{
								sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value = true;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Text = "";
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Locked = true;
							}
							else
							{
								sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value = false;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColC - 1].Locked = true;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColC - 1].Text = "";
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Locked = false;
								sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Text = "";
							}
						}
					}
					if (i == 0 && idxRow >= dataLocationTemplateRowStart - 1 && idxRow <= dataLocationTemplateRowStart + rowsValues.Count - 1 && idxCol == idxColH)
					{
						reloadGiayToTuyThan(winFpSpread, idxCol, idxRow, dataLocationTemplateRowStart - 1, dataLocationTemplateRowStart + rowsValues.Count - 1);
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value);
			bool flag2 = !flag;
			bool flag3 = flag || flag2;
			for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count && i != 1; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[i];
				HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
				int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				for (int j = 0; j < rowsValues.Count; j++)
				{
					if (!flag)
					{
						sheetView.Cells[CellId_O4[1], CellId_O4[0]].Value = true;
					}
					else
					{
						sheetView.Cells[CellId_O4[1], CellId_O4[0]].Value = false;
						sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColD - 1].Locked = true;
					}
					if (!flag2)
					{
						sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value = true;
						continue;
					}
					sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value = false;
					sheetView.Cells[dataLocationTemplateRowStart + j - 1, idxColC - 1].Locked = true;
				}
			}
			for (int j = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart; j < hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart + hTKKCMCSheetInfo.Tables[0].RowsValues.Count; j++)
			{
				sheetView.Cells[j - 1, idxColO - 1].Locked = false;
				string value = hTKKCMCSheetInfo.Tables[0].RowsValues[j - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[idxTableColAB_I];
				if ("VN".Equals(value))
				{
					sheetView.Cells[j - 1, idxColI - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[j - 1, idxColI - 1].Locked = true;
				}
				bool flag4 = hTKKCMCSheetInfo.Tables[0].RowsValues[j - hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart].RowValue[idxTableColI_I].ToBoolOrDefault();
				if ("VN".Equals(value) && !flag4)
				{
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Locked = false;
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				}
				else
				{
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Locked = true;
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[j - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			int num = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int j = 0; j < hTKKCMCTableInfo2.RowsValues.Count; j++)
			{
				sheetView.Rows[num + j].Visible = false;
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
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			reloadGiayToTuyThan(winFpSpread, iCol, iRow, dataLocationTemplateRowStart, dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count);
			if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count)
			{
				return;
			}
			if (isChangeByForm && iCol == idxColH)
			{
				reloadDiaChi(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, iRow, iCol);
			}
			if (iCol == idxColH)
			{
				string text = sheetView.Cells[num - 1, idxColAB - 1].Text;
				if ("VN".Equals(text))
				{
					sheetView.Cells[num - 1, idxColI - 1].Locked = false;
					sheetView.Cells[num - 1, idxColI - 1].Value = 0;
				}
				else
				{
					sheetView.Cells[num - 1, idxColI - 1].Locked = true;
					sheetView.Cells[num - 1, idxColI - 1].Value = 0;
				}
			}
			int num2 = idxColQ;
			int num3 = idxColR;
			int num4 = idxColS;
			int num5 = idxColT;
			int num6 = idxColAE;
			int num7 = idxColAF;
			int num8 = idxColAG;
			int num9 = idxColU;
			int num10 = idxColV;
			int num11 = idxColW;
			int num12 = idxColX;
			int num13 = idxColAH;
			int num14 = idxColAI;
			int num15 = idxColAJ;
			int num16 = idxColO;
			int num17 = idxColAD;
			int num18 = idxColAC;
			string text2 = "";
			string text3 = "";
			string text4 = "";
			string text5 = "";
			string text6 = "";
			string text7 = "";
			string text8 = "";
			string text9 = "";
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_K4[1], CellId_K4[0]].Value);
			if (iCol == num2)
			{
				text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text3 = sheetView.Cells[(int)iRow - 1, num6 - 1].Text;
				text4 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text;
				text5 = sheetView.Cells[(int)iRow - 1, num7 - 1].Text;
				text6 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text;
				text7 = sheetView.Cells[(int)iRow - 1, num8 - 1].Text;
			}
			if (iCol == num3)
			{
				text4 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text5 = sheetView.Cells[(int)iRow - 1, num7 - 1].Text;
				text6 = sheetView.Cells[(int)iRow - 1, num4 - 1].Text;
				text7 = sheetView.Cells[(int)iRow - 1, num8 - 1].Text;
				if (flag)
				{
					text2 = sheetView.Cells[(int)iRow - 1, num2 - 1].Text;
					text3 = sheetView.Cells[(int)iRow - 1, num6 - 1].Text;
					WinInsertCombox(sheetView, (int)iRow, num10, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[text3]);
					sheetView.Cells[(int)iRow - 1, num10 - 1].Text = text4;
					sheetView.Cells[(int)iRow - 1, num14 - 1].Text = text5;
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text6;
					sheetView.Cells[(int)iRow - 1, num15 - 1].Text = text7;
				}
			}
			if (iCol == num4)
			{
				text4 = sheetView.Cells[(int)iRow - 1, num3 - 1].Text;
				text5 = sheetView.Cells[(int)iRow - 1, num7 - 1].Text;
				WinInsertCombox(sheetView, (int)iRow, num11, HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen[text5]);
				text6 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				text7 = sheetView.Cells[(int)iRow - 1, num8 - 1].Text;
				if (flag)
				{
					sheetView.Cells[(int)iRow - 1, num11 - 1].Text = text6;
					sheetView.Cells[(int)iRow - 1, num15 - 1].Text = text7;
				}
			}
			if (iCol == num5)
			{
				text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				if (flag)
				{
					sheetView.Cells[(int)iRow - 1, num12 - 1].Text = text2;
				}
			}
			if (iCol == HTKKFile.ColumnLettersToIndex("H") || iCol == HTKKFile.ColumnLettersToIndex("I"))
			{
				string text = sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AB") - 1].Text;
				bool flag2 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("I") - 1].Value.ToBoolOrDefault();
				if ("VN".Equals(text) && !flag2)
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AD") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AD") - 1].Value = "";
				}
				else
				{
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Locked = true;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("K") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("N") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("O") - 1].Value = "";
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AD") - 1].Locked = false;
					sheetView.Cells[num - 1, HTKKFile.ColumnLettersToIndex("AD") - 1].Value = "";
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			if (iRow < dataLocationTemplateRowStart || iRow > dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count || iCol != idxColN)
			{
				return;
			}
			WinInsertCombox(sheetView, (int)iRow, idxColQ, HTKKWinDictionaryDiaDanh.TinhTP);
			WinInsertCombox(sheetView, (int)iRow, idxColR, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP["ALL"]);
			WinInsertCombox(sheetView, (int)iRow, idxColS, HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen["ALL"]);
			string text = "";
			string text2 = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
			foreach (KeyValuePair<string, string> quocGium in HTKKWinDictionaryDiaDanh.QuocGia)
			{
				if (quocGium.Value.Equals(text2))
				{
					text = quocGium.Key;
					break;
				}
			}
			if (text != "VN")
			{
				sheetView.Cells[(int)iRow - 1, idxColQ - 1].Locked = true;
				sheetView.Cells[(int)iRow - 1, idxColR - 1].Locked = true;
				sheetView.Cells[(int)iRow - 1, idxColS - 1].Locked = true;
				sheetView.Cells[(int)iRow - 1, idxColQ - 1].Value = "Khác";
				sheetView.Cells[(int)iRow - 1, idxColR - 1].Value = "Khác";
				sheetView.Cells[(int)iRow - 1, idxColS - 1].Value = "Khác";
				sheetView.Cells[(int)iRow - 1, idxColAC - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColAD - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColAE - 1].Value = "";
			}
			else
			{
				sheetView.Cells[(int)iRow - 1, idxColQ - 1].Locked = false;
				sheetView.Cells[(int)iRow - 1, idxColR - 1].Locked = false;
				sheetView.Cells[(int)iRow - 1, idxColS - 1].Locked = false;
				sheetView.Cells[(int)iRow - 1, idxColQ - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColR - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColS - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColAC - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColAD - 1].Value = "";
				sheetView.Cells[(int)iRow - 1, idxColAE - 1].Value = "";
			}
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
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
			if (iRow >= dataLocationTemplateRowStart && iRow <= dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count && iCol == idxColH)
			{
				string text = sheetView.Cells[(int)iRow - 1, idxColAB - 1].Text;
				if (!HTKKFormater.isNullString(text) && text != "VN")
				{
					sheetView.Cells[(int)iRow - 1, idxColQ - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColR - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColS - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColT - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColU - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColV - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColW - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColX - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColAE - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColAF - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColAG - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColAH - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColAI - 1].Value = "";
					sheetView.Cells[(int)iRow - 1, idxColAJ - 1].Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		isChangeByForm = false;
	}

	private void reloadGiayToTuyThan(FpSpread winFpSpread, long iCol, long iRow, int tuRow, int denRow)
	{
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			SheetView sheetView = winFpSpread.Sheets[0];
			bool flag = Convert.ToBoolean(sheetView.Cells[CellId_O4[1], CellId_O4[0]].Value);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, idxColO);
				CellChange_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i, idxColK);
				if (flag)
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColC - 1].Text = "";
				}
				else
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColD - 1].Text = "";
				}
				string[] array = hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray();
				string text = array[13];
				string text2 = array[12];
				string text3 = array[11];
				string text4 = array[17];
				string text5 = array[16];
				string text6 = array[15];
				if (text3.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColAE - 1].Text = "";
				}
				if (text2.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColAF - 1].Text = "";
				}
				if (text.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColAG - 1].Text = "";
				}
				if (text6.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColAH - 1].Text = "";
				}
				if (text5.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColAI - 1].Text = "";
				}
				if (text4.Trim().Equals("Khác"))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColAJ - 1].Text = "";
				}
				string value = hTKKCMCTableInfo.RowsValues[i].RowValue[idxTableColAB_I];
				if ("VN".Equals(value))
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColI - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + i - 1, idxColI - 1].Locked = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				sheetView.Rows[num + i].Visible = false;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
			int num2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string text = sheetView.Cells[num2 + i, HTKKFile.ColumnLettersToIndex("K") - 1].Text;
				if (!string.IsNullOrEmpty(text))
				{
					sheetView.Cells[num2 + i, HTKKFile.ColumnLettersToIndex("M") - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private static void WinInsertCombox(SheetView sView, int iRow, int iCol, string[] Values)
	{
		try
		{
			if (iCol >= 1 && iRow >= 1)
			{
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = Values;
				sView.Cells[iRow - 1, iCol - 1].CellType = comboBoxCellType;
				sView.Cells[iRow - 1, iCol - 1].Text = "";
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
			string text = "//CTieuTKhaiChinh/DSachNPTDaCoMST";
			string text2 = "CT6/ct6_soGiayTo";
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
}
