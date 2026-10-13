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

internal class Dnht01Nckt286
{
	private string cellId_SoDeNghiHoan = "H_6";

	private string cellId_HTTruocKTSau = "S_8";

	private string cellId_KTTruocHTSau = "S_10";

	private string cellId_LoaiHT = "L_12";

	private string cellId_CQTCuc = "L_14";

	private string cellId_CQTChiCuc = "L_14";

	private string cellId_LyDoHT = "J_23";

	private string cellId_MaLoaiHT = "BD_12";

	private string cellId_KeKhai_Thang = "E_17";

	private string cellId_KeKhai_Quy = "K_17";

	private string cellId_KeKhai_Tu = "G_19";

	private string cellId_KeKhai_Tu_Nam = "J_19";

	private string cellId_KeKhai_Den = "G_21";

	private string cellId_KeKhai_Den_Nam = "J_21";

	private string cellId_SoTienDeNghiHoan = "AA_70";

	private string cellId_TongTienNSNN_BangChu = "L_84";

	private string cellId_TongTienHoanTra_BangChu = "Z_86";

	private string cellId_TenChuDADT = "J_61";

	private string cellId_MSTDADT = "J_63";

	private string cellId_TenDADT = "J_57";

	private string cellId_NgayHoanThanhDADT = "AX_73";

	private string cellId_DiaChiDADT = "J_59";

	private string cellId_SoDADT = "J_67";

	private string cellId_NgayDADT = "J_69";

	private string cellId_CoQuanBHDADT = "J_71";

	private int idxRowNgayKy = 73;

	private int[] cellCQTCapCuc = GeneralFunctions.ArrayCellId("J_53");

	private int[] cellCQTMaCapCuc = GeneralFunctions.ArrayCellId("AX_53");

	private int[] cellCQTChiCuc = GeneralFunctions.ArrayCellId("AA_53");

	private int[] cellCQTMaChiCuc = GeneralFunctions.ArrayCellId("AY_53");

	private string maHoanXuatKhau = "0002";

	private string maHoanDuAnDT = "0039";

	private ComboBoxCellType comboQuy = new ComboBoxCellType();

	private ComboBoxCellType comboThang = new ComboBoxCellType();

	private int[] CellId_SoDeNghiHoan => GeneralFunctions.ArrayCellId(cellId_SoDeNghiHoan);

	private int[] CellId_HTTruocKTSau => GeneralFunctions.ArrayCellId(cellId_HTTruocKTSau);

	private int[] CellId_KTTruocHTSau => GeneralFunctions.ArrayCellId(cellId_KTTruocHTSau);

	private int[] CellId_LoaiHT => GeneralFunctions.ArrayCellId(cellId_LoaiHT);

	private int[] CellId_CQTCuc => GeneralFunctions.ArrayCellId(cellId_CQTCuc);

	private int[] CellId_CQTChiCuc => GeneralFunctions.ArrayCellId(cellId_CQTChiCuc);

	private int[] CellId_LyDoHT => GeneralFunctions.ArrayCellId(cellId_LyDoHT);

	private int[] CellId_MaLoaiHT => GeneralFunctions.ArrayCellId(cellId_MaLoaiHT);

	private int[] CellId_KeKhai_Thang => GeneralFunctions.ArrayCellId(cellId_KeKhai_Thang);

	private int[] CellId_KeKhai_Quy => GeneralFunctions.ArrayCellId(cellId_KeKhai_Quy);

	private int[] CellId_KeKhai_Tu => GeneralFunctions.ArrayCellId(cellId_KeKhai_Tu);

	private int[] CellId_KeKhai_Tu_Nam => GeneralFunctions.ArrayCellId(cellId_KeKhai_Tu_Nam);

	private int[] CellId_KeKhai_Den => GeneralFunctions.ArrayCellId(cellId_KeKhai_Den);

	private int[] CellId_KeKhai_Den_Nam => GeneralFunctions.ArrayCellId(cellId_KeKhai_Den_Nam);

	private int[] CellId_SoTienDeNghiHoan => GeneralFunctions.ArrayCellId(cellId_SoTienDeNghiHoan);

	private int[] CellId_TongTienNSNN_BangChu => GeneralFunctions.ArrayCellId(cellId_TongTienNSNN_BangChu);

	private int[] CellId_TongTienHoanTra_BangChu => GeneralFunctions.ArrayCellId(cellId_TongTienHoanTra_BangChu);

	private int[] CellId_TenChuDADT => GeneralFunctions.ArrayCellId(cellId_TenChuDADT);

	private int[] CellId_MSTDADT => GeneralFunctions.ArrayCellId(cellId_MSTDADT);

	private int[] CellId_TenDADT => GeneralFunctions.ArrayCellId(cellId_TenDADT);

	private int[] CellId_NgayHoanThanhDADT => GeneralFunctions.ArrayCellId(cellId_NgayHoanThanhDADT);

	private int[] CellId_DiaChiDADT => GeneralFunctions.ArrayCellId(cellId_DiaChiDADT);

	private int[] CellId_SoDADT => GeneralFunctions.ArrayCellId(cellId_SoDADT);

	private int[] CellId_NgayDADT => GeneralFunctions.ArrayCellId(cellId_NgayDADT);

	private int[] CellId_CoQuanBHDADT => GeneralFunctions.ArrayCellId(cellId_CoQuanBHDADT);

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			string kieuKyHoanThue = PeriodInfo.KieuKyHoanThue;
			try
			{
				sheetView.Cells[18, HTKKFile.ColumnLettersToIndex("G") - 1].Text = PeriodInfo.TuThang;
				sheetView.Cells[20, HTKKFile.ColumnLettersToIndex("J") - 1].Text = PeriodInfo.DenThang;
				if (kieuKyHoanThue == "1" || kieuKyHoanThue == "3")
				{
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("AY") - 1].Text = "M";
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("AZ") - 1].Text = "01/" + PeriodInfo.TuThang;
					DateTime obj = ConvertHelpers.ConvertToDateTime("01/" + PeriodInfo.TuThang, "dd/MM/yyyy").AddMonths(1).AddDays(-1.0);
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BA") - 1].Text = ConvertHelpers.ConvertDateTimeToStr(obj, "dd/MM/yyyy");
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BB") - 1].Text = PeriodInfo.TuThang;
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = PeriodInfo.TuThang;
				}
				if (kieuKyHoanThue == "4" || kieuKyHoanThue == "1")
				{
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = "M";
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = "01/" + PeriodInfo.DenThang;
					DateTime obj = ConvertHelpers.ConvertToDateTime("01/" + PeriodInfo.DenThang, "dd/MM/yyyy").AddMonths(1).AddDays(-1.0);
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = ConvertHelpers.ConvertDateTimeToStr(obj, "dd/MM/yyyy");
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BG") - 1].Text = PeriodInfo.DenThang;
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BH") - 1].Text = PeriodInfo.DenThang;
				}
				if (kieuKyHoanThue == "4" || kieuKyHoanThue == "2")
				{
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("AY") - 1].Text = "Q";
					DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(PeriodInfo.TuThang.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(PeriodInfo.TuThang.Split(new char[1] { '/' })[1]), 1, 1);
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("AZ") - 1].Text = ConvertHelpers.ConvertDateTimeToStr(ngayDauQuy, "dd/MM/yyyy");
					DateTime obj = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(PeriodInfo.TuThang.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(PeriodInfo.TuThang.Split(new char[1] { '/' })[1]), 1, 1);
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BA") - 1].Text = ConvertHelpers.ConvertDateTimeToStr(obj, "dd/MM/yyyy");
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BB") - 1].Text = ngayDauQuy.Month.ToString().PadLeft(2, '0') + "/" + ngayDauQuy.Year.ToString().PadLeft(4, '0');
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = obj.Month.ToString().PadLeft(2, '0') + "/" + obj.Year.ToString().PadLeft(4, '0');
				}
				if (kieuKyHoanThue == "3" || kieuKyHoanThue == "2")
				{
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = "Q";
					DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Split(new char[1] { '/' })[1]), 1, 1);
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = ConvertHelpers.ConvertDateTimeToStr(ngayDauQuy, "dd/MM/yyyy");
					DateTime obj = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(PeriodInfo.DenThang.Split(new char[1] { '/' })[1]), 1, 1);
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = ConvertHelpers.ConvertDateTimeToStr(obj, "dd/MM/yyyy");
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BG") - 1].Text = ngayDauQuy.Month.ToString().PadLeft(2, '0') + "/" + ngayDauQuy.Year.ToString().PadLeft(4, '0');
					sheetView.Cells[84, HTKKFile.ColumnLettersToIndex("BH") - 1].Text = obj.Month.ToString().PadLeft(2, '0') + "/" + obj.Year.ToString().PadLeft(4, '0');
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/chuyenKhoan"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tienMat"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/nguoiNhanTien"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/cmt-HoChieu"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/ngayCap"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tinh-ThanhPho_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tinh-ThanhPho_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo9) => hTKKCMCCellInfo9.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/noiNhanTien"));
				if (!PeriodInfo.TieuMucHachToanVal.Equals(maHoanXuatKhau) && !PeriodInfo.TieuMucHachToanVal.Equals(maHoanDuAnDT))
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].BackColor = Color.White;
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].BackColor = Color.White;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].BackColor = Color.White;
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].BackColor = Color.White;
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].BackColor = Color.White;
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].BackColor = Color.White;
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].BackColor = Color.White;
				}
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			}
			if (!PeriodInfo.NganhNgheKinhDoanh.Equals(""))
			{
				sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("L") - 1].Text = PeriodInfo.NganhNgheKinhDoanh;
				sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = PeriodInfo.TieuMucHachToanVal;
			}
			string text = sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("BD") - 1].Text;
			if (!string.IsNullOrEmpty(text))
			{
				string[] items = GetDM("..\\InterfaceIni\\Catalogue_Ly_Do_Hoan_Thue.xml", text).ToArray();
				ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
				comboBoxCellType.Items = items;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].CellType = comboBoxCellType;
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].Text = ToKhaiInfo.SheetInfos[0].StaticCells[7].Value;
			}
			string text2 = sheetView.Cells[11, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
			int num = HTKKFile.ColumnLettersToIndex("J");
			int num2 = HTKKFile.ColumnLettersToIndex("R");
			int num3 = HTKKFile.ColumnLettersToIndex("AX");
			int num4 = HTKKFile.ColumnLettersToIndex("AY");
			int num5 = HTKKFile.ColumnLettersToIndex("AZ");
			Cell cell = sheetView.Cells[51, num - 1];
			Cell cell2 = sheetView.Cells[51, num3 - 1];
			Cell cell3 = sheetView.Cells[51, num5 - 1];
			Cell cell4 = sheetView.Cells[51, num4 - 1];
			Cell cell5 = sheetView.Cells[52, num5 - 1];
			cell.Text = PeriodInfo.TenTinh;
			cell2.Text = PeriodInfo.MaTinh;
			cell3.Value = PeriodInfo.TenXa;
			cell4.Value = PeriodInfo.MaXa;
			string nameTable = "01_DNHT_II";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num6 = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			int num7 = HTKKFile.ColumnLettersToIndex("AM");
			for (int num8 = dataLocationTemplateRowStart; num8 <= num6; num8++)
			{
				int num9 = HTKKFile.ColumnLettersToIndex("AH");
				int num10 = HTKKFile.ColumnLettersToIndex("BA");
				string text3 = sheetView.Cells[num8 - 1, num10 - 1].Text;
				string text4 = sheetView.Cells[num8 - 1, num7 - 1].Text;
				if (!string.IsNullOrEmpty(text4))
				{
					sheetView.Cells[num8 - 1, num7 - 1].Text = text4.Replace("Q", "").Replace("/", "");
				}
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.MaskChar = '.';
				switch (text3)
				{
				case "Y":
					maskCellType.Mask = "0000";
					break;
				case "Q":
					maskCellType.Mask = "0/0000";
					break;
				case "M":
					maskCellType.Mask = "00/0000";
					break;
				case "D":
					maskCellType.Mask = "00/00/0000";
					break;
				}
				sheetView.Cells[num8 - 1, num7 - 1].ResetCellType();
				sheetView.Cells[num8 - 1, num7 - 1].CellType = maskCellType;
			}
			int[] array = GeneralFunctions.ArrayCellId("L_14");
			int[] array2 = GeneralFunctions.ArrayCellId("AG_14");
			int[] array3 = GeneralFunctions.ArrayCellId("BD_14");
			int[] array4 = GeneralFunctions.ArrayCellId("BD_12");
			sheetView.Cells[array[1], array[0]].Locked = true;
			string text5 = sheetView.Cells[array4[1], array4[0]].Text.Trim();
			string text6 = TTinChung.tenCQTCapCuc;
			string tenCQTQuanLy = PeriodInfo.tenCQTQuanLy;
			if (text6.IndexOf("-") == 4)
			{
				text6 = text6.Substring(6).Trim();
			}
			sheetView.Cells[array[1], array[0]].Text = PeriodInfo.tenCQTQuanLy;
			sheetView.Cells[array3[1], array3[0]].Text = PeriodInfo.maCQTQuanLy;
			Row row = sheetView.Rows[array[1]];
			float preferredHeight = row.GetPreferredHeight();
			string text7 = text6;
			if (text6.Length < tenCQTQuanLy.Length)
			{
				text7 = tenCQTQuanLy;
			}
			if (text7.Length > 30)
			{
				float num11 = preferredHeight / row.Height;
				preferredHeight += num11 * 5f;
				row.Height = preferredHeight;
			}
			if (PeriodInfo.MaHSDuAnDT.Equals(""))
			{
				for (int num12 = 49; num12 < 73; num12++)
				{
					sheetView.SetRowVisible(num12, visible: false);
				}
			}
			else
			{
				int[] array5 = GeneralFunctions.ArrayCellId("J_51");
				sheetView.Cells[array5[1], array5[0]].Text = PeriodInfo.MaHSDuAnDT;
				for (int num12 = 49; num12 < 73; num12++)
				{
					if (num12 != 52)
					{
						sheetView.SetRowVisible(num12, visible: true);
					}
				}
			}
			TextCellType textCellType = new TextCellType();
			textCellType.WordWrap = true;
			textCellType.Multiline = true;
			sheetView.Cells[cellCQTCapCuc[1], cellCQTCapCuc[0]].CellType = textCellType;
			sheetView.Cells[cellCQTChiCuc[1], cellCQTChiCuc[0]].CellType = textCellType;
			string[] tTDuAnDT = getTTDuAnDT(PeriodInfo.MaHSDuAnDT, HeaderInfo);
			if (tTDuAnDT != null && tTDuAnDT.Length > 0)
			{
				sheetView.Cells[CellId_TenDADT[1], CellId_TenDADT[0]].Text = tTDuAnDT[3];
				sheetView.Cells[cellCQTCapCuc[1], cellCQTCapCuc[0]].Text = tTDuAnDT[15];
				sheetView.Cells[cellCQTMaCapCuc[1], cellCQTMaCapCuc[0]].Text = tTDuAnDT[16];
				sheetView.Cells[cellCQTChiCuc[1], cellCQTChiCuc[0]].Text = tTDuAnDT[17];
				sheetView.Cells[cellCQTMaChiCuc[1], cellCQTMaChiCuc[0]].Text = tTDuAnDT[18];
				sheetView.Cells[CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]].Text = tTDuAnDT[4] + ", " + tTDuAnDT[9] + ", " + tTDuAnDT[5];
				sheetView.Cells[CellId_TenChuDADT[1], CellId_TenChuDADT[0]].Text = tTDuAnDT[11];
				sheetView.Cells[CellId_MSTDADT[1], CellId_MSTDADT[0]].Text = tTDuAnDT[12];
				sheetView.Cells[CellId_SoDADT[1], CellId_SoDADT[0]].Text = tTDuAnDT[19];
				sheetView.Cells[CellId_NgayDADT[1], CellId_NgayDADT[0]].Text = HTKKFormater.FormatStringToDate(tTDuAnDT[20]);
				sheetView.Cells[CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]].Text = tTDuAnDT[21];
				sheetView.Cells[CellId_NgayHoanThanhDADT[1], CellId_NgayHoanThanhDADT[0]].Text = HTKKFormater.FormatStringToDate(tTDuAnDT[22]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[1];
			setDateCell(sheetView.Cells[9, HTKKFile.ColumnLettersToIndex("AH") - 1], DateTime.Today);
			winFpSpread.Sheets[1].AutoCalculation = false;
			int index = 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int column = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num = dataLocationTemplateRowStart - 1;
			StyleInfo info = new StyleInfo();
			sheetView.GetStyleInfo(dataLocationTemplateRowStart - 1, column, info);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				Cell cell = sheetView.Cells[num + i, column];
				cell.Text = (i + 1).ToString();
				sheetView.SetStyleInfo(num + i, column, info);
			}
			winFpSpread.Sheets[1].AutoCalculation = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			setDateCell(sheetView.Cells[9, HTKKFile.ColumnLettersToIndex("AN") - 1], DateTime.Today);
			winFpSpread.Sheets[2].AutoCalculation = false;
			int index = 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[index];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int column = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num = dataLocationTemplateRowStart - 1;
			StyleInfo info = new StyleInfo();
			sheetView.GetStyleInfo(dataLocationTemplateRowStart - 1, column, info);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				Cell cell = sheetView.Cells[num + i, column];
				cell.Text = (i + 1).ToString();
				sheetView.SetStyleInfo(num + i, column, info);
			}
			winFpSpread.Sheets[2].AutoCalculation = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			setDateCell(sheetView.Cells[9, HTKKFile.ColumnLettersToIndex("AA") - 1], DateTime.Today);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[4];
			setDateCell(sheetView.Cells[13, HTKKFile.ColumnLettersToIndex("AB") - 1], DateTime.Today);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[5];
			setDateCell(sheetView.Cells[9, HTKKFile.ColumnLettersToIndex("AH") - 1], DateTime.Today);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void setDateCell(Cell cell, DateTime date)
	{
		cell.Value = date.ToString("dd/MM/yyyy");
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		//IL_099f: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("R");
			int num5 = HTKKFile.ColumnLettersToIndex("BD");
			int num6 = HTKKFile.ColumnLettersToIndex("BH");
			int num7 = HTKKFile.ColumnLettersToIndex("AH");
			int num8 = HTKKFile.ColumnLettersToIndex("AM");
			int num9 = HTKKFile.ColumnLettersToIndex("AA");
			int num10 = HTKKFile.ColumnLettersToIndex("AX");
			int num11 = HTKKFile.ColumnLettersToIndex("BA");
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo sInfo = ToKhaiInfo.SheetInfos[0];
			int num12 = HTKKFile.ColumnLettersToIndex("Q");
			int num13 = HTKKFile.ColumnLettersToIndex("AX");
			int num14 = GeneralFunctions.DetectIndexOfTable(sInfo, (int)iRow);
			string text = sheetView.Cells[(int)iRow - 1, num13 - 1].Text.Trim();
			if (iCol == num2 && iRow == 12)
			{
				foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.LoaiHoanThue)
				{
					if (sheetView.Cells[num - 1, num2 - 1].Text == item.Value)
					{
						sheetView.Cells[num - 1, num5 - 1].Text = item.Key;
						break;
					}
				}
			}
			else if (iCol == num3 && iRow == 23)
			{
				foreach (KeyValuePair<string, string> item2 in HTKKWinDictionaryDiaDanh.LyDoHoanThue)
				{
					if (sheetView.Cells[num - 1, num3 - 1].Text == item2.Value)
					{
						sheetView.Cells[num - 1, num6 - 1].Text = item2.Key;
						break;
					}
				}
				GeneralFunctions.SetRowHeight(sheetView, 87);
			}
			else if (iCol == num7)
			{
				string text2 = sheetView.Cells[num - 1, num11 - 1].Text;
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.MaskChar = '.';
				switch (text2)
				{
				case "Y":
					maskCellType.Mask = "0000";
					break;
				case "Q":
					maskCellType.Mask = "0/0000";
					break;
				case "M":
					maskCellType.Mask = "00/0000";
					break;
				case "D":
					maskCellType.Mask = "00/00/0000";
					break;
				}
				sheetView.Cells[(int)iRow - 1, num8 - 1].Text = "";
				sheetView.Cells[(int)iRow - 1, num8 - 1].ResetCellType();
				sheetView.Cells[(int)iRow - 1, num8 - 1].CellType = maskCellType;
			}
			else if (iCol == num8)
			{
				string text2 = sheetView.Cells[num - 1, num11 - 1].Text;
				string text3 = sheetView.Cells[(int)iRow - 1, num8 - 1].Text;
				frmMessage frmMessage = new frmMessage("0456", "3", "", "Thông báo");
				switch (text2)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text3, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = DateTime.ParseExact("01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = text3;
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy").ToString("MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = ConvertHelpers.ConvertToDateTime(text3, "dd/MM/yyyy").ToString("MM/yyyy");
					break;
				case "Y":
					if (ConvertHelpers.ConvertToDateTime("01/01/" + text3, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = DateTime.ParseExact("01/01/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = DateTime.ParseExact("01/12/" + text3, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = "01/" + text3;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = "12/" + text3;
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text3.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text3.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text3.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text3.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text3.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text3.Split(new char[1] { '/' })[1]), 1, 1).AddMonths(3).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = text3;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = text3;
					break;
				default:
					sheetView.Cells[(int)iRow - 1, num8 - 1].Text = "";
					((Form)frmMessage).ShowDialog();
					break;
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			int num15 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			if (iRow >= num15 && iRow <= hTKKCMCTableInfo.RowsValues.Count + num15 && (int)iCol == HTKKFile.ColumnLettersToIndex("AU"))
			{
				string text4 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BB") - 1].Text;
				if (text4 == "5")
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AW") - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AW") - 1].Locked = true;
				}
			}
			if (iRow >= num15 && iRow <= hTKKCMCTableInfo.RowsValues.Count + num15 && (int)iCol == HTKKFile.ColumnLettersToIndex("D"))
			{
				string text5 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text;
				if (string.IsNullOrEmpty(text5))
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("Q") - 1].Text = PeriodInfo.tenCQTCapCuc;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AX") - 1].Text = PeriodInfo.maCQTCapCuc;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CellChange_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("AH");
			int num3 = HTKKFile.ColumnLettersToIndex("AM");
			int num4 = HTKKFile.ColumnLettersToIndex("AA");
			int num5 = HTKKFile.ColumnLettersToIndex("AX");
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			if ((int)iCol != HTKKFile.ColumnLettersToIndex("L") || (int)iRow - 1 != 11)
			{
				return;
			}
			Dictionary<string, string> loaiHoanThue = HTKKWinDictionaryDiaDanh.LoaiHoanThue;
			string text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
			foreach (KeyValuePair<string, string> item in loaiHoanThue)
			{
				if (item.Value == text)
				{
					sheetView.Cells[GeneralFunctions.ArrayCellId("BD_12")[1], GeneralFunctions.ArrayCellId("BD_12")[0]].Text = item.Key;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/chuyenKhoan"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tienMat"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/nguoiNhanTien"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/cmt-HoChieu"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/ngayCap"));
			HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tinh-ThanhPho_ten"));
			HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tinh-ThanhPho_ma"));
			HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/noiNhanTien"));
			HTKKCMCCellInfo hTKKCMCCellInfo9 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tenChuTaiKhoan"));
			HTKKCMCCellInfo hTKKCMCCellInfo10 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/taiKhoanSo"));
			HTKKCMCCellInfo hTKKCMCCellInfo11 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/taiNganHang_ten"));
			if (idxRow == hTKKCMCCellInfo.CellTemplate.RowIndex - 1 && idxCol == hTKKCMCCellInfo.CellTemplate.ColIndex - 1)
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value))
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = false;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Text = "";
				}
				else
				{
					sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = true;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Text = "";
				}
			}
			if (idxRow == hTKKCMCCellInfo2.CellTemplate.RowIndex - 1 && idxCol == hTKKCMCCellInfo2.CellTemplate.ColIndex - 1)
			{
				if (ConvertHelpers.ConvertToBool(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value))
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = false;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Text = "";
				}
				else
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = true;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Locked = true;
					sheetView.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Text = "";
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Locked = false;
					sheetView.Cells[hTKKCMCCellInfo11.CellTemplate.RowIndex - 1, hTKKCMCCellInfo11.CellTemplate.ColIndex - 1].Text = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 2 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_1_HT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 2 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_2_HT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_3(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 2 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_3a_HT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
			if (activeRowIndex == 2 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_3b_HT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ButtonClick_5(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<string, int> DicSheet)
	{
		try
		{
			int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
			if (winFpSpread.ActiveSheet.ActiveRowIndex == 0 && activeColumnIndex == 2)
			{
				GeneralFunctions.ShowBangKe("Bang_Ke_01_4_HT.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void initComboboxQuy()
	{
		comboQuy = new ComboBoxCellType();
		string[] array = new string[4];
		for (int i = 1; i <= 4; i++)
		{
			array[i - 1] = "Q" + i;
		}
		comboQuy.ItemData = array;
		comboQuy.Items = array;
	}

	private void initComboboxThang()
	{
		comboThang = new ComboBoxCellType();
		string[] array = new string[12];
		for (int i = 1; i <= 12; i++)
		{
			array[i - 1] = i.ToString();
		}
		comboThang.ItemData = array;
		comboThang.Items = array;
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/maCQTNoiNop");
			XmlNode val2 = xmlNode.SelectSingleNode("HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/tenCQTNoiNop");
			if (val != null)
			{
				val.InnerText = PeriodInfo.maCQTQuanLy;
			}
			if (val2 != null)
			{
				val2.InnerText = PeriodInfo.tenCQTQuanLy;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo)
	{
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3a(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_3b(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			Validate_4(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
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
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			if (!PeriodInfo.MaHSDuAnDT.Equals(""))
			{
				string text = sheetView.Cells[CellId_MSTDADT[1], CellId_MSTDADT[0]].Text;
				string text2 = sheetView.Cells[CellId_TenChuDADT[1], CellId_TenChuDADT[0]].Text;
				string text3 = sheetView.Cells[CellId_TenDADT[1], CellId_TenDADT[0]].Text;
				string text4 = sheetView.Cells[CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]].Text;
				string text5 = sheetView.Cells[CellId_SoDADT[1], CellId_SoDADT[0]].Text;
				string text6 = sheetView.Cells[CellId_NgayDADT[1], CellId_NgayDADT[0]].Text;
				string text7 = sheetView.Cells[CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]].Text;
				string mst = TTinChung.mst;
				if (text3.Trim().Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_TenDADT[1], CellId_TenDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround1(sheetView, CellId_TenDADT[1], CellId_TenDADT[0]);
				}
				if (text4.Trim().Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_DiaChiDADT[1], CellId_DiaChiDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround1(sheetView, CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]);
				}
				if (text5.Trim().Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_SoDADT[1], CellId_SoDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround1(sheetView, CellId_SoDADT[1], CellId_SoDADT[0]);
				}
				if (text6.Trim().Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_NgayDADT[1], CellId_NgayDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else if (!(ConvertHelpers.ConvertToDateTime(text6, "dd/MM/yyyy") == DateTime.MinValue))
				{
					HTKKFile.RemoveValidateBackGround1(sheetView, CellId_NgayDADT[1], CellId_NgayDADT[0]);
				}
				if (text7.Trim().Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidateBackGround1(sheetView, CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]);
				}
				if (text.Trim().Equals(""))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_MSTDADT[1], CellId_MSTDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37")));
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThue(text.Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, CellId_MSTDADT[1], CellId_MSTDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
					value = 2;
				}
				else if (mst.Length == 13 || mst.Length == 14)
				{
					if (!mst.Substring(0, 10).Equals(text.Substring(0, 10)))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, CellId_MSTDADT[1], CellId_MSTDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_50")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidateBackGround1(sheetView, CellId_MSTDADT[1], CellId_MSTDADT[0]);
					}
				}
				else
				{
					HTKKFile.RemoveValidateBackGround1(sheetView, CellId_MSTDADT[1], CellId_MSTDADT[0]);
				}
				if (!mst.Trim().Equals(text.Trim()) && !mst.Trim().Equals(text.Trim().Replace("-", "")))
				{
					if (text2.Equals(""))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, CellId_TenChuDADT[1], CellId_TenChuDADT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51")));
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidateBackGround1(sheetView, CellId_TenChuDADT[1], CellId_TenChuDADT[0]);
					}
				}
			}
			else
			{
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_MSTDADT[1], CellId_MSTDADT[0]);
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_TenChuDADT[1], CellId_TenChuDADT[0]);
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_TenDADT[1], CellId_TenDADT[0]);
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]);
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_SoDADT[1], CellId_SoDADT[0]);
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_NgayDADT[1], CellId_NgayDADT[0]);
				HTKKFile.RemoveValidateBackGround1(sheetView, CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]);
			}
			string text8 = sheetView.Cells[CellId_SoDeNghiHoan[1], CellId_SoDeNghiHoan[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_SoDeNghiHoan[1], CellId_SoDeNghiHoan[0]);
			if (string.IsNullOrEmpty(text8))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_SoDeNghiHoan[1], CellId_SoDeNghiHoan[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			string text9 = sheetView.Cells[CellId_LoaiHT[1], CellId_LoaiHT[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_LoaiHT[1], CellId_LoaiHT[0]);
			if (string.IsNullOrEmpty(text9))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_LoaiHT[1], CellId_LoaiHT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			string text10 = sheetView.Cells[CellId_CQTChiCuc[1], CellId_CQTChiCuc[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_CQTChiCuc[1], CellId_CQTChiCuc[0]);
			if (string.IsNullOrEmpty(text10))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_CQTChiCuc[1], CellId_CQTChiCuc[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			string text11 = sheetView.Cells[CellId_LyDoHT[1], CellId_LyDoHT[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_LyDoHT[1], CellId_LyDoHT[0]);
			if (string.IsNullOrEmpty(text11))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_LyDoHT[1], CellId_LyDoHT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value);
				string text12 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				if (num2 < 300000000.0 && (text12.Equals("0002") || text12.Equals("0039")))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			num = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string text13 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("D") - 1].Text;
				string text14 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Text;
				string text15 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Q") - 1].Text;
				string text16 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("W") - 1].Text;
				string text17 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AC") - 1].Text;
				string text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Text;
				string text19 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AM") - 1].Text;
				double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AP") - 1].Value);
				string text20 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AU") - 1].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("W") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AP") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AU") - 1);
				if (!string.IsNullOrEmpty(text13) || !string.IsNullOrEmpty(text14) || !string.IsNullOrEmpty(text15) || !string.IsNullOrEmpty(text16) || !string.IsNullOrEmpty(text17) || !string.IsNullOrEmpty(text18) || !string.IsNullOrEmpty(text19) || num3 != 0.0 || !string.IsNullOrEmpty(text20))
				{
					if (string.IsNullOrEmpty(text13))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThue(text13.Trim()))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text15))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("W") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text17))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text19))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (num3 <= 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AP") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/BuTruChoKhoanPNop/tongCot7"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/TTinDNHT/tongCot6"));
				double num4 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value);
				double num5 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo2.Value);
				HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				if (num4 > num5)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_54")));
					value = 2;
				}
				string text21 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AM") - 1].Text;
				if (string.IsNullOrEmpty(text21))
				{
					continue;
				}
				string text22 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BA") - 1].Text;
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy"));
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
				switch (text22)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text21, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						value = 2;
					}
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text21, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						value = 2;
					}
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text21.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text21.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						value = 2;
					}
					break;
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
			SheetView sheetView2 = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[1];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(1, out value);
			int[] array = GeneralFunctions.ArrayCellId("BA_13");
			int[] array2 = GeneralFunctions.ArrayCellId("BB_13");
			string text = sheetView.Cells[array[1], array[0]].Text;
			string text2 = sheetView.Cells[array2[1], array2[0]].Text;
			DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", null);
			DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", null);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string text5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string text6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToStringOrDefault();
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Value.ToStringOrDefault();
				double num2 = Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				long num4 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AN") - 1].Value);
				string text7 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AQ") - 1].Value.ToStringOrDefault();
				long num5 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AT") - 1].Value);
				string text8 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AQ") - 1);
				if (text5.IsNulOrEmpty())
				{
					if (str.IsNulOrEmpty())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")), i + 1);
						value = 2;
					}
					if (str3.IsNulOrEmpty())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")), i + 1);
						value = 2;
					}
					if (text3.IsNulOrEmpty())
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")), i + 1);
						value = 2;
					}
					else
					{
						DateTime now = DateTime.Now;
						DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
						if (dateTime3 > now)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")), i + 1);
							value = 2;
						}
					}
					if (num5 == 0 && text7 != "Khác")
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
						value = 2;
					}
					continue;
				}
				if (!ValidateBase.CheckMaSoThue(text5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")), i + 1);
					value = 2;
				}
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")), i + 1);
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")), i + 1);
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")), i + 1);
					value = 2;
				}
				if (text3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")), i + 1);
					value = 2;
				}
				else
				{
					DateTime now = DateTime.Now;
					DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
					if (dateTime3 > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")), i + 1);
						value = 2;
					}
				}
				if (text4.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_64")), i + 1);
					value = 2;
				}
				else if (GeneralFunctions.HasInvalidXmlChars(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage("0661"), i + 1);
					value = 2;
				}
				if (text6.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_66")), i + 1);
					value = 2;
				}
				else if (GeneralFunctions.HasInvalidXmlChars(text6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage("0661"), i + 1);
					value = 2;
				}
				if (str4.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_67")), i + 1);
					value = 2;
				}
				if (num2 == 0.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
					value = 2;
				}
				if (num3 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
					value = 2;
				}
				if (num4 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
					value = 2;
				}
				if (text7.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AQ") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_71")), i + 1);
					value = 2;
				}
				if (num5 > 0 && num5 > num4)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_85")), i + 1);
					value = 2;
				}
				if (num5 == 0 && text7 != "Khác")
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
					value = 2;
				}
				if (num4 > 0 && num4 < num3 && num2 >= 1.0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_91")), i + 1);
					value = 2;
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[1].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_1_HT/tongCong"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/TTinDNHT/tongCot6"));
			double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value);
			double num7 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo2.Value);
			HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			if (num6 < num7)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_86")));
				value = 2;
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

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[2];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[2];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(2, out value);
			int[] array = GeneralFunctions.ArrayCellId("BA_13");
			int[] array2 = GeneralFunctions.ArrayCellId("BB_13");
			string text = sheetView.Cells[array[1], array[0]].Text;
			string text2 = sheetView.Cells[array2[1], array2[0]].Text;
			DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", null).AddMonths(-1);
			DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", null);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
				double num2 = Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AN") - 1].Value);
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Text;
				string text5 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				if (!GeneralFunctions.HaveToValidateRow(hTKKCMCTableInfo.RowsInfo, hTKKCMCTableInfo.RowsValues[i].RowValue.ToArray()))
				{
					continue;
				}
				if (!text3.IsNulOrEmpty())
				{
					DateTime now = DateTime.Now;
					DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
					if (dateTime3 > dateTime2 || dateTime3 < dateTime)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_76")), i + 1);
						value = 2;
					}
				}
				if (text4.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")), i + 1);
					value = 2;
				}
				if (num3 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")), i + 1);
					value = 2;
				}
			}
			switch (value)
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
			throw ex;
		}
	}

	public void Validate_4(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[5];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[5];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(5, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Value.ToStringOrDefault();
				string text = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("S") - 1].Value.ToStringOrDefault();
				string str5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string str6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string text2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AC") - 1].Text;
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Text;
				long num2 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AN") - 1].Text;
				string text5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AQ") - 1].Text;
				long num4 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AT") - 1].Value);
				long num5 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AW") - 1].Value);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AQ") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AW") - 1);
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (str4.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (text.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (str5.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (str6.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (text2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (text3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (num2 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (num3 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (text4.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AN") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (text5.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AQ") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (num4 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (num5 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AW") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_70")));
					value = 2;
				}
				if (!text.IsNulOrEmpty() && !ValidateBase.CheckMaSoThue(text))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				if (num3 > num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_80")));
					value = 2;
				}
				if (num4 > Convert.ToInt64((double)num3 * 0.85))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AT") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_81")));
					value = 2;
				}
				if (num5 > Convert.ToInt64((double)num3 * 0.15))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AW") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_82")));
					value = 2;
				}
				if (!text2.IsNulOrEmpty())
				{
					DateTime now = DateTime.Now;
					DateTime dateTime = DateTime.ParseExact(text2, "dd/MM/yyyy", null);
					if (dateTime > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")));
						value = 2;
					}
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[5].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[5].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[5].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_3a(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[3];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[3];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = ToKhaiInfo.SheetInfos[4];
			SheetView sheetView2 = winFpSpread.Sheets[3];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(3, out value);
			int[] array = GeneralFunctions.ArrayCellId("BA_13");
			int[] array2 = GeneralFunctions.ArrayCellId("BB_13");
			string text = sheetView.Cells[array[1], array[0]].Text;
			string text2 = sheetView.Cells[array2[1], array2[0]].Text;
			DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", null);
			DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", null);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("S") - 1].Value.ToStringOrDefault();
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string str5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string str6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToStringOrDefault();
				double num2 = Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num4 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				string text5 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				if (str.IsNulOrEmpty() && str2.IsNulOrEmpty() && str3.IsNulOrEmpty() && text3.IsNulOrEmpty() && str4.IsNulOrEmpty() && text4.IsNulOrEmpty() && str5.IsNulOrEmpty() && str6.IsNulOrEmpty() && num2 == 0.0 && num3 == 0 && num4 == 0)
				{
					continue;
				}
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")));
					value = 2;
				}
				if (text3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
					value = 2;
				}
				else
				{
					DateTime now = DateTime.Now;
					DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
					if (dateTime3 > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")));
						value = 2;
					}
				}
				if (!text4.IsNulOrEmpty() && !ValidateBase.CheckMaSoThue(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				if (num4 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75")));
					value = 2;
				}
			}
			hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("S") - 1].Value.ToStringOrDefault();
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string str5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string str6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToStringOrDefault();
				double num2 = Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num4 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				string text5 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				if (str.IsNulOrEmpty() && str2.IsNulOrEmpty() && str3.IsNulOrEmpty() && text3.IsNulOrEmpty() && str4.IsNulOrEmpty() && text4.IsNulOrEmpty() && str5.IsNulOrEmpty() && str6.IsNulOrEmpty() && num2 == 0.0 && num3 == 0 && num4 == 0)
				{
					continue;
				}
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")));
					value = 2;
				}
				if (text3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
					value = 2;
				}
				else
				{
					DateTime now = DateTime.Now;
					DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
					if (dateTime3 > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")));
						value = 2;
					}
				}
				if (!text4.IsNulOrEmpty() && !ValidateBase.CheckMaSoThue(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				if (num4 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75")));
					value = 2;
				}
			}
			hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[2];
			num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("S") - 1].Value.ToStringOrDefault();
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string str5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string str6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToStringOrDefault();
				double num2 = Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num4 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				string text5 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				if (str.IsNulOrEmpty() && str2.IsNulOrEmpty() && str3.IsNulOrEmpty() && text3.IsNulOrEmpty() && str4.IsNulOrEmpty() && text4.IsNulOrEmpty() && str5.IsNulOrEmpty() && str6.IsNulOrEmpty() && num2 == 0.0 && num3 == 0 && num4 == 0)
				{
					continue;
				}
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")));
					value = 2;
				}
				if (text3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
					value = 2;
				}
				else
				{
					DateTime now = DateTime.Now;
					DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
					if (dateTime3 > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")));
						value = 2;
					}
				}
				if (!text4.IsNulOrEmpty() && !ValidateBase.CheckMaSoThue(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				if (num4 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75")));
					value = 2;
				}
			}
			hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[3];
			num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string text3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("S") - 1].Value.ToStringOrDefault();
				string text4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string str5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string str6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToStringOrDefault();
				double num2 = Convert.ToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num4 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				string text5 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				if (str.IsNulOrEmpty() && str2.IsNulOrEmpty() && str3.IsNulOrEmpty() && text3.IsNulOrEmpty() && str4.IsNulOrEmpty() && text4.IsNulOrEmpty() && str5.IsNulOrEmpty() && str6.IsNulOrEmpty() && num2 == 0.0 && num3 == 0 && num4 == 0)
				{
					continue;
				}
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_60")));
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_61")));
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_62")));
					value = 2;
				}
				if (text3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_63")));
					value = 2;
				}
				else
				{
					DateTime now = DateTime.Now;
					DateTime dateTime3 = DateTime.ParseExact(text3, "dd/MM/yyyy", null);
					if (dateTime3 > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")));
						value = 2;
					}
				}
				if (!text4.IsNulOrEmpty() && !ValidateBase.CheckMaSoThue(text4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_44")));
					value = 2;
				}
				if (num4 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_75")));
					value = 2;
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[3].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_3a_HT/BangKeGTGT/HinhThucHoanTra/soTien"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[3].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_3a_HT/BangKeGTGT/HinhThucHoanTra/soTienHoanCahan"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/TTinDNHT/tongCot6"));
			double num5 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value);
			double num6 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo2.Value);
			double num7 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo3.Value);
			HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1);
			HTKKFile.RemoveValidateBackGround(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
			if (num5 + num6 != num7)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_88")));
				value = 2;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[3].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PLuc_01_3a_HT/BangKeGTGT/HHDV_XANGDUNGCQ/tongCong_ct12"));
			double num8 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo4.Value);
			if (num5 > num8)
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_89")));
				value = 2;
			}
			if (hTKKCMCSheetInfo2 != null)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[4].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("tongTienThueGTGTXacNhanHoan"));
				double num9 = ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo5.Value);
				if (num9 < num6)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_90")));
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[3].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[3].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[3].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_3b(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[4];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[4];
			if (hTKKCMCSheetInfo == null)
			{
				return;
			}
			int value = 0;
			levelWarning.TryGetValue(4, out value);
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				string str = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("G") - 1].Value.ToStringOrDefault();
				string str2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Value.ToStringOrDefault();
				string str3 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("M") - 1].Value.ToStringOrDefault();
				string text = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
				string text2 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("S") - 1].Text;
				string str4 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value.ToStringOrDefault();
				string str5 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Y") - 1].Value.ToStringOrDefault();
				string str6 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AB") - 1].Value.ToStringOrDefault();
				string str7 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AE") - 1].Value.ToStringOrDefault();
				long num2 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Value);
				long num3 = Convert.ToInt64(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AK") - 1].Value);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1);
				if (str.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("G") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (str2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (str3.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("M") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (!text.IsNulOrEmpty())
				{
					DateTime now = DateTime.Now;
					DateTime dateTime = DateTime.ParseExact(text, "dd/MM/yyyy", null);
					if (dateTime > now)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_73")));
						value = 2;
					}
				}
				else
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("P") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (text2.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("S") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (str4.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (str5.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Y") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (str6.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AB") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (str7.IsNulOrEmpty())
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AE") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_15")));
					value = 2;
				}
				if (num2 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_84")));
					value = 2;
				}
				if (num3 == 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_84")));
					value = 2;
				}
				if (num3 > num2)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AK") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_87")));
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[4].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[4].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[4].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public List<string> GetDM(string fileDmThue, string loaiHT)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		try
		{
			List<string> list = new List<string>();
			XmlDocument val = new XmlDocument();
			if (!File.Exists(fileDmThue))
			{
				return null;
			}
			val.Load(fileDmThue);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string text = array[0];
				string text2 = "";
				if (text == loaiHT)
				{
					text2 = array[2];
					list.Add(text2);
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private string[] getTTDuAnDT(string maDADT, HTKKHeaderInfo headerInfo)
	{
		//IL_006e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + headerInfo.mst;
			string path = text + "\\DK_Ma_du_an_DT_100.xml";
			if (File.Exists(path))
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
				XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("RowData");
				foreach (XmlNode item in elementsByTagName)
				{
					XmlNode val = item;
					string innerText = val.InnerText;
					string[] array = Regex.Split(innerText, "###");
					if (array[1] == maDADT)
					{
						return array;
					}
				}
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			winFpSpread.ActiveSheet.AutoCalculation = false;
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView activeSheet = winFpSpread.ActiveSheet;
			if (activeSheetIndex == 1 || activeSheetIndex == 2)
			{
				int column = HTKKFile.ColumnLettersToIndex("D") - 1;
				int num = dataLocationTemplateRowStart - 1;
				StyleInfo info = new StyleInfo();
				activeSheet.GetStyleInfo(dataLocationTemplateRowStart - 1, column, info);
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					Cell cell = activeSheet.Cells[num + i, column];
					cell.Text = (i + 1).ToString();
					activeSheet.SetStyleInfo(num + i, column, info);
				}
			}
			winFpSpread.ActiveSheet.AutoCalculation = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ClearData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView activeSheet = winFpSpread.ActiveSheet;
			int column = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num = dataLocationTemplateRowStart - 1;
			StyleInfo info = new StyleInfo();
			activeSheet.GetStyleInfo(dataLocationTemplateRowStart - 1, column, info);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				Cell cell = activeSheet.Cells[num + i, column];
				cell.Text = (i + 1).ToString();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ClearData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[activeSheetIndex];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView activeSheet = winFpSpread.ActiveSheet;
			int column = HTKKFile.ColumnLettersToIndex("D") - 1;
			int num = dataLocationTemplateRowStart - 1;
			StyleInfo info = new StyleInfo();
			activeSheet.GetStyleInfo(dataLocationTemplateRowStart - 1, column, info);
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				Cell cell = activeSheet.Cells[num + i, column];
				cell.Text = (i + 1).ToString();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
