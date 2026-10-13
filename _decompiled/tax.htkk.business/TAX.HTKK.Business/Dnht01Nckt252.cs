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

internal class Dnht01Nckt252
{
	private string cellId_SoDeNghiHoan = "H_6";

	private string cellId_HTTruocKTSau = "S_8";

	private string cellId_KTTruocHTSau = "S_10";

	private string cellId_LoaiHT = "L_12";

	private string cellId_CQTCuc = "L_14";

	private string cellId_CQTChiCuc = "T_14";

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
				HTKKCMCCellInfo hTKKCMCCellInfo = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/chuyenKhoan"));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tienMat"));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/nguoiNhanTien"));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/cmt-HoChieu"));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/ngayCap"));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tinh-ThanhPho_ten"));
				HTKKCMCCellInfo hTKKCMCCellInfo7 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/tinh-ThanhPho_ma"));
				HTKKCMCCellInfo hTKKCMCCellInfo8 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/NoiDungDNHT/HThucDNHT/HoanTraTrucTiep/HinhThucHoanTra/noiNhanTien"));
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
				sheetView.Cells[GeneralFunctions.ArrayCellId("J_23")[1], GeneralFunctions.ArrayCellId("J_23")[0]].Text = ToKhaiInfo.SheetInfos[0].StaticCells[8].Value;
			}
			string nameTable = "01_DNHT_II";
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo table) => table.TableName == nameTable);
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			int num2 = HTKKFile.ColumnLettersToIndex("AM");
			for (int num3 = dataLocationTemplateRowStart; num3 <= num; num3++)
			{
				int num4 = HTKKFile.ColumnLettersToIndex("AH");
				int num5 = HTKKFile.ColumnLettersToIndex("BA");
				string text2 = sheetView.Cells[num3 - 1, num5 - 1].Text;
				string text3 = sheetView.Cells[num3 - 1, num2 - 1].Text;
				if (!string.IsNullOrEmpty(text3))
				{
					sheetView.Cells[num3 - 1, num2 - 1].Text = text3.Replace("Q", "").Replace("/", "");
				}
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
				sheetView.Cells[num3 - 1, num2 - 1].ResetCellType();
				sheetView.Cells[num3 - 1, num2 - 1].CellType = maskCellType;
			}
			int[] array = GeneralFunctions.ArrayCellId("L_14");
			int[] array2 = GeneralFunctions.ArrayCellId("T_14");
			int[] array3 = GeneralFunctions.ArrayCellId("AG_14");
			int[] array4 = GeneralFunctions.ArrayCellId("BD_14");
			int[] array5 = GeneralFunctions.ArrayCellId("BD_12");
			sheetView.Cells[array[1], array[0]].Locked = true;
			sheetView.Cells[array2[1], array2[0]].Locked = true;
			string text4 = sheetView.Cells[array5[1], array5[0]].Text.Trim();
			string text5 = TTinChung.tenCQTCapCuc;
			string text6 = TTinChung.tenCQTNoiNop;
			if (text5.IndexOf("-") == 4)
			{
				text5 = text5.Substring(6).Trim();
			}
			if (text4 == "0009")
			{
				List<string> list = GeneralFunctions.SplitString(PeriodInfo.tenCQTCapCuc, " - ");
				text5 = ((list.Count > 1) ? list[1] : list[0]);
				sheetView.Cells[array[1], array[0]].Text = text5;
				sheetView.Cells[array2[1], array2[0]].Text = PeriodInfo.tenCQTQuanLy;
				text6 = PeriodInfo.tenCQTQuanLy;
				sheetView.Cells[array4[1], array4[0]].Text = TTinChung.maCQTQuanLy;
			}
			else
			{
				sheetView.Cells[array[1], array[0]].Text = text5;
				sheetView.Cells[array2[1], array2[0]].Text = TTinChung.tenCQTNoiNop;
				sheetView.Cells[array4[1], array4[0]].Text = TTinChung.maCQTNoiNop;
			}
			Row row = sheetView.Rows[array[1]];
			float preferredHeight = row.GetPreferredHeight();
			string text7 = text5;
			if (text5.Length < text6.Length)
			{
				text7 = text6;
			}
			if (text7.Length > 30)
			{
				float num6 = preferredHeight / row.Height;
				preferredHeight += num6 * 5f;
				row.Height = preferredHeight;
			}
			if (PeriodInfo.MaHSDuAnDT.Equals(""))
			{
				for (int num7 = 49; num7 < 71; num7++)
				{
					sheetView.SetRowVisible(num7, visible: false);
				}
			}
			else
			{
				int[] array6 = GeneralFunctions.ArrayCellId("J_51");
				sheetView.Cells[array6[1], array6[0]].Text = PeriodInfo.MaHSDuAnDT;
				for (int num7 = 49; num7 < 71; num7++)
				{
					sheetView.SetRowVisible(num7, visible: true);
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
				sheetView.Cells[CellId_DiaChiDADT[1], CellId_DiaChiDADT[0]].Text = tTDuAnDT[4] + ", " + tTDuAnDT[9] + ", " + tTDuAnDT[7] + ", " + tTDuAnDT[5];
				sheetView.Cells[CellId_TenChuDADT[1], CellId_TenChuDADT[0]].Text = tTDuAnDT[11];
				sheetView.Cells[CellId_MSTDADT[1], CellId_MSTDADT[0]].Text = tTDuAnDT[12];
				sheetView.Cells[CellId_SoDADT[1], CellId_SoDADT[0]].Text = tTDuAnDT[19];
				sheetView.Cells[CellId_NgayDADT[1], CellId_NgayDADT[0]].Text = HTKKFormater.FormatStringToDate(tTDuAnDT[20]);
				sheetView.Cells[CellId_CoQuanBHDADT[1], CellId_CoQuanBHDADT[0]].Text = tTDuAnDT[21];
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
		//IL_093f: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = (int)iRow;
			int num2 = HTKKFile.ColumnLettersToIndex("L");
			int num3 = HTKKFile.ColumnLettersToIndex("J");
			int num4 = HTKKFile.ColumnLettersToIndex("BD");
			int num5 = HTKKFile.ColumnLettersToIndex("BH");
			int num6 = HTKKFile.ColumnLettersToIndex("AH");
			int num7 = HTKKFile.ColumnLettersToIndex("AM");
			int num8 = HTKKFile.ColumnLettersToIndex("AA");
			int num9 = HTKKFile.ColumnLettersToIndex("AX");
			int num10 = HTKKFile.ColumnLettersToIndex("BA");
			int activeSheetIndex = winFpSpread.ActiveSheetIndex;
			SheetView sheetView = winFpSpread.Sheets[0];
			if (iCol == num2 && iRow == 12)
			{
				foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.LoaiHoanThue)
				{
					if (sheetView.Cells[num - 1, num2 - 1].Text == item.Value)
					{
						sheetView.Cells[num - 1, num4 - 1].Text = item.Key;
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
						sheetView.Cells[num - 1, num5 - 1].Text = item2.Key;
						break;
					}
				}
				GeneralFunctions.SetRowHeight(sheetView, 87);
			}
			else if (iCol == num6)
			{
				string text = sheetView.Cells[num - 1, num10 - 1].Text;
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.MaskChar = '.';
				switch (text)
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
				sheetView.Cells[(int)iRow - 1, num7 - 1].Text = "";
				sheetView.Cells[(int)iRow - 1, num7 - 1].ResetCellType();
				sheetView.Cells[(int)iRow - 1, num7 - 1].CellType = maskCellType;
			}
			else if (iCol == num7)
			{
				string text = sheetView.Cells[num - 1, num10 - 1].Text;
				string text2 = sheetView.Cells[(int)iRow - 1, num7 - 1].Text;
				frmMessage frmMessage = new frmMessage("0456", "3", "", "Thông báo");
				switch (text)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text2, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = DateTime.ParseExact("01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = DateTime.ParseExact("01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = text2;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = text2;
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = text2;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = text2;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy").ToString("MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy").ToString("MM/yyyy");
					break;
				case "Y":
					if (ConvertHelpers.ConvertToDateTime("01/01/" + text2, "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = DateTime.ParseExact("01/01/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = DateTime.ParseExact("01/12/" + text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(1).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = "01/" + text2;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = "12/" + text2;
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text2.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text2.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
					{
						return;
					}
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BC") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text2.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text2.Split(new char[1] { '/' })[1]), 1, 1).ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BD") - 1].Text = GeneralFunctions.GetNgayDauQuy(int.Parse(text2.Split(new char[1] { '/' })[0].Replace("Q", "")), int.Parse(text2.Split(new char[1] { '/' })[1]), 1, 1).AddMonths(3).AddDays(-1.0)
						.ToString("dd/MM/yyyy");
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BE") - 1].Text = text2;
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BF") - 1].Text = text2;
					break;
				default:
					sheetView.Cells[(int)iRow - 1, num7 - 1].Text = "";
					((Form)frmMessage).ShowDialog();
					break;
				}
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[1];
			int num11 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			if (iRow >= num11 && iRow <= hTKKCMCTableInfo.RowsValues.Count + num11 && (int)iCol == HTKKFile.ColumnLettersToIndex("AU"))
			{
				string text3 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("BB") - 1].Text;
				if (text3 == "5")
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AW") - 1].Locked = false;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("AW") - 1].Locked = true;
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
			string text10 = sheetView.Cells[CellId_CQTCuc[1], CellId_CQTCuc[0]].Text;
			string text11 = sheetView.Cells[CellId_CQTChiCuc[1], CellId_CQTChiCuc[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_CQTCuc[1], CellId_CQTCuc[0]);
			HTKKFile.RemoveValidate(sheetView, CellId_CQTChiCuc[1], CellId_CQTChiCuc[0]);
			if (string.IsNullOrEmpty(text10))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_CQTCuc[1], CellId_CQTCuc[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			if (string.IsNullOrEmpty(text11))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_CQTChiCuc[1], CellId_CQTChiCuc[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			string text12 = sheetView.Cells[CellId_LyDoHT[1], CellId_LyDoHT[0]].Text;
			HTKKFile.RemoveValidate(sheetView, CellId_LyDoHT[1], CellId_LyDoHT[0]);
			if (string.IsNullOrEmpty(text12))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, CellId_LyDoHT[1], CellId_LyDoHT[0], HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
				value = 2;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("V") - 1].Value);
				string text13 = sheetView.Cells[CellId_MaLoaiHT[1], CellId_MaLoaiHT[0]].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1);
				if (num2 < 300000000.0 && (text13.Equals("0002") || text13.Equals("0039")))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("V") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_17")));
					value = 2;
				}
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[1];
			num = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart - 1;
			for (int i = 0; i < hTKKCMCTableInfo2.RowsValues.Count; i++)
			{
				string text14 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("D") - 1].Text;
				string text15 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("J") - 1].Text;
				string text16 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("Q") - 1].Text;
				string text17 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("W") - 1].Text;
				string text18 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AC") - 1].Text;
				string text19 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AH") - 1].Text;
				string text20 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AM") - 1].Text;
				double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AP") - 1].Value);
				string text21 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AU") - 1].Text;
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("J") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("W") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AP") - 1);
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AU") - 1);
				if (!string.IsNullOrEmpty(text14) || !string.IsNullOrEmpty(text15) || !string.IsNullOrEmpty(text16) || !string.IsNullOrEmpty(text17) || !string.IsNullOrEmpty(text18) || !string.IsNullOrEmpty(text19) || !string.IsNullOrEmpty(text20) || num3 != 0.0 || !string.IsNullOrEmpty(text21))
				{
					if (string.IsNullOrEmpty(text14))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					else if (!ValidateBase.CheckMaSoThue(text14.Trim()))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("D") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_49")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text16))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("Q") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text17))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("W") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text18))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AC") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text19))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AH") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
						value = 2;
					}
					if (string.IsNullOrEmpty(text20))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_18")));
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
				string text22 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("AM") - 1].Text;
				if (string.IsNullOrEmpty(text22))
				{
					continue;
				}
				string text23 = sheetView.Cells[num + i, HTKKFile.ColumnLettersToIndex("BA") - 1].Text;
				HTKKCMCCellInfo hTKKCMCCellInfo3 = ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy"));
				HTKKFile.RemoveValidate(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1);
				switch (text23)
				{
				case "M":
					if (ConvertHelpers.ConvertToDateTime("01/" + text22, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						value = 2;
					}
					break;
				case "D":
					if (ConvertHelpers.ConvertToDateTime(text22, "dd/MM/yyyy") == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, HTKKFile.ColumnLettersToIndex("AM") - 1, HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_52")));
						value = 2;
					}
					break;
				case "Q":
					if (ConvertHelpers.ConvertToDateTime("01/" + (ConvertHelpers.ConvertToInt(text22.Split(new char[1] { '/' })[0].Replace("Q", "")) * 3 - 2).ToString().PadLeft(2, '0') + "/" + text22.Split(new char[1] { '/' })[1], "dd/MM/yyyy") == DateTime.MinValue)
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
}
