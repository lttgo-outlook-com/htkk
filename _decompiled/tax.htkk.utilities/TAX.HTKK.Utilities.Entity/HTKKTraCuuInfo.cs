using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKTraCuuInfo
{
	public string ID { get; set; }

	public string FinanceYear { get; set; }

	public string Year { get; set; }

	public string Month { get; set; }

	public string ThreeMonth { get; set; }

	public string Day { get; set; }

	public string Caption { get; set; }

	public string LoaiTK { get; set; }

	public string TkThangQuy { get; set; }

	public string StartDate { get; set; }

	public string EndDate { get; set; }

	public string DataFile { get; set; }

	public string ThueKhauTru { get; set; }

	public string ThuePhaiNop { get; set; }

	public string VersionAPP { get; set; }

	public string VersionXML { get; set; }

	public string KyKeKhai { get; set; }

	public string KieuKy { get; set; }

	public string TenToKhai { get; set; }

	public string FileName { get; set; }

	public string KyTinhThueShow { get; set; }

	public string Lan { get; set; }

	public string TuNgay { get; set; }

	public string DenNgay { get; set; }

	public string TuThang { get; set; }

	public string DenThang { get; set; }

	public string Thang { get; set; }

	public string Quy { get; set; }

	public string Ngay { get; set; }

	public string NgayLapKHBS { get; set; }

	public string DoanhNghiepHoatDongLienTuc { get; set; }

	public string LoaiTKShow { get; set; }

	public string DonViTinhMa { get; set; }

	public string DonViTinhTen { get; set; }

	public string TieuMucHachToan { get; set; }

	public string NganhNgheKD { get; set; }

	public string LanXuatBan { get; set; }

	public string Prefix { get; set; }

	public List<int> ListDateNumber { get; set; }

	public List<HTKKTraCuuValidity> Validity { get; set; }

	public string STT { get; set; }

	public string CheckValue { get; set; }

	public Dictionary<string, string> dicVar { get; set; }

	public string ThuHo { get; set; }

	public string MaHSDuAnDT { get; set; }

	public string tenCQTCapCuc { get; set; }

	public string maCQTCapCuc { get; set; }

	public string maCQTQuanLy { get; set; }

	public string tenCQTQuanLy { get; set; }

	public string maTinh { get; set; }

	public string tenTinh { get; set; }

	public string maHuyen { get; set; }

	public string tenHuyen { get; set; }

	public string maXa { get; set; }

	public string tenXa { get; set; }

	public string KieuKyHoanThue { get; set; }

	public string NopTheoTinhCQT { get; set; }

	public int ToChucCNDNVonChuyenNhuong { get; set; }

	public HTKKTraCuuInfo()
	{
		ID = string.Empty;
		FinanceYear = string.Empty;
		Year = string.Empty;
		Month = string.Empty;
		ThreeMonth = string.Empty;
		Day = string.Empty;
		Caption = string.Empty;
		LoaiTK = string.Empty;
		TkThangQuy = string.Empty;
		StartDate = string.Empty;
		EndDate = string.Empty;
		DataFile = string.Empty;
		ThueKhauTru = string.Empty;
		ThuePhaiNop = string.Empty;
		KyKeKhai = string.Empty;
		KieuKy = string.Empty;
		FileName = string.Empty;
		VersionAPP = string.Empty;
		VersionXML = string.Empty;
		TenToKhai = string.Empty;
		KyTinhThueShow = string.Empty;
		Lan = "0";
		ListDateNumber = new List<int>();
		STT = string.Empty;
		CheckValue = string.Empty;
		TuNgay = string.Empty;
		DenNgay = string.Empty;
		TuThang = string.Empty;
		DenThang = string.Empty;
		Thang = string.Empty;
		Quy = string.Empty;
		Ngay = string.Empty;
		NgayLapKHBS = string.Empty;
		DoanhNghiepHoatDongLienTuc = string.Empty;
		LoaiTKShow = string.Empty;
		DonViTinhMa = string.Empty;
		DonViTinhTen = string.Empty;
		TieuMucHachToan = string.Empty;
		NganhNgheKD = string.Empty;
		LanXuatBan = string.Empty;
		Prefix = string.Empty;
		ThuHo = string.Empty;
		MaHSDuAnDT = string.Empty;
		KieuKyHoanThue = string.Empty;
		Validity = new List<HTKKTraCuuValidity>();
		dicVar = new Dictionary<string, string>();
		maCQTCapCuc = string.Empty;
		tenCQTCapCuc = string.Empty;
		maCQTQuanLy = string.Empty;
		tenCQTQuanLy = string.Empty;
		maTinh = string.Empty;
		tenTinh = string.Empty;
		maHuyen = string.Empty;
		tenHuyen = string.Empty;
		NopTheoTinhCQT = string.Empty;
		ToChucCNDNVonChuyenNhuong = 0;
	}
}
