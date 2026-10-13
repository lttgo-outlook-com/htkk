using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKPeriodInfo
{
	public string pbanTKhaiXML { get; set; }

	public string MoTaBieuMau { get; set; }

	public string TenToKhai { get; set; }

	public string menuID { get; set; }

	public string MaToKhai { get; set; }

	public string VersionXML { get; set; }

	public string Ngay { get; set; }

	public string Thang { get; set; }

	public string Nam { get; set; }

	public string kyKKhai { get; set; }

	public string Ky { get; set; }

	public string Quy { get; set; }

	public string LoaiToKhai { get; set; }

	public string SoLan { get; set; }

	public string LanXuatBan { get; set; }

	public string KieuKy { get; set; }

	public string TuNgay { get; set; }

	public string DenNgay { get; set; }

	public string TuThang { get; set; }

	public string DenThang { get; set; }

	public string NgayLapToKhai { get; set; }

	public string NgayKyToKhai { get; set; }

	public string MaLyDoGiaHan { get; set; }

	public string LyDoGiaHan { get; set; }

	public string NgayLapKHBS { get; set; }

	public string NgayChiaCTLN { get; set; }

	public string TuNam { get; set; }

	public string DenNam { get; set; }

	public int TKLanXuatBan { get; set; }

	public int ToKhaiThang { get; set; }

	public int ToKhaiKy { get; set; }

	public int ToKhaiQuy { get; set; }

	public int ToKhaiNam { get; set; }

	public int ToKhaiTuThangDenThang { get; set; }

	public int ToKhaiPhatSinh { get; set; }

	public int ToKhaiTuNamDenNam { get; set; }

	public int QuyetToanChoRiengTungNam { get; set; }

	public int QuyetToanHetVaoNamDau { get; set; }

	public string DoanhNghiepHoatDongLienTuc { get; set; }

	public string DonViTinhTen { get; set; }

	public string DonViTinhMa { get; set; }

	public string NganhNgheKinhDoanh { get; set; }

	public bool CTDC { get; set; }

	public bool LNDC { get; set; }

	public string TieuMucHachToanVal { get; set; }

	public string ToChucKhaiThay { get; set; }

	public int SoNgayNopChamKHBS { get; set; }

	public Dictionary<string, object> dicVar { get; set; }

	public string KieuKyHoanThue { get; set; }

	public string MaHSDuAnDT { get; set; }

	public string ThuHo { get; set; }

	public string tenCQTCapCuc { get; set; }

	public string maCQTCapCuc { get; set; }

	public string maCQTQuanLy { get; set; }

	public string tenCQTQuanLy { get; set; }

	public string TenTinh { get; set; }

	public string MaTinh { get; set; }

	public string TenHuyen { get; set; }

	public string MaHuyen { get; set; }

	public string TenXa { get; set; }

	public string MaXa { get; set; }

	public string TenQuyetToan { get; set; }

	public string MaQuyetToan { get; set; }

	public string NgayTb { get; set; }

	public string SoTB { get; set; }

	public int ToChucCNDNVonChuyenNhuong { get; set; }

	public HTKKPeriodInfo()
	{
		pbanTKhaiXML = string.Empty;
		MoTaBieuMau = string.Empty;
		TenToKhai = string.Empty;
		MaToKhai = string.Empty;
		menuID = string.Empty;
		LoaiToKhai = string.Empty;
		Nam = string.Empty;
		kyKKhai = string.Empty;
		SoLan = string.Empty;
		KieuKy = string.Empty;
		NgayChiaCTLN = string.Empty;
		Ngay = string.Empty;
		Thang = string.Empty;
		Nam = string.Empty;
		Ky = string.Empty;
		Quy = string.Empty;
		LoaiToKhai = string.Empty;
		SoLan = string.Empty;
		LanXuatBan = string.Empty;
		KieuKy = string.Empty;
		TuNgay = string.Empty;
		DenNgay = string.Empty;
		TuThang = string.Empty;
		DenThang = string.Empty;
		TuNam = string.Empty;
		DenNam = string.Empty;
		NgayLapToKhai = string.Empty;
		NgayKyToKhai = string.Empty;
		MaLyDoGiaHan = string.Empty;
		LyDoGiaHan = string.Empty;
		NgayLapKHBS = string.Empty;
		CTDC = false;
		LNDC = false;
		TKLanXuatBan = 0;
		ToKhaiThang = 0;
		ToKhaiKy = 0;
		ToKhaiQuy = 0;
		ToKhaiNam = 0;
		ToKhaiTuThangDenThang = 0;
		ToKhaiTuNamDenNam = 0;
		ToKhaiPhatSinh = 0;
		SoNgayNopChamKHBS = 0;
		QuyetToanChoRiengTungNam = 0;
		QuyetToanHetVaoNamDau = 0;
		DoanhNghiepHoatDongLienTuc = string.Empty;
		DonViTinhTen = string.Empty;
		DonViTinhMa = string.Empty;
		KieuKyHoanThue = string.Empty;
		TieuMucHachToanVal = string.Empty;
		MaHSDuAnDT = string.Empty;
		ThuHo = string.Empty;
		ToChucKhaiThay = string.Empty;
		dicVar = new Dictionary<string, object>();
		tenCQTCapCuc = string.Empty;
		maCQTCapCuc = string.Empty;
		maCQTQuanLy = string.Empty;
		tenCQTQuanLy = string.Empty;
		MaTinh = string.Empty;
		TenTinh = string.Empty;
		MaHuyen = string.Empty;
		TenHuyen = string.Empty;
		MaXa = string.Empty;
		TenXa = string.Empty;
		ToChucCNDNVonChuyenNhuong = 0;
		TenQuyetToan = string.Empty;
		MaQuyetToan = string.Empty;
		NgayTb = string.Empty;
		SoTB = string.Empty;
	}
}
