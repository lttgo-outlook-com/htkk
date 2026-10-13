<?xml version="1.0" encoding="UTF-8"?>
<!-- edited with XMLSpy v2010 (http://www.altova.com) by MESMERiZE (MiZE) -->
<xs:schema xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue" xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="http://kekhaithue.gdt.gov.vn/TKhaiThue" elementFormDefault="qualified" attributeFormDefault="unqualified">
	<xs:redefine schemaLocation="..\TKhaiThue.xsd">
		<xs:complexType name="CTieuTKhaiChinhType">
			<xs:complexContent>
				<xs:extension base="CTieuTKhaiChinhType">
					<xs:sequence>
						<xs:element name="daKKTheoTT18" type="xs:boolean" nillable="true">
							<xs:annotation>
								<xs:documentation>Cho phép điều chỉnh, bổ sung tờ khai 01/XSBHDC đã kê khai kỳ tháng 7, tháng 8 năm 2026 theo Thông tư 18/2026/TT-BTC</xs:documentation>
							</xs:annotation>
						</xs:element>
						<xs:element name="SoPSinhTKy" type="ToKhai">
							<xs:annotation>
								<xs:documentation>Số phát sinh trong kỳ</xs:documentation>
							</xs:annotation>
						</xs:element>
						<xs:element name="SoLuyKe" type="ToKhai">
							<xs:annotation>
								<xs:documentation>Số lũy kế</xs:documentation>
							</xs:annotation>
						</xs:element>
					</xs:sequence>
				</xs:extension>
			</xs:complexContent>
		</xs:complexType>
		<xs:complexType name="PLucType">
			<xs:complexContent>
				<xs:extension base="PLucType">
					<xs:sequence>
						<xs:element name="PLuc_01_1_BK_XSBHDCMG" minOccurs="0">
							<xs:complexType>
								<xs:sequence>
									<xs:element name="XoSo">
										<xs:annotation>
											<xs:documentation>I. Hoạt động đại lý xổ số</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="CTietXoSo" nillable="true" maxOccurs="unbounded">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct07" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Họ tên cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct08" type="TINType" nillable="true">
																<xs:annotation>
																	<xs:documentation>Mã số thuế/Số định danh cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct09" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct10" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct11" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế TNCN đã khấu trừ trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct12" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct13" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
														<xs:attribute name="id" type="xs:string" use="required"/>
													</xs:complexType>
												</xs:element>
												<xs:element name="ct09_1" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng doanh thu trong năm với hoạt động xổ số</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct10_1" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V) với hoạt động xổ số</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct11_1" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế TNCN đã khấu trừ trong năm  với hoạt động xổ số</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct12_1" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) với hoạt động xổ số</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct13_1" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) với hoạt động xổ số</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="BaoHiem">
										<xs:annotation>
											<xs:documentation>II .Hoạt động đại lý bảo hiểm</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="CTietBaoHiem" nillable="true" maxOccurs="unbounded">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct07" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Họ tên cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct08" type="TINType" nillable="true">
																<xs:annotation>
																	<xs:documentation>Mã số thuế/Số định danh cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct09" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct10" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct11" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế TNCN đã khấu trừ trong năm </xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct12" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct13" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
														<xs:attribute name="id" type="xs:string" use="required"/>
													</xs:complexType>
												</xs:element>
												<xs:element name="ct09_2" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng doanh thu trong năm với Hoạt động đại lý bảo hiểm</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct10_2" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V) với Hoạt động đại lý bảo hiểm</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct11_2" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế TNCN đã khấu trừ trong năm của Hoạt động đại lý bảo hiểm</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct12_2" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) của Hoạt động đại lý bảo hiểm</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct13_2" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) của Hoạt động đại lý bảo hiểm</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="BHiemNThoBHiemKhongBatBuoc">
										<xs:annotation>
											<xs:documentation>III.Tiền phí tích lũy bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="CTietBaoHiem" nillable="true" maxOccurs="unbounded">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct07" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Họ tên cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct08" type="TINType" nillable="true">
																<xs:annotation>
																	<xs:documentation>Mã số thuế/Số định danh cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct09" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct10" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct11" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế TNCN đã khấu trừ trong năm </xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct12" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) </xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct13" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) </xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
														<xs:attribute name="id" type="xs:string" use="required"/>
													</xs:complexType>
												</xs:element>
												<xs:element name="ct09_3" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng doanh thu trong năm với Tiền phí tích lũy bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct10_3" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V) với Tiền phí tích lũy bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct11_3" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế TNCN đã khấu trừ trong năm với Tiền phí tích lũy bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct12_3" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) với Tiền phí tích lũy bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct13_3" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) với Tiền phí tích lũy bảo hiểm nhân thọ, bảo hiểm không bắt buộc khác</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="BanHangDC">
										<xs:annotation>
											<xs:documentation>IV .Hoạt động bán hàng đa cấp</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="CTietBanHangDC" nillable="true" maxOccurs="unbounded">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct07" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Họ tên cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct08" type="TINType" nillable="true">
																<xs:annotation>
																	<xs:documentation>Mã số thuế/Số định danh cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct09" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct10" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct11" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế TNCN đã khấu trừ trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct12" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct13" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
														<xs:attribute name="id" type="xs:string" use="required"/>
													</xs:complexType>
												</xs:element>
												<xs:element name="ct09_4" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng doanh thu trong năm của Hoạt động bán hàng đa cấp</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct10_4" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V) của Hoạt động bán hàng đa cấp</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct11_4" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế TNCN đã khấu trừ trong năm của Hoạt động bán hàng đa cấp</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct12_4" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) của Hoạt động bán hàng đa cấp</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct13_4" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) của Hoạt động bán hàng đa cấp</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="HoatDongMG">
										<xs:annotation>
											<xs:documentation>V .Hoạt động môi giới</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="CTietHoatDongMG" nillable="true" maxOccurs="unbounded">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct07" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Họ tên cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct08" type="TINType" nillable="true">
																<xs:annotation>
																	<xs:documentation>Mã số thuế/Số định danh cá nhân</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct09" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct10" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct11" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế TNCN đã khấu trừ trong năm</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct12" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct13" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V)</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
														<xs:attribute name="id" type="xs:string" use="required"/>
													</xs:complexType>
												</xs:element>
												<xs:element name="ct09_5" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng doanh thu trong năm của Hoạt động môi giới</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct10_5" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V) của Hoạt động môi giới</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct11_5" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế TNCN đã khấu trừ trong năm của Hoạt động môi giới</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct12_5" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) của Hoạt động môi giới</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ct13_5" type="xs:long" nillable="true">
													<xs:annotation>
														<xs:documentation>Tổng Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) của Hoạt động môi giới</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="ct14" type="xs:long" nillable="true">
										<xs:annotation>
											<xs:documentation>Doanh thu trong năm</xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="ct15" type="xs:long" nillable="true">
										<xs:annotation>
											<xs:documentation>Doanh thu tính thuế TNCN (chỉ áp dụng cho mục I, II, IV, V)</xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="ct16" type="xs:long" nillable="true">
										<xs:annotation>
											<xs:documentation>Số thuế TNCN đã khấu trừ trong năm </xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="ct17" type="xs:long" nillable="true">
										<xs:annotation>
											<xs:documentation>Doanh thu tính thuế GTGT (chỉ áp dụng cho mục I, II, IV,V) </xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="ct18" type="xs:long" nillable="true">
										<xs:annotation>
											<xs:documentation>Số thuế GTGT phải nộp trong năm (chỉ áp dụng cho mục I, II, IV,V) </xs:documentation>
										</xs:annotation>
									</xs:element>
								</xs:sequence>
							</xs:complexType>
						</xs:element>
						<xs:element name="PL01_KHBS" minOccurs="0">
							<xs:complexType>
								<xs:sequence>
									<xs:element name="Header">
										<xs:annotation>
											<xs:documentation>Thông tin chung</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="maTKhai" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Mã tờ khai</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="tenTKhai" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Tên tờ khai</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="maGiaoDich" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Mã giao dịch điện tử</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="KyKKhaiThue">
													<xs:annotation>
														<xs:documentation>Kỳ kê khai thuế</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="kieuKy">
																<xs:annotation>
																	<xs:documentation>Kiểu kỳ</xs:documentation>
																</xs:annotation>
																<xs:simpleType>
																	<xs:restriction base="xs:string">
																		<xs:pattern value="D|M|Y|K|Q"/>
																	</xs:restriction>
																</xs:simpleType>
															</xs:element>
															<xs:element name="kyKKhai" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Quý: Q/YYYY, ngày: DD/MM/YYYY, tháng: MM/YYYY, năm: YYYY, Kỳ: K/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiTuNgay" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng DD/MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiDenNgay" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng DD/MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiTuThang" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiDenThang" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
												<xs:element name="soLan" type="xs:nonNegativeInteger">
													<xs:annotation>
														<xs:documentation>Số lần</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="mst" type="TINType">
													<xs:annotation>
														<xs:documentation>Mã số thuế NNT</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="tenNNT" type="xs:token">
													<xs:annotation>
														<xs:documentation>Tên NNT</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="mstDLyThue" type="TINType" nillable="true">
													<xs:annotation>
														<xs:documentation>Mã số thuế đại lý thuế</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="tenDLyThue" type="xs:token" nillable="true">
													<xs:annotation>
														<xs:documentation>Tên đại lý thuế</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="soHDongDLyThue" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Số hợp đồng đại lý thuế</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ngayKyHDDLyThue" type="xs:date" nillable="true">
													<xs:annotation>
														<xs:documentation>Ngày ký hợp đồng</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="ma_DonViTien" type="xs:string" nillable="true">
										<xs:annotation>
											<xs:documentation>Mã đơn vị tiền</xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="ten_DonViTien" type="xs:string" nillable="true">
										<xs:annotation>
											<xs:documentation>Tên đơn vị tiền</xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="Muc_A">
										<xs:annotation>
											<xs:documentation>A. Xác định tăng/giảm số thuế phải nộp và tiền chậm nộp, tăng/giảm số thuế được khấu trừ, tăng/giảm số thuế đề nghị hoàn</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="Muc_I">
													<xs:annotation>
														<xs:documentation>I. Xác định tăng/giảm số thuế phải nộp và tiền chậm nộp:</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="Muc_1">
																<xs:annotation>
																	<xs:documentation>1. Số thuế phải nộp trên tờ khai điều chỉnh tăng/giảm</xs:documentation>
																</xs:annotation>
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="ChiTiet" maxOccurs="unbounded">
																			<xs:complexType>
																				<xs:sequence>
																					<xs:element name="ct2_ma" type="xs:string" nillable="true">
																						<xs:annotation>
																							<xs:documentation>Mã tiểu mục</xs:documentation>
																						</xs:annotation>
																					</xs:element>
																					<xs:element name="ct2_ten" type="xs:string" nillable="true">
																						<xs:annotation>
																							<xs:documentation>Tên tiểu mục</xs:documentation>
																						</xs:annotation>
																					</xs:element>
																					<xs:element name="ct3" type="xs:long" nillable="true">
																						<xs:annotation>
																							<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																						</xs:annotation>
																					</xs:element>
																				</xs:sequence>
																				<xs:attribute name="id" type="xs:string" use="required"/>
																			</xs:complexType>
																		</xs:element>
																		<xs:element name="tongCong_ct10" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																	</xs:sequence>
																</xs:complexType>
															</xs:element>
															<xs:element name="Muc_2">
																<xs:annotation>
																	<xs:documentation>2. Số thuế phải nộp trên Phụ lục phân bổ điều chỉnh tăng/giảm</xs:documentation>
																</xs:annotation>
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="DSachPLuc">
																			<xs:complexType>
																				<xs:sequence>
																					<xs:element name="BKePLuc" maxOccurs="unbounded">
																						<xs:complexType>
																							<xs:sequence>
																								<xs:element name="ma_PLuc" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Mã phụ lục</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ten_PLuc" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Tên phụ lục</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="CTietPLuc">
																									<xs:complexType>
																										<xs:sequence>
																											<xs:element name="ChiTiet" minOccurs="0" maxOccurs="unbounded">
																												<xs:complexType>
																													<xs:sequence>
																														<xs:element name="ct2_ma" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Mã đơn vị phụ thuộc, địa điểm kinh doanh/tên mặt hàng/tỉnh nơi không có đơn vị phụ thuộc/địa điểm kinh doanh/đơn vị tiêu thụ/cơ sở khai thác></xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct2_ten" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Tên đơn vị phụ thuộc, địa điểm kinh doanh/tên mặt hàng/tỉnh nơi không có đơn vị phụ thuộc/địa điểm kinh doanh/đơn vị tiêu thụ/cơ sở khai thác</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct3_ma" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Mã số thuế/mã số địa điểm kinh doanh (nếu có)</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct3_ten" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Tên Mã số thuế/mã số địa điểm kinh doanh (nếu có)</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct04_ma" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Mã huyện</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct04_ten" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Tên huyện</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct05_ma" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Mã tỉnh</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct05_ten" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Tên tỉnh</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct06_ma" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Mã cơ quan thuế quản lý địa bàn nhận phân bổ</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct06_ten" type="xs:string" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Tên cơ quan thuế quản lý địa bàn nhận phân bổ</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																														<xs:element name="ct7" type="xs:long" nillable="true">
																															<xs:annotation>
																																<xs:documentation>Số tiền thuế phải nộp điều chỉnh tăng/ giảm (tăng ghi (+), giảm ghi (-))</xs:documentation>
																															</xs:annotation>
																														</xs:element>
																													</xs:sequence>
																													<xs:attribute name="id" type="xs:string" use="required"/>
																												</xs:complexType>
																											</xs:element>
																										</xs:sequence>
																									</xs:complexType>
																								</xs:element>
																							</xs:sequence>
																							<xs:attribute name="id" type="xs:string" use="required"/>
																						</xs:complexType>
																					</xs:element>
																					<xs:element name="tongCong_ct11" type="xs:long" nillable="true">
																						<xs:annotation>
																							<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																						</xs:annotation>
																					</xs:element>
																				</xs:sequence>
																			</xs:complexType>
																		</xs:element>
																	</xs:sequence>
																</xs:complexType>
															</xs:element>
															<xs:element name="Muc_3">
																<xs:annotation>
																	<xs:documentation>3. Xác định số tiền chậm nộp điều chỉnh tăng/giảm (tăng ghi (+), giảm ghi (-)):</xs:documentation>
																</xs:annotation>
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="ct3a_1" type="xs:date" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Đến ngày</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ct3a_2" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Số ngày chậm nộp</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ct3b" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Số tiền chậm nộp tăng/giảm</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																	</xs:sequence>
																</xs:complexType>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
												<xs:element name="Muc_II">
													<xs:annotation>
														<xs:documentation>II. Số thuế được khấu trừ điều chỉnh tăng/giảm</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ChiTiet" maxOccurs="unbounded">
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="ct2_ma" type="xs:string" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Mã tiểu mục</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ct2_ten" type="xs:string" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tên tiểu mục</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ct3" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																	</xs:sequence>
																	<xs:attribute name="id" type="xs:string" use="required"/>
																</xs:complexType>
															</xs:element>
															<xs:element name="tongCong_ct12" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
												<xs:element name="Muc_III">
													<xs:annotation>
														<xs:documentation>III. Số thuế đề nghị hoàn điều chỉnh tăng/giảm</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ChiTiet" maxOccurs="unbounded">
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="ct2_ma" type="xs:string" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Mã tiểu mục</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ct2_ten" type="xs:string" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tên tiểu mục</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ct3" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																	</xs:sequence>
																	<xs:attribute name="id" type="xs:string" use="required"/>
																</xs:complexType>
															</xs:element>
															<xs:element name="tongCong_ct13" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>Tăng/giảm số thuế phải nộp (tăng ghi (+), giảm ghi (-))</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="Muc_B">
										<xs:annotation>
											<xs:documentation>B. Số thuế thu hồi hoàn và tiền chậm nộp (nếu có):</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="Muc_I">
													<xs:annotation>
														<xs:documentation>I. Số thuế thu hồi hoàn</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct_1" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>1. Số tiền thu hồi hoàn</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_2_so" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Quyết định hoàn thuế: Số </xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_2_ngay" type="xs:date" nillable="true">
																<xs:annotation>
																	<xs:documentation>Quyết định hoàn thuế: ngày</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_2_CQT_ma" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Quyết định hoàn thuế: CQT ban hành</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_2_CQT_ten" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Quyết định hoàn thuế: CQT ban hành</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_3_so" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Lệnh hoàn thuế: Số </xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_3_ngay" type="xs:date" nillable="true">
																<xs:annotation>
																	<xs:documentation>Lệnh hoàn thuế: Ngày</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
												<xs:element name="Muc_II">
													<xs:annotation>
														<xs:documentation>II. Tiền chậm nộp</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ct_1" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>1. Số ngày nhận được tiền hoàn thuế</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ct_2" type="xs:long" nillable="true">
																<xs:annotation>
																	<xs:documentation>2. Số tiền chậm nộp </xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
								</xs:sequence>
							</xs:complexType>
						</xs:element>
						<xs:element name="PL01_1_KHBS" minOccurs="0">
							<xs:complexType>
								<xs:sequence>
									<xs:element name="Header">
										<xs:annotation>
											<xs:documentation>Thông tin chung</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="maTKhai" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Mã tờ khai</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="tenTKhai" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Tên tờ khai</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="maGiaoDich" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Mã giao dịch điện tử</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="KyKKhaiThue">
													<xs:annotation>
														<xs:documentation>Kỳ kê khai thuế</xs:documentation>
													</xs:annotation>
													<xs:complexType>
														<xs:sequence>
															<xs:element name="kieuKy">
																<xs:annotation>
																	<xs:documentation>Kiểu kỳ</xs:documentation>
																</xs:annotation>
																<xs:simpleType>
																	<xs:restriction base="xs:string">
																		<xs:pattern value="D|M|Y|K|Q"/>
																	</xs:restriction>
																</xs:simpleType>
															</xs:element>
															<xs:element name="kyKKhai" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Quý: Q/YYYY, ngày: DD/MM/YYYY, tháng: MM/YYYY, năm: YYYY, Kỳ: K/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiTuNgay" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng DD/MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiDenNgay" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng DD/MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiTuThang" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="kyKKhaiDenThang" type="xs:string">
																<xs:annotation>
																	<xs:documentation>Định dạng MM/YYYY</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
												<xs:element name="soLan" type="xs:nonNegativeInteger">
													<xs:annotation>
														<xs:documentation>Số lần</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="mst" type="TINType">
													<xs:annotation>
														<xs:documentation>Mã số thuế NNT</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="tenNNT" type="xs:token">
													<xs:annotation>
														<xs:documentation>Tên NNT</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="mstDLyThue" type="TINType" nillable="true">
													<xs:annotation>
														<xs:documentation>Mã số thuế đại lý thuế</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="tenDLyThue" type="xs:token" nillable="true">
													<xs:annotation>
														<xs:documentation>Tên đại lý thuế</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="soHDongDLyThue" type="xs:string" nillable="true">
													<xs:annotation>
														<xs:documentation>Số hợp đồng đại lý thuế</xs:documentation>
													</xs:annotation>
												</xs:element>
												<xs:element name="ngayKyHDDLyThue" type="xs:date" nillable="true">
													<xs:annotation>
														<xs:documentation>Ngày ký hợp đồng</xs:documentation>
													</xs:annotation>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="ma_DonViTien" type="xs:string" nillable="true">
										<xs:annotation>
											<xs:documentation>Mã đơn vị tiền</xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="ten_DonViTien" type="xs:string" nillable="true">
										<xs:annotation>
											<xs:documentation>Tên đơn vị tiền</xs:documentation>
										</xs:annotation>
									</xs:element>
									<xs:element name="Muc_A">
										<xs:annotation>
											<xs:documentation>A. Thông tin khai bổ sung</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="DSachHSo">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="BKeHSo" maxOccurs="unbounded">
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="ma_HSo" type="xs:string" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Mã hồ sơ</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="ten_HSo" type="xs:string" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tên hồ sơ</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="CTietHSo">
																			<xs:complexType>
																				<xs:sequence>
																					<xs:element name="ChiTiet" minOccurs="0" maxOccurs="unbounded">
																						<xs:complexType>
																							<xs:sequence>
																								<xs:element name="ct2_ma" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Mã Tên chỉ tiêu 
điều chỉnh/ Tên hàng hóa, dịch vụ</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct2_ten" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Tên Tên chỉ tiêu 
điều chỉnh/ Tên hàng hóa, dịch vụ</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct3_ma" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Mã số chỉ tiêu</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct3_ten" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Tên Mã số chỉ tiêu</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct3_1_ma" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Mã số chỉ tiêu Dành cho đơn vị phụ thuộc/nhà máy...</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct3_1_ten" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Tên Mã số chỉ tiêu Dành cho đơn vị phụ thuộc/nhà máy...</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct3_2_ma" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Mã số chỉ tiêu Dành cho đ  những mẫu không có key, nhập dạng text </xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct3_2_ten" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Tên Mã số chỉ tiêu Dành cho  những mẫu không có key, nhập dạng text .</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct04" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Số đã kê khai</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct05" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Số điều chỉnh</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct06" type="xs:decimal" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Chênh lệch/thay đổi giữa số điều chỉnh với số đã kê khai</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct7" type="xs:decimal" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Điều chỉnh tăng/giảm nghĩa vụ thuế</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																								<xs:element name="ct8" type="xs:string" nillable="true">
																									<xs:annotation>
																										<xs:documentation>Lý do</xs:documentation>
																									</xs:annotation>
																								</xs:element>
																							</xs:sequence>
																							<xs:attribute name="id" type="xs:string" use="required"/>
																						</xs:complexType>
																					</xs:element>
																				</xs:sequence>
																			</xs:complexType>
																		</xs:element>
																	</xs:sequence>
																	<xs:attribute name="id" type="xs:string" use="required"/>
																</xs:complexType>
															</xs:element>
															<xs:element name="TongCong">
																<xs:complexType>
																	<xs:sequence>
																		<xs:element name="tongCong_7" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tổng cộng (tăng +, giảm -) phải nộp:</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="tongCong_8" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tổng cộng (tăng +, giảm -) khấu trừ:</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																		<xs:element name="tongCong_9" type="xs:long" nillable="true">
																			<xs:annotation>
																				<xs:documentation>Tổng cộng (tăng +, giảm -) đề nghị hoàn:</xs:documentation>
																			</xs:annotation>
																		</xs:element>
																	</xs:sequence>
																</xs:complexType>
															</xs:element>
														</xs:sequence>
													</xs:complexType>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
									<xs:element name="Muc_B">
										<xs:annotation>
											<xs:documentation>B. Tài liệu kèm theo (nếu có)</xs:documentation>
										</xs:annotation>
										<xs:complexType>
											<xs:sequence>
												<xs:element name="CTietTaiLieu" maxOccurs="unbounded">
													<xs:complexType>
														<xs:sequence>
															<xs:element name="ma_TLieu" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Mã tài liệu</xs:documentation>
																</xs:annotation>
															</xs:element>
															<xs:element name="ten_TLieu" type="xs:string" nillable="true">
																<xs:annotation>
																	<xs:documentation>Tên tài liệu</xs:documentation>
																</xs:annotation>
															</xs:element>
														</xs:sequence>
														<xs:attribute name="id" type="xs:string" use="required"/>
													</xs:complexType>
												</xs:element>
											</xs:sequence>
										</xs:complexType>
									</xs:element>
								</xs:sequence>
							</xs:complexType>
						</xs:element>
					</xs:sequence>
				</xs:extension>
			</xs:complexContent>
		</xs:complexType>
	</xs:redefine>
	<xs:complexType name="ToKhai">
		<xs:sequence>
			<xs:element name="ct09" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu trả trong kỳ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct10" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế TNCN</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct11" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế thu TNCN đã khấu trừ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct12" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế GTGT</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct13" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế giá trị gia tăng phải nộp</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct14" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu trả trong kỳ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct15" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế TNCN</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct16" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế GTGT</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct17" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế TNCN đã khấu trừ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct18" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế giá trị gia tăng phải nộp</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct19" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng thu nhập trả trong kỳ </xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct20" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế TNCN đã khấu trừ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct21" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu trả trong kỳ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct22" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế TNCN</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct23" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế TNCN đã khấu trừ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct24" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế GTGT</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct25" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế giá trị gia tăng phải nộp</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct26" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu trả trong kỳ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct27" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế TNCN</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct28" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế TNCN đã khấu trừ</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct29" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng doanh thu tính thuế GTGT</xs:documentation>
				</xs:annotation>
			</xs:element>
			<xs:element name="ct30" type="xs:long" nillable="true">
				<xs:annotation>
					<xs:documentation>Tổng số thuế giá trị gia tăng phải nộp</xs:documentation>
				</xs:annotation>
			</xs:element>
		</xs:sequence>
	</xs:complexType>
</xs:schema>
