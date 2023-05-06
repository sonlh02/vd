using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using QuanLyBanHang.Class;
using COMExcel = Microsoft.Office.Interop.Excel;
namespace QuanLyBanHang
{
    public partial class frmHoaDon : Form
    {
        DataTable ChiTietHoaDon;
        public frmHoaDon()
        {
            InitializeComponent();
        }

        private void frmHoaDon_Load(object sender, EventArgs e)
        {
            btnthem.Enabled = true;
            btnluu.Enabled = false;
            btnhuy.Enabled = false;
            btnin.Enabled = false;
            txtmahoadon.ReadOnly = true;
            txttennhanvien.ReadOnly = true;
            txtTenKhachHang.ReadOnly = true;
            txtDiaChi.ReadOnly = true;
            txtSoDienThoai.ReadOnly = true;
            txtTenXe.ReadOnly = true;
            txtdongia.ReadOnly = true;
            txtThanhTien.ReadOnly = true;
            txttongtien.ReadOnly = true;
            txtGiamGia.Text = "0";
            txttongtien.Text = "0";
            Functions.FillCombo("SELECT MaKhachHang, HoTen FROM KhachHang", cbomakhchhang, "MaKhachHang", "MaKhachHang");
            cbomakhchhang.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaNhanVien, HoTen FROM NhanVien", cbomanhanvien, "MaNhanVien", "MaNhanVien");
            cbomanhanvien.SelectedIndex = -1;
            Functions.FillCombo("SELECT MaXe, TenXe FROM XeDien", cbomaxe, "MaXe", "MaXe");
            cbomaxe.SelectedIndex = -1;
            //Hiển thị thông tin của một hóa đơn được gọi từ form tìm kiếm
            if (txtmahoadon.Text != "")
            {
                LoadInfoHoaDon();
                btnhuy.Enabled = true;
                btnin.Enabled = true;
            }
            LoadDataGridView();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT a.MaXe, b.TenXe, a.SoLuong, a.DonGia,a.GiamGia,a.ThanhTien FROM ChiTietHoaDon AS a, XeDien AS b WHERE a.MaHoaDon = N'" + txtmahoadon.Text + "' AND a.MaXe=b.MaXe";
            ChiTietHoaDon = Functions.GetDataToTable(sql);
            dgvHoaDonBan.DataSource = ChiTietHoaDon;
            dgvHoaDonBan.Columns[0].HeaderText = "Mã Xe";
            dgvHoaDonBan.Columns[1].HeaderText = "Tên Xe";
            dgvHoaDonBan.Columns[2].HeaderText = "Số lượng";
            dgvHoaDonBan.Columns[3].HeaderText = "Đơn giá";
            dgvHoaDonBan.Columns[4].HeaderText = "Giảm giá %";
            dgvHoaDonBan.Columns[5].HeaderText = "Thành tiền";
            dgvHoaDonBan.Columns[0].Width = 80;
            dgvHoaDonBan.Columns[1].Width = 130;
            dgvHoaDonBan.Columns[2].Width = 80;
            dgvHoaDonBan.Columns[3].Width = 90;
            dgvHoaDonBan.Columns[4].Width = 90;
            dgvHoaDonBan.Columns[4].Width = 90;
            dgvHoaDonBan.AllowUserToAddRows = false;
            dgvHoaDonBan.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        private void LoadInfoHoaDon()
        {
            string str;
            str = "SELECT NgayLap FROM HoaDon WHERE MaHoaDon = N'" + txtmahoadon.Text + "'";
            dtpngaylap.Value = DateTime.Parse(Functions.GetFieldValues(str));
            str = "SELECT MaNhanVien FROM HoaDon WHERE MaHoaDon = N'" + txtmahoadon.Text + "'";
            cbomanhanvien.Text = Functions.GetFieldValues(str);
            str = "SELECT MaKhachHang FROM HoaDon WHERE MaHoaDon = N'" + txtmahoadon.Text + "'";
            cbomakhchhang.Text = Functions.GetFieldValues(str);
            str = "SELECT TongTien FROM HoaDon WHERE MaHoaDon = N'" + txtmahoadon.Text + "'";
            txttongtien.Text = Functions.GetFieldValues(str);
            lblTongTien.Text = "Bằng chữ: " + Functions.ChuyenSoSangChuoi(double.Parse(txttongtien.Text));
        }

        private void btnthem_Click(object sender, EventArgs e)
        {
            btnhuy.Enabled = false;
            btnluu.Enabled = true;
            btnin.Enabled = false;
            btnthem.Enabled = false;
            ResetValues();
            txtmahoadon.Text = Functions.CreateKey("HDB");
            LoadDataGridView();
        }
        private void ResetValues()
        {
        }

        private void btnluu_Click(object sender, EventArgs e)
        {
           
            string sql;
            double sl, SLcon, tong, Tongmoi;
            sql = "SELECT MaHoaDon FROM HoaDon WHERE MaHoaDon=N'" + txtmahoadon.Text + "'";
            if (!Functions.CheckKey(sql))
            {
                // Mã hóa đơn chưa có, tiến hành lưu các thông tin chung
                // Mã HDBan được sinh tự động do đó không có trường hợp trùng khóa
                if (cbomanhanvien.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cbomanhanvien.Focus();
                    return;
                }
                if (cbomakhchhang.Text.Length == 0)
                {
                    MessageBox.Show("Bạn phải nhập khách hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cbomakhchhang.Focus();
                    return;
                }
                sql = "INSERT INTO HoaDon(MaHoaDon, MaKhachHang, MaNhanvien, NgayLap, TongTien) VALUES (N'" + txtmahoadon.Text.Trim() +
                   "',N'" + cbomakhchhang.SelectedValue +
                    "',N'" + cbomanhanvien.SelectedValue + 
                    "','" + dtpngaylap.Text.Trim() +
                    "'," + txttongtien.Text + ")";
                Functions.RunSQL(sql);
            }
            // Lưu thông tin của các mặt hàng
            if (cbomaxe.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cbomaxe.Focus();
                return;
            }
            if ((txtsoluong.Text.Trim().Length == 0) || (txtsoluong.Text == "0"))
            {
                MessageBox.Show("Bạn phải nhập số lượng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtsoluong.Text = "";
                txtsoluong.Focus();
                return;
            }
            if (txtGiamGia.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập giảm giá", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtGiamGia.Focus();
                return;
            }
            sql = "SELECT MaXe FROM ChiTietHoaDon WHERE MaXe=N'" + cbomaxe.SelectedValue + "' AND MaHoaDon = N'" + txtmahoadon.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("Mã hàng này đã có, bạn phải nhập mã khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetValuesHang();
                cbomaxe.Focus();
                return;
            }
            // Kiểm tra xem số lượng hàng trong kho còn đủ để cung cấp không?
            sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM XeDien WHERE MaXe = N'" + cbomaxe.SelectedValue + "'"));
            if (Convert.ToDouble(txtsoluong.Text) > sl)
            {
                MessageBox.Show("Số lượng mặt hàng này chỉ còn " + sl, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtsoluong.Text = "";
                txtsoluong.Focus();
                return;
            }
            sql = "INSERT INTO ChiTietHoaDon(MaHoaDon,MaXe,SoLuong,DonGia,GiamGia,ThanhTien) VALUES(N'" + txtmahoadon.Text.Trim() + "',N'" + cbomaxe.SelectedValue + "'," + txtsoluong.Text + "," + txtdongia.Text + "," + txtGiamGia.Text + "," + txtThanhTien.Text + ")";
            Functions.RunSQL(sql);
            LoadDataGridView();
            // Cập nhật lại số lượng của mặt hàng vào bảng tblHang
            SLcon = sl - Convert.ToDouble(txtsoluong.Text);
            sql = "UPDATE XeDien SET SoLuong =" + SLcon + " WHERE MaXe= N'" + cbomaxe.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại tổng tiền cho hóa đơn bán
            tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM HoaDon WHERE MaHoaDon = N'" + txtmahoadon.Text + "'"));
            Tongmoi = tong + Convert.ToDouble(txtThanhTien.Text);
            sql = "UPDATE HoaDon SET TongTien =" + Tongmoi + " WHERE MaHoaDon = N'" + txtmahoadon.Text + "'";
            Functions.RunSQL(sql);
            txttongtien.Text = Tongmoi.ToString();
            lblTongTien.Text = "Bằng chữ: " + Functions.ChuyenSoSangChuoi(double.Parse(Tongmoi.ToString()));
            ResetValuesHang();
            btnhuy.Enabled = true;
            btnthem.Enabled = true;
            btnin.Enabled = true;
        }
        private void ResetValuesHang()
        {
            cbomaxe.Text = "";
            txtsoluong.Text = "";
            txtGiamGia.Text = "0";
            txtThanhTien.Text = "0";
        }

        private void cbomaxe_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cbomaxe.Text == "")
            {
                txtTenXe.Text = "";
                txtdongia.Text = "";
            }
            // Khi chọn mã hàng thì các thông tin về hàng hiện ra
            str = "SELECT TenXe FROM XeDien WHERE MaXe =N'" + cbomaxe.SelectedValue + "'";
            txtTenXe.Text = Functions.GetFieldValues(str);
            str = "SELECT GiaBan FROM XeDien WHERE MaXe =N'" + cbomaxe.SelectedValue + "'";
            txtdongia.Text = Functions.GetFieldValues(str);
        }

        private void cbomakhchhang_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cbomakhchhang.Text == "")
            {
                txtTenKhachHang.Text = "";
                txtDiaChi.Text = "";
                txtSoDienThoai.Text = "";
            }
            //Khi chọn Mã khách hàng thì các thông tin của khách hàng sẽ hiện ra
            str = "Select HoTen from KhachHang where MaKhachHang = N'" + cbomakhchhang.SelectedValue + "'";
            txtTenKhachHang.Text = Functions.GetFieldValues(str);
            str = "Select DiaChi from KhachHang where MaKhachHang = N'" + cbomakhchhang.SelectedValue + "'";
            txtDiaChi.Text = Functions.GetFieldValues(str);
            str = "Select SoDienThoai from KhachHang where MaKhachHang= N'" + cbomakhchhang.SelectedValue + "'";
            txtSoDienThoai.Text = Functions.GetFieldValues(str);
        }

        private void cbomanhanvien_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cbomanhanvien.Text == "")
                txttennhanvien.Text = "";
            // Khi chọn Mã nhân viên thì tên nhân viên tự động hiện ra
            str = "Select HoTen from NhanVien where MaNhanVien =N'" + cbomanhanvien.SelectedValue + "'";
            txttennhanvien.Text = Functions.GetFieldValues(str);
        }

        private void txtsoluong_TextChanged(object sender, EventArgs e)
        {
            //Khi thay đổi số lượng thì thực hiện tính lại thành tiền
            double tt, sl, dg, gg;
            if (txtsoluong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtsoluong.Text);
            if (txtGiamGia.Text == "")
                gg = 0;
            else
                gg = Convert.ToDouble(txtGiamGia.Text);
            if (txtdongia.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtdongia.Text);
            tt = sl * dg - sl * dg * gg / 100;
            txtThanhTien.Text = tt.ToString();
        }

        private void txtGiamGia_TextChanged(object sender, EventArgs e)
        {
            //Khi thay đổi giảm giá thì tính lại thành tiền
            double tt, sl, dg, gg;
            if (txtsoluong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtsoluong.Text);
            if (txtGiamGia.Text == "")
                gg = 0;
            else
                gg = Convert.ToDouble(txtGiamGia.Text);
            if (txtdongia.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtdongia.Text);
            tt = sl * dg - sl * dg * gg / 100;
            txtThanhTien.Text = tt.ToString();
        }

        private void cboMaHD_DropDown(object sender, EventArgs e)
        {
            Functions.FillCombo("SELECT MaHoaDon FROM HoaDon", cboMaHD, "MaHoaDon", "MaHoaDon");
            cboMaHD.SelectedIndex = -1;
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            if (cboMaHD.Text == "")
            {
                MessageBox.Show("Bạn phải chọn một mã hóa đơn để tìm", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cboMaHD.Focus();
                return;
            }
            txtmahoadon.Text = cboMaHD.Text;
            LoadInfoHoaDon();
            LoadDataGridView();
            btnhuy.Enabled = true;
            btnluu.Enabled = true;
            btnin.Enabled = true;
            cboMaHD.SelectedIndex = -1;
        }

        private void btnin_Click(object sender, EventArgs e)
        {
            // Khởi động chương trình Excel
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable tblThongtinHD, tblThongtinHang;
            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            // Định dạng chung
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:Z300"].Font.Name = "Times new roman"; //Font chữ
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5; //Màu xanh da trời
            exRange.Range["A1:A1"].ColumnWidth = 7;
            exRange.Range["B1:B1"].ColumnWidth = 15;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "Cửa Hàng Xe Điện.";
            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "Bắc Từ Liêm - Hà Nội";
            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: 0869219002";
            exRange.Range["C2:E2"].Font.Size = 16;
            exRange.Range["C2:E2"].Font.Bold = true;
            exRange.Range["C2:E2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:E2"].MergeCells = true;
            exRange.Range["C2:E2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:E2"].Value = "HÓA ĐƠN BÁN";
            // Biểu diễn thông tin chung của hóa đơn bán
            sql = "SELECT a.MaHoaDon, a.NgayLap, a.TongTien, b.HoTen, b.DiaChi, b.SoDienThoai, c.HoTen FROM HoaDon AS a, KhachHang AS b, NhanVien AS c WHERE a.MaHoaDon = N'" + txtmahoadon.Text + "' AND a.MaKhachHang = b.MaKhachHang AND a.MaNhanVien = c.MaNhanVien";
            tblThongtinHD = Functions.GetDataToTable(sql);
            exRange.Range["B6:C9"].Font.Size = 12;
            exRange.Range["B6:B6"].Value = "Mã hóa đơn:";
            exRange.Range["C6:E6"].MergeCells = true;
            exRange.Range["C6:E6"].Value = tblThongtinHD.Rows[0][0].ToString();
            exRange.Range["B7:B7"].Value = "Khách hàng:";
            exRange.Range["C7:E7"].MergeCells = true;
            exRange.Range["C7:E7"].Value = tblThongtinHD.Rows[0][3].ToString();
            exRange.Range["B8:B8"].Value = "Địa chỉ:";
            exRange.Range["C8:E8"].MergeCells = true;
            exRange.Range["C8:E8"].Value = tblThongtinHD.Rows[0][4].ToString();
            exRange.Range["B9:B9"].Value = "Điện thoại:";
            exRange.Range["C9:E9"].MergeCells = true;
            exRange.Range["C9:E9"].Value = tblThongtinHD.Rows[0][5].ToString();
            //Lấy thông tin các mặt hàng
            sql = "SELECT b.TenXe, a.SoLuong, b.GiaBan, a.GiamGia, a.ThanhTien " +
                  "FROM ChiTietHoaDon AS a , XeDien AS b WHERE a.MaHoaDon = N'" +
                  txtmahoadon.Text + "' AND a.MaXe = b.MaXe";
            tblThongtinHang = Functions.GetDataToTable(sql);
            //Tạo dòng tiêu đề bảng
            exRange.Range["A11:F11"].Font.Bold = true;
            exRange.Range["A11:F11"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C11:F11"].ColumnWidth = 12;
            exRange.Range["A11:A11"].Value = "STT";
            exRange.Range["B11:B11"].Value = "Tên hàng";
            exRange.Range["C11:C11"].Value = "Số lượng";
            exRange.Range["D11:D11"].Value = "Đơn giá";
            exRange.Range["E11:E11"].Value = "Giảm giá";
            exRange.Range["F11:F11"].Value = "Thành tiền";
            for (hang = 0; hang < tblThongtinHang.Rows.Count; hang++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][hang + 12] = hang + 1;
                for (cot = 0; cot < tblThongtinHang.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 12
                {
                    exSheet.Cells[cot + 2][hang + 12] = tblThongtinHang.Rows[hang][cot].ToString();
                    if (cot == 3) exSheet.Cells[cot + 2][hang + 12] = tblThongtinHang.Rows[hang][cot].ToString() + "%";
                }
            }
            exRange = exSheet.Cells[cot][hang + 14];
            exRange.Font.Bold = true;
            exRange.Value2 = "Tổng tiền:";
            exRange = exSheet.Cells[cot + 1][hang + 14];
            exRange.Font.Bold = true;
            exRange.Value2 = tblThongtinHD.Rows[0][2].ToString();
            exRange = exSheet.Cells[1][hang + 15]; //Ô A1 
            exRange.Range["A1:F1"].MergeCells = true;
            exRange.Range["A1:F1"].Font.Bold = true;
            exRange.Range["A1:F1"].Font.Italic = true;
            exRange.Range["A1:F1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignRight;
            exRange.Range["A1:F1"].Value = "Bằng chữ: " + Functions.ChuyenSoSangChuoi(double.Parse(tblThongtinHD.Rows[0][2].ToString()));
            exRange = exSheet.Cells[4][hang + 17]; //Ô A1 
            exRange.Range["A1:C1"].MergeCells = true;
            exRange.Range["A1:C1"].Font.Italic = true;
            exRange.Range["A1:C1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            DateTime d = Convert.ToDateTime(tblThongtinHD.Rows[0][1]);
            exRange.Range["A1:C1"].Value = "Hà Nội, ngày " + d.Day + " tháng " + d.Month + " năm " + d.Year;
            exRange.Range["A2:C2"].MergeCells = true;
            exRange.Range["A2:C2"].Font.Italic = true;
            exRange.Range["A2:C2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:C2"].Value = "Nhân viên bán hàng";
            exRange.Range["A6:C6"].MergeCells = true;
            exRange.Range["A6:C6"].Font.Italic = true;
            exRange.Range["A6:C6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A6:C6"].Value = tblThongtinHD.Rows[0][6];
            exSheet.Name = "Hóa đơn bán hàng";
            exApp.Visible = true;
        }

        private void btndong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtsoluong_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar >= '0') && (e.KeyChar <= '9')) || (Convert.ToInt32(e.KeyChar) == 8))
                e.Handled = false;
            else e.Handled = true;
        }
    }
}
