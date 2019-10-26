using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Media;
using System.Net.Mail;
using System.Net;


namespace _Odev
{
    public partial class Form1 : Form
    {
        OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.Ace.OleDb.12.0;Data Source=rehber.accdb");
        DataTable tablo = new DataTable();
        OleDbCommand kmt = new OleDbCommand();
        bool duzen = false;

        public void listele()
        {
            tablo.Clear();
            OleDbDataAdapter adtr = new OleDbDataAdapter("Select * From rehber", baglanti);
            adtr.Fill(tablo);
            dataGridView1.DataSource = tablo;
        }

        private void dtaGridViewHazirla() {
            dataGridView1.ColumnHeadersVisible = false;
            dataGridView1.GridColor = Color.White;
            dataGridView1.RowTemplate.Height = 40;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            listele();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.Columns[1].Width = 190;

            DataGridViewButtonColumn dgBtnCol = new DataGridViewButtonColumn();

            dgBtnCol.Text = "Düzenle";
            dgBtnCol.Name = "dtBtDuzenle";
            dgBtnCol.UseColumnTextForButtonValue = true;
            dgBtnCol.Width = 77;
            dataGridView1.Columns.Add(dgBtnCol);

            DataGridViewButtonColumn dgBtnCol2 = new DataGridViewButtonColumn();

            dgBtnCol2.Text = "Sil";
            dgBtnCol2.Name = "dtBtSil";
            dgBtnCol2.UseColumnTextForButtonValue = true;
            dgBtnCol2.Width = 77;
            dataGridView1.Columns.Add(dgBtnCol2);
            dataGridView1.Rows[0].Selected = true;
        }

        private void kayitSil(int id) {
            baglanti.Open();
            kmt.Connection = baglanti;
            kmt.CommandText = "DELETE FROM rehber WHERE id= " + id + "";
            kmt.ExecuteNonQuery();
            baglanti.Close();
            listele();
        }

        private void kayitDuzenle()
        {
            panel3.Visible = true;
            this.Height = 280;
            txtİsim.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txtCep.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            txtEv.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txtEposta.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            txtAdres.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
        }

        private void kayitDuzenle(int id) {

            if (txtİsim.Text != "" && txtCep.Text != "" && txtEposta.Text != "")
            {
                baglanti.Open();
                kmt.Connection = baglanti;
                kmt.CommandText = " UPDATE rehber SET isim='" + txtİsim.Text + "',cep='" + txtCep.Text + "',ev='" + txtEv.Text + "',email='" + txtEposta.Text + "',adres='" + txtAdres.Text + "' WHERE id =" + id + "";
                kmt.ExecuteNonQuery();
                baglanti.Close();

                panel3.Visible = true;
                this.Height = 280;
                MessageBox.Show("Düzenleme İşlemi Tamamlandı", "Bilgi");
                txtİsim.Text = "";
                txtCep.Text = "";
                txtEv.Text = "";
                txtEposta.Text = "";
                txtAdres.Text = "";
                txtBul.Text = "";
                listele();
                panel3.Visible = false;
                this.Height = 507;
            }
            else { MessageBox.Show("İsim, Cep Telefonu ve Eposta alanları boş bırakılamaz", "Hata"); txtİsim.Focus(); }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dtaGridViewHazirla();  
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 6){
                duzen = true;
                kayitDuzenle();
            }
            else if (e.ColumnIndex == 7) {
               int a = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
               kayitSil(a);
            }
               
            
        }

        private void btnAra_Click(object sender, EventArgs e)
        {
            SoundPlayer player = new SoundPlayer();
            player.SoundLocation = "../Debug/music/1.wav";
            player.Play();
            MessageBox.Show("Aranıyor..." + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "\n" + dataGridView1.CurrentRow.Cells[2].Value.ToString());
            player.Stop();
        }

        private void btnMesaj_Click(object sender, EventArgs e)
        {
            SoundPlayer player = new SoundPlayer();
            player.SoundLocation = "../Debug/music/2.wav";
            player.Play();
            MessageBox.Show("Mesaj Gönderiliyor..." + dataGridView1.CurrentRow.Cells[1].Value.ToString() + "\n" + dataGridView1.CurrentRow.Cells[2].Value.ToString());
            player.Stop();
        }

        private void txtBul_TextChanged(object sender, EventArgs e)
        {
            if (txtBul.Text.Trim() == "")
            {
                listele();
            }
            else
            {
                tablo.Clear();
                OleDbDataAdapter adtr = new OleDbDataAdapter("Select * from rehber where isim Like '%" + txtBul.Text + "%'", baglanti);
                adtr.Fill(tablo);
                dataGridView1.DataSource = tablo;
            }
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            this.Height = 380;
            panel4.Visible = true;
            txtKime.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
        }

        private void btnEkle_Click(object sender, EventArgs e)
        {
            duzen = false;
            panel3.Visible = true;
            this.Height = 280;
        }

        private void btnIptal_Click(object sender, EventArgs e)
        {
            txtİsim.Text = "";
            txtCep.Text = "";
            txtEv.Text = "";
            txtEposta.Text = "";
            txtAdres.Text = "";
            txtBul.Text = "";
            listele();
            panel3.Visible = false;
            this.Height = 507;
        }

        private void btnKaydet_Click(object sender, EventArgs e)
        {
            if (duzen == true) {
                int a = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                kayitDuzenle(a);
                return;
            }
            if (txtİsim.Text != "" && txtCep.Text != "" && txtEposta.Text != "")
            {
                baglanti.Open();
                kmt.Connection = baglanti;
                kmt.CommandText = " INSERT INTO rehber (isim,cep,ev,email,adres) VALUES ('"+txtİsim.Text+"','"+txtCep.Text+"','"+txtEv.Text+"','"+txtEposta.Text+"','"+txtAdres.Text+"')";
                kmt.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Kayıt İşlemi Tamamlandı","Bilgi");
                txtİsim.Text = "";
                txtCep.Text = "";
                txtEv.Text = "";
                txtEposta.Text = "";
                txtAdres.Text = "";
                txtBul.Text = "";
                listele();
                panel3.Visible = false;
                this.Height = 507;
            }
            else { MessageBox.Show("İsim, Cep Telefonu ve Eposta alanları boş bırakılamaz", "Hata"); txtİsim.Focus(); }
        }

        private void btnCikis_Click(object sender, EventArgs e)
        {
            txtKime.Text = "";
            txtKonu.Text = "";
            txtIcerik.Text = "";
            txtKimden.Text = "";
            txtSifre.Text = "";
            panel4.Visible = false;
            this.Height = 507;
        }

        private void btnGonder_Click(object sender, EventArgs e)
        {
             MailMessage mail = new MailMessage();

            mail.From = new MailAddress(txtKimden.Text); //Zamanım az oldugu  için kimden kısmını gönderenin eposta adresine ayarladım.

            mail.To.Add(txtKime.Text); 

            mail.Subject = txtIcerik.Text;

            mail.Body = txtIcerik.Text;

            SmtpClient smtp = new SmtpClient("smtp.live.com", 587); // hotmail üzerinden gönderileceğinden smtp.live.com ve onun 587 nolu portu kullanılır.

            smtp.Credentials = new NetworkCredential(txtKimden.Text, txtSifre.Text); //hangi e-posta üzerinden gönderileceği. E posta, şifre'si yazılır.

            smtp.EnableSsl = true;
            try
            {
                smtp.Send(mail);
                txtKime.Text = "";
                txtKonu.Text = "";
                txtIcerik.Text = "";
                txtKimden.Text = "";
                txtSifre.Text = "";
                panel4.Visible = false;
                this.Height = 507;
                MessageBox.Show("E-Posta gönderildi","Bilgi");
            }
            catch
            {
                MessageBox.Show("E-Postaz gönderilemedi","Hata");
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Height = 380;
            panel4.Visible = true;
            txtKime.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            panel5.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            duzen = true;
            kayitDuzenle();
            panel5.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            kayitSil(a);
            panel5.Visible = false;
            this.Height = 507;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel5.Visible = false;
            this.Height = 507;
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            panel5.Visible = true;
            this.Height = 300;
            lblisim.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            lblCep.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            lblEv.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            lblEposta.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            lblAdres.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
        }

      

      

       

       
    }
}
