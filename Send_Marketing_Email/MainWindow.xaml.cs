using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Net.Mail;

namespace Send_Marketing_Email
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string connectionstring = "Server = 172.23.4.250; Database=MarketingEmail;User Id = sa;Password=p@ssw0rd;";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlCommand com = new SqlCommand();
                com.Connection = conn;
                com.CommandText = "Select fromemail from tblmailfrom";
                SqlDataReader dr = com.ExecuteReader();
                while (dr.Read())
                {
                    cboFrom.Items.Add(dr[0]);

                }
                dr.Close();
                com.CommandText = "Select TABLE_NAME from INFORMATION_SCHEMA.TABLES where TABLE_NAME like 'tbl_addr_%'";
                dr = com.ExecuteReader();
                while (dr.Read())
                {
                    cboTo.Items.Add(dr[0]);

                }
                dr.Close();
            }

        }

        private void cboFrom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lblFrom.Content = "";
            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlCommand com = new SqlCommand();
                com.Connection = conn;
                com.CommandText = "Select name from tblmailfrom where fromemail=@fromemail";
                com.Parameters.AddWithValue("@fromemail", cboFrom.SelectedValue.ToString());
                SqlDataReader dr = com.ExecuteReader();
                while (dr.Read())
                {
                    lblFrom.Content = dr[0];

                }
                dr.Close();
            }
        }

        private void cmdSend_Click(object sender, RoutedEventArgs e)
        {
            MailMessage mail = new MailMessage(cboFrom.SelectedValue.ToString(), "siuming@bpohk.com");
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Host = "172.23.4.5";
            mail.Subject = txtSubject.Text;
            mail.Body = txtBody.Text;
            mail.IsBodyHtml = true;
            client.Send(mail);
            MessageBox.Show("Done!");

        }
    }
}
