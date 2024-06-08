using ImapX.Collections;
using ImapX.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApp19
{
    /// <summary>
    /// Логика взаимодействия для Sending.xaml
    /// </summary>
    public partial class Sending : Window
    {
        private string path;
        public Sending(string path)
        {
            InitializeComponent();
            this.path = path;
        }

        private void SendBtm_Click(object sender, RoutedEventArgs e)
        {
            int port = 25;
            if (YourMailTbox.Text.Contains("@gmail.com") == true)
            {
                string protokol = "gmail.com";
                Send(protokol, port);
            }
            else if (YourMailTbox.Text.Contains("@mail.ru") == true)
            {
                string protokol = "mail.ru";
                port = 587;
                Send(protokol, port);
            }
            else if (YourMailTbox.Text.Contains("@rambler.ru") == true)
            {
                string protokol = "rambler.ru";
                Send(protokol, port);
            }
            else if (YourMailTbox.Text.Contains("@yandex.ru") == true)
            {
                string protokol = "yandex.ru";
                Send(protokol, port);
            }
            else
            {
                MessageBox.Show("Введён недоступный протокол");
            }
        }

        private void Send(string protokol, int port)
        {
            MailMessage mailMessage = new MailMessage(YourMailTbox.Text, FriendMailTbox.Text, ThemeTbox.Text, "отправила файлик с помощью приложения)))");
            mailMessage.Attachments.Add(new Attachment(path));
            SmtpClient smtpClient = new SmtpClient("smtp." + protokol, port);
            smtpClient.Credentials = new NetworkCredential(YourMailTbox.Text, PasswordTbox.Text);
            smtpClient.EnableSsl = true;
            smtpClient.Send(mailMessage);
            MessageBox.Show("Сообщение отправилось!");
        }
    }
}
