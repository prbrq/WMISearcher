using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WMITest2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Сканирование запущено
        /// </summary>
        bool IsRunning = false;

        CancellationTokenSource cts = new CancellationTokenSource();

        [STAThread]
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (IsRunning)
                {
                    button2.Enabled = false;
                }
                else
                {
                    button2.Enabled = true;

                    Cursor = Cursors.WaitCursor;

                    //отображение информации о запущенных процессах
                    dataGridView1.Columns.Clear();
                    dataGridView1.Columns.Add("ProcessId", "ProcessId");
                    dataGridView1.Columns.Add("Name", "Name");
                    dataGridView1.Columns.Add("WindowsVersion", "WindowsVersion");
                    dataGridView1.Columns.Add("VirtualSize", "VirtualSize");
                    dataGridView1.Columns.Add("CSName", "CSName");

                    //отображение информации о найденных в сети компьютерах
                    dataGridView2.Columns.Clear();
                    dataGridView2.Columns.Add("CSName", "CSName");
                    dataGridView2.Columns.Add("IP", "IP");
                    dataGridView2.Columns.Add("Description", "Description");
                    dataGridView2.Columns.Add("MACAddress", "MACAddress");

                    //в случае пустого адреса поле заполнится IP-адресом локального компьютера
                    if (string.IsNullOrEmpty(textBox1.Text))
                    {
                        textBox1.Text = "127.0.0.1";
                    }

                    //обработка введённого в textBox1.text диапазона адрессов, которые хранятся в переменной addresses  
                    string diapason = textBox1.Text;
                    var sae = diapason.Split('-').ToList();
                    if (sae.Count == 1)
                    {
                        sae.Add(sae[0]);
                    }
                    string diap1 = sae[0];
                    string diap2 = sae[1];
                    var array1 = diap1.Split('.');
                    var array2 = diap2.Split('.');

                    int start = BitConverter.ToInt32(new byte[]
                    {
                    Convert.ToByte(array1[3]),
                    Convert.ToByte(array1[2]),
                    Convert.ToByte(array1[1]),
                    Convert.ToByte(array1[0])
                    }, 0);
                    int end = BitConverter.ToInt32(new byte[]
                    {
                    Convert.ToByte(array2[3]),
                    Convert.ToByte(array2[2]),
                    Convert.ToByte(array2[1]),
                    Convert.ToByte(array2[0])
                    }, 0);

                    List<IPAddress> addresses = new List<IPAddress>();
                    for (int i = start; i <= end; i++)
                    {
                        byte[] bytes = BitConverter.GetBytes(i);
                        addresses.Add(new IPAddress(new[] { bytes[3], bytes[2], bytes[1], bytes[0] }));
                    }

                    foreach (var address in addresses)
                    {
                        if (TryPing(address))
                        {
                            GetWmi(address);
                        }
                        else
                        {
                            toolStripStatusLabel1.Text = $"{address} not available";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                button2.Enabled = IsRunning = false;
            }
        }

        private void GetWmi(IPAddress address)
        {
            

            //подключение к компьютеру в сети
            ManagementScope scope = new ManagementScope($@"\\{ address }\root\cimv2");
            var username = textBox2.Text;
            var password = textBox3.Text;
            if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Username = $@"{ textBox2.Text }";
                options.Password = $@"{ textBox3.Text }";
                scope.Options = options;
            }
            scope.Connect();
            
            var task1 = Task.Factory.StartNew(async delegate
            {
                // вывод в таблицу №2 информации о компьютерах в сети
                const string QueryN = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True";
                ManagementObjectSearcher networkEnumerator = new ManagementObjectSearcher(QueryN);
                networkEnumerator.Scope = scope;
                foreach (ManagementObject NetEd in networkEnumerator.Get())
                {
                    var Hname = NetEd["DNSHostName"];
                    var Desc = NetEd["Description"];
                    var array = NetEd["IPAddress"] as string[];
                    var IPA = string.Join(", ", array);
                    var MACA = NetEd["MACAddress"];
                    await Task.Run(delegate
                    {
                        dataGridView2.Invoke(new Action(delegate
                        {
                            dataGridView2.Rows.Add(Hname, IPA, Desc, MACA);
                        }));
                    });
                }
            }, cts);

            var task2 = Task.Factory.StartNew(async delegate
            {
                //вывод в таблицу №1 информации о процессах
                const string Query = "SELECT * FROM Win32_Process";
                ManagementObjectSearcher processEnumerator = new ManagementObjectSearcher(Query);
                processEnumerator.Scope = scope;
                foreach (ManagementObject process in processEnumerator.Get())
                {
                    var pid = process["ProcessId"];
                    var name = process["Name"];
                    var winver = process["WindowsVersion"];
                    var vsize = process["VirtualSize"];
                    var csn = process["CSName"];

                    await Task.Run(delegate
                    {
                        dataGridView1.Invoke(new Action(delegate
                        {
                            dataGridView1.Rows.Add(pid, name, winver, vsize, csn);
                        }));
                    });

                    if (checkBox1.Checked)
                    {
                        //подключение к базе данных
                        SqlConnection conn = new SqlConnection();
                        conn.ConnectionString = @"Server=localhost\SQLEXPRESS;Database=WMIDB;Trusted_Connection=True;";
                        conn.Open();

                        //очистка базы данных
                        const string QueryD = "DELETE FROM dbo.Win32_Process";
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = QueryD;
                            cmd.ExecuteNonQuery();
                        }

                        //добавление информации о полученных процессах в базу данных
                        const string Query1 = "INSERT INTO dbo.Win32_Process VALUES (@pid, @name, @winver, @vsize, @csn)";
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = Query1;
                            cmd.Parameters.Add(new SqlParameter("@pid", SqlDbType.NVarChar) { Value = pid });
                            cmd.Parameters.Add(new SqlParameter("@name", SqlDbType.NVarChar) { Value = name });
                            cmd.Parameters.Add(new SqlParameter("@winver", SqlDbType.NVarChar) { Value = winver });
                            cmd.Parameters.Add(new SqlParameter("@vsize", SqlDbType.NVarChar) { Value = vsize });
                            cmd.Parameters.Add(new SqlParameter("@csn", SqlDbType.NVarChar) { Value = csn });
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        var row = $"{pid};{name};{winver};{vsize};{csn}";
                        using (var writer = new StreamWriter("output.csv", true, Encoding.UTF8))
                        {
                            writer.WriteLine($@"{DateTime.Now}: {row}");
                        }
                    }
                }
            }, cts);

            Task.WaitAll(task1, task2);
        }

        // функция, используемая для проверки доступности IP-адресов
        public static bool TryPing(IPAddress address)
        {
            Ping pingSender = new Ping();
            PingReply reply = pingSender.Send(address);
            return reply.Status == IPStatus.Success;
        }

        private void OnRowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var dgv = sender as DataGridView;
            var ri = e.RowIndex;
            if (ri > -1 && ri < dgv.RowCount)
            {   // Черезстрочно закрашиваем в светлосерый
                dgv.Rows[ri].DefaultCellStyle.BackColor = (ri & 1) == 0 ? Color.LightGray : Color.LightGreen;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cts.Cancel();
        }

    }
}
