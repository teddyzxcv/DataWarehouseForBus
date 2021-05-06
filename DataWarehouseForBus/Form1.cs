using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataWarehouseForBus
{
    public partial class Form1 : Form
    {
        static List<Cloth> DataBase = new List<Cloth>();
        public Form1()
        {
            InitializeComponent();
            DataBase = XmlSer.DeserializationFromXml();
            RefreshDataGrid(DataBase);
        }
        private void RefreshDataGrid(List<Cloth> db)
        {
            DataTable dft = new DataTable();
            dft.Columns.Add("款号");
            dft.Columns.Add("厂家");
            dft.Columns.Add("名称");
            dft.Columns.Add("重量");
            dft.Columns.Add("进价");
            dft.Columns.Add("售价");
            dft.Columns.Add("描述");
            foreach (var item in db)
            {
                var row = dft.NewRow();
                row["款号"] = item.Article;
                row["厂家"] = item.Manufacturer;
                row["名称"] = item.ClothName;
                row["重量"] = item.Weight;
                row["进价"] = item.InPrice;
                row["售价"] = item.OutPrice;
                row["描述"] = item.Description;
                dft.Rows.Add(row);
            }
            this.dataGridView1.DataSource = dft;
            for (int i = 0; i < dft.Columns.Count; i++)
            {
                this.dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }

        }
        private void importStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV 文件 (*.csv)|*.csv";
            List<string> output = new List<string>();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                var clothimport = Cloth.ClothParser(ofd.FileName);
                clothimport = clothimport.Distinct().ToList();
                foreach (var item in clothimport)
                {
                    if (DataBase.Contains(item))
                        output.Add(item.Article);
                    else
                    {
                        if (DataBase.Count != 0)
                            item.Id = DataBase.Select(e => e.Id).Max() + 1;
                        else
                            item.Id = 1;
                        DataBase.Add(item);
                    }
                }
            }
            if (output.Count != 0)
                MessageBox.Show($"无法添加以下款号，因为他们已存在于数据库内: {String.Join(',', output)}", "错误！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            RefreshDataGrid(DataBase);
            XmlSer.SerializationToXml(DataBase);
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            var newdb = new List<Cloth>();
            DataTable db = (DataTable)(this.dataGridView1.DataSource);
            for (int i = 0; i < db.Rows.Count; i++)
            {
                var cloth = new Cloth((string)db.Rows[i]["款号"], (string)db.Rows[i]["厂家"], (string)db.Rows[i]["名称"], double.Parse((string)db.Rows[i]["重量"]), double.Parse((string)db.Rows[i]["进价"]), double.Parse((string)db.Rows[i]["售价"]), (string)db.Rows[i]["描述"]);
                newdb.Add(cloth);
            }
            DataBase = newdb;
            XmlSer.SerializationToXml(DataBase);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<Cloth> output = DataBase;
            string article = this.textBox1.Text;
            string manufactor = this.textBox2.Text;
            string price = this.textBox3.Text;
            string name = this.textBox4.Text;
            if (article != String.Empty)
                output = output.Where(e => e.Article == article).ToList();
            if (manufactor != String.Empty)
                output = output.Where(e => e.Manufacturer == manufactor).ToList();
            if (price != String.Empty)
                output = output.Where(e => e.OutPrice == double.Parse(price)).ToList();
            if (name != String.Empty)
                output = output.Where(e => e.ClothName == name).ToList();
            RefreshDataGrid(output);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RefreshDataGrid(DataBase);

        }
    }
}
