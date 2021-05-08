using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Aspose.Cells;
namespace DataWarehouseForBus
{
    public partial class Form1 : Form
    {
        static List<Cloth> DataBase = new List<Cloth>();
        static string FilePath;
        public Form1()
        {
            InitializeComponent();
            this.dataGridView1.ContextMenuStrip = this.contextMenuStrip2;

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
            ofd.Filter = "XLSX 文件 (*.xlsx)|*.xlsx";
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
            XmlSer.SerializationToXml(DataBase, FilePath);
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
            XmlSer.SerializationToXml(DataBase, FilePath);
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

        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.Filter = "XML 文件 (*.xml)|*.xml";
            if (sfd.ShowDialog() != DialogResult.OK)
            {
                this.Close();
                return;
            }
            FilePath = sfd.FileName;
            DataBase = XmlSer.DeserializationFromXml(FilePath);
            RefreshDataGrid(DataBase);
            this.Activate();

        }

        private void exportStripMenuItem1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XLSX 文件 (*.xlsx)|*.xlsx";
            var result = sfd.ShowDialog();

            if (result == DialogResult.OK)
            {
                ExportExcel((DataTable)this.dataGridView1.DataSource, sfd.FileName, new List<string>() { "款号", "厂家", "名称", "重量", "进价", "售价", "描述" });
            }
        }

        private static string ExportExcel(DataTable dt, string filename, List<string> listHeader)
        {
            try
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(); //Workbook
                Aspose.Cells.Worksheet sheet = workbook.Worksheets[0]; //Worksheet
                Aspose.Cells.Cells cells = sheet.Cells;//Cell
                Style style;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    cells[0, j].PutValue(listHeader[j]);

                    style = cells[0, j].GetStyle();
                    style.BackgroundColor = Color.Blue;
                    style.ForegroundColor = System.Drawing.Color.FromArgb(153, 204, 0);
                    style.Pattern = BackgroundType.Solid;
                    cells[0, j].SetStyle(style);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        cells[i + 1, j].PutValue(dt.Rows[i][j].ToString());
                    }
                }
                sheet.AutoFitColumns();
                cells.SetRowHeight(0, 30);
                workbook.Save(filename);

                return "";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        private void saveStripMenuItem1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XML 文件 (*.xml)|*.xml";
            var result = sfd.ShowDialog();
            if (result == DialogResult.OK)
            {
                XmlSer.SerializationToXml(DataBase, sfd.FileName);
            }
        }

        private void deleteStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtb = (DataTable)this.dataGridView1.DataSource;
                if (this.dataGridView1.SelectedRows.Count == 0)
                    return;
                string article = dtb.Rows[this.dataGridView1.SelectedRows[0].Index]["款号"].ToString();
                string manufacturer = dtb.Rows[this.dataGridView1.SelectedRows[0].Index]["厂家"].ToString();
                Cloth todelete = DataBase.Where(e => e.Article == article).Where(e => e.Manufacturer == manufacturer).First();
                DataBase.Remove(todelete);
                RefreshDataGrid(DataBase);
                XmlSer.SerializationToXml(DataBase, FilePath);
            }
            catch
            {
                return;
            }
        }
    }
}
