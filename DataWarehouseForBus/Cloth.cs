using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Linq;
using System.Data;
using Aspose.Cells;


namespace DataWarehouseForBus
{
    [Serializable]
    [XmlRoot("ClothProduct")]
    public class Cloth
    {
        [XmlElement("ID")]
        public int Id;
        [XmlElement("Article")]
        public string Article;
        [XmlElement("Manufacturer")]

        public string Manufacturer;
        [XmlElement("Name")]

        public string ClothName;
        [XmlElement("Weight")]

        public double Weight;
        [XmlElement("InPrice")]

        public double InPrice;
        [XmlElement("OutPrice")]
        public double OutPrice;
        [XmlElement("Description")]
        public string Description;

        public Cloth(string article, string manufacturer, string clothname, double weight, double inprice, double outprice, string description)
        {
            Article = article;
            Manufacturer = manufacturer;
            ClothName = clothname;
            Weight = weight;
            InPrice = inprice;
            OutPrice = outprice;
            Description = description;
        }
        public Cloth()
        {

        }
        public static List<Cloth> ClothParser(string filepath)
        {
            List<Cloth> output = new List<Cloth>();
            LoadOptions loadOptions2 = new LoadOptions(LoadFormat.Xlsx);
            Workbook wb = new Workbook(filepath, loadOptions2);
            var dt = wb.Worksheets[0].Cells.ExportDataTableAsString(1, 0, wb.Worksheets[0].Cells.MaxDataRow, wb.Worksheets[0].Cells.MaxDataColumn + 1);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var article = dt.Rows[i][0].ToString();
                var manufactor = dt.Rows[i][1].ToString();
                var name = dt.Rows[i][2].ToString();
                var weight = double.Parse(dt.Rows[i][3].ToString());
                var inp = double.Parse(dt.Rows[i][4].ToString());
                var outp = double.Parse(dt.Rows[i][5].ToString());
                var des = dt.Rows[i][6].ToString();
                output.Add(new Cloth(article, manufactor, name, weight, inp, outp, des));
            }
            return output;
        }
        public override bool Equals(object obj)
        {
            var cloth = (Cloth)obj;

            return cloth.Article == this.Article && cloth.Manufacturer == this.Manufacturer;
        }
        public override int GetHashCode()
        {
            return this.Article.GetHashCode() * 33 + this.Manufacturer.GetHashCode() * 66;
        }

    }
}
