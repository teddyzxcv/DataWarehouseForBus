using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Linq;
using System.Data;


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
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string csvstring = File.ReadAllText(filepath, Encoding.GetEncoding("GB2312"));
            using (var csvReader = new StringReader(csvstring))
            using (var tfp = new NotVisualBasic.FileIO.CsvTextFieldParser(csvReader))
            {
                // Get Some Column Names
                if (!tfp.EndOfData)
                {
                    string[] fields = tfp.ReadFields();
                }
                // Get Remaining Rows
                while (!tfp.EndOfData)
                {
                    string[] fields = tfp.ReadFields();
                    fields = fields[0].Split(';');
                    Cloth newCloth = new Cloth(fields[0], fields[1], fields[2], double.Parse(fields[3]), double.Parse(fields[4]), double.Parse(fields[5]), fields[6]);
                    output.Add(newCloth);
                }
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
