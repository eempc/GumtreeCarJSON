using System;
using System.Collections.Specialized;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using HtmlAgilityPack;
using System.Collections;

namespace GumtreeCarJSON {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            InitFBD();
            InitString();
        }

        OrderedDictionary currentCar = new OrderedDictionary();
        string[] attributes;
        string saveFile = @"D:\Temp\Car Database\Car v2.csv";

        public void InitString() {
            attributes = new string[] { "vrn", "vehicle_make", "vehicle_model", "vehicle_registration_year", "seller_type", "vehicle_mileage", "vehicle_colour", "price", "vehicle_not_writeoff", "vehicle_vhc_checked", "URL", "Location", "MOT expiry" };
            foreach (string a in attributes)
                currentCar.Add(a, "");
        }    

        public void InitFBD() {
            openFileDialog1.Filter = "HTML (*.html;*.htm)|*.html;*.htm|" + "All files (*.*)|*.*";
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = "HTML file browser";
        }

        public void GumtreeCarHTMLExtract() {
            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                foreach (string fileName in openFileDialog1.FileNames) {
                    //Retrieve JSON
                    string fullHTML = File.ReadAllText(fileName);
                    string dataLayer = ExtractRegex(fullHTML, "var dataLayer = ", "}];") + "}]";
                    dynamic obj = JsonConvert.DeserializeObject(dataLayer);

                    //Unfortunately I do not know if there is a way to turn a string into the json .dot thing, if I could, then the following would be a one line loop
                    //Maybe in a future commit
                    currentCar[attributes[0]] = obj[0].a.attr.vrn;
                    currentCar[attributes[1]] = obj[0].a.attr.vehicle_make;
                    currentCar[attributes[2]] = obj[0].a.attr.vehicle_model;
                    currentCar[attributes[3]] = obj[0].a.attr.vehicle_registration_year;
                    currentCar[attributes[4]] = obj[0].a.attr.seller_type;
                    currentCar[attributes[5]] = obj[0].a.attr.vehicle_mileage;
                    currentCar[attributes[6]] = obj[0].a.attr.vehicle_colour;
                    currentCar[attributes[7]] = obj[0].a.attr.price;
                    currentCar[attributes[8]] = obj[0].a.attr.vehicle_not_writeoff;
                    currentCar[attributes[9]] = obj[0].a.attr.vehicle_vhc_checked;

                    //Retrieve URL and location from downloaded HTML file on local drive
                    currentCar[attributes[10]] = ExtractRegex(fullHTML, "https://", " -->");
                    currentCar[attributes[11]] = ExtractLocation(fullHTML).Replace(",", "");

                    WriteToCSV();
                }
            }
        }

        //ClosedXML is huge! Best write to CSV instead for the moment
        public void WriteToCSV() {
            if (!Directory.Exists(Path.GetDirectoryName(saveFile))) 
                Directory.CreateDirectory(Path.GetDirectoryName(saveFile));

            if (!File.Exists(saveFile)) {
                //Create and write header
                using (StreamWriter writer = File.CreateText(saveFile)) {
                    foreach (string a in attributes)
                        writer.Write(a + ",");
                    writer.Write("\n");
                }               
            }

            using (StreamWriter writer = File.AppendText(saveFile)) {
                foreach (DictionaryEntry x in currentCar) 
                    writer.Write(x.Value.ToString() + ",");
                writer.Write("\n");
            }
        }

        public string ExtractLocation(string source) {
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(source);
            var singleNode = doc.DocumentNode.SelectSingleNode("//body/div/div/div/main/div/header/strong/span");
            return singleNode.InnerText.ToString();
        }

        public string ExtractRegex(string source, string start, string end) {
            Regex rx = new Regex(start + "(.*?)" + end);
            var match = rx.Match(source);
            return match.Groups[1].ToString();
        }

        private void button1_Click(object sender, EventArgs e) {
            GumtreeCarHTMLExtract();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) {
            Application.Exit();
        }

    }
}
