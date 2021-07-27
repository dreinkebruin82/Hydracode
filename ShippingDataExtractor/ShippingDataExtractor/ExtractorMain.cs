using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;

namespace ShippingDataExtractor
{
    public partial class ExtractorMain : Form
    {
        public ExtractorMain()
        {
            InitializeComponent();
        }

        #region Private Veriables
        private string loadNumberInitialText = "Load Number\n";
        private string tableInitialText = "Consigned to:";
        private string tableEndText = "Sub Total:\n";
        private string headerEndText = "\nSeal 2:\nSeal 1:\n";
        private string footerTextInitial = "UNLESS EXPRESSLY SUBJECT TO";
        private string manifestInitialText = "SHIP TO\n";
        private string bolEndText = "\nof";

        private string fileName = "";
        private string loadNumber = "";
        private string bolNumber = "";
        private string manifestNumber = "";
        private string orderNumber = "";
        private List<Record> records = null;
        #endregion

        #region Events

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                btnStart.Enabled = false;
                btnStart.Refresh();

                List<string> orderTableRetrieved = new List<string>();
                foreach (string path in openFileDialog1.FileNames)
                {
                    if (File.Exists(path))
                    {
                        try
                        {
                            List<string> textOfAllPages = ExtractPageTextFromPdf(path);

                            records = new List<Record>();
                            fileName = path.Substring(path.LastIndexOf(@"\") + 1);
                            loadNumber = GetLoadNumber(textOfAllPages);
                            bolNumber = GetBOL(textOfAllPages);
                            manifestNumber = GetManifestNumber(textOfAllPages);
                            string pageText = RemoveHeaderFooterAndJoinPages(textOfAllPages);
                            orderTableRetrieved = GetOrderTables(pageText);

                            foreach (string orderTable in orderTableRetrieved)
                            {
                                GetRecords(orderTable);
                            }
                            WriteRecordsInFile(fileName);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error during processing the file: " + fileName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("This path is not correct: " + path, "Invalid Path", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                if (openFileDialog1.FileNames.Length > 0 && orderTableRetrieved.Count > 0)
                {
                    MessageBox.Show("Process completed successfully!", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                btnStart.Enabled = true;
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "PDF|*.pdf";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtSelectedFile.Text = string.Join(Environment.NewLine, openFileDialog1.FileNames);
            }
            else
            {
                txtSelectedFile.Text = string.Empty;
            }
        }
        #endregion

        #region Priave Methods

        private List<string> ExtractPageTextFromPdf(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                List<string> pageTexts = new List<string>();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    pageTexts.Add(PdfTextExtractor.GetTextFromPage(reader, i));
                }
                return pageTexts;
            }
        }

        private string RemoveHeaderFooterAndJoinPages(List<string> textOfAllPages)
        {
            string textWithoutHeaderFooter = string.Empty;
            int pageNo = 1;
            foreach (string pageText in textOfAllPages)
            {
                if (pageNo++ == 1)
                {
                    continue;
                }
                string text = pageText;
                if (text.Contains(headerEndText))
                {
                    text = text.Substring(text.IndexOf(headerEndText) + headerEndText.Length);
                }
                if (text.Contains(footerTextInitial))
                {
                    text = text.Substring(0, text.IndexOf(footerTextInitial));
                }
                textWithoutHeaderFooter += text;
            }
            return textWithoutHeaderFooter;
        }

        private string GetLoadNumber(List<string> allPageTexts)
        {
            string pageText = string.Join("", allPageTexts);
            string loadNumber = string.Empty;
            if (pageText.Contains(loadNumberInitialText))
            {
                loadNumber = pageText.Substring(pageText.IndexOf(loadNumberInitialText) + loadNumberInitialText.Length, 15);
                loadNumber = loadNumber.Substring(0, loadNumber.IndexOf("\n"));
            }
            return loadNumber.Trim();
        }

        private string GetBOL(List<string> allPageTexts)
        {
            string pageText = string.Join("", allPageTexts);
            string bol = string.Empty;
            if (pageText.Contains(bolEndText))
            {
                bol = pageText.Substring(0, pageText.IndexOf(bolEndText));
                bol = bol.Substring(bol.LastIndexOf("\n") + 1);
                bol = bol.Contains(" ") ? bol.Substring(0, bol.IndexOf(" ")) : bol;
            }
            return bol.Trim();
        }

        private string GetManifestNumber(List<string> allPageTexts)
        {
            string pageText = string.Join("", allPageTexts);
            string text = string.Empty;
            string manifestNo = string.Empty;
            if (pageText.Contains(manifestInitialText))
            {
                text = pageText.Substring(pageText.IndexOf(manifestInitialText) + manifestInitialText.Length, 50);
                text = text.Substring(0, text.IndexOf("\n"));
                string[] values = text.Split((new string[] { " " }), StringSplitOptions.None);
                if (values.Length >= 2)
                {
                    manifestNo = values[values.Length - 2];
                }
            }
            return manifestNo.Trim();
        }

        private List<string> GetOrderTables(string pageText)
        {
            List<string> orderTables = new List<string>();
            string text = pageText;
            do
            {
                if (text.Contains(tableInitialText))
                {
                    int nextTableInitialIndex = text.IndexOf(tableInitialText, text.IndexOf(tableInitialText) + tableInitialText.Length);
                    int tableEndIndex = text.IndexOf(tableEndText);

                    string strippedText = "";
                    if (nextTableInitialIndex != -1 && nextTableInitialIndex < tableEndIndex)
                    {
                        strippedText = text.Substring(0, nextTableInitialIndex);
                    }
                    else
                    {
                        strippedText = text.Substring(0, text.IndexOf(tableEndText) + tableEndText.Length);
                    }
                    orderTables.Add(strippedText);
                    text = text.Substring(strippedText.Length);
                }
            } while (text.Contains(tableInitialText));

            return orderTables;
        }

        private void GetRecords(string orderTable)
        {
            orderNumber = GetOrderNumber(orderTable);
            string text = orderTable.Substring(FindStartIndexOfTableText(orderTable));

            string[] lines = text.Split((new string[] { "\n" }), StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                string[] values = lines[i].Split((new string[] { " " }), StringSplitOptions.None);
                if (values.Length >= 8)
                {
                    int weightIndex = GetIndexOfWeight(values);
                    if (!values[values.Length - 6].Contains(" ") && values[values.Length - 6].Length >= 10 && values[values.Length - 5].Contains(".") && values[values.Length - 1].Contains(","))
                    {
                        Record record = new Record();
                        record.LoadNumber = loadNumber;
                        record.BOL = bolNumber;
                        record.ManifestNumber = manifestNumber;
                        record.Order = orderNumber;

                        record.BasicWgt = values[weightIndex].Trim();
                        record.RollNumber = values[values.Length - 6].Trim();
                        record.Diameter = values[values.Length - 5].Trim();
                        record.Width = values[values.Length - 4].Trim();
                        record.LinerFeet = values[values.Length - 3].Trim();
                        record.SquareFeet = values[values.Length - 2].Trim();
                        record.Weight = values[values.Length - 1].Trim();

                        // Description
                        if (weightIndex == values.Length - 7)
                        {
                            record.ProductDescription = lines[++i].Trim(); // Getting next 
                        }
                        else if (weightIndex < values.Length - 7)
                        {
                            string productDescription = "";
                            for (int p = weightIndex + 1; p < values.Length - 6; p++)
                            {
                                productDescription += values[p] + " ";
                            }
                            record.ProductDescription = productDescription.Trim();
                        }
                        if (record.ProductDescription.Contains("#"))
                        {
                            record.ProductDescription = record.ProductDescription.Substring(record.ProductDescription.IndexOf("#") + 1).Trim();
                        }
                        // Customer PO
                        string customerPO = string.Empty;
                        for (int p = 0; p < weightIndex; p++)
                        {
                            customerPO += values[p] + " ";
                        }
                        record.CustomerPO = customerPO.Trim();

                        records.Add(record);
                    }
                }
            }
        }

        private int GetIndexOfWeight(string[] values)
        {
            int index = -1;
            double weight = 0;
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i].Contains("."))
                {
                    if (double.TryParse(values[i], out weight))
                    {
                        index = i;
                        break;
                    }
                }
            }
            return index;
        }

        private string GetOrderNumber(string orderTable)
        {
            string orderNumber = "";
            if (orderTable.Contains("Order:"))
            {
                orderNumber = orderTable.Substring(orderTable.IndexOf("Order:"), orderTable.IndexOf("\n", orderTable.IndexOf("Order:")) - orderTable.IndexOf("Order:"));
                orderNumber = orderNumber.Replace("Order:", "").Trim();
                if (orderNumber.Contains(" "))
                {
                    orderNumber = orderNumber.Substring(0, orderNumber.IndexOf(" ")).Trim();
                }
            }
            return orderNumber;
        }

        private int FindStartIndexOfTableText(string orderTable)
        {
            int maxIndex = 0;
            maxIndex = orderTable.LastIndexOf("(inch)\n") > maxIndex ? orderTable.LastIndexOf("(inch)\n") + "(inch)\n".Length : maxIndex;
            maxIndex = orderTable.LastIndexOf("Feet\n") > maxIndex ? orderTable.LastIndexOf("Feet\n") + "Feet\n".Length : maxIndex;
            maxIndex = orderTable.LastIndexOf("(lbs)\n") > maxIndex ? orderTable.LastIndexOf("(lbs)\n") + "(lbs)\n".Length : maxIndex;
            return maxIndex;
        }

        #endregion

        #region Write records into file

        private void WriteRecordsInFile(string fileName)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + DateTime.Now.ToString("MMM-dd-yyyy") + " - Extracted Data\\";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            fileName = fileName.Replace(".pdf", ".txt");

            //string header = "LOAD\tManifest\tORDER\tCustomer PO\tBasic Wgt (lbs)\tProduct Description\tRoll Number\tDiameter (inch)\tWidth (inch)\tLineal Feet\tSquare Feet\tWeight (lbs)\n";
            string header = "PG1/Product Description\tPG2/Width\tPG3/Customer PO\tPG4/Load\tPG5Manifest\tItem ID\tWeight\tCt in Pkg\tLocation\tLB/FT\tLNFT\n";
            string data = string.Empty;
            foreach (Record item in records)
            {
                data += item.BasicWgt + " " + item.ProductDescription + "\t"; //PG1/Product Description
                data += item.Width + "\t";          //PG2/Width
                data += item.CustomerPO + "\t";     //PG3/Customer PO
                data += item.BOL + "\t";            //PG4/Load ==> BOL#
                data += item.ManifestNumber + "\t"; //PG5Manifest
                data += item.RollNumber + "\t";     //Item ID
                data += item.Weight + "\t";         //Weight
                data += "1\t";                      //Ct in Pkg
                data += "\t";                       //Location
                data += "\t";                       //LB/FT
                data += item.LinerFeet + "\n";      //LNFT
            }
            write(path + fileName, header + data);
        }

        private void write(string path, string content)
        {
            using (TextWriter writer = new StreamWriter(path))
            {
                writer.Write(content);
            }
        }

        #endregion
    }
}
