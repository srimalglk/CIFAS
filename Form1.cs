using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace CIFASFormatter3
{
    public partial class Form1 : Form
    {
        const int CUSTOMER_ID = 17;

        DataGridView my_dataviewer = new DataGridView();
        DataGridView exceptionViewer = new DataGridView();
        DataTable my_datatable = new DataTable();
        DataTable exception_datatable = new DataTable();
        ContextMenuStrip cMenu = new ContextMenuStrip();

        Button btn_process = new Button();
        Button btn_process_flatnos = new Button();
        Button btn_process_housenos = new Button();
        Button btn_process_street = new Button();
        Button btn_process_mobile_phoneNos = new Button();
        Button btn_process_home_phoneNos = new Button();
        Button btn_generate_output = new Button();

        String ExceptionFilePath;

        Dictionary<string, string> postcodeRegionMap = new Dictionary<string, string>();
        Dictionary<string, string> postcodeTownMap = new Dictionary<string, string>();

        public Form1()
        {
            InitializeComponent();
        }



        private void exceptionViewer_MouseClick(object sender, MouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
            {
                //var hti = exceptionViewer.HitTest(e.X, e.Y);
                //exceptionViewer.ClearSelection();
                //exceptionViewer.Rows[hti.RowIndex].Selected = true;

                ContextMenuStrip exception_menu = new System.Windows.Forms.ContextMenuStrip();
                int position_xy_mouse_row = exceptionViewer.HitTest(e.X, e.Y).RowIndex;

                if (position_xy_mouse_row >= 0)
                {
                    exceptionViewer.Rows[position_xy_mouse_row].Selected = true;
                    exception_menu.Items.Add("Accept").Name = "Accept";
                    exception_menu.Items.Add("Remove").Name = "Remove";
                    exception_menu.Items.Add("Add To Known Exception").Name = "Known-Exception";
                    exception_menu.Items.Add("Show Corrections").Name = "Show Corrections";

                }
                exception_menu.Show(exceptionViewer, new Point(e.X, e.Y));
                exception_menu.ItemClicked += new ToolStripItemClickedEventHandler(exception_menu_ItemClicked);
            }
        }

        void exception_menu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            //throw new NotImplementedException();

            switch( e.ClickedItem.Name.ToString() )
            {
                case "Accept":
                    int rowindex = exceptionViewer.CurrentRow.Index;
                    DataGridViewRow selectedRow = exceptionViewer.Rows[rowindex];
                    //MessageBox.Show(selectedRow.Cells[4].Value.ToString() + selectedRow.Cells[5].Value.ToString());
                    
                    if ( (selectedRow.Cells["Address Line1"].Value.ToString().Trim() + selectedRow.Cells["Address Line2"].Value.ToString().Trim()).Length > 0)
                        {
                            MessageBox.Show("Please process \"Address Line1\" & \"Address Line2\" data columns before fixing!");
                        }
                        else
                        {
                            MessageBox.Show("Preaparing to add fixed row into processed output.");
                            DataRow dr = my_datatable.NewRow(); //exception_datatable.NewRow();
                            for (int j = 0; j < exception_datatable.Columns.Count-1; j++)
                            {
                                dr[my_datatable.Columns[j].ColumnName] = exception_datatable.Rows[rowindex][j];
                            }
                            exception_datatable.Rows.RemoveAt(rowindex);
                            my_datatable.Rows.Add(dr);
                        }
                    //MessageBox.Show(selectedRow.Cells["Address Line2"].Value.ToString() + selectedRow.Cells[4].Value.ToString());
                    break;
                case "Remove":
                    MessageBox.Show("Removing rejected row.");
                    break;
                case "Known-Exception":
                    MessageBox.Show("Add to Known Exception.");
                    break;
                case "Show Corrections":
                    MessageBox.Show("Show Corrections.");
                    break;
            }

        }

        private void mItem_Fix_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Please FIX ME!");
        }


        private void btn_generate_output_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Generating Output Files ...");

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //saveFileDialog.Filter = "CSV Files|*.csv";
            saveFileDialog.Title = "Filename for output";
            saveFileDialog.FileName = "ICIC-CIFAS.csv";
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            //saveFileDialog.CheckFileExists = true;

            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //System.IO.StreamReader sr = new System.IO.StreamReader(openFileDialog1.FileName);
                //MessageBox.Show(saveFileDialog.FileName);
                //LoadDataFile(saveFileDialog.FileName);
                //sr.Close();
                //File.AppendText("Testing");

                StreamWriter sW = null;

                try
                {


                    File.WriteAllText(saveFileDialog.FileName, my_datatable.ToString());

                    sW = new StreamWriter(saveFileDialog.FileName);

                    sW.WriteLine("Surname,FirstName,BirthDate,HomeTelephone,MobileTelephone,Flat,HouseName,HouseNumber,Street,Town,Postcode");

                    for (int row = 0; row < my_datatable.Rows.Count; row++)
                    {
                        string lines = "";
                        //for (int col = 0; col < 4; col++)
                        //{
                        //1,0,2,4,3,5,7,6,8,11,13
                        lines += (string.IsNullOrEmpty(lines) ? "" : ",")
                            + my_datatable.Rows[row][1].ToString() + ","
                            + my_datatable.Rows[row][0].ToString() + ","
                            + my_datatable.Rows[row][13].ToString() + ","
                            + my_datatable.Rows[row][14].ToString() + ","
                            + my_datatable.Rows[row][15].ToString() + ","
                            + my_datatable.Rows[row][5].ToString() + ","
                            + my_datatable.Rows[row][6].ToString() + ","
                            + my_datatable.Rows[row][7].ToString() + ","
                            + "," //my_datatable.Rows[row][8].ToString() + ","
                            + "," //my_datatable.Rows[row][11].ToString() + ","
                            + my_datatable.Rows[row][11].ToString()
                            ;
                        //}

                        sW.WriteLine(lines);
                    }

                    MessageBox.Show("CIFAS File Successfully Generated.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("CIFAS File Genration ERROR! - " + ex.ToString());
                }
                finally
                {
                    sW.Close();
                }


                

                
            }

            btn_process.Enabled = false;
        }

        private void btn_process_flatnos_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {
                //MessageBox.Show(my_datatable.Rows[i]["Address Line1"].ToString());
                //Flat No Filteration

                Match match = Regex.Match(my_datatable.Rows[i]["Address Line1"].ToString(), @"(.*)((?i)Flat Number|Flat No|Flat No:|Flat No.|Flat #|Flat,|Flat.|Flat|Apartment Number|Apartment No|Apartment #|Apartment )[\s*|\.*](\d*[\w/\w]*)\s*(.*)");

                // Here we check the Match instance.
                if (match.Success)
                {
                    // Finally, we get the Group value and display it.
                    string key = match.Groups[1].Value;
                    //MessageBox.Show(key);
                    //MessageBox.Show(match.Groups[2].Value);
                    //MessageBox.Show(match.Groups[3].Value);
                    //MessageBox.Show(match.Groups[4].Value);

                    if (match.Groups[3].Value.Length > 2 && match.Groups[3].Value.All(char.IsLetter))
                    {

                        DataRow dr = exception_datatable.NewRow();
                        for (int j = 0; j < my_datatable.Columns.Count; j++)
                        {
                            dr[my_datatable.Columns[j].ColumnName] = my_datatable.Rows[i][j];
                        }
                        dr["Exception"] = "Invalid Flat Number";
                        my_datatable.Rows.RemoveAt(i);
                        exception_datatable.Rows.Add(dr);
                    }
                    else
                    {
                        my_datatable.Rows[i]["Flat No"] = match.Groups[3].Value;
                        my_datatable.Rows[i]["Address Line"] = match.Groups[1].Value.Trim() + " " + match.Groups[4].Value.Trim();
                    }
                }
                else
                {
                    my_datatable.Rows[i]["Address Line"] = my_datatable.Rows[i]["Address Line1"].ToString().Trim();
                }
            }
            setRowNumber(my_dataviewer);
            setRowNumber(exceptionViewer);

            btn_process_flatnos.Enabled = false;
        }

        private void btn_process_housenos_Click(object sender, EventArgs e)
        {
            //House No
            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {
                String Address="";
                if( my_datatable.Rows[i]["Address Line"].ToString().Trim().Length > 0)
                    Address = my_datatable.Rows[i]["Address Line"].ToString().Trim() + " " + my_datatable.Rows[i]["Address Line2"].ToString().Trim();
                else
                    Address = my_datatable.Rows[i]["Address Line2"].ToString().Trim();
                    //Address = my_datatable.Rows[i]["Address Line1"].ToString().Trim() + " " + my_datatable.Rows[i]["Address Line2"].ToString().Trim();

                Address = Address.Replace(",", " ").Trim();

                //Match match = Regex.Match(my_datatable.Rows[i]["Address Line2"].ToString(), @"(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(.*)");
                //Match match = Regex.Match(Address, @"(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(.*)");
                // Latest Used 30/10/2017 //Match match = Regex.Match(Address, @"(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(.*)");
                //Match match = Regex.Match(Address, @"(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(\w)\s*(.*)");
                //Match match = Regex.Match(Address, @"(.*)((?i)Number|No)(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(.*)");

                string pattern = @"(?<houseNo>\d+\w?([-/\s]\d+\w?)?)\s*(.*)";
                Regex regex = new Regex(pattern);
                Match match = regex.Match(Address);

                // Here we check the Match instance.
                if (match.Success)
                {
                    //my_datatable.Rows[i]["House No"] = match.Groups[2].Value.Trim();
                    my_datatable.Rows[i]["House No"] = match.Groups["houseNo"].Value.Trim();
                    //my_datatable.Rows[i]["House No"] = match.Groups[1].Value.Trim() + match.Groups[2].Value.Trim();

                    //my_datatable.Rows[i]["Address Line1"] = match.Groups[2].Value;
                    //House Name
                    //MessageBox.Show(match.Groups[1].Value);
                    //MessageBox.Show(match.Groups["houseNo"].Value);
                    //MessageBox.Show(match.Groups[2].Value);

                    /*
                    Match matchHouseName = Regex.Match(match.Groups[2].Value, @"(.*\s*(?i)(house|cottage|rectory|vicarage|bungalow|lodge|farm|home|office|heights|mansions|retreat|block))\s*(.*)");
                    if (matchHouseName.Success)
                    {
                        //MessageBox.Show(matchHouseName.Groups[1].Value);
                        //MessageBox.Show(matchHouseName.Groups[2].Value);
                        my_datatable.Rows[i]["House Name"] = matchHouseName.Groups[1].Value;
                        my_datatable.Rows[i]["Street"] = matchHouseName.Groups[3].Value + my_datatable.Rows[i]["Address Line2"];
                        my_datatable.Rows[i]["Address Line1"] = "";
                        my_datatable.Rows[i]["Address Line2"] = "";
                    }
                    else
                    {
                        my_datatable.Rows[i]["Street"] = match.Groups[2].Value + my_datatable.Rows[i]["Address Line2"];
                        my_datatable.Rows[i]["Address Line1"] = "";
                        my_datatable.Rows[i]["Address Line2"] = "";
                    }
                     * */
                    processHouseName(i, match.Groups[2].Value.Trim());
                    //processHouseName(i, match.Groups[3].Value.Trim());
                }
                else
                {
                    //processHouseName(i, my_datatable.Rows[i]["Address Line2"].ToString().Trim());
                    processHouseName(i, Address);

                    Match match2 = Regex.Match(my_datatable.Rows[i]["Street"].ToString(), @"(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(.*)");

                    // Here we check the Match instance.
                    if (match2.Success)
                    {
                        my_datatable.Rows[i]["House No"] = match2.Groups[1].Value.Trim();
                        //my_datatable.Rows[i]["Street"] = match.Groups[2].Value.Trim();
                        my_datatable.Rows[i]["Address Line"] = match.Groups[2].Value.Trim();
                    }
                }
            }
            setRowNumber(my_dataviewer);
            setRowNumber(exceptionViewer);

            btn_process_housenos.Enabled = false;
        }

        private void btn_process_street_Click(object sender, EventArgs e)
        {



            // Generate Postcode Mapper
            //Look for postcode

            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {

                //If Address Line 2 contains 1 word that is town
                //if contains more than two words then last one is town

                //09/10/2017
                /*
                if (!my_datatable.Rows[i]["Address Line2"].ToString().Trim().Contains(" "))
                {
                    my_datatable.Rows[i]["Town"] = my_datatable.Rows[i]["Address Line2"].ToString().Trim();
                }
                else
                    my_datatable.Rows[i]["Town"] = my_datatable.Rows[i]["Address Line2"].ToString().
                        Substring(my_datatable.Rows[i]["Address Line2"].ToString().LastIndexOf(' ') + 1);
                */

                if (!my_datatable.Rows[i]["Address Line"].ToString().Trim().Contains(" "))
                {
                    my_datatable.Rows[i]["Town"] = my_datatable.Rows[i]["Address Line"].ToString().Trim();
                    //MessageBox.Show("First One");
                }
                else
                {
                    my_datatable.Rows[i]["Town"] = my_datatable.Rows[i]["Address Line"].ToString().
                        Substring(my_datatable.Rows[i]["Address Line"].ToString().LastIndexOf(' ') + 1);
                    my_datatable.Rows[i]["Street"] = my_datatable.Rows[i]["Address Line"].ToString().
                        Substring(0,my_datatable.Rows[i]["Address Line"].ToString().LastIndexOf(' '));
                }


                //MessageBox.Show(my_datatable.Rows[i]["Postcode"].ToString().Trim());

                //Match match = Regex.Match(my_datatable.Rows[i]["Postcode"].ToString(), @"^([a-zA-Z]{1,2}[0-9]{1,2}[a-zA-Z]?)\s?.*");
                //commented on 05/10/2017 for testing TW59LL is not picked up on below expression
                //Match match = Regex.Match(my_datatable.Rows[i]["Postcode"].ToString().Trim(), @"^([A-Z]{1,2}[0-9R][0-9A-Z]?) [0-9][ABD-HJLNP-UW-Z]{2}$");
                Match match = Regex.Match(my_datatable.Rows[i]["Postcode"].ToString().Trim(), @"^([A-Z]{1,2}[0-9R][0-9A-Z]?)\s?[0-9][ABD-HJLNP-UW-Z]{2}$");

                // Here we check the Match instance.
                if (match.Success)
                {
                    //MessageBox.Show(match.Groups[1].Value);
                    //MessageBox.Show(postcodeRegionMap[match.Groups[1].Value].ToString());
                    //MessageBox.Show(postcodeTownMap[match.Groups[1].Value].ToString());

                    try
                    {
                        String address = my_datatable.Rows[i]["Street"].ToString().ToUpper();

                        MessageBox.Show(postcodeRegionMap[match.Groups[1].Value].ToString().ToUpper() + " - " + postcodeTownMap[match.Groups[1].Value].ToString().ToUpper());

                        if (!postcodeRegionMap[match.Groups[1].Value].ToString().Equals("") && address.EndsWith(postcodeRegionMap[match.Groups[1].Value].ToString().ToUpper()))
                        {
                            MessageBox.Show("One");
                            my_datatable.Rows[i]["State"] = "-" + match.Groups[1].Value + "-" + postcodeRegionMap[match.Groups[1].Value].ToString();

                            my_datatable.Rows[i]["Town"] = postcodeRegionMap[match.Groups[1].Value].ToString().ToUpper();
                            if (my_datatable.Rows[i]["Town"].ToString().Length > 0)
                                my_datatable.Rows[i]["Street"] = my_datatable.Rows[i]["Street"].ToString().Replace(my_datatable.Rows[i]["Town"].ToString(), "").Trim();

                        }
                        else if (!postcodeTownMap[match.Groups[1].Value].ToString().Equals("") && address.EndsWith(postcodeTownMap[match.Groups[1].Value].ToString().ToUpper()))
                        {
                            MessageBox.Show("Two");
                            my_datatable.Rows[i]["State"] = "-" + match.Groups[1].Value + "-" + postcodeTownMap[match.Groups[1].Value].ToString();

                            my_datatable.Rows[i]["Town"] = postcodeTownMap[match.Groups[1].Value].ToString().ToUpper();
                            if (my_datatable.Rows[i]["Town"].ToString().Length > 0)
                                my_datatable.Rows[i]["Street"] = my_datatable.Rows[i]["Street"].ToString().Replace(my_datatable.Rows[i]["Town"].ToString(), "").Trim();
                        }

                    }
                    catch (KeyNotFoundException ex)
                    {
                        //MessageBox.Show(match.Groups[1].Value + ex);
                        //exception_datatable.Rows.Add(my_datatable.Rows[i].C);

                        DataRow dr = exception_datatable.NewRow();
                        for (int j = 0; j < my_datatable.Columns.Count; j++)
                        {
                            dr[my_datatable.Columns[j].ColumnName] = my_datatable.Rows[i][j];
                        }
                        dr["Exception"] = "Invalid Postcode - " + ex.ToString();
                        my_datatable.Rows.RemoveAt(i);
                        exception_datatable.Rows.Add(dr);

                    }
                }
            }
            setRowNumber(my_dataviewer);
            setRowNumber(exceptionViewer);

            btn_process_street.Enabled = false;
        }

        private void btn_process_mobile_phoneNos_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Processing Phone Numbers");

            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {
                //MessageBox.Show(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim());
                //Match match = Regex.Match(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim(), @"^(\(?(?:0(?:0|11)\)?[\s-]?\(?|\+)(44)\)?[\s-]?)?\(?0?(?:\)[\s-]?)?([1-9]\d{1,4}\)?[\d[\s-]]+)((?:x|ext\.?|\#)\d{3,4})?$");

                //Match match = Regex.Match("+447947517036", @"^((?(?:0(?:0|11))?[\s-]?(?|+)(44))?[\s-]?)?(?0?(?:)[\s-]?)?([1-9]\d{1,4})?[\d[\s-]]+)((?:x|ext\.?|\#)\d{3,4})?$");
                //Match match = Regex.Match(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim(), @"^([\s-]?\(?|\+)(44)\)?[\s-]?)?\(?0?(?:\)[\s-]?)?([1-9]\d{1,4}\)?[\d[\s-]]+)((?:x|ext\.?|\#)\d{3,4})?$");
                //Match match = Regex.Match("+44 7947517036", @"^((\(\(?|\+)(44)\)?[\s-]?)?\(?0?(?:\)[\s-]?)?([1-9]\d{1,4}\)?[\d[\s-]]+)((?:x|ext\.?|\#)\d{3,4})?$");

                //Match match = Regex.Match(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim(), @"^([00]?[\+]?44?)([\s-.]?\d{3}[\s-.]?\d{4})$");
                //Match match = Regex.Match(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim(), @"\D*([2-9]\d{2})(\D*)([2-9]\d{2})(\D*)(\d{4})\D*");
                //Match match = Regex.Match(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim(), @"\D*([\0\0|\+]?[4]\d{1})(\D*)([2-9]\d{1})(\D*)(\d{8})\D*");

                if (my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim().Length > 0)
                {

                    Match match = Regex.Match(my_datatable.Rows[i]["Mobile Telephone"].ToString().Trim(), @"^.*(?:^0|^00|\+0\(\)0|\+00|\+|4{2})[-\s]?(.*)$");

                    // Here we check the Match instance.
                    if (match.Success)
                    {
                        //MessageBox.Show(match.Groups[3].Value.Trim());
                        //my_datatable.Rows[i]["Mobile Telephone"] = "0" + match.Groups[3].Value.Trim() + match.Groups[5].Value.Trim();

                        //if match.Groups[1].Value.Trim() Not empty then format as follows
                        //MessageBox.Show(match.Groups[1].Value.Trim());
                        if (match.Groups[1].Value.Trim().Length > 0)
                            my_datatable.Rows[i]["Mobile Telephone"] = "0" + Regex.Replace(match.Groups[1].Value.Trim(), @"\s", "");
                        else
                            my_datatable.Rows[i]["Mobile Telephone"] = "";
                    }
                    else
                    {
                        DataRow dr = exception_datatable.NewRow();
                        for (int j = 0; j < my_datatable.Columns.Count; j++)
                        {
                            dr[my_datatable.Columns[j].ColumnName] = my_datatable.Rows[i][j];
                        }
                        dr["Exception"] = "Invalid Mobile Phone No";
                        my_datatable.Rows.RemoveAt(i);
                        exception_datatable.Rows.Add(dr);
                    }
                }

            }

            btn_process_mobile_phoneNos.Enabled = false;
        }

        private void btn_process_home_phoneNos_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {
                //Match match = Regex.Match(my_datatable.Rows[i]["Home Telephone"].ToString().Trim(), @"\D*([\0\0|\+]?[4]\d{1})(.*)");

                if (my_datatable.Rows[i]["Home Telephone"].ToString().Trim().Length > 0)
                {

                    //Match match = Regex.Match(my_datatable.Rows[i]["Home Telephone"].ToString().Trim(), @"^.*(?:^0|^00|\+|4{2})[-\s]?(.*)$");
                    Match match = Regex.Match(my_datatable.Rows[i]["Home Telephone"].ToString().Trim(), @"^.*(?:^0|^00|\+0\(\)0|\+|4{2})[-\s]?(.*)$");

                    // Here we check the Match instance.
                    if (match.Success)
                    {
                        //MessageBox.Show(match.Groups[3].Value.Trim());
                        //str = Regex.Replace(str, @"\s", "");
                        my_datatable.Rows[i]["Home Telephone"] = "0" + Regex.Replace(match.Groups[1].Value.Trim(), @"\s", "");
                        //if formated number is <11 then exception
                    }
                    else
                    {
                        DataRow dr = exception_datatable.NewRow();
                        for (int j = 0; j < my_datatable.Columns.Count; j++)
                        {
                            dr[my_datatable.Columns[j].ColumnName] = my_datatable.Rows[i][j];
                        }
                        dr["Exception"] = "Invalid Home Phone No";
                        my_datatable.Rows.RemoveAt(i);
                        exception_datatable.Rows.Add(dr);
                    }
                }
            }
            btn_process_home_phoneNos.Enabled = false;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            /*
            //my_datatable = my_datatable.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is System.DBNull || string.Compare((field as string).Trim(), string.Empty) == 0)).CopyToDataTable();
            
            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {
                //MessageBox.Show(my_datatable.Rows[i]["Address Line1"].ToString());
                //Flat No Filteration

                Match match = Regex.Match(my_datatable.Rows[i]["Address Line1"].ToString(), @"((?i)Flat Number|Flat No|Flat #|Flat|Apartment Number|Apartment No|Apartment #|Apartment)\s*(\d*[\w]*)\s*(.*)");

                // Here we check the Match instance.
                if (match.Success)
                {
                    // Finally, we get the Group value and display it.
                    //string key = match.Groups[1].Value;
                    //MessageBox.Show(key);
                    //MessageBox.Show(match.Groups[2].Value);
                    //MessageBox.Show(match.Groups[3].Value);
                    //MessageBox.Show(match.Groups[4].Value);
                    my_datatable.Rows[i]["Flat No"] = match.Groups[2].Value;
                    my_datatable.Rows[i]["Address Line1"] = match.Groups[3].Value;
                    
                }
                


            }

            //my_dataviewer.DataSource = null;
            //my_dataviewer.Update();
            //my_dataviewer.Refresh();
            //exceptionViewer.DataSource = exception_datatable;


            //House No
            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {

                Match match = Regex.Match(my_datatable.Rows[i]["Address Line1"].ToString(), @"(^\d*\w*\s*(?:[-\\\x2F]?\s*)?\d*\s*\d+\\?\x2F?\s*\d*)\s*(.*)");

                // Here we check the Match instance.
                if (match.Success)
                {
                    my_datatable.Rows[i]["House No"] = match.Groups[1].Value;
                    //my_datatable.Rows[i]["Address Line1"] = match.Groups[2].Value;
                    //House Name
                    //MessageBox.Show(match.Groups[2].Value);
            */
                    /*
                    Match matchHouseName = Regex.Match(match.Groups[2].Value, @"(.*\s*(?i)(house|cottage|rectory|vicarage|bungalow|lodge|farm|home|office|heights|mansions|retreat|block))\s*(.*)");
                    if (matchHouseName.Success)
                    {
                        //MessageBox.Show(matchHouseName.Groups[1].Value);
                        //MessageBox.Show(matchHouseName.Groups[2].Value);
                        my_datatable.Rows[i]["House Name"] = matchHouseName.Groups[1].Value;
                        my_datatable.Rows[i]["Street"] = matchHouseName.Groups[3].Value + my_datatable.Rows[i]["Address Line2"];
                        my_datatable.Rows[i]["Address Line1"] = "";
                        my_datatable.Rows[i]["Address Line2"] = "";
                    }
                    else
                    {
                        my_datatable.Rows[i]["Street"] = match.Groups[2].Value + my_datatable.Rows[i]["Address Line2"];
                        my_datatable.Rows[i]["Address Line1"] = "";
                        my_datatable.Rows[i]["Address Line2"] = "";
                    }
                     * */

            /*
                    processHouseName(i, match.Groups[2].Value);
                }
                else
                {
                    processHouseName(i, my_datatable.Rows[i]["Address Line1"].ToString());
                }
            }
            

            // Generate Postcode Mapper
            //Look for postcode

            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {

                //Match match = Regex.Match(my_datatable.Rows[i]["Postcode"].ToString(), @"^([a-zA-Z]{1,2}[0-9]{1,2}[a-zA-Z]?)\s?.*");
                Match match = Regex.Match(my_datatable.Rows[i]["Postcode"].ToString().Trim(), @"^([A-Z]{1,2}[0-9R][0-9A-Z]?) [0-9][ABD-HJLNP-UW-Z]{2}$");

                // Here we check the Match instance.
                if (match.Success)
                {
                    //MessageBox.Show(match.Groups[1].Value);
                    //MessageBox.Show(postcodeRegionMap[match.Groups[1].Value].ToString());
                    //MessageBox.Show(postcodeTownMap[match.Groups[1].Value].ToString());
                    
                    try
                    {
                        String address = my_datatable.Rows[i]["Street"].ToString().ToUpper();

                        if (!postcodeRegionMap[match.Groups[1].Value].ToString().Equals("") &&  address.EndsWith(postcodeRegionMap[match.Groups[1].Value].ToString().ToUpper()))
                        {
                            my_datatable.Rows[i]["State"]="-" + match.Groups[1].Value + "-" + postcodeRegionMap[match.Groups[1].Value].ToString();

                            my_datatable.Rows[i]["Town"] = postcodeRegionMap[match.Groups[1].Value].ToString().ToUpper();
                            if (my_datatable.Rows[i]["Town"].ToString().Length > 0)
                                my_datatable.Rows[i]["Street"] = my_datatable.Rows[i]["Street"].ToString().Replace(my_datatable.Rows[i]["Town"].ToString(), "").Trim();

                        }
                        else if (!postcodeTownMap[match.Groups[1].Value].ToString().Equals("") && address.EndsWith(postcodeTownMap[match.Groups[1].Value].ToString().ToUpper()))
                        {
                            my_datatable.Rows[i]["State"] = "-" + match.Groups[1].Value + "-" + postcodeTownMap[match.Groups[1].Value].ToString();

                            my_datatable.Rows[i]["Town"] = postcodeTownMap[match.Groups[1].Value].ToString().ToUpper();
                            if (my_datatable.Rows[i]["Town"].ToString().Length > 0)
                                my_datatable.Rows[i]["Street"] = my_datatable.Rows[i]["Street"].ToString().Replace(my_datatable.Rows[i]["Town"].ToString(), "").Trim();
                        }

                    }
                    catch (KeyNotFoundException ex)
                    {
                        //MessageBox.Show(match.Groups[1].Value + ex);
                        //exception_datatable.Rows.Add(my_datatable.Rows[i].C);
                        
                        DataRow dr = exception_datatable.NewRow();
                        for (int j = 0; j < my_datatable.Columns.Count; j++)
                        {
                            dr[my_datatable.Columns[j].ColumnName] = my_datatable.Rows[i][j];
                        }
                        my_datatable.Rows.RemoveAt(i);
                        exception_datatable.Rows.Add(dr);
                        
                    }
                }
            }

            setRowNumber(my_dataviewer);
            setRowNumber(exceptionViewer);

            */

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Text Files|*.txt|CSV Files|*.csv";
            openFileDialog1.Title = "Select a CSV File";

            //MessageBox.Show("On File Opening");

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //System.IO.StreamReader sr = new System.IO.StreamReader(openFileDialog1.FileName);
                //MessageBox.Show(openFileDialog1.FileName);
                LoadDataFile(openFileDialog1.FileName);
                //sr.Close();
                btn_process.Enabled = false;
            }
            else
            {
                btn_process.Enabled = true;
            }
        }

        private void processHouseName(int row, string address)
        {
            //if "Block comes first exception"

            Match matchHouseName = Regex.Match(address, @"(.*\s*(?i)(house|cottage|rectory|vicarage|bungalow|lodge|farm|home|office|heights|mansions|retreat|block))\s*(.*)");
            if (matchHouseName.Success)
            {
                //MessageBox.Show(matchHouseName.Groups[1].Value);
                //MessageBox.Show(matchHouseName.Groups[2].Value);

                //if (matchHouseName.Groups[1].Value.Trim().ToUpper().Equals("BLOCK"))
                if (matchHouseName.Groups[1].Value.Trim().Split(' ').Length == 1)
                {
                    //Exception
                    DataRow dr = exception_datatable.NewRow();
                    for (int j = 0; j < my_datatable.Columns.Count; j++)
                    {
                        dr[my_datatable.Columns[j].ColumnName] = my_datatable.Rows[row][j];
                    }
                    dr["Exception"] = "Invalid House Name";
                    my_datatable.Rows.RemoveAt(row);
                    exception_datatable.Rows.Add(dr);
                }
                else
                {

                    my_datatable.Rows[row]["House Name"] = matchHouseName.Groups[1].Value.Trim();
                    //09/10/2017
                    //my_datatable.Rows[row]["Street"] = (matchHouseName.Groups[3].Value + " " + my_datatable.Rows[row]["Address Line2"].ToString()).Trim();
                    my_datatable.Rows[row]["Address Line"] = (matchHouseName.Groups[3].Value).Trim();

                    //my_datatable.Rows[row]["Address Line1"] = "";
                    //my_datatable.Rows[row]["Address Line2"] = "";
                    //MessageBox.Show( "1 : " + (matchHouseName.Groups[3].Value).Trim());
                }
            }
            else
            {
                //my_datatable.Rows[row]["Street"] = ( address + " " + my_datatable.Rows[row]["Address Line2"].ToString()).Trim();
                my_datatable.Rows[row]["Address Line"] = (address).Trim();
                //my_datatable.Rows[row]["Address Line1"] = "";
                //my_datatable.Rows[row]["Address Line2"] = "";
                //MessageBox.Show("2 : " + (address).Trim());
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //DataGridView my_dataviewer = new DataGridView();
            //DataGridView exceptionViewer = new DataGridView();
            
            btn_process.Click += new EventHandler(btn_Click);

            
            btn_process_flatnos.Click += new EventHandler(btn_process_flatnos_Click);


            btn_process_housenos.Click += new EventHandler(btn_process_housenos_Click);

            
            btn_process_street.Click += new EventHandler(btn_process_street_Click);

            btn_process_mobile_phoneNos.Click += new EventHandler(btn_process_mobile_phoneNos_Click);

            btn_process_home_phoneNos.Click += new EventHandler(btn_process_home_phoneNos_Click);

            btn_generate_output.Click += new EventHandler(btn_generate_output_Click);

            my_dataviewer.Sorted += new EventHandler(my_dataviewer_sorted);

            
            this.Size = new Size(1200, 700);
            my_dataviewer.Size = new Size(1000, 400);
            my_dataviewer.Location = new Point(5, 5);

            exceptionViewer.Size = new Size(1000, 200);
            exceptionViewer.Location = new Point(5, 425);

            btn_process.Text = "Load CSV File";
            btn_process.Size = new Size(150, 30);
            btn_process.Location = new Point( 1000 + 30, 5);

            btn_process_flatnos.Text = "Format Flat Nos";
            btn_process_flatnos.Size = new Size(150, 30);
            btn_process_flatnos.Location = new Point( 1000 + 30, 40 +30);

            btn_process_housenos.Text = "Format House No && Name";
            btn_process_housenos.Size = new Size(150, 30);
            btn_process_housenos.Location = new Point( 1000 + 30, 75 +30);

            btn_process_street.Text = "Format Street && Town";
            btn_process_street.Size = new Size(150, 30);
            btn_process_street.Location = new Point(1000 + 30, 110 + 30);

            btn_process_mobile_phoneNos.Text = "Format Mobile Phone Numbers";
            btn_process_mobile_phoneNos.Size = new Size(150, 30);
            btn_process_mobile_phoneNos.Location = new Point(1000 + 30, 145 + 30);

            btn_process_home_phoneNos.Text = "Format Home Phone Numbers";
            btn_process_home_phoneNos.Size = new Size(150, 30);
            btn_process_home_phoneNos.Location = new Point(1000 + 30, 180 + 30);


            btn_generate_output.Text = "Generate CIFAS Files";
            btn_generate_output.Size = new Size(150, 30);
            btn_generate_output.Location = new Point(1000 + 30, 210 + 30 + 30);


            ToolStripItem mItem = cMenu.Items.Add("Fix");
            mItem.Click += new EventHandler(mItem_Fix_Click);

            DataGridViewRow defaultRow = new DataGridViewRow();
            defaultRow.ContextMenuStrip = cMenu;

            this.Cursor = Cursors.WaitCursor;
            GeneratePostcodeMap();
            //LoadDataFile();

            this.Cursor = Cursors.Default;

            
            //my_dataviewer.DataSource = my_datatable;
            //exceptionViewer.DataSource = exception_datatable;
            
            //my_dataviewer.Refresh();
            this.Controls.Add(my_dataviewer);
            this.Controls.Add(btn_process);
            this.Controls.Add(btn_process_flatnos);
            this.Controls.Add(btn_process_housenos);
            this.Controls.Add(btn_process_street);
            this.Controls.Add(btn_process_mobile_phoneNos);
            this.Controls.Add(btn_process_home_phoneNos);
            this.Controls.Add(btn_generate_output);
            this.Controls.Add(exceptionViewer);

            setRowNumber(my_dataviewer);
            setRowNumber(exceptionViewer);

            this.exceptionViewer.MouseClick += new System.Windows.Forms.MouseEventHandler(this.exceptionViewer_MouseClick);
            this.my_dataviewer.MouseClick += new MouseEventHandler(my_dataviewer_MouseClick);

            string exeRuntimeDirectory =
                System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string subDirectory = System.IO.Path.Combine(exeRuntimeDirectory, "KnownExceptions");
            if (!System.IO.Directory.Exists(subDirectory))
            {
                // Output directory does not exist, so create it.
                System.IO.Directory.CreateDirectory(subDirectory);
            }

            //string driveLetter = Path.GetPathRoot(Environment.CurrentDirectory);
            //string path
            ExceptionFilePath = subDirectory + @"\\Exception.txt";
            //StreamWriter sw = new StreamWriter(path);
            //string lines = "First line.\r\nSecond line.\r\nThird line.";

            if (!System.IO.File.Exists(ExceptionFilePath))
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(ExceptionFilePath);
                //file.WriteLine(lines);/
                file.Close();
            }

        }

        void my_dataviewer_MouseClick(object sender, MouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
            {
                //var hti = exceptionViewer.HitTest(e.X, e.Y);
                //exceptionViewer.ClearSelection();
                //exceptionViewer.Rows[hti.RowIndex].Selected = true;

                ContextMenuStrip dataviewer_menu = new System.Windows.Forms.ContextMenuStrip();
                //int position_xy_mouse_row = exceptionViewer.HitTest(e.X, e.Y).RowIndex;
                int position_xy_mouse_row = my_dataviewer.HitTest(e.X, e.Y).RowIndex;

                if (position_xy_mouse_row >= 0)
                {
                    my_dataviewer.Rows[position_xy_mouse_row].Selected = true;
                    dataviewer_menu.Items.Add("Add To Known Exception").Name = "Known-Exception";
                    dataviewer_menu.Items.Add("Remove From Known Exception").Name = "Remove-Exception";
                }
                dataviewer_menu.Show(my_dataviewer, new Point(e.X, e.Y));
                dataviewer_menu.ItemClicked += new ToolStripItemClickedEventHandler(my_dataviewer_menu_ItemClicked);
                }
        }

        void my_dataviewer_menu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            //throw new NotImplementedException();

            switch( e.ClickedItem.Name.ToString() )
            {
                
                case "Known-Exception":
                    MessageBox.Show("This is a known exception");
                    using (StreamWriter sw = File.AppendText(ExceptionFilePath))
                    {
                        //DataGridViewRow selectedRow = my_dataviewer.;
                        //sw.WriteLine(sender.ToString());
                        //sw.WriteLine(my_dataviewer.SelectedRows.ToString());
                        //int index = e.RowIndex;
                        //DataGridViewRow selectedRow = dataGridView1.Rows[index];
                        //textBoxID.Text = selectedRow.Cells[0].Value.ToString();
                        //textBoxFN.Text = selectedRow.Cells[1].Value.ToString();
                        //textBoxLN.Text = selectedRow.Cells[2].Value.ToString();
                        //textBoxAGE.Text = selectedRow.Cells[3].Value.ToString();

                        int rowIndex = my_dataviewer.CurrentRow.Index;
                        //sw.WriteLine( rowIndex.ToString());
                        DataGridViewRow row = my_dataviewer.Rows[rowIndex];
                        sw.WriteLine( row.Cells[CUSTOMER_ID].Value );
                        this.my_dataviewer.Rows[rowIndex].DefaultCellStyle.BackColor = Color.Salmon;
                        this.my_dataviewer.Rows[rowIndex].Selected = false;
                    }	
                    break;
                case "Remove-Exception":

                    MessageBox.Show("This is to Remove Known exception");

                    int rowIndex1 = my_dataviewer.CurrentRow.Index;
                        //sw.WriteLine( rowIndex.ToString());
                    DataGridViewRow row1 = my_dataviewer.Rows[rowIndex1];
                    //sw.WriteLine( row1.Cells[CUSTOMER_ID].Value );
                    this.my_dataviewer.Rows[rowIndex1].DefaultCellStyle.BackColor = Color.White;
                    this.my_dataviewer.Rows[rowIndex1].Selected = false;

                    break;
                //case "Remove":
                //    MessageBox.Show("Removing rejected row.");
                //    break;
            }

        }

        private void GeneratePostcodeMap()
        {
            string[] raw_data = System.IO.File.ReadAllLines("C:\\temp\\ICICI-04-08-2016\\uk_postcode_05_edited.csv");
            //string[] raw_data = System.IO.File.ReadAllLines("C:\\temp\\AddressTest.csv");
            string[] data_col = null;

            int x = 0;

            foreach (string text_line in raw_data)
            {
                //empty text_line need to be tested and filtered

                data_col = text_line.Split(',');

                if (x == 0)
                {
                    x++;
                }
                else
                {
                    postcodeRegionMap.Add(data_col[0].ToString(), data_col[5].ToString());
                    postcodeTownMap.Add(data_col[0].ToString(), data_col[6].ToString());
                }
            }
        }

        /*
        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1); // row.Index + 1;
            }
        }
         */

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }

            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void LoadDataFile(String filename)
        {
            //MessageBox.Show("TEst");

            //string[] raw_data = System.IO.File.ReadAllLines("C:\\temp\\AddressTest.csv");
            //string[] raw_data = System.IO.File.ReadAllLines("C:\\temp\\ICICI-04-08-2016\\misibg513-edisted.csv");
            //string[] raw_data = System.IO.File.ReadAllLines("C:\\temp\\ICICI-04-08-2016\\misibg513_edisted_new.csv");

            string[] raw_data = System.IO.File.ReadAllLines(filename);

            string[] data_col = null;

            int x = 0;
            int row =0;

            exception_datatable.Columns.Add("Exception");

            foreach (string text_line in raw_data)
            {
                //empty text_line need to be tested and filtered

                //data_col = text_line.Split(',');
                data_col = text_line.Split('|');

                if (x == 0)
                {
                    for (int i = 0; i <= data_col.Count() - 1; i++)
                    {
                        my_datatable.Columns.Add(data_col[i]);
                        exception_datatable.Columns.Add(data_col[i]);
                    }
                    x++;

                    //Additional data columns added on 14/09/2017
                    //Flat No
                    my_datatable.Columns.Add("Flat No");
                    exception_datatable.Columns.Add("Flat No");
                    //House Name
                    my_datatable.Columns.Add("House Name");
                    exception_datatable.Columns.Add("House Name");
                    //House No
                    my_datatable.Columns.Add("House No");
                    exception_datatable.Columns.Add("House No");
                    //Street
                    my_datatable.Columns.Add("Street");
                    exception_datatable.Columns.Add("Street");
                    //Address Line
                    my_datatable.Columns.Add("Address Line");
                    exception_datatable.Columns.Add("Address Line");

                }
                else
                {
                    try
                    {
                        //MessageBox.Show("" +data_col.Count());
                        //if (data_col[14].ToString().Length > 0 && data_col[14].ToString().ToUpper().Equals("UNITED KINGDOM"))
                        if (data_col[12].ToString().Length > 0 && data_col[12].ToString().ToUpper().Equals("UNITED KINGDOM"))
                        {
                            my_datatable.Rows.Add(data_col);
                        }
                        else
                        {
                            /*
                            Match match = Regex.Match(data_col[4].ToString().Trim(), @"^.*(?:^0|^00|\+0\(\)0|\+00|\+|4{2})[-\s]?(.*)$");
                            //Match match = Regex.Match(data_col[4].ToString().Trim(), @"^.*(?:+0044)[-\s]?(.*)$");

                            if (match.Success)
                            {
                                MessageBox.Show(data_col[4].ToString());
                            }
                             * */

                            if (
                                data_col[4].ToString().Trim().StartsWith("+0044") || 
                                data_col[4].ToString().Trim().StartsWith("+00044")
                               )
                            {
                                //MessageBox.Show(data_col[4].ToString());
                                //data_col[18] = "Invalid Address";
                                //MessageBox.Show(" " + data_col.Length);
                                exception_datatable.Rows.Add(data_col);
                                exception_datatable.Rows[row++]["Exception"] = "Invalid Address";
                                //["Exception"] = "Invalid Mobile Phone No";
                            }


                        }
                    }catch(Exception e){
                        MessageBox.Show("Unable to Find Country Column within data file : \n\n\n" + e.ToString());
                        break;
                    }
                }
                
            }
            

            ReorderTable( ref my_datatable,
                "First Name",
                "Surname",
                "Address Line1",
                "Address Line2",
                "Address Line",
                "Flat No",
                "House Name",
                "House No",
                "Street",
                "Town",
                "State",
                "Postcode",
                "Country",
                "Birth Date",
                "Home Telephone",
                "Mobile Telephone",
                "Email",
                "Customer ID",
                "END");


            ReorderTable(ref exception_datatable,
                 "First Name",
                 "Surname",
                 "Address Line1",
                 "Address Line2",
                 "Address Line",
                 "Flat No",
                 "House Name",
                 "House No",
                 "Street",
                 "Town",
                 "State",
                 "Postcode",
                 "Country",
                 "Birth Date",
                 "Home Telephone",
                 "Mobile Telephone",
                 "Email",
                 "Customer ID",
                 "END",
                 "Exception");



            //dataGridView.DataSource = null;
            //dataGridView.Update();
            //dataGridView.Refresh();

            my_dataviewer.DataSource = null;
            my_dataviewer.Update();
            my_dataviewer.Refresh();
            my_dataviewer.DataSource = my_datatable;
            my_dataviewer.Columns[17].Visible = true;
            my_dataviewer.Columns[4].Visible = true; //false;
            my_dataviewer.Columns[18].Visible = true; //false;

            exceptionViewer.DataSource = null;
            exceptionViewer.Update();
            exceptionViewer.Refresh();
            exceptionViewer.DataSource = exception_datatable;
            exceptionViewer.Columns[17].Visible = true;
            exceptionViewer.Columns[4].Visible = false;
            exceptionViewer.Columns[18].Visible = false;


            SelectRowsWithKnownExceptions();

        }

        private void SelectRowsWithKnownExceptions()
        {
            //Select rows with Known Exceptions
            //DataGridViewCellStyle cellStyle = new DataGridViewCellStyle();

            //read file and get all the Customer IDs
            string[] readText = File.ReadAllLines(ExceptionFilePath);

            for (int i = 0; i < my_datatable.Rows.Count; i++)
            {
                DataGridViewRow row = my_dataviewer.Rows[i];
                if (readText.Contains(row.Cells[CUSTOMER_ID].Value))
                {
                    this.my_dataviewer.Rows[i].DefaultCellStyle.BackColor = Color.Salmon;
                }
            }
        }

        public static void ReorderTable(ref DataTable table, params String[] columns)
        {
            if (columns.Length != table.Columns.Count)
                throw new ArgumentException("Count of columns must be equal to table.Column.Count", "columns");

            for (int i = 0; i < columns.Length; i++)
            {
                table.Columns[columns[i]].SetOrdinal(i);
            }
            
        }

        private void my_dataviewer_sorted(object sender, EventArgs e)
        {
            SelectRowsWithKnownExceptions();
            //this.my_dataviewer.FirstDisplayedCell = this.my_dataviewer.CurrentCell;
            //MessageBox.Show("DataGrid Sorted Completed");
        }
    }
}
