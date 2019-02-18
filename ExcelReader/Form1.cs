using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet result;
        DataSet result2;

        int ariv, le, ret, tstart, leave, leave2, tmid;
        double res, koren;

        private int indexdg3, indexdg4;

        List<WS> listaWS = new List<WS>();
        List<Student> listaStudent = new List<Student>();

        List<WS> upareno = new List<WS>();
        List<double> db = new List<double>();


        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);

                    reader.IsFirstRowAsColumnNames = true;
                    result = reader.AsDataSet();
                    cbxSheet.Items.Clear();
                    foreach(DataTable dt in result.Tables)
                    {
                        cbxSheet.Items.Add(dt.TableName);
                    }
                    var path = fs.Name;
                    var filename = Path.GetFileName(path);
                    label3.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
                    label3.Text = filename;

                    reader.Close();
                }
        }

        private void cbxSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < result.Tables[cbxSheet.SelectedIndex].Rows.Count; i++)
            {
                listaStudent = result.Tables[cbxSheet.SelectedIndex].AsEnumerable()
                .Select(row => new Student
                {
                    set_p = row.Field<string>("set"),
                    id_p = row.Field<string>("studentID_anon"),
                    ws_p = row.Field<string>("WS#"),
                    minGrade_p = row.Field<string>("midGrade"),
                    FinGrade_p = row.Field<string>("finalGrade"),
                    Tstart_p = row.Field<string>("Tstart-anonymized"),
                    Tmid_p = row.Field<string>("Tmid-anonymized"),
                    Tfinal_p = row.Field<string>("Tfinal-anonymized"),
                    TlastEdit_p = row.Field<string>("TlastEdit-anonymized"),
                }).ToList();
            }
            for (int j = 0; j < listaStudent.Count; j++)
            {
                if (listaStudent[j].id_p == "" || listaStudent[j].id_p == null)
                {
                    listaStudent.Remove(listaStudent[j]);
                    j--;
                }
            }
            if (label4.Text != "")
            {
                proveriID();
                ucitajLose();
            }
        }

        private void btnOpen2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd1 = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                if (ofd1.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs1 = File.Open(ofd1.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader2 = ExcelReaderFactory.CreateOpenXmlReader(fs1);

                    reader2.IsFirstRowAsColumnNames = true;
                    result2 = reader2.AsDataSet();
                    cbxSheet2.Items.Clear();
                    foreach (DataTable dt in result2.Tables)
                    {
                        cbxSheet2.Items.Add(dt.TableName);
                    }

                    var path = fs1.Name;
                    var filename = Path.GetFileName(path);
                    label4.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
                    label4.Text = filename;

                    reader2.Close();
                }
        }

        private void cbxSheet2_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < result2.Tables[cbxSheet2.SelectedIndex].Rows.Count; i++)
            {
                listaWS = result2.Tables[cbxSheet2.SelectedIndex].AsEnumerable()
                .Select(row => new WS
                {
                    WSid_p = row.Field<string>("WS#"),
                    WSorigin_p = row.Field<string>("WS#_origin"),
                    Tarival_p = row.Field<string>("Arrival-anonimized"),
                    Tleave_p = row.Field<string>("Leave-anonymized"),
                    Treturn_p = row.Field<string>("Return-anonymized"),
                    Tleave2_2 = row.Field<string>("Leave-anoymized"),
                }).ToList();
            }
            for (int j = 0; j < listaWS.Count; j++)
            {
                if(listaWS[j].WSid_p=="" || listaWS[j].WSid_p==null)
                {
                    listaWS.Remove(listaWS[j]);
                    j--;
                }
            }
            if (label3.Text != "")
            {
                proveriID();
                ucitajLose();
            }
        }

        private void ucitajLose()
        {
            listaStudent = listaStudent.OrderBy(x => x.Tstart_p).ToList();
            listaWS = listaWS.OrderBy(x => x.Tarival_p).ToList();

            var bindingList = new BindingList<Student>(listaStudent);
            var source = new BindingSource(bindingList, null);
            dataGridView3.DataSource = source;

            var bindingList2 = new BindingList<WS>(listaWS);
            var source2 = new BindingSource(bindingList2, null);
            dataGridView4.DataSource = source2;

            var bindingList1 = new BindingList<WS>(upareno);
            var source1 = new BindingSource(bindingList1, null);
            dataGridView.DataSource = source1;

            dataGridView.Columns["id_p"].Visible = false;
            dataGridView.Columns["minGrade_p"].Visible = false;
            dataGridView.Columns["FinGrade_p"].Visible = false;
            dataGridView.Columns["set_p"].Visible = false;
            dataGridView.Columns["ws_p"].Visible = false;
            dataGridView.Columns["Tstart_p"].Visible = false;
            dataGridView.Columns["Tmid_p"].Visible = false;
            dataGridView.Columns["Tfinal_p"].Visible = false;
            dataGridView.Columns["TlastEdit_p"].Visible = false;
            dataGridView.Columns["korisnik_p"].Visible = false;

            dataGridView4.Columns["id_p"].Visible = false;
            dataGridView4.Columns["minGrade_p"].Visible = false;
            dataGridView4.Columns["FinGrade_p"].Visible = false;
            dataGridView4.Columns["set_p"].Visible = false;
            dataGridView4.Columns["ws_p"].Visible = false;
            dataGridView4.Columns["Tstart_p"].Visible = false;
            dataGridView4.Columns["Tmid_p"].Visible = false;
            dataGridView4.Columns["Tfinal_p"].Visible = false;
            dataGridView4.Columns["TlastEdit_p"].Visible = false;
            dataGridView4.Columns["korisnik_p"].Visible = false;

            dataGridView4.Columns["studentID"].Visible = false;
            dataGridView4.Columns["midGrade"].Visible = false;
            dataGridView4.Columns["finalGrade"].Visible = false;
            dataGridView4.Columns["Tstart"].Visible = false;
            dataGridView4.Columns["Tmid"].Visible = false;
            dataGridView4.Columns["Tfinal"].Visible = false;
            dataGridView4.Columns["TlastEdit"].Visible = false;
            dataGridView4.Columns["korisnik_p"].Visible = false;

            label6.Text = "Upareno:  "+upareno.Count.ToString();
        }

        private void proveriID()
        {
            for (int i = 0; i < listaWS.Count; i++)
            {
                int index2 = 0;
                double min = Double.MaxValue;

                for (int j = 0; j < listaStudent.Count; j++)
                {
                    if (listaWS[i].WSid_p == listaStudent[j].ws_p || listaStudent[j].ws_p == listaWS[i].WSorigin_p)
                    {
                        if (listaWS[i].Tarival_p != "" && listaWS[i].Tleave_p != "" && listaStudent[j].Tstart_p != "" && listaStudent[j].TlastEdit_p != "")
                        {
                            res = Math.Pow((Double.Parse(listaWS[i].Tarival_p) - Double.Parse(listaStudent[j].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave_p) - Double.Parse(listaStudent[j].TlastEdit_p)), 2);
                            koren = Math.Sqrt(res);
                        }

                        if (listaWS[i].Treturn_p != "" && listaWS[i].Tleave2_2 != "" && listaStudent[j].Tstart_p != "" && listaStudent[j].TlastEdit_p != "" && listaWS[i].Tleave_p == "")
                        {
                            res = Math.Pow((Double.Parse(listaWS[i].Treturn_p) - Double.Parse(listaStudent[j].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave2_2) - Double.Parse(listaStudent[j].TlastEdit_p)), 2);
                            koren = Math.Sqrt(res);
                        }

                        if (listaWS[i].Tarival_p != "" && listaWS[i].Tleave2_2 != "" && listaStudent[j].Tstart_p != "" && listaStudent[j].TlastEdit_p != "" && listaWS[i].Tleave_p == "")
                        {
                            res = Math.Pow((Double.Parse(listaWS[i].Tarival_p) - Double.Parse(listaStudent[j].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave2_2) - Double.Parse(listaStudent[j].TlastEdit_p)), 2);
                            koren = Math.Sqrt(res);
                        }

                        if (koren < min)
                        {
                            min = koren;
                            index2 = j;
                        }
                    }
                }

                WS obj = new WS();
                obj.WSid_p = listaWS[i].WSid_p;
                obj.WSorigin_p = listaWS[i].WSorigin_p;
                obj.Tarival_p = listaWS[i].Tarival_p;
                obj.Tleave_p = listaWS[i].Tleave_p;
                obj.Treturn_p = listaWS[i].Treturn_p;
                obj.Tleave2_2 = listaWS[i].Tleave2_2;
                obj.korisnik_p = listaStudent[index2];

                upareno.Add(obj);
                listaWS.Remove(listaWS[i]);
                listaStudent.Remove(listaStudent[index2]);
            }
        }


        private void sortirajRucno()
        {
            int index = -1;
            if (cb1.Checked && cb2.Checked)
                MessageBox.Show("Morate selektovati samo jedan checkbox", "GRESKA!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            else if (cb1.Checked)
            {
                double min = Double.MaxValue;

                for (int i = 0; i < listaWS.Count; i++)
                {
                    if (listaWS[i].Tarival_p != "" && listaWS[i].Tleave_p != "" && listaStudent[indexdg3].Tstart_p != "" && listaStudent[indexdg3].TlastEdit_p != "")
                    {
                        res = Math.Pow((Double.Parse(listaWS[i].Tarival_p) - Double.Parse(listaStudent[indexdg3].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave_p) - Double.Parse(listaStudent[indexdg3].TlastEdit_p)), 2);
                        koren = Math.Sqrt(res);
                    }

                    if (listaWS[i].Treturn_p != "" && listaWS[i].Tleave2_2 != "" && listaStudent[indexdg3].Tstart_p != "" && listaStudent[indexdg3].TlastEdit_p != "" && listaWS[i].Tleave_p == "")
                    {
                        res = Math.Pow((Double.Parse(listaWS[i].Treturn_p) - Double.Parse(listaStudent[indexdg3].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave2_2) - Double.Parse(listaStudent[indexdg3].TlastEdit_p)), 2);
                        koren = Math.Sqrt(res);
                    }

                    if (listaWS[i].Tarival_p != "" && listaWS[i].Tleave2_2 != "" && listaStudent[indexdg3].Tstart_p != "" && listaStudent[indexdg3].TlastEdit_p != "" && listaWS[i].Tleave_p == "")
                    {
                        res = Math.Pow((Double.Parse(listaWS[i].Tarival_p) - Double.Parse(listaStudent[indexdg3].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave2_2) - Double.Parse(listaStudent[indexdg3].TlastEdit_p)), 2);
                        koren = Math.Sqrt(res);
                    }

                    if (koren < min)
                    {
                        min = koren;
                        index = i;
                    }
                }


                if (index != -1)
                {
                    WS obj = new WS();
                    obj.WSid_p = listaWS[index].WSid_p;
                    obj.WSorigin_p = listaWS[index].WSorigin_p;
                    obj.Tarival_p = listaWS[index].Tarival_p;
                    obj.Tleave_p = listaWS[index].Tleave_p;
                    obj.Treturn_p = listaWS[index].Treturn_p;
                    obj.Tleave2_2 = listaWS[index].Tleave2_2;
                    obj.korisnik_p = listaStudent[indexdg3];

                    upareno.Add(obj);
                    listaWS.Remove(listaWS[index]);
                    listaStudent.Remove(listaStudent[indexdg3]);
                }
            }

            else if (cb2.Checked)
            {
                double min = Double.MaxValue;
                for (int j = 0; j < listaStudent.Count; j++)
                {
                    for (int i = 0; i < listaWS.Count; i++)
                    {
                        if (listaWS[i].Tarival_p != "" && listaWS[i].Tleave_p != "" && listaStudent[j].Tstart_p != "" && listaStudent[j].TlastEdit_p != "")
                        {
                            res = Math.Pow((Double.Parse(listaWS[i].Tarival_p) - Double.Parse(listaStudent[j].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave_p) - Double.Parse(listaStudent[j].TlastEdit_p)), 2);
                            koren = Math.Sqrt(res);
                        }

                        if (listaWS[i].Treturn_p != "" && listaWS[i].Tleave2_2 != "" && listaStudent[j].Tstart_p != "" && listaStudent[j].TlastEdit_p != "" && listaWS[i].Tleave_p == "")
                        {
                            res = Math.Pow((Double.Parse(listaWS[i].Treturn_p) - Double.Parse(listaStudent[j].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave2_2) - Double.Parse(listaStudent[j].TlastEdit_p)), 2);
                            koren = Math.Sqrt(res);
                        }

                        if (listaWS[i].Tarival_p != "" && listaWS[i].Tleave2_2 != "" && listaStudent[j].Tstart_p != "" && listaStudent[j].TlastEdit_p != "" && listaWS[i].Tleave_p == "")
                        {
                            res = Math.Pow((Double.Parse(listaWS[i].Tarival_p) - Double.Parse(listaStudent[j].Tstart_p)), 2) + Math.Pow((Double.Parse(listaWS[i].Tleave2_2) - Double.Parse(listaStudent[j].TlastEdit_p)), 2);
                            koren = Math.Sqrt(res);
                        }

                        if (koren < min)
                        {
                            min = koren;
                            index = j;
                        }
                    }


                    if (index != -1)
                    {
                        WS obj = new WS();
                        obj.WSid_p = listaWS[index].WSid_p;
                        obj.WSorigin_p = listaWS[index].WSorigin_p;
                        obj.Tarival_p = listaWS[index].Tarival_p;
                        obj.Tleave_p = listaWS[index].Tleave_p;
                        obj.Treturn_p = listaWS[index].Treturn_p;
                        obj.Tleave2_2 = listaWS[index].Tleave2_2;
                        obj.korisnik_p = listaStudent[j];

                        upareno.Add(obj);
                        listaWS.Remove(listaWS[index]);
                        listaStudent.Remove(listaStudent[j]);
                    }
                }
            }
            ucitajLose();
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            indexdg3 = dataGridView3.CurrentCell.RowIndex;
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            indexdg4 = dataGridView4.CurrentCell.RowIndex;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //sortirajRucno();
        }

        private void btnUpari_Click(object sender, EventArgs e)
        {
            if (indexdg3 == 0 && indexdg4 == 0)
                MessageBox.Show("Morate selektovati vrste koje hocete da pridruzite!", "GRESKA!",MessageBoxButtons.OK,MessageBoxIcon.Error);
            else
            { 
                WS obj = new WS();
                obj.WSid_p = listaWS[indexdg4].WSid_p;
                obj.WSorigin_p = listaWS[indexdg4].WSorigin_p;
                obj.Tarival_p = listaWS[indexdg4].Tarival_p;
                obj.Tleave_p = listaWS[indexdg4].Tleave_p;
                obj.Treturn_p = listaWS[indexdg4].Treturn_p;
                obj.Tleave2_2 = listaWS[indexdg4].Tleave2_2;
                obj.korisnik_p = listaStudent[indexdg3];
                upareno.Add(obj);
                listaWS.Remove(listaWS[indexdg4]);
                listaStudent.Remove(listaStudent[indexdg3]);

                ucitajLose();
            }
        }
    }
}
