using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Priroda
{
    public partial class main : Form
    {
        public int lv;

        OleDbCommand cm;
        OleDbConnection cn;
        OleDbDataAdapter da;
        OleDbCommandBuilder cb;
        DataTable dt;
        
        public main()
        {
            InitializeComponent();
            if (lv == 3)
            {
                tabControl1.SelectedTab = tabPage3;
            }
            else if (lv == 2)
            {
                tabControl1.SelectedTab = tabPage2;
            }
            else if (lv == 1)
            {
                tabControl1.SelectedTab = tabPage1;
            }
            ZapolneniyData2();
            cn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() + @"\User.accdb");
            cn.Open();
            cm = new OleDbCommand("Select * from KormDB", cn);
            dt = new DataTable();
            da = new OleDbDataAdapter(cm);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView2.DataSource = dt;
            dataGridView3.DataSource = dt;
            int i = 0;
            while (i < 5)
            {
                dataGridView1.Columns[i].ReadOnly = true;
                i++;
            }
            cn.Close();
        }

        public void Cheklv(int id) //проверка на уровень полномочий
        {
            if (lv < id)
            {
                MessageBox.Show("Не достаточно полномочий");
                tabControl1.SelectedTab = tabPage1;
            }
        }

        private void button1_Click(object sender, EventArgs e) //Выход
        {
            Application.Exit();
        }

        private void tabPage2_Enter(object sender, EventArgs e) //Проверка
        {
            Cheklv(2);
        }

        private void tabPage3_Enter(object sender, EventArgs e) //Проверка
        {
            Cheklv(3);
        }

        private void button2_Click(object sender, EventArgs e) //Смена пользователя
        {
            Relogin z = new Relogin();
            z.ShowDialog();
            if (z.tmp == 3)
            {
                lv = 3;
                tabControl1.SelectedTab = tabPage3;
            }
            else if (z.tmp == 2)
            {
                lv = 2;
                tabControl1.SelectedTab = tabPage2;
            }
            else if (z.tmp == 1)
            {
                lv = 1;
                tabControl1.SelectedTab = tabPage1;
            }
        }

        private void button3_Click(object sender, EventArgs e) //Сохранение
        {
            cb = new OleDbCommandBuilder(da);
            da.Update(dt);
        }

        private void button4_Click(object sender, EventArgs e) //Сохранение
        {
            cb = new OleDbCommandBuilder(da);
            da.Update(dt);
        }

        private void button5_Click(object sender, EventArgs e) //
        {
            foreach (DataGridViewRow row in dataGridView4.Rows)
            {
                object[] rowData = new object[row.Cells.Count];
                for (int i = 0; i < rowData.Length; ++i)
                {
                    rowData[i] = row.Cells[i].Value;
                }
                this.dataGridView5.Rows.Add(rowData);
                this.dataGridView6.Rows.Add(rowData);
            }
        }

        private void button6_Click(object sender, EventArgs e) //
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                object[] rowData = new object[row.Cells.Count];
                for (int i = 0; i < rowData.Length; ++i)
                {
                    rowData[i] = row.Cells[i].Value;
                }
                rowData[5] = 0;
                this.dataGridView4.Rows.Add(rowData);
                int k = 0;
                while (k < 5)
                {
                    dataGridView2.Columns[k].ReadOnly = true;
                    k++;
                }
            }
        }
        
        void ZapolneniyData2() //заполнение 2 табл
        {
            dataGridView4.ColumnCount = 6;
            dataGridView4.Columns[0].Name = "Код";
            dataGridView4.Columns[1].Name = "Название";
            dataGridView4.Columns[2].Name = "Назначене";
            dataGridView4.Columns[3].Name = "Вес_в_Г";
            dataGridView4.Columns[4].Name = "Цена_в_Р";
            dataGridView4.Columns[5].Name = "Количество";
            dataGridView5.ColumnCount = 6;
            dataGridView5.Columns[0].Name = "Код";
            dataGridView5.Columns[1].Name = "Название";
            dataGridView5.Columns[2].Name = "Назначене";
            dataGridView5.Columns[3].Name = "Вес_в_Г";
            dataGridView5.Columns[4].Name = "Цена_в_Р";
            dataGridView5.Columns[5].Name = "Количество";
            dataGridView6.ColumnCount = 6;
            dataGridView6.Columns[0].Name = "Код";
            dataGridView6.Columns[1].Name = "Название";
            dataGridView6.Columns[2].Name = "Назначене";
            dataGridView6.Columns[3].Name = "Вес_в_Г";
            dataGridView6.Columns[4].Name = "Цена_в_Р";
            dataGridView6.Columns[5].Name = "Количество";
        }

        private void button7_Click(object sender, EventArgs e) //удаление
        {
            foreach (DataGridViewRow row in dataGridView4.SelectedRows)
            {
                if (row.Index != dataGridView4.RowCount - 1)
                {
                    dataGridView4.Rows.RemoveAt(row.Index);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e) //doc
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.doc)|*.doc";

            sfd.FileName = "export" + DateTime.Now.ToString("d-M-yyyy HH-mm-ss") + ".doc";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataGridView4, sfd.FileName);
            }
        }

        private void button10_Click(object sender, EventArgs e) //очистка
        {
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
            ZapolneniyData2();
        }

        public void Export_Data_To_Word(DataGridView DGV, string filename) 
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 12;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Color = Word.WdColor.wdColorBlack;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header 
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Заявление на покупку";
                    headerRange.Font.Size = 16;
                    headerRange.Font.Color = Word.WdColor.wdColorBlack;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                var pText = oDoc.Paragraphs.Add();
                pText.Format.SpaceAfter = 10f;
                pText.Range.Text = String.Format("ООО <<Природа>>");
                pText.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                pText.Range.InsertParagraphAfter();
                pText.Range.Text = String.Format("Дата: " + DateTime.Now.ToString());
                pText.Range.InsertParagraphAfter();
                pText.Range.Text = String.Format("Зам. Директор: Бараксанов С.М.");
                pText.Range.InsertParagraphAfter();
                pText.Range.Text = String.Format("Подпись: ____________");
                pText.Range.InsertParagraphAfter();
                pText.Format.SpaceAfter = 10f;
                pText.Range.Text = String.Format("Директор: Козырев А.С.");
                pText.Range.InsertParagraphAfter();
                pText.Range.Text = String.Format("Подпись: ____________");
                pText.Range.InsertParagraphAfter();
                //save
                oDoc.SaveAs2(filename);
            }
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox tbViewEdit = e.Control as TextBox;
            if (tbViewEdit != null)
            {
                tbViewEdit.TextChanged += new EventHandler(tbViewEdit_TextChanged);
            }
        }
        private void tbViewEdit_TextChanged(object sender, EventArgs e)
        {
            TextBox dgCellEdit = (TextBox)sender;
            if (dgCellEdit.Text == "")
            {

            }
            else
            {
                int a = Convert.ToInt32(dgCellEdit.Text);
                int i = Convert.ToInt32(dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[0].Value.ToString());
                a = a * Convert.ToInt32(dataGridView1.Rows[i - 1].Cells[4].Value.ToString());
                dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[4].Value = a;
                a = Convert.ToInt32(dgCellEdit.Text);
                a = a * Convert.ToInt32(dataGridView1.Rows[i - 1].Cells[3].Value.ToString());
                dataGridView4.Rows[dataGridView4.CurrentCell.RowIndex].Cells[3].Value = a;
            }

        }
    }
}
