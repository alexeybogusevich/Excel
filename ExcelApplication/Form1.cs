using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Antlr4.Runtime;
using System.IO;

namespace ExcelApplication
{
    public partial class Form1 : Form
    {
        const int Measures = 200;
        int rowNum = 0;
        int columnNum = 0;
        public Dictionary<string, MyCell> dictionary = new Dictionary<string, MyCell>();

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
            saveFileDialog1.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";

            CreateDataGridView(Measures, Measures);

            for (int i = 0; i < rowNum; ++i)
            {
                for (int j = 0; j < columnNum; ++j)
                {
                    string cellName = SetCellName(j, i);
                    MyCell cell = new MyCell(cellName);
                    cell.Value = 0;
                    cell.Exp = "0";
                    cell.Column = j;
                    cell.Row = i;
                    dictionary.Add(cellName, cell);
                }
            }

            this.ActiveControl = dgv;
        } //конструктор форми

        private void CreateDataGridView(int columns, int rows)
        {
            for (int i = 0; i < columns; ++i)
            {
                DataGridViewColumn A = new DataGridViewColumn();
                A.HeaderText = SetColumnName(i);
                A.Name = A.HeaderText;
                MyCell cellA = new MyCell();
                A.CellTemplate = cellA;
                dgv.Columns.Add(A);
                ++columnNum;
            }

            for (int i = 0; i < rows; ++i)
            {
                DataGridViewRow A = new DataGridViewRow();
                A.HeaderCell.Value = i.ToString();
                dgv.Rows.Add(A);
                ++rowNum;
            }
        } //створює таблицю

        private void AddRow_Click(object sender, EventArgs e)
        {
            DataGridViewRow A = new DataGridViewRow();
            A.HeaderCell.Value = rowNum.ToString();
            dgv.Rows.Add(A);
            ++rowNum;
            for (int j = 0; j < dgv.ColumnCount; ++j)
            {
                string cellName = SetCellName(j, rowNum - 1);
                MyCell cell = new MyCell(cellName);
                cell.Value = 0;
                cell.Exp = "0";
                cell.Column = j;
                cell.Row = A.Index;
                dictionary.Add(cellName, cell);
            }
        }   

        private void DeleteRow_Click(object sender, EventArgs e)
        {
            int ind = dgv.Rows.Count - 1;
            for(int i=0; i<dgv.ColumnCount; ++i)
            {
                if(dictionary[SetCellName(i, ind)].DependentOnMe.Count != 0 || dgv[i, ind].Value != null)
                {
                    MessageBox.Show("Some cells in the row may have dependencies." +
                        "\nUnable to delete this one.");
                    return;
                }
            }

            for (int j = 0; j < dgv.ColumnCount; ++j)
            {
                string cellName = SetCellName(j, dgv.RowCount-1);
                dictionary.Remove(cellName);
            }
            dgv.Rows.RemoveAt(ind);
            --rowNum;
            for(int i=0; i<rowNum; ++i)
                dgv.Rows[i].HeaderCell.Value = i.ToString();
        }

        private void AddColumn_Click(object sender, EventArgs e)
        {
            DataGridViewColumn A = new DataGridViewColumn();
            A.HeaderText = SetColumnName(columnNum);
            A.Name = A.HeaderText;
            MyCell cellA = new MyCell();
            A.CellTemplate = cellA;
            dgv.Columns.Add(A);
            ++columnNum;
            for (int j = 0; j < dgv.RowCount; ++j)
            {
                string cellName = SetCellName(columnNum - 1, j);
                MyCell cell = new MyCell(cellName);
                cell.Value = 0;
                cell.Exp = "0";
                cell.Row = j;
                cell.Column = A.Index;
                dictionary.Add(cellName, cell);
            }
        }

        private void DeleteColumn_Click(object sender, EventArgs e)
        {
            int ind = dgv.Columns.Count - 1;
            for (int i = 0; i < dgv.RowCount; ++i)
            {
                if (dictionary[SetCellName(ind, i)].DependentOnMe.Count != 0 || dgv[ind,i].Value != null)
                {
                    MessageBox.Show("Some cells in the column may have dependencies." +
                        "\nUnable to delete this one.");
                    return;
                }
            }

            for (int j = 0; j < dgv.RowCount; ++j)
            {
                string cellName = SetCellName(dgv.ColumnCount-1, j);
                dictionary.Remove(cellName);
            }

            dgv.Columns.RemoveAt(ind);
            --columnNum;
            for(int i=0; i<columnNum; ++i)
                dgv.Columns[i].HeaderText = SetColumnName(i);
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //string cellName = SetCellName(e.ColumnIndex, e.RowIndex);
            if (dgv[e.ColumnIndex, e.RowIndex].Value.ToString() == "0")
                dgv[e.ColumnIndex, e.RowIndex].Value = null;
        }

        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            string cellName = SetCellName(dgv.CurrentCell.ColumnIndex, dgv.CurrentCell.RowIndex);
            if (dictionary[cellName].Exp != "0")
            {
                textBox1.Text = dictionary[cellName].Exp;
                dgv[e.ColumnIndex, e.RowIndex].Value = dictionary[cellName].Value.ToString();
            }
            else
                textBox1.Text = "";
        }

        private void dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (Char.IsLetter((char)e.KeyValue) || Char.IsNumber((char)e.KeyValue))
            {
                if (textBox1.Text == "0")
                {
                    textBox1.Clear();
                }
                
                textBox1.Text += (char)e.KeyValue;
                dgv.CurrentCell.Value = textBox1.Text;
                this.ActiveControl = textBox1;
                textBox1.Focus();
                textBox1.SelectionStart = textBox1.Text.Length;
            }
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Focus();
            this.ActiveControl = textBox1;
            textBox1.SelectionStart = textBox1.Text.Length;
        }

        private void dgv_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            string cellName = SetCellName(dgv.CurrentCell.ColumnIndex, dgv.CurrentCell.RowIndex);
            if (dictionary[cellName].Value != 0)
                textBox1.Text = dictionary[cellName].Exp;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            string cellName = SetCellName(dgv.CurrentCell.ColumnIndex, dgv.CurrentCell.RowIndex);

            SetCell(cellName);
            RefreshCells(cellName);

            textBox1.BackColor = SystemColors.Window;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.BackColor = SystemColors.Control;
            dgv[dgv.CurrentCell.ColumnIndex, dgv.CurrentCell.RowIndex].Value = textBox1.Text;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = dgv;
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = saveFileDialog1.FileName;
            StreamWriter streamWriter = new StreamWriter(filename);
            try
            {
                streamWriter.WriteLine(dgv.ColumnCount);
                streamWriter.WriteLine(dgv.RowCount);

                for (int i = 0; i < dgv.ColumnCount; ++i)
                {
                    for (int j = 0; j < dgv.RowCount; ++j)
                    {
                        string cellName = SetCellName(i, j);
                        streamWriter.WriteLine(dictionary[cellName].Exp);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("An error occurred while writing the file.");
            }
            finally
            {
                streamWriter.Close(); 
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int columns = 0;
            int rows = 0;
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = openFileDialog1.FileName;
            StreamReader streamReader = new StreamReader(filename);
            columns = Convert.ToInt32(streamReader.ReadLine());
            rows = Convert.ToInt32(streamReader.ReadLine());

            Clear(dgv);
            CreateDataGridView(columns, rows);
            dictionary.Clear();

            try
            {
                for (int i = 0; i < columns; ++i)
                {
                    for (int j = 0; j < rows; ++j)
                    {
                        string cellName = SetCellName(i, j); 
                        MyCell cell = new MyCell(cellName);
                        cell.Value = 0;
                        cell.Exp = streamReader.ReadLine();
                        cell.Column = i;
                        cell.Row = j;
                        dictionary.Add(cellName, cell);
                    }
                }

                for( int i=0; i< columns; ++i)
                {
                    for(int j=0; j< rows; ++j)
                    {
                        string cellName = SetCellName(i, j);
                        string input = dictionary[cellName].AddressToValue(ref dictionary);
                        var output = Calculator.Evaluate(input);
                        dictionary[cellName].Value = output;
                        if (output != 0)
                            dgv[i, j].Value = output;
                    }
                }

                for(int i=0; i<columns; ++i)
                {
                    for(int j=0; j<rows; ++j)
                    {
                        string cellName = SetCellName(i, j);
                        RefreshCells(dictionary[cellName].Name);
                    }
                }

            }
            catch (Exception)
            {
                MessageBox.Show("An error occurred while writing the file.");
            }
            finally
            {
                streamReader.Close();
            }
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Welcome to KNUExcel 2018! \n" +
                "Actions you are allowed to do are the following: \n" +
                "# Calculate cell values using operators + - / * Min(;) Max(;) ^ \n" +
                "# Add/Delete the last current row/column \n" +
                "# Save a grid to/Load a grid from a text file \n" +
                "# Send us a good feedback via email: *link*");
        }

        private void SetRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1);
            }
        } //задає заголовки рядкам

        public string SetColumnName(int i)
        {
            const int alphabet = 26;
            string name = "";
            int n = i;
            if (i < alphabet)
            {
                char c = (char)(i + 65);
                return (name + c);
            }

            for (int j = i; j >= alphabet; j = j / alphabet)
            {
                int mod = i % alphabet;
                int div = i / alphabet - 1;

                name += (char)(div + 65);

                if (div < alphabet)
                {
                    name += (char)(mod + 65);
                }
            }
            return name;
        } //задає заголовки стовбчикам

        public string SetCellName(int columnIndex, int rowIndex)
        {
            return dgv.Columns[columnIndex].HeaderText
                + dgv.Rows[rowIndex].HeaderCell.Value;
        } //повертає ім'я клітинки за індексами

        private void RefreshCells(string initialCell)
        {
            foreach (string item in dictionary[initialCell].DependentOnMe)
            {
                dictionary[item].Value = Calculator.Evaluate(dictionary[item].AddressToValue(ref dictionary));
                if (dictionary[item].Value.ToString() == "0")
                    dgv[dictionary[item].Column, dictionary[item].Row].Value = null;
                else
                    dgv[dictionary[item].Column, dictionary[item].Row].Value = dictionary[item].Value;
                RefreshCells(item);
            }
        } //оновлення клітин, що залежать від даної

        public bool RecurrenceCheck(MyCell Current, MyCell Initial)
        {
            if (Current.IDependOn.Contains(Initial.Name))
                return false; //pominyav
            foreach (string cellName in Current.IDependOn)
            {
                if (RecurrenceCheck(dictionary[cellName], Initial))
                    return true;
            }
            return false;
        } //перевірка на наявність рекурсії

        public void ReccurenceSearch(string initialCell, ref List<string> visited)
        {
            if (!visited.Contains(initialCell))
                visited.Add(initialCell);
            foreach (string item in dictionary[initialCell].DependentOnMe)
            {
                if (!visited.Contains(item))
                    ReccurenceSearch(item, ref visited);
            }
        } //знаходження всіх комірок рекурсії

        private void SetCell(string cellName)
        {
            string input;
            try
            {
                if (textBox1.Text.Contains(cellName))
                {
                    MessageBox.Show("Reccurence is present\nThe cell will be cleared.");
                    throw new Exception();
                }
                if (textBox1.Text.Contains('.'))
                {
                    MessageBox.Show("Invalid Expression\nThe cell will be cleared.");
                    throw new Exception();
                }

                dictionary[cellName].Exp = textBox1.Text;

                input = dictionary[cellName].AddressToValue(ref dictionary);
                if (RecurrenceCheck(dictionary[cellName], dictionary[cellName]))
                {
                    MessageBox.Show("Reccurence is present\nThe cells in it will be cleared.");
                    List<string> reccurenceList = new List<string>();
                    ReccurenceSearch(cellName, ref reccurenceList);
                    foreach (string item in reccurenceList)
                    {
                        dictionary[item].Value = 0;
                        dictionary[item].Exp = "0";
                        dictionary[item].IDependOn.Clear();
                        dictionary[item].DependentOnMe.Clear();
                        dgv[dictionary[item].Column, dictionary[item].Row].Value = null;
                    }
                    throw new Exception();
                }
                input = dictionary[cellName].AddressToValue(ref dictionary);

                if (input == "" || dictionary[cellName].Exp == "0")
                {
                    dictionary[cellName].Value = 0;
                    dictionary[cellName].Exp = "0";
                    dictionary[cellName].IDependOn.Clear();
                    RefreshCells(cellName);
                    throw new Exception();
                }

                try
                {
                    var res = Calculator.Evaluate(input);
                    dictionary[cellName].Value = res;
                    if (res == 0)
                    {
                        dgv.CurrentCell.Value = null;
                        RefreshCells(cellName);
                    }
                    else
                        dgv.CurrentCell.Value = res.ToString();
                }
                catch (Exception)
                {
                    MessageBox.Show("Invalid Expression.\nThe cell will be cleared.");
                    dictionary[cellName].Value = 0;
                    dictionary[cellName].Exp = "0";
                    dictionary[cellName].IDependOn.Clear();
                    throw new Exception();
                }
                if (dgv.CurrentCell.Value.ToString() == "∞")
                {
                    MessageBox.Show("Unable to calculate infinity.\nThe cell will be cleared.");
                    dictionary[cellName].Value = 0;
                    dictionary[cellName].Exp = "0";
                    dictionary[cellName].IDependOn.Clear();
                    RefreshCells(cellName);
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                dgv.CurrentCell.Value = null;
                textBox1.BackColor = SystemColors.Window;
                return;
            }
        } //обробка комірки

        public void Clear(DataGridView dataGridView)
        {
            dgv.Rows.Clear();
            int count = dgv.ColumnCount;
            for (int i = 0; i < count; ++i)
                dgv.Columns.RemoveAt(0);
        } //знищує таблицю
    }

}
