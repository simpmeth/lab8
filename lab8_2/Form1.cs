using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace lab8_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void созданиеФайлаСобственногоФорматаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new Form();
            form.MdiParent = this;
            form.Show();
            form.FormClosing += new FormClosingEventHandler(this.Editor_FormClosing);
          

            Type t = typeof(Form);
            PropertyInfo pi = t.GetProperty("MdiClient", BindingFlags.Instance | BindingFlags.NonPublic);
            MdiClient cli = (MdiClient)pi.GetValue(form.MdiParent, null);
            ActiveMdiChild.Location = new Point(0, 0);
            ActiveMdiChild.Size = new Size(cli.Width - 4, cli.Height - 4);

            var dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.ContextMenuStrip = contextMenuStrip1;

            var column = new DataGridViewColumn();
            column.HeaderText = "column0";
            column.Name = "column0";
            column.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView.Columns.Add(column);
            dataGridView.Rows.Add();

            form.Controls.Add(dataGridView);
        }

        
        private void закрытиеФайлаСобственногоФорматаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ActiveMdiChild != null) {
                ActiveMdiChild.Close();
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var result = MessageBox.Show("Хотите выйти?", "", MessageBoxButtons.YesNo);
            if ( result== DialogResult.No)
            {
                
                e.Cancel = true;                
            }
        }

        private void Editor_FormClosing(object sender, FormClosingEventArgs e)
        {
            var result = MessageBox.Show("Хотите закрыть документ?", "", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {

                e.Cancel = true;
            }
        }

        private void расположениеОконКаскадомToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.Cascade);
        }

        private void расположениеОконМозаикойToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void расположениеОконВертикальноToolStripMenuItem_Click(object sender, EventArgs e)
        {

            this.LayoutMdi(MdiLayout.TileVertical);
        }

        private void расположениеОконГоризонтальноToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void заданиеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void документToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void добавитьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;


            var column = new DataGridViewColumn();
            column.HeaderText = "column" + (dataGridView.Columns.Count);
            column.Name = "column" + (dataGridView.Columns.Count); 
            column.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView.Columns.Add(column);
           
        }

        private void добавитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            dataGridView.Rows.Add();
        }

        private void удалитьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            if (dataGridView.SelectedCells.Count!=0 )
            for (int i = 0; i <= dataGridView.SelectedCells.Count-1; i++) {
                var index = dataGridView.SelectedCells[i].ColumnIndex;
                dataGridView.Columns.Remove(dataGridView.Columns[index].Name);
            }
        }

        private void удалитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            if (dataGridView.SelectedCells.Count != 0)
                for (int i = 0; i <= dataGridView.SelectedCells.Count-1; i++)
                {
                    var index = dataGridView.SelectedCells[i].RowIndex;
                    dataGridView.Rows.RemoveAt(index);
                }
        }

        private void среднееЗначениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
           try {
                var result = (double)0;
                var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
                if (dataGridView.SelectedCells.Count != 0)
                    for (int i = 0; i <= dataGridView.SelectedCells.Count - 1; i++)
                    {
                        var intValue = (double)0;
                        var value = (double) dataGridView.SelectedCells[i].Value;
                        result += value;
                    }
                MessageBox.Show("Среднее значения " + (result / dataGridView.SelectedCells.Count).ToString());
            }
            catch  {
                MessageBox.Show("Невозможно посчитать");
            }
        }

        private void словДлинаКоторыхМеньше4СимволовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var result = string.Empty;

                var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
                if (dataGridView.SelectedCells.Count != 0)
                    for (int i = 0; i <= dataGridView.SelectedCells.Count - 1; i++)
                    {
                        var intValue = (double)0;
                        var value =  dataGridView.SelectedCells[i].Value as string;
                        if (value !=null)
                        if (value.Length < 4)
                            result += value;
                    }
                
                dataGridView.Rows.Add();
                var filter = new DataGridViewElementStates();
                
                var lastindex = dataGridView.Rows.GetLastRow(filter);
                dataGridView.Rows[lastindex-1].Cells[0].Value = result;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Невозможно посчитать"+ex.ToString());
            }
        }

        private void открытиеФайлаСобственногоФорматаToolStripMenuItem_Click(object sender, EventArgs e)
        {


            string filename = "";
            var ofd = new OpenFileDialog();
            ofd.Filter = "FILE (*.file)|*.file";
            ofd.FileName = "Output.file";
            if (ofd.ShowDialog() == DialogResult.OK)
            {

                var form = new Form();
                form.MdiParent = this;
                form.Show();
                form.FormClosing += new FormClosingEventHandler(this.Editor_FormClosing);


                Type t = typeof(Form);
                PropertyInfo pi = t.GetProperty("MdiClient", BindingFlags.Instance | BindingFlags.NonPublic);
                MdiClient cli = (MdiClient)pi.GetValue(form.MdiParent, null);
                ActiveMdiChild.Location = new Point(0, 0);
                ActiveMdiChild.Size = new Size(cli.Width - 4, cli.Height - 4);

                var dataGridView = new DataGridView();
                dataGridView.Dock = DockStyle.Fill;
                dataGridView.ContextMenuStrip = contextMenuStrip1;

                


                form.Controls.Add(dataGridView);

                var isColumnName = true;
                using (FileStream file = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.Read, 4096))
                using (StreamReader reader = new StreamReader(file))
                {
                    while (!reader.EndOfStream)
                    {
                        var fields = reader.ReadLine().Split(',');
                        if (fields.Length >= 2 && (fields[0] != "" || fields[1] != ""))
                        {
                            if (isColumnName)
                            {
                                foreach (string header in fields)
                                {
                                    var column = new DataGridViewColumn();
                                    column.HeaderText = header.ToString();
                                    column.Name = header.ToString();
                                    column.CellTemplate = new DataGridViewTextBoxCell();
                                    dataGridView.Columns.Add(column);
                                }

                                isColumnName = false;
                            }
                            else
                            dataGridView.Rows.Add(fields);
                        }
                    }
                }
            }
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void сохранениеФайлаСобственногоФорматаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            string filename = "";
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "FILE (*.file)|*.file";
            sfd.FileName = "Output.file";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                
                if (File.Exists(filename))
                {
                    try
                    {
                        File.Delete(filename);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("Невозможно сохранить." + ex.Message);
                    }
                }
                int columnCount = dataGridView.ColumnCount;
                string columnNames = "";
                string[] output = new string[dataGridView.RowCount + 1];
                for (int i = 0; i < columnCount; i++)
                {
                    columnNames += dataGridView.Columns[i].Name.ToString() + ",";
                }
                columnNames= columnNames.Remove(columnNames.Length - 1);
                output[0] += columnNames;
                for (int i = 1; (i - 1) < dataGridView.RowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        var str = dataGridView.Rows[i - 1].Cells[j].Value;
                        if (str == null) str = string.Empty;

                        output[i] += str.ToString() + ",";
                    }
                    output[i]= output[i].Remove(output[i].Length - 1);
                }
                File.WriteAllLines(sfd.FileName, output, System.Text.Encoding.UTF8);
                MessageBox.Show("Сохранено");
            }
        }

        

        private void сохранениеТаблицыВФайлФорматаExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "Inventory_Adjustment_Export.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                
                copyAlltoClipboard();

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; 
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                
                Excel.Range rng = xlWorkSheet.get_Range("D:D").Cells;
                rng.NumberFormat = "@";

                
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                delRng.Delete(Type.Missing);
                xlWorkSheet.get_Range("A1").Select();

                
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

               
                Clipboard.Clear();
                dataGridView.ClearSelection();

                
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
            }

        }

        private void сохранениеТаблицыВФайлФорматаWordToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


        

        private void copyAlltoClipboard()
        {
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            dataGridView.SelectAll();
            DataObject dataObj = dataGridView.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
