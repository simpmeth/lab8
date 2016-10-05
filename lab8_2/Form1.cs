using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

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


        private void добавитьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild.Controls.Count == 0) return;
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            var column = new DataGridViewColumn();
            column.HeaderText = "column" + (dataGridView.Columns.Count);
            column.Name = "column" + (dataGridView.Columns.Count); 
            column.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView.Columns.Add(column);
           
        }

        private void добавитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild.Controls.Count == 0) return;
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            dataGridView.Rows.Add();
        }

        private void удалитьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild.Controls.Count == 0) return;
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            if (dataGridView.SelectedCells.Count!=0 )
            for (int i = 0; i <= dataGridView.SelectedCells.Count-1; i++) {
                var index = dataGridView.SelectedCells[i].ColumnIndex;
                dataGridView.Columns.Remove(dataGridView.Columns[index].Name);
            }
        }

        private void удалитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.ActiveMdiChild.Controls.Count == 0) return;
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;
            if (dataGridView.SelectedCells.Count != 0)
                for (int i = 0; i <= dataGridView.SelectedCells.Count-1; i++)
                {

                    var index = dataGridView.SelectedCells[i].RowIndex;
                    if (!dataGridView.CurrentRow.IsNewRow)
                    dataGridView.Rows.RemoveAt(index);
                }
        }


      
        private void среднееЗначениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
           try {
                
                if (this.ActiveMdiChild.Controls.Count == 0) return;
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
                if (this.ActiveMdiChild.Controls.Count == 0) return;
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
            if (this.ActiveMdiChild.Controls.Count == 0) return;
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
            if (this.ActiveMdiChild.Controls.Count == 0) return;
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
            if (this.ActiveMdiChild.Controls.Count == 0) return;
            var dataGridView = this.ActiveMdiChild.Controls[0] as DataGridView;

            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView, sfd.FileName);
            }
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
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


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
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

               
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

                //NASSIM LOUCHANI
            }
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
