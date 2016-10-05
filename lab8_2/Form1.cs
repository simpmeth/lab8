using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

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

        }
    }
}
