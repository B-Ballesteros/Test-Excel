using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            openButton.Click += OpenButtonClick;
            saveButton.Click += SaveButtonClick;
            addColumnButton.Click += addColumnClick;
        }

        void openFile()
        {
            if(oFDialog.ShowDialog() == DialogResult.OK)
            {
                sFDialog.FileName = oFDialog.FileName;
                try
                {
                    gridView.DataSource = ExcelHandler.ReadFile(oFDialog.FileName);
                }catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        void saveFile()
        {
            if(sFDialog.ShowDialog() ==DialogResult.OK)
            {
                try
                {
                    ExcelHandler.WriteFile(sFDialog.FileName, grid2Table());

                }catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        void addColumn2Grid()
        {
            gridView.Columns.Add(new DataGridViewTextBoxColumn());
        }


        DataTable grid2Table()
        {
            DataTable dt = new DataTable();
            int columns = gridView.ColumnCount;
            for(int i = 0; i  < columns; i++)
            {
                dt.Columns.Add();
            }
            foreach (DataGridViewRow row in gridView.Rows)
            {
                DataRow dRow = dt.NewRow();
                foreach(DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dRow);
            }
            return dt;
        }

        private void OpenButtonClick(object sender, EventArgs e)
        {
            openFile();
        }

        private void SaveButtonClick(object sender, EventArgs e)
        {
            saveFile();
            
        }

        private void addColumnClick(object sender, EventArgs e)
        {
            addColumn2Grid();
        }

    }
}
