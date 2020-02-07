using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Interop.Excel;
namespace siralama
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

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application objexcel = new Microsoft.Office.Interop.Excel.Application();
            objexcel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook objbook = objexcel.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet objsheet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range objrange;
            objrange = objsheet.get_Range("A1", System.Reflection.Missing.Value);
            objrange.set_Value(System.Reflection.Missing.Value, textBox1.Text);
        }
    }
}
