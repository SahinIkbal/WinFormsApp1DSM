using OfficeOpenXml;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private ExcelPackage excelPackage;

        public ExcelWorksheet? WS { get; set; }
        public Form1()
        {
            InitializeComponent();
            Start();
        }

        private void Start()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var stream = new FileInfo(@"C:\Users\Sahin\Desktop\New folder\Book1.xlsx");
            excelPackage = new ExcelPackage(stream);
            WS = excelPackage.Workbook.Worksheets[0];

            ReadMethod(WS,comboBox1,0);
            ReadMethod(WS,comboBox2,1);
            ReadMethod(WS,comboBox3,2);
            ReadMethod(WS,comboBox4,3);

        }

        private void ReadMethod(ExcelWorksheet ws,ComboBox comboBox,int index)
        {
            var state = true;
            int i = 0;
            while (state == true)
            {
                var test = ws.Cells.GetCellValue<string>(i, index);
                if (String.IsNullOrEmpty(test))
                    break;
                comboBox.Items.Add(test);
                i++;

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WriteMethod();
            excelPackage.Save();
            Close();
        }

        private void WriteMethod()
        {
            int i = 0;
            while (true)
            {
                if (string.IsNullOrEmpty(WS.Cells.GetCellValue<string>(i, 7)))
                {
                    if (i!=0)
                    {
                        i++; 
                    }
                    break;
                }
                i++;
            }
            var s = $"Project Code is:{comboBox1.Text}, Activity Number is:{comboBox2.Text}, File Name is:{comboBox3.Text}, Error Description is:" +
                $"{comboBox4.Text}";
            WS.Cells.SetCellValue(i, 7, s);
        }

    
    }
}