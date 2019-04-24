using System;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Shape

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void canvas_Paint(object sender, PaintEventArgs e)
        {
            // double Length, Breadth, Area;
            //Console.WriteLine("Length");
            //Length = double.Parse(Console.ReadLine());
            //Console.WriteLine("Breadth");
            //Breadth = double.Parse(Console.ReadLine());
            //Area = Length * Breadth;
            //Console.WriteLine("The area is " + Area);
            //Console.WriteLine("Hello World!");
            Graphics gObject = canvas_CreateGraphics();

            Brush red = new SolidBrush(Color.Red);
            Pen redPen = new Pen(red, 8);

            gObject.DrawLine(redPen, 10, 10, 400, 376); 

        }
    }
}
