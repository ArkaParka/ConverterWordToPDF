using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace Converter_WordToPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            var wordDocument = new Document();
            Word.Application appWord = new Word.Application();

            button1.Click += button1_Click;
            button2.Click += button2_Click;
            openFileDialog1.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
            saveFileDialog1.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
                        
        }


        private void button1_Click(object sender, EventArgs e)
        {
                       
            var wordDocument = new Document();

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {                           

                Word.Application appWord = new Word.Application();
                wordDocument = appWord.Documents.Add(@"C:\Users\Арина\Desktop\hello.docx");
                return;
            }
               
            
            MessageBox.Show("Файл открыт");
                        
        }


        private void button2_Click(object sender, EventArgs e)
        {            
            var wordDocument = new Document();
                        
            wordDocument.ExportAsFixedFormat(@"C:\Users\Арина\Desktop\hello.pdf", WdExportFormat.wdExportFormatPDF);
                      
            
            MessageBox.Show("Файл сохранен");
        }
                
    }
}
