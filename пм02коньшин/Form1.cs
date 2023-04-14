using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace пм02коньшин
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

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //вспомагательные переменные
            double width = Convert.ToDouble(textBox1.Text);
            double height = Convert.ToDouble(textBox2.Text);
            double cost1 = 213.15;
            double cost2 = 265.80;
            //выбор типа потолка
            if (radioButton1.Checked)
            {
                double answer = width * height * cost1;
                if (checkBox1.Checked) //реализация выбора доп. параметров
                {
                    answer += answer * 0.3;
                    richTextBox1.Text = $"Итоговая стоимость: {answer}";
                    label3.Text = Convert.ToString(answer);
                }
                else if (checkBox2.Checked)
                {
                    answer += answer * 0.26;
                    richTextBox1.Text = $"Итоговая стоимость: {answer}";
                    label3.Text = Convert.ToString(answer);
                }
                else {
                    richTextBox1.Text = $"Итоговая стоимость: {answer}";
                    label3.Text = Convert.ToString(answer);
                }
            }
            else if (radioButton2.Checked)
            {
                double answer = width * height * cost2;
                if (checkBox1.Checked)
                {
                    answer += answer * 0.3;
                    richTextBox1.Text = $"Итоговая стоимость: {answer}";
                    label3.Text = Convert.ToString(answer);
                }
                else if (checkBox2.Checked)
                {

                    answer += answer * 0.26;
                    richTextBox1.Text = $"Итоговая стоимость: {answer}";
                    label3.Text = Convert.ToString(answer);
                }
                else
                {
                    richTextBox1.Text = $"Итоговая стоимость: {answer}";
                    label3.Text = Convert.ToString(answer);
                }
            }
            else MessageBox.Show("Выберите тип потолка!");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string itogSum = label3.Text;
            try
            {
                //в файле хранится номер для заказа
                int code = Convert.ToInt32(File.ReadAllText(@"код.txt"));
                string date = DateTime.Now.ToShortDateString();
                File.Copy(@"чек.docx", @"чеки\" + code + "_" + date + "_" + itogSum + ".docx");
                //копируем образец чека и работаем с ним
                Word.Document doc = null;
                Word.Application app = new Word.Application();
                //путь к данным
                string sourse = @"C:\Users\student1\Desktop\пм02коньшин\пм02коньшин\bin\Debug\чеки\" + code + "_" + date + "_" + itogSum + ".docx";
                doc = app.Documents.Open(sourse);
                doc.Activate();
                Word.Bookmarks book = doc.Bookmarks;
                Word.Range range;
                int i = 0;
                
                string[] data = new string[] { code.ToString(), date.ToString(), radioButton1.Text, itogSum.ToString() };
                foreach (Word.Bookmark b in book)
                {
                    range = b.Range;
                    range.Text = data[i];
                    i++;
                }
                doc.Close();
                doc = null;
                code++;
                //перезаписываем код в файл
                File.WriteAllText(@"код.txt", code.ToString());
                MessageBox.Show("Успешно!");
            }
            catch
            {
                MessageBox.Show("Ошибка, попробуйте ещё раз");
            }
        }
    }
}
