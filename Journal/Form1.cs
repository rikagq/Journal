using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Journal
{
    public partial class Form1 : Form
    {

        string mainPath = "C:\\Users\\" + Environment.UserName + "\\Documents\\Journal.xlsx";
        string mainPath2 = "C:\\Users\\" + Environment.UserName + "\\Documents\\";

        private Excel.Application excelapp;
        private Excel.Window excelWindow;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;
        int numberOfStudents = 0;
        int errorCount = 0;
        string magic = "ZZ1";

        string A = "A";
        string B = "B";

        string C = "U";
        string D = "V";
        string E = "C";
        string F = "D";
        string G = "E";
        string H = "F";
        string I = "G";
        string J = "H";
        string K = "I";
        string L = "J";
        string M = "K";
        string N = "L";
        string O = "M";
        string P = "N";
        string Q = "O";
        string R = "P";
        string S = "Q";
        string T = "R";
        string U = "S";
        string V = "T";
        
        string W = "W";
        string X = "X";
        string Y = "Y";
        string Z = "Z";
        string AA = "AA";
        string AB = "AB";
        string AC = "AC";
        string AD = "AD";
        string AE = "AE";
        string AF = "AF";
        string AG = "AG";
        string AH = "AH";
        string AI = "AI";
        string AJ = "AJ";
        string AK = "AK";

        string AL = "AL";
        string AM = "AM";
        string AN = "AN";
        string AO = "AO";
        string AP = "AP";


        ////////////////////////////////////////////////////////////////
        bool debugMode = true;
        ////////////////////////////////////////////////////////////////
        
        public void CreateDir()
        {
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(mainPath2))
                {
                    return;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }
        public void CreateFile()
        {
            if (File.Exists(mainPath))
            {
                return;
            }
            else
            {
                try
                {
                    // Create the file, or overwrite if the file exists.
                    File.Create(mainPath);
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }
        public void clearThings()
        {
            //tB1.Clear();
            tB2.Clear();
            tB3.Clear();
            tB4.Clear();
            tB5.Clear();
            tB6.Clear();
            tB7.Clear();
            tB8.Clear();
            tB9.Clear();
            tB10.Clear();
            tB11.Clear();
            tB12.Clear();
            tB13.Clear();
            tB14.Clear();
            tB15.Clear();
            textBox1.Clear();
            comboBox1.SelectedIndex = 0;
            button1.PerformClick();
            button2.PerformClick();
        }
        public void testMode()
        {
            //tB1.Text = "11";
            tB2.Text = "test";
            tB3.Text = "test";
            tB4.Text = "test";
            tB5.Text = "test";
            tB6.Text = "12345678901234";
            tB7.Text = "test";
            tB8.Text = "test";
            tB9.Text = "1235";
            tB10.Text = "123456";
            tB11.Text = "teest";
            tB12.Text = "teest";
            tB13.Text = "teest";
            tB14.Text = "teest@asd.asd";
            tB15.Text = "1234567890";
            comboBox1.SelectedIndex = 0;
        }
        public void headerCreate()
        {
            excelworksheet.get_Range($"{A}1").Value = "Дата";
            excelworksheet.get_Range($"{B}1").Value = "Номер заявления";
            excelworksheet.get_Range($"{C}1").Value = "Образование";
            excelworksheet.get_Range($"{D}1").Value = "Форма обучения";
            excelworksheet.get_Range($"{E}1").Value = "Фамилия";
            excelworksheet.get_Range($"{F}1").Value = "Имя";
            excelworksheet.get_Range($"{G}1").Value = "Отчество";
            excelworksheet.get_Range($"{H}1").Value = "Дата рождения";
            excelworksheet.get_Range($"{I}1").Value = "26.02.03";
            excelworksheet.get_Range($"{I}2").Value = "Б";
            excelworksheet.get_Range($"{J}2").Value = "К";
            excelworksheet.get_Range($"{K}1").Value = "26.02.05";
            excelworksheet.get_Range($"{K}2").Value = "Б";
            excelworksheet.get_Range($"{L}2").Value = "К";
            excelworksheet.get_Range($"{M}1").Value = "15.02.06";
            excelworksheet.get_Range($"{M}2").Value = "Б";
            excelworksheet.get_Range($"{N}2").Value = "К";
            excelworksheet.get_Range($"{O}1").Value = "23.02.01";
            excelworksheet.get_Range($"{O}2").Value = "Б";
            excelworksheet.get_Range($"{P}2").Value = "К";
            excelworksheet.get_Range($"{Q}1").Value = "35.02.09";
            excelworksheet.get_Range($"{Q}2").Value = "Б";
            excelworksheet.get_Range($"{R}2").Value = "К";
            excelworksheet.get_Range($"{S}1").Value = "35.02.10";
            excelworksheet.get_Range($"{S}2").Value = "Б";
            excelworksheet.get_Range($"{T}2").Value = "К";
            excelworksheet.get_Range($"{U}1").Value = "35.02.11";
            excelworksheet.get_Range($"{U}2").Value = "Б";
            excelworksheet.get_Range($"{V}2").Value = "К";
            excelworksheet.get_Range($"{W}1").Value = "Средний балл аттестата";
            excelworksheet.get_Range($"{X}1").Value = "Аттестат";
            excelworksheet.get_Range($"{X}2").Value = "номер";
            excelworksheet.get_Range($"{Y}2").Value = "дата выдачи";
            excelworksheet.get_Range($"{Z}2").Value = "кем выдан";
            excelworksheet.get_Range($"{AA}2").Value = "код подразделения";
            excelworksheet.get_Range($"{AB}1").Value = "Паспорт";
            excelworksheet.get_Range($"{AB}2").Value = "серия";
            excelworksheet.get_Range($"{AC}2").Value = "номер";
            excelworksheet.get_Range($"{AD}2").Value = "кем выдан";
            excelworksheet.get_Range($"{AE}2").Value = "дата выдачи";
            excelworksheet.get_Range($"{AF}2").Value = "место рождения";
            excelworksheet.get_Range($"{AG}2").Value = "место регистрации";
            excelworksheet.get_Range($"{AH}1").Value = "Контакты";
            excelworksheet.get_Range($"{AH}2").Value = "e-mail";
            excelworksheet.get_Range($"{AI}2").Value = "Телефон";
            excelworksheet.get_Range($"{AJ}1").Value = "Особые отметки";
            excelworksheet.get_Range($"{AK}1").Value = "Вакантное место";
            excelworksheet.get_Range($"{AL}1").Value = "Впервые / не впервые";
            excelworksheet.get_Range($"{AM}1").Value = "Необходимость в общежитии";
            excelworksheet.get_Range($"{AN}1").Value = "Документ об образовании";
            excelworksheet.get_Range($"{AO}1").Value = "Статус заявления";
        }
        public void formating()
        {
            excelcells = excelworksheet.get_Range("A1", "FU1000");
            excelcells.WrapText = true;
            excelcells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            excelcells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

        }
        public async void WaitSomeTime()
        {
            await Task.Delay(200);
            this.Enabled = true;
        }
        public void switchHelpFix(string cell1, string cell2, RadioButton m_rB1, 
            RadioButton m_rB2, RadioButton m_rB3)
        {
            if (m_rB1.Checked)
                excelworksheet.get_Range($"{cell1}{numberOfStudents}").Value = m_rB1.Text;
            if (m_rB2.Checked)
                excelworksheet.get_Range($"{cell2}{numberOfStudents}").Value = m_rB2.Text;
            if (m_rB3.Checked)
            {
                excelworksheet.get_Range($"{cell1}{numberOfStudents}").Value = m_rB1.Text;
                excelworksheet.get_Range($"{cell2}{numberOfStudents}").Value = m_rB2.Text;
            }
        }
        public void switchFix(ComboBox cB, RadioButton m_rB1, RadioButton m_rB2, 
            RadioButton m_rB3, string m_I, string m_J, string m_K, string m_L,
            string m_M, string m_N, string m_O, string m_P, string m_Q,
            string m_R, string m_S, string m_T, string m_U, string m_V)
        {
            switch (cB.SelectedItem)
            {
                case "26.02.03":
                    switchHelpFix(m_I, m_J, m_rB1, m_rB2, m_rB3);
                    break;
                case "26.02.05":
                    switchHelpFix(m_K, m_L, m_rB1, m_rB2, m_rB3);
                    break;
                case "15.02.06":
                    switchHelpFix(m_M, m_N, m_rB1, m_rB2, m_rB3);
                    break;
                case "23.02.01":
                    switchHelpFix(m_O, m_P, m_rB1, m_rB2, m_rB3);
                    break;
                case "35.02.09":
                    switchHelpFix(m_Q, m_R, m_rB1, m_rB2, m_rB3);
                    break;
                case "35.02.10":
                    switchHelpFix(m_S, m_T, m_rB1, m_rB2, m_rB3);
                    break;
                case "35.02.11":
                    switchHelpFix(m_U, m_V, m_rB1, m_rB2, m_rB3);//v
                    break;
            }
        }
        public void setCellFromTextBox(string cell, System.Windows.Forms.TextBox tB)
        {
            excelworksheet.get_Range($"{cell}{numberOfStudents}").Value = tB.Text;
        }

        public Form1()
        {
            InitializeComponent();
            {
                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;
                comboBox4.SelectedIndex = 0;
                comboBox7.SelectedIndex = 0;
                comboBox8.SelectedIndex = 0;
            }
            if (File.Exists(mainPath))
            {
                excelapp = new Excel.Application();
                excelapp.Visible = false;
                excelapp.Workbooks.Open(mainPath);
                excelappworkbooks = excelapp.Workbooks;
                excelappworkbook = excelappworkbooks[1];
                excelsheets = excelappworkbook.Worksheets;
                //Получаем ссылку на лист 1
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                tB1.Text = $"{(excelworksheet.get_Range(magic).Value - 1):000}";
                {
                    excelappworkbook.Saved = true;
                    excelapp.DisplayAlerts = false;
                    excelappworkbook.SaveAs(mainPath);
                    excelapp.Quit();
                }
                if (tB1.Text == "")
                {
                    tB1.Text = "001";
                }
            }
            else
            {
                tB1.Text = "001";
            }

            if (debugMode == true)
            {
                testMode();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CreateDir();
            if (!File.Exists(mainPath))
            {
                excelapp = new Excel.Application();
                excelapp.Visible = false;

                excelapp.SheetsInNewWorkbook = 1;
                excelapp.Workbooks.Add(Type.Missing);

                excelappworkbooks = excelapp.Workbooks;
                excelappworkbook = excelappworkbooks[1];
                excelappworkbook.Saved = true;
                excelapp.DisplayAlerts = false;
                excelappworkbook.SaveAs(mainPath);
                excelapp.Quit();
            }
            excelapp = new Excel.Application();
            excelapp.Visible = false;
            this.Enabled = false;
            WaitSomeTime();
            //Остановка программы Здесь
            excelapp.Workbooks.Open(mainPath);
            excelappworkbooks = excelapp.Workbooks;
            excelappworkbook = excelappworkbooks[1];
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            errorProvider1.Clear();
            {
                var textB = new[] { tB1, tB2, tB3, tB4, tB5, tB6, tB7, tB8, tB10, tB11, tB12, tB13, tB14, tB15 };
                foreach (var control in textB.Where(ee => String.IsNullOrEmpty(ee.Text)))
                {
                    errorProvider1.SetError(control, "Поле не может быть пустым");
                    errorCount++;
                }
                var fiotB = new[] { tB2, tB3, tB4 };
                foreach (var control in fiotB.Where(ee => !Regex.Match(ee.Text, @"^[A-Za-zА-Яа-яЁё]*$").Success))
                {
                    errorProvider1.SetError(control, "Поля ФИО не могут содержать цифр");
                    errorCount++;
                }
                if (!Regex.Match(tB1.Text, @"^[0-9]*$").Success)
                {
                    errorProvider1.SetError(tB1, "Некорректный номер заявления");
                    errorCount++;
                }
                if (!Regex.Match(tB14.Text, @"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$").Success)
                {
                    errorProvider1.SetError(tB14, "Некорректный email");
                    errorCount++;
                }
                if (!Regex.Match(tB15.Text, @"^[0-9]{10}$").Success)
                {
                    errorProvider1.SetError(tB15, "Номер телефона должен быть в формате 12345567890");
                    errorCount++;
                }
            }
            if (errorCount == 0)
            {
                success2.Visible = true;
                timer1.Start();
                {
                    excelcells = excelworksheet.get_Range("A1");
                    if (excelcells.Value != "")
                    {
                        excelcells = excelworksheet.get_Range(magic); // МАГИЧЕСКОЕ ЧИСЛО
                        if (excelcells.Value == null)
                        {
                            excelcells.Value = 2;
                            formating();
                            headerCreate();
                        }
                        excelcells = excelworksheet.get_Range(magic); // МАГИЧЕСКОЕ ЧИСЛО
                        excelcells.Value += 1;
                        numberOfStudents = (int)excelcells.Value;
                    }
                }

                tB1.Text = $"{(excelworksheet.get_Range(magic).Value - 1):000}";

                //Info block///////////////////////////////////////////////
                {
                    excelcells = excelworksheet.get_Range($"{A}{numberOfStudents}");
                    excelcells.Value = dateTimePicker1.Value.Date;
                    excelcells.ColumnWidth = 10;
                    excelcells = excelworksheet.get_Range($"{B}{numberOfStudents}");
                    excelcells.NumberFormat = "@";
                    excelcells.Value = Convert.ToInt32(tB1.Text) - 1;
                    excelcells.Value = $"{(excelcells.Value):000}";
                    excelworksheet.get_Range($"{C}{numberOfStudents}").Value = comboBox7.SelectedItem;
                    excelworksheet.get_Range($"{D}{numberOfStudents}").Value = comboBox4.SelectedItem;
                    setCellFromTextBox(E, tB2);
                    setCellFromTextBox(F, tB3);
                    setCellFromTextBox(G, tB4);
                    excelcells = excelworksheet.get_Range($"{H}{numberOfStudents}");
                    excelcells.Value = dateTimePicker2.Value.Date;
                    excelcells.ColumnWidth = 10;
                    //Switch
                    switchFix(comboBox1, rB1, rB2, rB3, I, J, K, L, M, N, O, P, Q, R, S, T, U, V);
                    switchFix(comboBox5, rB4, rB5, rB6, I, J, K, L, M, N, O, P, Q, R, S, T, U, V);
                    switchFix(comboBox6, rB7, rB8, rB9, I, J, K, L, M, N, O, P, Q, R, S, T, U, V);
                    setCellFromTextBox(W, tB5);
                    excelcells = excelworksheet.get_Range($"{X}{numberOfStudents}");
                    excelcells.NumberFormat = "0";
                    excelcells.ColumnWidth = 16;
                    excelcells.Value = tB6.Text;
                    excelcells = excelworksheet.get_Range($"{Y}{numberOfStudents}");
                    excelcells.Value = dateTimePicker3.Value.Date;
                    excelcells.ColumnWidth = 10;
                    setCellFromTextBox(Z, tB7);
                    setCellFromTextBox(AA, tB9);
                    setCellFromTextBox(AB, tB9);
                    setCellFromTextBox(AC, tB10);
                    setCellFromTextBox(AD, tB11);
                    excelcells = excelworksheet.get_Range($"{AE}{numberOfStudents}");
                    excelcells.Value = dateTimePicker4.Value.Date;
                    excelcells.ColumnWidth = 10;
                    setCellFromTextBox(AF, tB12);
                    setCellFromTextBox(AG, tB13);
                    setCellFromTextBox(AH, tB14);
                    excelcells = excelworksheet.get_Range($"{AI}{numberOfStudents}");
                    excelcells.ColumnWidth = 12;
                    excelcells.NumberFormat = "@";
                    excelcells.Value = $"+7{tB15.Text}";
                    excelcells = excelworksheet.get_Range($"{AJ}{numberOfStudents}");
                    excelcells.Value = comboBox2.SelectedItem;
                    if (comboBox2.SelectedItem.ToString() == "Другое")
                    {
                        excelcells.Value = textBox1.Text;
                    }
                    excelcells = excelworksheet.get_Range($"{AK}{numberOfStudents}");
                    if (checkBox1.Checked)
                        excelcells.Value = "Да";
                    else
                        excelcells.Value = "Нет";
                    excelcells = excelworksheet.get_Range($"{AL}{numberOfStudents}");
                    if (checkBox2.Checked)
                        excelcells.Value = "Да";
                    else
                        excelcells.Value = "Нет";
                    excelcells = excelworksheet.get_Range($"{AM}{numberOfStudents}");
                    if (checkBox3.Checked)
                        excelcells.Value = "Да";
                    else
                        excelcells.Value = "Нет";
                    excelworksheet.get_Range($"{AN}{numberOfStudents}").Value = comboBox8.SelectedItem;
                    excelcells = excelworksheet.get_Range($"{AO}{numberOfStudents}");
                    excelcells.Value = comboBox3.SelectedItem;
                    if (comboBox3.SelectedItem.ToString() == "Другое")
                    {
                        excelcells.Value = textBox2.Text;
                    }
                }
                
                if (debugMode == false)
                {
                    clearThings();
                }
            }
            else
            {
                MessageBox.Show("Одно или несколько из полей введены неправильно");
            }
            errorCount = 0;
            {
                excelappworkbook.Saved = true;
                excelapp.DisplayAlerts = false;
                excelappworkbook.SaveAs(mainPath);
                excelapp.Quit();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem.ToString() == "Другое")
            {
                textBox1.Visible = true;
            }
            if (comboBox2.SelectedItem.ToString() != "Другое")
            {
                textBox1.Visible = false;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            success2.Visible = false;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            rB4.Checked = false;
            rB5.Checked = false;
            rB6.Checked = false;
            comboBox5.SelectedIndex = -1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            rB7.Checked = false;
            rB8.Checked = false;
            rB9.Checked = false;
            comboBox6.SelectedIndex = -1;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem.ToString() == "Другое")
            {
                textBox2.Visible = true;
            }
            if (comboBox3.SelectedItem.ToString() != "Другое")
            {
                textBox2.Visible = false;
            }
        }
    }
}
