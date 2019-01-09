using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Windows.Forms.Design;
using System.Threading;

namespace ToDocx
{
    public partial class Form1 : Form
    {
        string[] fs;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        public MSWord.Application wordApp;
        public MSWord.Document wordDoc;
        private void button1_Click(object sender, EventArgs e)
        {
            string _path;
            fldDlg.ShowDialog();
            _path = fldDlg.SelectedPath;
            if (!Directory.Exists(_path))
                return;
            textBox1.Text = _path;
            //fs = Directory.GetFiles(_path, "*.doc");
            //foreach (string f in fs)
            //{
            //    if (f.Split('.')[1].Length == 3)
            //        textBox2.AppendText(f + "\r\n");
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string _path;
            //fldDlg.ShowDialog();
            _path = fldDlg.SelectedPath;
            wordApp = new MSWord.Application();
            Thread T = new Thread(() =>
                {
                    if (!Directory.Exists(_path))
                    {
                        MessageBox.Show("你的路径是非法路径，请输入文件夹路径！","错误信息！");
                        return;
                    }

                    string[] file = Directory.GetFiles(_path, "*.doc", SearchOption.AllDirectories);
                    //foreach (string f in file)
                    //{
                    //    if(f.Split('.')[1].Length == 3)
                    //        textBox2.AppendText(f + "\r\n");
                    //}
                    if(file.Length == 0)
                    {
                        MessageBox.Show("当前文件夹不存在doc文档，请核对路径是否正确！");
                        return;
                    }

                    for (int i = 0; i < file.Length; i++)
                    {
                        try
                        {
                            wordDoc = wordApp.Documents.Open(file[i]);
                            wordDoc.SaveAs(file[i].Substring(0, file[i].LastIndexOf(".")) + ".docx",
                                MSWord.WdSaveFormat.wdFormatDocumentDefault);
                            wordDoc.Close(false);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(file[i]);
                            continue;
                        }
                    }
                    wordApp.Quit();
                    MessageBox.Show("转换完成！");
                }
            );
            T.Start();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string _path;
            _path = fldDlg.SelectedPath;
            wordApp = new MSWord.Application();
            Thread T = new Thread(() =>
                {
                    if (!Directory.Exists(_path))
                    {
                        MessageBox.Show("你的路径是非法路径，请输入文件夹路径！", "错误信息！");
                        return;
                    }

                    string[] file = Directory.GetFiles(_path, "*.docx", SearchOption.AllDirectories);
                    if (file.Length == 0)
                    {
                        MessageBox.Show("当前文件夹不存在docx文档，请核对路径是否正确！");
                        return;
                    }

                    for (int i = 0; i < file.Length; i++)
                    {
                        try
                        {
                            wordDoc = wordApp.Documents.Open(file[i]);
                            wordDoc.SaveAs(file[i].Substring(0, file[i].LastIndexOf(".")) + ".doc",
                                MSWord.WdSaveFormat.wdFormatDocumentDefault);
                            wordDoc.Close(false);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(file[i]);
                            continue;
                        }
                    }
                    wordApp.Quit();
                    MessageBox.Show("转换完成！");
                }
            );
            T.Start();
        }
    }
}
