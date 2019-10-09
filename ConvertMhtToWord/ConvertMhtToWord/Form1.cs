using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word.Document;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Collections.Generic;
using System.Linq;

namespace ConvertMhtToWord
{
    public partial class Form1 : Form
    {
        private string _imgDirectory = String.Empty;
        public static string strHtmlHead = @"<html xmlns=""http://www.w3.org/1999/xhtml""><head>";
        public static string strHtmlEnd = @"</table></body></html>";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            //file.Filter = @"*.mht";
            file.Filter = @"mht文件|*.mht|所有文件|*.*";
            //file.DefaultExt = ".mht";
            file.ShowDialog();
            this.textBox1.Text = file.FileName;
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists(textBox1.Text))
            {
                DoConvert(textBox1.Text, true);
            }
            else
            {
                MessageBox.Show(@"文件不存在");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (File.Exists(textBox1.Text))
            {
                DoConvert(textBox1.Text, false);
            }
            else
            {
                MessageBox.Show(@"文件不存在");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var imgList = new List<string>();
            using (var srReadFile = new StreamReader(_imgDirectory + "\\" + "ImgDictionary.txt"))
            {
                while (!srReadFile.EndOfStream)
                {
                    string strReadLine = srReadFile.ReadLine();
                    imgList.Add(strReadLine);
                }
            }

            //FileStream fsDic = new FileStream(_imgDirectory + "\\" + "ImgDictionary.txt", FileMode.Create);
            Application wordApp = new Application();
            Document doc = wordApp.Documents.Add();
            var parser = new HtmlParser();

            FileStream aFile = new FileStream(_imgDirectory + "\\" + "temp.html", FileMode.Open);
            var document = parser.ParseDocument(aFile);
            
            var tableChildren = document.QuerySelectorAll("table > tbody > tr > td");
            int i = 0;
            foreach (var child in tableChildren)
            {
                //if (i > 300)
                //    break;
                if (child.GetElementCount() == 0)
                {
                    WirteWord(doc, wordApp, child.TextContent);
                }
                else
                {
                    var time = child.QuerySelector("div");
                    if (time != null)
                        WirteWord(doc, wordApp, time.TextContent);

                    var font = child.QuerySelector("div font");
                    if (font != null)
                        WirteWord(doc, wordApp, font.TextContent);

                    var img = child.QuerySelector("img");
                    if (img != null)
                    {
                        var data = img.GetAttribute("src");
                        string sheetData = Regex.Match(data, @"\{(.*)\}", RegexOptions.Singleline).Groups[1].Value;

                        var imgUrl = imgList.Where(a => a.Contains(sheetData)).ToList();
                        if (imgUrl.Count > 0)
                        {
                            WirteWord(doc, wordApp, imgUrl[0], true);
                        }
                    }
                }
                i++;
            }
            var strSrcFilePath = textBox1.Text;
            var fileDirectory = strSrcFilePath.Substring(0, strSrcFilePath.IndexOf("."));
            var fileName = fileDirectory + ".docx";

            foreach (Paragraph para in doc.Paragraphs)
            {
                if (para.Range.Text.StartsWith("日期"))
                {
                    //Object styleHeading = WdBuiltinStyle.wdStyleHeading1; //四级标题
                    para.Range.set_Style("标题 1");
                }
            }

            if (!File.Exists(fileName))
            {
                FileInfo fi = new FileInfo(fileName);
                var di = fi.Directory;
                if (!di.Exists)
                    di.Create();
                doc.SaveAs2(fileName);
            }

            aFile.Close();
            doc.Close();
            wordApp.Quit();
        }

        private void DoConvert(string strSrcFilePath, bool isHtml)
        {
            //mht文件
            FileStream fsSrc = new FileStream(strSrcFilePath, FileMode.Open);
            StreamReader rsSrc = new StreamReader(fsSrc);
            StringBuilder sbSrc = new StringBuilder();
            FileStream fsDic = null;
            StreamWriter swDic = null;
            FileStream fsHtml = null;
            StreamWriter swHtml = null;

            _imgDirectory = strSrcFilePath.Substring(0, strSrcFilePath.IndexOf("."));
            if (!Directory.Exists(_imgDirectory))
                Directory.CreateDirectory(_imgDirectory);
            string strLine;
            string strSuffix = "";
            string strContent;
            string htmlContent;
            string strImgFileName = "";
            bool blBegin = false;           //表示到一个附件开头的标志位
            bool blEnd = false;             //表示到一个附件结尾的标志位

            bool htmlBegin = false;
            bool htmlEnd = false;

            if (isHtml)
            {
                //html文件
                 fsHtml = new FileStream(_imgDirectory + "\\" + "temp.html", FileMode.Create);
                 swHtml = new StreamWriter(fsHtml);
            }
            else
            {
                //txt文件
                 fsDic = new FileStream(_imgDirectory + "\\" + "ImgDictionary.txt", FileMode.Create);
                 swDic = new StreamWriter(fsDic);
            }
            
            while (!rsSrc.EndOfStream)
            {
                strLine = rsSrc.ReadLine().TrimEnd();
                if (isHtml)
                {
                    if (strLine.Contains(strHtmlHead))
                    {
                        htmlBegin = true;
                    }
                    else if (strLine.Contains(strHtmlEnd))
                    {
                        htmlEnd = true;
                    }
                    if (htmlBegin)
                    {
                        sbSrc.Append(strLine);
                    }
                    if (htmlEnd)
                    {
                        htmlContent = sbSrc.ToString();
                        swHtml.WriteLine(htmlContent);
                        break;
                    }
                }
                else
                {
                    //第1步操作,附件部分读取成相应的图片,并将图片名称和后缀信息保存成字典文件
                    if (strLine == "")
                    {
                        if (blBegin == true && blEnd == true)
                        {
                            blEnd = false;
                        }
                        else if (blBegin == true && blEnd == false)
                        {
                            blBegin = false;
                            blEnd = true;
                            strContent = sbSrc.ToString();
                            sbSrc.Remove(0, sbSrc.Length);
                            WriteToImage(strImgFileName, strContent, strSuffix);    //保存成图片文件
                            swDic.WriteLine(_imgDirectory + "\\" + strImgFileName + "." + strSuffix);  //写入到字典文件,用户读取正文时生成链接
                        }
                    }
                    else if (strLine.Contains("Content-Location:"))
                    {
                        blBegin = true;
                        strImgFileName = strLine.Substring(18, 36);
                    }
                    else if (strLine.Contains("Content-Type:image/"))
                    {
                        strSuffix = strLine.Replace("Content-Type:image/", "");
                    }
                    else if (blBegin == true)
                    {
                        sbSrc.Append(strLine);
                    }
                }

            }
            if (isHtml)
            {
                swHtml.Close();
                fsHtml.Close();
            }
            else
            {
                swDic.Close();
                fsDic.Close();
            }
            rsSrc.Close();
            fsSrc.Close();
        }

        //保存每个图片到对应的文件
        private void WriteToImage(string strFileName, string strContent, string strSuffix)
        {
            byte[] byteContent = Convert.FromBase64String(strContent);
            FileStream fs = new FileStream(_imgDirectory + "\\" + strFileName + "." + strSuffix, FileMode.Create);
            fs.Write(byteContent, 0, byteContent.Length);
            fs.Close();
        }

        void WirteWord(Document wordDoc, Application wordApp, string text, bool isImg = false)
        {
            wordApp.Selection.EndKey();
            if (isImg)
            {
                Object range = wordDoc.Paragraphs.Last.Range;
                //定义该插入的图片是否为外部链接
                Object linkToFile = false;               //默认,这里貌似设置为bool类型更清晰一些
                //定义要插入的图片是否随Word文档一起保存
                Object saveWithDocument = true;              //默认
                //使用InlineShapes.AddPicture方法(【即“嵌入型”】)插入图片
                wordDoc.Paragraphs.Last.Range.Text += "\n";
                wordDoc.InlineShapes.AddPicture(text, ref linkToFile, ref saveWithDocument, ref range);
            }
            else
            {
                wordDoc.Paragraphs.Last.Range.Text += text;
            }
        }
    }
}
