using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Drawing.Imaging;
using System.Windows.Controls;
using System.Windows.Data;
//using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;

namespace WpfTestCard
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
            //new teacher().Show();
        }
        /// <summary>
        /// 工作人员
        /// </summary>
        /// <returns></returns>
        public bool CreateWorddocument()
        {
            try
            {
                object defvalue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordapp = new Microsoft.Office.Interop.Word.Application();
                
                //Document doc = new Document();
                Document doc = wordapp.Documents.Add(ref defvalue, ref defvalue, ref defvalue, ref defvalue);
                wordapp.Visible = true;
                doc.PageSetup.PageWidth = 235f;
                doc.PageSetup.PageHeight = 350f;
                //doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                doc.PageSetup.LeftMargin = 15f;
                doc.PageSetup.RightMargin = 15f;
                doc.PageSetup.TopMargin = 20f;
                doc.PageSetup.BottomMargin = 20f;

                
                doc.Paragraphs.Last.Range.Font.Size = 18f;
                doc.Paragraphs.Last.Range.Font.Name = "黑体";
                doc.Paragraphs.Last.Range.Font.Bold = 1;
                doc.Paragraphs.Last.Range.Text = "全国硕士研究生招生考试\r\n";

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                doc.Paragraphs.Last.Range.Text ="2017年西华大学考点\r\n";


                doc.Paragraphs.Last.Range.Font.Size = 26f;
                doc.Paragraphs.Last.Range.Font.Bold = 2;
                doc.Paragraphs.Last.Range.Text = "工作人员证\r\n" ;

                wordapp.Selection.MoveDown(WdUnits.wdLine, 3,WdMovementType.wdExtend);//选中向下3行
                wordapp.Selection.ParagraphFormat.DisableLineHeightGrid=-1 ;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordapp.Selection.ParagraphFormat.LineUnitAfter = 0;
                wordapp.Selection.ParagraphFormat.LineUnitBefore = 0;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);// 移动光标到末尾


                /*插入图片*/
                doc.InlineShapes.AddPicture("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg", WdPictureLinkType.wdLinkNone,true,doc.Paragraphs.Last.Range);
                System.Drawing.Image pic = System.Drawing.Image.FromFile("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg");
                float dd = pic.Width / (float) pic.Height;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                doc.InlineShapes[1].Height = 135f;
                doc.InlineShapes[1].Width = 135*dd;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Text = "姓名：吴俊川";
                wordapp.Selection.ParagraphFormat.LeftIndent = 30f;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);
                wordapp.Selection.TypeParagraph();
                

                doc.Paragraphs.Last.Range.Text = "单位：研究生部";
                wordapp.Selection.ParagraphFormat.LeftIndent = 30f;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);
                wordapp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);

                /*
                 * 第二页
                 * 
                 */

                doc.Paragraphs.Last.Range.Font.Size = 10f;
                doc.Paragraphs.Last.Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                doc.Paragraphs.Last.Range.Text = "我是第一行，居中对齐\r\n";
                

                doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                doc.Paragraphs.Last.Range.Text = "第二，左";
                doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                doc.Paragraphs.Last.Range.Text += "第二，右\r\n";

                doc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute;
                doc.Paragraphs.Last.Range.Text = "文档结尾,分布对齐\r\n";
                doc.Paragraphs.CloseUp();
                //insertBreakNextPage();
                doc.SaveAsQuickStyleSet("1125");
                doc.Close();
                wordapp.Quit();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message+e.StackTrace, "文档生成信息");
                return false;
            }
            
            //
            //doc.PageSetup.PageHeight=16;
            //doc.PageSetup.PageWidth = 9;
            ////doc.PageSetup.PaperSize=4;
            //Paragraphs title = doc.Paragraphs;
            ////title.Alignment = ;
            //doc.Paragraphs.Add(new object());
        }
        public void CreateFlowdocument()
        {

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            new System.Threading.Tasks.Task(new Action(() => 
            {
                CreateWorddocument();
            })).Start();
            
        }
    }
}
