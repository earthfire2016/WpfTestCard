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
using MySql.Data.MySqlClient;
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
            mysqlCon.ConnectionString = ConnString;
            mysqlCon.Open();
            //new teacher().Show();
        }
        private string ConnString = "server=localhost;database=testcard;userid=root;password=123456";
        private MySqlConnection mysqlCon = new MySqlConnection();

        private MySqlDataReader GetDataReader()
        {
            MySqlDataReader reader;
            reader= MySqlHelper.ExecuteReader(mysqlCon,"select * from workerinfo");
            while (reader.Read())
            {
                string name, room, unit, id;
                object[] values=new object[4];
                reader.GetValues(values);
                name = values[0].ToString();
                room = values[1].ToString();
                unit = values[2].ToString();
                id = values[3].ToString();
                this.Dispatcher.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show("name: "+name+" room: "+room+" unit:"+unit+" id:"+id);
                }));
                //MessageBox.Show(values[0].ToString() + values[1].ToString() + values[2].ToString() + values[3].ToString());
            }
            return reader;
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

                wordapp.Selection.ParagraphFormat.LeftIndent = 0f;
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

                doc.Paragraphs.Last.Range.Text = "姓名：";
                doc.Paragraphs.Last.Range.InsertAfter("吴俊川");
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
                wordapp.Selection.ParagraphFormat.LeftIndent = 0f;
                doc.Paragraphs.Last.Range.Font.Size = 18f;
                doc.Paragraphs.Last.Range.Font.Name = "黑体";
                doc.Paragraphs.Last.Range.Font.Bold = 1;
                doc.Paragraphs.Last.Range.Text = "全国硕士研究生招生考试\r\n";

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                doc.Paragraphs.Last.Range.Text = "2017年西华大学考点\r\n";


                doc.Paragraphs.Last.Range.Font.Size = 26f;
                doc.Paragraphs.Last.Range.Font.Bold = 2;
                doc.Paragraphs.Last.Range.Text = "工作人员证\r\n";

                wordapp.Selection.MoveDown(WdUnits.wdLine, 3, WdMovementType.wdExtend);//选中向下3行
                wordapp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordapp.Selection.ParagraphFormat.LineUnitAfter = 0;
                wordapp.Selection.ParagraphFormat.LineUnitBefore = 0;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);// 移动光标到末尾


                /*插入图片*/
                doc.InlineShapes.AddPicture("C:\\Users\\Administrator\\Desktop\\文档\\tt.jpg", WdPictureLinkType.wdLinkNone, true, doc.Paragraphs.Last.Range);
                System.Drawing.Image p = System.Drawing.Image.FromFile("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg");
                float d = p.Width / (float)p.Height;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                doc.InlineShapes[2].Height = 135f;
                doc.InlineShapes[2].Width = 135 * d;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Text = "姓名：";
                doc.Paragraphs.Last.Range.InsertAfter("吴俊川");
                wordapp.Selection.ParagraphFormat.LeftIndent = 30f;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);
                wordapp.Selection.TypeParagraph();


                doc.Paragraphs.Last.Range.Text = "单位：研究生部";
                wordapp.Selection.ParagraphFormat.LeftIndent = 30f;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);
                wordapp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);


                doc.Paragraphs.CloseUp();
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
        /// <summary>
        /// 12.2 修改
        /// </summary>
        /// <returns></returns>
        public bool ProUse()
        {
            try
            {
                object defvalue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordapp = new Microsoft.Office.Interop.Word.Application();

                //Document doc = new Document();
                Document doc = wordapp.Documents.Add(ref defvalue, ref defvalue, ref defvalue, ref defvalue);
                wordapp.Visible = true;
                wordapp.Selection.WholeStory();
                wordapp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
                doc.PageSetup.PageWidth = 260f;
                doc.PageSetup.PageHeight = 350f;
                //doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                doc.PageSetup.LeftMargin = 15f;
                doc.PageSetup.RightMargin = 15f;
                doc.PageSetup.TopMargin = 26f;
                doc.PageSetup.BottomMargin = 15f;
                wordapp.Selection.ParagraphFormat.TabStops.Add(207, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);



                for (int i = 1; i < 10; i++)
                {
                    wordapp.Selection.ParagraphFormat.LeftIndent = 0f;
                    doc.Paragraphs.Last.Range.Font.Size = 18f;
                    doc.Paragraphs.Last.Range.Font.Name = "黑体";
                    doc.Paragraphs.Last.Range.Font.Bold = 1;
                    doc.Paragraphs.Last.Range.Text = "全国硕士研究生招生考试\r\n";

                    doc.Paragraphs.Last.Range.Font.Size = 16f;
                    doc.Paragraphs.Last.Range.Font.Bold = 0;
                    doc.Paragraphs.Last.Range.Text = "2017年西华大学考点\r\n";


                    doc.Paragraphs.Last.Range.Font.Size = 26f;
                    doc.Paragraphs.Last.Range.Font.Bold = 2;
                    doc.Paragraphs.Last.Range.Text = "监考员证\r\n";

                    wordapp.Selection.MoveDown(WdUnits.wdLine, 3, WdMovementType.wdExtend);//选中向下3行
                    wordapp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
                    wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    wordapp.Selection.ParagraphFormat.LineUnitAfter = 0;
                    wordapp.Selection.ParagraphFormat.LineUnitBefore = 0;
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);// 移动光标到末尾


                    /*插入图片*/
                    doc.InlineShapes.AddPicture("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg", WdPictureLinkType.wdLinkNone, true, doc.Paragraphs.Last.Range);
                    System.Drawing.Image pic = System.Drawing.Image.FromFile("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg");
                    float dd = pic.Width / (float)pic.Height;
                    wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    doc.InlineShapes[i].Height = 135f;
                    doc.InlineShapes[i].Width = 135 * dd;
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                    wordapp.Selection.TypeParagraph();

                    doc.Paragraphs.Last.Range.Font.Size = 16f;
                    doc.Paragraphs.Last.Range.Font.Bold = 0;
                    wordapp.Selection.TypeParagraph();

                    doc.Paragraphs.Last.Range.Text = "姓名：";

                    wordapp.Selection.ParagraphFormat.LeftIndent = 24f;
                    wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                    // 设置 name 字体 
                    doc.Paragraphs.Last.Range.InsertAfter("吴俊川");
                    wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                    wordapp.Selection.Font.Size = 18f;
                    wordapp.Selection.Font.Name = "华文新魏";

                    //doc.Paragraphs.TabStops.Add(180, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
                    //doc.Paragraphs.Last.Range.InsertAlignmentTab((int)WdTabAlignment.wdAlignTabRight,0);
                    doc.Paragraphs.Last.Range.InsertAfter("\t");
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);



                    doc.Paragraphs.Last.Range.InsertAfter("考场号：01");
                    //wordapp.Selection.ParagraphFormat.TabStops.Add(207, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
                    wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                    wordapp.Selection.Font.Size = 14f;
                    wordapp.Selection.Font.Name = "宋体";
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);


                    wordapp.Selection.TypeParagraph();


                    doc.Paragraphs.Last.Range.Text = "单位：";
                    wordapp.Selection.ParagraphFormat.LeftIndent = 24f;
                    doc.Paragraphs.Last.Range.Font.Size = 16f;
                    doc.Paragraphs.Last.Range.Font.Name = "黑体";
                    wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                    doc.Paragraphs.Last.Range.InsertAfter("1 考务室");
                    wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                    wordapp.Selection.Font.Size = 14f;
                    wordapp.Selection.Font.Name = "宋体";
                    wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                    wordapp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                }
                

                /*
                 * 第二页
                 * 
                 */
                doc.Paragraphs.CloseUp();
                doc.SaveAsQuickStyleSet("1125");
                doc.Close();
                wordapp.Quit();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + e.StackTrace, "文档生成信息");
                return false;
            }
        }
        public bool CreateTripdoc()
        {
            try
            {
                object defvalue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordapp = new Microsoft.Office.Interop.Word.Application();

                //Document doc = new Document();
                Document doc = wordapp.Documents.Add(ref defvalue, ref defvalue, ref defvalue, ref defvalue);
                wordapp.Visible = true;
                wordapp.Selection.WholeStory();
                wordapp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
                doc.PageSetup.PageWidth = 260f;
                doc.PageSetup.PageHeight = 350f;
                //doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                doc.PageSetup.LeftMargin = 15f;
                doc.PageSetup.RightMargin = 15f;
                doc.PageSetup.TopMargin = 26f;
                doc.PageSetup.BottomMargin = 15f;

                wordapp.Selection.ParagraphFormat.LeftIndent = 0f;
                doc.Paragraphs.Last.Range.Font.Size = 18f;
                doc.Paragraphs.Last.Range.Font.Name = "黑体";
                doc.Paragraphs.Last.Range.Font.Bold = 1;
                doc.Paragraphs.Last.Range.Text = "全国硕士研究生招生考试\r\n";

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                doc.Paragraphs.Last.Range.Text = "2017年西华大学考点\r\n";


                doc.Paragraphs.Last.Range.Font.Size = 26f;
                doc.Paragraphs.Last.Range.Font.Bold = 2;
                doc.Paragraphs.Last.Range.Text = "监考员证\r\n";

                wordapp.Selection.MoveDown(WdUnits.wdLine, 3, WdMovementType.wdExtend);//选中向下3行
                wordapp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordapp.Selection.ParagraphFormat.LineUnitAfter = 0;
                wordapp.Selection.ParagraphFormat.LineUnitBefore = 0;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);// 移动光标到末尾


                /*插入图片*/
                doc.InlineShapes.AddPicture("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg", WdPictureLinkType.wdLinkNone, true, doc.Paragraphs.Last.Range);
                System.Drawing.Image pic = System.Drawing.Image.FromFile("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg");
                float dd = pic.Width / (float)pic.Height;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                doc.InlineShapes[1].Height = 135f;
                doc.InlineShapes[1].Width = 135 * dd;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Text = "姓名：";
                
                wordapp.Selection.ParagraphFormat.LeftIndent = 24f;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                // 设置 name 字体 
                doc.Paragraphs.Last.Range.InsertAfter("吴俊川");
                wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                wordapp.Selection.Font.Size = 18f;
                wordapp.Selection.Font.Name = "华文新魏";

                //doc.Paragraphs.TabStops.Add(180, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
                //doc.Paragraphs.Last.Range.InsertAlignmentTab((int)WdTabAlignment.wdAlignTabRight,0);
                doc.Paragraphs.Last.Range.InsertAfter("\t");
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);
                


                doc.Paragraphs.Last.Range.InsertAfter("考场号：01");
                wordapp.Selection.ParagraphFormat.TabStops.Add(207,WdTabAlignment.wdAlignTabRight,WdTabLeader.wdTabLeaderSpaces);
                wordapp.Selection.MoveEnd(WdUnits.wdStory,10);
                wordapp.Selection.Font.Size = 14f;
                wordapp.Selection.Font.Name = "宋体";
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);


                wordapp.Selection.TypeParagraph();


                doc.Paragraphs.Last.Range.Text = "单位：";
                wordapp.Selection.ParagraphFormat.LeftIndent = 24f;
                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Name = "黑体";
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                doc.Paragraphs.Last.Range.InsertAfter("1 考务室");
                wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                wordapp.Selection.Font.Size = 14f;
                wordapp.Selection.Font.Name = "宋体";
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                wordapp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);

                /*
                 * 第二页
                 * 
                 */
                wordapp.Selection.ParagraphFormat.LeftIndent = 0f;
                doc.Paragraphs.Last.Range.Font.Size = 18f;
                doc.Paragraphs.Last.Range.Font.Name = "黑体";
                doc.Paragraphs.Last.Range.Font.Bold = 1;
                doc.Paragraphs.Last.Range.Text = "全国硕士研究生招生考试\r\n";

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                doc.Paragraphs.Last.Range.Text = "2017年西华大学考点\r\n";


                doc.Paragraphs.Last.Range.Font.Size = 26f;
                doc.Paragraphs.Last.Range.Font.Bold = 2;
                doc.Paragraphs.Last.Range.Text = "监考员证\r\n";

                wordapp.Selection.MoveDown(WdUnits.wdLine, 3, WdMovementType.wdExtend);//选中向下3行
                wordapp.Selection.ParagraphFormat.DisableLineHeightGrid = -1;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordapp.Selection.ParagraphFormat.LineUnitAfter = 0;
                wordapp.Selection.ParagraphFormat.LineUnitBefore = 0;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);// 移动光标到末尾


                /*插入图片*/
                doc.InlineShapes.AddPicture("C:\\Users\\Administrator\\Desktop\\文档\\tt.jpg", WdPictureLinkType.wdLinkNone, true, doc.Paragraphs.Last.Range);
                System.Drawing.Image pi = System.Drawing.Image.FromFile("C:\\Users\\Administrator\\Desktop\\文档\\te.jpg");
                float d = pi.Width / (float)pi.Height;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                doc.InlineShapes[2].Height = 135f;
                doc.InlineShapes[2].Width = 135 * d;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Bold = 0;
                wordapp.Selection.TypeParagraph();

                doc.Paragraphs.Last.Range.Text = "姓名：";

                wordapp.Selection.ParagraphFormat.LeftIndent = 24f;
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                // 设置 name 字体 
                doc.Paragraphs.Last.Range.InsertAfter("杨明华");
                wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                wordapp.Selection.Font.Size = 18f;
                wordapp.Selection.Font.Name = "华文新魏";

                //doc.Paragraphs.TabStops.Add(180, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
                //doc.Paragraphs.Last.Range.InsertAlignmentTab((int)WdTabAlignment.wdAlignTabRight,0);
                doc.Paragraphs.Last.Range.InsertAfter("\t");
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);



                doc.Paragraphs.Last.Range.InsertAfter("考场号：01");
                wordapp.Selection.ParagraphFormat.TabStops.Add(207, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
                wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                wordapp.Selection.Font.Size = 14f;
                wordapp.Selection.Font.Name = "宋体";
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);


                wordapp.Selection.TypeParagraph();


                doc.Paragraphs.Last.Range.Text = "单位：";
                wordapp.Selection.ParagraphFormat.LeftIndent = 24f;
                doc.Paragraphs.Last.Range.Font.Size = 16f;
                doc.Paragraphs.Last.Range.Font.Name = "黑体";
                wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                doc.Paragraphs.Last.Range.InsertAfter("1 考务室");
                wordapp.Selection.MoveEnd(WdUnits.wdStory, 10);
                wordapp.Selection.Font.Size = 14f;
                wordapp.Selection.Font.Name = "宋体";
                wordapp.Selection.EndKey(WdUnits.wdStory, ref defvalue);

                wordapp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);


                doc.Paragraphs.CloseUp();
                doc.SaveAsQuickStyleSet("1125");
                doc.Close();
                wordapp.Quit();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + e.StackTrace, "文档生成信息");
                return false;
            }
        }
        public bool UseTempate()
        {
            try
            {
                object defvalue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordapp = new Microsoft.Office.Interop.Word.Application();

                //Document doc = new Document();
                object templatename = "C:\\Users\\Administrator\\Desktop\\文档\\范围最新模板.dotx";
                Document doc = wordapp.Documents.Add(ref templatename, ref defvalue, ref defvalue, ref defvalue);
                wordapp.Visible = true;


                
                //doc.Bookmarks["name"].Range.Text = "黎明";
                wordapp.Selection.GoTo(WdGoToItem.wdGoToBookmark, defvalue, defvalue, "name");
                wordapp.Selection.TypeText("明天见");
                //doc.Bookmarks["name"].Range.Font.Name = "华文新魏";
                //doc.Bookmarks["name"].Range.Font.Size = 18f;

                doc.Bookmarks["num"].Range.Text = "03";


                doc.Bookmarks["ph"].Range.InlineShapes.AddPicture("C:\\Users\\Administrator\\Desktop\\文档\\tt.jpg", WdPictureLinkType.wdLinkNone, true);
                System.Drawing.Image pi = System.Drawing.Image.FromFile("C:\\Users\\Administrator\\Desktop\\文档\\tt.jpg");
                float d = pi.Width / (float)pi.Height;
                //wordapp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                doc.InlineShapes[1].Height = 135f;
                doc.InlineShapes[1].Width = 135 * d;


                doc.Bookmarks["unit"].Range.Text = "7考务室";
                //doc.Bookmarks[4].Range.Font.Name = "宋体";
                //doc.Bookmarks[4].Range.Font.Size = 14f;

                doc.Paragraphs.CloseUp();



                doc.SaveAsQuickStyleSet("1125");
                doc.Close();
                wordapp.Quit();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.StackTrace, "失败");
                return false;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    //CreateWorddocument();
                    //new teacher().Show();
                    //UseTempate();
                    GetDataReader();
                    //ProUse();
                }
                catch (Exception es)
                {
                    MessageBox.Show(es.ToString());
                    
                }
                
            }));
            //new System.Threading.Tasks.Task(new Action(() => 
            //{
            //    //CreateWorddocument();
            //    // CreateTripdoc();
                
                
            //})).Start();
            
        }
    }
}
