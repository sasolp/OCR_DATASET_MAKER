using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using openCV;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text.RegularExpressions;
namespace TXT2Word_Lable
{
    public partial class Form1 : Form
    {
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct BoundRect
        {
            public int x;
            public int y;
            public int width;
            public int height;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct PolyPoint
        {
            public int x;
            public int y;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct EngineBoundPolygon
        {
            public IntPtr pnts;
            public int pntCount;
            public double area;
        }
        struct BoundPolygon
        {
            public PolyPoint[] pnts;
            public double area;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct EnginePicRegion
        {
            public BoundRect boundRect;
            public EngineBoundPolygon boundPoly;
        }
        struct PicRegion
        {
            public BoundRect boundRect;
            public BoundPolygon boundPoly;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct EngineTextRegion
        {
            public BoundRect boundRect;
            public EngineBoundPolygon boundPoly;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string ocrText;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string fontName;
            public int fontSize;
        }
        struct TextRegion
        {
            public BoundRect boundRect;
            public BoundPolygon boundPoly;
            public string ocrText;
            public string fontName;
            public int fontSize;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct TableCell
        {
            public BoundRect boundRect;
            public int row;
            public int column;
            public int rowSpan;
            public int columnSpan;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string ocrText;            
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct EngineTableRegion
        {
            public BoundRect boundRect;
            public int rowsCount;
            public int columnsCount;
            public int cellsCount;
            public IntPtr Cells;
        }
        struct TableRegion
        {
            public BoundRect boundRect;
            public int rowsCount;
            public int columnsCount;
            public TableCell[] cells;
        }
        /*[StructLayout(LayoutKind.Sequential), Serializable]
        struct BoundRect
        {
            public int x;
            public int y;
            public int width;
            public int height;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct PolyPoint
        {
            public int x;
            public int y;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct BoundPolygon
        {
           // [MarshalAs(UnmanagedType.ByValArray)]
            public IntPtr pnts;
            public double area;
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct TextRegion
        {
            public BoundRect boundRect;
            public BoundPolygon boundPoly;
           // [MarshalAs(UnmanagedType.BStr)]
            public IntPtr ocrText;
            //[MarshalAs(UnmanagedType.BStr)]
            public IntPtr fontName;
            public int fontSize;
        }*/
        [DllImport("ImgProc.dll", EntryPoint = "FindBaseLine", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.Cdecl)]
        extern static void FindBaseLine(ref IplImage inputImg, IntPtr finalBaseLine, int font_size);

        [DllImport("ImgProc.dll", EntryPoint = "MapToOriginalSize", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        extern static IntPtr MapToOriginalSize(ref IplImage inputImg);
        [DllImport("ImgProc.dll", EntryPoint = "ReleaseMap", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        extern static void ReleaseMap(IntPtr inputImg);
        [DllImport("ImgProc.dll", EntryPoint = "GetOCRText", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        extern static bool GetOCRText(string strPicPath, [MarshalAs(UnmanagedType.LPWStr)]ref string strOcrText);
        [DllImport("ImgProc.dll", EntryPoint = "GetOCRText2", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        extern static int GetOCRText2(string strPicPath, string strOutPath, int fileCounter, [MarshalAs(UnmanagedType.LPWStr)]ref string strOcrText);
        [DllImport("ImgProc.dll", EntryPoint = "SaveImage", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        extern static void SaveImage(string strPicPath, ref IplImage inputImg);
        [DllImport("ImgProc.dll", EntryPoint = "AnalyzeLayout", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        unsafe extern static void AnalyzeLayout(IntPtr pImageData, int WIDTH, int HEIGHT, int WIDTHSTEP, int channels, out IntPtr arrTexts, out int textsCount, out IntPtr arrPics, out int picsCount, out IntPtr arrTables, out int tablesCount);
        [DllImport("ImgProc.dll", EntryPoint = "ProccessLayout", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        unsafe extern static void ProccessLayout(IntPtr pImageData, int WIDTH, int HEIGHT, int WIDTHSTEP, int channels, ref IntPtr arrTexts, ref int textsCount, ref IntPtr arrPics, ref int picsCount, ref IntPtr arrTables, ref int tablesCount);
        [DllImport("ImgProc.dll", EntryPoint = "FeatureSVMTrain", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        unsafe extern static int FeatureSVMTrain(string inputImagePath, string trainFilePath);
        [DllImport("ImgProc.dll", EntryPoint = "LoadSVMModels", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        unsafe extern static  int LoadSVMModels(ref IntPtr svm_classifier);
        [DllImport("ImgProc.dll", EntryPoint = "GetOCR", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        unsafe extern static int GetOCR(ref IntPtr svm_classifier, string inPath, [MarshalAs(UnmanagedType.LPWStr)]ref string strOcrText);
        public Form1()
        {
            InitializeComponent();
        }


        const int CF_METAFILE = 14;

        IntPtr intptr;
        bool bOneChannel;
        string inputFile;
        string outputFile;
        string outputFolder;
        const int NOISE_GAUSSIAN = 0;
        const int NOISE_UNIFORM = 1;
        const int NOISE_SALT_PEPPER = 2;
        double param1;
        double param2;
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.DefaultExt = "txt";
            of.Filter = "Text files (*.txt)|*.txt";
            of.Multiselect = true;
            if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                for (int i = 0; i < of.FileNames.Length; i++)
                {

                    inputFile = txtInputFile.Text += of.FileNames[i] + ";";
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            CreateHashArray();
            cmbNoise.SelectedIndex = 2;
            chkOneChannel.Checked = true;
            bOneChannel = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            sf.Filter = "Word Files (*.docx)|*.docx";
            sf.CheckFileExists = false;
            if (sf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputFile = txtOutputFile.Text = sf.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }


        private static string[] GetCode2(string WORD1)
        {
            
            //StreamWriter sww = new StreamWriter("report.txt");
            bool bMustBeReverse = false;
            Regex regex = new Regex("[0-9]+");
            Match match = regex.Match(WORD1);
            while (match != null && match.Length != 0)
            {
                char[] tmpWord = match.Value.ToArray();
                Array.Reverse(tmpWord);
                WORD1 = WORD1.Replace(match.Value, new string(tmpWord));
                match = match.NextMatch();
            }

            char[] word = WORD1.ToCharArray();

            string[] wordlab = new string[word.Length];
            int[] forward = new int[word.Length];
            int[] backward = new int[word.Length];
            int[] nimcheck = new int[word.Length];





            #region convert to latin
            for (int i = 0; i < wordlab.Length; i++)
            {

                switch (word[i])
                {
                    case 'ض':
                        wordlab[i] = "zad";
                        break;
                    case 'ص':
                        wordlab[i] = "sad";
                        break;
                    case 'ث':
                        wordlab[i] = "se";
                        break;
                    case 'ق':
                        wordlab[i] = "qaf";
                        break;
                    case 'ف':
                        wordlab[i] = "f";
                        break;
                    case 'غ':
                        wordlab[i] = "q";
                        break;
                    case 'ع':
                        wordlab[i] = "gs";
                        break;
                    case 'ه':
                        wordlab[i] = "h";
                        break;
                    case 'خ':
                        wordlab[i] = "x";
                        break;
                    case 'ح':
                        wordlab[i] = "he";
                        break;
                    case 'ج':
                        wordlab[i] = "j";
                        break;
                    case 'چ':
                        wordlab[i] = "ch";
                        break;
                    case 'پ':
                        wordlab[i] = "p";
                        break;
                    case 'ش':
                        wordlab[i] = "sh";
                        break;
                    case 'س':
                        wordlab[i] = "sin";
                        break;
                    case 'ی':
                        wordlab[i] = "y";
                        break;
                    case 'ي':
                        wordlab[i] = "y";
                        break;
                    case 'ب':
                        wordlab[i] = "b";
                        break;
                    case 'ل':
                        wordlab[i] = "l";
                        break;
                    case 'ا':
                        wordlab[i] = "a";
                        break;
                    case 'آ':
                        wordlab[i] = "aa";
                        break;
                    case 'ت':
                        wordlab[i] = "t";
                        break;
                    case 'ن':
                        wordlab[i] = "n";
                        break;
                    case 'م':
                        wordlab[i] = "m";
                        break;
                    case 'ک':
                        wordlab[i] = "k";
                        break;
                    case 'ك':
                        wordlab[i] = "k";
                        break;
                    case 'گ':
                        wordlab[i] = "g";
                        break;
                    case 'ظ':
                        wordlab[i] = "za";
                        break;
                    case 'ط':
                        wordlab[i] = "ta";
                        break;
                    case 'ز':
                        wordlab[i] = "ze";
                        break;
                    case 'ر':
                        wordlab[i] = "r";
                        break;
                    case 'ذ':
                        wordlab[i] = "zal";
                        break;
                    case 'د':
                        wordlab[i] = "d";
                        break;
                    case 'ژ':
                        wordlab[i] = "zh";
                        break;
                    case 'ئ':
                        wordlab[i] = "i";
                        break;
                    case 'و':
                        wordlab[i] = "v";
                        break;
                    case 'ؤ':
                        wordlab[i] = "vh";
                        break;
                    case 'أ':
                        wordlab[i] = "aht";
                        break;
                    case 'إ':
                        wordlab[i] = "ahb";
                        break;
                    case 'ة':
                        wordlab[i] = "hh";
                        break;
                    case 'ۀ':
                        wordlab[i] = "hh";
                        break;
                    case '‌':
                        wordlab[i] = "nimspace";
                        break;
                    case '‏':
                        wordlab[i] = "nimspace";
                        break;
                    case '#':
                        wordlab[i] = "allah";
                        break;
                    case 'ء':
                        wordlab[i] = "hamze";
                        break;
                    case '1':
                        wordlab[i] = "one";
                        break;
                    case '2':
                        wordlab[i] = "two";
                        break;
                    case '3':
                        wordlab[i] = "three";
                        break;
                    case '4':
                        wordlab[i] = "four";
                        break;
                    case '5':
                        wordlab[i] = "five";
                        break;
                    case '6':
                        wordlab[i] = "six";
                        break;
                    case '7':
                        wordlab[i] = "seven";
                        break;
                    case '8':
                        wordlab[i] = "eight";
                        break;
                    case '9':
                        wordlab[i] = "nine";
                        break;
                    case '0':
                        wordlab[i] = "zero";
                        break;
                    case '{':
                        wordlab[i] = "ako";
                        break;
                    case '}':
                        wordlab[i] = "akc";
                        break;
                    case '(':
                        wordlab[i] = "paro";
                        break;
                    case ')':
                        wordlab[i] = "parc";
                        break;
                    case '،':
                        wordlab[i] = "vir";
                        break;
                    case '؛':
                        wordlab[i] = "simi";
                        break;
                    case ';':
                        wordlab[i] = "Engsimi";
                        break;
                    case ':':
                        wordlab[i] = "tdot";
                        break;
                    case '.':
                        wordlab[i] = "dot";
                        break;
                    case '@':
                        wordlab[i] = "atan";
                        break;
                    case '/':
                        wordlab[i] = "slash";
                        break;
                    case '%':
                        wordlab[i] = "darsad";
                        break;
                    case '-':
                        wordlab[i] = "menha";
                        break;
                    case '_':
                        wordlab[i] = "under";
                        break;
                    case '$':
                        wordlab[i] = "la";
                        break;
                    case '?':
                        wordlab[i] = "que";
                        break;
                    case '؟':
                        wordlab[i] = "que";
                        break;
                    case '[':
                        wordlab[i] = "bro";
                        break;
                    case ']':
                        wordlab[i] = "brc";
                        break;
                    case ',':
                        wordlab[i] = "kama";
                        break;
                    case '÷':
                        wordlab[i] = "taqsim";
                        break;
                    case '×':
                        wordlab[i] = "zarb";
                        break;
                    case '=':
                        wordlab[i] = "eq";
                        break;
                    case '*':
                        wordlab[i] = "star";
                        break;
                    case '!':
                        wordlab[i] = "wonder";
                        break;
                    case '+':
                        wordlab[i] = "plus";
                        break;
                    case '»':
                        wordlab[i] = "_gume";
                        break;
                    case '«':
                        wordlab[i] = "gume_";
                        break;
                    case '"':
                        wordlab[i] = "tanvin";
                        break;

                    // case 'ً':
                    //  wordlab[i] = "tanvin";
                    //  break;
                    default:
                        wordlab[i] = "space";
                        // sww.WriteLine(wordlab[i]);
                        break;
                }

            }
            #endregion

            // sww.Close();

            char[] groupALL = new char[] { 'ض', 'ص', 'ث', 'ق', 'ف', 'غ', 'ع', 'ه', 'خ', 'ح', 'ج', 'چ', 'پ', 'ش', 'س', 'ی', 'ب', 'ل', 'ا', 'آ', 'ت', 'ن', 'م', 'ک', 'گ', 'ظ', 'ط', 'ز', 'ر', 'ذ', 'د', 'ژ', 'ئ', 'و', 'ي', 'ك', 'أ', 'ؤ', 'إ', 'ة', '@', '$', 'ً' };

            char[] group1 = new char[] { 'ض', 'ص', 'ث', 'ق', 'ف', 'غ', 'ع', 'ه', 'خ', 'ح', 'ج', 'چ', 'پ', 'ش', 'س', 'ی', 'ب', 'ل', 'ت', 'ن', 'م', 'ک', 'گ', 'ظ', 'ط', 'ئ', 'ي', 'ك', 'ة' };

            char[] group2 = new char[] { 'ا', 'آ', 'ر', 'ز', 'ژ', 'د', 'ذ', 'و', 'أ', 'ؤ', 'إ', '@', '$', 'ً' };

            char[] group3 = new char[] { 'ا', 'آ', 'ر', 'ز', 'ژ', 'د', 'ذ', 'و', 'أ', 'ؤ', 'إ', '@', '$', 'ً' };

            string[] group4 = new string[] { "space", "nimspace", "hamze", "allah", "ako", "akc", "paro", "parc", "dot", "tdot", "vir", "simi", "slash", "darsad", "under", "menha", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "zero", "que", "plus", "wonder", "star", "eq", "zarb", "taqsim", "kama", "bro", "brc" };

            string[] group5 = new string[] { "hamze", "hamze", "allah", "ako", "akc", "paro", "parc", "dot", "tdot", "vir", "simi", "slash", "darsad", "under", "menha", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "zero", "que", "plus", "wonder", "star", "eq", "zarb", "taqsim", "kama", "bro", "brc" };

            Array.Sort(group1);
            Array.Sort(group2);
            Array.Sort(group3);
            Array.Sort(group4);
            Array.Sort(groupALL);
            Array.Sort(group5);


            // Array.BinarySearch(group5,wordlab[i]>=0);


            #region nimcheck


            for (int i = 0; i < wordlab.Length - 1; i++)
            {
                if (Array.BinarySearch(group2, word[i]) >= 0 || Array.BinarySearch(group5, wordlab[i]) >= 0)//wordlab[i] == "hamze" || wordlab[i] == "allah")
                {
                    if (wordlab[i + 1] == "space" || wordlab[i + 1] == "nimspace")
                    {
                        nimcheck[i] = 0;
                    }
                    else
                    {
                        nimcheck[i] = 1;
                    }
                }
                else if (Array.BinarySearch(group5, wordlab[i + 1]) >= 0) //(wordlab[i + 1] == "hamze" || wordlab[i + 1] == "allah")
                {
                    nimcheck[i] = 1;
                }
                else
                {
                    nimcheck[i] = 0;
                }



            }

            #endregion


            #region forward


            for (int i = 0; i < wordlab.Length - 1; i++)
            {
                if (Array.BinarySearch(group1, word[i]) >= 0)
                {
                    //if (wordlab[i + 1] == "space" || wordlab[i + 1] == "nimspace")
                    if (Array.BinarySearch(group4, wordlab[i + 1]) >= 0)
                    {
                        forward[i] = 0;
                    }
                    else
                    {
                        forward[i] = 1;
                    }
                }
                else
                {
                    forward[i] = 0;
                }



            }

            #endregion

            #region backward


            for (int i = wordlab.Length - 1; i > 0; i--)
            {
                if (Array.BinarySearch(groupALL, word[i]) >= 0)
                {
                    if (Array.BinarySearch(group2, word[i - 1]) >= 0)
                    {
                        backward[i] = 0;
                    }
                    else if (Array.BinarySearch(group4, wordlab[i - 1]) >= 0)// if (wordlab[i - 1] == "space" || wordlab[i - 1] == "nimspace")
                    {
                        backward[i] = 0;
                    }
                    else
                    {
                        backward[i] = 1;
                    }
                }
                else
                {
                    backward[i] = 0;
                }



            }







            #endregion



            List<string> mainlab = new List<string>();
            string lastConnectedString = word[0].ToString();
            for (int i = 1; i < wordlab.Length; i++)
            {
                if(
                    //(forward[i] == 1 && backward[i - 1] == 1) ||
                    (forward[i - 1] == 1 && backward[i] == 1)
                    )
                {
                    lastConnectedString += word[i];
                }
                else
                {
                    mainlab.Add(lastConnectedString);
                    lastConnectedString = word[i].ToString();
                }
            }
            mainlab.Add(lastConnectedString);
            if (mainlab.Count == 0)
            {
                mainlab.Add(lastConnectedString);
            }

            //string []arrStr = mainlab.ToArray();
            //string strReturn = "";
            //for (int i = 0; i < arrStr.Length; i++)
            //{
            //    if (arrStr[i].Contains("nimspace")) continue;
            //    strReturn += arrStr[i];
            //}           
            //strReturn.Replace("h__l__l", "allah");
            //strReturn.Replace("a__l", "la");
            //strReturn.Replace("a__l_", "la_");
            //strReturn.Replace("tanvina", "atan");
            //strReturn.Replace("tanvina_", "atan_");
            return mainlab.ToArray();


        }

        private static string[] GetCode(string WORD1)
        {

            //StreamWriter sww = new StreamWriter("report.txt");
            bool bMustBeReverse = false;
            Regex regex = new Regex("[0-9]+");
            Match match = regex.Match(WORD1);
            while (match != null && match.Length != 0)
            {
                char[] tmpWord = match.Value.ToArray();
                Array.Reverse(tmpWord);
                WORD1 = WORD1.Replace(match.Value, new string(tmpWord));
                match = match.NextMatch();
            }

            char[] word = WORD1.ToCharArray();

            string[] wordlab = new string[word.Length];
            int[] forward = new int[word.Length];
            int[] backward = new int[word.Length];
            int[] nimcheck = new int[word.Length];





            #region convert to latin
            for (int i = 0; i < wordlab.Length; i++)
            {

                switch (word[i])
                {
                    case 'ض':
                        wordlab[i] = "zad";
                        break;
                    case 'ص':
                        wordlab[i] = "sad";
                        break;
                    case 'ث':
                        wordlab[i] = "se";
                        break;
                    case 'ق':
                        wordlab[i] = "qaf";
                        break;
                    case 'ف':
                        wordlab[i] = "f";
                        break;
                    case 'غ':
                        wordlab[i] = "q";
                        break;
                    case 'ع':
                        wordlab[i] = "gs";
                        break;
                    case 'ه':
                        wordlab[i] = "h";
                        break;
                    case 'خ':
                        wordlab[i] = "x";
                        break;
                    case 'ح':
                        wordlab[i] = "he";
                        break;
                    case 'ج':
                        wordlab[i] = "j";
                        break;
                    case 'چ':
                        wordlab[i] = "ch";
                        break;
                    case 'پ':
                        wordlab[i] = "p";
                        break;
                    case 'ش':
                        wordlab[i] = "sh";
                        break;
                    case 'س':
                        wordlab[i] = "sin";
                        break;
                    case 'ی':
                        wordlab[i] = "y";
                        break;
                    case 'ي':
                        wordlab[i] = "y";
                        break;
                    case 'ب':
                        wordlab[i] = "b";
                        break;
                    case 'ل':
                        wordlab[i] = "l";
                        break;
                    case 'ا':
                        wordlab[i] = "a";
                        break;
                    case 'آ':
                        wordlab[i] = "aa";
                        break;
                    case 'ت':
                        wordlab[i] = "t";
                        break;
                    case 'ن':
                        wordlab[i] = "n";
                        break;
                    case 'م':
                        wordlab[i] = "m";
                        break;
                    case 'ک':
                        wordlab[i] = "k";
                        break;
                    case 'ك':
                        wordlab[i] = "k";
                        break;
                    case 'گ':
                        wordlab[i] = "g";
                        break;
                    case 'ظ':
                        wordlab[i] = "za";
                        break;
                    case 'ط':
                        wordlab[i] = "ta";
                        break;
                    case 'ز':
                        wordlab[i] = "ze";
                        break;
                    case 'ر':
                        wordlab[i] = "r";
                        break;
                    case 'ذ':
                        wordlab[i] = "zal";
                        break;
                    case 'د':
                        wordlab[i] = "d";
                        break;
                    case 'ژ':
                        wordlab[i] = "zh";
                        break;
                    case 'ئ':
                        wordlab[i] = "i";
                        break;
                    case 'و':
                        wordlab[i] = "v";
                        break;
                    case 'ؤ':
                        wordlab[i] = "vh";
                        break;
                    case 'أ':
                        wordlab[i] = "aht";
                        break;
                    case 'إ':
                        wordlab[i] = "ahb";
                        break;
                    case 'ة':
                        wordlab[i] = "hh";
                        break;
                    case 'ۀ':
                        wordlab[i] = "hh";
                        break;
                    case '‌':
                        wordlab[i] = "nimspace";
                        break;
                    case '‏':
                        wordlab[i] = "nimspace";
                        break;
                    case '#':
                        wordlab[i] = "allah";
                        break;
                    case 'ء':
                        wordlab[i] = "hamze";
                        break;
                    case '1':
                        wordlab[i] = "one";
                        break;
                    case '2':
                        wordlab[i] = "two";
                        break;
                    case '3':
                        wordlab[i] = "three";
                        break;
                    case '4':
                        wordlab[i] = "four";
                        break;
                    case '5':
                        wordlab[i] = "five";
                        break;
                    case '6':
                        wordlab[i] = "six";
                        break;
                    case '7':
                        wordlab[i] = "seven";
                        break;
                    case '8':
                        wordlab[i] = "eight";
                        break;
                    case '9':
                        wordlab[i] = "nine";
                        break;
                    case '0':
                        wordlab[i] = "zero";
                        break;
                    case '{':
                        wordlab[i] = "ako";
                        break;
                    case '}':
                        wordlab[i] = "akc";
                        break;
                    case '(':
                        wordlab[i] = "paro";
                        break;
                    case ')':
                        wordlab[i] = "parc";
                        break;
                    case '،':
                        wordlab[i] = "vir";
                        break;
                    case '؛':
                        wordlab[i] = "simi";
                        break;
                    case ';':
                        wordlab[i] = "Engsimi";
                        break;
                    case ':':
                        wordlab[i] = "tdot";
                        break;
                    case '.':
                        wordlab[i] = "dot";
                        break;
                    case '@':
                        wordlab[i] = "atan";
                        break;
                    case '/':
                        wordlab[i] = "slash";
                        break;
                    case '%':
                        wordlab[i] = "darsad";
                        break;
                    case '-':
                        wordlab[i] = "menha";
                        break;
                    case '_':
                        wordlab[i] = "under";
                        break;
                    case '$':
                        wordlab[i] = "la";
                        break;
                    case '?':
                        wordlab[i] = "que";
                        break;
                    case '؟':
                        wordlab[i] = "que";
                        break;
                    case '[':
                        wordlab[i] = "bro";
                        break;
                    case ']':
                        wordlab[i] = "brc";
                        break;
                    case ',':
                        wordlab[i] = "kama";
                        break;
                    case '÷':
                        wordlab[i] = "taqsim";
                        break;
                    case '×':
                        wordlab[i] = "zarb";
                        break;
                    case '=':
                        wordlab[i] = "eq";
                        break;
                    case '*':
                        wordlab[i] = "star";
                        break;
                    case '!':
                        wordlab[i] = "wonder";
                        break;
                    case '+':
                        wordlab[i] = "plus";
                        break;
                    case '»':
                        wordlab[i] = "_gume";
                        break;
                    case '«':
                        wordlab[i] = "gume_";
                        break;
                    case '"':
                        wordlab[i] = "tanvin";
                        break;

                    // case 'ً':
                    //  wordlab[i] = "tanvin";
                    //  break;
                    default:
                        wordlab[i] = "space";
                        // sww.WriteLine(wordlab[i]);
                        break;
                }

            }
            #endregion

            // sww.Close();

            char[] groupALL = new char[] { 'ض', 'ص', 'ث', 'ق', 'ف', 'غ', 'ع', 'ه', 'خ', 'ح', 'ج', 'چ', 'پ', 'ش', 'س', 'ی', 'ب', 'ل', 'ا', 'آ', 'ت', 'ن', 'م', 'ک', 'گ', 'ظ', 'ط', 'ز', 'ر', 'ذ', 'د', 'ژ', 'ئ', 'و', 'ي', 'ك', 'أ', 'ؤ', 'إ', 'ة', '@', '$', 'ً' };

            char[] group1 = new char[] { 'ض', 'ص', 'ث', 'ق', 'ف', 'غ', 'ع', 'ه', 'خ', 'ح', 'ج', 'چ', 'پ', 'ش', 'س', 'ی', 'ب', 'ل', 'ت', 'ن', 'م', 'ک', 'گ', 'ظ', 'ط', 'ئ', 'ي', 'ك', 'ة' };

            char[] group2 = new char[] { 'ا', 'آ', 'ر', 'ز', 'ژ', 'د', 'ذ', 'و', 'أ', 'ؤ', 'إ', '@', '$', 'ً' };

            char[] group3 = new char[] { 'ا', 'آ', 'ر', 'ز', 'ژ', 'د', 'ذ', 'و', 'أ', 'ؤ', 'إ', '@', '$', 'ً' };

            string[] group4 = new string[] { "space", "nimspace", "hamze", "allah", "ako", "akc", "paro", "parc", "dot", "tdot", "vir", "simi", "slash", "darsad", "under", "menha", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "zero", "que", "plus", "wonder", "star", "eq", "zarb", "taqsim", "kama", "bro", "brc" };

            string[] group5 = new string[] { "hamze", "hamze", "allah", "ako", "akc", "paro", "parc", "dot", "tdot", "vir", "simi", "slash", "darsad", "under", "menha", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "zero", "que", "plus", "wonder", "star", "eq", "zarb", "taqsim", "kama", "bro", "brc" };

            Array.Sort(group1);
            Array.Sort(group2);
            Array.Sort(group3);
            Array.Sort(group4);
            Array.Sort(groupALL);
            Array.Sort(group5);


            // Array.BinarySearch(group5,wordlab[i]>=0);


            #region nimcheck


            for (int i = 0; i < wordlab.Length - 1; i++)
            {
                if (Array.BinarySearch(group2, word[i]) >= 0 || Array.BinarySearch(group5, wordlab[i]) >= 0)//wordlab[i] == "hamze" || wordlab[i] == "allah")
                {
                    if (wordlab[i + 1] == "space" || wordlab[i + 1] == "nimspace")
                    {
                        nimcheck[i] = 0;
                    }
                    else
                    {
                        nimcheck[i] = 1;
                    }
                }
                else if (Array.BinarySearch(group5, wordlab[i + 1]) >= 0) //(wordlab[i + 1] == "hamze" || wordlab[i + 1] == "allah")
                {
                    nimcheck[i] = 1;
                }
                else
                {
                    nimcheck[i] = 0;
                }



            }

            #endregion


            #region forward


            for (int i = 0; i < wordlab.Length - 1; i++)
            {
                if (Array.BinarySearch(group1, word[i]) >= 0)
                {
                    //if (wordlab[i + 1] == "space" || wordlab[i + 1] == "nimspace")
                    if (Array.BinarySearch(group4, wordlab[i + 1]) >= 0)
                    {
                        forward[i] = 0;
                    }
                    else
                    {
                        forward[i] = 1;
                    }
                }
                else
                {
                    forward[i] = 0;
                }



            }

            #endregion

            #region backward


            for (int i = wordlab.Length - 1; i > 0; i--)
            {
                if (Array.BinarySearch(groupALL, word[i]) >= 0)
                {
                    if (Array.BinarySearch(group2, word[i - 1]) >= 0)
                    {
                        backward[i] = 0;
                    }
                    else if (Array.BinarySearch(group4, wordlab[i - 1]) >= 0)// if (wordlab[i - 1] == "space" || wordlab[i - 1] == "nimspace")
                    {
                        backward[i] = 0;
                    }
                    else
                    {
                        backward[i] = 1;
                    }
                }
                else
                {
                    backward[i] = 0;
                }



            }







            #endregion



            for (int i = 0; i < wordlab.Length; i++)
            {
                if (forward[i] == 1)
                {
                    wordlab[i] = "_" + wordlab[i];

                }
                if (backward[i] == 1)
                {
                    wordlab[i] += "_";
                }


            }


            List<string> mainlab = new List<string>();

            for (int i = 0; i < wordlab.Length; i++)
            {



                if (nimcheck[i] == 1)
                {
                    mainlab.Add(wordlab[i]);
                   // mainlab.Add("nimspace");
                }

                else
                {
                    mainlab.Add(wordlab[i]);
                }
            }
            if (!bMustBeReverse)
                mainlab.Reverse();
            //string []arrStr = mainlab.ToArray();
            //string strReturn = "";
            //for (int i = 0; i < arrStr.Length; i++)
            //{
            //    if (arrStr[i].Contains("nimspace")) continue;
            //    strReturn += arrStr[i];
            //}           
            //strReturn.Replace("h__l__l", "allah");
            //strReturn.Replace("a__l", "la");
            //strReturn.Replace("a__l_", "la_");
            //strReturn.Replace("tanvina", "atan");
            //strReturn.Replace("tanvina_", "atan_");
            return mainlab.ToArray();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            fb.ShowNewFolderButton = true;
            if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputFolder = txtOutputFolder.Text = fb.SelectedPath;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                FontDialog fontDlg = new FontDialog();
                if (fontDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    outputFile = txtOutputFile.Text;
                    outputFolder = txtOutputFolder.Text;
                    string[] inputFiles = txtInputFile.Text.Split(';');
                    List<string> distinctiveWords = new List<string>();
                    string content = "";
                    for (int i = 0; i < inputFiles.Length; i++)
                    {
                        if (!File.Exists(inputFiles[i])) continue;
                        StreamReader sr = new StreamReader(inputFiles[i], Encoding.GetEncoding(1256));
                        content = sr.ReadToEnd();

                        content = content.Replace('\r', ' ');
                        content = content.Replace('\n', ' ');
                        string[] arrWords = content.Split(' ');
                        for (int j = 0; j < arrWords.Length; j++)
                        {
                            if (!distinctiveWords.Contains(arrWords[j]))
                                distinctiveWords.Add(arrWords[j]);
                        }
                        sr.Close();
                    }
                    content = "";
                    for (int i = 0; i < distinctiveWords.Count; i++)
                    {
                        content += distinctiveWords[i] + " ";
                    }
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    object nullObject = Type.Missing;
                    // outputFile = txtOutputFile.Text.Substring(0, txtOutputFile.Text.LastIndexOf(".docx"));
                    outputFolder = txtOutputFolder.Text;
                    inputFile = txtInputFile.Text;
                    Document doc = wordApp.Documents.Add(ref nullObject, ref nullObject, ref nullObject, ref nullObject);
                    /*doc = wordApp.Documents.Open(outputFile, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject);*/
                    // doc.SaveAs2(outputFile);
                    doc.Activate();
                    Paragraph para = doc.Paragraphs.Add();

                    para.Range.Text = content;
                    para.Alignment = WdParagraphAlignment.wdAlignParagraphRight | WdParagraphAlignment.wdAlignParagraphJustify;

                    doc.ActiveWindow.Selection.MoveStart();
                    doc.ActiveWindow.Selection.HomeKey();
                    doc.ActiveWindow.Selection.MoveEnd(WdUnits.wdStory, 1);

                    doc.ActiveWindow.Selection.Font.NameBi = fontDlg.Font.Name;
                    doc.ActiveWindow.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    doc.ActiveWindow.Selection.Font.BoldBi = (fontDlg.Font.Bold) ? 1 : 0;
                    outputFile = outputFolder + "\\" + fontDlg.Font.Name;
                    for (int i = 12; i <= 18; i++)
                    {
                        doc.ActiveWindow.Selection.Font.SizeBi = (float)i;
                        doc.SaveAs2(outputFile + i.ToString() + ".docx");
                    }
                    for (int i = 22; i < 31; i += 4)
                    {
                        doc.ActiveWindow.Selection.Font.SizeBi = (float)i;
                        doc.SaveAs2(outputFile + i.ToString() + ".docx");
                    }
                    object SaveCahnge = WdSaveOptions.wdSaveChanges;

                    object Saveformat = WdOriginalFormat.wdOriginalDocumentFormat;
                    //doc.Close(ref SaveCahnge, ref Saveformat, ref nullObject);
                    doc.Close();
                    wordApp.Quit();
                    MessageBox.Show("Finished");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);

            }
        }
        private Bitmap GetClipboardImage(Microsoft.Office.Interop.Word.Application app, out System.Drawing.Size newSize)
        {
            Bitmap bmp = null;
            System.Drawing.Imaging.Metafile myMetaFile = null;
            newSize = new System.Drawing.Size(1, 1);
            while (true)
            {
                if (Program.OpenClipboard(IntPtr.Zero))
                {

                    if (Program.IsClipboardFormatAvailable(CF_METAFILE) != 0)
                    {

                        intptr = Program.GetClipboardData(CF_METAFILE);

                        myMetaFile = new System.Drawing.Imaging.Metafile(intptr, true);
                        Stream imageStream = new MemoryStream();
                        myMetaFile.Save(imageStream, System.Drawing.Imaging.ImageFormat.Png);
                        Image img = Image.FromStream(imageStream);
                        object h = false;
                        object v = true;
                        newSize = new System.Drawing.Size((int)app.PointsToPixels(app.MillimetersToPoints(myMetaFile.PhysicalDimension.Width / 100), ref h),
                                                                              (int)app.PointsToPixels(app.MillimetersToPoints(myMetaFile.PhysicalDimension.Height / 100), ref v));
                        bmp = new Bitmap(img);
                        myMetaFile.Dispose();
                    }

                    Program.CloseClipboard();
                    break;
                }
            }
            return bmp;
        }
        private void GetDefaultParams(int fontSize, int phase, out int baseLine, out int height)
        {
            int H = 0, def = 0;

            if (fontSize == 12 || fontSize == 13 || fontSize == 14)
            {
                H = 88;
                def = 62;
            }
            if (fontSize == 43 || fontSize == 44 || fontSize == 45)
            {
                H = 216;
                def = 155;
            }
            if (phase == 1)
            {
                if (fontSize == 15 || fontSize == 16 || fontSize == 17 || fontSize == 18)
                {
                    H = 88;
                    def = 62;
                }
                if (fontSize == 19 || fontSize == 20 || fontSize == 21 || fontSize == 22)
                {
                    H = 104;
                    def = 73;
                }
                if (fontSize == 23 || fontSize == 24 || fontSize == 25 || fontSize == 26)
                {
                    H = 120;
                    def = 88;
                }
                if (fontSize == 27 || fontSize == 28 || fontSize == 29 || fontSize == 30)
                {
                    H = 144;
                    def = 103;
                }
                if (fontSize == 31 || fontSize == 32 || fontSize == 33 || fontSize == 34)
                {
                    H = 160;
                    def = 115;
                }
                if (fontSize == 35 || fontSize == 36 || fontSize == 37 || fontSize == 38)
                {
                    H = 176;
                    def = 128;
                }
                if (fontSize == 39 || fontSize == 40 || fontSize == 41 || fontSize == 42)
                {
                    H = 200;
                    def = 142;
                }
            }
            else
            {
                if (fontSize == 15 || fontSize == 16 || fontSize == 17 || fontSize == 18)
                {
                    H = 104;
                    def = 73;
                }
                if (fontSize == 19 || fontSize == 20 || fontSize == 21 || fontSize == 22)
                {
                    H = 120;
                    def = 88;
                }
                if (fontSize == 23 || fontSize == 24 || fontSize == 25 || fontSize == 26)
                {
                    H = 144;
                    def = 103;
                }
                if (fontSize == 27 || fontSize == 28 || fontSize == 29 || fontSize == 30)
                {
                    H = 160;
                    def = 115;
                }
                if (fontSize == 31 || fontSize == 32 || fontSize == 33 || fontSize == 34)
                {
                    H = 176;
                    def = 128;
                }
                if (fontSize == 35 || fontSize == 36 || fontSize == 37 || fontSize == 38)
                {
                    H = 200;
                    def = 142;
                }
                if (fontSize == 39 || fontSize == 40 || fontSize == 41 || fontSize == 42)
                {
                    H = 216;
                    def = 155;
                }
            }
            baseLine = def - 1;
            height = H;
        }

        private void AlignBaseLine(ref IplImage initImage, ref IplImage originalImage, ref int baseLine, int initBaseLine)
        {
            const int heightTlorance = 2;
            const int MIN_PROJECTION = 10;
            unsafe
            {
                byte* pData1 = (byte*)initImage.imageData;
                double maxSum = 0;
                double sum = 0;
                int minimumY = -1;
                int maximumY = -1;
                for (int i = 0; i < initImage.height; i++)
                {
                    sum = 0;
                    for (int j = 0; j < initImage.width; j++)
                    {
                        sum += pData1[i * initImage.widthStep + j * 3];
                    }
                    if (sum > maxSum)
                    {
                        maxSum = sum;
                        if (sum / 255 > 5 && minimumY < 0)
                            minimumY = i;
                    }
                    if (sum / 255 > 5)
                        maximumY = i;
                }
                int minY = 0, maxY = 0;
                if (maxSum < MIN_PROJECTION || minY < 0 || maxY < 0)
                {
                    baseLine = -1;
                    return;
                }

                //initBaseLine = baseLine;
                int fontHeight = initBaseLine - minimumY;
                int roiHeight = maximumY - minimumY;
                if ((baseLine - fontHeight) + (maximumY - minimumY + heightTlorance) > originalImage.height)
                {
                    roiHeight = originalImage.height - (baseLine - fontHeight);
                }
                cvlib.CvSetImageROI(ref initImage, cvlib.cvRect(0, initBaseLine - fontHeight, originalImage.width, roiHeight));
                cvlib.CvSetImageROI(ref originalImage, cvlib.cvRect(0, baseLine - fontHeight, originalImage.width, roiHeight));
                cvlib.CvCopy(ref initImage, ref originalImage);
                //cvlib.CvSaveImage("D:\\b.bmp", ref originalImage);
                cvlib.CvResetImageROI(ref initImage);
                cvlib.CvResetImageROI(ref originalImage);
            }
        }
        private void Binarize(ref IplImage srcImage)
        {
            unsafe
            {
                int W = srcImage.widthStep, H = srcImage.height;
                byte* image = (byte*)srcImage.imageData;
                int[] hist = new int[256];
                for (int i = 0; i < 256; i++)
                {
                    hist[i] = 0;
                }
                for (int y = 0; y < H; y++)
                {
                    for (int x = 0; x < W; x++)
                    {
                        int value = image[y * W + x];
                        if (value <= 0)
                        {
                            hist[0]++;
                        }
                        else
                        {
                            if (value >= 255)
                                hist[255]++;
                            else
                                hist[value]++;
                        }
                    }
                }
                double mean = 0;
                double std_dev = 0;
                for (int i = 0; i < 256; i++)
                {
                    mean += i * hist[i];
                    std_dev += i * i * hist[i];
                }
                mean /= (W * H);
                std_dev = std_dev / (W * H) - mean * mean;
                if (std_dev <= 0)
                    std_dev = 0;
                else
                    std_dev = Math.Sqrt(std_dev);
                double k = 0.5;
                // tempNi = (tempmean * (1 - K * (1 - (tempstd / 128))));

                int tsh = (int)(mean * (1 - k * (1 - std_dev / 128)));
                //int tsh = 1;
                for (int i = 0; i < W * H; i++)
                    image[i] = (image[i] >= tsh) ? (byte)255 : (byte)0;

            }
        }
        private void DestroyImage(ref IplImage srcImage, ref IplImage noiseImage, int dilateRectWidth, int dilateRectHeight)
        {

            const int NTimes = 2;
            int[] values = new int[9];
            param1 = double.Parse(txtParam1.Text);
            param2 = double.Parse(txtParam2.Text);
            for (int i = 0; i < 9; i++)
            {
                values[i] = 1;
            }
            unsafe
            {
                //cvlib.CvShowImage("", ref noiseImage); cvlib.CvWaitKey(0);
                byte* image = (byte*)noiseImage.imageData;
                byte* src = (byte*)srcImage.imageData;
                int W = srcImage.widthStep, H = srcImage.height;
                Random rand = new Random();
                IplConvKernel convKernel = cvlib.CvCreateStructuringElementEx(3, 3, 1, 1, cvlib.CV_SHAPE_RECT, values);
                for (int i = 0; i < (H - dilateRectHeight * 3 - 1); i++)
                    for (int j = 0; j < (W - dilateRectWidth * 3 - 1); j += 3)
                    {
                        //if( (image[i * W + j] > (param1 - 3) && image[i * W + j] < (param1 + 3) ) || (image[i * W + j] > (param2 - 3) && image[i * W + j] < (param2 + 3)) )
                        if (image[i * W + j] == param1 || image[i * W + j] == param2)
                        {

                            cvlib.CvSetImageROI(ref srcImage, cvlib.cvRect(j / 3, i, rand.Next(2, dilateRectWidth), rand.Next(2, dilateRectHeight)));
                            cvlib.CvDilate(ref srcImage, ref srcImage, ref convKernel, NTimes);
                            cvlib.CvErode(ref srcImage, ref srcImage, ref convKernel, NTimes);

                        }
                        if (image[i * noiseImage.widthStep + j] == param1)
                        {
                            src[i * noiseImage.widthStep + j] = 0;
                            src[i * noiseImage.widthStep + j + 1] = 0;
                            src[i * noiseImage.widthStep + j + 2] = 0;
                        }
                        else
                            if (image[i * noiseImage.widthStep + j] == param2)
                            {
                                src[i * noiseImage.widthStep + j] = 255;
                                src[i * noiseImage.widthStep + j + 1] = 255;
                                src[i * noiseImage.widthStep + j + 2] = 255;

                            }
                    }
                cvlib.CvReleaseStructuringElement(ref convKernel);
            }
            cvlib.CvResetImageROI(ref srcImage);
        }
        private void DestroyImage2(ref IplImage srcImage, ref IplImage edgeImage, double pdf, int MaxPathLenght)
        {
            CvPoint pnt, nextPnt;
            int index;
            byte exchangeValue;
            int pathLen = 0;
            Random rand = new Random();
            unsafe
            {
                byte* edgePtr = (byte*)edgeImage.imageData;
                byte* srcPtr = (byte*)srcImage.imageData;
                List<int> pnts = new List<int>();
                int W = edgeImage.widthStep, H = edgeImage.height;
                int srcW = srcImage.widthStep, srcH = srcImage.height;
                for (int i = 0; i < H; i++)
                {
                    for (int j = 0; j < edgeImage.width; j++)
                    {
                        if (edgePtr[i * W + j] > 0)
                            pnts.Add(i * W + j);
                    }
                }
                int randomCount = (int)(pdf * pnts.Count);
                for (int i = 0; i < randomCount; i++)
                {
                    while (true)
                    {
                        index = pnts.ElementAt(rand.Next(0, pnts.Count));
                        pnt.y = index / W;
                        pnt.x = index % W;
                        if (edgePtr[pnt.y * W + pnt.x] == 255)
                            break;
                    }
                    if (rand.Next(0, 1) == 0)//dilation
                    {
                        exchangeValue = 255;
                    }
                    else//erosion
                    {
                        exchangeValue = 0;
                    }
                    pathLen = rand.Next(1, MaxPathLenght);
                    nextPnt = pnt;
                    for (int j = 0; j < pathLen; j++)
                    {
                        pnt = nextPnt;
                        edgePtr[pnt.y * W + pnt.x] = 128;
                        if (pnt.y >= (srcImage.height - 2) || pnt.x >= (srcImage.width - 2))
                            break;
                        for (int m = -1; m <= 1; m++)
                        {
                            for (int n = -1; n <= 1; n++)
                            {
                                index = (pnt.y - m) * srcW + (pnt.x - n) * 3;
                                srcPtr[index] = exchangeValue; srcPtr[index + 1] = exchangeValue; srcPtr[index + 2] = exchangeValue;
                                if (edgePtr[(pnt.y - m) * W + (pnt.x - n)] == 255)
                                {
                                    nextPnt.x = (pnt.x - n);
                                    nextPnt.y = (pnt.y - m);
                                }
                            }
                        }
                    }
                }

            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            outputFile = txtOutputFile.Text;
            outputFolder = txtOutputFolder.Text;
            inputFile = txtInputFile.Text;
            if (outputFile == "" || outputFile == null)
            {
                MessageBox.Show("Please, Enter or browse Path of output File which contains Word File");
                return;
            }
            /*  if (outputFolder == "" || outputFolder == null)
              {
                  MessageBox.Show("Please, Enter or browse Path of output Folder");
                  return;
              }*/
            if (cmbNoise.SelectedIndex < 0 || cmbNoise.SelectedIndex > 2)
            {
                MessageBox.Show("Please, Select a Noise Distribution");
                return;
            }
            if (txtParam1.Text == "" || txtParam2.Text == "")
            {
                MessageBox.Show("Please, Enter Params's value of Noise Distribution");
                return;
            }
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object nullObject = Type.Missing;
            Document doc = null;

            //System.Threading.Thread.CurrentThread.SetApartmentState(System.Threading.ApartmentState.STA); 
            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                int fileCounter = 22;
                string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database\Nazanin";
                //while (fileCounter <= 45)
                {
                    outputFile = txtOutputFile.Text;//defaultPath + fileCounter.ToString() + ".docx";
                    doc = wordApp.Documents.Open(outputFile, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject);
                    doc.Activate();

                    int counter = 0;
                    object x = 5, y = 1, z = WdMovementType.wdExtend;
                    bool bBreak = false;
                    doc.ActiveWindow.Selection.StartOf();
                    doc.ActiveWindow.Selection.HomeKey();
                    doc.ActiveWindow.Selection.MoveEnd(WdUnits.wdWord, 1);
                    outputFile = outputFile.Substring(0, txtOutputFile.Text.LastIndexOf(".docx"));
                    Directory.CreateDirectory(outputFile);
                    outputFolder = outputFile;
                    string HFilesPath = outputFolder + "\\" + outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2) + "H\\"; ;
                    string LFilesPath = outputFolder + "\\" + outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2) + "L\\"; ;
                    param1 = double.Parse(txtParam1.Text);
                    param2 = double.Parse(txtParam2.Text);
                    int fontSize = 0, baseLine = 0, imageHeight = 0, imageWidth = 0;
                    fontSize = 15;//int.Parse(outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2));//(int)doc.ActiveWindow.Selection.Font.Size;
                    Directory.CreateDirectory(LFilesPath);
                    Directory.CreateDirectory(HFilesPath);
                    doc.ActiveWindow.Selection.StartOf();
                    ulong rgn = 1;
                    while (true)
                    {
                        doc.ActiveWindow.Selection.HomeKey();
                        doc.ActiveWindow.Selection.MoveEnd(x, z);
                        if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                            bBreak = true;
                        doc.ActiveWindow.Selection.Copy();
                        System.Drawing.Size newSize;
                        Bitmap bmp = GetClipboardImage(wordApp, out newSize);
                        string strImagePath = System.IO.Path.GetTempFileName().Replace(".tmp", ".png");
                        double heightRatio = 1.95;
                        double widthRatio = 2.05;
                        bmp.Save(strImagePath);
                        bmp.Dispose();
                        bmp = null;
                        IplImage imagePNG = cvlib.CvLoadImage(strImagePath, cvlib.CV_LOAD_IMAGE_UNCHANGED);

                        IplImage initImage = cvlib.CvCreateImage(cvlib.CvSize((int)(imagePNG.width / widthRatio), (int)(imagePNG.height / heightRatio)), 8, 3);
                        IplImage tmpImage = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 3);
                        IplImage image = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1);
                        IplImage image1 = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1),
                                 image2 = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1),
                                 image3 = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1);
                        cvlib.CvSplit(ref imagePNG, ref image1, ref image2, ref image3, ref image);
                        cvlib.CvMerge(ref image, ref image, ref image, ref tmpImage);
                        cvlib.CvResize(ref tmpImage, ref initImage, 1);
                        //cvlib.CvSaveImage("D:\\a.bmp", ref initImage);
                        // for (int phase = 0; phase < 2; phase++)
                        for (int phase = 1; phase < 2; phase++)
                        {
                            string picOutPath = ((phase == 0) ? HFilesPath : LFilesPath) + Path.GetFileNameWithoutExtension(outputFile) + "_" + counter.ToString() + ".bmp";
                            string picOutPath2 = ((phase == 0) ? HFilesPath : LFilesPath) + Path.GetFileNameWithoutExtension(outputFile) + "_" + counter.ToString() + "_B.bmp";

                            GetDefaultParams(fontSize, phase, out baseLine, out imageHeight);
                            IntPtr finalBaseLine;
                            int initBaseLine = 0;
                            unsafe { finalBaseLine = new IntPtr(&initBaseLine); }
                            IntPtr p = MapToOriginalSize(ref initImage);
                            IplImage outputImage = (IplImage)Marshal.PtrToStructure(p, typeof(IplImage));
                            //cvlib.CvShowImage("d:\\temp1.bmp", ref outputImage); cvlib.CvWaitKey(0);
                            IplImage originalImage = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, imageHeight), 8, 3);
                            IplImage edgeImage = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, imageHeight), 8, 1);
                            IplImage matRnd = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, originalImage.height), 8, 3);
                            IplImage matRndOneChannel = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, originalImage.height), 8, 1);
                            cvlib.CvSet(ref originalImage, cvlib.cvScalar(0, 0, 0, 0));
                            // MessageBox.Show("1");
                            FindBaseLine(ref outputImage, finalBaseLine, fontSize);
                            AlignBaseLine(ref outputImage, ref originalImage, ref baseLine, initBaseLine);
                            if (baseLine < 0)
                            {
                                cvlib.CvReleaseImage(ref matRnd);
                                cvlib.CvReleaseImage(ref originalImage);
                                cvlib.CvReleaseImage(ref matRndOneChannel);
                                counter--;
                                break;
                            }
                            cvlib.CvNot(ref originalImage, ref originalImage);
                            Binarize(ref originalImage);

                            // MessageBox.Show(System.Windows.Forms.Application.StartupPath);
                            // MessageBox.Show(Directory.GetCurrentDirectory());

                            //Directory.SetCurrentDirectory(System.Windows.Forms.Application.StartupPath);
                            SaveImage(picOutPath, ref originalImage);
                            //cvlib.CvShowImage("2", ref originalImage); cvlib.CvWaitKey(0);
                            //cvlib.CvShowImage("1ee", ref originalImage); cvlib.CvWaitKey(0);

                            cvlib.CvSetImageCOI(ref originalImage, 0);
                            cvlib.CvCanny(ref originalImage, ref edgeImage, 50, 100, 5);
                            //cvlib.CvShowImage("0", ref originalImage); 
                            DestroyImage2(ref originalImage, ref edgeImage, 0.0081, 3);


                            //cvlib.CvShowImage("1", ref edgeImage); cvlib.CvShowImage("2", ref originalImage); cvlib.CvWaitKey(0);
                            if (cmbNoise.SelectedIndex == NOISE_SALT_PEPPER)
                            {
                                cvlib.CvRandArr(ref rgn, ref matRnd, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(0, 0, 0, 0), cvlib.cvScalar(255, 255, 255, 0));
                                unsafe
                                {
                                    byte* pData1 = (byte*)matRnd.imageData;
                                    byte* pData2 = (byte*)originalImage.imageData;

                                    for (int i = 0; i < originalImage.height; i++)
                                        for (int j = 0; j < originalImage.widthStep; j++)
                                        {
                                            if (bOneChannel)
                                            {
                                                if (pData1[i * matRnd.widthStep + j] == param1)
                                                {
                                                    pData2[i * matRnd.widthStep + j] = 0;
                                                    pData2[i * matRnd.widthStep + j + 1] = 0;
                                                    pData2[i * matRnd.widthStep + j + 2] = 0;
                                                }
                                                else
                                                    if (pData1[i * matRnd.widthStep + j] == param2)
                                                    {
                                                        pData2[i * matRnd.widthStep + j] = 255;
                                                        pData2[i * matRnd.widthStep + j + 1] = 255;
                                                        pData2[i * matRnd.widthStep + j + 2] = 255;

                                                    }
                                                j += 2;
                                            }
                                            else
                                            {
                                                if (pData1[i * matRnd.widthStep + j] < param1)
                                                {
                                                    pData2[i * matRnd.widthStep + j] = 0;
                                                }
                                                else
                                                    if (pData1[i * matRnd.widthStep + j] > param2)
                                                    {
                                                        pData2[i * matRnd.widthStep + j] = 255;
                                                    }
                                            }
                                        }
                                }
                            }
                            else
                            {
                                if (bOneChannel)
                                {
                                    cvlib.CvRandArr(ref rgn, ref matRndOneChannel, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(param1, 0, 0, 0), cvlib.cvScalar(param2, 0, 0, 0));
                                    cvlib.CvMerge(ref matRndOneChannel, ref matRndOneChannel, ref matRndOneChannel, ref matRnd);
                                }
                                else
                                    cvlib.CvRandArr(ref rgn, ref matRnd, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(param1, param1, param1, 0), cvlib.cvScalar(param2, param2, param2, 0));

                                cvlib.CvAdd(ref originalImage, ref matRnd, ref originalImage);
                            }
                            Binarize(ref originalImage);
                            if (cmbNoise.SelectedIndex != NOISE_SALT_PEPPER)
                            {
                                param1 = 1;
                                param2 = 255;
                            }
                            {
                                //DestroyImage(ref originalImage, ref matRnd, fontSize / 4, fontSize / 4);
                                // cvlib.CvRandArr(ref rgn, ref matRnd, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(0, 0, 0, 0), cvlib.cvScalar(255, 255, 255, 0));
                                // DestroyImage(ref originalImage, ref matRnd, 7, 7);
                            }
                            cvlib.CvSmooth(ref originalImage, ref originalImage, cvlib.CV_MEDIAN, 3, 3, 0, 0);
                            SaveImage(picOutPath2, ref originalImage);
                            //cvlib.CvShowImage("", ref originalImage); cvlib.CvWaitKey(0);
                            cvlib.CvReleaseImage(ref matRnd);
                            cvlib.CvReleaseImage(ref originalImage);
                            cvlib.CvReleaseImage(ref matRndOneChannel);
                            cvlib.CvReleaseImage(ref edgeImage);
                            Marshal.StructureToPtr(outputImage, p, false);
                            ReleaseMap(p);
                        }
                        cvlib.CvReleaseImage(ref image);
                        cvlib.CvReleaseImage(ref image1);
                        cvlib.CvReleaseImage(ref image2);
                        cvlib.CvReleaseImage(ref image3);
                        cvlib.CvReleaseImage(ref imagePNG);
                        cvlib.CvReleaseImage(ref tmpImage);
                        cvlib.CvReleaseImage(ref initImage);
                        File.Delete(strImagePath);
                        doc.ActiveWindow.Selection.MoveDown(x, 1, WdMovementType.wdMove);

                        if (bBreak)
                            break;
                        counter++;
                    }


                    // doc.Close(ref nullObject, ref nullObject, ref nullObject);
                    fileCounter++;
                }
                // wordApp.Quit();
                wordApp = null;
                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);
                if (doc != null) doc.Close(ref nullObject, ref nullObject, ref nullObject);
                if (wordApp != null) wordApp.Quit();
                MessageBox.Show("Finished unsuccessfully");

            }
        }

        private void cmbNoise_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbNoise.SelectedIndex)
            {
                case NOISE_GAUSSIAN:
                    toolTip1.SetToolTip(txtParam1, "Please Enter value of Mean Of GAUSSIAN Noise Distribution");
                    toolTip1.SetToolTip(txtParam2, "Please Enter value of Standard Deviation Of GAUSSIAN Noise Distribution");
                    break;
                case NOISE_UNIFORM:
                    toolTip1.SetToolTip(txtParam1, "Please Enter Minimum value Of UNIFORM Noise Distribution");
                    toolTip1.SetToolTip(txtParam2, "Please Enter Maximum value Of UNIFORM Noise Distribution");
                    break;
                case NOISE_SALT_PEPPER:
                    toolTip1.SetToolTip(txtParam1, "Please Enter Low Cut value Of Salt&Pepper Noise Distribution\n.ie 30 according to range(0 - 255)");
                    toolTip1.SetToolTip(txtParam2, "Please Enter High Cut value Of Salt&Pepper Noise Distribution\n.ie 220 according to range(0 - 255)");
                    break;
                default:
                    break;
            }
        }

        private void chkOneChannel_CheckedChanged(object sender, EventArgs e)
        {
            bOneChannel = chkOneChannel.Checked;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            outputFile = txtOutputFile.Text;
            outputFolder = txtOutputFolder.Text;
            inputFile = txtInputFile.Text;
            if (outputFile == "" || outputFile == null)
            {
                MessageBox.Show("Please, Enter or browse Path of output File which contains Word File");
                return;
            }
            if (outputFolder == "" || outputFolder == null)
            {
                MessageBox.Show("Please, Enter or browse Path of output Folder");
                return;
            }
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object nullObject = Type.Missing;
            Document doc = null;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                doc = wordApp.Documents.Open(outputFile, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                    , ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                    , ref nullObject, ref nullObject
                    , ref nullObject, ref nullObject);
                doc.Activate();

                int counter = 0;
                object x = 5, y = 1, z = WdMovementType.wdExtend;
                doc.ActiveWindow.Selection.StartOf();
                bool bBreak = false;
                doc.ActiveWindow.Selection.HomeKey();
                while (true)
                {

                    doc.ActiveWindow.Selection.MoveRight(WdUnits.wdCharacter, 1, z);
                    if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                        bBreak = true;
                    doc.ActiveWindow.Selection.CopyAsPicture();

                    System.Drawing.Size newSize;
                    Bitmap bmp = GetClipboardImage(wordApp, out newSize);
                    string picOutPath = outputFolder + "\\" + Path.GetFileNameWithoutExtension(inputFile) + counter.ToString() + ".bmp";
                    bmp.Save(picOutPath);
                    doc.ActiveWindow.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                    if (bBreak)
                        break;
                    counter++;
                }


                doc.Close(ref nullObject, ref nullObject, ref nullObject);
                wordApp.Quit();
                wordApp = null;
                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);
                if (doc != null) doc.Close(ref nullObject, ref nullObject, ref nullObject);
                if (wordApp != null) wordApp.Quit();
                MessageBox.Show("Finished unsuccessfully");

            }
        }
        struct XParamInfo
        {
            public string i;
            public string o;
            public string opt;
            public string sw;
            public string ls;
            public string fp;
            public string defbl;
            public string h;
            public string fs;
            public string is_b;
            public string si;
            public XParamInfo(string x_i, string x_o, string x_opt, string x_sw, string x_ls, string x_fp,
                              string x_defbl, string x_h, string x_fs, string x_is_b, string x_si)
            {
                i = x_i; o = x_o; opt = x_opt; sw = x_sw; ls = x_ls; fp = x_fp;
                defbl = x_defbl; h = x_h; fs = x_fs; is_b = x_is_b; si = x_si;
            }
        }
        static XParamInfo Model1Params = new XParamInfo("", "", "1", "4", "1", "0.01", "62", "88", "", "0", "10");
        static XParamInfo Model2Params = new XParamInfo("", "", "1", "4", "1", "0.01", "73", "104", "", "0", "10");
        private void button8_Click(object sender, EventArgs e)
        {

            try
            {
                int fontSize = 12;
                //string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database2\Nazanin";
                string defaultPath = txtOutputFolder.Text;
                //Model1(12 - 18 L)
                //while (fontSize <= 31)
                {
                    //outputFile = defaultPath + fontSize.ToString() + ".docx";
                    //if (!File.Exists(outputFile))
                    //{
                    //    fontSize++;
                    //    continue;
                    //}
                    //outputFolder = defaultPath + fontSize.ToString();

                    // string LFilesPath = outputFolder + "\\" + fontSize.ToString() + "L\\";
                    string[] filenames = Directory.GetFiles(defaultPath, "*.bmp", SearchOption.AllDirectories);

                    Process p = new Process();
                    ProcessStartInfo pp = new ProcessStartInfo();
                    pp.FileName = "fe_ex.exe";
                    pp.CreateNoWindow = true;
                    pp.WindowStyle = ProcessWindowStyle.Hidden;
                    int baseLine = 62;
                    foreach (string file in filenames)
                    {
                        try
                        {
                            lblCurrentFile.Text = "Current File : " + file;
                            Refresh();
                            string strImagePath = System.IO.Path.GetTempFileName().Replace(".tmp", ".bmp");
                            File.Delete(Path.GetDirectoryName(strImagePath) + "\\" + Path.GetFileNameWithoutExtension(strImagePath) + ".tmp");
                            if (!File.Exists(file))
                                continue;
                            if (File.Exists(Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + ".fea"))
                                continue;
                            IplImage imgSrc = cvlib.CvLoadImage(file, cvlib.CV_LOAD_IMAGE_UNCHANGED);
                            int initBaseLine = 0;
                            IntPtr finalBaseLine;
                            unsafe { finalBaseLine = new IntPtr(&initBaseLine); }
                            cvlib.CvNot(ref imgSrc, ref imgSrc);
                            IplImage originalImage = cvlib.CvCreateImage(cvlib.CvSize(imgSrc.width, imgSrc.height), 8, 3);
                            cvlib.CvSet(ref originalImage, cvlib.cvScalar(0, 0, 0, 0));
                            FindBaseLine(ref imgSrc, finalBaseLine, 15);
                            if (baseLine != initBaseLine)
                            {
                                AlignBaseLine(ref imgSrc, ref originalImage, ref baseLine, initBaseLine);
                                cvlib.CvNot(ref originalImage, ref originalImage);
                                SaveImage(strImagePath, ref originalImage);

                            }
                            cvlib.CvReleaseImage(ref imgSrc);
                            cvlib.CvReleaseImage(ref originalImage);
                            if (baseLine != initBaseLine)
                                Model1Params.i = strImagePath;
                            else
                                Model1Params.i = file;
                            Model1Params.fs = "15";
                            Model1Params.o = Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + ".fea";
                            pp.Arguments = string.Format("-i {0} -o {1} -opt {2} -sw {3} -ls {4} -fp {5} -defbl {6} -h {7} -fs {8} -is_b {9} -si {10}",
                                Model1Params.i, Model1Params.o, Model1Params.opt, Model1Params.sw, Model1Params.ls, Model1Params.fp, Model1Params.defbl, Model1Params.h, Model1Params.fs, Model1Params.is_b, Model1Params.si);
                            p.StartInfo = pp;
                            p.Start();
                            p.WaitForExit();
                            if (baseLine != initBaseLine)
                                File.Delete(strImagePath);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);

                        }
                    }
                    fontSize++;
                }
                return;
                //Model2(15 - 18 H)
                fontSize = 15;
                while (fontSize <= 18)
                {
                    outputFolder = defaultPath + fontSize.ToString();
                    string HFilesPath = outputFolder + "\\" + fontSize.ToString() + "H\\";
                    string[] filenames = Directory.GetFiles(HFilesPath, "*.bmp", SearchOption.AllDirectories);
                    Process p = new Process();
                    ProcessStartInfo pp = new ProcessStartInfo();
                    pp.FileName = "fe_ex.exe";
                    pp.CreateNoWindow = true;
                    pp.WindowStyle = ProcessWindowStyle.Hidden;
                    foreach (string file in filenames)
                    {
                        try
                        {

                            lblCurrentFile.Text = "Current File : " + file;
                            Refresh();
                            Model2Params.i = file;
                            Model2Params.o = Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + ".fea"; ;
                            Model2Params.fs = fontSize.ToString();
                            pp.Arguments = string.Format("-i {0} -o {1} -opt {2} -sw {3} -ls {4} -fp {5} -defbl {6} -h {7} -fs {8} -is_b {9} -si {10}",
                                Model2Params.i, Model2Params.o, Model2Params.opt, Model2Params.sw, Model2Params.ls, Model2Params.fp, Model2Params.defbl, Model2Params.h, Model2Params.fs, Model2Params.is_b, Model2Params.si);
                            p.StartInfo = pp;
                            p.Start();
                            p.WaitForExit();

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);

                        }
                    }
                    fontSize++;
                }
                //Model2(19 - 22 L)
                fontSize = 19;
                while (fontSize <= 22)
                {
                    outputFolder = defaultPath + fontSize.ToString();
                    string LFilesPath = outputFolder + "\\" + fontSize.ToString() + "L\\";
                    string[] filenames = Directory.GetFiles(LFilesPath, "*.bmp", SearchOption.AllDirectories);
                    Process p = new Process();
                    ProcessStartInfo pp = new ProcessStartInfo();
                    pp.FileName = "fe_ex.exe";
                    pp.CreateNoWindow = true;
                    pp.WindowStyle = ProcessWindowStyle.Hidden;
                    foreach (string file in filenames)
                    {
                        try
                        {

                            lblCurrentFile.Text = "Current File : " + file;
                            Refresh();
                            Model2Params.i = file;
                            Model2Params.o = Path.GetDirectoryName(file) + "\\" + Path.GetFileNameWithoutExtension(file) + ".fea"; ;
                            Model2Params.fs = fontSize.ToString();
                            pp.Arguments = string.Format("-i {0} -o {1} -opt {2} -sw {3} -ls {4} -fp {5} -defbl {6} -h {7} -fs {8} -is_b {9} -si {10}",
                                Model2Params.i, Model2Params.o, Model2Params.opt, Model2Params.sw, Model2Params.ls, Model2Params.fp, Model2Params.defbl, Model2Params.h, Model2Params.fs, Model2Params.is_b, Model2Params.si);
                            p.StartInfo = pp;
                            p.Start();
                            p.WaitForExit();

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);

                        }
                    }
                    fontSize++;
                }
                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string strOut = new string(' ', 5 * 1024);
                GetOCRText(of.FileName, ref strOut);
                string outputFile = Path.GetDirectoryName(of.FileName) + "\\" + Path.GetFileNameWithoutExtension(of.FileName) + ".txt";
                OCRResult ocrFrm = new OCRResult();
                ocrFrm.lblOCR.Text = strOut;
                StreamWriter sw = new StreamWriter(outputFile);
                sw.Write(strOut);
                sw.Close();
                ocrFrm.Show();
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

            try
            {
                int fontSize = 20;
                //string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database2\Nazanin";
                string defaultPath = txtOutputFolder.Text;//@"D:\Projects\A-OCR\new_ocr\Database2\\Nazanin16\\16H\\";
                //string[] defaultPaths = {@"D:\Projects\A-OCR\new_ocr\Database2\Nazanin26\26H\", @"D:\Projects\A-OCR\new_ocr\Database2\Nazanin30\30H\"};

                // for (int i = 0; i < 2; i++)
                {
                    // defaultPath = defaultPaths[i];
                    string[] filenames = Directory.GetFiles(defaultPath, "*.bmp", SearchOption.AllDirectories);
                    int dstHeight = 88;
                    CvRect roi = cvlib.cvRect(0, 10, 0, dstHeight);
                    double angle = 0;
                    CvPoint2D32f center;
                    CvMat rotMat = cvlib.CvCreateMat(2, 3, cvlib.CV_32FC1);
                    double[] angles = { -0.1, -0.05, 0.05, 0.1 };
                    string dstFile;
                    center.x = 0;
                    for (int rotCounter = 0; rotCounter < 4; rotCounter++)
                    {
                        angle = angles[rotCounter];

                        foreach (string file in filenames)
                        {
                            try
                            {
                                lblCurrentFile.Text = file;
                                this.Refresh();
                                dstFile = Path.GetDirectoryName(file) + "\\rot" + (rotCounter + 1).ToString() + "\\" + Path.GetFileName(file);
                                if (!Directory.Exists(Path.GetDirectoryName(file) + "\\rot" + (rotCounter + 1).ToString()))
                                    Directory.CreateDirectory(Path.GetDirectoryName(file) + "\\rot" + (rotCounter + 1).ToString());
                                IplImage imgSrc = cvlib.CvLoadImage(file, cvlib.CV_LOAD_IMAGE_UNCHANGED);
                                center.y = imgSrc.height / 2;
                                cvlib.Cv2DRotationMatrix(center, angle, 1.0, ref rotMat);

                                IplImage imgDst = cvlib.CvCreateImage(cvlib.CvSize(imgSrc.width, imgSrc.height), 8, 3);
                                cvlib.CvWarpAffine(ref imgSrc, ref imgDst, ref rotMat, cvlib.CV_INTER_CUBIC + cvlib.CV_WARP_FILL_OUTLIERS, cvlib.cvScalarAll(255));
                                cvlib.CvThreshold(ref imgDst, ref imgDst, 128, 255, cvlib.CV_THRESH_BINARY);
                                SaveImage(dstFile, ref imgDst);
                                cvlib.CvReleaseImage(ref imgSrc);
                                cvlib.CvReleaseImage(ref imgDst);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex);

                            }
                        }

                    }
                }

                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {

            try
            {
                int fontSize = 20;
                //string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database2\Nazanin";
                string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database2";


                string[] filenames = Directory.GetFiles(defaultPath, "*.bmp", SearchOption.AllDirectories);
                int dstHeight = 88;
                CvRect roi = cvlib.cvRect(0, 10, 0, dstHeight);
                foreach (string file in filenames)
                {
                    try
                    {
                        lblCurrentFile.Text = file;
                        IplImage imgSrc = cvlib.CvLoadImage(file, cvlib.CV_LOAD_IMAGE_UNCHANGED);
                        roi.width = imgSrc.width;
                        IplImage imgDst = cvlib.CvCreateImage(cvlib.CvSize(imgSrc.width, dstHeight), 8, 3);
                        cvlib.CvSetImageROI(ref imgSrc, roi);
                        cvlib.CvCopy(ref imgSrc, ref imgDst);
                        cvlib.CvSaveImage(file, ref imgDst);
                        cvlib.CvReleaseImage(ref imgSrc);
                        cvlib.CvReleaseImage(ref imgDst);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);

                    }
                }

                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);

            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog of = new OpenFileDialog();
                of.DefaultExt = "txt";
                of.Filter = "Text files (*.txt)|*.txt";
                of.Multiselect = true;
                if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    for (int m = 0; m < of.FileNames.Length; m++)
                    {

                        SaveFileDialog sf = new SaveFileDialog();
                        sf.DefaultExt = "lab";
                        sf.AddExtension = true;
                        sf.FileName = Path.GetFileNameWithoutExtension(of.FileName);
                        sf.Filter = "Lable files (*.lab)|*.lab";
                        // if ( sf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            outputFile = Path.GetDirectoryName(of.FileNames[m]) + "\\" + Path.GetFileNameWithoutExtension(of.FileNames[m]) + ".lab"; //sf.FileNames[i];
                            inputFile = of.FileNames[m];
                            StreamReader sr = new StreamReader(inputFile, Encoding.GetEncoding(1256));
                            StreamWriter sw = new StreamWriter(outputFile);
                            string content = "";
                            content = sr.ReadToEnd();
                            content = content.Replace('\r', ' ');
                            content = content.Replace('\n', ' ');
                            content = content.Replace('\t', ' ');

                            string[] words = content.Split(new char[] { ' ', '\n', '\t', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                            string strTmp = "";
                            for (int i = words.Length - 1; i >= 0; i--)
                            {
                                strTmp = words[i];

                                Regex regex = new Regex(@"\W*لله$");
                                Match match = regex.Match(strTmp);
                                while (match != null && match.Length != 0)
                                {
                                    strTmp = strTmp.Replace(match.Value, "#");
                                    match = match.NextMatch();
                                }
                                // strTmp = strTmp.Replace("لله", "#");
                                strTmp = strTmp.Replace("اً", "@");
                                strTmp = strTmp.Replace("لا", "$");

                                string[] conLine = GetCode(strTmp);

                                for (int j = 0; j < conLine.Length; j++)
                                {
                                    if (conLine[j].Contains("nimspace"))
                                        continue;
                                    sw.Write("{0}\n", conLine[j]);
                                }

                            }
                            sw.Close();
                            sr.Close();
                        }

                    }
                    MessageBox.Show("Finished");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);

            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            outputFile = txtOutputFile.Text;
            outputFolder = txtOutputFolder.Text;
            inputFile = txtInputFile.Text;
            if (outputFile == "" || outputFile == null)
            {
                MessageBox.Show("Please, Enter or browse Path of output File which contains Word File");
                return;
            }
            /*  if (outputFolder == "" || outputFolder == null)
              {
                  MessageBox.Show("Please, Enter or browse Path of output Folder");
                  return;
              }*/
            if (cmbNoise.SelectedIndex < 0 || cmbNoise.SelectedIndex > 2)
            {
                MessageBox.Show("Please, Select a Noise Distribution");
                return;
            }
            if (txtParam1.Text == "" || txtParam2.Text == "")
            {
                MessageBox.Show("Please, Enter Params's value of Noise Distribution");
                return;
            }
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object nullObject = Type.Missing;
            Document doc = null;

            //System.Threading.Thread.CurrentThread.SetApartmentState(System.Threading.ApartmentState.STA); 
            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                int fileCounter = 22;
                string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database\Nazanin";
                //while (fileCounter <= 45)
                {
                    outputFile = txtOutputFile.Text;//defaultPath + fileCounter.ToString() + ".docx";
                    doc = wordApp.Documents.Open(outputFile, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject);
                    doc.Activate();

                    int counter = 0;
                    object x = 5, y = 1, z = WdMovementType.wdExtend;
                    bool bBreak = false;
                    doc.ActiveWindow.Selection.StartOf();
                    doc.ActiveWindow.Selection.HomeKey();
                    doc.ActiveWindow.Selection.MoveEnd(WdUnits.wdWord, 1);
                    outputFile = outputFile.Substring(0, txtOutputFile.Text.LastIndexOf(".docx"));
                    Directory.CreateDirectory(outputFile);
                    outputFolder = outputFile;
                    string HFilesPath = outputFolder + "\\" + outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2) + "H\\"; ;
                    string LFilesPath = outputFolder + "\\" + outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2) + "L\\"; ;
                    param1 = double.Parse(txtParam1.Text);
                    param2 = double.Parse(txtParam2.Text);
                    int fontSize = 0, baseLine = 0, imageHeight = 0, imageWidth = 0;
                    fontSize = int.Parse(outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2));//(int)doc.ActiveWindow.Selection.Font.Size;
                    Directory.CreateDirectory(LFilesPath);
                    Directory.CreateDirectory(HFilesPath);
                    doc.ActiveWindow.Selection.StartOf();
                    ulong rgn = 1;
                    while (true)
                    {
                        doc.ActiveWindow.Selection.HomeKey();
                        doc.ActiveWindow.Selection.MoveEnd(x, z);
                        if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                            bBreak = true;
                        doc.ActiveWindow.Selection.Copy();
                        System.Drawing.Size newSize;
                        Bitmap bmp = GetClipboardImage(wordApp, out newSize);
                        string strImagePath = System.IO.Path.GetTempFileName().Replace(".tmp", ".png");
                        double heightRatio = 1.95;
                        double widthRatio = 2.05;
                        bmp.Save(strImagePath);
                        bmp.Dispose();
                        bmp = null;
                        IplImage imagePNG = cvlib.CvLoadImage(strImagePath, cvlib.CV_LOAD_IMAGE_UNCHANGED);

                        IplImage initImage = cvlib.CvCreateImage(cvlib.CvSize((int)(imagePNG.width / widthRatio), (int)(imagePNG.height / heightRatio)), 8, 3);
                        IplImage tmpImage = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 3);
                        IplImage image = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1);
                        IplImage image1 = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1),
                                 image2 = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1),
                                 image3 = cvlib.CvCreateImage(cvlib.CvSize(imagePNG.width, imagePNG.height), 8, 1);
                        cvlib.CvSplit(ref imagePNG, ref image1, ref image2, ref image3, ref image);
                        cvlib.CvMerge(ref image, ref image, ref image, ref tmpImage);
                        cvlib.CvResize(ref tmpImage, ref initImage, 1);
                        //cvlib.CvSaveImage("D:\\a.bmp", ref initImage);
                        // for (int phase = 0; phase < 2; phase++)
                        for (int phase = 1; phase < 2; phase++)
                        {
                            string picOutPath = ((phase == 0) ? HFilesPath : LFilesPath) + Path.GetFileNameWithoutExtension(outputFile) + "_" + counter.ToString() + ".bmp";
                            string picOutPath2 = ((phase == 0) ? HFilesPath : LFilesPath) + Path.GetFileNameWithoutExtension(outputFile) + "_" + counter.ToString() + "_B.bmp";

                            GetDefaultParams(fontSize, phase, out baseLine, out imageHeight);
                            IntPtr finalBaseLine;
                            int initBaseLine = 0;
                            unsafe { finalBaseLine = new IntPtr(&initBaseLine); }
                            //IntPtr p = MapToOriginalSize(ref initImage);
                            IplImage outputImage = cvlib.CvCloneImage(ref initImage);// (IplImage)Marshal.PtrToStructure(p, typeof(IplImage));
                            //cvlib.CvShowImage("d:\\temp1.bmp", ref outputImage); cvlib.CvWaitKey(0);
                            IplImage originalImage = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, imageHeight), 8, 3);
                            IplImage edgeImage = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, imageHeight), 8, 1);
                            IplImage matRnd = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, originalImage.height), 8, 3);
                            IplImage matRndOneChannel = cvlib.CvCreateImage(cvlib.CvSize(outputImage.width, originalImage.height), 8, 1);
                            cvlib.CvSet(ref originalImage, cvlib.cvScalar(0, 0, 0, 0));
                            // MessageBox.Show("1");
                            FindBaseLine(ref outputImage, finalBaseLine, fontSize);
                            AlignBaseLine(ref outputImage, ref originalImage, ref baseLine, initBaseLine);
                            if (baseLine < 0)
                            {
                                cvlib.CvReleaseImage(ref matRnd);
                                cvlib.CvReleaseImage(ref originalImage);
                                cvlib.CvReleaseImage(ref matRndOneChannel);
                                counter--;
                                break;
                            }
                            cvlib.CvNot(ref originalImage, ref originalImage);
                            Binarize(ref originalImage);

                            // MessageBox.Show(System.Windows.Forms.Application.StartupPath);
                            // MessageBox.Show(Directory.GetCurrentDirectory());

                            //Directory.SetCurrentDirectory(System.Windows.Forms.Application.StartupPath);
                            SaveImage(picOutPath, ref originalImage);
                            //cvlib.CvShowImage("2", ref originalImage); cvlib.CvWaitKey(0);
                            //cvlib.CvShowImage("1ee", ref originalImage); cvlib.CvWaitKey(0);

                            cvlib.CvSetImageCOI(ref originalImage, 0);
                            cvlib.CvCanny(ref originalImage, ref edgeImage, 50, 100, 5);
                            //cvlib.CvShowImage("0", ref originalImage); 
                            DestroyImage2(ref originalImage, ref edgeImage, 0.0081, 3);


                            //cvlib.CvShowImage("1", ref edgeImage); cvlib.CvShowImage("2", ref originalImage); cvlib.CvWaitKey(0);
                            if (cmbNoise.SelectedIndex == NOISE_SALT_PEPPER)
                            {
                                cvlib.CvRandArr(ref rgn, ref matRnd, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(0, 0, 0, 0), cvlib.cvScalar(255, 255, 255, 0));
                                unsafe
                                {
                                    byte* pData1 = (byte*)matRnd.imageData;
                                    byte* pData2 = (byte*)originalImage.imageData;

                                    for (int i = 0; i < originalImage.height; i++)
                                        for (int j = 0; j < originalImage.widthStep; j++)
                                        {
                                            if (bOneChannel)
                                            {
                                                if (pData1[i * matRnd.widthStep + j] == param1)
                                                {
                                                    pData2[i * matRnd.widthStep + j] = 0;
                                                    pData2[i * matRnd.widthStep + j + 1] = 0;
                                                    pData2[i * matRnd.widthStep + j + 2] = 0;
                                                }
                                                else
                                                    if (pData1[i * matRnd.widthStep + j] == param2)
                                                    {
                                                        pData2[i * matRnd.widthStep + j] = 255;
                                                        pData2[i * matRnd.widthStep + j + 1] = 255;
                                                        pData2[i * matRnd.widthStep + j + 2] = 255;

                                                    }
                                                j += 2;
                                            }
                                            else
                                            {
                                                if (pData1[i * matRnd.widthStep + j] < param1)
                                                {
                                                    pData2[i * matRnd.widthStep + j] = 0;
                                                }
                                                else
                                                    if (pData1[i * matRnd.widthStep + j] > param2)
                                                    {
                                                        pData2[i * matRnd.widthStep + j] = 255;
                                                    }
                                            }
                                        }
                                }
                            }
                            else
                            {
                                if (bOneChannel)
                                {
                                    cvlib.CvRandArr(ref rgn, ref matRndOneChannel, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(param1, 0, 0, 0), cvlib.cvScalar(param2, 0, 0, 0));
                                    cvlib.CvMerge(ref matRndOneChannel, ref matRndOneChannel, ref matRndOneChannel, ref matRnd);
                                }
                                else
                                    cvlib.CvRandArr(ref rgn, ref matRnd, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(param1, param1, param1, 0), cvlib.cvScalar(param2, param2, param2, 0));

                                cvlib.CvAdd(ref originalImage, ref matRnd, ref originalImage);
                            }
                            Binarize(ref originalImage);
                            if (cmbNoise.SelectedIndex != NOISE_SALT_PEPPER)
                            {
                                param1 = 1;
                                param2 = 255;
                            }
                            {
                                //DestroyImage(ref originalImage, ref matRnd, fontSize / 4, fontSize / 4);
                                // cvlib.CvRandArr(ref rgn, ref matRnd, cvlib.CV_RAND_NORMAL, cvlib.cvScalar(0, 0, 0, 0), cvlib.cvScalar(255, 255, 255, 0));
                                // DestroyImage(ref originalImage, ref matRnd, 7, 7);
                            }
                            cvlib.CvSmooth(ref originalImage, ref originalImage, cvlib.CV_MEDIAN, 3, 3, 0, 0);
                            SaveImage(picOutPath2, ref originalImage);
                            //cvlib.CvShowImage("", ref originalImage); cvlib.CvWaitKey(0);
                            cvlib.CvReleaseImage(ref matRnd);
                            cvlib.CvReleaseImage(ref originalImage);
                            cvlib.CvReleaseImage(ref matRndOneChannel);
                            cvlib.CvReleaseImage(ref edgeImage);
                            //Marshal.StructureToPtr(outputImage, p, false);
                            // ReleaseMap(p);
                        }
                        cvlib.CvReleaseImage(ref image);
                        cvlib.CvReleaseImage(ref image1);
                        cvlib.CvReleaseImage(ref image2);
                        cvlib.CvReleaseImage(ref image3);
                        cvlib.CvReleaseImage(ref imagePNG);
                        cvlib.CvReleaseImage(ref tmpImage);
                        cvlib.CvReleaseImage(ref initImage);
                        File.Delete(strImagePath);
                        doc.ActiveWindow.Selection.MoveDown(x, 1, WdMovementType.wdMove);

                        if (bBreak)
                            break;
                        counter++;
                    }


                    // doc.Close(ref nullObject, ref nullObject, ref nullObject);
                    fileCounter++;
                }
                // wordApp.Quit();
                wordApp = null;
                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);
                if (doc != null) doc.Close(ref nullObject, ref nullObject, ref nullObject);
                if (wordApp != null) wordApp.Quit();
                MessageBox.Show("Finished unsuccessfully");

            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {


            outputFile = txtOutputFile.Text;
            outputFolder = txtOutputFolder.Text;
            inputFile = txtInputFile.Text;
            if (outputFile == "" || outputFile == null)
            {
                MessageBox.Show("Please, Enter or browse Path of output File which contains Word File");
                return;
            }
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object nullObject = Type.Missing;
            Document doc = null;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                doc = wordApp.Documents.Open(outputFile, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                    , ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                    , ref nullObject, ref nullObject
                    , ref nullObject, ref nullObject);
                doc.Activate();

                int counter = 0;
                object x = 5, y = 1, z = WdMovementType.wdExtend;
                doc.ActiveWindow.Selection.StartOf();
                bool bBreak = false;
                doc.ActiveWindow.Selection.HomeKey();
                outputFile = outputFile.Substring(0, txtOutputFile.Text.LastIndexOf(".docx"));
                Directory.CreateDirectory(outputFile);
                outputFolder = outputFile;
                string LFilesPath = outputFolder + "\\" + outputFile.Substring(txtOutputFile.Text.LastIndexOf(".docx") - 2, 2) + "L\\";
                while (true)
                {


                    doc.ActiveWindow.Selection.HomeKey();
                    doc.ActiveWindow.Selection.MoveEnd(x, z);
                    if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                        bBreak = true;
                    string labPath1 = LFilesPath + Path.GetFileNameWithoutExtension(outputFile) + "_" + counter.ToString() + ".lab";
                    string labPath2 = LFilesPath + Path.GetFileNameWithoutExtension(outputFile) + "_" + counter.ToString() + "_B.lab";


                    StreamWriter sw1 = new StreamWriter(labPath1);
                    StreamWriter sw2 = new StreamWriter(labPath2);
                    string content = doc.ActiveWindow.Selection.Text;

                    string[] words = content.Split(' ');
                    string strTmp = "";
                    for (int i = words.Length - 1; i >= 0; i--)
                    {
                        strTmp = words[i];

                        Regex regex = new Regex(@"\W*لله$");
                        Match match = regex.Match(strTmp);
                        while (match != null && match.Length != 0)
                        {
                            strTmp = strTmp.Replace(match.Value, "#");
                            match = match.NextMatch();
                        }
                        // strTmp = strTmp.Replace("لله", "#");
                        strTmp = strTmp.Replace("اً", "@");
                        strTmp = strTmp.Replace("لا", "$");

                        string[] conLine = GetCode(strTmp);

                        for (int j = 0; j < conLine.Length; j++)
                        {
                            if (conLine[j].Contains("nimspace"))
                                continue;
                            sw1.Write("{0}\n", conLine[j]);
                            sw2.Write("{0}\n", conLine[j]);
                        }
                        sw1.Write("{0}\n", "space");
                        sw2.Write("{0}\n", "space");

                    }
                    sw1.Close();
                    sw2.Close();
                    doc.ActiveWindow.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                    if (bBreak)
                        break;
                    counter++;
                }




                doc.Close(ref nullObject, ref nullObject, ref nullObject);
                wordApp.Quit();
                wordApp = null;
                MessageBox.Show("Finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : \n" + ex.Message);
                if (doc != null) doc.Close(ref nullObject, ref nullObject, ref nullObject);
                if (wordApp != null) wordApp.Quit();
                MessageBox.Show("Finished unsuccessfully");

            }

        }

        private void button16_Click(object sender, EventArgs e)
        {
            string strOutFolder = txtPicFolder.Text;
            int fileCounter = int.Parse(txtFileCounter.Text);
            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string strOut = new string(' ', 5 * 1024);
                int newfileCounter = GetOCRText2(of.FileName, strOutFolder, fileCounter, ref strOut);
                string[] files = Directory.GetFiles(strOutFolder);
                foreach (string item in files)
                {
                    if (item.LastIndexOf(".bmp") > 0)
                    {
                        string path = item.Substring(0, item.LastIndexOf(".bmp"));
                        if (File.Exists(path + ".lab"))
                        {
                            linePic.ImageLocation = item;
                            numericUpDown.Enabled = true;
                            numericUpDown.Value = fileCounter;
                            btnDelete.Enabled = btnSaveResult.Enabled = true;
                            break;
                        }
                    }
                }
                txtFileCounter.Text = newfileCounter.ToString();
                OCRResult ocrFrm = new OCRResult();
                ocrFrm.lblOCR.Text = strOut;

                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object nullObject = Type.Missing;
                outputFolder = txtOutputFolder.Text;
                inputFile = txtInputFile.Text;
                string outputFile = Path.GetDirectoryName(of.FileName) + "\\" + Path.GetFileNameWithoutExtension(of.FileName) + ".docx";
                Document doc = wordApp.Documents.Add(ref nullObject, ref nullObject, ref nullObject, ref nullObject);
                doc.Activate();
               
                Paragraph para = doc.Paragraphs.Add();
                doc.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 100);
                para.Range.Text = strOut;
                para.Alignment = WdParagraphAlignment.wdAlignParagraphRight | WdParagraphAlignment.wdAlignParagraphJustify;

                doc.ActiveWindow.Selection.MoveStart();
                doc.ActiveWindow.Selection.HomeKey();
                doc.ActiveWindow.Selection.MoveEnd(WdUnits.wdStory, 1);

                doc.ActiveWindow.Selection.Font.NameBi = "Nazanin";
                doc.ActiveWindow.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                doc.SaveAs2(outputFile);
                doc.Close();
                wordApp.Quit();
                /*
                string outputFile = Path.GetDirectoryName(of.FileName) + "\\" + Path.GetFileNameWithoutExtension(of.FileName) + ".txt";
                StreamWriter sw = new StreamWriter(outputFile);
                sw.Write(strOut);
                sw.Close();*/
                ocrFrm.Show();
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {

        }

        private void button12_Click_1(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            /*string srcPath = @"D:\OCR_DATASET\17.bmp";
            IntPtr arrTexts0 = new IntPtr(0);
            IntPtr arrPics0 = new IntPtr(0);
            IntPtr arrTables0 = new IntPtr(0);
            EngineTextRegion currentText = new EngineTextRegion();
            EnginePicRegion currentPic = new EnginePicRegion();
            EngineTableRegion currentTable = new EngineTableRegion();
            int textsCount = 0;
            int picsCount = 0;
            int tablesCount = 0; 
            PolyPoint tmpPoint = new PolyPoint();
            TableCell tmpCell = new TableCell();

            System.Drawing.Bitmap bmp = new Bitmap(srcPath);
            System.Drawing.Imaging.BitmapData bmpData = bmp.LockBits(new System.Drawing.labtangle(0,0, bmp.Width, bmp.Height), System.Drawing.Imaging.ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            AnalyzeLayout(bmpData.Scan0, bmp.Width, bmp.Height, bmpData.Stride, bmpData.Stride / bmp.Width, out arrTexts0, out textsCount, out arrPics0, out picsCount, out arrTables0, out tablesCount);
            
            TextRegion[] arrTexts = new TextRegion[textsCount];
            PicRegion[] arrPics = new PicRegion[picsCount];
            TableRegion[] arrTables = new TableRegion[tablesCount];
            for (int i = 0; i < textsCount; i++)
            {
                IntPtr pCurrentText = new IntPtr(arrTexts0.ToInt32() + i * Marshal.SizeOf(currentText));
                currentText = (EngineTextRegion)Marshal.PtrToStructure(pCurrentText, typeof(EngineTextRegion));
                arrTexts[i].boundRect = currentText.boundRect;
                arrTexts[i].boundPoly.area = currentText.boundPoly.area;
                arrTexts[i].boundPoly.pnts = new PolyPoint[currentText.boundPoly.pntCount];
                for (int j = 0; j < currentText.boundPoly.pntCount; j++)
                {
                    arrTexts[i].boundPoly.pnts[j] = (PolyPoint)Marshal.PtrToStructure(new IntPtr(currentText.boundPoly.pnts.ToInt32() + Marshal.SizeOf(tmpPoint) * j), typeof(PolyPoint));
                }
                arrTexts[i].fontName = currentText.fontName;
                arrTexts[i].ocrText = currentText.ocrText;
                arrTexts[i].fontSize = currentText.fontSize;
            }
            for (int i = 0; i < picsCount; i++)
            {

                IntPtr pCurrentPic = new IntPtr(arrPics0.ToInt32() + i * Marshal.SizeOf(currentPic));
                currentPic = (EnginePicRegion)Marshal.PtrToStructure(pCurrentPic, typeof(EnginePicRegion));
                arrPics[i].boundRect = currentPic.boundRect;
                arrPics[i].boundPoly.area = currentPic.boundPoly.area;
                arrPics[i].boundPoly.pnts = new PolyPoint[currentPic.boundPoly.pntCount];

                for (int j = 0; j < currentPic.boundPoly.pntCount; j++)
                {
                    arrPics[i].boundPoly.pnts[j] = (PolyPoint)Marshal.PtrToStructure(new IntPtr(currentPic.boundPoly.pnts.ToInt32() + Marshal.SizeOf(tmpPoint) * j), typeof(PolyPoint));
                }
            }
            for (int i = 0; i < tablesCount; i++)
            {

                IntPtr pCurrentTable = new IntPtr(arrTables0.ToInt32() + i * Marshal.SizeOf(currentTable));
                currentTable = (EngineTableRegion)Marshal.PtrToStructure(pCurrentTable, typeof(EngineTableRegion));

                arrTables[i].boundRect = currentTable.boundRect;
                arrTables[i].columnsCount = currentTable.columnsCount;
                arrTables[i].rowsCount = currentTable.rowsCount;
                arrTables[i].cells = new TableCell[currentTable.cellsCount];
                for (int j = 0; j < currentTable.cellsCount; j++)
                {
                    arrTables[i].cells[j] = (TableCell)Marshal.PtrToStructure(new IntPtr(currentTable.Cells.ToInt32() + Marshal.SizeOf(tmpCell) * j), typeof(TableCell));
                }

            }
            IntPtr tmpPtr = IntPtr.Zero; 
            IntPtr pArrTextRegions = Marshal.AllocHGlobal(arrTexts.Length * Marshal.SizeOf(currentText));
            //MessageBox.Show(pArrTextRegions.ToInt32().ToString("x"));
            for (int i = 0; i < arrTexts.Length; i++)
            {
                currentText.boundRect = arrTexts[i].boundRect;
                currentText.fontName = arrTexts[i].fontName;
                currentText.fontSize = arrTexts[i].fontSize;
                currentText.ocrText = arrTexts[i].ocrText;
                currentText.boundPoly.area = arrTexts[i].boundPoly.area  ;
                currentText.boundPoly.pntCount = arrTexts[i].boundPoly.pnts.Length;
                currentText.boundPoly.pnts = Marshal.AllocHGlobal(currentText.boundPoly.pntCount * Marshal.SizeOf(tmpPoint));
                for (int j = 0; j < currentText.boundPoly.pntCount; j++)
                {
                    tmpPtr = new IntPtr(currentText.boundPoly.pnts.ToInt32() + j * Marshal.SizeOf(tmpPoint));
                    Marshal.StructureToPtr(arrTexts[i].boundPoly.pnts[j], tmpPtr, false);                     
                }
                Marshal.StructureToPtr(currentText, new IntPtr(pArrTextRegions.ToInt32() + i * Marshal.SizeOf(currentText)), false);
            }
            IntPtr pArrPicRegions = Marshal.AllocHGlobal(arrPics.Length * Marshal.SizeOf(currentPic));
            for (int i = 0; i < arrPics.Length; i++)
            {
                currentPic.boundRect = arrPics[i].boundRect;
                currentPic.boundPoly.area = arrPics[i].boundPoly.area;
                currentPic.boundPoly.pntCount = arrPics[i].boundPoly.pnts.Length;
                currentPic.boundPoly.pnts = Marshal.AllocHGlobal(currentPic.boundPoly.pntCount * Marshal.SizeOf(tmpPoint));
                for (int j = 0; j < currentPic.boundPoly.pntCount; j++)
                {
                    tmpPtr = new IntPtr(currentPic.boundPoly.pnts.ToInt32() + j * Marshal.SizeOf(tmpPoint));
                    Marshal.StructureToPtr(arrPics[i].boundPoly.pnts[j], tmpPtr, false);                     
                }
                Marshal.StructureToPtr(currentPic, new IntPtr(pArrPicRegions.ToInt32() + i * Marshal.SizeOf(currentPic)), false);
            }
            IntPtr pArrTablesRegions = Marshal.AllocHGlobal(arrTables.Length * Marshal.SizeOf(currentTable));
            //MessageBox.Show(pArrTablesRegions.ToInt32().ToString("x"));
            for (int i = 0; i < arrTables.Length; i++)
            {

                currentTable.boundRect = arrTables[i].boundRect ;
                currentTable.columnsCount = arrTables[i].columnsCount;
                currentTable.rowsCount = arrTables[i].rowsCount;
                currentTable.cellsCount = arrTables[i].cells.Length;
                currentTable.Cells = Marshal.AllocHGlobal(currentTable.cellsCount * Marshal.SizeOf(tmpCell));                
                for (int j = 0; j < currentTable.cellsCount; j++)
                {
                    tmpPtr = new IntPtr(currentTable.Cells.ToInt32() + j * Marshal.SizeOf(tmpCell));
                    Marshal.StructureToPtr(arrTables[i].cells[j], tmpPtr, false);                      
                }
                Marshal.StructureToPtr(currentTable, new IntPtr(pArrTablesRegions.ToInt32() + i * Marshal.SizeOf(currentTable)), false);
            }
            ProccessLayout(bmpData.Scan0, bmp.Width, bmp.Height, bmpData.Stride, bmpData.Stride / bmp.Width, ref pArrTextRegions, ref textsCount, ref pArrPicRegions, ref picsCount, ref pArrTablesRegions, ref tablesCount);
            TextRegion[] retArrTexts = new TextRegion[textsCount];
            PicRegion[] retArrPics = new PicRegion[picsCount];
            TableRegion[] retArrTables = new TableRegion[tablesCount];
            for (int i = 0; i < textsCount; i++)
            {
                IntPtr pCurrentText = new IntPtr(pArrTextRegions.ToInt32() + i * Marshal.SizeOf(currentText));
                currentText = (EngineTextRegion)Marshal.PtrToStructure(pCurrentText, typeof(EngineTextRegion));
                retArrTexts[i].boundRect = currentText.boundRect;
                retArrTexts[i].boundPoly.area = currentText.boundPoly.area;
                retArrTexts[i].boundPoly.pnts = new PolyPoint[currentText.boundPoly.pntCount];
                for (int j = 0; j < currentText.boundPoly.pntCount; j++)
                {
                    retArrTexts[i].boundPoly.pnts[j] = (PolyPoint)Marshal.PtrToStructure(new IntPtr(currentText.boundPoly.pnts.ToInt32() + Marshal.SizeOf(tmpPoint) * j), typeof(PolyPoint));
                }
                retArrTexts[i].fontName = currentText.fontName;
                retArrTexts[i].ocrText = currentText.ocrText;
                retArrTexts[i].fontSize = currentText.fontSize;
            }
            for (int i = 0; i < picsCount; i++)
            {

                IntPtr pCurrentPic = new IntPtr(pArrPicRegions.ToInt32() + i * Marshal.SizeOf(currentPic));
                currentPic = (EnginePicRegion)Marshal.PtrToStructure(pCurrentPic, typeof(EnginePicRegion));
                retArrPics[i].boundRect = currentPic.boundRect;
                retArrPics[i].boundPoly.area = currentPic.boundPoly.area;
                retArrPics[i].boundPoly.pnts = new PolyPoint[currentPic.boundPoly.pntCount];

                for (int j = 0; j < currentPic.boundPoly.pntCount; j++)
                {
                    retArrPics[i].boundPoly.pnts[j] = (PolyPoint)Marshal.PtrToStructure(new IntPtr(currentPic.boundPoly.pnts.ToInt32() + Marshal.SizeOf(tmpPoint) * j), typeof(PolyPoint));
                }
            }
            for (int i = 0; i < tablesCount; i++)
            {

                IntPtr pCurrentTable = new IntPtr(pArrTablesRegions.ToInt32() + i * Marshal.SizeOf(currentTable));
                currentTable = (EngineTableRegion)Marshal.PtrToStructure(pCurrentTable, typeof(EngineTableRegion));

                retArrTables[i].boundRect = currentTable.boundRect;
                retArrTables[i].columnsCount = currentTable.columnsCount;
                retArrTables[i].rowsCount = currentTable.rowsCount;
                retArrTables[i].cells = new TableCell[currentTable.cellsCount];
                for (int j = 0; j < currentTable.cellsCount; j++)
                {
                    retArrTables[i].cells[j] = (TableCell)Marshal.PtrToStructure(new IntPtr(currentTable.Cells.ToInt32() + Marshal.SizeOf(tmpCell) * j), typeof(TableCell));
                }

            }
           int x = 4;
           x++;*/
            FolderBrowserDialog fb = new FolderBrowserDialog();
            fb.ShowNewFolderButton = true;
            fb.SelectedPath = txtPicFolder.Text; //@"D:\OCR_DATASET\newDataset";
            //fb.RootFolder = Environment.SpecialFolder.labent;
            //if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtPicFolder.Text = fb.SelectedPath;
                string[] files = Directory.GetFiles(fb.SelectedPath);
                int max = -1;
                foreach (string item in files)
                {
                    if (item.LastIndexOf(".bmp") > 0)
                    {
                        int number = int.Parse(item.Substring(item.LastIndexOf("\\") + 1, item.LastIndexOf(".bmp") - item.LastIndexOf("\\") - 1));
                        if (number > max)
                        {
                            max = number;
                        }
                    }
                }
                max++;
                txtFileCounter.Text = max.ToString();
                foreach (string item in files)
                {
                    if (item.LastIndexOf(".bmp") > 0)
                    {
                        string path = item.Substring(0, item.LastIndexOf(".bmp"));
                        if (File.Exists(path + ".lab"))
                        {
                            linePic.ImageLocation = item;
                            numericUpDown.Enabled = true;
                            numericUpDown.Value = int.Parse(item.Substring(item.LastIndexOf("\\") + 1, item.LastIndexOf(".bmp") - item.LastIndexOf("\\") - 1));
                            btnDelete.Enabled = btnSaveResult.Enabled = true;
                            break;
                        }
                    }
                }
            }
        }

        private void linePic_LocationChanged(object sender, EventArgs e)
        {

        }
        int[] arrHashCodes = new int[172];
        string[] arrModelSymbols = {
			"a",           
			"space",       
			"_b",          
			"_b_",         
			"_i",          
			"_i_",         
			"_l",          
			"_l_",         
			"_n",          
			"_n_",         
			"_p",          
			"_se",         
			"_se_",        
			"_t",          
			"_t_",         
			"_y",          
			"_y_",         
			"a_",          
			"h",           
			"hamze",       
			"r",           
			"v",           
			"vh",          
			"ze",          
			"_ch",         
			"_ch_",        
			"_f",          
			"_f_",         
			"_g",          
			"_g_",         
			"_gs",         
			"_gs_",        
			"_h",          
			"_h_",         
			"_he",         
			"_he_",        
			"_j",          
			"_j_",         
			"_k",          
			"_k_",         
			"_m",          
			"_m_",         
			"_p_",         
			"_q",          
			"_q_",         
			"_qaf",        
			"_qaf_",       
			"_ta",         
			"_ta_",        
			"_x",          
			"_x_",         
			"_za",         
			"_za_",        
			"aa",          
			"aa_",         
			"ch",          
			"ch_",         
			"d",           
			"d_",          
			"gs",          
			"gs_",         
			"h_",          
			"he",          
			"he_",         
			"i",           
			"j",           
			"j_",          
			"l",           
			"m",           
			"n",           
			"q",           
			"q_",          
			"qaf",         
			"r_",          
			"ta",          
			"ta_",         
			"v_",          
			"vh_",         
			"x",           
			"x_",          
			"y",           
			"za",          
			"za_",         
			"zal",         
			"zal_",        
			"ze_",         
			"zh",          
			"zh_",         
			"_sad",        
			"_sad_",       
			"_sh",         
			"_sh_",        
			"_sin",        
			"_sin_",       
			"_zad",        
			"_zad_",       
			"b",           
			"f",           
			"f_",          
			"g",           
			"i_",          
			"k",           
			"l_",          
			"m_",          
			"n_",          
			"p",           
			"qaf_",        
			"se",          
			"t",           
			"y_",          
			"b_",          
			"g_",          
			"k_",          
			"p_",          
			"se_",         
			"t_",          
			"allah",       
			"sad",         
			"sad_",        
			"sh",          
			"sh_",         
			"sin",         
			"sin_",        
			"zad",         
			"zad_",        
			"aht",         
			"aht_",        
			"ahb",         
			"ahb_",        
			"atan",        
			"atan_",       
			"hh",          
			"hh_",         
			
			"zero",        
			"one",         
			"two",         
			"three",       
			"four",        
			"five",        
			"six",         
			"seven",       
			"eight",       
			"nine",        
			
			"la",          
			"la_",         
			
			"ako",         
			"akc",         
			"paro",        
			"parc",        
			"vir",         
			"simi",        
			"tdot",        
			"dot",         
			"slash",       
			"darsad",      
			"menha",       
			"under",       
			"que",         
			"bro",         
			"brc",         
			"kama",        
			"taqsim",      
			"zarb",        
			"eq",          
			"star",        
			"wonder",      
			"plus",  
			"_gume",
			"gume_",
			"tanvin",
			"Engsimi",
			"nimspace"      
		};
        string[] arrSymbols = {

			"ا",
			" ",
			"ب",
			"ب",
			"ئ",
			"ئ",
			"ل",
			"ل",
			"ن",
			"ن",
			"پ",
			"ث",
			"ث",
			"ت",
			"ت",
			"ي",
			"ي",
			"ا",
			"ه",
			"ء",
			"ر",
			"و",
			"ؤ",
			"ز",
			"چ",
			"چ",
			"ف",
			"ف",
			"گ",
			"گ",
			"ع",
			"ع",
			"ه",
			"ه",
			"ح",
			"ح",
			"ج",
			"ج",
			"ک",
			"ک",
			"م",
			"م",
			"پ",
			"غ",
			"غ",
			"ق",
			"ق",
			"ط",
			"ط",
			"خ",
			"خ",
			"ظ",
			"ظ",
			"آ",
			"آ",
			"چ",
			"چ",
			"د",
			"د",
			"ع",
			"ع",
			"ه",
			"ح",
			"ح",
			"ئ",
			"ج",
			"ج",
			"ل",
			"م",
			"ن",
			"غ",
			"غ",
			"ق",
			"ر",
			"ط",
			"ط",
			"و",
			"ؤ",
			"خ",
			"خ",
			"ي",
			"ظ",
			"ظ",
			"ذ",
			"ذ",
			"ز",
			"ژ",
			"ژ",
			"ص",
			"ص",
			"ش",
			"ش",
			"س",
			"س",
			"ض",
			"ض",
			"ب",
			"ف",
			"ف",
			"گ",
			"ئ",
			"ک",
			"ل",
			"م",
			"ن",
			"پ",
			"ق",
			"ث",
			"ت",
			"ي",
			"ب",
			"گ",
			"ک",
			"پ",
			"ث",
			"ت",
			"لله",
			"ص",
			"ص",
			"ش",
			"ش",
			"س",
			"س",
			"ض",
			"ض",
			"أ",
			"أ",
			"إ",
			"إ",
			"اً",
			"اً",
			"ة",
			"ة",
			
			"0",
			"1",
			"2",
			"3",
			"4",
			"5",
			"6",
			"7",
			"8",
			"9",
			
			"لا",
			"لا",
			
			"{",
			"}",
			"(",
			")",
			"،",
			"؛",
			":",
			".",
			"/",
			"%",
			"-",
			"_",
			"?",
			"[",
			"]",
			",",
			"÷",
			"×",
			"=",
			"*",
			"!",
			"+",
			"»",
			"«",
			"\"",
			";",
			"‌"          

		};

        void CreateHashArray()
        {
            for (int i = 0; i < 172; i++)
            {
                for (int j = 0; j < arrModelSymbols[i].Length; j++)
                    arrHashCodes[i] |= (arrModelSymbols[i][j] << j * 8);
            }
        }
        string model(string p)
        {
            string pp = "";
            int hashCode = 0;
            bool bBeFound = false;
            for (int i = 0; i < p.Length; i++)
                hashCode |= (p[i] << i * 8);
            for (int i = 0; i < 172; i++)
            {
                if (hashCode == arrHashCodes[i])
                {
                    pp = arrSymbols[i];
                    bBeFound = true;
                    break;
                }
            }

            if (!bBeFound)
                pp = p;
            return pp;
        }
        string reverse_model(string p)
        {
            string pp = "";
            bool bBeFound = false;
            for (int i = 0; i < 172; i++)
            {
                if (p == arrSymbols[i])
                {
                    pp = arrModelSymbols[i];
                    bBeFound = true;
                    break;
                }
            }

            if (!bBeFound)
                pp = p;
            return pp;
        }
        void SaveResults(string filePath, string text)
        {


            StreamWriter sr = new StreamWriter(filePath);


            int i = 0;

            Regex regex = new Regex(@"\W*لله$");
            Match match = regex.Match(text);
            while (match != null && match.Length != 0)
            {
                text = text.Replace(match.Value, "#");
                match = match.NextMatch();
            }
            // strTmp = strTmp.Replace("لله", "#");
            text = text.Replace("اً", "@");
            text = text.Replace("لا", "$");

            string[] conLine = GetCode(text);

            for (int j = 0; j < conLine.Length; j++)
            {
                //if (conLine[j].Contains("nimspace"))
                //    continue;
                sr.Write("{0}\n", conLine[j]);
            }

            sr.Close();


        }
        void ReadResults2(string filePath, ref Dictionary<string, int> arrPhonemes)
        {

            List<string> words = new List<string>();
            string buff;

            StreamReader sr = new StreamReader(filePath);


            string matn = "";
            string lastText = "";
            while (!sr.EndOfStream)
            {

                buff = sr.ReadLine();
                if (buff == "") break;

                string[] temp = buff.Split(' ');
                if(temp.Length > 2)
                    words.Add(temp[2]);
                else
                    words.Add(temp[0]);

            }
            sr.Close();
            for (int i = 0; i < words.Count; i++)
            {
                if (arrPhonemes.Keys.Contains(words[i].ToLower()))
                {
                    arrPhonemes[words[i]]++;
                }
                else
                {
                    arrPhonemes.Add(words[i].ToLower(), 1);
                }
            }


        }
        string ReadResults(string filePath)
        {

            List<string> words = new List<string>();
            string buff;

            StreamReader sr = new StreamReader(filePath);


            string matn = "";
            string lastText = "";
            while (!sr.EndOfStream)
            {

                buff = sr.ReadLine();
                if (buff == "") break;

                string[] temp = buff.Split(' ');
                if(temp.Length > 2)
                    words.Add(temp[2]);
                else
                    words.Add(temp[0]);

            }
            sr.Close();


            string tmpStr;
            int i = 0;
            for (i = words.Count - 1; i >= 0; i--)
            {

                string farsi = model(words[i]);
                matn += farsi;

            }
            string strOut = "";
            for (i = 0; i < matn.Length; i++)
            {
                if (matn[i] == 10)
                    strOut += " ";
                else
                    strOut += matn[i].ToString();

            }
            Regex regex = new Regex("[0-9]+");
            Match match = regex.Match(matn);
            while (match != null && match.Length != 0)
            {
                char[] tmpWord = match.Value.ToArray();
                Array.Reverse(tmpWord);
                matn = matn.Replace(match.Value, new string(tmpWord));
                match = match.NextMatch();
            }

            return matn;


        }
        decimal prevFile;
        public string ReplaceNumbers(string WORD1)
        {          
            //Regex regex = new Regex(@"([0-9]+[ ])+[0-9]$");
            Regex regex = new Regex(@"([0-9]+[ ]+)+");
            StringBuilder tmpStr = new StringBuilder(WORD1);
            Match match = regex.Match(WORD1);
            int spcaeCounter = 0;
            while (match != null && match.Length != 0)
            {
                if (match.Length == 1)
                    continue;
                char[] tmpWord = match.Value.ToArray();
                
                Array.Reverse(tmpWord);

                string newStr = new string(tmpWord);
                newStr = newStr.Replace(" ", "");
                newStr = newStr + " ";
                tmpStr = tmpStr.Replace(match.Value, newStr, match.Index - spcaeCounter, match.Value.Length/* + (match.Value.Length - newStr.Length) * 2*/);
                spcaeCounter += match.Value.Length - newStr.Length;
                match = match.NextMatch();
            }
            WORD1 = tmpStr.ToString();
            regex = new Regex(@"([0-9]+[ ])[0-9]");
            match = regex.Match(WORD1);
            spcaeCounter = 0;
            while (match != null && match.Length != 0)
            {
                if (match.Length == 1)
                    continue;
                string[] tmpWord = match.Value.Split(new char[] { ' ' });
                string newStr = "";
                for (int i = tmpWord.Length - 1; i >= 0; i--)
                {
                    newStr += tmpWord[i];
                }
                newStr = newStr.Replace(" ", "");
                //tmpStr = tmpStr.Replace(match.Value, newStr, match.Index, match.Value.Length);
                tmpStr = tmpStr.Replace(match.Value, newStr, match.Index - spcaeCounter, match.Value.Length/* + (match.Value.Length - newStr.Length) * 2*/);
                spcaeCounter += match.Value.Length - newStr.Length;
                match = match.NextMatch();
            }
            /*for (int i = 1; i < tmpStr.Length - 1; i++)
            {
                if (tmpStr[i + 1] == ' ')
                {
                    if (tmpStr[i] == '{' || tmpStr[i] == '[' || tmpStr[i] == '(' )
                    {
                        tmpStr.Remove(i + 1, 1);
                        i--;
                        continue;
                    }

                }
                if (tmpStr[i - 1] == ' ')
                {
                    if (tmpStr[i] == '}' || tmpStr[i] == ']' || tmpStr[i] == ')' || tmpStr[i] == ',' || tmpStr[i] == '.' ||
                        tmpStr[i] == ':' || tmpStr[i] == '؛' || tmpStr[i] == '،' || tmpStr[i] == '؟' || tmpStr[i] == '!')
                    {
                        tmpStr.Remove(i - 1,1);
                        i--;
                        continue;
                    }

                }
            }*/
            return tmpStr.ToString();
        }
        private void numericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (bTextChanged)
            {
                if (MessageBox.Show("do you want to save changes?", "", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    
                    string labelFilePath1 = txtPicFolder.Text + "\\" + prevFile.ToString() + ".lab";
                    SaveResults(labelFilePath1, richTextBox1.Text);
                    bTextChanged = false;
                }
            }
            prevFile = numericUpDown.Value;
            string labelFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".lab";
            string picFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".bmp";
            if (File.Exists(labelFilePath))
            {
                string context = ReadResults(labelFilePath);
                richTextBox1.Text = ReplaceNumbers(context);
                if (context != richTextBox1.Text)
                {
                    bTextChanged = true;
                }
                else
                {
                    bTextChanged = false;                        
                }
                linePic.ImageLocation = picFilePath;
            }
            else
                numericUpDown.Value++;
        }

        private void numericUpDown_EnabledChanged(object sender, EventArgs e)
        {
            string labelFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".lab";
            string picFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".bmp";
            prevFile = numericUpDown.Value;
            if (File.Exists(labelFilePath))
            {
                string context = ReadResults(labelFilePath);
                
                richTextBox1.Text = ReplaceNumbers(context);
                if (context != richTextBox1.Text)
                {
                    bTextChanged = true;
                }
                else
                {
                    bTextChanged = false;
                }
                linePic.ImageLocation = picFilePath;
            }
            else
                numericUpDown.Value++;

        }

        private void btnSaveResult_Click(object sender, EventArgs e)
        {
            string labelFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".lab";
            SaveResults(labelFilePath, richTextBox1.Text);
            bTextChanged = false;
            
        }
        bool bTextChanged;
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            bTextChanged = true;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            string labelFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".lab";
            string feaFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".fea";
            string bmpFilePath = txtPicFolder.Text + "\\" + numericUpDown.Value.ToString() + ".bmp";
            File.Delete(labelFilePath);
            File.Delete(feaFilePath);
            File.Delete(bmpFilePath);
            if (numericUpDown.Value > 0)
                numericUpDown.Value--;
            else
                numericUpDown.Value = 1;
            MessageBox.Show("deletion is successfully");
        }
        string iconPath;
        private void button17_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                iconPath = of.FileName;
                if (File.Exists(iconPath))
                {
                    pictureBox1.ImageLocation = iconPath;
                }
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (File.Exists(iconPath))
            {
                pictureBox1.Image.Save(iconPath);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (File.Exists(iconPath))
            {
                ColorDialog cd = new ColorDialog();
                if (cd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Color selectedColor = cd.Color;
                    string strImagePath = System.IO.Path.GetTempFileName().Replace(".tmp", ".png");
                    openCV.IplImage srcImg = openCV.cvlib.CvLoadImage(iconPath, cvlib.CV_LOAD_IMAGE_UNCHANGED);
                     IplImage image1 = cvlib.CvCreateImage(cvlib.CvSize(srcImg.width, srcImg.height), 8, 1);
                     IplImage image2 = cvlib.CvCreateImage(cvlib.CvSize(srcImg.width, srcImg.height), 8, 1),
                              image3 = cvlib.CvCreateImage(cvlib.CvSize(srcImg.width, srcImg.height), 8, 1),
                              image4 = cvlib.CvCreateImage(cvlib.CvSize(srcImg.width, srcImg.height), 8, 1);
                    cvlib.CvSplit(ref srcImg, ref image1, ref image2, ref image3, ref image4);
                    cvlib.CvSet(ref image1, cvlib.cvScalarAll(selectedColor.B));
                    cvlib.CvSet(ref image2, cvlib.cvScalarAll(selectedColor.G));
                    cvlib.CvSet(ref image3, cvlib.cvScalarAll(selectedColor.R));
                    cvlib.CvMerge(ref image1, ref image2, ref image3, ref image4, ref srcImg);
                    SaveImage(strImagePath, ref srcImg);
                    cvlib.CvReleaseImage(ref image1);
                    cvlib.CvReleaseImage(ref image2);
                    cvlib.CvReleaseImage(ref image3);
                    cvlib.CvReleaseImage(ref image4);
                    cvlib.CvReleaseImage(ref srcImg);
                    pictureBox1.ImageLocation = "";
                    pictureBox1.ImageLocation = strImagePath;
                }
            }
            else
            {
                MessageBox.Show("please, open an image");
            }
        }

        int GetNormalHeight(int fontSize)
        {
            int retVal = 0;
            switch (fontSize)
            {
                case 10:
                    retVal = 230;
                    break;
                case 12:
                    retVal = 245;
                    break;
                case 13:
                    retVal = 260;
                    break;
                case 14:
                    retVal = 275;
                    break;
                case 15:
                    retVal = 290;
                    break;
                case 16:
                    retVal = 305;
                    break;
                case 17:
                    retVal = 320;
                    break;
                case 18:
                    retVal = 335;
                    break;
                case 22:
                    retVal = 380;
                    break;
                case 26:
                    retVal = 435;
                    break;
                case 30:
                    retVal = 490;
                    break;
            }
            return retVal;
        }
        private void button19_Click(object sender, EventArgs e)
        {
            string strOutFolder = txtPicFolder.Text;
            int fileCounter = int.Parse(txtFileCounter.Text);
            OpenFileDialog of = new OpenFileDialog();
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object nullObject = Type.Missing;
            Document doc = null;

            string secondFont = "Nazanin";
            string firstFont = "Nazanin";
            FontDialog fontDlg = new FontDialog();
            //if (fontDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                secondFont = "b nazanin";
            }
            //if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            string[] files = Directory.GetFiles(txtOutputFolder.Text , "*.bmp");
            for (int m = 0; m < files.Length; m++ )
            {
                outputFile = Path.GetDirectoryName(files[m]) + "\\" + Path.GetFileNameWithoutExtension(files[m]) + ".docx";
                string strOut = new string(' ', 5 * 1024);
                fileCounter = int.Parse(txtFileCounter.Text);
                int newfileCounter = GetOCRText2(files[m], strOutFolder, fileCounter, ref strOut);
                txtFileCounter.Text = newfileCounter.ToString();
                try
                {
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    doc = wordApp.Documents.Open(outputFile, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject
                        , ref nullObject, ref nullObject);
                    doc.Activate();

                    int counter = 0;
                    object x = 5, y = 1, z = WdMovementType.wdExtend;
                    doc.ActiveWindow.Selection.StartOf();
                    bool bBreak = false;
                    doc.ActiveWindow.Selection.HomeKey();
                    outputFile = outputFile.Substring(0, outputFile.LastIndexOf(".docx"));
                    //Directory.CreateDirectory(outputFile);
                    outputFolder = outputFile;
                    string LFilesPath = strOutFolder + "\\";
                    int fontSize = 0;
                    fontSize = 10;
                    Regex regex1 = new Regex("[0-9]+");
                    while (true)
                    {
                        string strImagePath = "";

                        doc.ActiveWindow.Selection.HomeKey();
                        doc.ActiveWindow.Selection.MoveEnd(x, z);
                        if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                            bBreak = true;

                        string text = doc.ActiveWindow.Selection.Text;

                        firstFont = doc.ActiveWindow.Selection.Font.Name;
                        Match match1 = regex1.Match(text);
                        if (match1 != null && match1.Length != 0)
                            doc.ActiveWindow.Selection.Font.NameBi = secondFont;
                        doc.ActiveWindow.Selection.Copy();
                        System.Drawing.Size newSize;
                        Bitmap bmp = GetClipboardImage(wordApp, out newSize);
                        strImagePath = System.IO.Path.GetTempFileName().Replace(".tmp", ".png");
                        double heightRatio = 1.95;
                        double widthRatio = 2.05;
                        bmp.Save(strImagePath);

                        if (bmp.Height > 1.5 * GetNormalHeight(fontSize))
                        {

                            doc.ActiveWindow.Selection.Font.NameBi = firstFont;
                            doc.ActiveWindow.Selection.Font.Name = firstFont;
                            bmp.Dispose();
                            bmp = null;
                            if (File.Exists(strImagePath)) File.Delete(strImagePath);
                            doc.ActiveWindow.Selection.MoveDown(x, 1, WdMovementType.wdMove);
                            if (doc.ActiveWindow.Selection.Bookmarks.Exists(@"\EndOfDoc"))
                                break;
                            counter++;
                            continue;
                        }

                        doc.ActiveWindow.Selection.Font.Name = firstFont;
                        doc.ActiveWindow.Selection.Font.NameBi = firstFont;
                        bmp.Dispose();
                        bmp = null;
                        if (File.Exists(strImagePath)) File.Delete(strImagePath);

                        string labPath1 = LFilesPath + /*Path.GetFileNameWithoutExtension(outputFile) + "_" + */(counter + fileCounter).ToString() + ".lab";


                        StreamWriter sw1 = new StreamWriter(labPath1);
                        string content = doc.ActiveWindow.Selection.Text;

                        string[] words = content.Split(' ');
                        string strTmp = "";
                        for (int i = words.Length - 1; i >= 0; i--)
                        {
                            strTmp = words[i];

                            Regex regex = new Regex(@"\W*لله$");
                            Match match = regex.Match(strTmp);
                            while (match != null && match.Length != 0)
                            {
                                strTmp = strTmp.Replace(match.Value, "#");
                                match = match.NextMatch();
                            }
                            // strTmp = strTmp.Replace("لله", "#");
                            strTmp = strTmp.Replace("اً", "@");
                            strTmp = strTmp.Replace("لا", "$");

                            string[] conLine = GetCode(strTmp);

                            for (int j = 0; j < conLine.Length; j++)
                            {
                                if (conLine[j].Contains("nimspace"))
                                    continue;
                                sw1.Write("{0}\n", conLine[j]);
                            }
                            sw1.Write("{0}\n", "space");

                        }
                        sw1.Close();
                        doc.ActiveWindow.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
                        if (bBreak)
                            break;
                        counter++;
                    }




                    doc.Close(ref nullObject, ref nullObject, ref nullObject);
                    wordApp.Quit();
                    wordApp = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error : \n" + ex.Message);
                    if (doc != null) doc.Close(ref nullObject, ref nullObject, ref nullObject);
                    if (wordApp != null) wordApp.Quit();
                    MessageBox.Show("Finished unsuccessfully");

                }
            }
                    MessageBox.Show("Finished");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles(txtPicFolder.Text,"*.lab");
            string[]symbols = Directory.GetFiles(@"D:\fontSize\Units","*.bmp");
            Dictionary<string, int> arrPhonemes = new Dictionary<string,int>();
            for (int i = 0; i < symbols.Length; i++)
			{
                arrPhonemes.Add(Path.GetFileNameWithoutExtension(symbols[i]).ToLower(), 1);
            }
            for (int i = 0; i < files.Length; i++)
			{
			 
                ReadResults2(files[i], ref arrPhonemes);
			}
            List<string> missedList = new List<string>();
            for (int i = 0; i < arrPhonemes.Count; i++)
            {
                if (arrPhonemes.Values.ElementAt(i) == 1)
                {
                    missedList.Add(arrPhonemes.Keys.ElementAt(i));
                }
            }
        }

        void ReadResults3(string filePath, ref Dictionary<string, Dictionary<int,int>> arrPhonemes)
        {
            string unitsPath = @"D:\fontSize\RealUnits\";
            List<string> words = new List<string>();
            string buff;

            StreamReader sr = new StreamReader(filePath);
            string picPath = filePath.Replace(".lab", ".bmp");
            string matn = "";
            string lastText = "";
            IplImage image = openCV.cvlib.CvLoadImage(picPath, cvlib.CV_LOAD_IMAGE_UNCHANGED);
            string newPath = "";
            while (!sr.EndOfStream)
            {

                buff = sr.ReadLine();
                if (buff == "") break;

                string[] temp = buff.Split(' ');
                int start = int.Parse(temp[0]) / 100000 ;
                int end = int.Parse(temp[1]) / 100000 ;
                string phonem = temp[2];
                int len = end - start;
                if (arrPhonemes.Keys.Contains(phonem))
                {
                    Dictionary<int,int> phonemSamples= arrPhonemes[phonem];
                    if (phonemSamples.ContainsKey(len))
                    {                        
                        arrPhonemes[phonem][len]++;
                    }
                    else if (phonemSamples.ContainsKey(len - 1))
                    {
                        len--;
                        arrPhonemes[phonem][len]++;
                    }
                    else if (phonemSamples.ContainsKey(len + 1))
                    {
                        len++;
                        arrPhonemes[phonem][len]++;
                    }
                    else
                    {
                        arrPhonemes[phonem].Add(len, 1);
                    }

                    if (arrPhonemes[phonem][len] <= 10)
                    {
                        if (end < image.width)
                        {
                            int lenIndex = 0;
                            for (int i = 0; i < arrPhonemes[phonem].Keys.Count; i++)
                            {
                                if (arrPhonemes[phonem].Keys.ElementAt(i) == len)
                                {
                                    lenIndex = i;
                                    break;
                                }
                            }
                            newPath = unitsPath + phonem + "-" + lenIndex.ToString()+"__" + arrPhonemes[phonem][len].ToString() + ".bmp";
                            cvlib.CvSetImageROI(ref image, cvlib.cvRect(start, 0, end - start, image.height));
                            IplImage imgDst = cvlib.CvCreateImage(cvlib.CvSize(end - start, image.height), 8, 1);
                            cvlib.CvCopy(ref image, ref imgDst);
                            cvlib.CvSaveImage(newPath, ref imgDst);
                            cvlib.CvReleaseImage(ref imgDst);
                        }
                    }
                }

            }
            cvlib.CvReleaseImage(ref image);
            sr.Close();

        }
        private void button21_Click(object sender, EventArgs e)
        {
            string[] files = Directory.GetFiles(txtPicFolder.Text, "*.lab");
            string[] symbols = Directory.GetFiles(@"D:\fontSize\Units", "*.bmp");
            Dictionary<int, int> tmp = new Dictionary<int, int>();
            Dictionary<string, Dictionary<int,int>> arrPhonemes = new Dictionary<string, Dictionary<int,int>>();
            for (int i = 0; i < symbols.Length; i++)
            {
                arrPhonemes.Add(Path.GetFileNameWithoutExtension(symbols[i]).ToLower(), new Dictionary<int, int>());
            }

            for (int i = 0; i < files.Length; i++)
            {

                ReadResults3(files[i], ref arrPhonemes);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            string unitsPath = txtOutputFile.Text;
            string[] symbols = Directory.GetFiles(unitsPath, "*.bmp");
            Dictionary<string, int[]> arrPhonemes = new Dictionary<string, int[]>();
            Dictionary<string, int> arrPhonemesCounter = new Dictionary<string, int>();
            for (int i = 0; i < symbols.Length; i++)
            {
                string name = Path.GetFileNameWithoutExtension(symbols[i]).ToLower();
                int startIndex = name.LastIndexOf("-");
                string strTmp = name.Substring(startIndex + 1, name.Length - startIndex - 1);
                string strSymbolName  = name.Substring(0, startIndex );
                string[] counters = strTmp.Split(new string[] {"__"}, StringSplitOptions.None);
                int firstCounter = int.Parse(counters[0]);
                int secondCounter = int.Parse(counters[1]);
                if (arrPhonemes.ContainsKey(strSymbolName))
                {
                    if (arrPhonemes[strSymbolName][0] < firstCounter)
                    {
                        arrPhonemes[strSymbolName][0] = firstCounter;
                    }
                    if (arrPhonemes[strSymbolName][1] < secondCounter)
                    {
                        arrPhonemes[strSymbolName][1] = secondCounter;
                    }
                }
                else
                {
                    arrPhonemes.Add(strSymbolName, new int[] {firstCounter, secondCounter});
                    arrPhonemesCounter.Add(strSymbolName, 1);
                }

            }
            string[] inputFiles = txtInputFile.Text.Split(';');
            for (int i = 0; i < inputFiles.Length; i++)
            {
                StreamReader sr = new StreamReader(inputFiles[i]);
                string text = sr.ReadToEnd();
                sr.Close();
                Regex regex = new Regex("\n");
                Match match = regex.Match(text);
                int spaceCounter = 1;
                const int imageSpaceCount = 20;
                IplImage image = cvlib.CvCreateImage(cvlib.CvSize(imageSpaceCount * 100, 50), 8, 3);
                cvlib.CvSet(ref image, cvlib.cvScalar(255,255,255,255));
                int start = 0;
                int end = 0;
                int len = 0;
                int prevPos = 0;
                int firstCounter = 0, secondCounter = 0;
                int PicCounter = 0;
                string textLab = "";
                while(match != null && match.Length != 0)
                {
                    if (spaceCounter % 20 == 0)
                    {
                        if (image.ptr != IntPtr.Zero)
                        {
                            IplImage originalImage = cvlib.CvCreateImage(cvlib.CvSize(start, image.height), 8, 3);
                            cvlib.CvSetImageROI(ref image, cvlib.cvRect(0, 0, start, image.height));
                            cvlib.CvCopy(ref image, ref originalImage);
                            string originalImagePath = txtOutputFolder.Text + "\\" + Path.GetFileNameWithoutExtension(inputFiles[i]) + "_" + PicCounter.ToString() + ".bmp";
                            string originalLablePath = txtOutputFolder.Text + "\\" + Path.GetFileNameWithoutExtension(inputFiles[i]) + "_" + PicCounter.ToString() + ".lab";
                            StreamWriter sw = new StreamWriter(originalLablePath);
                            sw.Write(textLab);
                            sw.Close();
                            cvlib.CvSaveImage(originalImagePath, ref originalImage);
                            cvlib.CvReleaseImage(ref originalImage);
                            cvlib.CvReleaseImage(ref image);
                            start = 0;
                            PicCounter++;
                            spaceCounter++;
                             textLab = "";
                        }
                        image = cvlib.CvCreateImage(cvlib.CvSize(imageSpaceCount * 300, image.height), 8, 3);
                    }
                    string phonem = text.Substring(prevPos , match.Index - prevPos);
                    if (!arrPhonemes.ContainsKey(phonem))
                    {

                        prevPos = match.Index + 1;
                        match = match.NextMatch();
                        continue;
                    }
                    textLab += phonem + "\n";
                    
                    string phonemPath = "";
                    while(true)
                    {
                        arrPhonemesCounter[phonem]++;
                        int phonemCount = ((arrPhonemes[phonem][0] + 1) * arrPhonemes[phonem][1]);
                        int mode1 = arrPhonemesCounter[phonem] % phonemCount;
                        firstCounter = mode1 / arrPhonemes[phonem][1] ;
                        secondCounter = mode1 % arrPhonemes[phonem][1] + 1;
                        phonemPath = unitsPath + phonem + "-" + firstCounter.ToString() + "__" + secondCounter.ToString() + ".bmp";
                        if (File.Exists(phonemPath))
                            break;
                    }
                    IplImage phonemPic = cvlib.CvLoadImage(phonemPath, cvlib.CV_LOAD_IMAGE_COLOR);
                    cvlib.CvSetImageROI(ref image, cvlib.cvRect(start, 0, phonemPic.width , image.height));
                    cvlib.CvCopy(ref phonemPic, ref image);
                    start += phonemPic.width;
                    cvlib.CvReleaseImage(ref phonemPic);
                    if(phonem == "space")
                        spaceCounter++;
                    prevPos = match.Index + 1;
                    match = match.NextMatch();

                }

            }
        }
        public SortedList<char, char[]> arrCorrespondCharsFirst;
        public SortedList<char, char[]> arrCorrespondCharsLast;
        public SortedList<char, char[]> arrCorrespondCharsMiddle;
        public int CorrespondCharsFirstCount;
        public int CorrespondCharsLastCount;
        public int CorrespondCharsMiddleCount;
        public SortedSet<string> lexicon;
        public SortedList<string, int> unigramLexicon;
        public SortedList<int, string> unigramIDLexicon;
        public SortedList<string, double> wightedLexicon;
        public void LoadChars()
        {
            arrCorrespondCharsFirst = new System.Collections.Generic.SortedList<char, char[]>();
            arrCorrespondCharsLast = new System.Collections.Generic.SortedList<char, char[]>();
            arrCorrespondCharsMiddle = new System.Collections.Generic.SortedList<char, char[]>();
            arrCorrespondCharsFirst.Add('ا', new char[] { 'آ', 'أ' });
            arrCorrespondCharsFirst.Add('ب', new char[] { 'پ', 'ی', 'ن', 'ت' });
            arrCorrespondCharsFirst.Add('پ', new char[] { 'ب', 'ی' });
            arrCorrespondCharsFirst.Add('ی', new char[] { 'ب', 'پ' });
            arrCorrespondCharsFirst.Add('ت', new char[] { 'ث', 'ن' });
            arrCorrespondCharsFirst.Add('ث', new char[] { 'ت', 'ن' });
            arrCorrespondCharsFirst.Add('ج', new char[] { 'ح', 'خ', 'چ' });
            arrCorrespondCharsFirst.Add('چ', new char[] { 'ج', 'ح', 'خ' });
            arrCorrespondCharsFirst.Add('ح', new char[] { 'خ', 'ج', 'چ' });
            arrCorrespondCharsFirst.Add('خ', new char[] { 'ح', 'ج', 'ج' });
            arrCorrespondCharsFirst.Add('د', new char[] { 'ذ', 'ر', 'ز' });
            arrCorrespondCharsFirst.Add('ذ', new char[] { 'د', 'ن' });
            arrCorrespondCharsFirst.Add('ر', new char[] { 'ز', 'د', 'ژ', 'و' });
            arrCorrespondCharsFirst.Add('ز', new char[] { 'د', 'ر', 'ژ', 'و' });
            arrCorrespondCharsFirst.Add('ژ', new char[] { 'ز', 'ر', 'و' });
            arrCorrespondCharsFirst.Add('س', new char[] { 'ش', 'ص', 'ض' });
            arrCorrespondCharsFirst.Add('ش', new char[] { 'س', 'ض', 'ص' });
            arrCorrespondCharsFirst.Add('ص', new char[] { 'ض', 'س', 'ش' });
            arrCorrespondCharsFirst.Add('ض', new char[] { 'ظ', 'ص', 'ش' });
            arrCorrespondCharsFirst.Add('ط', new char[] { 'ص', 'ض', 'ظ' });
            arrCorrespondCharsFirst.Add('ظ', new char[] { 'ص', 'ض', 'ط' });
            arrCorrespondCharsFirst.Add('ع', new char[] { 'غ', 'ئ' });
            arrCorrespondCharsFirst.Add('غ', new char[] { 'ف', 'ق', 'ع', 'ئ' });
            arrCorrespondCharsFirst.Add('ف', new char[] { 'غ', 'ق', 'ئ' });
            arrCorrespondCharsFirst.Add('ق', new char[] { 'ف', 'ت', 'ن', 'ئ' });
            arrCorrespondCharsFirst.Add('ک', new char[] { 'گ', 'ل', 'د', 'ذ' });
            arrCorrespondCharsFirst.Add('گ', new char[] { 'ک', 'ل', 'د', 'ذ' });
            arrCorrespondCharsFirst.Add('ن', new char[] { 'ت', 'ب', 'ئ' });
            arrCorrespondCharsFirst.Add('و', new char[] { 'ر', 'ز' });
            arrCorrespondCharsFirst.Add('ه', new char[] { 'م' });
            arrCorrespondCharsFirst.Add('ل', new char[] { 'ک', 'گ' });
            arrCorrespondCharsFirst.Add('أ', new char[] { 'آ', 'ا' });
            arrCorrespondCharsFirst.Add('آ', new char[] { 'أ', 'ا' });
            arrCorrespondCharsFirst.Add('ئ', new char[] { 'ت', 'ب','پ', 'ی', 'ن', 'ت' });


            arrCorrespondCharsMiddle.Add('ا', new char[] { 'ل', 'آ', 'أ' });
            arrCorrespondCharsMiddle.Add('ب', new char[] { 'پ', 'ی', 'ن', 'ت' });
            arrCorrespondCharsMiddle.Add('پ', new char[] { 'ب', 'ی' });
            arrCorrespondCharsMiddle.Add('ی', new char[] { 'ب', 'پ','س','ت' });
            arrCorrespondCharsMiddle.Add('ت', new char[] { 'ث', 'ن' });
            arrCorrespondCharsMiddle.Add('ث', new char[] { 'ت', 'ن' });
            arrCorrespondCharsMiddle.Add('ج', new char[] { 'ح', 'خ', 'چ' });
            arrCorrespondCharsMiddle.Add('چ', new char[] { 'ج', 'ح', 'خ' });
            arrCorrespondCharsMiddle.Add('ح', new char[] { 'خ', 'ج', 'چ' });
            arrCorrespondCharsMiddle.Add('خ', new char[] { 'ح', 'ج', 'ج' });
            arrCorrespondCharsMiddle.Add('د', new char[] { 'ذ', 'ر', 'ز', 'ک', 'گ' });
            arrCorrespondCharsMiddle.Add('ذ', new char[] { 'د', 'ن', 'ک', 'ک', 'گ' });
            arrCorrespondCharsMiddle.Add('ر', new char[] { 'ز', 'د', 'ژ', 'و', 'ؤ' });
            arrCorrespondCharsMiddle.Add('ز', new char[] { 'د', 'ر', 'ژ', 'و', 'ؤ' });
            arrCorrespondCharsMiddle.Add('ژ', new char[] { 'ز', 'ر', 'و' });
            arrCorrespondCharsMiddle.Add('س', new char[] { 'ش', 'ص', 'ض' });
            arrCorrespondCharsMiddle.Add('ش', new char[] { 'س', 'ض', 'ص' });
            arrCorrespondCharsMiddle.Add('ص', new char[] { 'ض', 'س', 'ش' });
            arrCorrespondCharsMiddle.Add('ض', new char[] { 'ص', 'س', 'ش', 'ظ','غ','ف','ق'});
            arrCorrespondCharsMiddle.Add('ط', new char[] { 'ص', 'ض', 'ظ', 'ا' });
            arrCorrespondCharsMiddle.Add('ظ', new char[] { 'ص', 'ض', 'ط', 'ا' });
            arrCorrespondCharsMiddle.Add('ع', new char[] { 'غ', 'ئ', 'ف' });
            arrCorrespondCharsMiddle.Add('غ', new char[] { 'ف', 'ق', 'ئ', 'ع' });
            arrCorrespondCharsMiddle.Add('ف', new char[] { 'غ', 'ق', 'ئ','ض' });
            arrCorrespondCharsMiddle.Add('ق', new char[] { 'ف','غ', 'ت', 'ن', 'ئ' });
            arrCorrespondCharsMiddle.Add('ک', new char[] { 'گ', 'ل', 'د', 'ذ' });
            arrCorrespondCharsMiddle.Add('گ', new char[] { 'ک', 'ل', 'د', 'ذ' });
            arrCorrespondCharsMiddle.Add('ن', new char[] { 'ت', 'ب', 'ئ' });
            arrCorrespondCharsMiddle.Add('و', new char[] { 'ر', 'ز', 'ؤ' });
            arrCorrespondCharsMiddle.Add('ه', new char[] { 'م' });
            arrCorrespondCharsMiddle.Add('ل', new char[] { 'ک', 'ا' });
            arrCorrespondCharsMiddle.Add('ؤ', new char[] { 'و' });
            arrCorrespondCharsMiddle.Add('أ', new char[] { 'آ', 'ا' });
            arrCorrespondCharsMiddle.Add('آ', new char[] { 'أ', 'ا' });
            arrCorrespondCharsMiddle.Add('ئ', new char[] { 'ت', 'ب', 'پ', 'ی', 'ن', 'ت' });



            arrCorrespondCharsLast.Add('ا', new char[] { 'ل', 'آ', 'أ' });
            arrCorrespondCharsLast.Add('ب', new char[] { 'پ', 'ت', 'ک', 'گ' });
            arrCorrespondCharsLast.Add('پ', new char[] { 'ب', 'ت', 'ک', 'گ' });
            arrCorrespondCharsLast.Add('ی', new char[] { 'ئ' });
            arrCorrespondCharsLast.Add('ت', new char[] { 'ث', 'ب', 'ک', 'گ', 'ف' });
            arrCorrespondCharsLast.Add('ث', new char[] { 'ت', 'ب', 'ک', 'گ' });
            arrCorrespondCharsLast.Add('ج', new char[] { 'ح', 'خ', 'چ', 'ع', 'غ' });
            arrCorrespondCharsLast.Add('چ', new char[] { 'ج', 'ح', 'خ', 'ع', 'غ' });
            arrCorrespondCharsLast.Add('ح', new char[] { 'خ', 'ج', 'چ', 'ع', 'غ' });
            arrCorrespondCharsLast.Add('خ', new char[] { 'ح', 'ج', 'ج', 'ع', 'غ' });
            arrCorrespondCharsLast.Add('د', new char[] { 'ذ', 'ر', 'ز', 'ک', 'گ' });
            arrCorrespondCharsLast.Add('ذ', new char[] { 'د' });
            arrCorrespondCharsLast.Add('ر', new char[] { 'ز', 'د', 'ژ', 'و' });
            arrCorrespondCharsLast.Add('ز', new char[] { 'د', 'ر', 'ژ', 'و' });
            arrCorrespondCharsLast.Add('ژ', new char[] { 'ز', 'ر', 'و' });
            arrCorrespondCharsLast.Add('س', new char[] { 'ش', 'ص', 'ض', 'ن', 'ق' });
            arrCorrespondCharsLast.Add('ش', new char[] { 'س', 'ض', 'ص', 'ن', 'ق' });
            arrCorrespondCharsLast.Add('ص', new char[] { 'ض', 'س', 'ش', 'ن', 'ق' });
            arrCorrespondCharsLast.Add('ض', new char[] { 'ظ', 'ص', 'ش', 'ن', 'ق' });
            arrCorrespondCharsLast.Add('ط', new char[] { 'ظ', 'ا' });
            arrCorrespondCharsLast.Add('ظ', new char[] { 'ط', 'ا' });
            arrCorrespondCharsLast.Add('ع', new char[] { 'غ', 'ح', 'خ', 'چ', 'ج' });
            arrCorrespondCharsLast.Add('غ', new char[] { 'ع', 'ح', 'خ', 'چ', 'ج' });
            arrCorrespondCharsLast.Add('ف', new char[] { 'ب', 'ت' });
            arrCorrespondCharsLast.Add('ق', new char[] { 'ش', 'س', 'ض', 'ص', 'ن' });
            arrCorrespondCharsLast.Add('ک', new char[] { 'گ', 'د', 'ذ' });
            arrCorrespondCharsLast.Add('گ', new char[] { 'ک', 'د', 'ذ' });
            arrCorrespondCharsLast.Add('ن', new char[] { 'س', 'ش', 'ص', 'ض', 'ق' });
            arrCorrespondCharsLast.Add('و', new char[] { 'ر', 'ز', 'ؤ' });
            arrCorrespondCharsLast.Add('ه', new char[] { 'ة', 'ۀ' });
            arrCorrespondCharsLast.Add('ة', new char[] { 'ه', 'ۀ' });
            arrCorrespondCharsLast.Add('ۀ', new char[] { 'ة', 'ه' });
            arrCorrespondCharsLast.Add('ل', new char[] { 'ک' });
            arrCorrespondCharsLast.Add('ؤ', new char[] { 'و' });
            arrCorrespondCharsLast.Add('أ', new char[] { 'آ', 'ا' });
            arrCorrespondCharsLast.Add('آ', new char[] { 'أ', 'ا' });

            CorrespondCharsFirstCount = arrCorrespondCharsFirst.Count;
            CorrespondCharsLastCount = arrCorrespondCharsLast.Count;
            CorrespondCharsMiddleCount = arrCorrespondCharsMiddle.Count;
        }
        public void LoadWightedLexicon(string lexPath)
        {
            StreamReader sr = new StreamReader(lexPath, Encoding.GetEncoding(1256));
           
            string strWord;
            bool bBigramIsStarted = false;
            while (!sr.EndOfStream)
            {
                strWord = sr.ReadLine();
                string[] values = strWord.Split(new char[] { '\t' });
                if (!bBigramIsStarted && strWord == "\\2-grams:")
                   bBigramIsStarted = true;
                if(values.Length >= 2)
                {
                    if (!bBigramIsStarted)
                    {
                        wightedLexicon.Add(values[1], double.Parse(values[0]));
                        unigramLexicon.Add(values[1], unigramLexicon.Count + 1);
                    }
                    else
                    {
                        wightedLexicon.Add(values[1], double.Parse(values[0]));
                    }
                }
            } 
            sr.Close();           
        }
        public void LoadLexicon(string lexPath, Encoding encoding) 
        {
            StreamReader sr = new StreamReader(lexPath, encoding);
            string strWord;
            while (!sr.EndOfStream)
            {
                strWord = PreProcess(sr.ReadLine());
                if (strWord == null || strWord == "")
                {
                    continue;
                }
                lexicon.Add(strWord.Split(new char[] {'\t'})[0]);
            }    
            sr.Close();        
        }
        protected string PreProcess(string BaseText)
        {
            StringBuilder TempText = new StringBuilder();
            //Declare the standard Alphabet
            char[] Alphabet = { 'ا', 'آ', 'ب', 'پ', 'ت', 'ث', 'ج', 'چ', 'ح', 'خ', 'د', 'ذ',
                          'ر','ز','ژ','س','ش','ص','ض','ط','ظ','ع','غ','ف','ق','ک','گ',
                          'ل','م','ن','و','ه','ی','(',')',']','{','[','}','آ','.','،',
                          '،','!','?','؟','%','@','-','_','/','پ',':','"',';',' ','%',
                          Convert.ToChar("\n"),'.','؛','»','«','+'};

            char[] Separator = { '/', '\\', '%', '%', '؛', ':', '"', '[', ']', '(', ')' };
            int Length = BaseText.Length;
            for (int i = 0; i < Length; i++)
            {
                switch (BaseText[i])
                {
                    //case '‌': break;NimFasele
                    case '=':TempText.Append('='); break;
                    case '\\':TempText.Append('\\'); break;
                    case 'ت': TempText.Append('ت'); break;
                    case 'ي': TempText.Append('ی'); break;
                    case 'ێ': TempText.Append('ی'); break;
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case 'ئ':
                        if (i != (Length - 1) && BaseText[i + 1] == 'ی')
                        {
                            TempText.Append("ی");
                            break;
                        }
                        else
                        {
                            TempText.Append('ئ');
                            break;
                        }
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '“': TempText.Append(" \" "); break;
                    case '”': TempText.Append(" \" "); break;
                    case 'ۀ': TempText.Append('ه'); break;
                    case 'ة': TempText.Append('ه'); break;
                    case '،': TempText.Append(" " + '،' + " "); break;
                    case 'ؤ': TempText.Append('و'); break;
                    case 'إ': TempText.Append('ا'); break;
                    case 'ۇ': TempText.Append('و'); break;
                    case 'أ': TempText.Append('ا'); break;
                    case 'ك': TempText.Append('ک'); break;
                    case 'ھ': TempText.Append('ه'); break;
                    case 'ﻳ': TempText.Append('ی'); break;
                    case 'ﯽ': TempText.Append('ی'); break;
                    case 'ﻲ': TempText.Append('ی'); break;
                    case 'ﻟ': TempText.Append('ل'); break;
                    case 'ﺳ': TempText.Append('س'); break;
                    case 'ﺘ': TempText.Append('ت'); break;
                    case 'ﺎ': TempText.Append('ا'); break;
                    case 'ﺗ': TempText.Append('ت'); break;
                    case 'ﻴ': TempText.Append('ی'); break;
                    case 'ﻦ': TempText.Append('ن'); break;
                    case 'ں': TempText.Append('ن'); break;
                    case 'ﻣ': TempText.Append('م'); break;
                    case 'ﻯ': TempText.Append('ی'); break;
                    case 'ﭙ': TempText.Append('پ'); break;
                    case 'ﭘ': TempText.Append('پ'); break;
                    case 'ﮔ': TempText.Append('گ'); break;
                    case 'ﺞ': TempText.Append('ج'); break;
                    case 'ﺇ': TempText.Append('ا'); break;
                    case 'ﻐ': TempText.Append('غ'); break;
                    case 'ﺮ': TempText.Append('ر'); break;
                    case 'ﺴ': TempText.Append('ﺴ'); break;
                    case 'ﻘ': TempText.Append('ق'); break;
                    case 'ﻢ': TempText.Append('م'); break;
                    case 'ﻨ': TempText.Append('ن'); break;
                    case 'ﻅ': TempText.Append('ظ'); break;
                    case 'ﺩ': TempText.Append('د'); break;
                    case 'ﺫ': TempText.Append('ذ'); break;
                    case 'ﺍ': TempText.Append('ا'); break;
                    case 'ﺭ': TempText.Append('ر'); break;
                    case '–': TempText.Append(" – "); break;
                    case 'ﻭ': TempText.Append('و'); break;
                    case '-': TempText.Append(" - "); break;
                    case 'ـ': TempText.Append(""); break;
                    case ',': TempText.Append("،"); break;
                    case 'ْ': TempText.Append(""); break;
                    case 'ٌ': TempText.Append(""); break;
                    case 'ٍ': TempText.Append(""); break;
                    case 'ً': TempText.Append(""); break;
                    case 'ُ': TempText.Append(""); break;
                    case 'ِ': TempText.Append(""); break;
                    case 'َ': TempText.Append(""); break;
                    case 'ّ': TempText.Append(""); break;
                    case 'ٔ': TempText.Append(""); break;
                    case '': TempText.Append(""); break;
                    case '"':
                        if ((i > 0 && BaseText[i - 1] == '.') && (i > 1 && BaseText[i - 1] == ' '))
                        {
                            TempText.Append(" \"" + "\n");
                        }
                        else if (i > 0 && BaseText[i - 1] == '؟')
                        {
                            TempText.Append(" \"" + "\n");
                        }
                        else if (i > 0 && BaseText[i - 1] == '!')
                        {
                            TempText.Append(" \"" + "\n");
                        }
                        else
                        {
                            TempText.Append(" \" ");
                        }
                        break;
                    case '»':
                        if ((i > 0 && BaseText[i - 1] == '.') && (i > 1 && BaseText[i - 1] == ' '))
                        {
                            TempText.Append("»" + "\n");
                        }
                        else if (i > 0 && BaseText[i - 1] == '؟')
                        {
                            TempText.Append("»" + "\n");
                        }
                        else if (i > 0 && BaseText[i - 1] == '!')
                        {
                            TempText.Append("»" + "\n");
                        }
                        else
                        {
                            TempText.Append(" »");
                        }
                        break;
                    case '’': TempText.Append(" \" "); break;
                    case '‘': TempText.Append(" \" "); break;
                    //Persian Number
                    case '۰':
                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("0" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "0"); break;
                        }
                        TempText.Append("0"); break;
                    case '۱':
                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("1" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "1"); break;
                        }
                        TempText.Append("1"); break;
                    case '۲':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("2" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "2"); break;
                        }
                        TempText.Append("2"); break;
                    case '۳':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("3" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "3"); break;
                        }
                        TempText.Append("3"); break;
                    case '۴':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("4" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "4"); break;
                        }
                        TempText.Append("4"); break;
                    case '۵':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("5" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "5"); break;
                        }
                        TempText.Append("5"); break;
                    case '۶':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("6" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "6"); break;
                        }
                        TempText.Append("6"); break;
                    case '۷':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("7" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "7"); break;
                        }
                        TempText.Append("7"); break;
                    case '۸':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("8" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "8"); break;
                        }
                        TempText.Append("8"); break;
                    case '۹':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("9" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "9"); break;
                        }
                        TempText.Append("9"); break;
                    //English Number
                    case '0':
                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("0" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "0"); break;
                        }
                        TempText.Append("0"); break;
                    case '1':
                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("1" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "1"); break;
                        }
                        TempText.Append("1"); break;
                    case '2':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("2" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "2"); break;
                        }
                        TempText.Append("2"); break;
                    case '3':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("3" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "3"); break;
                        }
                        TempText.Append("3"); break;
                    case '4':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("4" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "4"); break;
                        }
                        TempText.Append("4"); break;
                    case '5':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("5" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "5"); break;
                        }
                        TempText.Append("5"); break;
                    case '6':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("6" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "6"); break;
                        }
                        TempText.Append("6"); break;
                    case '7':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("7" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "7"); break;
                        }
                        TempText.Append("7"); break;
                    case '8':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("8" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "8"); break;
                        }
                        TempText.Append("8"); break;
                    case '9':

                        if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                        {
                            TempText.Append("9" + " "); break;
                        }
                        if (i > 1 && Char.IsLetter(BaseText[i - 1]))
                        {
                            TempText.Append(" " + "9"); break;
                        }
                        TempText.Append("9"); break;
                    case '': TempText.Append(""); break;
                    //case '‌': TempText.Append(" "); break;
                    case '«': TempText.Append("« "); break;
                    case '+':
                        if ((i > 0) && ((BaseText[i - 1] == '5') || (BaseText[i - 1] == '۵') || (BaseText[i - 1] == '۱') || (BaseText[i - 1] == '1')))
                        {
                            TempText.Append("+");
                        }
                        else
                        {
                            TempText.Append(" + ");
                        }
                        break;
                    case ':': TempText.Append(" : "); break;

                        if ((i > 0 && BaseText[i - 1] == '.') && (i > 1 && BaseText[i - 1] == ' '))
                        {
                            TempText.Append(" \"" + "\n");
                        }
                        else if (i > 0 && BaseText[i - 1] == '؟')
                        {
                            TempText.Append(" \"" + "\n");
                        }
                        else if (i > 0 && BaseText[i - 1] == '!')
                        {
                            TempText.Append(" \"" + "\n");
                        }
                        else
                        {
                            TempText.Append(" \" ");
                        }
                        break;
                    case '.':
                        if (i < (Length - 2) && BaseText[i + 1] == '.' && BaseText[i + 2] == '.')
                        {
                            TempText.Append(" ... ");
                            i += 2; break;
                        }
                        //Today problem with 3.6
                        if (i > 0 && i != (Length - 1) && ((BaseText[i + 1] == '0') || (BaseText[i + 1] == '1') || (BaseText[i + 1] == '2') || (BaseText[i + 1] == '3') || (BaseText[i + 1] == '4') || (BaseText[i + 1] == '5') || (BaseText[i + 1] == '6') || (BaseText[i + 1] == '7') || (BaseText[i + 1] == '8') || (BaseText[i + 1] == '9'))
                            && ((BaseText[i - 1] == '0') || (BaseText[i - 1] == '1') || (BaseText[i - 1] == '2') || (BaseText[i - 1] == '3') || (BaseText[i - 1] == '4') || (BaseText[i - 1] == '5') || (BaseText[i - 1] == '6') || (BaseText[i - 1] == '7') || (BaseText[i - 1] == '8') || (BaseText[i - 1] == '9')))
                        {
                            TempText.Append("."); break;
                        }
                        if (i > 0 && i != (Length - 1) && ((BaseText[i + 1] == '۰') || (BaseText[i + 1] == '۱') || (BaseText[i + 1] == '۲') || (BaseText[i + 1] == '۳') || (BaseText[i + 1] == '۴') || (BaseText[i + 1] == '۵') || (BaseText[i + 1] == '۶') || (BaseText[i + 1] == '۷') || (BaseText[i + 1] == '۸') || (BaseText[i + 1] == '۹'))
                            && ((BaseText[i - 1] == '۰') || (BaseText[i - 1] == '۱') || (BaseText[i - 1] == '۲') || (BaseText[i - 1] == '۳') || (BaseText[i - 1] == '۴') || (BaseText[i - 1] == '۵') || (BaseText[i - 1] == '۶') || (BaseText[i - 1] == '۷') || (BaseText[i - 1] == '۸') || (BaseText[i - 1] == '۹')))
                        {
                            TempText.Append(","); break;
                        }
                        if ((i > 0 && (Char.IsLetter(BaseText[i - 1]))) && ((i != (Length - 1)) && (Char.IsLetter(BaseText[i + 1]))))
                        {
                            TempText.Append("."); break;
                        }
                        if (i > 0 && i < (Length - 1))
                        {
                            if ((Char.IsLetter(BaseText[i - 1]) && (BaseText[i + 1] == ' ') && (i < (Length - 2) && (BaseText[i + 2] != '"' || BaseText[i + 2] != '»' || BaseText[i + 2] != '”'))) || (Char.IsLetter(BaseText[i + 1]) && BaseText[i - 1] == ' '))
                            {
                                TempText.Append(" . " + "\n"); break;
                            }
                        }
                        if (i == (Length - 1))
                        {
                            TempText.Append(" . "); break;
                        }
                        if ((i != (Length - 1) && BaseText[i + 1] != '.' && BaseText[i + 1] != '»' && BaseText[i + 1] != '”' && BaseText[i + 1] != '"') && (i != (Length - 2) && BaseText[i + 2] != '»') && (i != (Length - 1) && BaseText[i - 1] != '.'))
                        {
                            TempText.Append(" . " + "\n"); break;
                        }
                        TempText.Append(" . "); break;
                    case '!':

                        if (i != (Length - 1) && (BaseText[i + 1] == '!' || BaseText[i + 1] == '؟' || BaseText[i + 1] == ':'))
                        {
                            TempText.Append(" " + "! "); break;
                        }
                        if (i < (Length - 2) && (BaseText[i + 2] == '"' || BaseText[i + 2] == '»' || BaseText[i + 2] == ':'))
                        {
                            TempText.Append(" " + "!"); break;
                        }
                        if (i != (Length - 1) && BaseText[i + 1] != '"' && BaseText[i + 1] != '»' && BaseText[i + 1] != '”')
                        {
                            TempText.Append(" " + "!" + "\n"); break;
                        }
                        TempText.Append(" " + "! "); break;
                    case '؟':
                        if (i != (Length - 1) && (BaseText[i + 1] == '!' || BaseText[i + 1] == '؟'))
                        {
                            TempText.Append(" " + "؟ "); break;
                        }
                        if (i < (Length - 2) && (BaseText[i + 2] == '"' || BaseText[i + 2] == '»'))
                        {
                            TempText.Append(" " + "؟"); break;
                        }
                        if (i != (Length - 1) && BaseText[i + 1] != '"' && BaseText[i + 1] != '»' && BaseText[i + 1] != '”')
                        {
                            TempText.Append(" " + "؟" + "\n"); break;
                        }
                        TempText.Append(" " + "! "); break;
                    case '':
                        TempText.Append(""); break;
                    default:
                        if (Char.IsLetter(BaseText[i]) || Char.IsDigit(BaseText[i]) || Alphabet.Contains(BaseText[i]))
                        {
                            if (Separator.Contains(BaseText[i]))
                            {
                                TempText.Append(" " + BaseText[i] + " "); break;
                            }
                            if (Char.IsDigit(BaseText[i]))
                            {
                                if (i != (Length - 1) && Char.IsLetter(BaseText[i + 1]))
                                {
                                    TempText.Append(BaseText[i] + " "); break;
                                }
                                if (i > 0 && Char.IsLetter(BaseText[i - 1]))
                                {
                                    TempText.Append(" " + BaseText[i]); break;
                                }
                                TempText.Append(BaseText[i]); break;
                            }
                            if (BaseText[i] == ' ')
                            {
                                if (i > 0 && i != (Length - 1) && (BaseText[i - 1] == '.' || BaseText[i - 1] == '!' || BaseText[i - 1] == '؟'))
                                {
                                    TempText.Append(""); break;
                                }
                                else
                                {
                                    TempText.Append(BaseText[i]); break;
                                }
                            }
                            TempText.Append(BaseText[i]);
                        }
                        break;
                }
            }
            string FinalString = TempText.ToString();
            FinalString = Regex.Replace(FinalString, @"\t", " ", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString.ToString(), @"\n\n", "\n", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"^\s+$[\r\n\b\t]*", "", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"  ", " ", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"  ", " ", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"  ", " ", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"!\n\n", "!\n", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"؟\n\n", "؟\n", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @".\n\n", ".\n", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"\n\n", "\n", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"\t", "", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @";\n", ";", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"ــ", "-", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"هها", "ه ها", RegexOptions.Multiline);
            FinalString = Regex.Replace(FinalString, @"\n ", "\n", RegexOptions.Multiline);


            return FinalString;
        }
        public bool IsNumber(string word)
        {
            bool retVal = false;

            Regex regex = new Regex("^[0-9]+\\.?[0-9]+$");
            if (regex.Match(word).Length > 0)
                retVal = true;
            
            return retVal;
        }
        public string FindCorrectWord(string word, ref SortedList<string, double> foundedWords)
        {
            string correctedWord = word;
            if (word.Length <= 1)
                return correctedWord;
            char[] arrChars = word.ToArray();
            bool bFound = false;

            if (arrCorrespondCharsFirst.ContainsKey(word[0]))
            for (int j = 0; j < arrCorrespondCharsFirst[word[0]].Length; j++)
            {
                arrChars = word.ToArray();
                arrChars[0] = arrCorrespondCharsFirst[word[0]][j];
                correctedWord = new string(arrChars);
                if (lexicon.Contains(correctedWord))
                {
                    foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                    //bFound = true;
                    //break;
                }
                correctedWord = word;
            }
            if (!bFound && word[word.Length - 1] == 'ا' || word[word.Length - 1] == 'أ')
            {

                correctedWord = word.Substring(0, word.Length - 1) + "اً";
                if (lexicon.Contains(correctedWord))
                {
                    foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                    //bFound = true;
                }
                correctedWord = word;
            }
            if (!bFound && word[word.Length - 1] == 'ً' && word[word.Length - 2] == 'ا')
            {

                correctedWord = word.Substring(0, word.Length - 2) + "ا";
                if (lexicon.Contains(correctedWord))
                {
                    foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                    //bFound = true;
                }
                else
                { 
                    correctedWord = word.Substring(0, word.Length - 2) + "آ";
                    if (lexicon.Contains(correctedWord))
                    {
                        //bFound = true;
                        foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                    }
                    else
                    {
                        correctedWord = word.Substring(0, word.Length - 2) + "أ";
                        if (lexicon.Contains(correctedWord))
                        {
                            //bFound = true;
                            foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                        }
                    }
                }
                correctedWord = word;
            }
            if (arrCorrespondCharsFirst.ContainsKey(word[word.Length - 1]))
            for (int j = 0; !bFound && j < arrCorrespondCharsLast[word[word.Length - 1]].Length; j++)
            {
                arrChars = word.ToArray();
                arrChars[word.Length - 1] = arrCorrespondCharsLast[word[word.Length - 1]][j];
                correctedWord = new string(arrChars);
                if (lexicon.Contains(correctedWord))
                {
                    foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                    //bFound = true;
                    //break;
                }
                correctedWord = word;
            }
            for (int i = 1; !bFound && i < word.Length - 1; i++)
            {
                if(arrCorrespondCharsMiddle.ContainsKey(word[i]))
                for (int j = 0; j < arrCorrespondCharsMiddle[word[i]].Length; j++)
                {
                    arrChars = word.ToArray();
                    arrChars [i] = arrCorrespondCharsMiddle[word[i]][j];
                    correctedWord = new string(arrChars);
                    if (lexicon.Contains(correctedWord))
                    {
                        foundedWords.Add(correctedWord, FindUnigramItem(FindUnigramIndex(correctedWord)).prob);
                        //bFound = true;
                        //break;
                    }
                    correctedWord = word;
                }
            }
            for (int i = 0; i < foundedWords.Count; i++)
            {
                foundedWords[foundedWords.Keys[i]] = Math.Pow(10, foundedWords.Values[i]);              
            }
            return correctedWord;
        }
        public string CheckSpell(string strContent)
        {
            //strContent = PreProcess(strContent);
            string finalText = "";
            int startIndex = 0;
            int lastInex = -1;
            string word = "";
            string lastWord = "";
            char ch ;
            int wordsCount = strContent.Length;
            for (int i = 0; i < wordsCount; i++)
            {
                ch = strContent[i];

                switch (ch)
                {
                    case '':
                    case '':
                    case '':
                    case '':
                    case '':
                    case '"':
                    case '\'':
                    case '\r':
                    case '':
                    case '‌':
                    case '{':
                    case '}':
                    case '(':
                    case ')':
                    case '،':
                    case '؛':
                    case ':':
                    case '.':
                    case '/':
                    case '%':
                    case '-':
                    case '_':
                    case '?':
                    case '[':
                    case ']':
                    case ',':
                    case '÷':
                    case '×':
                    case '=':
                    case '*':
                    case '!':
                    case '+':
                    case '»':
                    case '«':
                    case '\\':
                    case '>':
                    case '<':
                    case '&':
                    case '@':
                    case '$':
                    case '^':
                    case ';':
                    case ' ':
                        break;
                    case '\n':
                        lastWord = "";
                        break;
                    default:
                        continue;
                }
                word = strContent.Substring(lastInex + 1, i - lastInex - 1 );
                if (lexicon.Contains(word) || IsNumber(word) || (word.Length > 0 && word[word.Length - 1] == 'ی' && lexicon.Contains(word.Substring(0, word.Length - 1))))
                {
                    finalText += word;
                    lastWord = word;
                }
                else
                {
                    SortedList<string, double> foundedWords = new SortedList<string,double>();
                    //finalText += FindCorrectWord(word, ref foundedWords);
                    FindCorrectWord(word, ref foundedWords);
                    string maxProbWord = "";
                    if (lastWord != "")
                    {
                        double max = 0;
                        double pdf = 0;
                        for (int j = 0; j < foundedWords.Count; j++)
                        {
                            pdf = FindBigramProbability(lastWord, foundedWords.Keys[j]);
                            if (pdf > max)
                            {
                                maxProbWord = foundedWords.Keys[j];
                                max = pdf;
                            }
                        }
                    }
                    else
                    {
                        double max = 0;
                        for (int j = 0; j < foundedWords.Count; j++)
                        {
                            if (foundedWords.Values[j] > max)
                            {
                                maxProbWord = foundedWords.Keys[j];
                                max = foundedWords.Values[j];
                            }
                        }
                    }
                    finalText += lastWord = maxProbWord;
                }
                finalText += ch.ToString();
                lastInex = i ;
            }
            return finalText;
        }
        private void button23_Click(object sender, EventArgs e)
        {

            lexicon = new SortedSet<string>();
            wightedLexicon = new SortedList<string, double>();
            //LoadWightedLexicon(@"D:\_OCR_DATASET\gholipour\khabari_hamcode_wordlist_DIC_TARKIB1_2.lm");
            //LoadLexicon("D:\\fulldict.dat", Encoding.UTF8);
            LoadLexicon("D:\\dict.dat", Encoding.GetEncoding(1256));
            LoadLexicon("D:\\dict2.dat", Encoding.GetEncoding(1256));
            LoadLexicon("D:\\PersianAcademy.dat", Encoding.GetEncoding(1256));
            LoadLexicon("D:\\PersianAcademy2.dat", Encoding.GetEncoding(1256));
            LoadChars();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.DefaultExt = "txt";
            of.Filter = "Text files (*.txt)|*.txt";
            of.Multiselect = true;
            if (of.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                for (int i = 0; i < of.FileNames.Length; i++)
                {

                     txtSpellCheckFile.Text += of.FileNames[i] + ";";
                }
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            string[] inputFiles = txtSpellCheckFile.Text.Split(';');
            string text = "";
            for (int i = 0; i < inputFiles.Length - 1; i++)
            {
                StreamReader sr = new StreamReader(inputFiles[i], Encoding.GetEncoding(1256));
                text += sr.ReadToEnd();
                sr.Close();
            }
            text = CheckSpell(text.Replace('ي','ی'));
            richTextBox2.Text = text;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            SortedList<string, int> lexicon = new SortedList<string, int>();
            string[] inputFiles = txtSpellCheckFile.Text.Split(';');
            string strContent = "";
            for (int i = 0; i < inputFiles.Length - 1; i++)
            {
                try
                {
                StreamReader sr = new StreamReader(inputFiles[i], Encoding.GetEncoding(1256));
                strContent += sr.ReadToEnd();
                sr.Close();

                }
                catch (OutOfMemoryException ex)
                {

                    break;
                }
            }
            int lastInex = -1;
            string word = "";
            char ch ;
            int wordsCount = strContent.Length;
            for (int i = 0; i < wordsCount; i++)
            {
                ch = strContent[i];

                switch (ch)
                {
                    case '#':
                    case '"':
                    case '\'':
                    case 'ء':
                    case '1':
                    case '2':
                    case '3':
                    case '4':
                    case '5':
                    case '6':
                    case '7':
                    case '8':
                    case '9':
                    case '0':
                    case '{':
                    case '}':
                    case '(':
                    case ')':
                    case '،':
                    case '؛':
                    case ';':
                    case ':':
                    case '.':
                    case '@':
                    case '/':
                    case '%':
                    case '-':
                    case '_':
                    case '$':
                    case '?':
                    case '؟':
                    case '[':
                    case ']':
                    case ',':
                    case '÷':
                    case '×':
                    case '=':
                    case '*':
                    case '!':
                    case '+':
                    case '»':
                    case '«':
                    case'':
                    case '\\':
                    case ' ':
                    case '\n':
                    case  '^':
                    case  '&':
                    case  '>':
                    case  '<':
                    case  '~':
                    case  '|':
                    case '\r':
                    case '\t':
                        break;
                    default:
                        continue;
                }
                word = strContent.Substring(lastInex + 1, i - lastInex - 1 );
                word = word.Replace('ي', 'ی');
                if (!lexicon.ContainsKey(word) && word.Length > 2)
                {
                    lexicon.Add(word, 1);
                }
                else if (word.Length > 2)
                {
                    lexicon[word]++;
                }
                lastInex = i ;
            }
            StreamWriter sw = new StreamWriter(txtDicPath.Text);
            string strWord = "";
            int MinCountForWords = 50;
            for (int i = 0; i < lexicon.Count; i++)
            {
                strWord = lexicon.Keys[i];
                if(lexicon[strWord] > MinCountForWords)
                    sw.WriteLine(strWord);
            }
            sw.Close();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            StreamReader sr = new StreamReader("D:\\model02");
            string content = sr.ReadToEnd();
            sr.Close();
            StreamWriter sw = new StreamWriter("D:\\model");
            BinaryWriter bw = new BinaryWriter(sw.BaseStream);
            byte[] bytes = new byte[content.Length];
            for (int i = 0; i < content.Length; i++)
            {
                bytes[i] = (byte)1;
            }
            bw.Write(bytes);
            bw.Close();
           
            //string token = "\\3-grams:";
            //string inputFile = txtSpellCheckFile.Text;
            //string strContent = "";
            //StreamReader sr = new StreamReader(inputFile, Encoding.GetEncoding(1256));
            //StreamWriter sw = new StreamWriter(Path.GetDirectoryName(inputFile) + "\\" + Path.GetFileNameWithoutExtension(inputFile) + "_2.lm", true, Encoding.GetEncoding(1256));
            //while (!sr.EndOfStream)
            //{
            //    strContent = ( sr.ReadLine());
            //    if (strContent == token)
            //    {
            //        break;
            //    }
            //    sw.WriteLine(strContent);
            //}
            //sr.Close();
            //sw.Close();
            
        }

        struct UnigramItem : IComparable<UnigramItem>
        {
            public double prob;
            public int id;
            public int firstIndex;
            public int length;
            public int CompareTo(UnigramItem a)
            {
                if (a.id < id)
                    return 1;
                else if (a.id > id)
                    return -1;
                else
                    return 0;
            }
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct Unigram
        {            
            public double prob;
            public int id;
            public int firstIndex;
            public int length;    
        }
        [StructLayout(LayoutKind.Sequential), Serializable]
        struct Bigram
        {
            public double prob;
            public int secondWordIndex;            
        }
        IntPtr arrBiGrams;
        IntPtr arrUniGrams;
        long arrBiGramsCount;
        long arrUniGramsCount;
        private void button28_Click(object sender, EventArgs e)
        {
            string strBigramBinary = System.Windows.Forms.Application.StartupPath +  @"\\khabari_hamcode_wordlist_DIC_TARKIB_BigramBinary.lm";
            string strUnigramBinary = System.Windows.Forms.Application.StartupPath + @"\\khabari_hamcode_wordlist_DIC_TARKIB_UnigramBinary.lm";
            string strUnigramStrings = System.Windows.Forms.Application.StartupPath + @"\\khabari_hamcode_wordlist_DIC_TARKIB_UnigramStrings.lm";
            string strInputLM = System.Windows.Forms.Application.StartupPath + @"\\khabari_hamcode_wordlist_DIC_TARKIB1_2.lm";
            lexicon = new SortedSet<string>();
            wightedLexicon = new SortedList<string, double>();
            SortedList<int, double> tmpwightedLexicon = new SortedList<int, double>();
            unigramLexicon = new SortedList<string, int>();
            SortedSet<UnigramItem> uniList = new SortedSet<UnigramItem>();
            
            LoadWightedLexicon(strInputLM);
            int wightedLexiconCount = wightedLexicon.Count;
            Unigram uniGram = new Unigram();
            Bigram biGram = new Bigram();
            
            
            int unigramLexiconCount = unigramLexicon.Count;
            int lastIndex = 0;
            double prevProb = 0.0f;
            int prevId = -1;
            StreamWriter swUnigramStrings = new StreamWriter(strUnigramStrings);
            StreamWriter swBigramBinary = new StreamWriter(strBigramBinary);
            StreamWriter swUnigramBinary = new StreamWriter(strUnigramBinary);
            BinaryWriter bwBigramBinary = new BinaryWriter(swBigramBinary.BaseStream);
            BinaryWriter bwUnigramBinary = new BinaryWriter(swUnigramBinary.BaseStream);
            int bigramItemsCounter = 0;
            unsafe
            {
                for (int i = 0; i < unigramLexiconCount; i++)
                {
                    swUnigramStrings.WriteLine(unigramLexicon.Keys[i] + "\t" + unigramLexicon.Values[i]);
                }
                byte[] biGramBytes = new byte[sizeof(Bigram)];
                int biGramBytesCount = biGramBytes.Length;
                byte[] uniGramBytes = new byte[sizeof(Unigram)];
                int uniGramBytesCount = uniGramBytes.Length;
                byte* biGramPtr = (byte*)&biGram;
                byte* uniGramPtr = (byte*)&uniGram;
                unigramLexiconCount = 0;
                for (int i = 0; i < wightedLexiconCount; i++)
                {
                    try
                    {

                        string[] words = wightedLexicon.Keys[i].Split(' ');
                        if (words.Length == 1)
                        {
                            if (prevId >= 0)
                            {
                                UnigramItem tmpUniGram = new UnigramItem();
                                tmpUniGram.id = prevId;//unigramLexicon[words[0]];
                                tmpUniGram.prob = prevProb;//(float)wightedLexicon[words[0]];
                                int tmpwightedLexiconCount = tmpwightedLexicon.Count;
                                if (unigramLexiconCount > 0 && lastIndex == 0)
                                {
                                    tmpUniGram.firstIndex = -1;
                                    tmpUniGram.length = 0;
                                }
                                else
                                {
                                    tmpUniGram.firstIndex = bigramItemsCounter;//lastIndex - unigramLexiconCount + 1;
                                    tmpUniGram.length = tmpwightedLexiconCount;
                                }
                                uniList.Add(tmpUniGram);                                
                                for (int j = 0; j < tmpwightedLexiconCount; j++)
                                {
                                    biGram.prob = tmpwightedLexicon.Values[j];
                                    biGram.secondWordIndex = tmpwightedLexicon.Keys[j];
                                    for (int m = 0; m < biGramBytesCount; m++)
                                    {
                                        biGramBytes[m] = biGramPtr[m];
                                    }
                                    bwBigramBinary.Write(biGramBytes);
                                }
                                bigramItemsCounter += tmpwightedLexiconCount;
                                tmpwightedLexicon.Clear();
                            }
                            prevId = unigramLexicon[words[0]];
                            prevProb = (double)wightedLexicon.Values[i];
                            lastIndex = i;
                            unigramLexiconCount++;
                        }
                        else
                        {
                            biGram.prob = (double)wightedLexicon.Values[i];
                            biGram.secondWordIndex = unigramLexicon[words[1]];
                            if (!tmpwightedLexicon.ContainsKey(biGram.secondWordIndex))
                                tmpwightedLexicon.Add(biGram.secondWordIndex, biGram.prob);
                            //swBigramBinary.WriteLine(biGram);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            int uniListCount = uniList.Count;
            UnigramItem uniGramItem = new UnigramItem();
            for (int i = 0; i < uniListCount; i++)
            {
                uniGramItem = uniList.ElementAt(i);
                uniGram.id = uniGramItem.id;
                uniGram.firstIndex = uniGramItem.firstIndex;
                uniGram.length = uniGramItem.length;
                uniGram.prob = uniGramItem.prob;
                for (int j = 0; j < uniGramBytesCount; j++)
                {
                    uniGramBytes[j] = uniGramPtr[j];
                }
                bwUnigramBinary.Write(uniGramBytes);
            }
            }
            bwBigramBinary.Close();
            bwUnigramBinary.Close();
            swUnigramStrings.Close();
            
        }
        private int FindUnigramIndex(string word)
        {
            int wordIndex = 0;
            word = word.Replace('ی','ي');
            word = word.Replace('ک', 'ك');
            if(unigramLexicon.ContainsKey(word))
                wordIndex = unigramLexicon[word];
            return wordIndex;
        }
        private Unigram FindUnigramItem(int id)
        {
            Unigram retUnigram = new Unigram();

            unsafe
            {
                Unigram* unigrams = (Unigram*)arrUniGrams;
                retUnigram = unigrams[id - 1];
            }
            return retUnigram;
        }
        private Bigram FindBigramItem(int id, int startIndex, int length)
        {
            Bigram retBigram = new Bigram();
            bool bFound = false;
            int min = startIndex;
            int max = startIndex + length;
            int index = min + (max - min) / 2;
            unsafe
            {
                Bigram* bigrams = (Bigram*)arrBiGrams;
                while (!bFound &&  max >= min)
                {
                    if (bigrams[index].secondWordIndex == id)
                    {
                        bFound = true;
                        retBigram = bigrams[index];
                    }
                    else if (bigrams[index].secondWordIndex < id)
                    {
                        min = index + 1;
                        index =  min  + (max - min) / 2;
                    }
                    else
                    {
                        max = index - 1;
                        index = min  + (max - min) / 2;
                    }
                }

            }
            return retBigram;
        }
        private double FindBigramProbability(string firstWord, string secondWord)
        {
            double retProb = 0.0;
            int firstId = unigramLexicon[firstWord];
            int secondId = unigramLexicon[secondWord];
            Unigram firstUnigram = FindUnigramItem(firstId);
            Bigram foundBigram = FindBigramItem(secondId, firstUnigram.firstIndex, firstUnigram.length);
            retProb = foundBigram.prob;
            return Math.Pow(10,  retProb);
        }
        void LoadBigrams()
        {
            string strBigramBinary = @"khabari_hamcode_wordlist_DIC_TARKIB_BigramBinary.lm";
            string strUnigramBinary = @"khabari_hamcode_wordlist_DIC_TARKIB_UnigramBinary.lm";
            string strUnigramStrings = @"khabari_hamcode_wordlist_DIC_TARKIB_UnigramStrings.lm";
            StreamReader srUni = new StreamReader(strUnigramBinary);
            System.IO.BinaryReader brUni = new BinaryReader(srUni.BaseStream);
            StreamReader srBi = new StreamReader(strBigramBinary);
            StreamReader srUniStrings = new StreamReader(strUnigramStrings);
            System.IO.BinaryReader brBi = new BinaryReader(srBi.BaseStream);
            string strWord = "";
            unigramLexicon = new SortedList<string, int>();
            while (!srUniStrings.EndOfStream)
            {
                strWord = srUniStrings.ReadLine();
                string[] values = strWord.Split(new char[] { '\t' });
                if (values.Length >= 2)
                {
                    unigramLexicon.Add(values[0], int.Parse(values[1]));
                }
            }
            srUniStrings.Close();
            unsafe
            {
                arrUniGramsCount = srUni.BaseStream.Length / sizeof(Unigram);
                arrBiGramsCount = srBi.BaseStream.Length / sizeof(Bigram);
                byte[] bytes = brUni.ReadBytes((int)srUni.BaseStream.Length);
                {
                    arrUniGrams = Marshal.AllocHGlobal((int)srUni.BaseStream.Length);
                    Marshal.Copy(bytes, 0, arrUniGrams, (int)srUni.BaseStream.Length);
                    Unigram* unigrams = (Unigram*)arrUniGrams;
                    unigrams[0].length = 0;

                }
                byte[] Bibytes = brBi.ReadBytes((int)srBi.BaseStream.Length);
                {
                    arrBiGrams = Marshal.AllocHGlobal((int)srBi.BaseStream.Length);
                    Marshal.Copy(Bibytes, 0, arrBiGrams, (int)srBi.BaseStream.Length);
                    Bigram* bigrams = (Bigram*)arrBiGrams;
                    bigrams[0].prob = 0;

                }
                FindBigramProbability("مذكور", "منفك");
                FindBigramProbability("<s>", "ابا");
            }


        }
        private void button29_Click(object sender, EventArgs e)
        {
            LoadBigrams();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            fb.ShowNewFolderButton = true;
            fb.RootFolder = Environment.SpecialFolder.MyComputer;
            if (fb.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtSVMInputFolder.Text = fb.SelectedPath;
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            if (sf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtSVMOutputFile.Text = sf.FileName;
            }
            
        }

        private void button32_Click(object sender, EventArgs e)
        {
            //txtSVMInputFolder.Text
            string[] filenames = Directory.GetFiles(txtSVMInputFolder.Text, "*.bmp", SearchOption.AllDirectories);
            string strExceptions = "";
            foreach (var path in filenames)
            {
               // string path = "D:\\fontSize\\thining\\16sizes\\16Sizes\\btraffic_REG\\_letter_\\_g2_.bmp";
                try
                {
                    string whichModel = "";
                    string letterName = Path.GetFileNameWithoutExtension(path);
                    if (path.Contains("_letter_"))
                    {
                        whichModel = txtSVMOutputFile.Text + "_letter_";
                    }
                    else if(path.Contains("_letter"))
                    {
                        whichModel = txtSVMOutputFile.Text + "_letter";
                    }
                    else if(path.Contains("letter_"))
                    {
                        whichModel = txtSVMOutputFile.Text + "letter_";
                    }
                    else 
                    {
                        if (letterName[0] == '_' && letterName.Substring(1).Contains("_"))
                        {
                            whichModel = txtSVMOutputFile.Text + "_letter_";
                        }
                        else if (letterName[0] == '_')
                        {
                            whichModel = txtSVMOutputFile.Text + "_letter";
                        }
                        else if (letterName.Substring(1).Contains("_"))
                        {
                            whichModel = txtSVMOutputFile.Text + "letter_";
                        }
                        else
                        {
                            whichModel = txtSVMOutputFile.Text + "letter";
                        }
                    }
                    if (FeatureSVMTrain(path, whichModel) > 0)
                    {
                        //MessageBox.Show("This image is Corrupt:\n" + path);
                    }
                }
                catch (Exception ex)
                {
                    strExceptions += Path.GetFileNameWithoutExtension(path) + "\n";
                }
            }
            StreamWriter sw = new StreamWriter("D:\\exceptions.txt");
            sw.Write(strExceptions);
            sw.Close();

        }

        private void button34_Click(object sender, EventArgs e)
        {
            OpenFileDialog sf = new OpenFileDialog();
            if (sf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtInputImage.Text = sf.FileName;
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            IntPtr ptr = IntPtr.Zero ;
            LoadSVMModels(ref ptr);
            string strOcrText = "";
            GetOCR(ref ptr, txtInputImage.Text, ref strOcrText);
            MessageBox.Show(strOcrText);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            //FontDialog fontDlg = new FontDialog();
            //if (fontDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputFile = txtOutputFile.Text;
                outputFolder = txtOutputFolder.Text;
                string[] inputFiles = txtInputFile.Text.Split(';');
                List<string> distinctiveWords = new List<string>();
                List<string> connectedSegsOfWords = new List<string>();
                Dictionary<string, int> dic = new Dictionary<string,int>();
                string content = "";
                for (int i = 0; i < inputFiles.Length; i++)
                {
                    if (!File.Exists(inputFiles[i])) continue;
                    StreamReader sr = new StreamReader(inputFiles[i], Encoding.GetEncoding(1256));
                    content = sr.ReadToEnd();

                    content = content.Replace('\r', ' ');
                    content = content.Replace('\n', ' ');
                    string[] arrWords = content.Split(' ');
                    for (int j = 0; j < arrWords.Length; j++)
                    {
                        if (arrWords[j] != "")
                        if (!distinctiveWords.Contains(arrWords[j]))
                            distinctiveWords.Add(arrWords[j]);
                    }
                    sr.Close();
                }
                int distinctiveWordsCount = distinctiveWords.Count;
                int lastSaved = 0;
                StreamWriter sw = new StreamWriter("connected.txt", false, Encoding.UTF8);
                sw.Close();
                for (int i = 0; i < distinctiveWordsCount; i++)
                {
                    string[] arrConnectedStrings = GetCode2(distinctiveWords[i]);
                    for (int j = 0; j < arrConnectedStrings.Length; j++)
                    {
                        if (!connectedSegsOfWords.Contains(arrConnectedStrings[j]))
                            connectedSegsOfWords.Add(arrConnectedStrings[j]);
                    }
                    if (connectedSegsOfWords.Count > 0 && lastSaved < connectedSegsOfWords.Count && connectedSegsOfWords.Count % 1000 == 0)
                    {
                        sw = new StreamWriter("connected.txt", true, Encoding.UTF8);
                        for (int m = 0; m < 1000; m++)
                        {
                            sw.WriteLine(connectedSegsOfWords[m + lastSaved]);
                        }
                        sw.Close();
                        lastSaved += 1000;
                    }
                }
                sw = new StreamWriter("connected.txt", true, Encoding.UTF8);
                int end = connectedSegsOfWords.Count;
                for (int m = lastSaved; m < end; m++)
                {
                    sw.WriteLine(connectedSegsOfWords[m]);
                }
                sw.Close();
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            {
                outputFile = txtOutputFile.Text;
                outputFolder = txtOutputFolder.Text;
                string[] inputFiles = txtInputFile.Text.Split(';');
                List<string> distinctiveWords = new List<string>();
                List<string> connectedSegsOfWords = new List<string>();
                Dictionary<int, List<string>> dic = new Dictionary<int, List<string>>();
                string content = "";
                for (int i = 0; i < inputFiles.Length; i++)
                {
                    if (!File.Exists(inputFiles[i])) continue;
                    StreamReader sr = new StreamReader(inputFiles[i], Encoding.GetEncoding(1256));
                    content = sr.ReadToEnd();

                    content = content.Replace('\r', ' ');
                    content = content.Replace('\n', ' ');
                    string[] arrWords = content.Split(' ');
                    for (int j = 0; j < arrWords.Length; j++)
                    {
                        if (arrWords[j] != "")
                            if (!distinctiveWords.Contains(arrWords[j]))
                                distinctiveWords.Add(arrWords[j]);
                    }
                    sr.Close();
                }
                int distinctiveWordsCount = distinctiveWords.Count;
                int lastSaved = 0;
                StreamWriter sw;
                for (int i = 0; i < distinctiveWordsCount; i++)
                {
                    int len = distinctiveWords[i].Length;
                    if (!dic.ContainsKey(len))                   
                        dic.Add(len, new List<string>());
                    dic[len].Add(distinctiveWords[i]);
                }
                for (int i = 0; i < dic.Count; i++)
                {
                    int key =  dic.Keys.ElementAt(i);
                    sw = new StreamWriter("connected_" + key.ToString() + ".txt", false, Encoding.UTF8);
                    int end = dic[key].Count;
                    for (int m = 0; m < end; m++)
                    {
                        sw.WriteLine(dic[key][m]);
                    }
                    sw.Close();
                }
                
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            //string defaultPath = @"D:\Projects\A-OCR\new_ocr\Database2\Nazanin";
            string defaultPath = @"E:\ocr-dataset\1";


            string[] filenames = Directory.GetFiles(defaultPath, "*.png", SearchOption.AllDirectories);
            int dstHeight = 88;
            CvRect roi = cvlib.cvRect(0, 0, 0, 0);
            RectangleF rect;
            foreach (string file in filenames)
            {
                try
                {
                    lblCurrentFile.Text = file;
                    IplImage imgSrc = cvlib.CvLoadImage(file, cvlib.CV_LOAD_IMAGE_UNCHANGED);
                    GetFitW(ref imgSrc, out rect);
                    roi.x = (int)rect.X;
                    roi.y = (int)rect.Y;
                    roi.width = (int)rect.Width;
                    roi.height = (int)rect.Height;
                    IplImage imgDst = cvlib.CvCreateImage(cvlib.CvSize(roi.width, roi.height), 8, 4);
                    cvlib.CvSetImageROI(ref imgSrc, roi);
                    cvlib.CvCopy(ref imgSrc, ref imgDst);
                    cvlib.CvSaveImage(file, ref imgDst);
                    cvlib.CvReleaseImage(ref imgSrc);
                    cvlib.CvReleaseImage(ref imgDst);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);

                }
            }

            MessageBox.Show("Finished");

        }
        private void GetFitW(ref IplImage initImage, out RectangleF rect)
        {
            unsafe
            {
                byte* pData1 = (byte*)initImage.imageData;
                int maxY = 0, maxX = 0, minY = initImage.height - 1, minX = initImage.width - 1;
                int input_rows = initImage.height;
                int input_cols = initImage.width;
                bool bFound = false;
                for (int i = 0; i < input_rows && !bFound; i++)
                {
                    for (int j = 0; j < input_cols; j++)
                    {
                        if (pData1[i * initImage.widthStep + j * 4] == 0)
                        {
                            minY = i;
                            bFound = true;
                            break;
                        }
                    }
                }
                bFound = false;
                for (int i = input_rows - 1; i >= 0 && !bFound; i--)
                {
                    for (int j = 0; j < input_cols; j++)
                    {
                        if (pData1[i * initImage.widthStep + j * 4] == 0)
                        {
                            maxY = i;
                            bFound = true;
                            break;
                        }
                    }
                }
                bFound = false;
                for (int j = 0; j < input_cols && !bFound; j++)
                {
                    for (int i = 0; i < input_rows; i++)
                    {
                        if (pData1[i * initImage.widthStep + j * 4] == 0)
                        {
                            minX = j;
                            bFound = true;
                            break;
                        }
                    }
                }
                bFound = false;
                for (int j = input_cols - 1; j >= 0 && !bFound; j--)
                {
                    for (int i = input_rows - 1; i >= 0; i--)
                    {
                        if (pData1[i * initImage.widthStep + j * 4] == 0)
                        {
                            maxX = j;
                            bFound = true;
                            break;
                        }
                    }
                }
                rect = new RectangleF(minX, minY, maxX - minX + 1, maxY - minY + 1);
               

            }
        }
        private void button11_Click_1(object sender, EventArgs e)
        {
            {
                outputFile = txtOutputFile.Text;
                outputFolder = txtOutputFolder.Text;
                string[] inputFiles = txtInputFile.Text.Split(';');
                List<string> distinctiveWords = new List<string>();
                List<string> connectedSegsOfWords = new List<string>();
                Dictionary<int, List<string>> dic = new Dictionary<int, List<string>>();
                string content = "";
             
                int height = 300;
                int len = 0;
                string[] fonts = { "B Nazanin", "B Homa", "B Yagut", "B Lotus", "B Titr" ,
                                 "B Zar", "B Traffic","B Mitra","B Tahoma","Times New Roman",
                                 "B Yekan","B Badr"};
                for (int i = 0; i < inputFiles.Length; i++)
                {
                    if (!File.Exists(inputFiles[i])) continue;
                    len = int.Parse(inputFiles[i][inputFiles[i].LastIndexOf('_')+1].ToString());
                    StreamReader sr = new StreamReader(inputFiles[i], Encoding.GetEncoding(1256));
                    content = sr.ReadToEnd();

                    content = content.Replace('\r', ' ');
                    content = content.Replace('\n', ' ');
                    string[] arrWords = content.Split(' ');
                    for (int j = 0; j < arrWords.Length; j++)
                    {
                        if (arrWords[j] != "")
                            if (!distinctiveWords.Contains(arrWords[j]))
                                distinctiveWords.Add(arrWords[j]);
                    }
                    sr.Close();
                }
                int distinctiveWordsCount = distinctiveWords.Count;
                Bitmap bmp = new Bitmap(height * len, height);
                Graphics g = Graphics.FromImage(bmp);
                string path = txtOutputFolder.Text + "\\" + len.ToString();
                Directory.CreateDirectory(path);
                for (int i = 0; i < distinctiveWordsCount; i++)
                {
                    string txt = distinctiveWords[i];
                    txt = txt.Replace('ی', 'ي');
                    txt = txt.Replace('ک', 'ك');
                    string code = "";
                    string [] tmp = GetCode(txt);
                    for (int m = 0; m < tmp.Length; m++)
                    {
                        code += "-" + tmp[m];
                    }
                    for (int j = 0; j < fonts.Length; j++)
                    {
                        try
                        {
                            g.FillRectangle(Brushes.White, new RectangleF(0, 0, height * len, height));
                            g.DrawString(txt, new System.Drawing.Font(new FontFamily(fonts[j]), 100, FontStyle.Regular), Brushes.Black, new PointF(10, 10));
                            bmp.Save(path + "\\" + code + "_" + fonts[j].Split(new char[] { ' ' })[1] + "_R.png");

                        }
                        catch (Exception)
                        {
                            
                        }
                        try
                        {
                            g.FillRectangle(Brushes.White, new RectangleF(0, 0, height * len, height));
                            g.DrawString(txt, new System.Drawing.Font(new FontFamily(fonts[j]), 100, FontStyle.Bold), Brushes.Black, new PointF(10, 10));
                            bmp.Save(path + "\\" + code + "_" + fonts[j].Split(new char[] { ' ' })[1] + "_B.png");
                                            
                        }
                        catch (Exception)
                        {
                            
                        }
                         
                    }
                }
                

            }
        }
    }
}
