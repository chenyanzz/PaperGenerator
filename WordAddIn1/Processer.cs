using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordAddIn1
{
    public class Processer
    {

        Application app;
        Document doc_fmt, doc_out, doc_head;
        ParagraphFormat text_format;
        List<ParagraphFormat> title_format = new List<ParagraphFormat>();
        Font text_font;
        List<Font> title_font = new List<Font>();
        List<string> papersIncluded = new List<string>();

        List<UInt32> label_index = new List<UInt32>();

        string MdDirectory;

        public Processer(Application app)
        {
            this.app = app;
        }

        void setFont(Font toSet, Font fmt)
        {
            toSet.Bold = fmt.Bold;
            toSet.Size = fmt.Size;
            toSet.Name = fmt.Name;
            toSet.Color = fmt.Color;
        }

        void readFormat(string path)
        {

            doc_fmt = app.Documents.Open(path);
            var paragraphs = doc_fmt.Paragraphs;
            text_format = paragraphs.First.Format;
            text_font = paragraphs.First.Range.Font;
            for (int i = 2; i <= paragraphs.Count; i++)
            {
                title_format.Add(paragraphs[i].Format);
                title_font.Add(paragraphs[i].Range.Font);
            }
        }

        void readMarkdown(string path)
        {
            System.IO.StreamReader reader = null;
            reader = new System.IO.StreamReader(path);

            int line_id = 0;
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                parseLine(line, ++line_id);
            }
            addPaperInclude(++line_id);
        }

        private void addTitle(string line)
        {
            int depth;
            for (depth = 0; depth < line.Length && line[depth] == '#'; depth++) ;
            string text = "";

            //eg. '##' so, depth = 2

            //if 1
            //to 1.0 
            while (depth < label_index.Count)
            {
                label_index.RemoveAt(label_index.Count - 1);
            }

            //if 1.1.3
            //to 1.1
            while (depth > label_index.Count)
            {
                label_index.Add(0);
            }

            label_index[depth - 1]++;

            text += getSectionName();

            if (line.Substring(depth).Length != 0)
            {
                if (line[depth] != ' ') text += " ";
            }

            text += line.Substring(depth);

            var font = title_font[depth - 1];
            setFont(app.Selection.Font, font);
            app.Selection.ParagraphFormat = title_format[depth - 1];
            app.Selection.TypeText(text);
        }


        int pic_id = 0, table_id = 0;
        void parseLine(string line, int id)
        {
            if (line.StartsWith("~")) return;//战术空行。
            if (line.Trim() == "") return;//战术空行。

            setFont(app.Selection.Font, text_font);
            app.Selection.ParagraphFormat = text_format;
            app.Selection.ParagraphFormat.IndentFirstLineCharWidth(2);

            if (line.Length != 0)
            {
                switch (line[0])
                {
                    case '#':
                        table_id = 0;
                        pic_id = 0;
                        addTitle(line);
                        break;

                    case '!':
                        string text = line;
                        //![aaa]("")
                        line.Replace("(\"", "(");
                        line.Replace("\")", ")");

                        //[bcd] --> bcd
                        Func<string, char, char, string> string_cutter =
                            (string str, char l, char r) =>
                            {
                                int pr, pl;
                                pl = line.IndexOf(l);
                                pr = line.IndexOf(r);
                                if (pr == -1 || pr == -1 || pl >= pr) return null;

                                return line.Substring(pl + 1, pr - pl - 1);
                            };

                        string pic_name = string_cutter(line, '[', ']');
                        string pic_path = string_cutter(line, '(', ')');
                        pic_name = pic_name.Trim('\"');
                        pic_path = pic_path.Trim('\"');

                        if (!System.IO.Path.IsPathRooted(pic_path))
                        {
                            pic_path = System.IO.Path.Combine(MdDirectory, pic_path);
                        }
                        if (pic_name == null || pic_path == null)
                        {
                            throw new Exception("Error Picture Definition in line" + id);
                        }

                        bool isTable = false;
                        if (line.EndsWith("table") ||
                            line.EndsWith("t") ||
                            line.EndsWith("tbl")) isTable = true;

                        pic_name = (isTable ? "表" : "图") + getSectionName() + "-" + (isTable? (++table_id):(++pic_id)) + " " + pic_name;

                        Action addPic = delegate
                        {
                            app.Selection.ParagraphFormat.IndentFirstLineCharWidth(0);
                            app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            app.Selection.ParagraphFormat.FirstLineIndent = 0;
                            app.Selection.ParagraphFormat.LeftIndent = 0;
                            var pic = app.Selection.InlineShapes.AddPicture(pic_path);
                            pic.LockAspectRatio = Office.MsoTriState.msoTrue;
                            //System.Windows.Forms.MessageBox.Show(pic_name+"$"+pic.Width);
                            //max width:415.3
                            const float maxWidth = 320;
                            if(pic.Width > maxWidth) pic.Width = maxWidth;
                        };
                        Action addInfo = delegate
                        {
                            app.Selection.ParagraphFormat = text_format;
                            setFont(app.Selection.Font, text_font);
                            app.Selection.Font.Size -= 2;
                            app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            app.Selection.TypeText(pic_name);
                        };


                        if (isTable) addInfo(); else addPic();
                        app.Selection.TypeParagraph();
                        if (isTable) addPic(); else addInfo();
                        break;

                    default:
                        //in case of `paper`
                        char[] squ = { '`' };
                        var parts = line.Split(squ);

                        if (parts.Length % 2 != 1)
                        {
                            throw new Exception("Error Paper Include in line" + id);
                        }

                        for (int i = 0; i < parts.Length; i++)
                        {
                            if (i % 2 == 0) addPureText(parts[i]);
                            else
                            {
                                papersIncluded.Add(parts[i]);
                                addPureText("[" + papersIncluded.Count + "]", true);
                            }
                        }
                        break;
                }
            }
            app.Selection.ParagraphFormat.Space15();
            app.Selection.TypeParagraph();
            app.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        private void addPureText(string line, bool isSuperScript = false)
        {
            setFont(app.Selection.Font, text_font);
            app.Selection.ParagraphFormat = text_format;
            app.Selection.ParagraphFormat.IndentFirstLineCharWidth(2);
            app.Selection.Font.Superscript = isSuperScript ? 1 : 0;
            app.Selection.TypeText(line);
        }

        void addPaperInclude(int lineId)
        {
            parseLine("# 参考文献", lineId);
            for (int i = 0; i < papersIncluded.Count; i++)
            {
                addPureText("[" + (i + 1) + "]" + papersIncluded[i]);
                app.Selection.TypeParagraph();
            }
        }

        string getSectionName(char sep = '.')
        {
            string name = "";
            int depth = label_index.Count;
            for (int i = 0; i < depth - 1; i++)
            {
                name += label_index[i].ToString() + sep.ToString();
            }
            name += label_index[depth - 1].ToString();

            return name;
        }

        void writeDocx(string path)
        {
            doc_out.SaveAs2(path);
        }

        public bool process(string FomatFilePath, string MdDilePath, string HeadingDocxPath, string OutputFilePath)
        {
            label_index.Clear();
            papersIncluded.Clear();
            try
            {
                MdDirectory = System.IO.Path.GetDirectoryName(MdDilePath);
                readFormat(FomatFilePath);
                doc_out = app.Documents.Add();
                if (HeadingDocxPath != null)
                {
                    doc_head = app.Documents.Open(HeadingDocxPath);
                    doc_head.Activate();
                    app.Selection.WholeStory();
                    app.Selection.Copy();
                    doc_out.Activate();
                    app.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
                    app.Selection.InsertNewPage();
                    doc_head.Close();
                }
                doc_out.Activate();
                readMarkdown(MdDilePath);
                writeDocx(OutputFilePath);
                doc_fmt.Close();
                doc_out.Activate();
                return true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error Occured:\n" + e.Message + "\n" + e.StackTrace);
                return false;
            }
        }
    }
}
