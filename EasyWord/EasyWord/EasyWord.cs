using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSWord=Microsoft.Office.Interop.Word;
using System.IO;

namespace EasyWord
{
    public class EasyWord:IDisposable
    {
        public bool IsOpen { get; set; }
        public string FileName { get; set; }

        private MSWord._Application wordApp;

        private MSWord._Document wordDoc;

        private static Object missing = System.Reflection.Missing.Value;

        public EasyWord(string filePath)
        {
            this.FileName = filePath;
            //一些初始化的操作
            wordApp = new MSWord.Application();
            wordApp.Visible = true;
            
        }

        /// <summary>
        /// 生成word文档
        /// </summary>
        /// <param name="directory">目录</param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool CreateWord(string directory)
        {
          if(!Directory.Exists(directory))
                throw new DirectoryNotFoundException();

            wordDoc=wordApp.Documents.Add(ref missing,ref missing,ref missing,ref missing);
            object fullName=directory+"\\"+this.FileName;

            wordDoc.SaveAs(ref fullName, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);
            CloseWord();
            return true;
        }

        private void OpenWord()
        {
            this.IsOpen = true;
            if (!File.Exists(this.FileName))
                throw new FileNotFoundException();
            object fileName = this.FileName;
            wordDoc = wordApp.Documents.Open(ref fileName);
        }

        private void CloseWord()
        {
            this.IsOpen = false;
            wordDoc.Close(MSWord.WdSaveOptions.wdSaveChanges);
            wordApp.Quit();
        }

        public void Dispose()
        {
            if(wordDoc!=null)
                wordDoc.Close();
            if(wordApp!=null)
                wordApp.Quit();
            
        }

    }


}
