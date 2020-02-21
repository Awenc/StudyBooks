using ICSharpCode.SharpZipLib.Zip;

/// <summary>
/// doc文档工具类
/// </summary>
public class DocUtility
{
    /// <summary>
    /// doc文档加载工具
    /// </summary>
    private IDocLoadHelper m_DocLoadHelper;

    public DocUtility ()
    {
        m_DocLoadHelper = null;
    }

    /// <summary>
    /// 读取doc文档中的数据
    /// </summary>
    /// <param name="filePath"></param>
    public void ReadDoc(string filePath)
    {
        if (m_DocLoadHelper == null)
        {
            throw new System.Exception("m_DocLoadHelper is null");
        }

        byte[] docBytes = m_DocLoadHelper.LoadDoc(filePath);
        
        MemoryStream streamm = new MemoryStream(docBytes);

        byte[] ReadData = new byte[2048];

        using (ZipInputStream s = new ZipInputStream(streamm)) //zip流读取 定义一个范围，在范围结束时处理对象。
        {
            ZipEntry theEntry;  //文档 zip 包含的所有文件信息
            while ((theEntry = s.GetNextEntry()) != null)
            {
                Debug.log("name:"+theEntry.name); // 文件名

                //eg：读取word/document.xml中所有的数据  （读取当前ZipEntry中的流数据）
                if (theEntry.Name == "word/document.xml")
                {
                    int readSize = s.Read(ReadData, 0, ReadData.Length); //将流中的数据读取到ReadData中
                    string documentTxt = "";
                    while (readSize > 0)
                    {
                        if (readSize < ReadData.Length)
                        {
                            documentTxt += Encoding.UTF8.GetString(ReadData, 0, readSize);
                            readSize = -1;
                        }
                        else
                        {
                            documentTxt += Encoding.UTF8.GetString(ReadData, 0, ReadData.Length);
                            readSize = s.Read(ReadData, 0, ReadData.Length);
                        }
                    }

                    Debug.log(documentTxt);
                }
            }

            s.Close();
        }
    }


    /// <summary>
    /// 替换doc文档中的文本内容并保存
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="outPath">输出路径</param>
    /// <param name="changeMap">替换文本字典</param>
    public void ChangeDocument(string filePath, string outPath, Dictionary<string, string> changeMap)
    {
        if (m_DocLoadHelper == null)
        {
            throw new System.Exception("m_DocLoadHelper is null");
        }

        byte[] docBytes = m_DocLoadHelper.LoadDoc(filePath);

        MemoryStream streamm = new MemoryStream(docBytes);

        byte[] outData = new byte[2048];

        ZipInputStream s = new ZipInputStream(streamm);

        ZipOutputStream outputStream = new ZipOutputStream(File.Create(OutPutPath));
        
        outputStream.SetLevel(0);

        string documentTxt = "";
        ZipEntry theEntry;
        while ((theEntry = s.GetNextEntry()) != null)
        {
            ZipEntry theEntryOut = new ZipEntry(theEntry.Name);
            Log.Info(theEntry.Name);

            outputStream.PutNextEntry(theEntryOut);
            if (theEntry.Name == "word/document.xml")
            {
                int outSize1 = s.Read(outData, 0, outData.Length);
                while (outSize1 > 0)
                {
                    if (outSize1 < outData.Length)
                    {
                        documentTxt += Encoding.UTF8.GetString(outData, 0, outSize1);
                        outSize1 = -1;
                    }
                    else
                    {
                        documentTxt += Encoding.UTF8.GetString(outData, 0, outData.Length);
                        outSize1 = s.Read(outData, 0, outData.Length);
                    }
                }
                byte[] txt = new byte[0];
                foreach (KeyValuePair<string, string> keyValuePair in changeMap)
                {
                    documentTxt = Regex.Replace(res, keyValuePair.Key, keyValuePair.Value);
                }
                txt = Utility.Converter.GetBytes(documentTxt);
                outputStream.Write(txt, 0, txt.Length);
            }
            s.Close();
            outputStream.Finish();
            outputStream.Close();
        }
    }

    public void ChangeImage()
    {
        public static void ReadDoc()
        {

            if (m_DocLoadHelper == null)
            {
                throw new System.Exception("m_DocLoadHelper is null");
            }

            byte[] docBytes = m_DocLoadHelper.LoadDoc(filePath);

            MemoryStream streamm = new MemoryStream(docBytes);

            byte[] outData = new byte[2048];

            //todo 输出操作

            string OutPutPath = Path.Combine(Application.dataPath + "/Resources/Test.docx");

            ZipOutputStream outputStream = new ZipOutputStream(File.Create(OutPutPath));
            outputStream.SetLevel(0);

            string documentTxt = "";

            using (ZipInputStream s = new ZipInputStream(streamm))
            {
                ZipEntry theEntry;
                while ((theEntry = s.GetNextEntry()) != null)
                {
                    ZipEntry theEntryOut = new ZipEntry(theEntry.Name);
                    Log.Info(theEntry.Name);

                    outputStream.PutNextEntry(theEntryOut);

                    if (theEntry.Name == "word/media/image2.png")
                    {
                        Texture2D texture2D = Resources.Load<Texture2D>("awen");
                        byte[] img = texture2D.EncodeToPNG();
                        outputStream.Write(img, 0, img.Length);
                    }
                    else if (theEntry.Name == "word/media/image3.png")
                    {
                        Texture2D texture2D = Resources.Load<Texture2D>("awen");
                        byte[] img = texture2D.EncodeToPNG();
                        outputStream.Write(img, 0, img.Length);
                    }
                    else if (theEntry.Name == "word/media/image11.png")
                    {
                        Texture2D texture2D = Resources.Load<Texture2D>("SignName");
                        byte[] img = texture2D.EncodeToPNG();
                        outputStream.Write(img, 0, img.Length);
                    }
                    else
                    {
                        int outSize = s.Read(outData, 0, outData.Length);
                        while (outSize > 0)
                        {
                            if (outSize < outData.Length)
                            {
                                outputStream.Write(outData, 0, outSize);
                                outSize = -1;
                            }
                            else
                            {
                                outputStream.Write(outData, 0, outData.Length);
                                outSize = s.Read(outData, 0, outData.Length);
                            }
                        }
                    }
                    
                }

                s.Close();
                outputStream.Finish();
                outputStream.Close();
            }
        }
    }


    //设置加载doc帮助类
    public void SetDocLoadHelper(IDocLoadHelper docLoadHelper)
    {
        if(docLoadHelper == null)
        {
            throw new System.Exception("docLoadHelper is null");
        }
        
        m_DocLoadHelper = docLoadHelper;
    }

}