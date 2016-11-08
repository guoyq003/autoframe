package com.guoyq.auto.util;
import com.guoyq.auto.config.Constants;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

import java.io.File;
import java.io.IOException;

public class Log {

    protected final Logger logger = Logger.getLogger(this.getClass());
    private  static Log log;

    /**
     * 构造方法，初始化读入log4j.xml
     */
    public Log()
    {

        //返回读取指定资源的输入流
        //InputStream is=this.getClass().getResourceAsStream("/resource/res.txt");
        //InputStream is=当前类.class.getResourceAsStream("XX.config");
        //BufferedReader br=new BufferedReader(new InputStreamReader(is));

        File   directory =   new   File(Constants.LogProperty);
        try
        {
            directory.getCanonicalPath();
            //	URL fileUrl = this.getClass().getResource("");
            System.out.println("Load Config fileName:"+directory.getCanonicalPath());
            PropertyConfigurator.configure(directory.getCanonicalPath());
//            DOMConfigurator.configure(directory.getCanonicalPath());
        } catch (IOException e) {
            e.printStackTrace();
        }//取得当前路径
    }

    /**
     * 单例模式，获取Logger实例
     * @return
     */
    public static Logger getInstance()
    {
        if (log == null)
        {
            log = new Log();
        }
        return log.logger;
    }

}
