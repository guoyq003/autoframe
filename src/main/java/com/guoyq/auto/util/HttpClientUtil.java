package com.guoyq.auto.util;

import org.apache.commons.httpclient.Header;
import org.apache.commons.httpclient.HttpClient;
import org.apache.commons.httpclient.HttpMethodBase;
import org.apache.commons.httpclient.methods.GetMethod;
import org.apache.commons.httpclient.methods.PostMethod;
import org.apache.commons.httpclient.params.HttpClientParams;
import org.apache.commons.httpclient.params.HttpConnectionManagerParams;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class HttpClientUtil {
    //返回两个信息，code 和 返回结果
    public static Map<String, String> request(Map<String, String> baseParam, Map<String, String> headerMap) {
        Map<String, String> returnMap = new HashMap<String, String>();
        String url = (String) baseParam.get("url");
        String methodName = (String) baseParam.get("methodName");
        returnMap.put("status", "500");
        if (isEmpty(url)) {
            returnMap.put("returnBody", "url不能为空");
            return returnMap;
        }
        if (isEmpty(methodName)) {
            returnMap.put("returnBody", "methodName不能为空");
            return returnMap;
        }
        String encoding = (String) baseParam.get("encoding");
        String connectionTimeout = (String) baseParam.get("connectionTimeout");
        String soTimeout = (String) baseParam.get("soTimeout");
        String ip = (String) baseParam.get("ip");
        String portStr = (String) baseParam.get("portStr");
        String paramBody = (String) baseParam.get("paramBody");
        HttpMethodBase method = null;
        HttpClient client = new HttpClient();
        //获取方法
        if (methodName.equalsIgnoreCase("post")) {
            PostMethod method1 = new PostMethod(url);
            if (!isEmpty(paramBody)) {
                method1.setRequestBody(paramBody);
            }
            method = method1;
        } else {
            method = new GetMethod(url);
        }
        //动态代理Host
        int port = 80;
        if (!isEmpty(ip)) {
            if (isEmpty(portStr)) {
                client.getHostConfiguration().setProxy(ip, port);
            } else {
                client.getHostConfiguration().setProxy(ip, Integer.parseInt(portStr));
            }
        }
        //设置编码格式
        HttpClientParams clientParams = client.getParams();
        if (isEmpty(encoding)) {
            clientParams.setContentCharset("UTF-8");
        } else {
            clientParams.setContentCharset(encoding);
        }
        //设置超时
        HttpConnectionManagerParams params = client.getHttpConnectionManager().getParams();
        if (isEmpty(connectionTimeout)) {
            params.setConnectionTimeout(5000);
        } else {
            params.setConnectionTimeout(Integer.parseInt(connectionTimeout));
        }
        if (isEmpty(soTimeout)) {
            params.setSoTimeout(1000 * 60);
        } else {
            params.setSoTimeout(Integer.parseInt(connectionTimeout));
        }

        //设置消息头
        if (null != headerMap) {
            for (String key : headerMap.keySet()) {
                method.setRequestHeader(key, headerMap.get(key));
            }
        }
        int status = 0;
        String returnBody = "";
        try {
            status = client.executeMethod(method);
            returnBody = method.getResponseBodyAsString();
            method.releaseConnection();
        } catch (IOException e) {
            e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
        }
        if (status != 200) {
            if (status == 303) {
                Header header = method.getResponseHeader("location"); // 跳转的目标地址是在 HTTP-HEAD 中的
                String newUri = header.getValue();
                returnMap.put("newUri",newUri);
                return returnMap;
            }
            String exceptionMsg = codeMap.get(status);
            if (isEmpty(exceptionMsg)) {
                returnMap.put("returnBody", "有异常错误，百度查询一下code码的含义" + ",,,returnBody:" + returnBody);
            } else {
                returnMap.put("returnBody", codeMap.get(status) + ",,,returnBody:" + returnBody);
            }
            returnMap.put("status", status + "");
        } else {
            returnMap.put("returnBody", returnBody);
            returnMap.put("status", status + "");
        }
        return returnMap;
    }

    static Map<Integer, String> codeMap = new HashMap<Integer, String>();

    static {
        codeMap.put(401, "执行方法没有授权");
        codeMap.put(405, "执行方法不对,请确认是get还是post请求");
        codeMap.put(415, "请确认content-type类型是否正确");
    }

    private static boolean isEmpty(String n) {
        boolean f = true;
        if (null != n && !"".equals(n)) {
            f = false;
        }
        return f;
    }
}
