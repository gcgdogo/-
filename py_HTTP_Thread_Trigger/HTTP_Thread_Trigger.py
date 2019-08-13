"""模块通过Fire实现命令行参数，运行时建议使用长标注 *.py --Name1=Value1 --Name2=Value2
将 HTTP_Thread_Trigger.start_server() 放在主模块最后以启动服务

建议在运行时对默认参数 [ --Server/Name=HTTP_Thread_Trigger  --Server/PortSeed=8400 ] 进行设置 便于链接
为避免端口冲突，运行时程序会按 [ port = PortSeed + i*567 ] 进行测试 i = 0 ~ 4，链接时需要进行尝试。
建议通过遍历端口并检测 [ http://127.0.0.1:[port]/Parameter/Server/Name ] 来测试实际端口，如有必要，可以通过随机设定 --Server/Name 来避免错误链接

在VBA中 WinHTTP 和 XMLHttp 在调用均已成功，初步测量响应时间 WinHTTP:2.2ms , XMLHttp:4.2ms
不要使用 [ localhost ] 进行访问，速度缓慢，整齐的卡在一秒一次，不知道为啥
XMLHttp 的 GET 没有尝试成功，且每一次send之后需要重新进行open不然会有各种错误
WinHTTP 的 GET 和 POST 没有发现明显区别
进行高频连续访问时可能会传输失败， vba中 on errror goto 无法处理 ， WinHTTP 十万次连续访问，在重复使用同一次open的情况下访问成功，平均1.516ms

目前效率最高的方法：open一次，send中追加DoEvents降低错误（成功循环40万次，1.59ms）
:    Dim WinReq As New WinHttpRequest
:    WinReq.Open "GET", "http://127.0.0.1:8400/Execute"
:    For i = 1 To 400000
:        DoEvents
:        WinReq.Send
:    Next"""

from werkzeug.serving import run_simple

from werkzeug.wrappers import Response

#import time #测试时使用
#import tqdm #测试时使用
import socket #用于检测端口
import fire #实现命令行参数

import markdown

def test_Thread(dict_parameter):
    if not('Execution_Count' in dict_parameter): dict_parameter['Execution_Count'] = 0
    #for i in dict_parameter:
        #print('{} = {}'.format(repr(i),repr(dict_parameter[i])))
    dict_parameter['Execution_Count'] = dict_parameter['Execution_Count'] + 1
    return "test_Thread Executed"

dict_parameter={
    'Server/Name':'HTTP_Thread_Trigger',
    'Server/PortSeed':8400,
    'Server/Port':8400,
    'Thread/Running':False,
    'Thread/Source':test_Thread,
    'Thread/str_Path':'',
    'Thread/str_Value':''
}

dict_ResponseText={
    '未找到对应命令' : 'ERR:Command Not Found 未找到对应命令',
    '不能多重启动' : 'ERR:Thread is Running Already 不能多重启动',
    '无法读取变量' : 'ERR:Patameter Not Found 未找到/无法读取变量'
}

saved_enivrion={}

def fun_Dict_to_HTML(x_dict,x_title = "Dict_to_HTML"):
    str_FrontEnd = """
    <head><meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * {color:#3e3e3e;}
        p{letter-spacing: 1px !important;line-height: 1.7 !important;min-height: 1em !important;box-sizing: border-box !important;
            word-wrap: break-word !important;text-align: justify;margin: 23.7px 0 !important;}
        blockquote {border-left: 10px solid rgba(128,128,128,0.075);
            background-color: rgba(128,128,128,0.05);padding: 13px 15px !important;margin: 0px;}
        blockquote p {color: #898989;margin: 0px;}
        strong {font-weight: normal;color: #A23400;}
        body {font-family: "Helvetica Neue",Helvetica,"Hiragino Sans GB","Microsoft YaHei",Arial,sans-serif;font-size: 15px;}
        pre {background-color: #f8f8f8;border-radius: 3px;word-wrap: break-word;padding: 12px 13px;font-size: 13px;color: #898989;}
        h1,h2,h3,h4,h5,h6 {word-break: break-all !important;margin: 20px 0;line-height: 1.2;text-align: left;font-weight: bold;padding-left: 15px;}
        h1 {border-left: 6px solid #71BA51 !important;font-size: 20px !important;}
        h2 {border-left: 4px solid #71BA51 !important;font-size: 18px !important;}
        h3 {border-left: 4px solid #71BA51 !important;}
        a {color: #4183C4 !important;text-decoration: none !important;}
        ul, ol {padding-left: 30px;}
        li {line-height: 24px;}
        hr {height: 4px;padding: 0;margin: 16px 0;background-color: #e7e7e7;border: 0 none;overflow: hidden;
            box-sizing: content-box;border-bottom: 1px solid #ddd;}
        code {font-family: "Helvetica Neue",Helvetica,"Hiragino Sans GB","Microsoft YaHei",Arial,sans-serif;
            background-color: #d0d0d0;border-radius: 3px;word-wrap: break-word;overflow: scroll;padding: 2px 6px;font-size: 13px;color: #4F4F4F;}
    </style></head>
    <body>"""

    str_BackEnd = """</body>"""

    # x_dict 重新分组
    x_dict_grouped = {
        "wsgi.":{},
        "HTTP_":{},
        "SERVER_":{},
        "Server/":{},
        "Thread/":{},
        "Parameters":{}
    }
    for i in x_dict:
        for j in x_dict_grouped:
            if j == i[0:len(j)] : break
        x_dict_grouped[j][i]=x_dict[i]
    
    x_string='# {}   \n  ----'.format(x_title)

    for i_Group in x_dict_grouped:

        if x_dict_grouped[i_Group]=={} : continue

        x_string = x_string + '\n##  {}'.format(i_Group)
        x_dict_group_i = x_dict_grouped[i_Group]

        for i in x_dict_group_i:
            i_value=x_dict_group_i[i]
            i_head=repr(i_value)[0]

            format_text=' \n  *  **{}:**  {}  '
            if '<([{'.find(i_head)>=0:
                format_text=' \n  *  **{}:**  ``` {} ```  '
            if i_head=="'" or i_head=='"':
                #print(repr(i_value))
                #print(str(i_value).find('\n'))
                if str(i_value).find('\n')>=0 :
                    format_text=" \n  *  **{}:**  \n ``` {} ``` \n"
                    i_value = i_value.replace('\n','  ```   \n   ```  ')
                else:
                    format_text=" \n  *  **{}:**  ``` '{}' ```  "
            
            x_string = x_string + format_text.format(i,str(i_value))
    #print(x_string)
    x_string =  markdown.markdown(x_string)
    return str_FrontEnd + x_string + str_BackEnd


def application(environ, start_response):
    
    dict_doc={}

    str_Command=environ['PATH_INFO'].split('/')[1]
    str_Path=environ['PATH_INFO'][2+len(str_Command):]
    str_Value=environ['QUERY_STRING']


    dict_doc['Enivron'] = """命令: localhost:[port]/Enivron/['Read_Only'/*]
    显示当前请求的environ信息，用于调试
    /Environ/Read_Only 为只读模式，不会更新显示结果，即显示上一次访问的相关数值，后面的其他url会直接忽略"""

    global saved_enivrion
    #检测 'Read_Only' 只读模式
    if environ['PATH_INFO'] != '/Environ/Read_Only' : saved_enivrion = environ.copy()
    
    if str_Command == 'Environ':
        return_string = fun_Dict_to_HTML(saved_enivrion,"Enivron:请求信息")
        response = Response( return_string, mimetype='text/html')
        return response(environ, start_response)


    #调用函数
    dict_doc['Execute'] = """命令: localhost:[port]/Execute/[Thread/str_Path]?[Thread/str_Value]
    执行函数 return_string = dict_parameter['Thread/Source'](dict_parameter)
    Execute后面的url会存储在如上所示的两个变量中，字符支持有限，中文直接乱码。暂不支持POST方式。"""
    if str_Command == 'Execute':
        if dict_parameter['Thread/Running'] == True:
            response = Response(dict_ResponseText['未找到对应命令'], mimetype='text/plain')
            return response(environ, start_response)
        dict_parameter['Thread/str_Path']=str_Path
        dict_parameter['Thread/str_Value']=str_Value

        #开始运行
        dict_parameter['Thread/Running'] = True
        return_string = dict_parameter['Thread/Source'](dict_parameter)
        dict_parameter['Thread/Running'] = False

        #print(return_string)
        response = Response(return_string, mimetype='text/plain')
        return response(environ, start_response)


    dict_doc['Parameter'] = """命令: localhost:[port]/('','Para','Parameter')
    空白默认命令
    显示dict_parameter内容
    
    命令: localhost:[port]/('Para','Parameter')/[Parameter_Name]
    如果方法为 POST 则对dict_parameter[Parameter_Name]进行设定，系统会尝试对 POST内容 进行 eval() 处理，文本需要用 ' 括起来
    必定返回dict_parameter[Parameter_Name]的值"""
    if (str_Command == 'Parameter' or str_Command == 'Para') and str_Path!='':
        if environ['REQUEST_METHOD']=='POST':
            print("Got_Post : {}".format(str_Path))
            POST_val = environ['wsgi.input'].read(int(environ['CONTENT_LENGTH'])).decode('utf-8')
            try:
                dict_parameter[str_Path] = eval(POST_val)
            except:
                dict_parameter[str_Path] = POST_val

        return_string = str(dict_parameter[str_Path])
        response = Response(return_string, mimetype='text/plain')
        return response(environ, start_response)
    
    if str_Command == '' or str_Command == 'Parameter' or str_Command == 'Para':
        return_string = fun_Dict_to_HTML(dict_parameter,'Parameter:变量列表')
        response = Response(return_string, mimetype='text/html')
        return response(environ, start_response)


    dict_doc['Help'] = """命令: localhost:[port]/('Help','H','h')
    显示本帮助信息"""
    #其他说明放这里

    dict_doc['ELSE:```__doc__```'] = __doc__

    dict_doc['ELSE:```ResponseText```'] = str(dict_ResponseText)

    if dict_parameter['Thread/Source'].__doc__ :
        dict_doc['ELSE:```Thread/Source.__doc__```'] = dict_parameter['Thread/Source'].__doc__

    if str_Command == 'H' or str_Command == 'h' or str_Command == 'Help':
        print(dict_doc)
        return_string = fun_Dict_to_HTML(dict_doc,'Help:帮助文档')
        response = Response(return_string, mimetype='text/html')
        return response(environ, start_response)


    #调用结束
    response = Response(dict_ResponseText['未找到对应命令'], mimetype='text/plain')
    return response(environ, start_response)

def get_blankport(PortSeed):
    test_link = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    test_link.settimeout(0.2)
    for i in range(5):

        #端口不成功就以 567 递增
        test_port=PortSeed + i*567

        try:
            test_link.connect(('localhost', test_port))
        except:
            return test_port
    print ('无法获取空白端口 可能是已有多次运行并未关闭服务')
    return -999999

def main(**kwargs):
    #将命令行参数整合进 dict_parameter
    for i in kwargs:
        dict_parameter[i] = kwargs[i]
    #获取可用端口
    dict_parameter['Server/Port'] = get_blankport(dict_parameter['Server/PortSeed'])

    run_simple('localhost', dict_parameter['Server/Port'], application, use_reloader=True , threaded=True)

def start_server():
    print(__doc__)
    fire.Fire(main)

if __name__ == '__main__':
    start_server()