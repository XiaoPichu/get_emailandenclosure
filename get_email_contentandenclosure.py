#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import poplib
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

#解码
def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

#对头文件进行解析并回送邮件主题和时间用于后续命名 
def get_email_headers(msg):
    headers = {}
    for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['Date'] = value
            elif header == 'Subject':
                subject = decode_str(value)
                headers['Subject'] = subject
            elif header == 'From':
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                from_addr = u'%s <%s>' % (name, addr)
                headers['From'] = from_addr
            elif header == 'To':
                all_cc = value.split(',')
                to = []
                for x in all_cc:
                    hdr, addr = parseaddr(x) #parseaddr(address)是模块中专门用来解析邮件地址的函数,返回一个元祖
                    name = decode_str(hdr)
                    to_addr = u'%s <%s>' % (name, addr)
                    to.append(to_addr)
                headers['To'] = ','.join(to)
            elif header == 'Cc':
                all_cc = value.split(',')
                cc = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    cc.append(to_addr)
                headers['Cc'] = ','.join(cc) #多人阅读进行群发
    return headers
            
#获取邮件的字符编码，首先在message中寻找编码，如果没有，就在header的Content-Type中寻找
def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos+8:].strip()
    return charset

#下载正文附件
def get_content_enclosure(message,str1,str2, savepath_enclosure, savepath_subject):
    attachments = [] 
    cnt = 0
    for part in message.walk():
        filename = part.get_filename() #附件的文件加文件后缀
        content_type = part.get_content_type()
        if filename:
            filename = decode_str(filename).lower() #对附件名进行解码并且将后缀转化为小写
            data = part.get_payload(decode=True)
            attachments.append(filename)
            with open(os.path.join(savepath_enclosure,str1+'_'+str2+'_'+'_'+str(cnt)+'_'+filename), 'wb') as fx:
                fx.write(data)
            cnt += 1
        else:
            if content_type == 'text/plain':
                charset = guess_charset(part)
                if charset:
                    try:
                        content = part.get_payload(decode=True).decode(charset)
                        with open(os.path.join(savepath_subject,str1+'_'+str2+'.txt'), 'wb') as fx:
                            fx.write(content.encode('utf-8'))
                    except AttributeError:
                        print('type error')
                    except LookupError:
                        print("unknown encoding: utf-8")
            elif content_type == 'text/html':
                print('html 格式 跳过')
                continue
            else:
                continue
    return attachments 
    
#登录邮箱
def log_server(pop3,e_mail,pass_word):
    # 连接到POP3服务器，带SSL的:
    ser = poplib.POP3_SSL(pop3)   
    # POP3服务器的欢迎文字: b'+OK 237 174238271' list()响应的状态/邮件数量/邮件占用的空间大小
    print(bytes.decode(ser.getwelcome()))
    # 身份认证:
    ser.user(e_mail)
    ser.pass_(pass_word)
    return ser
    
if __name__ == '__main__':
     # 账户信息
    email = 'XXX@qq.com'
    password = 'password'
    pop3_server = 'pop.qq.com'
    server = log_server(pop3_server,email,password)#登录
    savepath_subject = '正文'
    savepath_enclosure = '附件'
    if not os.path.exists(savepath_subject): os.mkdir(savepath_subject)#如果文件夹不存在，新建对应文件夹
    if not os.path.exists(savepath_enclosure): os.mkdir(savepath_enclosure)
    month_maps = {'Jan':'1','Feb':'2','Mar':'3','Apr':'4','May':'5','Jun':'6','Jul':'7','Aug':'8','Sept':'9','Oct':'10','Nov':'11','Dec':'12'}

    msg_count, msg_size = server.stat() #stat()返回邮件数量和占用空间:
    resp, mails, octets = server.list() #list()返回邮件的响应情况，邮件的序号和大小
    index = len(mails)#邮件的总数
    print('message count:', index)
    print('message size:', msg_size, 'bytes')

    m=5#倒数第几份邮件
    #此处的循环是取最近的几封邮件
    for i in range(index-m,index+1):
        resp, lines, octets = server.retr(i)  #取邮件内容
        msg_content = b'\r\n'.join(lines).decode('utf-8','ignore')

        # 把邮件内容解析为Message对象
        msg = Parser().parsestr(msg_content)
        headers = get_email_headers(msg)
        
        str1 = headers['Date'][5:25]
        # 得到日月年时间
        dy,mh,yr,te = str1.split(' ')[:4]
        # 月变为数字
        mh = month_maps[mh]
        # 时间去除：
        te = te.replace(':','_')
        str1 = '{}_{}_{}_{}'.format(yr,mh,dy,te)
        #拼成年月日时
        str2 = headers['Subject'].replace(':',' ') #将邮件主题回送并且当做附件和正文的名字的一部分

        attachments = get_content_enclosure(msg,str1,str2,savepath_enclosure,savepath_subject)

        print('from:', headers['From'])
        print('to:', headers['To'])
        print('subject:', headers['Subject'])
        if 'cc' in headers:  print('cc:', headers['Cc'])
        print('date:', headers['Date'])
        print('attachments: ', attachments)
        print('-----------------------------')

    server.quit()

