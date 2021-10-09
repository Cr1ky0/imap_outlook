import imaplib
import email
import time
import requests

def connect(user,password):
    # 链接到outlook imap端口
    conn = imaplib.IMAP4_SSL(port='993', host='outlook.office365.com')
    print('已连接服务器')
    conn.login(user,password )  # 链接指定邮箱
    print('已登陆')
    conn.select()  # 选择信箱默认inbox
    return conn

def get_index(conn,index):
    """提取FULL标准下的字符集编码所在位置

    :param conn: imap对象
    :param index:传入邮件序号 注意是data[0].split()列表下的值
    :return:返回FULL标准下字符集编码所在下标位置
    """
    type, data = conn.fetch(index, '(FULL)') # FULL标准包含了编码集
    j = 0
    for i in data[0].split():
        if(i.decode('UTF-8') == '("charset"'): # 注意解码
            return j+1
        j += 1
    return None;

def get_encode(conn,index):
    """提取该邮件字符集编码

    :param conn: imap对象
    :param index: 传入邮件序号 注意是data[0].split()列表下的值
    :return: 返回字符集编码str
    """
    char_index = get_index(conn,index)
    type, data = conn.fetch(index, '(FULL)')
    ENCODE = str(data[0].split()[char_index].decode('UTF-8'))
    encode = []
    for i in ENCODE:
        if(i.isalnum()):
            encode.append(i)
    encode = "".join(encode)
    return encode

def get_mail_title(conn,index):
    """获取邮件标题

    :param conn: imap对象
    :param index: conn.search().split()[i]，即序列表对应的序号(邮件编号)
    :return: 返回邮件标题str
    """
    encode = get_encode(conn,index)     # 获取编码
    type, data = conn.fetch(index, '(RFC822)')

    msg = email.message_from_string(data[0][1].decode(encode))  # 解码所有RFC882标准邮件信息
    sub = msg.get('subject')  # 获取标题:
    sub_encode = email.header.decode_header(sub)[0][0] # 提取未解码标题
    code = email.header.decode_header(sub)[0][1] # 提取对应解码的编码
    if isinstance(sub_encode,bytes):
        # 网络传输只能用bytes类型，故一些未转化的bytes类型要转成str
        # 而且由于bytes类型上面code提取的编码方式为None,故还是需要手动提取encode
        return sub_encode.decode(encode)
    if code is None:
        return sub_encode
    return sub_encode.decode(encode)

def mail_count(sequence):
    return len(sequence)

def main(conn):
    """主函数，主要进行邮箱状态的获取和标题的存储
        返回当前邮箱内邮件数以及标题列表
    :param conn: imap对象
    :return: 返回标题列表与当前邮箱邮件数
    """
    type, data = conn.search(None, 'ALL')# 提取所有邮件
    sequence = data[0].split()  # 邮箱序列表，注意是倒叙
    count = mail_count(sequence)  # 提取邮件数

    # 记录邮件标题(逆序后为邮件正常顺序 )
    mail_list = []
    for i in sequence:
        try:
            mail_list.append(get_mail_title(conn, i))
        except:
            print("wrong!")
            continue
    mail_list.reverse()

    return mail_list,count

def send_push(title):
    # 向中继服务器发送推送
    # 利用Android端的“消息接收”app，是基于MiPush的(酷安：唐大土土)
    # 发送post请求即可对指定别名的设备进行消息推送
    url = "https://tdtt.top/send?alias=别名&title=" + title + "&content=Outlook来信"
    res = requests.post(url=url)
    print(res.text)

if __name__ == '__main__':
    conn = connect('user','password') # 建立连接
    # 初始化邮箱
    count = 0;
    mail_list = []
    mail_list,count = main(conn)
    while True:
        # 提取当前状态
        new_count = 0
        new_mail_list = []
        new_mail_list,new_count = main(conn)
        # 当前状态与之前不同则将当前状态修改为新状态，并向服务器发送最新的邮件标题（表示推送最新邮件）
        # 这里需要解决一个问题，如果有多封邮件到达怎么推送
        if new_count != count:
            # 注意，如果在运行时删除邮件会出错
            try:
            # 推送邮件
                num = new_count - count
                for i in range(num):
                    send_push(new_mail_list[i])
                # 推送完成，修改状态为当前状态
                count = new_count
                mail_list = new_mail_list
            except:
                print("wrong!")
                exit(1)
        time.sleep(0.5) # 设置0.5秒刷新间隔







