# -*- coding: utf-8 -*-
# @Time    : 2021/3/5 10:23
# @Author  : chenkang19736
# @File    : mail.py
# @Software: PyCharm

import os
import smtplib
import email
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.audio import MIMEAudio
from email.mime.image import MIMEImage
from email.encoders import encode_base64

from Jira.public import logging


class Mail:
    def __init__(self, sender, password):
        self.sender = sender
        self.password = password

    def sendMail(self, subject, recipient, text, *attachmentFilePaths):
        '''发送邮件函数：参数（邮件主题，邮件内容，邮件附件（可多选））'''
        msg = MIMEMultipart()  # 发送附件时需调用 MIMEMultipart类，创建 MIMEMultipart,并添加信息头
        '''
    MIME邮件的基本信息、格式信息、编码方式等重要内容都记录在邮件内的各种域中，
    域的基本格式：{域名}：{内容}，域由域名后面跟“：”再加上域的信息内容构成，一条域在邮件中占一行或者多行，
    域的首行左侧不能有空白字符，比如空格或者制表符，占用多行的域其后续行则必须以空白字符开头。
    域的信息内容中还可以包含属性，属性之间以“;”分隔，属性的格式如下：{属性名称}=”{属性值}”。    
        '''
        msg['From'] = self.sender
        msg['To'] = ";".join(recipient)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        for attachmentFilePath in attachmentFilePaths:  #判断添加哪些附件
            msg.attach(self.getAttachment(attachmentFilePath))  #如果入参给定附件文件，使用attach 发放添加msg头信息
        try:
            mailServer = smtplib.SMTP('smtp.qq.com', 587)  # 连接腾讯邮件的服务器；SMTP（Simple Mail Transfer Protocol）即简单邮件传输协议，用于由源地址到目的地址传送邮件的规则，由它来控制信件的中转方式
            mailServer.ehlo()  # 使用starttls 方法必须先 ehlo 判断是否是同一个地址。。。
            mailServer.starttls()  # 以下SMTP都会加密;Put the SMTP connection in TLS (Transport Layer Security) mode. All SMTP commands that follow will be encrypted.
            mailServer.ehlo()  # You should then call ehlo() again.
            mailServer.login(self.sender, self.password)  # 登录邮箱
            mailServer.sendmail(self.sender, recipient, msg.as_string())  # 发送邮件（发件人，收件人，发送内容）
            mailServer.close() # 关闭邮件发送服务
            logging.info('Sent email to %s successfully' % recipient)
        except Exception as e:
            logging.error('sendEmai failed %s' % e)

    def getAttachment(self, attachmentFilePath):
        """
        获取附件， 根据附件的不同类型进行封装
        :param attachmentFilePath:  文件路径
        :return: 返回封装后的附件对象
        """
        contentType, encoding = mimetypes.guess_type(attachmentFilePath)  # 根据 guess_type方法判断文件的类型和编码方式

        if contentType is None or encoding is not None:  # 如果根据文件的名字/后缀识别不出是什么文件类型
            contentType = 'application/octet-stream'   # 使用默认类型，usable for a MIME content-type header.

        mainType, subType = contentType.split('/', 1)  # 根据contentType 判断主类型与子类型
        file = open(attachmentFilePath, 'rb')

        if mainType == 'text':  # 根据主类型不同，调用不同的文件读取方法
            attachment = MIMEBase(mainType, subType)  # A subclass of MIMENonMultipart, the MIMEText class is used to create MIME objects of major type text.
            attachment.set_payload(file.read())  # Set the entire message object’s payload（负载） to payload.
            encode_base64(attachment)  # Encodes the payload into base64 form and sets the Content-Transfer-Encoding header to base64.
        elif mainType == 'message':
            attachment = email.message_from_file(file)  # 使用message_from_file方法，Return a message object structure tree from an open file object
        elif mainType == 'image':  # 图片
            attachment = MIMEImage(file.read())  #A subclass of MIMENonMultipart, the MIMEImage class is used to create MIME message objects of major type image.
        #elif mainType == 'audio':  # 音频
            #attachment = MIMEAudio(file.read(), _subType=subType)  #A subclass of MIMENonMultipart, the MIMEAudio class is used to create MIME message objects of major type audio.

        else:
            attachment = MIMEBase(mainType, subType)  # The MIMEBase class always adds a Content-Type header (based on _maintype, _subtype, and _params), and a MIME-Version header (always set to 1.0).
            attachment.set_payload(file.read())  # Set the entire message object’s payload（负载） to payload.
            encode_base64(attachment)  # Encodes the payload into base64 form and sets the Content-Transfer-Encoding header to base64.

        file.close()
        """
    Content-disposition 是 MIME 协议的扩展，MIME 协议指示 MIME 用户代理如何显示附加的文件。Content-disposition其实可以控制用户请求所得的内容存为一个文件的时候提供一个默认的文件名，
    文件直接在浏览器上显示或者在访问时弹出文件下载对话框。 content-disposition = "Content-Disposition" ":" disposition-type *( ";" disposition-parm ) 。    
        """
        attachment.add_header('Content-Disposition', 'attachment',   filename=os.path.basename(attachmentFilePath))  #Content-Disposition为属性名 disposition-type是以什么方式下载，如attachment为以附件方式下载 disposition-parm为默认保存时的文件名
        return attachment


if __name__ == "__main__":
    from Jira import ROOT_PATH
    from Jira.public.loading_config import config_obj_of_authority

    sender = config_obj_of_authority['email']['username'].strip()
    password = config_obj_of_authority['email']['password'].strip()
    recipients = recipient = config_obj_of_authority['email']['recipient'].strip()

    mail = Mail(sender=sender, password=password)
    if "," in recipient:
        recipients = recipient.split(",")

    file = os.path.join(os.path.dirname(ROOT_PATH, "static/bug_list.xls"))
    mail.sendMail('邮件主题', recipients, "Its a test.", file)
    logging.info("Have Send Email Successfully.")
