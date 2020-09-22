#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ----------------------------------------------
# common module import
# ----------------------------------------------
import os
import sys
import codecs

# ----------------------------------------------
# outlookからメール情報を取得する関数
# ----------------------------------------------
import win32com.client

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39fe001e"
PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

# ----------------------------------------------
# パラメータ
# ----------------------------------------------
# Outlook内のメールアドレスを正しく取得したい時に使用
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39fe001e"

# Outlook内のメールヘッダーを取得したい場合にのみ使用
PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

# こちらで指定したメールのカウントを取得
MAX_OUTLOOK_ITEM_COUNT = 2000      # 2000

#Outlookメールのキャッシュを保存するディレクトリ
OUTLOOK_CACHE_DIR = 'outlook_cache'

# ---------------------------------
# outlookからメール内の文章を読み込み
# ファイルを作成する
# ---------------------------------

def make_sentence_from_outlook():
    outlook_mapi = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  # Outlook API読み込み
    outlook_folder = outlook_mapi.GetDefaultFolder(6)  # inboxのフォルダを取得

    outlook_readitem_count = 0

    for outlook_item in outlook_folder.Items:  # アイテムをループ
        if MAX_OUTLOOK_ITEM_COUNT < outlook_readitem_count:
            return  # 指定カウント以上で取得を終了
        if outlook_item.Class != 43:
            continue # 通常のメールではない会議通知などを除外する

        sys.stdout.write('\rOutlookからメール取得中 {} / {}'.format(outlook_readitem_count, MAX_OUTLOOK_ITEM_COUNT))

        mail_entryid = outlook_item.entryid
        #mail_conversationid = outlook_item.conversationid
        #print('entryid: ',outlook_item.entryid)                # outlook mailitemのユニークなID
        #print('ConversationID',outlook_item.conversationid)    # メールの識別子 返信などグループ化されている内容を捜索したい時に使用

        mail_senton = str(outlook_item.senton)
        mail_receivedtime = str(outlook_item.receivedtime)
        mail_subject = outlook_item.subject
        mail_body = outlook_mail_load_cache(outlook_item.entryid)

        #print('SentOn: ',outlook_item.senton)                  # 送信日時
        #print('ReceivedTime: ',str(outlook_item.receivedtime)) # 受信日時
        #print('Subject: ',outlook_item.subject)                # 件名
    
        #outlook_item_mailheader = outlook_item.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)    # mail headerを取得

        #mail_header = outlook_item_mailheader
        #mail_size = outlook_item.size
        mail_to = outlook_item.to
        mail_cc = outlook_item.cc
        #mail_body = outlook_item.body
        mail_sender_address = outlook_item.sender.address

        #print('Header: ', outlook_item_mailheader)
        #print('Size: ',outlook_item.size)                      # mailitemサイズ
        #print('To: ',outlook_item.to)                          # 宛先
        #print('CC: ',outlook_item.cc)                          # CC
        #print('Body: ',outlook_item.body)                      # 本文
        #print('SenderAddress: ',outlook_item.sender.address)   # 送信者メールアドレス

        if not mail_subject: mail_subject='(件名なし)'
        if not mail_senton: mail_senton='(送信日時なし)'
        if not mail_receivedtime: mail_receivedtime='(受信日時なし)'
        if not mail_to: mail_to='(宛先なし)'
        if not mail_cc: mail_cc='(CC)なし'
        if not mail_sender_address: mail_sender_address='(送信者なし)'

        if mail_body == None:
            mail_body = '送信日時: ' + mail_senton + "\n" + '受信日時: ' + mail_receivedtime + "\n" +  'Subject: ' + mail_subject + "\n" + 'SenderAddress: ' + mail_sender_address + "\n" + 'To: ' + mail_to + "\n" + 'CC: ' + mail_cc + "\n" + 'Body: ' + outlook_item.body
            outlook_mail_save_cache(outlook_item.entryid,mail_body)


        # Office365環境などでメールアドレスが正しく取得できない場合
        for outlook_item_rec in outlook_item.Recipients:
            if outlook_item_rec.address == 'Unknown':continue
            mail_str_address = ""
            mail_address_type = outlook_item_rec.Type # どこのアドレスかを取得する 1:to , 2:cc , 3:bcc
            mail_str_address = outlook_item_rec.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            print(outlook_item_rec.name,mail_str_address)

        outlook_readitem_count += 1


# ---------------------------------
# outlookのメールをキャッシュ
# ---------------------------------
#読み込み  空ファイル作成してOPEN or 既にあるファイルをOPENして、読み込み結果を返す
def outlook_mail_load_cache(mailid):
    if os.path.exists(OUTLOOK_CACHE_DIR+'/'+mailid):
        with codecs.open(OUTLOOK_CACHE_DIR+'/'+mailid, 'r' , 'utf-8') as f:
            return f.read()
    return None
#書き込み
def outlook_mail_save_cache(mailid,txt):
    with codecs.open(OUTLOOK_CACHE_DIR+'/'+mailid, 'w' , 'utf-8') as f:
        f.write(txt)


# ---------------------------------
# Main ファイルへメール抜き出し
# ---------------------------------

make_sentence_from_outlook()


