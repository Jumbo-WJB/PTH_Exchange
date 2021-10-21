#coding:utf-8
#author:Jumbo

import argparse
from string import Template
import xml.etree.cElementTree as ET
import base64
import requests
from requests_ntlm import HttpNtlmAuth
import logging
import os
import warnings
warnings.filterwarnings('ignore', message='Unverified HTTPS request')

def DownloadEmails(target,session,folder):
    logger.debug("[Stage 777] Get Mails Stage 1 Finditem ing... ")
    FindItem_body = convertFromTemplate({'folder':folder},templatesFolder + "FindItem.xml")
    FindItem_request = session.post(
        f"{target}/ews/exchange.asmx", headers=headers,
        data=FindItem_body,
        verify=False
        )
    # If status code 200 is NOT returned, the request failed
    if FindItem_request.status_code != 200:
        logger.error("[Stage 777] Request failed - Get Mails Stage 1 Finditem Error!")
        exit()
    folderXML = ET.fromstring(FindItem_request.content.decode())
    i = 0
    for item in folderXML.findall(".//t:ItemId", exchangeNamespace):
        params = {'Id': item.get('Id'), 'ChangeKey': item.get('ChangeKey')}
        logger.debug("[Stage 777] Get Mails Stage 2 GetItem ing... ")
        GetItem_body = convertFromTemplate(params, templatesFolder + "GetItem.xml")
        GetItem_request = session.post(
            f"{target}/ews/exchange.asmx", headers=headers,
            data=GetItem_body,
            verify=False
            )   
        if GetItem_request.status_code != 200:
            logger.error("[Stage 777] Request failed - Get Mails Stage 2 GetItem Error!")
            exit()
        itemXML = ET.fromstring(GetItem_request.content.decode())
        mimeContent = itemXML.find(".//t:MimeContent", exchangeNamespace).text
        logger.debug("[Stage 777] Get Mails Stage 3 Downloaditem ing... ")
        try:
            extension = "eml"
            outputDir = "output"
            if not os.path.exists(outputDir):
                os.makedirs(outputDir)
            fileName = outputDir + "/item-{}.".format(i) + extension
            with open(fileName, 'wb+') as fileHandle:
                fileHandle.write(base64.b64decode(mimeContent))
                fileHandle.close()
                print("[+] Item [{}] saved successfully".format(fileName))
        except IOError:
            print("[!] Could not write file [{}]".format(fileName))
        DownAttachment(target,item.get('Id'),i)
        i = i + 1


def DownAttachment(target,id,i):
    logger.debug("[Stage 555] Ready Download Attachmenting... ")
    params2 = {'Id': id}
    GetItem_body = convertFromTemplate(params2, templatesFolder + "GetAttachmentID.xml")
    GetItem_request = session.post(
        f"{target}/ews/exchange.asmx", headers=headers,
        data=GetItem_body,
        verify=False
        )
    logger.debug("[Stage 555] Determine if there are attachments in the email... ")
    if "AttachmentId" in GetItem_request.content.decode():
        itemXML = ET.fromstring(GetItem_request.content.decode())
        logger.debug("[Stage 555] This Mail Has Attachment... ")
        AttachmentIds = itemXML.findall(".//t:AttachmentId", exchangeNamespace)
        for AttachmentId in AttachmentIds:
            # print(AttachmentId.get('Id'))
            logger.debug("[Stage 555] Start Get Attachment Content... ")
            Attachment_body = convertFromTemplate({'AttachmentId':AttachmentId.get('Id')},templatesFolder + "GetAttachmentbody.xml")
            Attachment_request = session.post(
                f"{target}/ews/exchange.asmx", headers=headers,
                data=Attachment_body,
                verify=False
                )
            AttachmentXML = ET.fromstring(Attachment_request.content.decode())
            AttachmentXMLname = AttachmentXML.find(".//t:Name", exchangeNamespace).text
            if "<t:Content>" in Attachment_request.content.decode():
                AttachmentXMLcontent = AttachmentXML.find(".//t:Content", exchangeNamespace).text
                logger.debug("[Stage 555] Start Download Attachment... ")
                try:
                    outputDir = "output"
                    if not os.path.exists(outputDir):
                        os.makedirs(outputDir)
                    fileName = outputDir + "/item-{}-".format(i) + AttachmentXMLname
                    with open(fileName, 'wb+') as fileHandle:
                        fileHandle.write(base64.b64decode(AttachmentXMLcontent))
                        fileHandle.close()
                        print("[+] Item [{}] saved successfully".format(fileName))
                except IOError:
                    print("[!] Could not write file [{}]".format(fileName))
            elif "</t:Body>" in Attachment_request.content.decode():
                AttachmentXMLcontent = AttachmentXML.find(".//t:Body", exchangeNamespace).text
                logger.debug("[Stage 555] Start Download Attachment With Body... ")
                try:
                    outputDir = "output"
                    if not os.path.exists(outputDir):
                        os.makedirs(outputDir)
                    fileName = outputDir + "/item-{}-".format(i) + AttachmentXMLname
                    with open(fileName, 'wb+') as fileHandle:
                        fileHandle.write(AttachmentXMLcontent.encode())
                        fileHandle.close()
                        print("[+] Item [{}] saved successfully".format(fileName))
                except IOError:
                    print("[!] Could not write file [{}]".format(fileName))
            else:
                logger.warning("Attachment ERROR ,Download Failed!")





def convertFromTemplate(shellcode, templateFile):
    try:
        with open(templateFile) as f:
            src = Template(f.read())
            result = src.substitute(shellcode)
            f.close()
            return result
    except IOError as e:
        print("[!] Could not open or read template file [{}]".format(templateFile))
        return e
 
def encode_to_base64(text):
    message_bytes = text.encode('ascii')
    base64_bytes = base64.b64encode(message_bytes)
    return base64_bytes.decode('ascii')

 
if __name__ == '__main__':
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    templatesFolder = "ews_template/"
    exchangeNamespace = {'m': 'http://schemas.microsoft.com/exchange/services/2006/messages', 't': 'http://schemas.microsoft.com/exchange/services/2006/types'}
    parser = argparse.ArgumentParser()
    parser.add_argument('--target',required=True, help='Exchange Target')
    parser.add_argument('--username',required=True, help='email username')
    parser.add_argument('--password',required=True, help='email password or ntlm hash')
    parser.add_argument('--action',required=True,choices=['Brute','Search','Download'], help='The action you want to take')
    parser.add_argument("--keyword", help="keyword with you want search")
    parser.add_argument("--folder",default="inbox", help="folder name with you want download")
    args = parser.parse_args()
    session = requests.Session()
    USER_AGENT = 'Microsoft Office/16.0 (Windows NT 10.0; MAPI 16.0.9001; Pro)'
    headers = {'User-Agent': USER_AGENT,
        'Connection': 'Close',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Content-Type': 'text/xml'}
if args.target and args.username and args.password and args.action == "Download" and args.folder:
    session.auth = HttpNtlmAuth(username=args.username, password=args.password)
    DownloadEmails(args.target,session,args.folder)