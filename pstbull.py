import pandas as pd
import re
import pypff
import os
import argparse
import requests
from telnetlib import STATUS
from urllib import response



def print_ips(i_list): 
    print("\n\n[+]Printing the IP addresses obtained...\n\n\n")
    print("===============================")
    print("         IP ADDRESSES          ")
    print("===============================\n\n")
    for i in i_list:
        response = requests.get(f'https://ipapi.co/{i}/json/').json()
        location_data = {
            "ip": i,
            "city": response.get("city"),
            "region": response.get("region"),
            "country": response.get("country_name")
        }
        print("IP  :  ",location_data["ip"])
        print("City : ",location_data["city"])
        print("Region : ",location_data["region"])
        print("Country : ",location_data["country"])
        print("\n")            

def check(url):
    #checks the web exists
    try:
        response = get(url)
    #check response status ie wether the site runs or gets terminated if runs properly 200 is returned
        status = response.status_code
        if status == 200:
            return True
        else:
            return False
        #true is returned when site is valid and false when invalid
    except:
        return False

def url_spliter(u):

    sym = [
        "=",
        ";",
        "+",
        "*",
        ")",
        "(",
        "'",
        "&",
        "$",
        "!",
        "@",
        "]",
        "[",
        "#",
        "/",
        ",",
        ".",
        "-",
        ":",
        "_",
        "~",
        "?",
    ]
    s = ""
    for i in u:
        if i.isalpha() or i in sym or i.isdigit():
            s = s + i
            if i.isdigit():
                s = s + i
        else:
            break
    return s

def  get_arguements() :
    parser = argparse.ArgumentParser()
    parser.add_argument("-A","--aggressive", help="Aggressive mode displays all the collected info at a time",action="store_true")
    parser.add_argument("-e", "--emails", help="Display the emails obtained from the pst file",action="store_true")
    parser.add_argument("-n", "--names", help="Display the Names obtained from the pst file",action="store_true")
    parser.add_argument("-u", "--urls", help="Display the Urls obtained from the pst file",action="store_true")
    parser.add_argument("-i", "--ips", help="Display the IP Addresses obtained from the pst file",action="store_true")
    parser.add_argument("-f", "--file", dest="fileName", help="Pst file to be analyzed")
     
    option = parser.parse_args()
    if not option.fileName :
        parser.error("[-] Please specify the file to be analyzed.. or use --help for help")
    
    return (option)

def print_emails(e_list):
  print("\n\n[+]Printing the emails obtained...\n\n\n")
  print("===============================")
  print("            EMAILS             ")
  print("===============================\n\n")
  for email in e_list:
        print(email)
  print("\n\n")

def print_urls(u_list):
  print("\n\n[+]Printing the urls obtained...\n\n\n")
  print("===============================")
  print("             URLS              ")
  print("===============================\n\n")
  for url in u_list:
        print(url,"\t",end="")
        if check(url):
            print("Status-Working")
        else:
            print ("Status-Not Working")    
  print("\n\n") 




  
def print_names(n_list):
  print("\n\n[+]Printing the names obtained...\n\n\n")
  print("===============================")
  print("            NAMES              ")
  print("===============================\n\n")
  for name in n_list:
        print(name)
  print("\n\n")
  
def sender_name_extractor(message):
    name = message.sender_name
    name = " ".join(name.split()[:2])
    name = name if "@" not in name else ""
    name = "".join(i for i in name if i.isalnum())
    name_list.append(name.strip())
    return name
def reciever_names_extractor(message):
    header = message.transport_headers.split("X-ZL")
    if len(header) < 4:
        return
    reciever = ""
    cc = ""
    if "-to:" in header[2]:
        reciever = header[2].split("-to:")
        reciever = (reciever[1].split("<")[0])
    if "-cc:" in header[3]:
        cc = header[3].split("-cc:")
        cc = (cc[1].split("<")[0])
    name = ""
    if "@" in reciever :
      reciever = "" 
    if "@" in cc:
      cc = ""
    name = reciever + cc
    name = "\n".join([i.strip() for i in name.split(",")])
    name =  "".join(i for i in name if i.isalnum())
    name_list.append(name)
    return name


def email_checker(email):
    temp = email.split("@")
    after_at = temp[1]
    if after_at.count(".") > 1:
        return False
    return True


def extractors(message):
    temp_email_lst = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+',
                                message.transport_headers)
    temp_email_lst = [
        i for i in temp_email_lst if "+" not in i and email_checker(i)
    ]
    email_list.extend(temp_email_lst)
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|(([^\s()<>]+|(([^\s()<>]+)))))+(?:(([^\s()<>]+|(([^\s()<>]+))))|[^\s`!()[]{};:'\".,<>?«»“”‘’]))"
    temp_link_list = re.findall(
        regex,
        (str(message.transport_headers) + "\n" + str(message.plain_text_body)))
    temp_link_list = [i[0] for i in temp_link_list]
    temp_link_list = [i.split("\r")[0] for i in temp_link_list]
    temp_link_list = [url_spliter(i) for i in temp_link_list]
    # print(temp_link_list)
    link_list.extend(temp_link_list)

    temp_ip_list = re.findall(
        r"\(\[[0-9]*\.[0-9]+[0-9]*\.[0-9]+[0-9]*\.[0-9]+\]\)",
        message.transport_headers)
    temp_ip_list = [i[2:-2] for i in temp_ip_list]
    ip_list.extend(temp_ip_list)
    return {
        "email": temp_email_lst,
        "links": temp_link_list,
        "ips": temp_ip_list
    }



def processMessage(message):
    reciever_names_extractor(message)
    temp_extract = extractors(message)
    return {
        "subject": message.subject,
        "sender": message.sender_name,
        "sender_name": sender_name_extractor(message),
        "reciever_names": reciever_names_extractor(message),
        "header": message.transport_headers,
        "body": message.plain_text_body,
        "creation_time": message.creation_time,
        "submit_time": message.client_submit_time,
        "delivery_time": message.delivery_time,
        "attachment_count": message.number_of_attachments,
        "emails": "\n".join(temp_extract["email"]),
        "links": "\n".join(temp_extract['links']),
        "ips": "\n".join(temp_extract['ips'])
    }


def folderReport(msg_list, folder_name):
    if not len(msg_list):
        return
    temp_dict = {}
    for col in msg_list[0]:
        temp_dict[col] = []
    for msg in msg_list:
        for id in msg:
            temp_dict[id].append(msg[id])
    try:
        pd.DataFrame(temp_dict).to_csv(f"{fileName[:-4]}/{folder_name}.csv",
                                       index=False)
    except:
        return


def checkForMessage(folder):
    msg_list = []
    for message in folder.sub_messages:
        msg_dict = processMessage(message)
        msg_list.append(msg_dict)
    folderReport(msg_list, folder.name)


def folderTraverse(root_base):
    for folder in root_base.sub_folders:
        if folder.number_of_sub_folders:
            folderTraverse(folder)
        checkForMessage(folder)


def printClearOutput(folderName):
  import os
  import pandas as pd
  root = folderName
  sub_folders = os.listdir(root)
  to_remove = ['Names.csv', 'Ips.csv', 'Links.csv', 'Emails.csv']
  for name in to_remove:
    sub_folders.remove(name)
  for name in sub_folders:
    df=pd.read_csv(folderName+"/"+name)
    columns = [i for i in df]
    print(name.upper()+"\n")
    print("MESSAGES\n")
    for index,msg in df.iterrows():
      msg = [str(i) for i  in msg]
      body = "".join([i for i in msg[5] if i.isalnum() or i.isspace()])
      print("Subject: "+msg[0])
      print("From: "+msg[1])
      print("To: ",msg[3])
      print("\nBody: \n\n",body+"\n")
      print("Number of Attachments: ",msg[9])
      print("Creation Time: ",msg[6])
      print("Submit Time: ",msg[7])
      print("Delivery Time: ",msg[8])
      print("Ip Address: ",msg[12])
      print("\n\t\t\t\t**************************************\n")
def printPst(folderName):
    root = folderName
    sub_folders = os.listdir(root)
    sub_folders.remove("Emails.csv")
    if sub_folders == 0:
        return
    tab_space = "\t"
    print(root)
    for file in sub_folders:
        file_dir = root + "/" + file
        print(tab_space + " - " + file)
        df = pd.read_csv(file_dir)
        for index, row in df.iterrows():
            print(tab_space * 2 + " - " + "Sender : " + str(row['sender']) +
                  " | Subject : " + str(row['subject']))


def PstFileAnalyser():
    pst = pypff.open(fileName)
    root = pst.get_root_folder()
    global email_list, link_list, ip_list, name_list
    email_list = []
    link_list = []
    ip_list = []
    name_list = []
    try:
        os.mkdir(f"{fileName[:-4]}")
    except:
        pass
    folderTraverse(root)
    email_list = list(set(email_list))
    link_list = list(set(link_list))
    ip_list = list(set(ip_list))
    name_list = [i for i  in list(set(name_list)) if i!=""]
    pd.DataFrame({"names":name_list}).to_csv(f"{fileName[:-4]}/Names.csv",index=False)
    pd.DataFrame({
        "ip": ip_list
    }).to_csv(f"{fileName[:-4]}/Ips.csv", index=False)
    pd.DataFrame({
        "links": link_list
    }).to_csv(f"{fileName[:-4]}/Links.csv", index=False)
    pd.DataFrame({
        "email": email_list
    }).to_csv(f"{fileName[:-4]}/Emails.csv", index=False)
    if args.aggressive:
      printClearOutput(fileName[:-4])
      return  
    if args.emails :
      print_emails(email_list)
    if args.urls :
      print_urls(link_list)
    if args.ips :
      print_ips(ip_list)



args=get_arguements()
fileName = args.fileName
PstFileAnalyser()
