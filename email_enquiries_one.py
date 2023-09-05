import os,re
import pyodbc
import numpy as np
import pandas as pd
import duckdb
from tabulate import tabulate
import win32com.client
from pathlib import Path
from collections import OrderedDict
from fuzzywuzzy import process, fuzz
import rapidfuzz
#other libraries to be used in this script
import datetime
# from datetime import datetime, timedelta
import spacy
from spacy import displacy
from collections import Counter
import warnings
warnings.filterwarnings( 'ignore' )

#Get the Current Date and Time.
dt = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S:000")

#***************************************Creator's info table*********************************************************************************/
Creator_data = pd.DataFrame({"fldCreatedBy":["PRAVIN SUBRAMANIAN"] ,"fldCreatedOn":[dt]})
# print(Creator_data)
#********************************************************************************************************************************************/

#Connecting Database**********************************************************************************
conn_str = pyodbc.connect(
    r'Driver=SQL Server;'
    r'Server=XXXXXX;'
    r'Database=XXXX_Email;'
    r'Trusted_Connection=yes;'
    )
cursor = conn_str.cursor()
#*****************************************************************************************************


#****************************************************************************************************
# #Code to delete the records
# cursor.execute("""DELETE FROM [XXXX_Email].[dbo].[tblCustomerStage]""")

# #Code to reset the Identity Value (Auto-Increment)
# cursor.execute("""DBCC CHECKIDENT ('[XXXX_Email].[dbo].[tblCustomerStage]', RESEED, 0)""")

# conn_str.commit()
#********************************************************************************************************

# cust_ass = pd.read_sql("SELECT  * FROM [XXXX_Email].[dbo].[tblCustAsset] where fldIHSMarkit=1", conn_str)
# print(cust_ass)


#***Load the Model!!****
nlp = spacy.load(r'C:\Users\pravin.subramanian\Downloads\Python_Scripts\SpacyModels\Spacy_CUST_Model_with_emails_V2.08')



otlk = win32com.client.Dispatch("Outlook.Application")
outlook = otlk.GetNamespace("MAPI")
#The email of the sender after Quote preparation.
email_sender = outlook.Session.Accounts['pravin.subramanian@university.com']

for i in (range(len(outlook.Folders))):
    #Condition to take emails from that particular account.
    if (outlook.Folders.Item(i+1).Name == 'autoenquirytest@university.com'):
        root_folder = outlook.Folders.Item(i+1)
        #Looping to find the inbox folder and extracting the emails
        for folder in root_folder.Folders:
            if (folder.Name == 'Inbox'):
                messages = folder.Items
                messages.Sort("[ReceivedTime]", True) 
                msg_no = 0
                for message in messages:
                    if message.UnRead == True:
                        # message.UnRead = False
                        # message.Save()
                        
                        mail = otlk.CreateItem(0)
                        mail.ReadReceiptRequested
                        mail.To = 'pravin.subramanian@university.com; pravin.subramanian@university.com' #Recipients list
                        
                        #Code to take the email address of the Quote requestor
                        email_sender = ''
                        if message.Class==43:
                            if message.SenderEmailType=='EX':
                                email_sender = (message.Sender.GetExchangeUser().PrimarySmtpAddress)
                                email_sender = str(message.Sender) + ' <' + email_sender + '>'
                            else:
                                email_sender = (message.SenderEmailAddress)
                                email_sender = str(message.Sender) + ' <' + email_sender + '>'
                        # print(email_sender)      

                        # Get the recipient email addresses
                        recipients = message.Recipients
                        recipient_email_addresses = []
                        for recipient in recipients:
                            if recipient.AddressEntry is not None and recipient.AddressEntry.GetExchangeDistributionList() is not None:
                                # Get the email address of the group
                                group_email_address = recipient.AddressEntry.GetExchangeDistributionList().PrimarySmtpAddress
                                recipient_email_addresses.append(group_email_address)
                            # Check if the recipient is an individual email address
                            elif recipient.AddressEntry is not None and recipient.AddressEntry.GetExchangeUser() is not None:
                                # Get the email address of the individual
                                individual_email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                                recipient_email_addresses.append(individual_email_address)
                            elif recipient.AddressEntry is not None and recipient.AddressEntry.Address is not None:
                                individual_email_address = recipient.AddressEntry.Address
                                recipient_email_addresses.append(individual_email_address)


                        # Concatenate all the recipient email addresses into a single string#
                        recipient_email_string = ";".join(recipient_email_addresses)



                        #Ignore 24IT and Microsoft emails and consider other emails.****************************#
                        if  ('@twenty-four.it' not in email_sender) and ('noreply@microsoft' not in email_sender):
                            msg_no+=1
                            body = message.Body

                            try:
                                message.UnRead = False
                                message.Save()
                                # print(email_sender, recipient_email_addresses)

                                body_no_sign_part = re.sub(r'(?:With best regards,\s*.*From:)', 'From:', body, flags = re.IGNORECASE|re.DOTALL|re.MULTILINE) # Code to remove the signature from body.#
                                body_no_title = re.sub(r'(?:Subject:|From:|Sent:|To:|Cc:)\s*.*\n?', '', body_no_sign_part)# Code to remove from,to,subjects from body.#
                                body_no_Rambo = re.sub(r'(?:Thanks and Regards\s*.*(?:Rambo|company Ltd\s*.*\n.*___))', '', body_no_title, flags = re.IGNORECASE|re.DOTALL|re.MULTILINE)
                                body_no_caution = re.sub(r'(?:CAUTION - EXTERNAL SENDER !)', '', body_no_Rambo) #Code to remove caution message.#
                                body_with_no_links = re.sub(r'[<]?http[s]?://\S+', '', body_no_caution, flags=re.MULTILINE) #Code to remove any links from the body.#
                                body_with_no_sign = re.sub(r'^(?:Thank you\n|Thanks and Best Regards|Best Regards|With best regards\
                                |With kind regards|With best regards|Disclaimer:|Deniz Yenicerioglu)\s*.*', '', body_with_no_links, flags=re.MULTILINE|re.DOTALL|re.IGNORECASE)
                                body_no_ship_sign = re.sub(r'(?:Please do not delete this email\s*.*companyanduniv.com>  .)', '', body_with_no_sign, flags = re.IGNORECASE|re.DOTALL|re.MULTILINE)
                                comp_body = (body_no_ship_sign.rstrip()).lstrip()
                                db_email_body = "\n".join(list(OrderedDict.fromkeys(comp_body.split("\n"))))

                                #Detecting the entities from the email body using Spacy model.
                                doc=nlp(db_email_body)
                                ent_list = [(e.text, e.label_)for e in doc.ents]                                
                                
                                
#******************************************************************Part 1*************************************************************************************************************************
                                #If the Customer name is not detected in the body of the email, we are checking with the email signature
                                if 'CUST' not in (dict(ent_list)).values():
                                    # print('Getting in')
                                    body_v1_no_Rambo = re.sub(r'(?:Thanks and Regards\s*.*(?:Rambo|company Ltd\s*.*\n.*___))', '', body, flags = re.IGNORECASE|re.DOTALL|re.MULTILINE)
                                    body_v1_no_from = re.sub(r'(?:\s*.*\n.*EXTERNAL SENDER !)', '', body_v1_no_Rambo, flags = re.IGNORECASE|re.DOTALL|re.MULTILINE)
                                    body_newV = nlp(body_v1_no_from)
                                    Cust_list = [(e.text, e.label_)for e in body_newV.ents]
                                    entities = []
                                    text_ent  = {}
                                    for ent in body_newV.ents:
                                        ENT_Label = ent.label_
                                        ENT_Text = ent.text
                                        if ENT_Text not in entities:
                                            entities.append(ENT_Text)
                                            text_ent[ENT_Label]=ENT_Text
                                                                        
                                
                                
                                else:
                                    entities = []
                                    text_ent  = {}
                                    
                                    # print('Not getting in')
                                    for ent in doc.ents:
                                        ENT_Label = ent.label_
                                        ENT_Text = ent.text
                                        if ENT_Text not in entities:
                                            entities.append(ENT_Text)
                                            text_ent[ENT_Label]=ENT_Text
                                    # print(text_ent.items(),'2')
                                print(text_ent.items(), 'initial dicts')
#*********************************************************End of Part 1****************************************************************************************#

#********************************************************Part 2************************************************************************************************#
                                #Getting the sender of the email from the 'FROM' address
                                #New Code to get the sender email address when the email enters company environment.
                                t0 = re.findall(r'(?:From:|Von:)\s.*\nSent:.*\nTo:.*',body)
                                if (len(t0))>0:
                                    t0_wthSndRec = 'From: '+ email_sender + '\n' +'To: '+ recipient_email_string
                                    t0.insert(0, t0_wthSndRec)
                               
                                #*******Duplicate creation****************#
                                # dup_from = [re.findall(r'(?:From:|Von:)\s\S+@\S+',body)] #[0][0]
                                dup_from = ';'.join(','.join(email_address) for email_address in [re.findall(r'(?:From:|Von:)\s\S+@\S+',body)])
                                
                                dup_to = email_sender
                                dup = dup_from + '\n' + 'To: '+dup_to
                                # dup = 'From: '+email_sender + '\n' + 'To: '+recipient_email_addresses
                                dup_list = []
                                dup_list.append(dup) if (len(dup_from))>0 else dup_list
                                
                                #****************************************#
                                #Condition to take the email address if we receive the eamil driectly from client.
                                dir_email = 'From: '+ email_sender + '\n' +'To: '+ recipient_email_string
                                dir_emaillist = []
                                dir_emaillist.append(dir_email)
                                
                                t0 = t0 if (len(t0))>0 else dup_list if (len(dup_list))>0 else dir_emaillist
                                # print(t0,'t0')
                                dk = 0
                                check = ''
                                db_email_add = ''
                                db_email_nameadd = ''
                                for i in (range(len(t0))):
                                    dk -=1
                                    senderadd = re.findall(r'(?:From:|Von:)\s.*',t0[dk])
                                    receiveradd = re.findall(r'(?:To:)\s.*',t0[dk])
                                    
                                    #Take 'From' address only if the receipients list contains 'company'.
                                    if 'company' in receiveradd[0] and check == '':
                                        check = re.findall(r'\S*@\S*', (re.sub(r'(?:<|>|]|\[|mailto:|;|,)', '', senderadd[0])))
                                        if len(check)==0:
                                            check = ''
                                        else:
                                            # print(re.findall(r'\S*@\S*', (re.sub(r'(?:<|>|]|\[|mailto:|;|,)', '', senderadd[0]))), 'sender' )
                                            db_email_add = check[0] #Got the final email address of the Sender

                                            #Code to get the sender name
                                            snd_lst = [s.replace('From:','') for s in senderadd]
                                            db_email_name = [re.sub(r'(?:<|\[)', '', k) for j in [re.findall(r'.*(?:<|\[)', i) for i in snd_lst] for k in j]
                                            db_email_nameadd = (db_email_name[0].rstrip()).lstrip() #Got the final name of the Sender
                                        # print(db_email_nameadd)
                                
                                #If none of the recipients contains 'company' keywords. Then directly take the final 'From' Address
                                if db_email_add == '':
                                    db_email_add = email_sender
                                    check  = [email_sender]
                                print(check)    
                                print(db_email_add)
                                # print(text_ent)
#*************************************************Enf of Part 2***************************************************************************************************#

#************************************************Part 3***********************************************************************************************************#                                
                                #Passin only if the Customer name is detected from the email body.
                                if 'CUST' not in text_ent:
                                    sender = pd.read_sql("SELECT fldCustomerID, fldEmail FROM [XXXX_Email].[dbo].[tblCustContact] where fldEmail = ? and fldCustomerID IS NOT NULL and not fldCustomerID like 'null'", conn_str, params=[db_email_add])
                                    if (len(sender))>0:
                                        custID = sender['fldCustomerID'][0]
                                        cust_info = pd.read_sql("SELECT fldCustomerID, fldCustomerName FROM [XXXX_Email].[dbo].[tblCustomer] where fldCustomerID = ? ", conn_str, params=[str(custID)])
                                        if (len(cust_info))>0 :
                                            cust_name = cust_info['fldCustomerName'][0]
                                            text_ent['CUST'] = cust_name
                                        #  print(text_ent.items())
                                        else: #!!!!!!!!The below else code and this code are same.. So while making changes, update the changes in the below else code also!!!!!
                                            company_names_email = re.findall(re.escape('@')+"(.*?)"+re.escape('.'),db_email_add)
                                            cust_data0 = pd.read_sql("SELECT fldCustomerName, fldCustomerName as lw_custname,fldCustomerID FROM [XXXX_Email].[dbo].[tblCustomer] where not fldCustomerName like ''", conn_str)
                                            # print(company_names_email[0], db_email_add)
                                            cust_data0['Cust_Finder'] = company_names_email[0].lower()
                                            cust_data0['lw_custname'] = cust_data0['lw_custname'].str.lower()
                                            cust_data0['cust_matching_ratio'] = cust_data0.apply(lambda x:rapidfuzz.fuzz.partial_ratio(x.Cust_Finder, x.lw_custname), axis=1).to_list()
                                            cust_rev0 = cust_data0[cust_data0['cust_matching_ratio']>=60].reset_index()
                                            cust_rev0 = cust_rev0.sort_values(by="cust_matching_ratio",ascending=False).reset_index()
                                            text_ent['CUST'] = cust_rev0['fldCustomerName'][0]
                                            # cust_rev0 = cust_rev0[['fldCustomerName','Cust_Finder','fldCustomerID','cust_matching_ratio']]
                                            # print(cust_rev0)

                                        
                                    #Using the company domain name from email address, trying to fetch the customer name from the table using fuzzy search.
                                    else:  #!!!!!!!!The above else code and this code are same.. So while making changes, update the changes in the above code also!!!!!
                                        company_names_email = re.findall(re.escape('@')+"(.*?)"+re.escape('.'),db_email_add)
                                        # print(db_email_add,'db_email_add')
                                        if len(company_names_email)>0:
                                            cust_data0 = pd.read_sql("SELECT fldCustomerName, fldCustomerName as lw_custname,fldCustomerID FROM [XXXX_Email].[dbo].[tblCustomer] where not fldCustomerName like ''", conn_str)
                                            # print(company_names_email[0])
                                            cust_data0['Cust_Finder'] = company_names_email[0].lower()
                                            cust_data0['lw_custname'] = cust_data0['lw_custname'].str.lower()
                                            cust_data0['cust_matching_ratio'] = cust_data0.apply(lambda x:rapidfuzz.fuzz.partial_ratio(x.Cust_Finder, x.lw_custname), axis=1).to_list()
                                            cust_rev0 = cust_data0[cust_data0['cust_matching_ratio']>=60].reset_index()
                                            cust_rev0 = cust_rev0.sort_values(by="cust_matching_ratio",ascending=False).reset_index()
                                            text_ent['CUST'] = cust_rev0['fldCustomerName'][0]
                                        # cust_rev0 = cust_rev0[['fldCustomerName','Cust_Finder','fldCustomerID','cust_matching_ratio']]
                                        # print(cust_rev0)

                                    

                                    

                                # print(text_ent.items(), 'After customer addition')
                                #/******Conditiuon to remove vessel name from main dictionary and store in a new variable*****/
                                #/The main reason for doing this is to check the vessel name first in the email subject and append it with dict./
                                # print(text_ent.items())
                                no_vess_dict = {key: val for key, val in text_ent.items() if key != 'VESS'}
                                vess_dict = {key: val for key, val in text_ent.items() if key == 'VESS'}
                                



                                # if 'VESS' not in (dict(ent_list)).values():
                                vess_sub = nlp(message.Subject)
                                vess_name = [(e.text, e.label_)for e in vess_sub.ents]
                                # print(vess_name)
                                if ('VESS' in (dict(vess_name)).values()) or ('IMO NO.' in (dict(vess_name)).values()):
                                    vess_entities = []
                                    vess_text_ent  = {}
                                    for ent in vess_sub.ents:
                                        ENT_Label = ent.label_
                                        ENT_Text = ent.text
                                        if ENT_Text not in entities:
                                            vess_entities.append(ENT_Text)
                                            vess_text_ent[ENT_Label]=ENT_Text
                                    no_vess_dict.update(vess_text_ent)
                                    # print(vess_text_ent.items(), 'now')
                                if 'VESS' in (no_vess_dict):
                                    text_ent = no_vess_dict
                                    text_ent = {key: val for key, val in text_ent.items() if 'company' not in val.lower()}
                                else:
                                    no_vess_dict.update(vess_dict)
                                    text_ent = no_vess_dict
                                    text_ent = {key: val for key, val in text_ent.items() if 'company' not in val.lower()}

                                # print(text_ent.items(),'final dict')

                                imo_data = ''
                                df_final_rev = ''
                                cust_rev = ''
                                G_Var = 'YES'
                                # print(text_ent)
                                print(msg_no)
                                
                                
                                # text_ent['VESS'] = 'Vessel Check'
                                # text_ent['IMO NO.'] = '3258758'
                                # text_ent['CUST'] = 'Cust Check 014'
                                # print(text_ent.items())
                                
                                #/****Condition to take only the emails that has the customer information in it.****/
                                if 'CUST' in text_ent:
                                    # print('its entering')
                                    #Getting the Imo no, cust name, vessel name while looping the entities
                                    for enty,texts in text_ent.items():
                                        # print(enty, texts)
                                        
                                        if enty == 'IMO NO.':  #Condition to take the Vessel info using the IMO Number detected from the email. 
                                            imo_data = pd.read_sql("SELECT fldVesselID, fldVName, fldVIMO FROM [XXXX_Email].[dbo].[tblVessel] where fldIHSMarkit=1 and fldVIMO = ?", conn_str, params=[texts])
                                            imo_data['matching_ratio'] = 100.0
                                            
                                            
                                        if enty =='VESS': #Condition to take the Vessel info using the Vessel name detected from the email. 
                                            data = pd.read_sql("SELECT fldVesselID, fldVName, fldVName as lw_fldVName, fldVIMO FROM [XXXX_Email].[dbo].[tblVessel] where fldIHSMarkit=1", conn_str)
                                            data['Vess_Finder'] =  ((texts.lower()).replace('///','')).replace('urgent','')
                                            data['lw_fldVName'] = data['lw_fldVName'].str.lower()
                                            data['matching_ratio'] = data.apply(lambda x:rapidfuzz.fuzz.QRatio(x.Vess_Finder, x.lw_fldVName), axis=1).to_list()
                                            df_final_rev = data[data['matching_ratio']>=75].reset_index()
                                            df_final_rev = df_final_rev.sort_values(by="matching_ratio",ascending=False).reset_index()
                                            
                                
                                        if enty =='CUST': #Condition to take the Customer info using the Customer name detected from the email. 
                                            cust_data = pd.read_sql("SELECT fldCustomerName, fldCustomerName as lw_custname,fldCustomerID FROM [XXXX_Email].[dbo].[tblCustomer] where not fldCustomerName like ''", conn_str)
                                            cust_data['Cust_Finder'] = texts.lower()
                                            cust_data['lw_custname'] = cust_data['lw_custname'].str.lower()
                                            # print(texts)
                                            cust_data['cust_matching_ratio'] = cust_data.apply(lambda x:rapidfuzz.fuzz.QRatio(x.Cust_Finder, x.lw_custname), axis=1).to_list()
                                            cust_rev = cust_data[cust_data['cust_matching_ratio']>=80].reset_index()
                                            cust_rev = cust_rev.sort_values(by="cust_matching_ratio",ascending=False).reset_index()                                         
                                            cust_rev = cust_rev[:1]
                                            # print(tabulate(cust_rev, headers='keys', tablefmt='psql'))
                                            # print(text_ent.get('VESS'))
                            
                                            
                                            #/****If the detected customer name is not available in 'tblCustomer' table, append the customer info as a new record in the main and stage table.****/
                                            if (len(cust_rev))==0: 
                                                # print('passed')
                                                cusID_serial = pd.read_sql("SELECT MAX(fldCustomerID) as max_cusid FROM [XXXX_Email].[dbo].[tblCustomer]", conn_str)
                                                # print(cusID_serial['max_cusid'][0],(cusID_serial['max_cusid'][0])+1)

                                                cursor.execute(""" 
                                                INSERT INTO [XXXX_Email].[dbo].[tblCustomer] (fldCustomerID, fldCustomerName, fldCoCreatedDate, fldCoCreatedBy) 
                                                SELECT ?, ?, ?, ?
                                                WHERE NOT EXISTS (
                                                    SELECT fldCustomerID, fldCustomerName, fldCoCreatedDate, fldCoCreatedBy 
                                                    FROM [XXXX_Email].[dbo].[tblCustomer]
                                                    WHERE fldCustomerName LIKE (?) 
                                                    AND fldCoCreatedDate LIKE (?) 
                                                    AND fldCoCreatedBy LIKE (?)
                                                )
                                                """, (int(cusID_serial['max_cusid'][0])+1), texts, str(dt), 'PRAVIN SUBRAMANIAN', texts, str(dt), 'PRAVIN SUBRAMANIAN')


                                                conn_str.commit()

                                                stage_id = pd.read_sql("SELECT fldCustomerName,fldCustomerID FROM [XXXX_Email].[dbo].[tblCustomer] where fldCustomerName = ?", conn_str, params=[texts])
                                                stage_id0 = stage_id['fldCustomerID'][0]

                                                cursor.execute(""" 
                                                insert into [XXXX_Email].[dbo].[tblCustomerStage] (fldCustomerID,fldCustomerName,fldCreatedOn,fldCreatedBy) 
                                                Select ?,?,?,?
                                                Where not exists(select fldCustomerID,fldCustomerName,fldCreatedOn,fldCreatedBy from [XXXX_Email].[dbo].[tblCustomerStage]
                                                where fldCustomerID like (?) and fldCustomerName like (?) and fldCreatedOn like (?) and fldCreatedBy like (?))
                                                """,str(stage_id0),texts,str(dt),'PRAVIN SUBRAMANIAN',str(stage_id0),texts,str(dt),'PRAVIN SUBRAMANIAN')

                                                conn_str.commit()

                                                cust_rev = pd.read_sql("SELECT fldCustomerName,fldCustomerID FROM [XXXX_Email].[dbo].[tblCustomer] where fldCustomerName = ?", conn_str, params=[texts])
                                    # print(cust_rev[['fldCustomerName','fldCustomerID']])
                                    if ('IMO NO.' in text_ent) and ('VESS' in text_ent) and ((len(imo_data))==0) and ((len(df_final_rev))==0):
                                        
                                        vessID_serial = pd.read_sql("SELECT MAX(fldVesselID) as max_vessid FROM [XXXX_Email].[dbo].[tblVessel]", conn_str)
                                        
                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblVessel] (fldVesselID,fldVName,fldVIMO,fldCustomerID,MergedOn,MergedBy) 
                                        Select ?,?,?,?,?,?
                                        Where not exists(select fldVName,fldVIMO,fldCustomerID from [XXXX_Email].[dbo].[tblVessel]
                                        where fldVName like (?) or fldVIMO like (?) and fldCustomerID like (?))
                                        """,(int(vessID_serial['max_vessid'][0])+1),text_ent.get('VESS'),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]),str(dt),'PRAVIN SUBRAMANIAN',text_ent.get('VESS'),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]))
                                        conn_str.commit()

                                        vess_stage_id = pd.read_sql("SELECT fldVName,fldVIMO,fldVesselID FROM [XXXX_Email].[dbo].[tblVessel] where fldVName = ? and fldVIMO = ?", conn_str, params=[text_ent.get('VESS'),text_ent.get('IMO NO.')])
                                        vess_stage_id0 = vess_stage_id['fldVesselID'][0]

                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblVesselStage] (fldVesselID,fldVName,fldVIMO,fldCustomerID,fldCreatedOn,fldCreatedBy) 
                                        Select ?,?,?,?,?,?
                                        Where not exists(select fldVesselID,fldVName,fldVIMO,fldCustomerID from [XXXX_Email].[dbo].[tblVesselStage]
                                        where fldVName like (?) or fldVIMO like (?) and fldCustomerID like (?) and fldVesselID like (?))
                                        """,str(vess_stage_id0),text_ent.get('VESS'),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]),str(dt),'PRAVIN SUBRAMANIAN',text_ent.get('VESS'),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]),str(vess_stage_id0))
                                        conn_str.commit()

                                        imo_data = pd.read_sql("SELECT TOP 1 fldVesselID, fldVName, fldVIMO, fldCustomerID FROM [XXXX_Email].[dbo].[tblVessel] where fldVIMO = ?", conn_str, params=[text_ent.get('IMO NO.')])
                                        imo_data['matching_ratio'] = 100.0

                                        df_final_rev = pd.read_sql("SELECT TOP 1 fldVesselID, fldVName, fldVIMO, fldCustomerID FROM [XXXX_Email].[dbo].[tblVessel] where fldVName = ?", conn_str, params=[text_ent.get('VESS')])
                                        df_final_rev['matching_ratio'] = 100.0

                                        G_Var = 'NO'

                                    elif ('IMO NO.' in text_ent) and ('VESS' not in text_ent) and ((len(imo_data))==0):
                                        vessID_serial = pd.read_sql("SELECT MAX(fldVesselID) as max_vessid FROM [XXXX_Email].[dbo].[tblVessel]", conn_str)
                                    
                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblVessel] (fldVesselID,fldVIMO,fldCustomerID,MergedOn,MergedBy) 
                                        Select ?,?,?,?,?
                                        Where not exists(select fldVIMO,fldCustomerID from [XXXX_Email].[dbo].[tblVessel]
                                        where fldVIMO like (?) and fldCustomerID like (?))
                                        """,(int(vessID_serial['max_vessid'][0])+1),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]),str(dt),'PRAVIN SUBRAMANIAN',text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]))
                                        conn_str.commit()

                                        vess_stage_id = pd.read_sql("SELECT fldVIMO,fldVesselID FROM [XXXX_Email].[dbo].[tblVessel] where fldVIMO = ?", conn_str, params=[text_ent.get('IMO NO.')])
                                        vess_stage_id0 = vess_stage_id['fldVesselID'][0]
                                        
                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblVesselStage] (fldVesselID,fldVIMO,fldCustomerID,fldCreatedOn,fldCreatedBy) 
                                        Select ?,?,?,?,?
                                        Where not exists(select fldVesselID,fldVIMO,fldCustomerID from [XXXX_Email].[dbo].[tblVesselStage]
                                        where  fldVesselID like (?) and fldVIMO like (?) and fldCustomerID like (?))
                                        """,str(vess_stage_id0),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]),str(dt),'PRAVIN SUBRAMANIAN',str(vess_stage_id0),text_ent.get('IMO NO.'),str(cust_rev['fldCustomerID'][0]))
                                        conn_str.commit()

                                        imo_data = pd.read_sql("SELECT TOP 1 fldVesselID, fldVName, fldVIMO, fldCustomerID FROM [XXXX_Email].[dbo].[tblVessel] where fldVIMO = ?", conn_str, params=[text_ent.get('IMO NO.')])
                                        imo_data['matching_ratio'] = 100.0

                                        df_final_rev = pd.DataFrame(columns=['fldVesselID', 'fldVName', 'fldVIMO', 'fldCustomerID', 'matching_ratio'])
                                        G_Var = 'NO'

                                    elif ('IMO NO.' not in text_ent) and ('VESS' in text_ent) and ((len(df_final_rev))==0):
                                        
                                        vessID_serial = pd.read_sql("SELECT MAX(fldVesselID) as max_vessid FROM [XXXX_Email].[dbo].[tblVessel]", conn_str)
                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblVessel] (fldVesselID,fldVName,fldCustomerID,MergedOn,MergedBy) 
                                        Select ?,?,?,?,?
                                        Where not exists(select fldVName,fldCustomerID from [XXXX_Email].[dbo].[tblVessel]
                                        where fldVName like (?) and fldCustomerID like (?))
                                        """,(int(vessID_serial['max_vessid'][0])+1),text_ent.get('VESS'),str(cust_rev['fldCustomerID'][0]),str(dt),'PRAVIN SUBRAMANIAN',text_ent.get('VESS'),str(cust_rev['fldCustomerID'][0]))
                                        conn_str.commit()

                                        vess_stage_id = pd.read_sql("SELECT fldVName,fldVesselID FROM [XXXX_Email].[dbo].[tblVessel] where fldVName = ?", conn_str, params=[text_ent.get('VESS')])
                                        vess_stage_id0 = vess_stage_id['fldVesselID'][0]

                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblVesselStage] (fldVesselID,fldVName,fldCustomerID,fldCreatedOn,fldCreatedBy) 
                                        Select ?,?,?,?,?
                                        Where not exists(select fldVesselID,fldVName,fldCustomerID from [XXXX_Email].[dbo].[tblVesselStage]
                                        where fldVesselID like (?) and fldVName like (?) and fldCustomerID like (?) )
                                        """,str(vess_stage_id0),text_ent.get('VESS'),str(cust_rev['fldCustomerID'][0]),str(dt),'PRAVIN SUBRAMANIAN',str(vess_stage_id0),text_ent.get('VESS'),str(cust_rev['fldCustomerID'][0]))
                                        conn_str.commit()

                                        imo_data = pd.DataFrame(columns=['fldVesselID', 'fldVName', 'fldVIMO', 'fldCustomerID', 'matching_ratio'])

                                        df_final_rev = pd.read_sql("SELECT TOP 1 fldVesselID, fldVName, fldVIMO, fldCustomerID FROM [XXXX_Email].[dbo].[tblVessel] where fldVName = ?", conn_str, params=[text_ent.get('VESS')])
                                        df_final_rev['matching_ratio'] = 100.0
                                        G_Var = 'NO'




                                #If the customer entity is not detected in email body, then pass just empty dataframe
                                else: 
                                    cust_rev = pd.DataFrame(columns=['fldCustomerName','fldCustomerID'])
                                # print(tabulate(cust_rev, headers='keys', tablefmt='psql'))
#******************************************End of Part 3****************************************************************************************************************************************#

#*********************************************Part 4********************************************************************************************************************************************#                                                            

    #***************************************Code to find the Vessel ID and the list of customers for the vessel.*********************************/
                                Vess_data = imo_data if (len(imo_data))>0 else df_final_rev

                                #If the Vessel name/ IMO No is empty, then create a empty dataframe to show the vessel details as null in Quotes.
                                empty_df = pd.DataFrame(columns=['fldVesselID', 'fldVName', 'fldVIMO','matching_ratio'])
                                Vess_data = Vess_data if (len(Vess_data))>0 else empty_df
                                # print(Vess_data)
                                # print(tabulate(Vess_data, headers='keys', tablefmt='psql')) #Here it comes matching ratio sorted in desc

                                if G_Var == 'YES':
                                    # vessID= Vess_data['fldVesselID'][0] if (len(Vess_data['fldVesselID']))==1  else 'null'
                                    cust_ass = pd.read_sql("SELECT distinct  fldAssetID as fldVesselID, fldCustomerID FROM [XXXX_Email].[dbo].[tblCustAsset] where fldIHSMarkit=1", conn_str)
                                    Vess_Cust = duckdb.query("select fldVesselID, fldVName, fldVIMO, fldCustomerID,matching_ratio from (select * from Vess_data left join cust_ass on Vess_data.fldVesselID = cust_ass.fldVesselID)").df()
                                
                                
                                else:
                                    Vess_Cust = duckdb.query("select fldVesselID, fldVName, fldVIMO, fldCustomerID,matching_ratio from Vess_data").df()
                                # print(cust_ass, 'its')
                                Vess_Cust = Vess_Cust.sort_values(by="matching_ratio",ascending=False).reset_index()
                                # print('Passed line 574 Vessel details')
                                # print(tabulate(Vess_Cust, headers='keys', tablefmt='psql'))
#*****************************************End of Part 4 (Vessel Info)*******************************************************************************#

#*******************************************Part 5 (Cust Contact)***********************************************************************************#
    #****************************************Found the List of customers for the vessel.*******************************************************/
    #****************************************Code to find the customer contact details**************************************************/
                                cust_cont = pd.read_sql("SELECT fldCustContactID,  fldEmail, fldCustomerID FROM tblCustContact", conn_str)
                                cusID= cust_rev['fldCustomerID'][0] if (len(cust_rev['fldCustomerID']))==1  else 'null'
                                
                                #Getting the list of contact details for one customer ID.
                                cust_contact = duckdb.query(f"select * from cust_cont where cust_cont.fldCustomerID like {cusID}").df()
                                # print(cust_contact)
                                
                                for i in check:
                                    if i in cust_contact['fldEmail']:
                                        cust_contact_det = duckdb.query(f"select * from cust_contact where cust_cont.fldEmail = {i}").df()
                                        # print(cust_contact_det)
                                    else:
                                        # print('its here')
                                        # print(db_email_add, ' its ', db_email_nameadd)
                                        #Updating the records in the internal table
                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblCustContactInternal] (fldEmail,fldCustomerID,fldCName) 
                                        Select ?,?,?
                                        Where not exists(select fldEmail,fldCustomerID,fldCName from [XXXX_Email].[dbo].[tblCustContactInternal] 
                                        where fldEmail like (?) and fldCustomerID like (?) and fldCName like (?))
                                        """,db_email_add,str(cusID),db_email_nameadd,db_email_add,str(cusID),db_email_nameadd)
                                        
                                        #Updating the records in Main Cust Conatct table
                                        cursor.execute(""" 
                                        insert into [XXXX_Email].[dbo].[tblCustContact] (fldEmail,fldCustomerID,fldContCreatedDate,fldContCreatedBy) 
                                        Select ?,?,?,?
                                        Where not exists(select fldEmail,fldCustomerID from [XXXX_Email].[dbo].[tblCustContact] 
                                        where fldEmail like (?) and fldCustomerID like (?))
                                        """,db_email_add,str(cusID),str(dt),'PRAVIN SUBRAMANIAN',db_email_add,str(cusID))

                                        conn_str.commit()

                                        cust_contact_det = pd.read_sql("SELECT fldCustContactID,fldEmail,fldCustomerID FROM [XXXX_Email].[dbo].[tblCustContact] where fldCustomerID = ? and fldEmail = ?", conn_str, params=[str(cusID),(db_email_add)])
                                if (len(check))==0:
                                    cust_contact_det = pd.DataFrame(columns=['fldCustContactID', 'fldEmail', 'fldCustomerID'])

#***************************************End of Part 5 (Cust Contact)********************************************************************************************************#
# 
# ****************************************Part 6 (Cust Asset)***************************************************************************************************************#                           
                                # print(tabulate(Vess_Cust, headers='keys', tablefmt='psql'))
                                Creator_data['fldCustomerID'] = cusID
                                vess_chk = duckdb.query(f"select * from Vess_Cust where fldCustomerID like {cusID}").df()
                                
                                print(vess_chk)
                                
                                #Condition to append the vessel data in Cust Asset table which doesn't have the vessel info previously in it.
                                if (len(cust_rev)) > 0 and (len(vess_chk)) == 0 and (len(Vess_Cust)) > 0:
                                    # print('Passed inside')
                                    vess_id = duckdb.query("select distinct fldVesselID from Vess_Cust").df()
                                    # print(str(vess_id['fldVesselID'][0]) + 'Cust ID : '+str(cusID))
                                    cursor.execute(""" 
                                    insert into [XXXX_Email].[dbo].[tblCustAsset] (fldAssetID,fldCustomerID,fldIHSMarkit,fldCreatedOn,fldCreatedBy) 
                                    Select ?,?,?,?,?
                                    Where not exists(select fldAssetID,fldCustomerID,fldIHSMarkit from [XXXX_Email].[dbo].[tblCustAsset] 
                                    where fldAssetID like (?) and fldCustomerID like (?) and fldIHSMarkit like (?))
                                    """,str(vess_id['fldVesselID'][0]),str(cusID),str('1'),str(dt),'PRAVIN SUBRAMANIAN',str(vess_id['fldVesselID'][0]),str(cusID),str('1'))

                                    conn_str.commit()

                                    vess_pull = pd.read_sql("SELECT fldAssetID as fldVesselID, fldCustomerID FROM [XXXX_Email].[dbo].[tblCustAsset] where fldCustomerID = ? and fldIHSMarkit = ?", conn_str, params=[str(cusID),str(1)])
                                    Vess_Cust = duckdb.query("select fldVesselID, fldVName, fldVIMO, fldCustomerID from (select * from Vess_data left join vess_pull on Vess_data.fldVesselID = vess_pull.fldVesselID)").df()
                                #******************************End of Vessel addition to Cust Asset********************
#****************************************End of Part 6 (Cust Asset)************************************************************************************************************#


#*******************************************Final Joins***********************************************************************************************************************#
                                cust_vess_det = duckdb.query("select fldCustContactID,fldCustomerID,fldEmail,fldVesselID,fldVName,fldVIMO from (select * from cust_contact_det left join Vess_Cust on cust_contact_det.fldCustomerID = Vess_Cust.fldCustomerID)").df()
                                cust_det = duckdb.query(f"select fldCustContactID,fldCustomerID, fldCustomerName ,fldEmail,fldVesselID,fldVName,fldVIMO from (select * from cust_rev left join cust_vess_det on cust_rev.fldCustomerID = cust_vess_det.fldCustomerID)").df()
                                Cust_quote = duckdb.query("select distinct fldCustomerID, fldCustomerName, fldCustContactID, fldEmail,fldVesselID,fldVName,fldVIMO,fldCreatedOn,fldCreatedBy from (select * from cust_det left join Creator_data on cust_det.fldCustomerID = Creator_data.fldCustomerID)").df()

#********************************************End of Final Joins***************************************************************************************************************#                                                            
                                

                                if (len(Cust_quote))>0:
                                    # print(tabulate(Cust_quote, headers='keys', tablefmt='psql'))
                                    

                                    cursor.execute(""" 
                                    insert into [XXXX_Email].[dbo].[tblQuoteStage] (fldCustomerID,fldAssetID,fldCustContactID,fldCreatedOn,fldCreatedBy) 
                                    Select ?,?,?,?,?
                                    Where not exists(select fldCustomerID,fldAssetID,fldCustContactID from [XXXX_Email].[dbo].[tblQuoteStage] 
                                    where fldCustomerID like (?) and fldAssetID like (?) and fldCustContactID like (?) )
                                    """,str(Cust_quote['fldCustomerID'][0]),str(Cust_quote['fldVesselID'][0]),str(Cust_quote['fldCustContactID'][0]),str(dt),'PRAVIN SUBRAMANIAN',str(Cust_quote['fldCustomerID'][0])
                                    ,str(Cust_quote['fldVesselID'][0]),str(Cust_quote['fldCustContactID'][0]))
                                    
                                    conn_str.commit()

                                    Quote_id = pd.read_sql("SELECT distinct  fldQuoteStageID,fldCustomerID,fldCreatedOn,fldCreatedBy FROM [XXXX_Email].[dbo].[tblQuoteStage] where fldCustomerID = ? and fldAssetID = ? and fldCustContactID = ?", conn_str, params=[str(Cust_quote['fldCustomerID'][0]),str(Cust_quote['fldVesselID'][0]),str(Cust_quote['fldCustContactID'][0])])
                                    Quote_id_0 = Quote_id['fldQuoteStageID'][0]
                                    Quote_CreatedOn = Quote_id['fldCreatedOn'][0]
                                    Quote_CreatedBy = Quote_id['fldCreatedBy'][0]
                                    # Quote = duckdb.query("select fldCustomerID as CustomerID, fldCustomerName as Customer_Name, fldVIMO, fldCustomerID from (select * from Vess_data left join vess_pull on Vess_data.fldVesselID = vess_pull.fldVesselID)").df()
                                    Cust_quote = Cust_quote.rename(columns={"fldCustomerName": "Customer Name", "fldEmail":"Email","fldVName":"Vessel Name","fldVIMO":"Vessel IMO"})
                                    Orig_Quote = Cust_quote[['Customer Name','Email','Vessel Name','Vessel IMO']]
                                    Orig_Quote['Quote ID'] = Quote_id_0
                                    Orig_Quote['Created On'] = Quote_CreatedOn
                                    Orig_Quote['Created By'] = Quote_CreatedBy
                                    Orig_Quote = Orig_Quote[['Quote ID','Customer Name','Email','Vessel Name','Vessel IMO','Created On','Created By']]



                                    html_table = Cust_quote.to_html(index=False)

                                    columns = list(Orig_Quote.columns)
                                    values = list(Orig_Quote.values.flatten())


                                    # Create the HTML email body
                                    html_body = '''
                                    <html>
                                    <head>
                                    <style>
                                    table, th, td {{
                                    border: 1px solid black;
                                    border-collapse: collapse;
                                    text-align: center;
                                    padding: 5px;
                                    }}
                                    th {{
                                    font-weight: bold;
                                    }}
                                    </style>
                                    <td> <span style="font-size:14px;font-family: Calibri;float:left;"> A new quote has been prepared by {creatBy} in the XXX DB with the following details: </span> </td>
                                    </head>
                                    <body>
                                    <td></td><br><br>
                                    <h2>Pre Quote Information</h2>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Quote No.: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{QuoteID} </b></span></td><br>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Customer Name: &nbsp;&nbsp;&nbsp;&nbsp;<b>{Custname} </b></span></td><br>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Customer Email: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{email} </b></span></td><br>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Asset Name: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{vessname} </b></span></td><br>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Asset IMO No.: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{vessimo} </b></span></td><br>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Quote Created On: <b>{creatOn} </b></span></td>
                                    <br><br><br><br><td><span style="font-size:15px;font-family: Calibri;float:left;">Please Open XXX Database to view quote.
                                    <br><br>Thanks and regards,<br>SYSTEM ADMINISTRATOR</span></td>
                                    <table>
                                    
                                    </table>
                                    </body>
                                    </html>
                                    '''.format( #rows='\n'.join(['<tr><td>{0}</td><td>{1}</td></tr>'.format(c, v) for c, v in zip(columns, values)]),
                                            QuoteID = Orig_Quote['Quote ID'][0], Custname = Orig_Quote['Customer Name'][0],email = Orig_Quote['Email'][0],vessname = Orig_Quote['Vessel Name'][0]
                                            ,vessimo = Orig_Quote['Vessel IMO'][0],creatOn = Orig_Quote['Created On'][0],creatBy = 'SA')

                                    mail.Subject = 'Quote No: ' + str(Quote_id_0) + ' for Customer ' + Cust_quote['Customer Name'][0].upper()
                                    mail.HTMLBody = html_body
                                    mail.Send()

                                    # print(Cust_quote)
                            except:
                                # Create the HTML email body
                                    html_body = '''
                                    <html>
                                    <head>
                                    <style>
                                    table, th, td {{
                                    border: 1px solid black;
                                    border-collapse: collapse;
                                    text-align: center;
                                    padding: 5px;
                                    }}
                                    th {{
                                    font-weight: bold;
                                    }}
                                    </style>
                                    <td> <span style="font-size:14px;font-family: Calibri;float:left;"> The Quote preparation failed for {ESender} </span> </td>
                                    </head>
                                    <body>
                                    <td></td><br><br>
                                    <h2>The Quote is not prepared for the below email</h2>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Email Sender: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>{ESender} </b></span></td><br>
                                    <td> <span style="font-size:15px;font-family: Calibri;float:left;"> Email Subject: &nbsp;&nbsp;&nbsp;&nbsp;<b>{ESubject} </b></span></td><br>
                                    <td><span style="font-size:15px;font-family: Calibri;float:left;">
                                    <br><br>Thanks and regards,<br>SYSTEM ADMINISTRATOR</span></td>
                                    <table>
                                    
                                    </table>
                                    </body>
                                    </html>
                                    '''.format(ESubject = message.Subject, ESender = email_sender)

                                    mail.Subject = 'The Quote preparation failed for' +  email_sender
                                    mail.HTMLBody = html_body
                                    mail.Send()




#********************************************************************End of Code*********************************************************************************************************************#