from exchangelib import Credentials, Account, Configuration, DELEGATE, FileAttachment
from email import message_from_file
import requests
import hashlib
import pyodbc
from datetime import datetime, timezone


# Variables go here
journal_mailboxes = ['journaledmail@corp.gslab.com']
username = 'gsladmin@corp.gslab.com'
password = 'PENT@gon500!!!'
server_url = 'http://exchange2019.eastus.cloudapp.azure.com/EWS/Exchange.asmx'
scratch_directory = 'C:/temp/'
start_time = datetime.strptime('2023-05-01T00:00:00', '%Y-%m-%dT%H:%M:%S').replace(tzinfo=timezone.utc)
end_time = datetime.strptime('2023-06-02T23:59:59', '%Y-%m-%dT%H:%M:%S').replace(tzinfo=timezone.utc)
sql_server = 'auditdemo.database.windows.net'
sql_database = 'auditdemo'
sql_username = 'sreese'
sql_password = 'ICN45admin199!'


# ESTABLISHING SERVICE CONNECTIONS TO EXCHANGE AND SQL
# Connect to SQL server
conn = pyodbc.connect(
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=tcp:"+ sql_server + ",1433;"
    "Database=" + sql_database +";"
    "Authentication=SQLPassword;"
    "Uid=" + sql_username + ";"
    "Pwd=" + sql_password + ";"
    )

# Create a cursor object to execute SQL commands
cursor = conn.cursor()

# Disable SSL certificate verification for Exchange connectivity due to certificate not created (remove in prod)
requests.packages.urllib3.disable_warnings()

# Set EWS credentials and server URL
credentials = Credentials(username=username, password=password)
ews_config = Configuration(service_endpoint=server_url, credentials=credentials)





# FUNCTIONS
# Function to create a new Audit ID record number
def create_audit_id():
    # Create Audit ID by checking for the max row in the sql table and adding 1
    cursor.execute("SELECT MAX(auditID) FROM dbo.emailauditdata")
    audit_id = cursor.fetchone()[0] + 1
    return audit_id

# Function to retrieve emailIDs
def get_inbox_ids(mailbox, ews_config):

    # Connect account using the configuration
    account = Account(primary_smtp_address=mailbox, config=ews_config, autodiscover=False, access_type=DELEGATE)

    # Get email Ids from the inbox
    inbox_folder = account.inbox
    email_ids = [item.id for item in inbox_folder.filter(datetime_received__range=(start_time, end_time))]
    return email_ids

# Function to retrieve eml file from jounaled message
def get_journaled_email(mailbox, ews_config, item_id):
    # Connect account using the configuration
    account = Account(primary_smtp_address=mailbox, config=ews_config, autodiscover=False, access_type=DELEGATE)
    email = account.inbox.get(id=item_id)
    return email

# Function to retrieve hash value of a file
def calculate_file_sha256(file_path):
    # Create a SHA-256 hash object
    sha256_hash = hashlib.sha256()

    # Read the file in chunks to efficiently handle large files
    with open(file_path, "rb") as file:
        chunk_size = 4096  # Read 4KB chunks at a time
        for chunk in iter(lambda: file.read(chunk_size), b""):
            sha256_hash.update(chunk)

    # Get the hexadecimal representation of the SHA-256 hash
    sha256_hash_value = sha256_hash.hexdigest()

    return sha256_hash_value

#Function to write Audit Report to SQL database
def write_email_audit(audit_id, audit_datetime, message_id, message_from, message_to, message_cc, message_bcc, message_subject, message_body, message_attachments, message_other, message_audit_result, mailbox):
    
    # Write SQL Datat to Database
    query = "INSERT INTO emailauditdata (auditId, AuditDateTime, MessageID, MessageSender, MessageTO, MessageCC, MessageBCC, MessageSubject, MessageBody, MessageAttachments, MessageOther, MessageAuditResult, Mailbox) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
    cursor.execute(query, audit_id, audit_datetime, message_id, message_from, message_to, message_cc, message_bcc, message_subject, message_body, message_attachments, message_other, message_audit_result, mailbox)

#Function to retreive jounaled email attachment
def open_eml_as_attachment(eml_file_path):
    with open(eml_file_path, 'rb') as file:
        eml_content = file.read()

    attachment = FileAttachment(name='eml_attachment.eml', content=eml_content)
    return attachment

# Function to compare two email files and report the results
def compare_email_files(file_path1, file_path2):
    # Parse EML files
    with open(file_path1, 'r') as file1:
        eml1 = message_from_file(file1)

    with open(file_path2, 'r') as file2:
        eml2 = message_from_file(file2)

    # Compare subject, sender, recipients, and body
    subject1 = eml1.get('subject', '')
    subject2 = eml2.get('subject', '')
    if subject1 != subject2:
        #print(f"Difference in subject:\nFile 1: {subject1}\nFile 2: {subject2}\n")
        subjectaudit = "Fail"
    else:
        #print("Subject is the same.")
        subjectaudit = "Pass"

    sender1 = eml1.get('from', '')
    sender2 = eml2.get('from', '')
    if sender1 != sender2:
        #print(f"Difference in sender:\nFile 1: {sender1}\nFile 2: {sender2}\n")
        senderaudit = "Fail"
    else:
        #print("Sender is the same.")
        senderaudit = "Pass"

    recipients1 = eml1.get_all('to', [])
    recipients2 = eml2.get_all('to', [])
    if recipients1 != recipients2:
        #print(f"Difference in recipients:\nFile 1: {recipients1}\nFile 2: {recipients2}\n")
        recipientaudit = "Fail"
    else:
        #print("Recipients are the same.")
        recipientaudit = "Pass"

    cc1 = eml1.get_all('cc', [])
    cc2 = eml2.get_all('cc', [])
    if cc1 != cc2:
        #print(f"Difference in cc:\nFile 1: {recipients1}\nFile 2: {recipients2}\n")
        ccaudit = "Fail"
    else:
        #print("CC are the same.")
        ccaudit = "Pass"
    
    bcc1 = eml1.get_all('bcc', [])
    bcc2 = eml2.get_all('bcc', [])
    if bcc1 != bcc2:
        #print(f"Difference in BCC:\nFile 1: {recipients1}\nFile 2: {recipients2}\n")
        bccaudit = "Fail"
    else:
        #print("BCC are the same.")
        bccaudit = "Pass"

    body1 = eml1.get_payload()
    body2 = eml2.get_payload()
    if body1 != body2:
        #print(f"Difference in body:\nFile 1:\n{body1}\nFile 2:\n{body2}\n")
        bodyaudit = "Fail"
    else:
        #print("Body is the same.")
        bodyaudit = "Pass"

    # Compare attachments
    attachments1 = get_attachments(eml1)
    attachments2 = get_attachments(eml2)

    if len(attachments1) != len(attachments2):
        #print("Difference in attachment count")
        attachmentaudit = "Fail"
    else:
        #print("Attachments are the same.")
        attachmentaudit = "Pass"

    #Check to see if the reason has been identified for the audit failure. If no then mark audit failure as other
    combinedaudit = [subjectaudit, senderaudit, recipientaudit, bodyaudit, attachmentaudit]
    for audit in combinedaudit:
        if audit == "Fail":
            otheraudit = "Pass"
        else:
            otheraudit = "Fail"
    
    return subjectaudit, senderaudit, recipientaudit, bodyaudit, attachmentaudit, ccaudit, bccaudit, otheraudit

#Function to retireve email attachments
def get_attachments(eml):
    attachments = []
    for part in eml.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            attachment = {
                'name': part.get_filename(),
                'size': len(part.get_payload(decode=True))
            }
            attachments.append(attachment)
    return attachments





# AUTOMATION CODE
# check for all messageIDs in Mailboxes one at a time
for journal_mailbox in journal_mailboxes:
    messageids = get_inbox_ids(journal_mailbox, ews_config)

    # check each mailbox item
    for messageid in messageids:
        # use function retrieve email from exchange
        email = get_journaled_email(journal_mailbox, ews_config, messageid)

        # Extract each email attachment from message save locally and provide the file hash for total file comparison
        for attachment in email.attachments:
            print("")
            print("**Begin new Audit**")
            print("MessageID: " + attachment.item.message_id)

            # save file locally
            with open(scratch_directory + 'attachment.eml', 'wb') as file:
                file.write(attachment.item.mime_content)
            
            exchangemessage_hashvalue = calculate_file_sha256(scratch_directory + 'attachment.eml')
            print("Exchange Message Hash Value: " + exchangemessage_hashvalue)

            # Extract matching messageID from the hitachi content platform
            # Enter the code here when done for now using a temp place holder file to emulate the process
            # The emulated file is stored at the scrath directory in a file called attachment2.eml
            obsmessage_hashvalue = calculate_file_sha256(scratch_directory + 'attachment2.eml')
            print("OBS Email Hash Value:" + obsmessage_hashvalue)

            # Compare the hash results and if results match write a passing email audit record tot he database
            if obsmessage_hashvalue == exchangemessage_hashvalue:
                print("The hash values match - AUDIT PASSED!")
                audit_id = create_audit_id()
                print("Creating Audit Record " + str(audit_id))
                write_email_audit(audit_id, datetime.now(), attachment.item.message_id, "Pass", "Pass", "Pass", "Pass", "Pass", "Pass", "Pass", "Pass", "Pass", journal_mailbox) 
                print("Completed writing audit record " + str(audit_id))
                print("")
            else:
                print("The hash values do no match - AUDIT FAILED!")
                audit_id = create_audit_id()
                print("Creating Audit Record " + str(audit_id))
                # Compare email files to identify delta
                auditresult = compare_email_files(scratch_directory + 'attachment.eml', scratch_directory + 'attachment2.eml')
                write_email_audit(audit_id, datetime.now(), attachment.item.message_id, auditresult[0], auditresult[1], auditresult[2], auditresult[3], auditresult[4], auditresult[5], auditresult[6], auditresult[7], "Fail", journal_mailbox)
                print("Completed writing audit record " + str(audit_id))
                print("")
#1










