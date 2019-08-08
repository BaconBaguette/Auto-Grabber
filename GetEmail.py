from exchangelib import Credentials, Account, FileAttachment, ItemAttachment, Message
import datetime
import EmailLookup
import configparser
import os


config = configparser.ConfigParser()

config.read('AutoTest.INI')

lookup = EmailLookup.lookupDict

now = datetime.datetime.now()   # Gets the current date time
errNow = now.strftime("%H:%M:%S  %d/%m/%Y") # Converts datetime to format to be used in logging
fNow = now.strftime("%Y%m%d_%H%M%S")    # Converts datetime to format to be used in file naming

email_add = config['login']['email']      
email_pass = config['login']['emailpass']

directory = config['paths']['targetdir']     # The directory in which email attachments will be saved

with open('whitelist.txt', 'r') as f:   # Creates the whitelist from a local 'whitelist.txt' file
    whitelist = f.read().splitlines()

credentials = Credentials(email_add, email_pass)    # Create the credential class used for the account

try:
    account = Account(email_add, credentials=credentials, autodiscover=True)      # Tries logging in with the given credentials and writes an error to the log if it fails
except:
    with open('log.txt', 'a') as f:
        f.write(errNow)
        f.write(' - ERROR - Problem connecting to exchange server. This could be due to incorrect credentials, corrupt temporary files or a network issue.')
        f.write('\n')
    quit()

try:
    searchFolder = account.root / 'Top of Information Store' / folderName      # Attempts to connect to the mailbox. If it fails connects to inbox instead. ## THIS MIGHT WANT TO BE REMOVED ##
except:
    searchFolder = account.inbox

def dlAtt(folder, acc):     # Runs through the specified mailbox and downloads all attachments with a specified attachment type from whitelisted addresses
    for item in folder.all():
        count = 1
        if item.sender.email_address in whitelist:
            customer = lookup[item.sender.email_address]
            for attachment in item.attachments:            # FileType here
                attName = attachment.name.upper()
                if isinstance(attachment, FileAttachment) and 'GOUNITPO' in attName and 'CSV' in attName:
                    filename = 'UnitPO-' + customer + '-' + fNow + '_' + str(count) + '.csv'
                    local_path = directory + customer + '\\' + filename
                    if not os.path.exists(directory + customer):
                        os.makedirs(directory + customer)
                    try:
                        with open(local_path, 'wb') as f:
                            f.write(attachment.content)
                        with open('log.txt', 'a') as l:     # Writes to log, datetime downloaded, sender and attachment name.
                            l.write(errNow)
                            l.write(' - Saved attachment ')
                            l.write(filename)
                            l.write(' in ')
                            l.write(directory)
                            l.write('. Sent from ')
                            l.write(item.sender.email_address)
                            l.write('\n') 
                        count = count + 1
                    except:
                        with open('log.txt', 'a') as l:
                            l.write(errNow)
                            l.write(' - ERROR - Problem writing attachment to local drive.\n')  
            item.move(account.inbox / 'Imported EDIs')       

try:
    dlAtt(searchFolder, account)
except:
    with open('log.txt', 'a') as l:
        l.write(errNow)
        l.write(' - ERROR - Unspecified problem downloading attachments.\n')
