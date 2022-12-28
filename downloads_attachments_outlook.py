'''
    File name: download_attachments_outlook.py
    Author: Marcos Samuel Mattos Santos
    Date created: 11/23/2021
    Date last modified: 12/27/2022
    Python Version: 3.8.8
'''

import os
import win32com.client # needs installation (pywin32)
import re
import unidecode # needs installation
import getpass


downloads_dir_email = r'C:\Users\USER\TYPE HERE THE REST OF YOUR DIRECTORY'

usuario = getpass.getuser()

downloads_dir_email = re.sub('USER',usuario,downloads_dir_email)


def download_attachments_outlook(directory, search_term_subject = '', search_term_body = '', search_term_attachment = ''):
    '''
    This function is pretty simple and receives 4 parameters.
    Only the first one is required to make it run. The other ones are optional.
    See details below:

    directory: system folder to save the attachment files
    search_term_subject: string regular expression to search email's subject
    search_term_body: string regular expression to search email's body
    search_term_attachment: string regular expression to search email's attachments

    Overall:
    The program enters your outlook application inbox and loops through all messages.
    For each one, it'll test if the desired Regular Expression is present either on subject or body.
    If True, it'll loop through the attachments, searching for the desired Regular Expression.
    If it's True again, it'll remove any invalid character of the attachment name and proceed to the next condition.
    If the string length is greater than 156 (max length for naming windows files), it will cut the first (n - 156) characters and strip (remove edge spaces).
    Then, it'll save the attachment to the desired directory.

    If no attachments were saved, it'll return a message informing it.
    Else, a list containing other lists which contains the message subject and the attachment name will be returned to the user.
    '''

    # entering outlook
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')

    # entering inbox
    inbox = outlook.GetDefaultFolder(6)

    # getting inbox items and sorting by received time (most recent to oldest)
    e_mails = inbox.Items
    e_mails.Sort('[ReceivedTime]',True)

    # creating empty list
    list_messages = []

    # looping through all e-mails
    for counter in range(len(e_mails)):
        message = e_mails[counter]
        # testing re
        if re.search(search_term_subject,message.subject) or re.search(search_term_body,message.body):
            # looping through attachments
            for item in range(1,len(message.Attachments) + 1):
                attachment = message.Attachments.Item(item)
                attachment_name = str(attachment)
                # testing re
                if re.search(search_term_attachment, attachment_name):
                    attachment_name = re.sub("/|\\\|\\?|\\*|\\:|<|>|\\|","-",attachment_name)
                    # cutting and stripping big names
                    if len(attachment_name) > 156:
                        attachment_name = attachment_name[-156:].strip()
                    path = os.path.join(directory,attachment_name)
                    print(f'Subject: {message.subject}\nIndex: {counter}\nAttachment Name: {attachment_name}\nPath: {path}\n\n')
                    # saving attachment
                    attachment.SaveAsFile(path)
                    list_messages.append([message.subject,attachment_name])

    if len(list_messages) == 0:
        return 'No attachments downloaded'

    return list_messages


if __name__ == '__main__':
    test = download_attachments_outlook(directory = downloads_dir_email,search_term_subject = '', search_term_body = '', search_term_attachment = '')
    test
