import win32com.client
import re


OUTLOOK_SUBFOLDER_NAME = "arxiv"
OUTLOOK_FOLDER_NAME = "papers"


# Regular expression pattern to extract titles and authors
PATTERN = r"Title:(.*?)Authors:(.*?)Categories:" 
PARSED_FILENAME = "arxiv_parsed.txt"


ADDRESS_LIST = ["thomas_a_hahn@mac.com", "itammar.steinberg@weizmann.ac.il"]
EMAIL_TITLE = "Parsed daily arxiv from Ilya"

def extract_titles_and_authors(input_text):
    # Find all matches using re.findall
    matches = re.findall(PATTERN, input_text, re.DOTALL)

    # Clean up the extracted titles and authors
    papers_info = [(title.strip(), authors.strip()) for title, authors in matches]
    # Remove newlines
    papers_info = [(title.replace('\n', ' '), authors.replace('\n', ' ')) for title, authors in papers_info]
    # Remove consecutive blank spaces
    papers_info = [(re.sub(r'\s+', ' ', title), re.sub(r'\s+', ' ', authors)) for title, authors in papers_info]

    return papers_info


def mail_import():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook.Folders.Item(1)  # Access the root folder

    # Replace "papers" with the name of your main folder
    papers_folder = root_folder.Folders(OUTLOOK_FOLDER_NAME)

    # Replace "arxiv" with the name of your subfolder
    arxiv_folder = papers_folder.Folders(OUTLOOK_SUBFOLDER_NAME)

    messages = arxiv_folder.Items
    message_count = messages.Count

    email_list = list()
    for i in range(message_count):
        message = messages.Item(i + 1)
        
        # Check if the email is unread (not opened)
        if not message.UnRead:
            continue  # Skip this email if it has been opened
        
        # Mark the email as read (opened)
        message.UnRead = False
        message.Save()
        email_list.append(message.Body)
    return email_list
    
def body_generator(mail_list):
    body = ""
    title_author_pairs = list()
    for mail in mail_list:
        title_author_pairs += extract_titles_and_authors(mail)
    for title, authors in title_author_pairs:
        body += title
        body += '\n'
        body += authors
        body += '\n'
        body += '\n'
        body += '\n'
    return body

def write_data(body):
    with open(PARSED_FILENAME, 'a') as f:
        f.write(body)
    f.close()

def dispatch(email_body):
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    if not email_body:
        return

    # Create a new email
    for address in ADDRESS_LIST:
        new_mail = outlook.CreateItem(0)  # 0 represents the Outlook MailItem

        # Replace the following details with your information
        new_mail.Subject = EMAIL_TITLE
        new_mail.Body = email_body
        new_mail.To = address

        # Send the email
        new_mail.Send()
        print(f"Email sent to {address}")

def parse():
    mail_list = mail_import()
    body = body_generator(mail_list)
    write_data(body)
    dispatch(body)

parse()