# Email Sender for CQU@MSC
## Tools
- Using ali notification mailbox as the default mailbox  
- Using SMTP to send emails
- Using Word docfile as the email template

## How to use
1. Change the password of the mailbox. (If U use another mailbox, change the address too)
2. Modify `to_addrs`, include all the receivers' mailbox address.
3. Modify the title of the mail you want to send in the main function.
4. Give the disk address of the email content docfile in the main function.
5. Use `python main.py` to run the script.

## Building
You may need python3 with the following packs
- python-docx
- more packs i forgot (plz submit issues)