from converter import converter
import telebot
import os

# The token is stored in a file called apicode.txt
with open('apicode.txt', 'r') as f:
    mytoken = f.read()    


bot = telebot.TeleBot(mytoken)

welcome_message = '''

Hello! I'm a bot that converts excel files to vcf files. Send me an excel file and I'll send you a vcf file.

Please note that the excel file must be in the following format:

Name     | Phone number
John Doe | 123456789
Jane Doe | 987654321

The excel file must be in .xlsx format. The vcf file will be sent to you in the same chat.
You should open the vcf file with your phone's contacts app and save the contacts to your phone.
Have a nice day!

'''


@bot.message_handler(commands=['help'])
def send_welcome(message):
    bot.reply_to(message, welcome_message)

@bot.message_handler(content_types=['document'])
def send_content(message):
    '''This function is called when the bot receives a document.'''
    # Check if the document is an excel file
    if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        bot.reply_to(message, "File received. Converting to vcf...")

        file_info = bot.get_file(message.document.file_id)
        # Download the file
        downloaded_file = bot.download_file(file_info.file_path)
        with open('contacts.xlsx', 'wb') as new_file:
            new_file.write(downloaded_file)
        # Convert the file
        c = converter('contacts.xlsx')
        c.create_vcard_file('contacts.vcf','artitu')
        # Send the file
        with open('contacts.vcf', 'rb') as f:
            bot.send_document(message.chat.id, f)

        bot.reply_to(message, "File sent. We're hoping you're satisfied with our service. Have a nice day!")
        os.remove('contacts.xlsx')
        os.remove('contacts.vcf')
    

while True:
    try:
        bot.polling(none_stop=True)
        
    except Exception as e:
        print(e)

