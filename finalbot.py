import sharepy
import sys
import time
import random
from slackclient import SlackClient
from fuzzywuzzy import process


#import webbrowser

#chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'

#

#webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))

try:
    
    
    conn = sharepy.connect("https://jda365.sharepoint.com", username = "raguram.s@jda.com", password = "FerHomme43")    
    
    site = "https://jda365.sharepoint.com/sites/CategoryManagement/"    
    
    library = "Documents"    
    
    files = conn.get("{}/_api/web/lists/getbyTitle('{}')/Items?$select=FileLeafRef,FileRef,Id&$top=5000"
		  .format(site, library)).json()["d"]["results"]   
    
    array = []
    
    Mapper = {}    
    
    farewell = ["Bye", "Good Bye", "Until next time...", "Au revoir", "Adios", "Ciao", "See you soon", "Auf Wiedersehen..!"]
            
            
    for file in files:
        
                
        p = file['FileLeafRef']
        
        array.append(p)
        
        source = "https://" + conn.site + file["FileRef"]

#        folder = conn.get("{}/_api/web/GetFolderByServerRelativeUrl('{}')"
#                         .format(site, file["FileRef"])).status_code == 200                  

        Mapper[p] = source
                
   
    
    print('Data build successful...')
    
    
    def get_bot_id():
        
        api_call = slack_client.api_call('users.list')
        
        if api_call.get('ok') :
            
            users = api_call.get('members')
            
            for user in users :
                
                if 'name' in user and user.get('name') == bot_name :
                    
                    return '<@' + user.get('id') + '>'
    
        return None
    
    
    def search():
        
        return 'Mention the file you would like to search'
    
    def close():
        
        return random.choice(farewell) + ', ' + "I'm only a text away...!"
    
    def help():
        
        response = 'You can use the following commands : \r\n'
        
        for command in commands:
            
            response = response + command + '\r\n'
            
        response = response + 'In case you are using search command, type the file to be searched followed by the search command' + '\r\n' + "(eg) search hook.zip"
        
        return response
    
    
    
    def handle_command(user, command) :
        
        response = '<@' + user + '> '
        
        if command in commands :
            
            response = response + commands[command]()
            
        else :
            
            response = response + "Sorry, I am not sure what you mean by " + "'" + command + "'" + "." + '\r\n' + 'Type help to know the supported commands'
            
            
        return response
    
    
    def command_handler(user, command, contents):
        
        response = '<@' + user + '> '
        
        r1 = process.extractOne(contents , array)
        
        if r1[1] > 75 :
            
            file_name = r1[0] 
            
            res = Mapper[file_name]
            
            res = res.replace(' ', '%20')
            
            response = response + 'Click the link below to open the file'+ ' - ' +file_name + '\n' + res

#            webbrowser.get('chrome').open_new_tab(res)
            
        elif ( r1[1] > 65 and r1[1] < 76 ) :
            
            response = response + 'A liitle more accuracy of the file name would be appreciated...!'
            
            
        else :
            
            response = response + "I don't think such a directory exists in your Sharepoint directory...!"
        
        
        return response
    
    

    def handle_event(user, command, channel) :
        
        if command and channel :
            
            print('Received command ' + command + ' in channel ' + channel + ' from user ' + user)
            
            result = command
            
            if 'search' in command :
                
                command = 'search'
                
                resultant  = result.split('search')
                
                contents = resultant[1]
                
                contents = contents.lstrip()
                
                if contents == '' :
                    
                    response = handle_command(user, command)
                    
                else : 
                    
                    response = command_handler(user, command, contents)
                    
            else : 
                
                response = handle_command(user, command)
                
            slack_client.api_call('chat.postMessage', channel = channel, text = response, as_user = True)
                    
                    
    
    
    def parse_event(event):
        
        if event and 'text' in event and bot_id in event['text'] :
            
            handle_event(event['user'], event['text'].split(bot_id)[1].strip().lower(), event['channel'])
            
    
    
    def wait_for_event():
        
        events = slack_client.rtm_read()
        
        if events and len(events) > 0 : 
            
            for event in events :
                
                print(event)
                
                parse_event(event)
    
    
    def listen():
        
        if slack_client.rtm_connect(with_team_state = True):
            
            print('\r\n')
            
            print('Connection established Successfully...Listening for commands...!')
            
            while(True) :
                
                wait_for_event()
                
                time.sleep(1)
                
        else :
            
            exit('Error, Connection could not be established...!')
                
    
    slack_client = SlackClient('xoxb-579364393873-597438149168-pHJc50Vl8nX7sIleiptxSc87')
    
    bot_name = 'jarvis'
    
    bot_id = get_bot_id()
    
    commands = {
            
            'search' : search,
            
            'help'   : help,
            
            'bye'    : close
            } 
    
    if bot_id is None :
        
        exit('Error, could not find ' + bot_name)
        

    try :
        
        listen()
        
    except:
        
        print('\r\n')
        
        print('Ooops...An error occurred...!')
        
    
    
                
except:

    print('\r\n')
    
    print("Program Terminated Abruptly...!")
    
    sys.exit(0)