import os

import pandas as pd


def createhtmltables(pathToFolder, user, arrayContactos, arrayEventCall, arrayCallOneToOne,
                     dictionaryConversationDetails, tm, tmCSV):
    css_string = '''
        .mystyle {
            font-size: 11pt; 
            font-family: Arial;
            border-collapse: collapse; 
            border: 1px solid silver;
            width: 100%;
        }
        .mystyle th {
            background-color: royalblue;
            color: white;
        }

        .mystyle td, th {
            padding: 5px;
        }

        .mystyle tr:nth-child(even) {
            background: #f2f2f2;;
        }

        .mystyle tr:hover {
            background: silver;
            cursor: pointer;
        }
        '''
    with open(os.path.join(pathToFolder, "df_style_{}_{}.css".format(user, tm)), 'w') as f:
        f.write(css_string)

    html_string = '''
    <html>
      <head><title>{title}</title></head>
      <link rel="stylesheet" type="text/css" href="df_style_{user}_{tm}.css"/>
      <body>
        <h1>{title}</h1>
        {table}
      </body>
    </html>.
    '''

    pd.set_option('display.width', 1000)
    pd.set_option('colheader_justify', 'center')

    # -------------------------------------------------

    def first_element(elem):
        return elem[0]

    # Create html table to present Contacts
    contacts_dict = {}
    for k, v in arrayContactos.items():
        contacts_dict[k] = [v.nome, v.email, user]

    sorted_by_name_contacts = {}
    iterator = 1
    for k, v in sorted(contacts_dict.items(), key=lambda item: item[1][0].upper()):
        sorted_by_name_contacts[iterator] = v
        iterator = iterator + 1

    frame_contacts = pd.DataFrame.from_dict(sorted_by_name_contacts, orient='index', columns=['Name', 'Email', 'User'])

    with open(os.path.join(pathToFolder, "User_Contacts_{}_{}.html".format(user, tm)), 'w', encoding="utf-8") as file:

        file.write(
            html_string.format(title="CONTACTS", table=frame_contacts.to_html(classes='mystyle'), user=user, tm=tm))

    # Create html table to present Calls in Teams
    teams_calls = []
    for call_team in arrayEventCall:
        creator = call_team.creator.nome + ", " + call_team.creator.email
        new = [call_team.calldate, call_team.duration, call_team.count, creator, call_team.participants, user]
        teams_calls.append(new)

    teams_calls.sort(key=first_element)

    frame_team_calls = pd.DataFrame(teams_calls, columns=["End date", "Duration (minutes)", "Number of participants",
                                                          "Call Creator", "Call Participants", "user"])
    with open(os.path.join(pathToFolder, "Calls_Teams_{}_{}.html".format(user, tm)), 'w', encoding="utf-8") as file:
        file.write(
            html_string.format(title="TEAM CALLS", table=frame_team_calls.to_html(classes='mystyle'), user=user, tm=tm))

    # Create html table to present Calls in private conversations
    private_calls = []
    for call_private in arrayCallOneToOne:
        new = [call_private.timestart, call_private.timefinish, call_private.criador.nome, call_private.criador.email,
               call_private.presente.nome, call_private.presente.email, user]
        private_calls.append(new)

    private_calls.sort(key=first_element)

    frame_private_calls = pd.DataFrame(private_calls, columns=["Date Start Call", "Date End Call", "Call Creator Name",
                                                               "Call Creator Email", "Call Participant Name",
                                                               "Call Participant Email", "User"])
    with open(os.path.join(pathToFolder, "Calls_Private_Conversations_{}_{}.html".format(user, tm)), 'w',
              encoding="utf-8") as file:
        file.write(
            html_string.format(title="PRIVATE CALLS", table=frame_private_calls.to_html(classes='mystyle'), user=user,
                               tm=tm))

    # Create html table to present formation of new Teams
    teams_formation = []
    for k, v in dictionaryConversationDetails.items():
        for member in v.members:
            new = [v.conversation_id, v.date, v.creator.nome, v.creator.email, member.nome, member.email, user]
            teams_formation.append(new)

    teams_formation.sort(key=first_element)

    frame_teams_formation = pd.DataFrame(teams_formation,
                                         columns=["Conversation", "Date", "Name of User who added member",
                                                  "Email of User who added member", "Member Name", "Member Email",
                                                  "User"])

    with open(os.path.join(pathToFolder, "Teams_Formation_{}_{}.html".format(user, tm)), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="TEAMS FORMATION", table=frame_teams_formation.to_html(classes='mystyle'),
                                      user=user, tm=tm))

    # Create html table to present messages files
    frame_files = pd.read_csv(os.path.join(pathToFolder, "Files_{}_{}.csv".format(user, tmCSV)), delimiter=";",
                              header=0,
                              names=["Message ID", "File name", "File link", 'User'])
    with open(os.path.join(pathToFolder, "User_Messages_Files_{}_{}.html".format(user, tm)), 'w',
              encoding="utf-8") as file:
        file.write(
            html_string.format(title="MESSAGES FILES", table=frame_files.to_html(classes='mystyle'), user=user, tm=tm))

    # Create html table to present messages reactions
    frame_reacts = pd.read_csv(os.path.join(pathToFolder, "Reacts_{}_{}.csv".format(user, tmCSV)), delimiter=";",
                               header=0,
                               names=["Message ID", "Reaction", "Reacted by", 'Date', 'User'])
    with open(os.path.join(pathToFolder, "User_Messages_Reacts_{}_{}.html".format(user, tm)), 'w',
              encoding="utf-8") as file:
        file.write(
            html_string.format(title="MESSAGES REACTIONS", table=frame_reacts.to_html(classes='mystyle'), user=user,
                               tm=tm))

    # Create html table to present messages
    frame_messages = pd.read_csv(os.path.join(pathToFolder, "Mensagens_{}_{}.csv".format(user, tmCSV)), delimiter=";",
                                 header=0,
                                 names=["Message ID", "Message", "Date", 'Sender', 'Conversation ID', 'User'],
                                 encoding="utf-8")
    with open(os.path.join(pathToFolder, "User_Messages_{}_{}.html".format(user, tm)), 'w', encoding="utf-8") as file:
        file.write(
            html_string.format(title="MESSAGES", table=frame_messages.to_html(classes='mystyle'), user=user, tm=tm))

    # generate index html page
    html_string_index = '''
                <!DOCTYPE html>
            <html>
                <head>
                    <title>MS Teams Data</title>
                    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
                    <link rel="stylesheet" type="text/css" href="style_{user}_{tm}.css"/
                </head>
                <body>
                    <h1 style="text-align: center; padding: 10% 0;">Microsoft Teams extracted user data</h1>
                    <h1 style="text-align: center; padding: 10% 0;">User: {user}</h1>
                    <ul>
                        <li>
                            <a href="User_Contacts_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-address-book"></i>
                                    <i class="fa fa-address-book"></i>
                                </div>
                                <div class="name">Contacts</div>
                                <div class="number">{contacts_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="User_Messages_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-envelope"></i>
                                    <i class="fa fa-envelope"></i>
                                </div>
                                <div class="name">Messages</div>
                                <div class="number">{messages_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="User_Messages_Files_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-file"></i>
                                    <i class="fa fa-file"></i>
                                </div>
                                <div class="name">Messages Files</div>
                                <div class="number">{messages_files_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="User_Messages_Reacts_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-thumbs-up"></i>
                                    <i class="fa fa-thumbs-up"></i>
                                </div>
                                <div class="name">Messages Reactions</div>
                                <div class="number">{messages_reacts_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="Calls_Private_Conversations_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-phone"></i>
                                    <i class="fa fa-phone"></i>
                                </div>
                                <div class="name">Private Calls</div>
                                <div class="number">{private_calls_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="Teams_Formation_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-users"></i>
                                    <i class="fa fa-users"></i>
                                </div>
                                <div class="name">Teams Formations</div>
                                <div class="number">{teams_formation_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="Calls_Teams_{user}_{tm}.html">
                                <div class="icon">
                                    <i class="fa fa-phone"></i>
                                    <i class="fa fa-phone"></i>
                                </div>
                                <div class="name">Teams Calls</div>
                                <div class="number">{team_calls_number}</div>
                            </a>
                        </li>
                    </ul>
                </body>
            </html>
            '''
    with open(os.path.join(pathToFolder, "index_{}_{}.html".format(user, tm)), 'w', encoding="utf-8") as file:
        file.write(html_string_index.format(contacts_number=str(frame_contacts.__len__()),
                                            messages_number=str(frame_messages.__len__()),
                                            messages_files_number=str(frame_files.__len__()),
                                            messages_reacts_number=str(frame_reacts.__len__()),
                                            private_calls_number=str(frame_private_calls.__len__()),
                                            teams_formation_number=str(frame_teams_formation.__len__()),
                                            team_calls_number=str(frame_team_calls.__len__()),
                                            user=user, tm=tm))

    css_string = '''
                    body 
                    {
                        margin: 0;
                        padding: 0;
                        font-family: sans-serif;
                    }

                    ul 
                    {
                        position: absolute;
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%);
                        margin: 0;
                        padding: 20px 0;
                        display: flex;
                    }

                    ul li 
                    {
                        list-style: none;
                        text-align: center;
                        display: block;
                    }

                    ul li:last-child 
                    {
                        border-right: none;
                    }

                    ul li a 
                    {
                        text-decoration: none;
                        padding: 0 20px;
                        display: block;
                    }

                    ul li a .icon
                    {
                        width: 150px;
                        height: 40px;
                        text-align: center;
                        overflow: hidden;
                        margin: 0 auto 10px;
                    }

                    ul li a .icon .fa
                    {
                        width: 100%;
                        height: 100%;
                        line-height: 40px;
                        font-size: 34px;
                        transition: 0.3s;
                        color: #000;
                    }

                    ul li a:hover .icon .fa:last-child
                    {
                        color: royalblue;
                    }

                    ul li a:hover .icon .fa
                    {
                        transform: translateY(-100%);
                    }

                    .name
                    {
                        color: #000;
                    }

                    .number
                    {
                        color: red;
                    }

            '''

    with open(os.path.join(pathToFolder, "style_{}_{}.css".format(user, tm)), 'w') as f:
        f.write(css_string)

    html = open(os.path.join(pathToFolder, "User_Messages_{}_{}.html".format(user, tm)), 'r', encoding="utf-8")
    source_code = html.read()
    tables = pd.read_html(source_code)
    for i, table in enumerate(tables):
        table.to_csv(os.path.join(pathToFolder, 'messages_from_html_{}_{}.csv'.format(user, tm)), ';')
