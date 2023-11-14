import axios, {AxiosRequestConfig, AxiosResponse} from "axios";
import * as fs from "fs";

type ChatMessage = {
    "id": string;
    "replyToId": string;
    "etag": string;
    "messageType": string;
    "createdDateTime": string;
    "lastModifiedDateTime": string;
    "lastEditedDateTime": string;
    "deletedDateTime": string;
    "subject": string;
    "summary": string;
    "chatId": string;
    "importance": string;
    "locale": string;
    "webUrl": string;
    "channelIdentity": string;
    "policyViolation": string;
    "eventDetail": string;
    "from": {
        "application": string;
        "device": string;
        "user": {
            "@odata.type": string;
            "id": string;
            "displayName": string;
            "userIdentityType": string;
            "tenantId": string;
        }
    },
    "body": {
        "contentType": string;
        "content": string;
    },
    "attachments": string[],
    "mentions": string[][],
    "reactions": string[][]
}
type ChatMessageData = {
    "@odata.context": string;
    "@odata.count": number;
    "@odata.nextLink": string;
    "@microsoft.graph.tips": string;
    "value": ChatMessage[]

}

type Chat = {
    "id": string;
    "topic": string;
    "createdDateTime": string;
    "lastUpdatedDateTime": string;
    "chatType": string;
    "chatViewpoint": {
        "isHidden": false,
        "lastMessageReadDateTime": string;
    },
    "webUrl": string;
};
type ChatList = {
    "@odata.context": string;
    "@odata.count": number;
    "@odata.nextLink": string;
    "@microsoft.graph.tips": string;
    "value": Chat []
}
const DL_FILES = false;
const token = '';
const options: AxiosRequestConfig = {
    headers: {
        'Authorization': `Bearer ${token}`,
        'Host': 'graph.microsoft.com'
    },
    // proxy: {
    //     'host': '127.0.0.1',
    //     'port': 9000
    // }
};

export async function downloadFile(fileUrl: string, outputLocationPath: string): Promise<any> {
    return axios.get(fileUrl, {
        responseType: 'stream',
        ...options
    }).then(response => {
        response.data.pipe(fs.createWriteStream(outputLocationPath));
    }).catch(() => console.log('error download',fileUrl));
}

async function main() {


    let nextChatListLink: string = 'https://graph.microsoft.com/v1.0/chats?$top=1'
    while (nextChatListLink) {
        try {
            const chatIdsResponse = await axios.get<ChatList>(nextChatListLink, options)
            nextChatListLink = chatIdsResponse.data["@odata.nextLink"]
            const chats: ChatList = chatIdsResponse.data

            for (const chat of chats.value) {
                const chatMessagesUrl: string = `https://graph.microsoft.com/v1.0/chats/${chat.id}/messages?$top=50&$orderby=createdDateTime desc`;
                let messages: string[] = [];
                const users = new Set<string>()
                let nextLink: string = chatMessagesUrl
                while (nextLink) {


                    const chatMessagesResponse: AxiosResponse<ChatMessageData> = await axios.get<ChatMessageData>(nextLink, options)

                    const chatMessages = chatMessagesResponse.data
                    nextLink = chatMessages["@odata.nextLink"];
                    const currentMessages = chatMessages.value.filter(message => message.body.content && message.body.content !== '<systemEventMessage/>')

                    currentMessages.map((message => message.from?.user?.displayName)).forEach(user => users.add(user));
                    messages = [
                        ...messages,
                        ...(currentMessages.map(message => {
                                let messageContent = message.body.content.replace(/(\r\n|\n|\r|<p>&nbsp;<\/p>|<p><\/p>)/gm, "");
                                if (!messageContent.includes('<p>')) {
                                    messageContent = `<p>${messageContent}</p>`
                                }
                                if (messageContent.includes('src')) {

                                    const match = messageContent.match(/https:\/\/(.*)hostedContents\/(.+?)\/\$value/);
                                    if (match) {
                                        const newFileName = `files/${match[2]}`
                                        messageContent = messageContent.replace(match[0], newFileName);
                                        DL_FILES && downloadFile(match[0], `./output/${newFileName}`).then()
                                    }
                                }
                                return `<div class="username">${message.from?.user?.displayName} - ${message.createdDateTime}</div> 
                                ${messageContent}
                            `
                            }
                        )),
                    ]

                }
                const messageNb = messages.length
                if (messageNb > 0) {

                    const fileName = `${chat.topic || [...users].join(',')}-${messageNb}`
                        .split('"').join('-')
                        .split('?').join('-')
                        .split('/').join('-')
                        .split('\\').join('-')
                        .split(':').join('-')
                        .split(',').join('-')
                        .split(' ').join('-')
                    const content = `<html lang="fr">
                        <head>
                            <meta charset="utf-8">
                            <style>
                            body {background-color: #f5f5f5;
                                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif
                            }
                            .username {color: #616161}
                            p    {color: black; 
                                background-color: white; 
                                padding: 10px;
                                margin-left: 25px;
                                border-radius: 5px;
                            }
                            </style><title>fileName</title>
                        </head>
                        <body>${messages.reverse().join('')}</body>
                        </html>`;
                    fs.writeFileSync(`./output/${fileName}.html`, content);
                    console.log(`${chat.topic || [...users].join(',')} || ${messageNb}`)
                }
            }
        } catch (e) {

        }
    }

}

main()
