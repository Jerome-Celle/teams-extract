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

const token = '';

async function main() {

    const options: AxiosRequestConfig = {
        headers: {
            'Authorization': `Bearer ${token}`,
            'Host': 'graph.microsoft.com'
        }
    };
    let nextChatListLink: string = 'https://graph.microsoft.com/v1.0/chats?$top=1'
    while (nextChatListLink) {
        try {
            const chatIdsResponse = await axios.get<ChatList>(nextChatListLink, options)
            nextChatListLink = chatIdsResponse.data["@odata.nextLink"]
            const chats: ChatList = chatIdsResponse.data
            console.log('NB chats', chats.value.length)

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
                        ...currentMessages.map(message =>
                            `${message.from?.user?.displayName}::${message.createdDateTime}::${message.body.content.replace(/(\r\n|\n|\r)/gm, "")}`
                        ),
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
                    fs.writeFileSync(`./output/${fileName}.txt`, messages.join('\r'));
                    console.log(`${chat.topic || [...users].join(',')} || ${messageNb}`)
                }
            }
        } catch (e) {

        }
    }

}

main()
