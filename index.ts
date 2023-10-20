import axios, {AxiosRequestConfig, AxiosResponse} from "axios";

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
    "value": Chat []
}

const token = ''

async function main() {

    const options: AxiosRequestConfig = {
        headers: {
            'Authorization': `Bearer ${token}`,
            'Host': 'graph.microsoft.com'
        }
    };
    const chatListUrl: string = 'https://graph.microsoft.com/v1.0/chats';
    const chatIdsResponse = await axios.get<ChatList>(chatListUrl, options)
    const chats: ChatList = chatIdsResponse.data
    console.log('NB chats', chats.value.length)

    for (const chat of chats.value) {
        const chatMessagesUrl: string = `https://graph.microsoft.com/v1.0/chats/${chat.id}/messages?$top=50&$orderby=createdDateTime desc`;
        // let messages: string[] = [];
        let messageNb = 0;
        let lastDate = ''
        const users = new Set<string>()
        let nextLink: string | null = chatMessagesUrl
        while (nextLink) {

            const chatMessagesResponse: AxiosResponse<ChatMessageData> = await axios.get<ChatMessageData>(nextLink, options)
            const chatMessages = chatMessagesResponse.data
            nextLink = chatMessages["@odata.nextLink"];
            messageNb += chatMessages.value.length
            lastDate = chatMessages.value[chatMessages.value.length - 1].createdDateTime
            users.add(chatMessages.value[chatMessages.value.length - 1].from?.user?.displayName)
            // messages = [
            //     ...messages,
            //     ...chatMessages.value.map(message => message.createdDateTime)
            // ]
        }
        console.log(chat.topic, `|${[...users].join('|')}|`, lastDate, messageNb)
    }

}

main()
