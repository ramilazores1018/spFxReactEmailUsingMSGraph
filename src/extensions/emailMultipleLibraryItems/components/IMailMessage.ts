export interface IMailMessage {
    message: {
        toRecipients: IRecipient[],
        ccRecipients: IRecipient[],
        subject:string,
        body: {
            contentType:string;
            content: string;
        },
        attachments: IAttachment[],

    };
    saveToSentItems: boolean;

}

export interface IRecipient {

    emailAddress: {
        address: string;
    };

}

export interface IAttachment {

    
    contentBytes: string;
    odatatype: string;
    isInline: boolean;
    name: string;

}