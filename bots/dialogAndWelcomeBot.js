// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { CardFactory } = require('botbuilder');
const { DialogBot } = require('./dialogBot');
const WelcomeCard = require('./resources/welcomeCard.json');
const WELCOMED_USER = 'welcomedUserProperty';
const UN_SUCCESSFUL_CNT = 'unSuccessfulCntProperty';

class DialogAndWelcomeBot extends DialogBot {
    constructor(conversationState, userState, dialog) {
        super(conversationState, userState, dialog);
        this.welcomedUserProperty = userState.createProperty(WELCOMED_USER);
        this.unSuccessfulCntProperty = userState.createProperty(UN_SUCCESSFUL_CNT);
        this.userState = userState;

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const didBotWelcomedUser = await this.welcomedUserProperty.get(context, false);
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    if (didBotWelcomedUser === false) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                        await dialog.run(context, conversationState.createProperty('DialogState'));
                        await this.welcomedUserProperty.set(context, true);
                        await this.unSuccessfulCntProperty.set(context, 0);
                    }

                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
        this.onMessage(async (context, next) => {

            const didBotWelcomedUser = await this.welcomedUserProperty.get(context, false);
           // console.log(`Inside welcomeBot onMessage and prop=${didBotWelcomedUser} `);

        })

    }
}

module.exports.DialogAndWelcomeBot = DialogAndWelcomeBot;
