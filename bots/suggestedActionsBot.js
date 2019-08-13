// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

class SuggestedActionsBot extends ActivityHandler {
    constructor() {
        super();

        this.onMembersAdded(async (context, next) => {
            await this.sendWelcomeMessage(context);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            const text = context.activity.text;
          
            // checks to see which agent should handle the problem
            const testLCOS = [`lcos`, 'new computer', 'lcos server'];
            if (testLCOS.includes(text)) {
                await context.sendActivity(`Contact Jen`);
            }
            const James = ['Printer', 'Printers', 'printer', 'printers', 'scanner', 'Scanner', 'scanning','Phones', 'phones', 'Outlook/Email', 'Oulook', 'Email', 'Locked out of Profile', 'unable to sign in'];
            if (James.includes(text)) {
                await context.sendActivity('Contact James');
            }
            const testEd = ['Navigator', 'navigator', 'eCenter', 'performance series', 'UC Navigator', 'initial confrence', 'students', 'uc navigator', 'UC navigator'];
            if (testEd.includes(text)) {
                await context.sendActivity('Contact Ed Torres');
            } else {
                await context.sendActivity('Please select a option.');
            }

            // After the bot has responded send the suggested actions.
            await this.sendSuggestedActions(context);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    /**
     * Send a welcome message along with suggested actions for the user to click.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendWelcomeMessage(turnContext) {
        const { activity } = turnContext;

        // Iterate over all new members added to the conversation.
        for (const idx in activity.membersAdded) {
            if (activity.membersAdded[idx].id !== activity.recipient.id) {
                const welcomeMessage = `Welcome to James's Support Bot ${ activity.membersAdded[idx].name }. ` +
                    `This bot will help you with your IT problem. ` +
                    `Please select an option:`;
                await turnContext.sendActivity(welcomeMessage);
                await this.sendSuggestedActions(turnContext);
           
            }
        }
    }

    /**
     * Send suggested actions to the user.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendSuggestedActions(turnContext) {
        var reply = MessageFactory.suggestedActions(['Printers', 'LCOS', 'Navigator', 'Phones', 'Outlook/Email', 'Locked out of Profile'], 'What is your issue?');
        await turnContext.sendActivity(reply);
    }
}

module.exports.SuggestedActionsBot = SuggestedActionsBot;
