// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TimexProperty } = require('@microsoft/recognizers-text-data-types-timex-expression');
const { MessageFactory, InputHints } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { ComponentDialog, DialogSet, DialogTurnStatus, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';


class MainDialog extends ComponentDialog {
    constructor(luisRecognizer, boxCSDialog) {
        super('MainDialog');

        if (!luisRecognizer) throw new Error('[MainDialog]: Missing parameter \'luisRecognizer\' is required');
        this.luisRecognizer = luisRecognizer;


        // Define the main dialog and its related components.

        this.addDialog(new TextPrompt('TextPrompt'))
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                this.introStep.bind(this),
                this.actStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    /**
     * First step in the waterfall dialog. Prompts the user for a command.
     */
    async introStep(stepContext) {
        if (!this.luisRecognizer.isConfigured) {
            const messageText = 'NOTE: LUIS is not configured. To enable all capabilities, add `LuisAppId`, `LuisAPIKey` and `LuisAPIHostName` to the .env file.';
            await stepContext.context.sendActivity(messageText, null, InputHints.IgnoringInput);
            return await stepContext.next();
        }

        const messageText = stepContext.options.restartMsg ? stepContext.options.restartMsg : 'How can I help you? Select an option above or type your question:';
        const promptMessage = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);
        return await stepContext.prompt('TextPrompt', { prompt: promptMessage });
    }

    /**
    * Second step in the waterfall.
    */
    async actStep(stepContext) {
        let qryDetails = {};

        if (!this.luisRecognizer.isConfigured) {
            return await stepContext.beginDialog('boxCSDialog', qryDetails);
        }

        // Call LUIS
        var boxSetupMessageTxt = "";
        const luisResult = await this.luisRecognizer.executeLuisQuery(stepContext.context);
        console.log(`top Intent= ${LuisRecognizer.topIntent(luisResult)}`);
        switch (LuisRecognizer.topIntent(luisResult)) {

            case 'setupboxdrive':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/getting-started/box-drive) to get started.";
                await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'setupfolder':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://alexion.service-now.com/ask?id=sc_cat_item&sys_id=2bc828c313df6200faed51a63244b0cc) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'searchfiles':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/getting-started/sharing-collaborating) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'getoverview':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/getting-started) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'gettraining':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://alexion.service-now.com/ask?id=kb_article&sys_id=53c26b2ddb46a3840dfde9ec0b961923) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'acceesfrommobile':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/box-mobile/setting-up-box-for-mobile-devices) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'viewofficedocument':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/box-and-office-online/opening-editing-files) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'viewofficedocumentonmobile':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/box-mobile/office-apps-for-ios) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'editofficedocument':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/box-and-office-online/opening-editing-files) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'senddoctodocusign':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/apps-integrations/box-docusign) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'linkfile':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/getting-started/sharing-collaborating/shared-links-deep-dive) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'revokeaccess':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/help/Did-You-Know/DYK-Remove-Collaborator) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'customizeURL':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/help/Did-You-Know/DYK-Customize-URL) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'bestpractices':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://alexion.app.box.com/s/gf23rxrvy406hku08y8nbufa8p03acmo) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'None':
                // Catch all for unhandled intents
                console.log(`Inside None Intent`);

                const didntUnderstandMessageText = `Sorry I still don’t understand your question. Click this [link](https://alexion.service-now.com/ask) to open a ticket with IT Helpdesk and someone will get in touch with you. Thank you.`;
                await stepContext.context.sendActivity(didntUnderstandMessageText, didntUnderstandMessageText, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
        }

        return await stepContext.next();
    }

    /**
    * This is the final step in the main waterfall dialog.
    * 
    */
    async finalStep(stepContext) {

        if (stepContext.result) {
            /*
            const result = stepContext.result;
    
            const msg = `Still need to work on this. Like for sending confirmation message after completing any activity.`;
            await stepContext.context.sendActivity(msg, msg, InputHints.IgnoringInput);
            */
          //  console.log(`Inside stepcontext Result`);
            return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: 'What else can I do for you?' });
        }
        /*
        else {
            console.log(`Inside else && MaxAttempts=${MaxAttempts}`);
            
            if(cnt == MaxAttempts) {
                cnt =0;
                const didntUnderstandMessageText = `Sorry I still don’t understand your question. Click this [link](https://alexion.service-now.com/ask) to open a ticket with IT Helpdesk and someone will get in touch with you. Thank you.`;
                return await stepContext.context.sendActivity(didntUnderstandMessageText, didntUnderstandMessageText, InputHints.IgnoringInput);
            
            }
            
        }  */

        // Restart the main dialog with a different message the second time around
        return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: 'What else can I do for you?' });
    }


}

module.exports.MainDialog = MainDialog;
