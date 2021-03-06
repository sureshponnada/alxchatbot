// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TimexProperty } = require('@microsoft/recognizers-text-data-types-timex-expression');
const { MessageFactory, InputHints } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { ComponentDialog, DialogSet, DialogTurnStatus, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const UN_SUCCESSFUL_CNT = 'unSuccessfulCntProperty';




class MainDialog extends ComponentDialog {
    constructor(luisRecognizer, userState) {
        super('MainDialog', userState);


        if (!luisRecognizer) throw new Error('[MainDialog]: Missing parameter \'luisRecognizer\' is required');
        this.luisRecognizer = luisRecognizer;

        if (!userState) throw new Error('userState is undefined in MainDialog');
        this.userState = userState;

        this.unSuccessfulCntProperty = this.userState.createProperty(UN_SUCCESSFUL_CNT);

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
        await this.unSuccessfulCntProperty.set(stepContext.context, 0);
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
        if(LuisRecognizer.topIntent(luisResult) != 'None') {
            await this.unSuccessfulCntProperty.set(stepContext.context, 0);
            await this.userState.saveChanges(stepContext.context, false);
        }
        switch (LuisRecognizer.topIntent(luisResult)) {

            case 'setupboxdrive':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/getting-started/box-drive) to get started.";
                await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'setupfolder':
                boxSetupMessageTxt = "It is pretty easy to setup a Box folder. Submit a helpdesk ticket here [link](https://alexion.service-now.com/ask?id=sc_cat_item&sys_id=2bc828c313df6200faed51a63244b0cc) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'searchfile':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/guides/getting-started/sharing-collaborating) to know how to search in Box.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'getoverview':
                boxSetupMessageTxt = "Check this video [link](https://box.alexion.com/guides/getting-started) that provides an insightful overview of the Box @Alexion.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'gettraining':
                boxSetupMessageTxt = "Check out upcoming trainings at [link](https://alexion.service-now.com/ask?id=kb_article&sys_id=53c26b2ddb46a3840dfde9ec0b961923) and add a training class to your calendar.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'acceesfrommobile':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/guides/box-mobile/setting-up-box-for-mobile-devices) to access Box files from your mobile phone.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'viewofficedocument':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/guides/box-and-office-online/opening-editing-files) to learn the steps to view office documents.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'viewofficedocumentonmobile':
                boxSetupMessageTxt = "You work on Box files right from your laptop. Check out this [link](https://box.alexion.com/guides/box-mobile/office-apps-for-ios) to get started.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'editofficedocument':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/guides/box-and-office-online/opening-editing-files) to learn the steps to edit office documents.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'senddoctodocusign':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/guides/apps-integrations/box-docusign) to learn the steps to docusign a document in Box.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'linkfile':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/guides/getting-started/sharing-collaborating/shared-links-deep-dive) to learn the steps to share a document in Box.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'revokeaccess':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/help/Did-You-Know/DYK-Remove-Collaborator) to learn the steps to revoke access of existing collaborator.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'customizeURL':
                boxSetupMessageTxt = "Check out this [link](https://box.alexion.com/help/Did-You-Know/DYK-Customize-URL) to learn the steps to customize a Link to Box document.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'bestpractices':
                boxSetupMessageTxt = "Check out this [link](https://alexion.app.box.com/s/gf23rxrvy406hku08y8nbufa8p03acmo) to learn Best Practices for Box usage.";
                return await stepContext.context.sendActivity(boxSetupMessageTxt, boxSetupMessageTxt, InputHints.IgnoringInput);
                return await stepContext.next();
                break;
            case 'None':
                var cnt = await this.unSuccessfulCntProperty.get(stepContext.context);
                var iCnt = parseInt(cnt);

                if (!isNaN(iCnt) && iCnt >= 2) {
                    const didntUnderstandMessageText = `Sorry I still don’t understand your question. Click this [link](https://alexion.service-now.com/ask) to open a ticket with IT Helpdesk and someone will get in touch with you. Thank you.`;
                    await stepContext.context.sendActivity(didntUnderstandMessageText, didntUnderstandMessageText, InputHints.IgnoringInput);
                    await this.unSuccessfulCntProperty.set(stepContext.context, 0);
                    await this.userState.saveChanges(stepContext.context, false);                  
                    return await stepContext.next();
                    break;
                } else {
                    iCnt = iCnt + 1;
                    await this.unSuccessfulCntProperty.set(stepContext.context, iCnt);
                    const didntUnderstandMessageText = `Sorry I don’t understand your question. Please type what you are looking for.`;
                    const promptMessage = MessageFactory.text(didntUnderstandMessageText, didntUnderstandMessageText, InputHints.ExpectingInput);
                    await this.userState.saveChanges(stepContext.context, false);
                    return await stepContext.prompt('TextPrompt', { prompt: promptMessage });
                    break;

                }
        }

        return await stepContext.next();
    }

    /**
    * This is the final step in the main waterfall dialog.
    * 
    */
    async finalStep(stepContext) {

        if (stepContext.result) {

            return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: 'What else can I do for you?' });
        }
        else {
            return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: 'What else can I do for you?' });
        }

    }

}

module.exports.MainDialog = MainDialog;
