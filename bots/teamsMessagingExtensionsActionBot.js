// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { htmlToText } = require("html-to-text");
require("dotenv").config();
const {
  TeamsActivityHandler,
  CardFactory,
  TeamsInfo,
  MessageFactory,
} = require("botbuilder");
const baseurl = process.env.BaseUrl;

class TeamsMessagingExtensionsActionBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  /**
   *
   * @param {*} context Chat
   * @param {*} action Command
   * @returns It will listen the actions being performed inside the chat and perform actions such as sending adaptive card inside the chat
   */
  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    switch (action.commandId) {
      case "createCard":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      case "webView":
        return await webViewResponse(action);
    }
  }

  /**
   *
   * @param {*} context : Chat information from the MS teams chat
   * @param {*} action : We only have one action which is getting webview of frontend inside the task module, other one is default bioler code to handle corner cases
   * @returns
   */

  async handleTeamsMessagingExtensionFetchTask(context, action) {
    const serviceUrl = context.activity.serviceUrl;
    console.log("serviceUrl: ", serviceUrl);
    switch (action.commandId) {
      case "webView":
        return recognizeTab();

      default:
        try {
          const member = await this.getSingleMember(context);
          return {
            task: {
              type: "continue",
              value: {
                card: GetAdaptiveCardAttachment(),
                height: 400,
                title: `Hello ${member}`,
                width: 300,
              },
            },
          };
        } catch (e) {
          if (e.code === "BotNotInConversationRoster") {
            return {
              task: {
                type: "continue",
                value: {
                  card: GetJustInTimeCardAttachment(),
                  height: 400,
                  title: "Adaptive Card - App Installation",
                  width: 300,
                },
              },
            };
          }
          throw e;
        }
    }
  }

  /**
   *
   * @param {*} context : chat
   * @returns : checks if the person is part of convo or not
   */
  async getSingleMember(context) {
    try {
      const member = await TeamsInfo.getMember(
        context,
        context.activity.from.id
      );
      return member.name;
    } catch (e) {
      if (e.code === "MemberNotFoundInConversation") {
        context.sendActivity(MessageFactory.text("Member not found."));
        return e.code;
      }
      throw e;
    }
  }
}

/**
 * Boiler code (Leaving it as can be used for future use-cases) such as sending warning, error message
 * @returns
 */
function GetJustInTimeCardAttachment() {
  return CardFactory.adaptiveCard({
    actions: [
      {
        type: "Action.Submit",
        title: "Continue",
        data: { msteams: { justInTimeInstall: true } },
      },
    ],
    body: [
      {
        text: "Looks like you have not used Action Messaging Extension app in this team/chat. Please click **Continue** to add this app.",
        type: "TextBlock",
        wrap: true,
      },
    ],
    type: "AdaptiveCard",
    version: "1.0",
  });
}

/**
 * Boilder code An other template for adaptive card can be used in future
 * @returns
 */
function GetAdaptiveCardAttachment() {
  return CardFactory.adaptiveCard({
    actions: [{ type: "Action.Submit", title: "Close" }],
    body: [
      {
        text: "This app is installed in this conversation. You can now use it to do some great stuff!!!",
        type: "TextBlock",
        isSubtle: false,
        wrap: true,
      },
    ],
    type: "AdaptiveCard",
    version: "1.0",
  });
}

/**
 *
 * @param {*} to receiver name
 * @param {*} from sender name
 * @param {*} ApplauseID will be used for opening the task module for preview
 * @param {*} HeaderURL image appearing on task module / adaptive card
 * @param {*} richText Message
 * @param {*} HeaderTitle Recignized for
 * @param {*} StrengthIcons Icons for strength
 * @param {*} coins Points / Bonuses
 * @returns It will send an adaptive card in the chat with all above information
 */
function GetAdaptiveCard2(
  to,
  from,
  ApplauseID,
  HeaderURL,
  richText,
  HeaderTitle,
  StrengthIcons,
  coins
) {
  console.log("Hello", HeaderTitle, StrengthIcons, coins);

  const cardBody = [
    {
      type: "ColumnSet",
      columns: [
        {
          type: "Column",
          width: 40,
          items: [
            {
              type: "Image",
              url: `${HeaderURL}`,
            },
          ],
        },
        {
          type: "Column",
          width: 60,
          items: [
            {
              type: "TextBlock",
              size: "Medium",
              weight: "Bolder",
              text: HeaderTitle,
            },
            {
              type: "TextBlock",
              text: `To:${to} From:${from}`,
              wrap: true,
              weight: "Bolder",
              color: "Accent",
              isSubtle: true,
            },
            {
              type: "TextBlock",
              text: `${richText}`,
              wrap: true,
            },
          ],
          height: "stretch",
        },
      ],
    },
  ];

  if (coins > 0) {
    cardBody[0].columns[1].items.push({
      type: "TextBlock",
      text: `You have earned ${coins} points`,
      color: "Attention",
      isSubtle: true,
      weight: "Bolder",
      size: "Small",
      fontType: "Default",
    });
  }

  if (StrengthIcons?.length > 0) {
    const imageSet = {
      type: "ImageSet",
      images: StrengthIcons.map((iconUrl) => ({
        type: "Image",
        size: "Small",
        url: iconUrl,
      })),
      horizontalAlignment: "Right",
      height: "stretch",
      imageSize: "Small",
    };
    cardBody[0].columns[1].items.push(imageSet);
  }

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: cardBody,
    actions: [
      {
        type: "Action.OpenUrl",
        title: "Join the Conversation",

        // url: "https://teams.microsoft.com/l/entity/ADMIN CENTER APP ID/74bb6af2-3743-4ffc-bfc0-ca80f5152772?tenantId=44f47d0f-9269-44ca-9bbd-8d7df9cf7363",
        url: `https://teams.microsoft.com/l/task/36919edb-7aa7-4d47-a45c-bf5f8287f599?url=https://dev-teams-rn-frontend.azurewebsites.net/message-pop-up/${ApplauseID}&height=1000&width=1300&title=Preview&completionBotId=5501328b-3464-4280-9f97-9c6a98654850`,
      },
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.3",
    minHeight: "200px",
    msteams: {
      width: "full",
    },
    verticalContentAlignment: "Center",
    rtl: false,
  });
}

function GetGenericCard(cardType) {
  const cardBody = [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: `${cardType.toUpperCase()} Card`,
      style: "heading",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: "This Card has been sent to the channel",
      wrap: true,
      style: "default",
      fontType: "Default",
      size: "Default",
      spacing: "Medium",
    },
    {
      type: "TextBlock",
      spacing: "None",
      text: `Published On ${new Date().toLocaleDateString()}`,
      isSubtle: true,
      wrap: true,
    },
  ];

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: cardBody,
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
  });
}

function GetStandardCard(message, file) {
  const cardBody = [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "Getting To Know You, Share A Photo, Opinions and Debate", // You can replace this with a dynamic title if needed
    },
    {
      type: "ColumnSet",
      columns: [
        {
          type: "Column",
          width: "stretch",
          items: [
            {
              type: "TextBlock",
              text: message,
              wrap: true,
            },
            {
              type: "Image",
              url: image,
              size: "Stretch",
              width: "100%",
              altText: "Award Image",
            },
            {
              type: "TextBlock",
              spacing: "None",
              text: `Published On ${new Date().toLocaleDateString()}`,
              isSubtle: true,
              wrap: true,
            },
          ],
        },
      ],
    },
  ];

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: cardBody,
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
  });
}

function GetDelayedCard(question, delayedFollowUp, image) {
  const cardBody = [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "Brain Teaser/Delayed",
      style: "heading",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: `🧠Brain teaser time!🧠: ${question}`,
      wrap: true,
    },
    {
      type: "Image",
      url: image,
      size: "Stretch",
      width: "100%",
      altText: "Award Image",
    },
    {
      type: "TextBlock",
      text: `The answer will be posted in ${delayedFollowUp} minutes`,
      wrap: true,
      weight: "Bolder",
      horizontalAlignment: "Center",
    },
    {
      type: "TextBlock",
      spacing: "None",
      text: `Published On ${new Date().toLocaleDateString()}`,
      isSubtle: true,
      wrap: true,
    },
  ];

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: cardBody,
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
  });
}

function GetTeamAwardCard(awardName, awardCriteria, votingDuration, image) {
  const cardBody = [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "Team Award",
      style: "heading",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: "Award Type:",
      wrap: true,
      style: "default",
      fontType: "Default",
      size: "Default",
      weight: "Bolder",
      spacing: "Medium",
    },
    {
      type: "TextBlock",
      text: awardName,
      wrap: true,
    },
    {
      type: "TextBlock",
      text: "Award Criteria:",
      wrap: true,
      weight: "Bolder",
      spacing: "Medium",
    },
    {
      type: "RichTextBlock",
      inlines: [
        {
          type: "TextRun",
          text: awardCriteria,
        },
      ],
    },
    {
      type: "Image",
      url: image,
      size: "Stretch",
      width: "100%",
      altText: "Award Image",
    },
    {
      type: "TextBlock",
      text: `The voting time is ${votingDuration} minutes`,
      wrap: true,
      weight: "Bolder",
      horizontalAlignment: "Center",
    },
    {
      type: "TextBlock",
      spacing: "None",
      text: `Published On ${new Date().toLocaleDateString()}`,
      isSubtle: true,
      wrap: true,
    },
  ];

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: cardBody,
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
  });
}

function GetPollCard(message, pollOptions, image) {
  // Prepare the poll options as choice set
  const choiceSet = {
    type: "Input.ChoiceSet",
    id: "pollChoice",
    style: "expanded",
    choices: pollOptions.map((option) => ({
      title: option,
      value: option,
    })),
  };

  const cardBody = [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "Poll",
      style: "heading",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: "Question:",
      wrap: true,
      style: "default",
      fontType: "Default",
      size: "Default",
      weight: "Bolder",
      spacing: "Medium",
    },
    {
      type: "TextBlock",
      text: message,
      wrap: true,
    },
    choiceSet, // Add the choice set to the card
    {
      type: "Image",
      url: image,
      size: "Stretch",
      width: "100%",
      altText: "Poll Image",
    },
    {
      type: "TextBlock",
      spacing: "None",
      text: `Published On ${new Date().toLocaleDateString()}`,
      isSubtle: true,
      wrap: true,
    },
  ];

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: cardBody,
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.6",
  });
}

// User decided to create a new card
function createCardCommand(context, action) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

// For Sharing message
function shareMessageCommand(context, action) {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Messaging Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload.attachments &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

// this.onMessage(async (context, next) => {
//   const teamsChannelId = teamsGetChannelId(context.activity);
//   const activity = MessageFactory.text('This will be the first message in a new thread');
//   const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, process.env.MicrosoftAppId);

//   await context.adapter.continueConversationAsync(process.env.MicrosoftAppId, reference, async turnContext => {
//       await turnContext.sendActivity(MessageFactory.text('This will be the first response to the new thread'));
//   });

//   await next();
// });

/**
 * Render the frontend inside the task module
 * @returns
 */
function recognizeTab() {
  console.log("hello from tab");
  return {
    task: {
      type: "continue",
      value: {
        width: 800,
        height: 800,
        title: "Pick the channel you want to add Banter",
        url: `${baseurl}`,
      },
    },
  };
}

/**
 *  Helper function for removing markdown content
 * @param {*} htmlString
 * @returns
 */
function convertHtmlToMarkdown(htmlString) {
  const markdownString = htmlToText(htmlString, {
    wordwrap: false, // Prevent line breaks
    ignoreHref: true, // Ignore links
    ignoreImage: true, // Ignore images
  });

  return markdownString;
}

/**
 *
 * @param {*} action
 * @returns
 */
async function webViewResponse(action) {
  console.log("Hello from webview");
  const data = await action.data;
  console.log("Data received = ", data);

  let attachment;

  switch (data.cardType) {
    case "standard":
      attachment = GetStandardCard(data.message, data.image);
      break;
    case "delayed":
      attachment = GetDelayedCard(
        data.brainTeaserQuestion,
        data.delayFollowUp,
        data.image
      ); // Assuming a GetDelayedCard function
      break;
    case "teamAward":
      attachment = GetTeamAwardCard(
        data.awardName,
        data.awardCriteria,
        data.votingDuration,
        data.image
      ); // Assuming a GetTeamAwardCard function
      break;
    case "poll":
      attachment = GetPollCard(data.message, data.pollOptions, data.image); // Assuming a GetPollCard function
      break;
    default:
      console.warn("Unknown card type:", data.cardType);
      attachment = GetStandardCard(data.message, data.image); // Default case
      break;
  }

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

module.exports.TeamsMessagingExtensionsActionBot =
  TeamsMessagingExtensionsActionBot;
