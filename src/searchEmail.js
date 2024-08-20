const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { CardFactory, MessageFactory } = require("botbuilder");

class SearchEmail {
  triggerPatterns = ["find email", "find -e"];

  async handleCommandReceived(context, message) {   
    // verify the command arguments which are received from the client if needed.
    console.log(`App received message: ${message.text}`);

    // do something to process your command and return message activity as the response

    // render your adaptive card for reply message
    async function performSearch(queryStringg) {
      const searchEndpoint = "https://graph.microsoft.com/v1.0/search/query";
      const authToken =
        "";

      const requestBody = {
        requests: [
          {
            entityTypes: ["message"],
            query: {
              queryString: queryStringg,
            },
          },
        ],
      };

      // Make a POST request to the search endpoint with the auth token
      const response = await fetch(searchEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${authToken}`,
        },
        body: JSON.stringify(requestBody),
      });

      // Process the response and return the result
      const result = await response.json();

      return result;
    }

    const searchResults = await performSearch(message.text.split(" ").slice(1).slice(1).join(" "));

    let cardData;
    let searchURL;

    let allResults = [
      {
        type: "TextBlock",
        size: "Medium",
        weight: "Bolder",
        text: "${title}",
      },
    ];

    let i = 0;

    

    console.log(searchResults);

    if (searchResults.value[0].hitsContainers[0].total != 0) {
      await searchResults.value[0].hitsContainers[0].hits.forEach((result) => {
        i++;
        console.log(result);
        allResults.push({
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: `${i}: ${result.summary}`,
              size: "Small",
              wrap: true,
            },
            {
              type: "ActionSet",
              actions: [
                {
                  type: "Action.OpenUrl",
                  title: "Open Link",
                  url: `${result.resource.webLink}`,
                },
              ],
            },
          ],
        });
      });
    }

    if (searchResults.value[0].hitsContainers[0].total === 0) {
      allResults = [
        {
          type: "TextBlock",
          size: "Medium",
          weight: "Bolder",
          text: "No results found.",
        },
        {
          type: "TextBlock",
          text: "Try searching for something else.",
        },
      ];

      searchURL = "http://docs.aug.local:8888";
    } else {
      cardData = {
        title: `Search: ${message.text.split(" ").slice(1).join(" ")}`,
      };

      searchURL = `${searchResults.value[0].hitsContainers[0].hits[0].resource.webUrl}`;
    }

    console.log(allResults);

    const cardJson = AdaptiveCards.declare({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.3",
      body: allResults,
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
    }).render(cardData);
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}

module.exports = {
  SearchEmail,
};
