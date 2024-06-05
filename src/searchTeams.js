const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { CardFactory, MessageFactory } = require("botbuilder");

class SearchTeams {
  triggerPatterns = ["find message", "find -m"];

  async handleCommandReceived(context, message) {
    // verify the command arguments which are received from the client if needed.
    console.log(`App received message: ${message.text}`);

    // do something to process your command and return message activity as the response

    // render your adaptive card for reply message
    async function performSearch(queryStringg) {
      const searchEndpoint = "https://graph.microsoft.com/v1.0/search/query";
      const authToken =
        "eyJ0eXAiOiJKV1QiLCJub25jZSI6InlvSTBmbG5EMDFOTnFFMEtfSkFTOHViYnZBTzVkQmd4dEx1REk3dm5TSkUiLCJhbGciOiJSUzI1NiIsIng1dCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCIsImtpZCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8wMjFlZGI4Ni04NzU4LTRjNTYtOTM3Ny00ZTY2ZDIxOTZmYzAvIiwiaWF0IjoxNzE3NTk3NDQzLCJuYmYiOjE3MTc1OTc0NDMsImV4cCI6MTcxNzY4NDE0MywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhXQUFBQStiWXFjU0hrN3IxZ2N1eHJkM0JwYi84Mm9nOXNaellqTEcyQngzbGtONi9rc1RsbkVldnBtZUcvUzMyUWZGL0toVmNlNk1zQTRxblhkUllaK01kdVpOZzRQRlNsSGE1U2ltUTNjMG5ueGVzPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiVW5ndXJlYW51IiwiZ2l2ZW5fbmFtZSI6IkFsZXgiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIxNDIuMTEyLjIzLjg3IiwibmFtZSI6IkFsZXggVW5ndXJlYW51Iiwib2lkIjoiNGI2MGU4MzQtNTJkOC00MDRjLWEyYzItZDkxNThjNDgxOWU4IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAzN0I4NDI4OTgiLCJyaCI6IjAuQVdJQWh0c2VBbGlIVmt5VGQwNW0waGx2d0FNQUFBQUFBQUFBd0FBQUFBQUFBQUNNQUFRLiIsInNjcCI6IkNhbGVuZGFycy5SZWFkIENoYW5uZWxNZXNzYWdlLlJlYWQuQWxsIENoYXQuUmVhZCBFeHRlcm5hbEl0ZW0uUmVhZC5BbGwgRmlsZXMuUmVhZC5BbGwgTWFpbC5SZWFkIG9wZW5pZCBwcm9maWxlIFNpdGVzLlJlYWQuQWxsIFVzZXIuUmVhZCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IlBjT1pRY0hNbEJpTEJ1SHV4UEphSTFySllXeVNncU5ULXE0dVpXTmZRaWsiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiIwMjFlZGI4Ni04NzU4LTRjNTYtOTM3Ny00ZTY2ZDIxOTZmYzAiLCJ1bmlxdWVfbmFtZSI6ImFsZXhAYXVnc2lnbmFscy5jb20iLCJ1cG4iOiJhbGV4QGF1Z3NpZ25hbHMuY29tIiwidXRpIjoiTDVZa2hub2xtRVdmdlgxbi1HejVBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiZjI4YTFmNTAtZjZlNy00NTcxLTgxOGItNmExMmYyYWY2YjZjIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19jYyI6WyJDUDEiXSwieG1zX3NzbSI6IjEiLCJ4bXNfc3QiOnsic3ViIjoiTWRhX1FjZ0Y4SkxwclZoWGE4RC1rNlJVZExZemhwa0V4SmZFamh4MFNwSSJ9LCJ4bXNfdGNkdCI6MTU2NDU4NTExNn0.gIZZ_l2H0wzfi0JVpX_Q16GZezNfESpDkXCQ06hjzqWqQhl9maJcuOHc5Y1UtmFMA8Si-XATbNm07lnHUjAzmTKi6K9LruFiiD1GQYL3YYI_VvNLnZdLPuwbD3xy3ndnlwWBTrjr_RIEIWf_qiBm5qaEsOtw60AtOe41FlBvDGP1pxcFPCnibHx6z3rkJRRcUfj-fo3skAv0S_2jTCcWSUcNE_ulWsuv-we3aYkNxI6W9VafwKlRaLqUnFBG-I6yjmSY9nlsv66gcdJ1_XqfdSbb_k3Nmb56DPknI3Vsffdxwn-DPWcvwaTY1Twp82GAS2euBzeFWMCOq93oAMFOHA";

      const requestBody = {
        requests: [
          {
            entityTypes: ["chatMessage"],
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

    let length = 10;

    let query;

    if (message.text.split(" ").includes("--limit")) {
      query = message.text.split(" ").slice(1).slice(1).slice(0,-1).filter(e => e !== "--limit").join(" ");
      length = message.text.split(" ")[message.text.split(" ").length - 1];
      if(length == "all" || length == "ALL"){length = 1000}
    } else {
        query = message.text.split(" ").slice(1).slice(1).join(" ");
    }

    const searchResults = await performSearch(query);

    console.log(query);

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
        console.log(length);
        console.log(result);
        if(i <= length){allResults.push({
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
        })}
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
  SearchTeams,
};
